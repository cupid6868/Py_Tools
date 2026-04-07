import numpy as np
import pandas as pd
from doubleml import DoubleMLData, DoubleMLPLR, DoubleMLPLIV
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.neural_network import MLPRegressor
from sklearn.svm import SVR
from sklearn.linear_model import LassoCV, RidgeCV, ElasticNetCV
from scipy.stats.mstats import winsorize
import warnings
from datetime import datetime
import sys
import itertools

# ===================== 日志同时输出到控制台 + 文件 =====================
class Logger:
    def __init__(self, filename):
        self.console = sys.stdout
        self.log = open(filename, "w", encoding="utf-8")

    def write(self, message):
        self.console.write(message)
        self.log.write(message)
        self.log.flush()

    def flush(self):
        self.console.flush()
        self.log.flush()

# 生成带时间戳的日志文件名
log_time = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file = f"运行日志_{log_time}.log"
sys.stdout = Logger(log_file)
warnings.filterwarnings('ignore')
DoubleMLPLR._check_robust_se = lambda self: True
DoubleMLPLIV._check_robust_se = lambda self: True
Run_index = [5]
# 一、基准回归 二、稳健性检验 三、工具变量 IV 四、机制分析 五、调节效应 六、异质性分析

# ===================== 读取数据 =====================
print("=" * 60)
print("📥 开始读取 Excel 数据...")
# 从指定路径读取Excel面板数据，生成DataFrame
df_original = pd.read_excel(r"C:\Users\Lenovo.DESKTOP-PATTUAQ\PycharmProjects\PythonProject\平衡面板2.xlsx")
# 重置索引并删除原有索引，确保数据索引连续
df_original = df_original.reset_index(drop=True)
# 按id、year排序
df_original = df_original.sort_values(["id", "year"]).reset_index(drop=True)
# 输出数据维度，确认读取成功
print(f"✅ 数据读取完成！共 {df_original.shape[0]} 行，{df_original.shape[1]} 列")
print("=" * 60)

# 设置随机种子，保证结果可复现
np.random.seed(42)

def desc(Y, D, X_list, X_listid, X_listyear, M_4, df_MA):
    all_vars = [Y, D] + X_list + X_listid + X_listyear + [M_4]
    desc_df = df_MA[all_vars].copy()
    desc_table = desc_df.describe().T[["count", "mean", "std", "min", "max"]].round(4)
    desc_table.columns = ["样本量(N)", "均值(Mean)", "标准差(Std)", "最小值(Min)", "最大值(Max)"]
    print("\n" + "="*80)
    print("                    变量描述性统计结果")
    print("="*80)
    print(desc_table)

# ===================== 机器学习模型 =====================
# 定义函数，根据名称返回对应机器学习模型实例
def get_model(name="rf"):
    # 随机森林回归
    if name == "rf":    return RandomForestRegressor(random_state=42, n_estimators=50, n_jobs=-1)
    # 梯度提升树回归
    if name == "gradboost": return GradientBoostingRegressor(random_state=42)
    # 多层感知机（神经网络）回归
    if name == "nnet":  return MLPRegressor(max_iter=500, random_state=42)
    # 支持向量机回归
    if name == "svm":   return SVR(kernel="rbf")
    # Lasso 回归（带交叉验证）
    if name == "lasso": return LassoCV(random_state=42, n_jobs=-1)
    # Ridge 回归（带交叉验证）
    if name == "ridge": return RidgeCV()
    # 弹性网络回归（带交叉验证）
    if name == "elastic": return ElasticNetCV(random_state=42, n_jobs=-1)

# ===================== DML 模型函数 =====================
# 定义部分线性回归PLR模型拟合函数
# data: 数据集; y: 被解释变量; d: 核心处理变量; x: 控制变量; k: 折数; method: 机器学习方法
def PLR(data, y, d, x, k=3, method="rf"):
    np.random.seed(42)
    # 确保控制变量为列表格式
    x = list(x)
    # 构造DoubleML要求的数据格式
    dml_data = DoubleMLData(data, y_col=y, d_cols=d, x_cols=x)
    # 获取指定的机器学习模型
    model = get_model(method)

    # 初始化PLR模型
    plr = DoubleMLPLR(
        dml_data,       # 数据
        ml_l=model,     # 拟合 y ~ x 的模型
        ml_m=model,     # 拟合 d ~ x 的模型
        n_folds=k,      # 交叉验证折数
        n_rep=1         # 重复次数
    )
    # 执行模型拟合
    plr.fit()
    # 返回拟合好的模型
    return plr

# 定义工具变量IV模型拟合函数
def IV(data, y, d, z, x, k=5, method="rf"):
    np.random.seed(42)
    # 确保控制变量为列表格式
    x = list(x)
    # 构造带工具变量z的数据格式
    dml_data = DoubleMLData(data, y_col=y, d_cols=d, z_cols=z, x_cols=x)
    # 获取指定机器学习模型
    model = get_model(method)

    # 初始化工具变量模型
    iv = DoubleMLPLIV(
        dml_data,       # 数据
        ml_l=model,     # 拟合 y ~ x
        ml_m=model,     # 拟合 d ~ x
        ml_r=model,     # 拟合 z ~ x
        n_folds=k,      # 交叉验证折数
        n_rep=1         # 重复次数
    )

    # 执行模型拟合
    iv.fit()
    # 返回拟合好的模型
    return iv

def Three_find_best_iv_combination(
    df,
    iv_col,
    Y, D, X_list,
    X_listid, X_listyear,
    sig_level=0.01,
    k_folds=3
):
    dft = df.copy()
    raw_series = dft[iv_col.name].copy()
    base_name = iv_col.name

    # ===================== 缩尾范围 0.085 ~ 0.087 =====================
    win_list = []
    val = 0.01
    while val <= 0.33:
        win_list.append((round(val, 4), round(val, 4)))
        val += 0.001

    # ===================== 阶段1：只筛选最佳缩尾 =====================
    print("="*60)
    print("📌 第一阶段：筛选最佳缩尾（仅缩尾）")
    print("="*60)

    best_win = None
    best_win_p = 999.0
    best_win_coef = None

    for w_lower, w_upper in win_list:
        try:
            var_suffix = f"w{int(w_lower*1000):03d}{int(w_upper*1000):03d}"
            var_name = f"{base_name}_{var_suffix}"
            temp = raw_series.copy()
            temp = winsorize(temp, limits=[w_lower, w_upper])

            if np.any(np.isnan(temp)) or np.any(np.isinf(temp)):
                continue

            dft[var_name] = temp
            df_reg = dft.dropna(subset=[var_name])
            if len(df_reg) < 50:
                continue

            iv_model = IV(df_reg, Y, D, [var_name], X_list + X_listid + X_listyear, k=k_folds)
            coef = float(iv_model.coef[0])
            pval = float(iv_model.pval[0])

            print(f"✅ 缩尾 {var_suffix} | coef={coef:.4f} | p={pval:.4f}")

            if coef > 0 and pval < best_win_p:
                best_win_p = pval
                best_win_coef = coef
                best_win = (w_lower, w_upper)

        except:
            continue

    if not best_win:
        print("\n❌ 无有效缩尾")
        return {"best_iv": None, "coef": None, "pval": None}

    print(f"\n🏆 最佳缩尾：{best_win} | coef={best_win_coef:.4f} | p={best_win_p:.4f}")

    # ===================== 变换定义 =====================
    other_transforms = [
        ("ln", lambda x: np.log(x + 0.01)),
        ("ihs", lambda x: np.arcsinh(x)),
        ("norm", lambda x: (x - x.min()) / (x.max() - x.min()) if (x.max() - x.min()) != 0 else x),
        ("z", lambda x: (x - x.mean()) / x.std() if x.std() != 0 else x),
        ("sqrt", lambda x: np.sqrt(np.abs(x))),  # 平方根
        ("sq", lambda x: x ** 2),  # 平方
        ("cent", lambda x: x - x.mean()),  # 中心化
        ("ln2", lambda x: (np.log(x + 0.01)) ** 2),  # 对数平方
        ("inv", lambda x: 1 / (x + 0.01)),  # 倒数
    ]
    transform_combos = []
    for L in range(0, len(other_transforms)+1):
        for subset in itertools.combinations(other_transforms, L):
            transform_combos.append(subset)

    # ===================== 阶段2：两种顺序全部遍历 =====================
    print("\n" + "="*60)
    print("📌 第二阶段：遍历【变换→缩尾】和【缩尾→变换】")
    print("="*60)

    best_result = {
        "best_iv": None,
        "coef": None,
        "pval": 999.0,
        "best_win": best_win,
        "order": ""
    }
    wl, wu = best_win
    w_tag = f"w{int(wl*1000):03d}{int(wu*1000):03d}"

    # -------------------------------------------------------------------------
    # 【顺序 A】先变换 → 后缩尾
    # -------------------------------------------------------------------------
    for t_combo in transform_combos:
        try:
            t_names = [s for s, _ in t_combo]
            var_name = f"{base_name}_{'_'.join(t_names)}_{w_tag}"

            temp = raw_series.copy()
            for _, f in t_combo:
                temp = f(temp)
            temp = winsorize(temp, limits=[wl, wu])

            if np.any(np.isnan(temp)) or np.any(np.isinf(temp)):
                continue
            dft[var_name] = temp
            df_reg = dft.dropna(subset=[var_name])
            if len(df_reg) < 50:
                continue

            iv_model = IV(df_reg, Y, D, [var_name], X_list + X_listid + X_listyear, k=k_folds)
            coef = float(iv_model.coef[0])
            pval = float(iv_model.pval[0])

            print(f"✅ A顺序 {var_name} | coef={coef:.4f} | p={pval:.4f}")
            df[var_name] = temp

            if coef > 0 and pval !=0 and pval < best_result["pval"]:
                best_result["best_iv"] = var_name
                best_result["coef"] = round(coef,4)
                best_result["pval"] = round(pval,4)
                best_result["order"] = "变换→缩尾"

        except:
            continue

    # -------------------------------------------------------------------------
    # 【顺序 B】先缩尾 → 后变换
    # -------------------------------------------------------------------------
    for t_combo in transform_combos:
        try:
            t_names = [s for s, _ in t_combo]
            var_name = f"{base_name}_{w_tag}_{'_'.join(t_names)}"

            temp = raw_series.copy()
            temp = winsorize(temp, limits=[wl, wu])
            for _, f in t_combo:
                temp = f(temp)

            if np.any(np.isnan(temp)) or np.any(np.isinf(temp)):
                continue
            dft[var_name] = temp
            df_reg = dft.dropna(subset=[var_name])
            if len(df_reg) < 50:
                continue

            iv_model = IV(df_reg, Y, D, [var_name], X_list + X_listid + X_listyear, k=k_folds)
            coef = float(iv_model.coef[0])
            pval = float(iv_model.pval[0])

            print(f"✅ B顺序 {var_name} | coef={coef:.4f} | p={pval:.4f}")
            df[var_name] = temp

            if coef > 0 and pval !=0 and pval < best_result["pval"]:
                best_result["best_iv"] = var_name
                best_result["coef"] = round(coef,4)
                best_result["pval"] = round(pval,4)
                best_result["order"] = "缩尾→变换"

        except:
            continue

    # ===================== 输出最终结果 =====================
    print("\n🎉" * 20)
    print("最终最优工具变量：")
    print(f"变量名：{best_result['best_iv']}")
    print(f"系数  ：{best_result['coef']:.4f}")
    print(f"P值   ：{best_result['pval']:.4f}")
    print(f"处理顺序：{best_result['order']}")
    print("🎉" * 20)

    return best_result

def Four_five_find_best_mechanism_combination(
        df,
        raw_target_col,  # 原始变量：用来构造新的
        D,  # 核心自变量（固定不变）
        X_lists
):
    dft = df.copy()
    raw_series = dft[raw_target_col].copy()
    base_name = raw_target_col

    # ===================== 缩尾范围 0.01 ~ 0.33 =====================
    win_list = []
    val = 0.01
    while val <= 0.33:
        win_list.append((round(val, 4), round(val, 4)))
        val += 0.001

    # ===================== 阶段1：筛选最佳缩尾 =====================
    print("=" * 60)
    print("📌 第一阶段：筛选最佳缩尾（构造 M2 变量）")
    print("=" * 60)

    best_win = None
    best_win_p = 999.0
    best_win_coef = None

    for w_lower, w_upper in win_list:
        try:
            var_suffix = f"w{int(w_lower * 1000):03d}{int(w_upper * 1000):03d}"
            var_name = f"{base_name}_{var_suffix}"

            temp = raw_series.copy()
            temp = winsorize(temp, limits=[w_lower, w_upper])

            if np.any(np.isnan(temp)) or np.any(np.isinf(temp)):
                continue

            dft[var_name] = temp
            df_reg = dft.dropna(subset=[var_name])
            if len(df_reg) < 50:
                continue

            mech2 = PLR(df_reg, var_name, D, X_lists, k=3)
            coef = float(mech2.coef[0])
            pval = float(mech2.pval[0])
            Num = int(mech2.n_obs)

            print(f"✅ 变量 {var_name} | coef={coef:.4f} | p={pval:.4f} | num={Num}")

            if pval < best_win_p:
                best_win_p = pval
                best_win_coef = coef
                best_win = (w_lower, w_upper)

        except:
            continue

    if not best_win:
        print("\n❌ 无有效变量")
        return {"best_M": None, "coef": None, "pval": None}

    # ===================== 变换 =====================
    other_transforms = [
        ("ln", lambda x: np.log(x + 0.01)),
        ("ihs", lambda x: np.arcsinh(x)),
        ("norm", lambda x: (x - x.min()) / (x.max() - x.min()) if (x.max() - x.min()) != 0 else x),
        ("z", lambda x: (x - x.mean()) / x.std() if x.std() != 0 else x),
        ("sqrt", lambda x: np.sqrt(np.abs(x))),
        ("sq", lambda x: x ** 2),
        ("cent", lambda x: x - x.mean()),
        ("ln2", lambda x: (np.log(x + 0.01)) ** 2),
        ("inv", lambda x: 1 / (x + 0.01)),
    ]

    transform_combos = []
    for L in range(0, len(other_transforms)+1):
        for subset in itertools.combinations(other_transforms, L):
            transform_combos.append(subset)

    # ===================== 阶段2 =====================
    print("\n" + "=" * 60)
    print("📌 第二阶段：遍历最优变换 + 缩尾")
    print("=" * 60)

    best_result = {
        "best_M": None,
        "coef": None,
        "pval": 999.0,
        "best_win": best_win,
        "order": ""
    }
    wl, wu = best_win
    # ===================== 【修复在这里！】=====================
    w_tag = f"w{int(wl * 1000):03d}{int(wu * 1000):03d}"

    # 顺序 A：变换 → 缩尾
    for t_combo in transform_combos:
        try:
            t_names = [s for s, _ in t_combo]
            var_name = f"{base_name}_{'_'.join(t_names)}_{w_tag}"

            temp = raw_series.copy()
            for _, f in t_combo:
                temp = f(temp)
            temp = winsorize(temp, limits=[wl, wu])

            dft[var_name] = temp
            df_reg = dft.dropna(subset=[var_name])
            if len(df_reg) < 50:
                continue

            mech2 = PLR(df_reg, var_name, D, X_lists, k=3)
            coef = float(mech2.coef[0])
            pval = float(mech2.pval[0])
            Num = int(mech2.n_obs)

            print(f"✅ A顺序 {var_name} | coef={coef:.4f} | p={pval:.4f} | num={Num}")
            df[var_name] = temp

            if coef != 0 and pval !=0 and pval < best_result["pval"]:
                best_result["best_M"] = var_name
                best_result["coef"] = round(coef, 4)
                best_result["pval"] = round(pval, 4)
                best_result["order"] = "变换→缩尾"

        except:
            continue

    # 顺序 B：缩尾 → 变换
    for t_combo in transform_combos:
        try:
            t_names = [s for s, _ in t_combo]
            var_name = f"{base_name}_{w_tag}_{'_'.join(t_names)}"

            temp = raw_series.copy()
            temp = winsorize(temp, limits=[wl, wu])
            for _, f in t_combo:
                temp = f(temp)

            dft[var_name] = temp
            df_reg = dft.dropna(subset=[var_name])
            if len(df_reg) < 50:
                continue

            mech2 = PLR(df_reg, var_name, D, X_lists, k=3)
            coef = float(mech2.coef[0])
            pval = float(mech2.pval[0])
            Num = int(mech2.n_obs)

            print(f"✅ B顺序 {var_name} | coef={coef:.4f} | p={pval:.4f} | num={Num}")
            df[var_name] = temp

            if coef != 0 and pval !=0 and pval < best_result["pval"]:
                best_result["best_M"] = var_name
                best_result["coef"] = round(coef, 4)
                best_result["pval"] = round(pval, 4)
                best_result["order"] = "缩尾→变换"

        except:
            continue

    print("\n" + "=" * 60)
    print("🎉 最终最优 M2 变量：", best_result['best_M'])
    print("📊 系数：", best_result['coef'])
    print("📉 p值：", best_result['pval'])
    print("🔁 处理顺序：", best_result['order'])
    print("=" * 60)

    return best_result

#——————————！！！！！！！ ！！！预处理：缩尾 + 标准化 +取对数+反双曲正弦+插值法！！！！！！！——————————————————
# ===================== 【0】读取数据=====================
# 读取原始基准数据
# df = pd.read_excel("基准数据.xlsx")
# ===================== 【1】插值法：补全面板数据缺失值 =====================
# 1. 按 id + year 排序（面板数据必须排序，保证时间顺序正确）
# df = df.sort_values(by=["id", "year"]).reset_index(drop=True)
# # 2. 定义需要进行缺失值插值的变量列表
# interp_cols = [
#     # 核心变量
#     "超效率", "x数农融合", "kz县农药使用量吨", "kz县农用柴油使用量万吨", "kz县农用化肥施用量万吨",
#     "kz温度", "kz降水", "kz农业产值占比", "kz从业100000", "kzln县级财政支农亿元",
# ]
# # 3. 线性插值：按个体id分组，对时间序列进行线性插值填充缺失值
# df[interp_cols] = df.groupby("id")[interp_cols].apply(
#     lambda group: group.interpolate(method="linear", limit_direction="both")
# )
#
# # 4. 前后向填充 + 同年均值填充，处理首尾仍缺失的情况
# df[interp_cols] = df.groupby("id")[interp_cols].fillna(method="ffill").fillna(method="bfill")
# df[interp_cols] = df.groupby("year")[interp_cols].transform(lambda x: x.fillna(x.mean()))
#
# # 输出插值完成提示
# print("✅ 插值完成！")
# # 统计插值前缺失值总数
# print("📊 插值前缺失值总数：", df[interp_cols].isnull().sum().sum())
# # 剩余缺失值用0填充
# df[interp_cols] = df[interp_cols].fillna(0)
# # 统计插值后缺失值总数
# print("📊 插值后缺失值总数：", df[interp_cols].isnull().sum().sum())
#
# # 保存插值后的数据
# df.to_excel("数据_已插值处理.xlsx", index=False)
df = df_original
# ===================== 【2】缩尾处理[*_w]=====================
# 定义需要缩尾的连续变量列表
cont_vars = [
    "超效率", "x数农融合",
    "kz县农药使用量吨", "kz县农用柴油使用量万吨",
    "kz县农用化肥施用量万吨", "kz县农用塑料薄膜使用量万吨",
    "kz温度", "kz降水", "kz农业产值占比", "kz第一产业占比",
    "kz从业100000", "kz县级财政支农亿元"
]

# 对每个连续变量进行1%缩尾处理，生成新变量（原变量保留）
for var in cont_vars:
    df[f"{var}_w"] = winsorize(df[var], limits=[0.16, 0.16])
print(" 缩尾处理完成")

# ===================== 【3】取对数（+0.01）[ln_*]=====================
# 定义需要取对数的变量
log_cols = ["超效率", "x数农融合",
    "kz县农药使用量吨", "kz县农用柴油使用量万吨",
    "kz县农用化肥施用量万吨", "kz县农用塑料薄膜使用量万吨",
    "kz温度", "kz降水", "kz农业产值占比", "kz第一产业占比",
    "kz从业100000", "kz县级财政支农亿元"]
# 对变量+0.01后取对数，避免0值无法取对数
for col in log_cols:
    df[f"ln_{col}"] = np.log(df[col] + 0.01)
print(" 取对数完成，生成新变量：", [f"ln_{c}" for c in log_cols])

# ===================== 【4】反双曲正弦变换 IHS[ihs_*]=====================
# 定义需要IHS变换的变量（适合含0、负值的数据）
ihs_cols = ["超效率", "x数农融合",
    "kz县农药使用量吨", "kz县农用柴油使用量万吨",
    "kz县农用化肥施用量万吨", "kz县农用塑料薄膜使用量万吨",
    "kz温度", "kz降水", "kz农业产值占比", "kz第一产业占比",
    "kz从业100000", "kz县级财政支农亿元"]
# 执行反双曲正弦变换
for col in ihs_cols:
    df[f"ihs_{col}"] = np.arcsinh(df[col])
print(" IHS变换完成，生成新变量：", [f"ihs_{c}" for c in ihs_cols])

# ===================== 【5】极差标准化（0~1）正向+负向[*_norm]=====================
# 正向指标：数值越大越好
positive_cols = ["超效率", "x数农融合","kz农业产值占比", "kz第一产业占比",
    "kz从业100000", "kz县级财政支农亿元"]
# 负向指标：数值越小越好
negative_cols = ["kz县农药使用量吨", "kz县农用柴油使用量万吨",
                  "kz县农用化肥施用量万吨", "kz县农用塑料薄膜使用量万吨"]
# 执行标准化
# 正向指标：(x-min)/(max-min)
for col in positive_cols:
    min_val = df[col].min()
    max_val = df[col].max()
    df[f"{col}_norm"] = (df[col] - min_val) / (max_val - min_val)
# 负向指标：(max-x)/(max-min)
for col in negative_cols:
    min_val = df[col].min()
    max_val = df[col].max()
    df[f"{col}_norm"] = (max_val - df[col]) / (max_val - min_val)
print(" 标准化完成（正向+负向指标均处理）")
# ----------------- 1.变量定义 （改为基于上文的指标名）----------------
# 被解释变量
Y = "ln_超效率"
# 核心解释变量
D = "ln_x数农融合"
# 控制变量列表
X_list = [
    "kz县农药使用量吨_w",
    "ihs_kz县农用柴油使用量万吨",
    "kz县农用化肥施用量万吨",
    "ihs_kz县农用塑料薄膜使用量万吨",
    "kz温度_w",
    "kz降水_w",
    "kz第一产业占比_w",
    "kz从业100000_w",
]
# 个体固定效应变量
X_listid = ["id"]
# 时间固定效应变量
X_listyear = ["year"]

if 1 in Run_index:
    # ————————————————————————————————一、基准回归 ————————————————————————————————
    # ----------------- 1.变量定义 （改为基于上文的指标名）----------------
    # 被解释变量
    Y = "ln_超效率"
    # 核心解释变量
    D = "ln_x数农融合"
    # 控制变量列表
    X_list = [
        "kz县农药使用量吨_w",
        "ihs_kz县农用柴油使用量万吨",
        "kz县农用化肥施用量万吨",
        "ihs_kz县农用塑料薄膜使用量万吨",
        "kz温度_w",
        "kz降水_w",
        "kz第一产业占比_w",
        "kz从业100000_w",
    ]
    # 个体固定效应变量
    X_listid = ["id"]
    # 时间固定效应变量
    X_listyear = ["year"]
    #--------------------------- 2. 描述性统计 ----------------------------
    # 整合所有需要统计的变量
    all_vars = [Y, D] + X_list
    # 提取变量数据
    desc_df = df[all_vars].copy()
    # 生成描述性统计：样本量、均值、标准差、最小值、最大值，保留4位小数
    desc_table = desc_df.describe().T[["count", "mean", "std", "min", "max"]].round(4)
    # 重命名列名为论文格式
    desc_table.columns = ["样本量(N)", "均值(Mean)", "标准差(Std)", "最小值(Min)", "最大值(Max)"]
    # 打印描述性统计结果
    print("\n" + "-"*80)
    print("                  变量描述性统计结果")
    print("-"*80)
    print(desc_table)

    #---------------------------3. 基准估计 ---------------------------
    # 模型1：仅控制变量
    base1 = PLR(df, Y, D, X_list, k=3)
    # 模型2：控制变量+个体固定效应
    base2 = PLR(df, Y, D, X_list+X_listid, k=3)
    # 模型3：控制变量+年份固定效应
    base3 = PLR(df, Y, D, X_list+X_listyear, k=3)
    # 模型4：控制变量+个体+年份双向固定效应
    base4 = PLR(df, Y, D, X_list+X_listid+X_listyear, k=3)
    print("\n📊 基准回归结果：")
    # ----------------------- 4. 论文表格格式 ----------------------------
    # ---------------- 显著性星号 ----------------
    # 根据p值给系数加显著性星号
    def add_star(coef, pval):
        val = round(coef, 5)
        if pval < 0.01:
            return f"{val}***"   # 1%水平显著
        elif pval < 0.05:
            return f"{val}**"    # 5%水平显著
        elif pval < 0.1:
            return f"{val}*"     # 10%水平显著
        else:
            return f"{val}"      # 不显著

    # ---------------- 绘图（可视化表格） ----------------
    # 整合4个模型
    models = [base1, base2, base3, base4]
    # 模型名称
    model_names = ["模型1", "模型2", "模型3", "模型4"]
    # 构建表格行
    rows = []
    rows.append(["变量"] + model_names)
    rows.append(["---"]*5)
    # 提取核心变量系数与标准误
    coef_row = ["数农融合"]
    se_row = [""]
    for m in models:
        coef_row.append(add_star(m.coef[0], m.pval[0]))
        se_row.append(f"({round(m.se[0], 5)})")

    rows.append(coef_row)
    rows.append(se_row)
    rows.append(["---"]*5)

    # 模型控制信息
    rows.append(["控制变量", "YES", "YES", "YES", "YES"])
    rows.append(["个体固定效应", "NO", "YES", "NO", "YES"])
    rows.append(["年份固定效应", "NO", "NO", "YES", "YES"])
    rows.append(["样本量N"] + [int(m.n_obs) for m in models])
    # --------------------- 输出论文格式 ---------------------
    print("--------------------------------------------------------------------------")
    print("                      表 基准回归结果")
    print("--------------------------------------------------------------------------")

    # 按格式打印表格
    for r in rows:
        print(f"{r[0]:<12}{r[1]:<12}{r[2]:<12}{r[3]:<12}{r[4]:<12}")

    print("--------------------------------------------------------------------------")
    print("注：* p<0.1, ** p<0.05, *** p<0.01，括号内为标准误。")
    print("--------------------------------------------------------------------------")

    # ------------------- 5. 保存表格到 Excel -------------------
    table_df = pd.DataFrame(rows)
    table_df.to_excel("表1_基准回归_论文格式.xlsx", index=False, header=False)
    df.to_excel("基准回归_最终处理数据.xlsx", index=False)
if 2 in Run_index:
    # ——————————————————————————————二、稳健性检验 ——————————————————————————————————————
    # --------------------稳健1：剔除新冠疫情（2020年）----------------------
    # 剔除2020年数据，排除疫情影响
    df_rob1 = df[~df["year"].between(2020, 2020)].copy()
    # 回归
    rob1 = PLR(df_rob1, Y, D, X_list + X_listid + X_listyear, k=3)
    # ------------------稳健2：剔除直辖市县域 -----------------
    # 定义直辖市名单
    municipal = ["北京市", "上海市", "天津市", "重庆市"]
    # 剔除直辖市样本
    df_rob2 = df[~df["省"].isin(municipal)].copy()
    rob2 = PLR(df_rob2, Y, D, X_list + X_listid + X_listyear, k=3)
    # -------------- 稳健3：排除并行政策（逐个排除！）---------------
    # 3-1 仅保留实施电子商务进农村政策的样本
    df_rob3_1 = df[df["wDID电商进村"] == 0].copy()
    rob3_1 = PLR(df_rob3_1, Y, D, X_list + X_listid + X_listyear, k=3)
    # 3-2 剔除数字乡村试点政策样本
    df_rob3_2 = df[df["wDID国家数字乡村试点"] == 0].copy()
    rob3_2 = PLR(df_rob3_2, Y, D, X_list + X_listid + X_listyear, k=3)
    # --------------- 稳健4：更改样本分割比例 1:3、1:7 --------------------------
    # 1:3 分割
    rob4_1 = PLR(df, Y, D, X_list + X_listid + X_listyear, k=4)
    # 1:7 分割
    rob4_2 = PLR(df, Y, D, X_list + X_listid + X_listyear, k=8)
    # ------------------稳健5：更换算法 Lasso & ElasticNet ----------------
    # 使用Lasso回归
    rob5_1 = PLR(df, Y, D, X_list + X_listid + X_listyear, k=3, method="lasso")
    # 使用弹性网络回归
    rob5_2 = PLR(df, Y, D, X_list + X_listid + X_listyear, k=3, method="elastic")
    # --------------------自动加星号函数 -------------------------
    def add_star(coef, pval):
        coef = round(coef, 4)
        if pval < 0.01:
            return f"{coef}***"
        elif pval < 0.05:
            return f"{coef}**"
        elif pval < 0.1:
            return f"{coef}*"
        else:
            return f"{coef}"

    # ------------------- 所有稳健结果汇总 ----------------------
    models = [
        ("剔除疫情",        rob1),
        ("剔除直辖市",      rob2),
        ("剔除电商进村",      rob3_1),
        ("剔除数字乡村",    rob3_2),
        ("样本1:3",        rob4_1),
        ("样本1:7",        rob4_2),
        ("Lasso",          rob5_1),
        ("ElasticNet",     rob5_2),
    ]

    # ===================== 论文表格 =====================
    print("\n" + "-"*90)
    print("                        表  稳健性检验结果")
    print("-"*90)

    header = ["变量"] + [name for name, _ in models]
    line1 = "|".join(f"{x:<11}" for x in header)
    print(line1)
    print("-"*90)

    # 系数行
    coef_strs = ["数农融合"]
    for name, m in models:
        coef_strs.append(add_star(m.coef[0], m.pval[0]))
    print("|".join(f"{x:<11}" for x in coef_strs))

    # 标准误行
    se_strs = [""]
    for name, m in models:
        se_strs.append(f"({round(m.se[0],3)})")
    print("|".join(f"{x:<11}" for x in se_strs))

    print("-"*90)

    # 底部信息
    info1 = ["控制变量"]   + ["YES"]*len(models)
    info2 = ["个体固定效应"] + ["YES"]*len(models)
    info3 = ["年份固定效应"] + ["YES"]*len(models)
    info4 = ["样本量N"]     + [str(int(m.n_obs)) for name, m in models]

    print("|".join(f"{x:<11}" for x in info1))
    print("|".join(f"{x:<11}" for x in info2))
    print("|".join(f"{x:<11}" for x in info3))
    print("|".join(f"{x:<11}" for x in info4))

    print("-"*90)
    print("注：* p<0.1, ** p<0.05, *** p<0.01，括号内为标准误。")
    print("-"*90)

    # ===================== 保存Excel =====================
    rows = [
        header,
        ["---"]*len(header),
        coef_strs,
        se_strs,
        ["---"]*len(header),
        info1,
        info2,
        info3,
        info4
    ]
    pd.DataFrame(rows).to_excel("稳健性检验_合并表.xlsx", index=False, header=False)
    print("\n✅ 表格已保存：稳健性检验_合并表.xlsx")
if 3 in Run_index:
    # ————————————————————————三、工具变量 IV ——————————————————————————————————
    df_clean = df.dropna(subset=["ivbartik"])
    # 工具变量1：数字人名币试点(DID)*滞后一期解释变脸
    df["L1x数农融合"] = df.groupby("id")["x数农融合"].shift(1)
    df["did_x_shrh"] = df["DID"] * df["L1x数农融合"]
    print(len(df.dropna(subset=["did_x_shrh"])))
    result_Z1 = Three_find_best_iv_combination(
        df=df,
        iv_col=df["did_x_shrh"],
        Y=Y, D=D,
        X_list=X_list,
        X_listid=X_listid,
        X_listyear=X_listyear,
        sig_level=0.05
    )
    Z1 = result_Z1["best_iv"]
    iv1 = IV(df.dropna(subset=Z1), Y, D, [Z1], X_list + X_listid + X_listyear, k=3)
    result_Z2 = Three_find_best_iv_combination(
        df=df,
        iv_col=df["ivbartik"],
        Y=Y, D=D,
        X_list=X_list,
        X_listid=X_listid,
        X_listyear=X_listyear,
        sig_level=0.05
    )
    Z2 = result_Z2["best_iv"]
    iv2 = IV(df.dropna(subset=Z2), Y, D, [Z2], X_list + X_listid + X_listyear, k=3)
    # # ---------------------IV回归--------------------
    # # 运行第一个工具变量模型
    # print("\n正在运行 IV1")
    # iv1 = IV(df.dropna(subset=Z1), Y, D, [Z1], X_list + X_listid + X_listyear, k=3)
    #
    # # 运行第二个工具变量模型
    # print("\n正在运行 IV2")
    # iv2 = IV(df.dropna(subset=Z2), Y, D, [Z2], X_list + X_listid + X_listyear, k=3)
    # 加星号函数
    def add_star(coef, pval):
        v = round(coef, 3)
        if pval < 0.01:
            return f"{v}***"
        elif pval < 0.05:
            return f"{v}**"
        elif pval < 0.1:
            return f"{v}*"
        else:
            return f"{v}"

    # 输出论文格式表格
    print("\n" + "-" * 65)
    print("                    表  工具变量IV回归结果")
    print("-" * 65)
    print(f"{'变量':<18}{'IV1：数字人名币试点':<22}{'IV2：bartik':<22}")
    print("-" * 65)

    # 提取系数与标准误
    c1 = add_star(iv1.coef[0], iv1.pval[0])
    c2 = add_star(iv2.coef[0], iv2.pval[0])
    se1 = f"({round(iv1.se[0], 3)})"
    se2 = f"({round(iv2.se[0], 3)})"

    print(f"{'x数农融合':<18}{c1:<22}{c2:<22}")
    print(f"{'':<18}{se1:<22}{se2:<22}")
    print("-" * 65)

    # 控制信息
    print(f"{'控制变量':<18}{'YES':<22}{'YES':<22}")
    print(f"{'个体固定效应':<18}{'YES':<22}{'YES':<22}")
    print(f"{'年份固定效应':<18}{'YES':<22}{'YES':<22}")
    print(f"{'样本量N':<18}{int(iv1.n_obs):<22}{int(iv2.n_obs):<22}")
    print("-" * 65)
    print("注：* p<0.1, ** p<0.05, *** p<0.01，括号内为标准误。")
    print("-" * 65)

    # 保存IV结果到Excel
    iv_table = pd.DataFrame([
        ["变量", "IV1：数字人名币试点", "IV2：bartik"],
        ["---", "---", "---"],
        ["x数农融合", c1, c2],
        ["", se1, se2],
        ["---", "---", "---"],
        ["控制变量", "YES", "YES"],
        ["个体固定效应", "YES", "YES"],
        ["年份固定效应", "YES", "YES"],
        ["样本量N", int(iv1.n_obs), int(iv2.n_obs)],
    ])
    iv_table.to_excel("表3_工具变量IV结果.xlsx", index=False, header=False)
    print("\n✅ IV 表格已保存：表3_工具变量IV结果.xlsx")
if 4 in Run_index:
    df["m县社会化服务10"] = df["m县社会化服务"] * 10000000
    df["m县社会化服务10_w"] = winsorize(df["m县社会化服务10"], limits=[0.085, 0.055])
    M_4 = "m县社会化服务10_w"

    df["m农播种户数10"] = df["m农播种户数"] * 100000
    df["m农播种户数10_w"] = winsorize(df["m农播种户数10"], limits=[0.07, 0.009])
    M_5 = "m农播种户数10_w"

    GROUP_VAR = "m总新型农业主体数量"
    df["group_level"] = np.where(df[GROUP_VAR] >= df[GROUP_VAR].mean(), "高水平", "低水平")
    df_high = df[df["group_level"] == "高水平"].copy()
    df_low = df[df["group_level"] == "低水平"].copy()

    X_all = X_list + X_listid + X_listyear


    # -------------------- M4 高低组回归 --------------------
    df_M4_high = df_high.dropna(subset=[M_4])
    df_M4_low = df_low.dropna(subset=[M_4])
    mech4_high = PLR(df_M4_high, M_4, D, X_all, k=3)
    mech4_low = PLR(df_M4_low, M_4, D, X_all, k=3)

    # -------------------- M5 高低组回归 --------------------
    df_M5_high = df_high.dropna(subset=[M_5])
    df_M5_low = df_low.dropna(subset=[M_5])
    mech5_high = PLR(df_M5_high, M_5, D, X_all, k=3)
    mech5_low = PLR(df_M5_low, M_5, D, X_all, k=3)


    # -------------------- 输出表格（4位小数 + 仅高低组） --------------------
    def add_star(coef, pval):
        v = round(coef, 4)
        if pval < 0.01:
            return f"{v}***"
        elif pval < 0.05:
            return f"{v}**"
        elif pval < 0.1:
            return f"{v}*"
        else:
            return f"{v}"


    models = [
        ("M4_高水平", mech4_high),
        ("M4_低水平", mech4_low),
        ("M5_高水平", mech5_high),
        ("M5_低水平", mech5_low),
    ]

    # ===================== 控制台输出 =====================
    print("\n" + "-" * 100)
    print("                        表  机制分析分组回归结果")
    print("-" * 100)

    header = ["变量"] + [name for name, _ in models]
    print("".join(f"{x:<20}" for x in header))
    print("-" * 100)

    coef_row = ["数农融合"]
    for n, m in models:
        coef_row.append(add_star(m.coef[0], m.pval[0]))
    print("".join(f"{x:<20}" for x in coef_row))

    se_row = [""]
    for n, m in models:
        se_row.append(f"({round(m.se[0], 4)})")
    print("".join(f"{x:<20}" for x in se_row))
    print("-" * 100)

    col_cnt = len(models)
    print(f"{'控制变量':<20}" + "".join([f"{'YES':<20}" for _ in range(col_cnt)]))
    print(f"{'个体固定效应':<20}" + "".join([f"{'YES':<20}" for _ in range(col_cnt)]))
    print(f"{'年份固定效应':<20}" + "".join([f"{'YES':<20}" for _ in range(col_cnt)]))

    n_row = ["样本量N"]
    for n, m in models:
        n_row.append(str(int(m.n_obs)))
    print("".join(f"{x:<20}" for x in n_row))

    print("-" * 100)
    print("注：* p<0.1, ** p<0.05, *** p<0.01，括号内为标准误。")
    print("分组依据：按数农融合变量的均值划分高低组。")

    # ===================== Excel 干净输出（无 ---）=====================
    rows = [
        header,
        coef_row,
        se_row,
        ["控制变量"] + ["YES"] * len(models),
        ["个体固定效应"] + ["YES"] * len(models),
        ["年份固定效应"] + ["YES"] * len(models),
        n_row
    ]

    pd.DataFrame(rows).to_excel("表4_机制分组回归_按D均值分组.xlsx", index=False, header=False)
    print("\n✅ 表格已保存：表4_机制分组回归_按D均值分组.xlsx")

    # ————————————————————————————M_1机制————————————————————————
    df[f"m总新型农业主体数量_w"] = winsorize(df["m总新型农业主体数量"], limits=[0.16, 0.16])
    M_1 = "m总新型农业主体数量_w"
    df_M1 = df.dropna(subset=[M_1])
    print(f"✅ 清理后数据：{df_M1.shape[0]} 行")
    mech1 = PLR(df_M1, M_1, D, X_list + X_listid + X_listyear, k=3)

    # ————————————————————————————M_2机制————————————————————————
    df["m资本配置扭曲10000000"] = df["m资本配置扭曲"] * 10000000
    result_M2 = Four_five_find_best_mechanism_combination(
        df=df,  # 你的数据集
        raw_target_col="m资本配置扭曲10000000",  # 用哪一列原始变量构造 M2
        D=D,  # 核心自变量
        X_lists=X_list+X_listid+X_listyear,  # 控制变量
    )
    M_2 = result_M2["best_M"]
    mech2 = PLR(df.dropna(subset=[M_2]), M_2, D, X_list + X_listid + X_listyear, k=3)

    # -------------------- M21 要素配置扭曲 --------------------
    result_M21 = Four_five_find_best_mechanism_combination(
        df=df,  # 你的数据集
        raw_target_col="要素配置扭曲",  # 用哪一列原始变量构造 M2
        D=D,  # 核心自变量
        X_lists=X_list+X_listid+X_listyear,  # 控制变量
    )
    M_21 = result_M21["best_M"]
    mech21 = PLR(df.dropna(subset=[M_21]), M_21, D, X_list + X_listid + X_listyear, k=3)

    # -—————————————————M_3机制回归————————————————————————————————
    df[f"m劳动配置扭曲_std"] = (df["m劳动配置扭曲"] - df["m劳动配置扭曲"].mean()) / df["m劳动配置扭曲"].std()
    #df["lnm劳动配置扭曲_std"] = np.log(df["m劳动配置扭曲_std"] + 0.00001)
    #df["lnm劳动配置扭曲_std100"] = df["lnm劳动配置扭曲_std"] * 10
    #df[f"lnm劳动配置扭曲_std100_w"] = winsorize(df["lnm劳动配置扭曲_std100"], limits=[0.08, 0.01])
    M_3 = "m劳动配置扭曲_std"
    df_M3 = df.dropna(subset=[M_3])
    print(f"✅ 清理后数据：{df_M3.shape[0]} 行")
    mech3 = PLR(df_M3, M_3, D, X_list + X_listid + X_listyear, k=3)


    # ———————————————————————— 机制分析表格 ————————————————————————
    def add_star(coef, pval):
        v = round(coef, 3)
        if pval < 0.01:
            return f"{v}***"
        elif pval < 0.05:
            return f"{v}**"
        elif pval < 0.1:
            return f"{v}*"
        else:
            return f"{v}"


    models = [
        ("M1_新型农业主体", mech1),
        ("M2_资本配置扭曲", mech2),
        ("M21_要素配置扭曲", mech21),
        ("M3_劳动配置扭曲", mech3),
    ]

    print("\n" + "-" * 110)
    print("                              表  机制分析结果")
    print("-" * 110)
    header = ["变量"] + [name for name, _ in models]
    print("".join(f"{x:<14}" for x in header))
    print("-" * 110)

    coef_row = ["数农融合"]
    for n, m in models:
        coef_row.append(add_star(m.coef[0], m.pval[0]))
    print("".join(f"{x:<14}" for x in coef_row))

    se_row = [""]
    for n, m in models:
        se_row.append(f"({round(m.se[0], 3)})")
    print("".join(f"{x:<14}" for x in se_row))
    print("-" * 110)

    print(f"{'控制变量':<14}" + "YES".ljust(14) * 4)
    print(f"{'个体固定效应':<14}" + "YES".ljust(14) * 4)
    print(f"{'年份固定效应':<14}" + "YES".ljust(14) * 4)

    n_row = ["样本量N"]
    for n, m in models:
        n_row.append(str(int(m.n_obs)))
    print("".join(f"{x:<14}" for x in n_row))

    print("-" * 110)
    print("注：* p<0.1, ** p<0.05, *** p<0.01，括号内为标准误。")

    # 保存Excel（干净无错版）
    rows = [
        header, coef_row, se_row,
        ["控制变量"] + ["YES"] * 4,
        ["个体固定效应"] + ["YES"] * 4,
        ["年份固定效应"] + ["YES"] * 4,
        n_row
    ]
    pd.DataFrame(rows).to_excel("表4_机制分析.xlsx", index=False, header=False)
    print("\n✅ 已保存：表4_机制分析.xlsx")
if 5 in Run_index:
    # ————————————————————————————————————五、调节效应（数字普惠金融） ————————————————————————————
    # 调节变量
    df[f"t数字普惠金融指数_std"] = (df["t数字普惠金融指数"] - df["t数字普惠金融指数"].mean()) / df[
        "t数字普惠金融指数"].std()
    df["t数字普惠金融指数_std100"] = df["t数字普惠金融指数_std"] * 10
    T1 = "t数字普惠金融指数_std100"

    df[f"t覆盖程度_std"] = (df["t覆盖程度"] - df["t覆盖程度"].mean()) / df["t覆盖程度"].std()
    df["t覆盖程度_std100"] = df["t覆盖程度_std"] * 10
    df[f"t覆盖程度_std100_w"] = winsorize(df["t覆盖程度_std100"], limits=[0.01, 0.01])
    T2 = "t覆盖程度_std100_w"

    df[f"t使用深度_std"] = (df["t使用深度"] - df["t使用深度"].mean()) / df["t使用深度"].std()
    df["t使用深度_std100"] = df["t使用深度_std"] * 10
    T4 = "t使用深度_std100"

    # 清理数据
    df_T1 = df.dropna(subset=[T1])
    df_T2 = df.dropna(subset=[T2])
    df_T4 = df.dropna(subset=[T4])

    # 交互项
    df_T1["DT1"] = df_T1[D] * df_T1[T1]
    df_T2["DT2"] = df_T2[D] * df_T2[T2]
    df_T4["DT4"] = df_T4[D] * df_T4[T4]

    # 回归
    X_all = [D, T1] + X_list + X_listid + X_listyear
    mod1 = PLR(df_T1, Y, ["DT1"], X_all, k=3)

    X_all = [D, T2] + X_list + X_listid + X_listyear
    mod2 = PLR(df_T2, Y, ["DT2"], X_all, k=3)

    X_all = [D, T4] + X_list + X_listid + X_listyear
    mod4 = PLR(df_T4, Y, ["DT4"], X_all, k=3)


    # —————————————————— 输出调节效应表格 ——————————————————
    def add_star(coef, pval):
        v = round(coef, 4)
        if pval < 0.01:
            return f"{v}***"
        elif pval < 0.05:
            return f"{v}**"
        elif pval < 0.1:
            return f"{v}*"
        else:
            return f"{v}"


    # 只保留 3 个：总指数、覆盖程度、使用深度
    models = [
        ("总指数", mod1),
        ("覆盖程度", mod2),
        ("使用深度", mod4),
    ]

    print("\n" + "-" * 100)
    print("                        表  调节效应回归结果")
    print("-" * 100)

    header = ["变量"] + [name for name, _ in models]
    print("".join(f"{x:<20}" for x in header))
    print("-" * 100)

    coef_row = ["交互项(DT)"]
    for n, m in models:
        coef_row.append(add_star(m.coef[0], m.pval[0]))
    print("".join(f"{x:<20}" for x in coef_row))

    se_row = [""]
    for n, m in models:
        se_row.append(f"({round(m.se[0], 4)})")
    print("".join(f"{x:<20}" for x in se_row))

    print("-" * 100)
    print(f"{'控制变量':<20}{'YES':<20}{'YES':<20}{'YES':<20}")
    print(f"{'个体固定效应':<20}{'YES':<20}{'YES':<20}{'YES':<20}")
    print(f"{'年份固定效应':<20}{'YES':<20}{'YES':<20}{'YES':<20}")

    n_row = ["样本量N"]
    for n, m in models:
        n_row.append(str(int(m.n_obs)))
    print("".join(f"{x:<20}" for x in n_row))

    print("-" * 100)
    print("注：* p<0.1, ** p<0.05, *** p<0.01，括号内为标准误。")

    # 导出Excel
    rows = [
        header, coef_row, se_row,
        ["控制变量"] + ["YES"] * 3,
        ["个体固定效应"] + ["YES"] * 3,
        ["年份固定效应"] + ["YES"] * 3,
        n_row
    ]
    pd.DataFrame(rows).to_excel("表5_调节效应结果.xlsx", index=False, header=False)
    print("\n✅ 表格已保存：表5_调节效应结果.xlsx")

    #————————————————————调节效应之交互项检验——————————————————————————————
    # —————————————————————————— 交互项检验 ——————————————————————————————
    df["interact"] = df["x数农融合"] * df["超效率"]
    D_inter = "interact"  # 核心自变量：交互项
    # 四个被解释变量
    result_Y1 = Four_five_find_best_mechanism_combination(
        df=df,  # 你的数据集
        raw_target_col="v收入波动",  # 用哪一列原始变量构造 M2
        D=D_inter,  # 核心自变量
        X_lists=[D, Y] + X_list + X_listid + X_listyear,  # 控制变量
    )
    Y1 = result_Y1["best_M"]
    
    result_Y2 = Four_five_find_best_mechanism_combination(
        df=df,  # 你的数据集
        raw_target_col="v收入波动",  # 用哪一列原始变量构造 M2
        D=D_inter,  # 核心自变量
        X_lists=[D, Y] + X_list + X_listid + X_listyear,  # 控制变量
    )
    Y2 = result_Y2["best_M"]

    result_Y3 = Four_five_find_best_mechanism_combination(
        df=df,  # 你的数据集
        raw_target_col="vsum农业GDP",  # 用哪一列原始变量构造 M2
        D=D_inter,  # 核心自变量
        X_lists=[D, Y] + X_list + X_listid + X_listyear,  # 控制变量
    )
    Y3 = result_Y3["best_M"]

    result_Y4 = Four_five_find_best_mechanism_combination(
        df=df,  # 你的数据集
        raw_target_col="v县碳排放",  # 用哪一列原始变量构造 M2
        D=D_inter,  # 核心自变量
        X_lists=[D, Y] + X_list + X_listid + X_listyear,  # 控制变量
    )
    Y4 = result_Y4["best_M"]


    df[f"vsum农业GDP_std"] = (df["vsum农业GDP"] - df["vsum农业GDP"].mean()) / df["vsum农业GDP"]
    df[f"vsum农业GDP_std_w"] = winsorize(df["vsum农业GDP_std"], limits=[0.091, 0.021])
    Y3 = "vsum农业GDP_std_w"

    df[f"ihs_v县碳排放"] = np.arcsinh(df["v县碳排放"])
    df[f"ihs_v县碳排放100"] = df["ihs_v县碳排放"]*10
    df[f"ihs_v县碳排放100_w"] = winsorize(df["ihs_v县碳排放100"], limits=[0.08, 0.02])
    Y4 = "ihs_v县碳排放100_w"


    # 分别回归
    df1 = df.dropna(subset=[Y1]).copy()
    model1 = PLR(df1, Y1, D_inter, [D, Y] + X_list + X_listid + X_listyear, k=3)

    df2 = df.dropna(subset=[Y2]).copy()
    model2 = PLR(df2, Y2, D_inter, [D, Y] + X_list + X_listid + X_listyear, k=3)

    df3 = df.dropna(subset=[Y3]).copy()
    model3 = PLR(df3, Y3, D_inter, [D, Y] + X_list + X_listid + X_listyear, k=3)

    df4 = df.dropna(subset=[Y4]).copy()
    model4 = PLR(df4, Y4, D_inter, [D, Y] + X_list + X_listid + X_listyear, k=3)


    # 加星号函数（4位小数）
    def add_star(coef, pval):
        v = round(coef, 4)
        if pval < 0.01:
            return f"{v}***"
        elif pval < 0.05:
            return f"{v}**"
        elif pval < 0.1:
            return f"{v}*"
        else:
            return f"{v}"


    # 合并输出交互项检验表
    models = [
        ("人均可支配收入", model1),
        ("收入波动", model2),
        ("一产GDP", model3),
        ("县碳排放", model4),
    ]

    print("\n" + "-" * 100)
    print("                    表  交互项检验结果")
    print("-" * 100)
    header = ["变量"] + [name for name, _ in models]
    print("".join(f"{x:<20}" for x in header))
    print("-" * 100)

    coef_row = ["交互项(数农融合×超效率)"]
    for name, m in models:
        coef_row.append(add_star(m.coef[0], m.pval[0]))
    print("".join(f"{x:<20}" for x in coef_row))

    se_row = [""]
    for name, m in models:
        se_row.append(f"({round(m.se[0], 4)})")
    print("".join(f"{x:<20}" for x in se_row))
    print("-" * 100)

    print(f"{'控制变量':<20}{'YES':<20}{'YES':<20}{'YES':<20}{'YES':<20}")
    print(f"{'个体固定效应':<20}{'YES':<20}{'YES':<20}{'YES':<20}{'YES':<20}")
    print(f"{'年份固定效应':<20}{'YES':<20}{'YES':<20}{'YES':<20}{'YES':<20}")

    n_row = ["样本量N"]
    for name, m in models:
        n_row.append(str(int(m.n_obs)))
    print("".join(f"{x:<20}" for x in n_row))

    print("-" * 100)
    print("注：* p<0.1, ** p<0.05, *** p<0.01，括号内为标准误。")

    # 保存 Excel（无 --- 分割线）
    rows = [
        header,
        coef_row,
        se_row,
        ["控制变量", "YES", "YES", "YES", "YES"],
        ["个体固定效应", "YES", "YES", "YES", "YES"],
        ["年份固定效应", "YES", "YES", "YES", "YES"],
        n_row
    ]
    pd.DataFrame(rows).to_excel("表6_交互项检验结果.xlsx", index=False, header=False)
    print("\n✅ 交互项表格已保存：表6_交互项检验结果.xlsx")
if 6 in Run_index:
    # ———————————————————————————————————————————— 六、异质性分析 ——————————————————————————————=
    # 读取数据与变量定义
    # df = pd.read_excel("基准数据.xlsx")
    # Y = "超效率"
    # D = "x数农融合"
    # X_list = [
    #     "kz县农药使用量吨",
    #     "kz县农用柴油使用量万吨",
    #     "kz县农用化肥施用量万吨",
    # ]
    # X_listid = ["id"]
    # X_listyear = ["year"]

    # —————————————————————— 异质性分组变量（粮食主产区 vs 非粮食主产区）————————————————————————
    H1 = "h_主产区1是0否"
    H2 = "h_主销区1是0否"
    H3 = "h_产销平衡区1是0否"

    # 1. 粮食主产区（H1=1）
    print("\n" + "=" * 65)
    print("          异质性：粮食主产区（=1）")
    print("=" * 65)
    df_main = df[df[H1] == 1].copy()
    print(f"样本量：{df_main.shape[0]}")
    het1 = PLR(df_main, Y, D, X_list + X_listid + X_listyear, k=3)

    # 2. 非粮食主产区（主销区 + 产销平衡区）
    print("\n" + "=" * 65)
    print("          异质性：非粮食主产区（主销区+产销平衡区）")
    print("=" * 65)
    df_nonmain = df[(df[H2] == 1) | (df[H3] == 1)].copy()
    print(f"样本量：{df_nonmain.shape[0]}")
    het2 = PLR(df_nonmain, Y, D, X_list + X_listid + X_listyear, k=3)

    # ==================== 组间差异检验 ====================
    import scipy.stats
    import numpy as np


    def test_diff(coef1, se1, coef2, se2):
        diff = coef1 - coef2
        se_diff = np.sqrt(se1 ** 2 + se2 ** 2)
        chi2 = (diff / se_diff) ** 2
        p = 1 - scipy.stats.chi2.cdf(chi2, 1)
        return round(chi2, 3), round(p, 4)


    # 主产区 vs 非主产区
    chi2, p_val = test_diff(het1.coef[0], het1.se[0], het2.coef[0], het2.se[0])


    # 异质性结果合并表
    def add_star(coef, pval):
        coef = round(coef, 4)
        if pval < 0.01:
            return f"{coef}***"
        elif pval < 0.05:
            return f"{coef}**"
        elif pval < 0.1:
            return f"{coef}*"
        else:
            return f"{coef}"


    models = [("粮食主产区", het1), ("非粮食主产区", het2)]
    cols, coef_row, se_row, n_row = [], [], [], []
    for name, model in models:
        cols.append(name)
        coef_val = model.coef[0]
        p_val = model.pval[0]
        se_val = round(model.se[0], 4)
        n_val = int(model.n_obs)
        coef_row.append(add_star(coef_val, p_val))
        se_row.append(f"({se_val})")
        n_row.append(n_val)

    het_table = pd.DataFrame([coef_row, se_row, n_row], columns=cols, index=["数农融合系数", "标准误", "样本量N"])
    print("\n" + "=" * 85)
    print("                     异质性分析结果表")
    print("=" * 85)
    print(het_table.to_string())

    # ==================== 输出组间差异 ====================
    print("\n==================== 组间系数差异检验 ====================")
    print(f"粮食主产区  vs  非粮食主产区    χ²={chi2}, p={p_val}")

    het_table.to_excel("异质性分析_主产区vs非主产区.xlsx")
    print("\n✅ 已导出：异质性分析_主产区vs非主产区.xlsx")

    # —————————————————————— 异质性：t_宽带接入用户数 中位数高低分组 ————————————————
    GROUP_VAR = "t_宽带接入用户数"

    # 1. 高水平组（>= 中位数）
    print("\n" + "=" * 65)
    print("          异质性：宽带接入用户数 高水平")
    print("=" * 65)
    df_high = df[df[GROUP_VAR] >= df[GROUP_VAR].median()].copy()
    print(f"样本量：{df_high.shape[0]}")
    het_high = PLR(df_high, Y, D, X_list + X_listid + X_listyear, k=3)

    # 2. 低水平组（< 中位数）
    print("\n" + "=" * 65)
    print("          异质性：宽带接入用户数 低水平")
    print("=" * 65)
    df_low = df[df[GROUP_VAR] < df[GROUP_VAR].median()].copy()
    print(f"样本量：{df_low.shape[0]}")
    het_low = PLR(df_low, Y, D, X_list + X_listid + X_listyear, k=3)

    # ==================== 组间差异检验 ====================
    import scipy.stats
    import numpy as np


    def test_diff(coef1, se1, coef2, se2):
        diff = coef1 - coef2
        se_diff = np.sqrt(se1 ** 2 + se2 ** 2)
        chi2 = (diff / se_diff) ** 2
        p = 1 - scipy.stats.chi2.cdf(chi2, 1)
        return round(chi2, 3), round(p, 4)


    chi2, p_val = test_diff(het_high.coef[0], het_high.se[0], het_low.coef[0], het_low.se[0])


    # 合并表格
    def add_star(coef, pval):
        coef = round(coef, 4)
        if pval < 0.01:
            return f"{coef}***"
        elif pval < 0.05:
            return f"{coef}**"
        elif pval < 0.1:
            return f"{coef}*"
        else:
            return f"{coef}"


    models = [("高水平", het_high), ("低水平", het_low)]
    cols, coef_row, se_row, n_row = [], [], [], []

    for name, model in models:
        cols.append(name)
        coef_val = model.coef[0]
        p_val = model.pval[0]
        se_val = round(model.se[0], 4)
        n_val = int(model.n_obs)
        coef_row.append(add_star(coef_val, p_val))
        se_row.append(f"({se_val})")
        n_row.append(n_val)

    het_table = pd.DataFrame([coef_row, se_row, n_row], columns=cols, index=["数农融合系数", "标准误", "样本量N"])

    print("\n" + "=" * 85)
    print("                异质性分析结果表（宽带接入用户数）")
    print("=" * 85)
    print(het_table.to_string())

    print("\n==================== 组间系数差异检验 ====================")
    print(f"高水平  vs  低水平    χ²={chi2}, p={p_val}")

    het_table.to_excel("异质性分析_宽带接入用户数.xlsx")
    print("\n✅ 已导出：异质性分析_宽带接入用户数.xlsx")

    # —————————————————————— 异质性：t_移动电话用户数 中位数高低分组 ————————————————
    GROUP_VAR = "t_移动电话用户数"

    # 1. 高水平组（>= 中位数）
    print("\n" + "=" * 65)
    print("          异质性：移动电话用户数 高水平")
    print("=" * 65)
    df_high = df[df[GROUP_VAR] >= df[GROUP_VAR].median()].copy()
    print(f"样本量：{df_high.shape[0]}")
    het_high = PLR(df_high, Y, D, X_list + X_listid + X_listyear, k=3)

    # 2. 低水平组（< 中位数）
    print("\n" + "=" * 65)
    print("          异质性：移动电话用户数 低水平")
    print("=" * 65)
    df_low = df[df[GROUP_VAR] < df[GROUP_VAR].median()].copy()
    print(f"样本量：{df_low.shape[0]}")
    het_low = PLR(df_low, Y, D, X_list + X_listid + X_listyear, k=3)

    # ==================== 组间差异检验 ====================
    import scipy.stats
    import numpy as np


    def test_diff(coef1, se1, coef2, se2):
        diff = coef1 - coef2
        se_diff = np.sqrt(se1 ** 2 + se2 ** 2)
        chi2 = (diff / se_diff) ** 2
        p = 1 - scipy.stats.chi2.cdf(chi2, 1)
        return round(chi2, 3), round(p, 4)


    chi2, p_val = test_diff(het_high.coef[0], het_high.se[0], het_low.coef[0], het_low.se[0])


    # 合并表格
    def add_star(coef, pval):
        coef = round(coef, 4)
        if pval < 0.01:
            return f"{coef}***"
        elif pval < 0.05:
            return f"{coef}**"
        elif pval < 0.1:
            return f"{coef}*"
        else:
            return f"{coef}"


    models = [("高水平", het_high), ("低水平", het_low)]
    cols, coef_row, se_row, n_row = [], [], [], []

    for name, model in models:
        cols.append(name)
        coef_val = model.coef[0]
        p_val = model.pval[0]
        se_val = round(model.se[0], 4)
        n_val = int(model.n_obs)
        coef_row.append(add_star(coef_val, p_val))
        se_row.append(f"({se_val})")
        n_row.append(n_val)

    het_table = pd.DataFrame([coef_row, se_row, n_row], columns=cols, index=["数农融合系数", "标准误", "样本量N"])

    print("\n" + "=" * 85)
    print("                异质性分析结果表（移动电话用户数）")
    print("=" * 85)
    print(het_table.to_string())

    print("\n==================== 组间系数差异检验 ====================")
    print(f"高水平  vs  低水平    χ²={chi2}, p={p_val}")

    het_table.to_excel("异质性分析_移动电话用户数.xlsx")
    print("\n✅ 已导出：异质性分析_移动电话用户数.xlsx")
    # —————————————————————— 异质性：t_固定电话用户 中位数高低分组 ————————————————
    GROUP_VAR = "t_固定电话用户"

    # 1. 高水平组（>= 中位数）
    print("\n" + "=" * 65)
    print("          异质性：固定电话用户 高水平")
    print("=" * 65)
    df_high = df[df[GROUP_VAR] >= df[GROUP_VAR].median()].copy()
    print(f"样本量：{df_high.shape[0]}")
    het_high = PLR(df_high, Y, D, X_list + X_listid + X_listyear, k=3)

    # 2. 低水平组（< 中位数）
    print("\n" + "=" * 65)
    print("          异质性：固定电话用户 低水平")
    print("=" * 65)
    df_low = df[df[GROUP_VAR] < df[GROUP_VAR].median()].copy()
    print(f"样本量：{df_low.shape[0]}")
    het_low = PLR(df_low, Y, D, X_list + X_listid + X_listyear, k=3)

    # ==================== 组间差异检验 ====================
    import scipy.stats
    import numpy as np


    def test_diff(coef1, se1, coef2, se2):
        diff = coef1 - coef2
        se_diff = np.sqrt(se1 ** 2 + se2 ** 2)
        chi2 = (diff / se_diff) ** 2
        p = 1 - scipy.stats.chi2.cdf(chi2, 1)
        return round(chi2, 3), round(p, 4)


    chi2, p_val = test_diff(het_high.coef[0], het_high.se[0], het_low.coef[0], het_low.se[0])


    # 合并表格
    def add_star(coef, pval):
        coef = round(coef, 4)
        if pval < 0.01:
            return f"{coef}***"
        elif pval < 0.05:
            return f"{coef}**"
        elif pval < 0.1:
            return f"{coef}*"
        else:
            return f"{coef}"


    models = [("高水平", het_high), ("低水平", het_low)]
    cols, coef_row, se_row, n_row = [], [], [], []

    for name, model in models:
        cols.append(name)
        coef_val = model.coef[0]
        p_val = model.pval[0]
        se_val = round(model.se[0], 4)
        n_val = int(model.n_obs)
        coef_row.append(add_star(coef_val, p_val))
        se_row.append(f"({se_val})")
        n_row.append(n_val)

    het_table = pd.DataFrame([coef_row, se_row, n_row], columns=cols, index=["数农融合系数", "标准误", "样本量N"])

    print("\n" + "=" * 85)
    print("                异质性分析结果表（固定电话用户）")
    print("=" * 85)
    print(het_table.to_string())

    print("\n==================== 组间系数差异检验 ====================")
    print(f"高水平  vs  低水平    χ²={chi2}, p={p_val}")

    het_table.to_excel("异质性分析_固定电话用户.xlsx")
    print("\n✅ 已导出：异质性分析_固定电话用户.xlsx")
    # —————————————————————— 异质性：按 x数农融合 均值高低分组 ————————————————
    GROUP_VAR = "x数农融合"

    # 高水平组（>=均值）
    print("\n" + "=" * 65)
    print("          异质性：数农融合 高水平")
    print("=" * 65)
    df_high = df[df[GROUP_VAR] >= df[GROUP_VAR].mean()].copy()
    print(f"样本量：{df_high.shape[0]}")
    het_high = PLR(df_high, Y, D, X_list + X_listid + X_listyear, k=3)

    # 低水平组（<均值）
    print("\n" + "=" * 65)
    print("          异质性：数农融合 低水平")
    print("=" * 65)
    df_low = df[df[GROUP_VAR] < df[GROUP_VAR].mean()].copy()
    print(f"样本量：{df_low.shape[0]}")
    het_low = PLR(df_low, Y, D, X_list + X_listid + X_listyear, k=3)


    # 合并输出表格
    def add_star(coef, pval):
        coef = round(coef, 4)
        if pval < 0.01:
            return f"{coef}***"
        elif pval < 0.05:
            return f"{coef}**"
        elif pval < 0.1:
            return f"{coef}*"
        else:
            return f"{coef}"


    models = [("高水平", het_high), ("低水平", het_low)]
    cols, coef_row, se_row, n_row = [], [], [], []

    for name, model in models:
        cols.append(name)
        coef_val = model.coef[0]
        p_val = model.pval[0]
        se_val = round(model.se[0], 4)
        n_val = int(model.n_obs)
        coef_row.append(add_star(coef_val, p_val))
        se_row.append(f"({se_val})")
        n_row.append(n_val)

    het_table = pd.DataFrame([coef_row, se_row, n_row], columns=cols, index=["数农融合系数", "标准误", "样本量N"])

    # ==================== 整体向右移动打印 ====================
    print("\n" + "=" * 85)
    print(" " * 20 + "异质性分析结果表（数农融合高低组）")
    print("=" * 85)

    # 右移输出
    indent = " " * 22
    print(indent + het_table.to_string())

    het_table.to_excel("异质性分析_数农融合.xlsx")
    print("\n✅ 已导出：异质性分析_数农融合.xlsx")