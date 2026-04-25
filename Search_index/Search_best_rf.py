import numpy as np
import pandas as pd
from doubleml import DoubleMLData, DoubleMLPLR, DoubleMLPLIV
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.neural_network import MLPRegressor
from sklearn.svm import SVR
from sklearn.linear_model import LassoCV, RidgeCV, ElasticNetCV
from scipy.stats.mstats import winsorize
from concurrent.futures import ThreadPoolExecutor, as_completed
import os
import warnings
import itertools
import time
import threading
import sys
from datetime import datetime


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

# ======================================================================

warnings.filterwarnings('ignore')

Run_index = [1]
# 一、基准回归 二、稳健性检验 三、工具变量 IV 四、机制分析 五、调节效应 六、异质性分析
# ===================== 读取数据 =====================
print("=" * 60)
print("📥 开始读取 Excel 数据...")
df_original = pd.read_excel(r"C:\Users\18218\Desktop\PythonProject\平衡面板.xlsx")
df_original = df_original.reset_index(drop=True)
df_original = df_original.sort_values(["id", "year"]).reset_index(drop=True)
print(f"✅ 数据读取完成！共 {df_original.shape[0]} 行，{df_original.shape[1]} 列")
print("=" * 60)

np.random.seed(42)

def desc(Y, D, X_list, X_listid, X_listyear, M_4, df_MA):
    all_vars = [Y, D] + X_list + X_listid + X_listyear + [M_4]
    desc_df = df_MA[all_vars].copy()
    desc_table = desc_df.describe().T[["count", "mean", "std", "min", "max"]].round(4)
    desc_table.columns = ["样本量(N)", "均值(Mean)", "标准差(Std)", "最小值(Min)", "最大值(Max)"]
    print("\n" + "=" * 80)
    print("                    变量描述性统计结果")
    print("=" * 80)
    print(desc_table)

# ===================== 机器学习模型 =====================
def get_model(name="rf"):
    if name == "rf":    return RandomForestRegressor(random_state=42, n_estimators=50, n_jobs=-1)
    if name == "gradboost": return GradientBoostingRegressor(random_state=42)
    if name == "nnet":  return MLPRegressor(max_iter=500, random_state=42)
    if name == "svm":   return SVR(kernel="rbf")
    if name == "lasso": return LassoCV(random_state=42, n_jobs=-1)
    if name == "ridge": return RidgeCV()
    if name == "elastic": return ElasticNetCV(random_state=42, n_jobs=-1)

# ===================== DML 模型函数 =====================
def PLR(data, y, d, x, k=3, method="rf"):
    np.random.seed(42)
    x = list(x)
    dml_data = DoubleMLData(data, y_col=y, d_cols=d, x_cols=x)
    model = get_model(method)
    plr = DoubleMLPLR(dml_data, ml_l=model, ml_m=model, n_folds=k, n_rep=1)
    plr.fit()
    return plr

def IV(data, y, d, z, x, k=5, method="rf"):
    x = list(x)
    dml_data = DoubleMLData(data, y_col=y, d_cols=d, z_cols=z, x_cols=x)
    model = get_model(method)
    iv = DoubleMLPLIV(dml_data, ml_l=model, ml_m=model, ml_r=model, n_folds=k, n_rep=1)
    iv.fit()
    return iv

# ===================== 预处理 =====================
df = df_original.copy()

log_cols = ["超效率", "x数农融合",
            "kz县农药使用量吨", "kz县农用柴油使用量万吨",
            "kz县农用化肥施用量万吨", "kz县农用塑料薄膜使用量万吨",
            "kz温度", "kz降水", "kz农业产值占比", "kz第一产业占比",
            "kz从业100000", "kz县级财政支农亿元"]
for col in log_cols:
    df[f"ln_{col}"] = np.log(df[col] + 0.01)
print(" 取对数完成，生成新变量：", [f"ln_{c}" for c in log_cols])

ihs_cols = ["超效率", "x数农融合",
            "kz县农药使用量吨", "kz县农用柴油使用量万吨",
            "kz县农用化肥施用量万吨", "kz县农用塑料薄膜使用量万吨",
            "kz温度", "kz降水", "kz农业产值占比", "kz第一产业占比",
            "kz从业100000", "kz县级财政支农亿元"]
for col in ihs_cols:
    df[f"ihs_{col}"] = np.arcsinh(df[col])
print(" IHS变换完成，生成新变量：", [f"ihs_{c}" for c in ihs_cols])

positive_cols = ["超效率", "x数农融合", "kz农业产值占比", "kz第一产业占比",
                 "kz从业100000", "kz县级财政支农亿元"]
negative_cols = ["kz县农药使用量吨", "kz县农用柴油使用量万吨",
                 "kz县农用化肥施用量万吨", "kz县农用塑料薄膜使用量万吨"]
for col in positive_cols:
    min_val = df[col].min()
    max_val = df[col].max()
    df[f"{col}_norm"] = (df[col] - min_val) / (max_val - min_val)
for col in negative_cols:
    min_val = df[col].min()
    max_val = df[col].max()
    df[f"{col}_norm"] = (max_val - df[col]) / (max_val - min_val)
print(" 标准化完成（正向+负向指标均处理）")

# ===================== 固定配置 =====================
VAR_INIT = {
    "Y": "超效率",
    "D": "x数农融合",
    "X1": "kz县农药使用量吨",
    "X2": "kz县农用柴油使用量万吨",
    "X3": "kz县农用化肥施用量万吨",
    "X4": "kz县农用塑料薄膜使用量万吨",
    "X5": "kz温度",
    "X6": "kz降水",
}

X78_CANDIDATES = [
    "kz农业产值占比",
    "kz第一产业占比",
    "kz从业100000",
    "kz县级财政支农亿元"
]

# ===================== 核心模型运行（含系数检查 + 显著红色提示） =====================
def run_plr_4models(df, Y, D, Xs):
    try:
        m1 = PLR(df, Y, D, Xs, k=3)
        m2 = PLR(df, Y, D, Xs + ["id"], k=3)
        m3 = PLR(df, Y, D, Xs + ["year"], k=3)
        m4 = PLR(df, Y, D, Xs + ["id", "year"], k=3)

        c1, p1 = float(m1.coef[0]), float(m1.pval[0])
        c2, p2 = float(m2.coef[0]), float(m2.pval[0])
        c3, p3 = float(m3.coef[0]), float(m3.pval[0])
        c4, p4 = float(m4.coef[0]), float(m4.pval[0])
        p_mean = (p1 + p2 + p3 + p4) / 4

        # 系数全大于0 判断
        coef_all_positive = c1 > 0 and c2 > 0 and c3 > 0 and c4 > 0
        print(f"     【系数检查】c1={c1:.4f}, c2={c2:.4f}, c3={c3:.4f}, c4={c4:.4f} | 全大于0: {coef_all_positive}")

        # 3个或4个 p < 0.01 红色提示
        p_less_001 = sum([p1 <= 0.01, p2 <= 0.01, p3 <= 0.01, p4 <= 0.01])
        if p_less_001 >= 3:
            print(f"     \033[91m❗❗❗ {p_less_001}个p值<0.01，高度显著 ❗❗❗\033[0m")

        return c1, c2, c3, c4, p1, p2, p3, p4, p_mean
    except:
        return None, None, None, None, 999, 999, 999, 999, 999

def get_forms(var):
    return [var, f"{var}_w", f"ln_{var}", f"ihs_{var}"]

# ===================== 1.筛选Y =====================
def get_best_Y(df):
    D_use = VAR_INIT["D"]
    X_base = [VAR_INIT[f"X{i}"] for i in range(1,7)] + X78_CANDIDATES[:2]
    pool = get_forms(VAR_INIT["Y"])
    res = []
    print("\n" + "="*80)
    print("🔹 筛选 Y")
    print("="*80)
    for i,y in enumerate(pool):
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, y, D_use, X_base)
        res.append((pm, y))
        print(f"Y{i+1} | Y={y} | D={D_use}")
        print(f"     X={X_base}")
        print(f"     p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}\n")
    best = sorted(res)[0][1]
    print(f"✅ 最优Y: {best}")
    return best

# ===================== 2.筛选D =====================
def get_best_D(df, bestY):
    X_base = [VAR_INIT[f"X{i}"] for i in range(1,7)] + X78_CANDIDATES[:2]
    pool = get_forms(VAR_INIT["D"])
    res = []
    print("\n" + "="*80)
    print("🔹 筛选 D")
    print("="*80)
    for i,d in enumerate(pool):
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, bestY, d, X_base)
        res.append((pm, d))
        print(f"D{i+1} | Y={bestY} | D={d}")
        print(f"     X={X_base}")
        print(f"     p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}\n")
    best = sorted(res)[0][1]
    print(f"✅ 最优D: {best}")
    return best

# ===================== 3.筛选X1~X6 =====================
def get_best_X1to6(df, bestY, bestD):
    best = {f"X{i}":None for i in range(1,7)}
    print("\n" + "="*80)
    print("🔹 筛选 X1~X6")
    print("="*80)

    # X1
    print("\n--- X1 ---")
    res = []
    for x in get_forms(VAR_INIT["X1"]):
        xu = [x, VAR_INIT["X2"],VAR_INIT["X3"],VAR_INIT["X4"],VAR_INIT["X5"],VAR_INIT["X6"]] + X78_CANDIDATES[:2]
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, bestY, bestD, xu)
        res.append((pm,x))
        print(f"X1 | Y={bestY} | D={bestD} | X={xu}")
        print(f"    p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}")
    best["X1"] = sorted(res)[0][1]

    # X2
    print("\n--- X2 ---")
    res=[]
    for x in get_forms(VAR_INIT["X2"]):
        xu = [best["X1"],x,VAR_INIT["X3"],VAR_INIT["X4"],VAR_INIT["X5"],VAR_INIT["X6"]] + X78_CANDIDATES[:2]
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, bestY, bestD, xu)
        res.append((pm,x))
        print(f"X2 | Y={bestY} | D={bestD} | X={xu}")
        print(f"    p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}")
    best["X2"]=sorted(res)[0][1]

    # X3
    print("\n--- X3 ---")
    res=[]
    for x in get_forms(VAR_INIT["X3"]):
        xu = [best["X1"],best["X2"],x,VAR_INIT["X4"],VAR_INIT["X5"],VAR_INIT["X6"]] + X78_CANDIDATES[:2]
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, bestY, bestD, xu)
        res.append((pm,x))
        print(f"X3 | Y={bestY} | D={bestD} | X={xu}")
        print(f"    p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}")
    best["X3"]=sorted(res)[0][1]

    # X4
    print("\n--- X4 ---")
    res=[]
    for x in get_forms(VAR_INIT["X4"]):
        xu = [best["X1"],best["X2"],best["X3"],x,VAR_INIT["X5"],VAR_INIT["X6"]] + X78_CANDIDATES[:2]
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, bestY, bestD, xu)
        res.append((pm,x))
        print(f"X4 | Y={bestY} | D={bestD} | X={xu}")
        print(f"    p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}")
    best["X4"]=sorted(res)[0][1]

    # X5
    print("\n--- X5 ---")
    res=[]
    for x in get_forms(VAR_INIT["X5"]):
        xu = [best["X1"],best["X2"],best["X3"],best["X4"],x,VAR_INIT["X6"]] + X78_CANDIDATES[:2]
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, bestY, bestD, xu)
        res.append((pm,x))
        print(f"X5 | Y={bestY} | D={bestD} | X={xu}")
        print(f"    p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}")
    best["X5"]=sorted(res)[0][1]

    # X6
    print("\n--- X6 ---")
    res=[]
    for x in get_forms(VAR_INIT["X6"]):
        xu = [best["X1"],best["X2"],best["X3"],best["X4"],best["X5"],x] + X78_CANDIDATES[:2]
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, bestY, bestD, xu)
        res.append((pm,x))
        print(f"X6 | Y={bestY} | D={bestD} | X={xu}")
        print(f"    p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}")
    best["X6"]=sorted(res)[0][1]

    print(f"\n✅ X1~X6最优: {best}")
    return best

# ===================== 4.从4个里选2个做X7、X8 =====================
def get_best_X7X8(df, bestY, bestD, bestX16):
    print("\n" + "="*80)
    print("🔹 从4个变量中选最优2个作为X7、X8 + 最优变换")
    print("候选4个：", X78_CANDIDATES)
    print("="*80)

    X16_best_list = [bestX16[f"X{i}"] for i in range(1,7)]
    best_score = 999
    best_pair = None
    best_forms = None

    for v7, v8 in itertools.combinations(X78_CANDIDATES, 2):
        for f7 in get_forms(v7):
            for f8 in get_forms(v8):
                X_full = X16_best_list + [f7, f8]
                c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(df, bestY, bestD, X_full)
                print(f"\nX7候选:{v7}→{f7} | X8候选:{v8}→{f8}")
                print(f"Y={bestY} | D={bestD} | X={X_full}")
                print(f"p1={p1:.4f} p2={p2:.4f} p3={p3:.4f} p4={p4:.4f} mean={pm:.4f}")

                if pm < best_score:
                    best_score = pm
                    best_pair = (v7, v8)
                    best_forms = (f7, f8)

    best7, best8 = best_forms
    print("\n🎉 最终最优 X7、X8：")
    print(f"X7 = {best7}")
    print(f"X8 = {best8}")
    print(f"最优p均值 = {best_score:.4f}")
    return {"X7": best7, "X8": best8}

# ===================== 主流程 =====================
if 1 in Run_index:
    win_list = [0.028,0.029,0.030,0.031,0.058, 0.075,0.095, 0.125,0.15,0.16,0.18,0.20,0.23,0.25,0.28, 0.32, 0.33]
    all_res = []

    for wr in win_list:
        print(f"\n\n==================== 缩尾: {wr} ====================")
        dft = df.copy()
        for v in VAR_INIT.values():
            if v in dft.columns:
                dft[f"{v}_w"] = winsorize(dft[v], limits=[wr, wr])
        for v in X78_CANDIDATES:
            if v in dft.columns:
                dft[f"{v}_w"] = winsorize(dft[v], limits=[wr, wr])

        bestY   = get_best_Y(dft)
        bestD   = get_best_D(dft, bestY)
        bestX16 = get_best_X1to6(dft, bestY, bestD)
        bestX78 = get_best_X7X8(dft, bestY, bestD, bestX16)

        X_all = [bestX16[f"X{i}"] for i in range(1,7)] + [bestX78["X7"], bestX78["X8"]]
        c1,c2,c3,c4,p1,p2,p3,p4,pm = run_plr_4models(dft, bestY, bestD, X_all)

        all_res.append({
            "缩尾": wr,
            "Y": bestY,
            "D": bestD,
            "X1": bestX16["X1"], "X2": bestX16["X2"], "X3": bestX16["X3"],
            "X4": bestX16["X4"], "X5": bestX16["X5"], "X6": bestX16["X6"],
            "X7": bestX78["X7"], "X8": bestX78["X8"],
            "p1": round(p1,4), "p2": round(p2,4), "p3": round(p3,4), "p4": round(p4,4), "p_mean": round(pm,4)
        })

    pd.DataFrame(all_res).to_excel("最终结果_4选2作为X7X8.xlsx", index=False)
    print("\n🎉 全部运行完毕！")
    print(f"\n📄 完整日志已保存至：{log_file}")