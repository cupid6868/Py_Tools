import pandas as pd
import numpy as np

# ===================== 路径 =====================
input_path = r"C:\Users\Lenovo.DESKTOP-PATTUAQ\Desktop\极端天气_县级年度统计.xlsx"
output_expanded = r"C:\Users\Lenovo.DESKTOP-PATTUAQ\Desktop\年份扩展后原始数据.xlsx"  # 新增输出
output_A1 = r"C:\Users\Lenovo.DESKTOP-PATTUAQ\Desktop\A1_第一阶段_正向传染.xlsx"
output_A2 = r"C:\Users\Lenovo.DESKTOP-PATTUAQ\Desktop\A2_第二阶段_最终结果.xlsx"

fixed_cols = ["洪水", "干旱", "热浪与寒潮", "风暴", "野火", "滑坡"]
max_year = 2025
lag_years = 3

# ===================== 1. 读取数据 =====================
df = pd.read_excel(input_path)
df = df.drop_duplicates(subset=["省份", "地级", "县级", "年份"])
df = df.sort_values(["省份", "地级", "县级", "年份"]).reset_index(drop=True)

# ===================== 2. 年份扩展 =====================
def expand_county(group):
    group = group.sort_values("年份").reset_index(drop=True)
    info = group.iloc[0][["省份", "地级", "县级"]].to_dict()
    exist_years = set(group["年份"])
    rows = []

    for _, row in group.iterrows():
        rows.append(row.to_dict())

    for _, row in group.iterrows():
        base_year = row["年份"]
        for i in range(1, lag_years + 1):
            ny = base_year + i
            if ny > max_year:
                continue
            if ny not in exist_years:
                new_row = info.copy()
                new_row["年份"] = ny
                for c in fixed_cols:
                    new_row[c] = 0
                rows.append(new_row)
                exist_years.add(ny)
    return pd.DataFrame(rows)

df_expanded = df.groupby(["省份", "地级", "县级"], group_keys=False).apply(expand_county)
df_expanded = df_expanded.drop_duplicates(subset=["省份", "地级", "县级", "年份"])
df_expanded = df_expanded.sort_values(["省份", "地级", "县级", "年份"]).reset_index(drop=True)

# ===================== 🆕 新增：输出年份扩展后的表格 =====================
df_expanded.to_excel(output_expanded, index=False)
print(f"✅ 已输出【年份扩展后完整表格】：{output_expanded}")
print("=" * 90)

# ===================== 打印核心信息 =====================
print("=" * 90)
county_total = df_expanded.groupby(["省份", "地级", "县级"]).ngroups
print(f"📊 总计县级数量：{county_total} 个")
print(f"📋 原始总记录数：{len(df)} 条")
print(f"📋 年份扩展后记录数：{len(df_expanded)} 条")
print("=" * 90)
print("【各县级 原始年份 / 扩展后年份】")
print("=" * 90)

for name, group in df_expanded.groupby(["省份", "地级", "县级"]):
    pro, city, county = name
    ori_y = sorted(df[(df["省份"] == pro) & (df["地级"] == city) & (df["县级"] == county)]["年份"].tolist())
    new_y = sorted(group["年份"].tolist())
    print(f"• {pro}-{city}-{county:<15} | 原始：{ori_y} | 扩展：{new_y}")

print("=" * 90)
print("▶️ 开始执行第一阶段：A → A1 正向传染处理")

# ===================== 【A1：无污染版】 =====================
def make_A1(group):
    g = group.copy().reset_index(drop=True)
    n = len(g)
    base = 3

    for col in fixed_cols:
        ori = g[col].values.copy()
        out = ori.copy()

        # 前3年
        flag = 0
        for i in range(min(base, n)):
            if ori[i] == 1:
                flag = 1
            if flag:
                out[i] = 1

        # 超过3年：只看原始ori，防污染
        for i in range(base, n):
            start = max(0, i - lag_years)
            past_ori = ori[start:i]
            if np.any(past_ori == 1):
                out[i] = 1

        g[col] = out
    return g

df_A1 = df_expanded.groupby(["省份", "地级", "县级"], group_keys=False).apply(make_A1)
df_A1.to_excel(output_A1, index=False)
print(f"✅ 第一阶段完成，已输出A1文件：{output_A1}")

# ===================== 【A2：100% 无污染版】 =====================
print("▶️ 开始执行第二阶段：A1 → A2 基准清零处理")

def make_A2(group):
    g = group.copy().reset_index(drop=True)
    n = len(g)
    base = 3

    # 拿到当前县级的“年份扩展后原始数据”（未经过A1传染）
    key = tuple(g[["省份", "地级", "县级"]].iloc[0])
    orig_group = df_expanded[
        (df_expanded["省份"] == key[0]) &
        (df_expanded["地级"] == key[1]) &
        (df_expanded["县级"] == key[2])
    ].sort_values("年份")

    for col in fixed_cols:
        a1_data = g[col].values.copy()
        out = a1_data.copy()

        # ✅ 判断依据：年份扩展后的原始数据（从df_expanded来）
        expanded_ori = orig_group[col].values.copy()

        # 前3年基准规则
        if n >= 1:
            out[0] = 0
        if n >= 2:
            if expanded_ori[1] == 1 and expanded_ori[0] == 0:
                out[1] = 0
        if n >= 3:
            if expanded_ori[2] == 1 and expanded_ori[0] == 0 and expanded_ori[1] == 0:
                out[2] = 0

        # 超过3年规则
        for i in range(base, n):
            start = max(0, i - lag_years)
            past_ori = expanded_ori[start:i]
            curr_ori = expanded_ori[i]

            if np.all(past_ori == 0) and curr_ori == 1:
                out[i] = 0

        g[col] = out
    return g

df_final = df_A1.groupby(["省份", "地级", "县级"], group_keys=False).apply(make_A2)
df_final.to_excel(output_A2, index=False)

# ===================== 结束打印 =====================
print(f"✅ 第二阶段完成，已输出A2最终文件：{output_A2}")
print("=" * 90)
print("🎉 全部处理完毕 —— A1、A2 均 100% 无污染、无漏洞")
print("=" * 90)