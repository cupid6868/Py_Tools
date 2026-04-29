import pandas as pd
import numpy as np

# ===================== 路径 =====================
input_path = r"C:\Users\Lenovo.DESKTOP-PATTUAQ\Desktop\极端天气_县级年度统计.xlsx"
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
            if ny > max_year: continue
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

# ===================== 🖨️ 打印信息 =====================
print("="*80)
county_count = df_expanded.groupby(["省份", "地级", "县级"]).ngroups
print(f"✅ 总县级数量：{county_count} 个")
print("="*80)

# ===================== 【第一步 A → A1 ✅ 修复完毕】 =====================
import pandas as pd
import numpy as np

# ===================== 路径 =====================
input_path = r"C:\Users\Lenovo.DESKTOP-PATTUAQ\Desktop\极端天气_县级年度统计.xlsx"
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

    for col in fixed_cols:
        a1_ori = g[col].values.copy()
        out = a1_ori.copy()

        # 前3年：只看 a1_ori，绝不看out
        if n >= 1:
            out[0] = 0
        if n >= 2:
            if a1_ori[1] == 1 and a1_ori[0] == 0:
                out[1] = 0
        if n >= 3:
            if a1_ori[2] == 1 and a1_ori[0] == 0 and a1_ori[1] == 0:
                out[2] = 0

        # 超过3年：只看 a1_ori，彻底防污染
        for i in range(base, n):
            start = max(0, i - lag_years)
            past_ori = a1_ori[start:i]
            curr_ori = a1_ori[i]

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

df_A1 = df_expanded.groupby(["省份", "地级", "县级"], group_keys=False).apply(make_A1)

# ===================== 【导出第一阶段 A1】 =====================
df_A1.to_excel(output_A1, index=False)
print(f"✅ 第一阶段 A1 已导出：{output_A1}")

# ===================== 【第二步 A1 → A2】 =====================
def make_A2(group):
    g = group.copy().reset_index(drop=True)
    n = len(g)
    base = 3

    for col in fixed_cols:
        a1 = g[col].values.copy()
        out = a1.copy()

        # 前3年严格清零
        if n >= 1:
            out[0] = 0
        if n >= 2:
            if out[1] == 1 and out[0] == 0:
                out[1] = 0
        if n >= 3:
            if out[2] == 1 and out[0] == 0 and out[1] == 0:
                out[2] = 0

        # 超过前3年
        for i in range(base, n):
            start = max(0, i - lag_years)
            past = out[start:i]
            if np.all(past == 0) and out[i] == 1:
                out[i] = 0

        g[col] = out
    return g

df_final = df_A1.groupby(["省份", "地级", "县级"], group_keys=False).apply(make_A2)

# ===================== 导出最终结果 =====================
df_final.to_excel(output_A2, index=False)
print(f"✅ 最终阶段 A2 已导出：{output_A2}")
print("\n🎉 全部完成！")