import pandas as pd
import numpy as np

# ===================== 1. 文件路径配置 =====================
weather_path = r"C:\Users\Lenovo.DESKTOP-PATTUAQ\Desktop\极端天气事件0428.xlsx"
area_path = r"E:\Test_Code\2023年县级_县级数据.xlsx"
output_final_path = r"C:\Users\Lenovo.DESKTOP-PATTUAQ\Desktop\极端天气_县级年度统计.xlsx"

# ===================== 2. 读取数据 =====================
df_weather = pd.read_excel(weather_path, sheet_name="Ori")
df_area = pd.read_excel(area_path).drop_duplicates()

# ===================== 3. 四大直辖市：地级 = 省级名称 =====================
municipalities = ["北京市", "上海市", "天津市", "重庆市"]
df_area.loc[df_area["省级"].isin(municipalities), "地级"] = df_area["省级"]

# ===================== 4. 构建行政区划映射 =====================
# 省份 → 该省所有地级
province_to_city = {k: v.tolist() for k, v in df_area.groupby("省级")["地级"].unique().items()}

# 地级 → 该地级下所有县级（县+区）
city_to_all_county = {k: v.tolist() for k, v in df_area.groupby("地级")["县级"].unique().items()}

# 地级 → 该地级下仅区（不含县）
city_to_only_district = {}
df_area_filter = df_area[(df_area["县级"] == "市辖区") | (df_area["县级"].str.endswith("市区"))]
for k, v in df_area_filter.groupby("地级")["县级"].unique().items():
    city_to_only_district[k] = v.tolist()

# ===================== 5. 步骤1：处理 地级=1 → 拆分为该省所有地级 =====================
def split_city_row(row):
    if pd.isna(row["地级"]) or row["地级"] != 1:
        return [row.to_dict()]

    province = row["省份"]
    city_list = province_to_city.get(province, [])

    if len(city_list) == 0:
        return [row.to_dict()]

    new_rows = []
    base = row.to_dict()
    for city in city_list:
        new_row = base.copy()
        new_row["地级"] = city
        new_rows.append(new_row)
    return new_rows

# 执行地级拆分
step1_rows = []
for _, row in df_weather.iterrows():
    step1_rows.extend(split_city_row(row))

# ===================== 6. 步骤2：处理县级拆分 =====================
def split_county_row(row):
    county_val = row["县级"]
    city = row["地级"]

    if pd.isna(county_val) or pd.isna(city):
        return [row.to_dict()]

    # 规则1：县级 == 1 → 取全部县级
    if county_val == 1:
        county_list = city_to_all_county.get(city, [])
        if len(county_list) == 0:
            return [row.to_dict()]

        new_rows = []
        base = row.to_dict()
        for cty in county_list:
            new_row = base.copy()
            new_row["县级"] = cty
            new_rows.append(new_row)
        return new_rows

    # 规则2：县级 == 市辖区 / **市区 → 只取区
    elif county_val == "市辖区" or str(county_val).endswith("市区"):
        district_list = city_to_only_district.get(city, [])
        if len(district_list) == 0:
            return [row.to_dict()]

        new_rows = []
        base = row.to_dict()
        for dist in district_list:
            new_row = base.copy()
            new_row["县级"] = dist
            new_rows.append(new_row)
        return new_rows

    # 不拆分
    else:
        return [row.to_dict()]

# 执行县级拆分
final_rows = []
for row_dict in step1_rows:
    row = pd.Series(row_dict)
    final_rows.extend(split_county_row(row))

df_split = pd.DataFrame(final_rows)

# ===================== 7. 灾害类型编码处理 =====================
code_map = {
    "1": "洪水",
    "2": "干旱",
    "3": "热浪与寒潮",
    "4": "风暴",
    "5": "野火",
    "6": "滑坡"
}
dis_order = ["洪水", "干旱", "热浪与寒潮", "风暴", "野火", "滑坡"]

# 处理灾害类型字段
df_split["灾害类型_str"] = df_split["灾害类型"].astype(str).str.strip()
df_split["灾害数字"] = df_split["灾害类型_str"].str.split("-").str[-1]
df_split = df_split[df_split["灾害数字"].isin(code_map.keys())].copy()
df_split["灾害名称"] = df_split["灾害数字"].map(code_map)

# ===================== 8. 按 年份+省份+地级+县级 年度统计（0/1） =====================
pivot = df_split.groupby(["年份", "省份", "地级", "县级", "灾害名称"]).size().unstack(fill_value=0)

# 确保6类灾害列完整
for col in dis_order:
    if col not in pivot.columns:
        pivot[col] = 0

pivot = pivot[dis_order]
pivot = pivot.where(pivot == 0, 1).astype(int)  # 有记录=1，无=0
result = pivot.reset_index()

# ===================== 9. 导出最终结果 =====================
result.to_excel(output_final_path, index=False)

# ===================== 10. 输出日志 =====================
print("✅ 全部处理完成！")
print(f"原始天气记录数：{len(df_weather)}")
print(f"拆分后明细记录数：{len(df_split)}")
print(f"县级年度统计结果：{len(result)} 条")
print(f"文件已保存至：{output_final_path}")