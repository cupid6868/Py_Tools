import geopandas as gpd
import pandas as pd
import matplotlib.pyplot as plt

# -------------------------- 1. 文件路径配置（直接使用你的路径） --------------------------
county_shp_path = r"D:\Data_2\2023年县级\县级.shp"  # 县级行政区划shp
station_shp_path = r"C:\Users\18218\Desktop\PythonProject\Ming-route-data\Ming_Stations_2016.shp"  # 明朝驿站shp
output_csv_path = r"县域明朝驿站数量统计.csv"  # 输出CSV路径
output_png_path = r"县域驿站分布地图.png"  # 输出图片路径

# -------------------------- 2. 读取地理数据 --------------------------
# 读取县级行政区划数据
county_gdf = gpd.read_file(county_shp_path, encoding='utf-8')
# 读取明朝驿站数据
station_gdf = gpd.read_file(station_shp_path, encoding='utf-8')

# 统一坐标系（关键：必须保证两个数据坐标系一致才能空间匹配）
if county_gdf.crs != station_gdf.crs:
    station_gdf = station_gdf.to_crs(county_gdf.crs)

# -------------------------- 3. 空间连接：统计每个县的驿站数量 --------------------------
# 空间连接：判断每个驿站属于哪个县
joined_gdf = gpd.sjoin(station_gdf, county_gdf, how='left', predicate='within')

# 按县级字段分组统计驿站数量（使用第一个非几何字段作为县唯一标识，适配通用shp）
county_id_col = county_gdf.columns[0]  # 县唯一标识字段
station_count = joined_gdf.groupby(county_id_col).size().reset_index(name='驿站数量')

# 合并统计结果到县级地理数据
result_gdf = county_gdf.merge(station_count, on=county_id_col, how='left')
# 无驿站的县填充为0
result_gdf['驿站数量'] = result_gdf['驿站数量'].fillna(0).astype(int)

# -------------------------- 4. 输出CSV文件 --------------------------
# 提取属性数据（不含几何信息）输出CSV
csv_data = result_gdf.drop(columns='geometry')
csv_data.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
print(f"✅ 统计完成！CSV文件已保存至：{output_csv_path}")

# -------------------------- 5. 绘制地图（标注驿站位置） --------------------------
plt.rcParams['font.sans-serif'] = ['SimHei']  # 解决中文显示
plt.rcParams['axes.unicode_minus'] = False    # 解决负号显示

fig, ax = plt.subplots(figsize=(15, 12))

# 绘制县级行政区划（按驿站数量填色）
result_gdf.plot(
    ax=ax,
    column='驿站数量',
    cmap='Blues',        # 蓝色系配色
    edgecolor='black',   # 县边界颜色
    linewidth=0.5,       # 边界线宽
    legend=True,         # 显示图例
    legend_kwds={'label': '明朝驿站数量', 'orientation': 'vertical'}
)

# 绘制驿站点位（红色圆点突出显示）
station_gdf.plot(
    ax=ax,
    color='red',
    markersize=1,
    marker='o',
    label='明朝驿站'
)

# 为每个驿站标注名称（如果有名称字段）
if 'name' in station_gdf.columns:
    for x, y, label in zip(station_gdf.geometry.x, station_gdf.geometry.y, station_gdf['name']):
        ax.annotate(label, xy=(x, y), xytext=(3, 3), textcoords='offset points', fontsize=8, color='darkred')

# 地图样式设置
ax.set_title('明朝驿站县域分布统计地图', fontsize=16, pad=20)
ax.legend(loc='upper right')
ax.set_axis_off()  # 关闭坐标轴

# 保存图片
plt.tight_layout()
plt.savefig(output_png_path, dpi=300, bbox_inches='tight')
plt.close()

print(f"✅ 地图绘制完成！图片已保存至：{output_png_path}")
print(f"📊 统计概况：共{len(result_gdf)}个县，{len(station_gdf)}个明朝驿站")