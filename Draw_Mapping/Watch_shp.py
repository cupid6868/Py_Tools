import geopandas as gpd
import matplotlib.pyplot as plt

# ---------------------- 1. 读取SHP文件 ----------------------
shp_path = r"E:\Test_Code\中国气候区划\Climate_quhua.shp"
gdf = gpd.read_file(shp_path)

# ---------------------- 2. 自动列出所有列名（控制台显示） ----------------------
print("="*50)
print("✅ SHP文件包含的【所有列名】如下：")
print("="*50)
columns = gdf.columns.tolist()
for idx, col in enumerate(columns):
    print(f"  {idx} → {col}")
print("="*50)

# ---------------------- 3. 控制台输入数字选择列名 ----------------------
while True:
    try:
        choice = int(input("请输入你要用来着色的列的【数字编号】："))
        if 0 <= choice < len(columns):
            color_column = columns[choice]
            break
        else:
            print(f"❌ 请输入 0 ~ {len(columns)-1} 之间的数字！")
    except ValueError:
        print("❌ 请输入有效数字！")

print(f"\n✅ 已选择列：【{color_column}】，即将开始着色...")

# ---------------------- 4. 绘图：按列值不同颜色填充 ----------------------
plt.rcParams["font.sans-serif"] = ["SimHei"]  # 正常显示中文
plt.rcParams["axes.unicode_minus"] = False    # 正常显示负号

fig, ax = plt.subplots(figsize=(12, 10))

# 核心：按选中的列的值分类着色
gdf.plot(
    ax=ax,
    column=color_column,      # 按选中列着色
    cmap="tab10",             # 配色方案
    edgecolor="black",        # 边界黑色
    linewidth=0.6,            # 边界粗细
    legend=True,              # 显示图例
    legend_kwds={
        "loc": "upper left",
        "bbox_to_anchor": (1, 1)
    }
)

# ---------------------- 5. 图表美化 ----------------------
ax.set_title(f"中国东中西三大区域 - 按列【{color_column}】着色", fontsize=16, pad=20)
ax.set_axis_off()  # 关闭坐标轴
plt.tight_layout()

# 保存高清图片到同一目录
save_path = r"E:\Test_Code\区域着色结果图-climate.png"
plt.savefig(save_path, dpi=300, bbox_inches="tight")
# plt.show()

print(f"\n🎉 绘图完成！图片已保存到：\n{save_path}")