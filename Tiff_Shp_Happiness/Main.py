import os
import glob
import rasterio
import geopandas as gpd
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from rasterio.mask import mask
from shapely.geometry import mapping, box
from matplotlib.colors import LinearSegmentedColormap
import time

# 设置中文字体
plt.rcParams["font.family"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题

# 设置环境变量修复可能缺失的.shx文件
os.environ["SHAPE_RESTORE_SHX"] = "YES"


def calculate_township_suitability(tiff_path, townships, admin_field_mapping,
                                   shp_filename, output_folder, visualize=True):
    """处理单个TIFF文件与矢量数据，包含矢量字段输出功能"""
    start_time = time.time()
    tiff_filename = os.path.basename(tiff_path).split('.')[0]
    dynamic_fields = list(admin_field_mapping.keys())

    # 1. 输出矢量数据包含的所有字段及使用情况
    vector_fields = townships.columns.tolist()
    print(f"\n{'=' * 60}")
    print(f"【{shp_filename}】矢量数据字段信息")
    print(f"{'=' * 60}")
    print(f"共检测到 {len(vector_fields)} 个字段：")
    for i, field in enumerate(vector_fields, 1):
        # 获取字段数据类型
        field_type = str(townships[field].dtype)
        # 检查是否为当前ST_Class使用的字段
        is_used = field in admin_field_mapping.values()
        usage_mark = "✅ 已使用" if is_used else "  未使用"
        print(f"   {i:2d}. 字段名: {field:<15} 类型: {field_type:<10} {usage_mark}")

    # 检查缺失字段
    missing_fields = []
    for level, field in admin_field_mapping.items():
        if field and field not in vector_fields:
            missing_fields.append(f"{level}（字段名：{field}）")

    if missing_fields:
        print(f"\n❌ 以下必要字段不存在：")
        for mf in missing_fields:
            print(f"   - {mf}")
        return None

    # 2. 读取TIFF数据并处理
    print(f"\n{'=' * 60}")
    print(f"处理文件：SHP={shp_filename} | TIFF={tiff_filename}")
    print(f"{'=' * 60}")
    with rasterio.open(tiff_path) as src:
        tiff_crs = src.crs
        tiff_bounds = src.bounds
        tiff_extent = box(*tiff_bounds)
        print(f"✅ TIFF坐标系: {tiff_crs}")

        # 处理矢量坐标系
        if townships.crs is None:
            print("⚠️  矢量无坐标系，默认设为EPSG:4326")
            townships = townships.set_crs("EPSG:4326")
        print(f"✅ 矢量坐标系: {townships.crs}")

        # 坐标系转换
        if townships.crs != tiff_crs:
            print(f"🔄 转换矢量坐标系至 {tiff_crs}")
            try:
                townships = townships.to_crs(tiff_crs)
                print("✅ 坐标系转换完成")
            except Exception as e:
                print(f"❌ 坐标系转换失败: {str(e)}")
                return None

        # 检查空间重叠
        townships_extent = box(*townships.total_bounds)
        if not tiff_extent.intersects(townships_extent):
            print("⚠️  矢量与TIFF无空间重叠，跳过")
            return None

        # 3. 提取信息并计算适宜性
        total_townships = len(townships)
        results = []
        valid_indices = []

        for idx, row in townships.iterrows():
            if idx % 100 == 0:
                progress = (idx / total_townships) * 100
                print(f"🔄 进度: {progress:.1f}% ({idx}/{total_townships})")

            try:
                # 提取行政信息
                admin_info = {}
                for level, field in admin_field_mapping.items():
                    if field:
                        value = row[field] if field in row.index else "未知"
                        admin_info[level] = str(value) if pd.notna(value) and str(value).strip() != '' else "未知"
                    else:
                        admin_info[level] = "未知"

                # 空间检查
                row_geom = row['geometry']
                if not row_geom.intersects(tiff_extent):
                    continue

                # 计算均值
                geom = [mapping(row_geom)]
                out_image, _ = mask(src, geom, crop=True)
                nodata = src.nodata

                if nodata is not None:
                    values = out_image[out_image != nodata]
                else:
                    values = out_image.flatten()

                if len(values) > 0 and not np.all(np.isnan(values)):
                    suitability_mean = round(np.nanmean(values), 4)
                    results.append({
                        **admin_info,
                        "适宜性均值": suitability_mean
                    })
                    valid_indices.append(idx)

            except Exception as e:
                print(f"❌ 处理第{idx}个单元时出错: {str(e)}")

    # 4. 保存结果（添加序号列）
    if not results:
        print(f"❌ 无有效结果，不保存CSV")
        return None

    result_df = pd.DataFrame(results)
    column_order = dynamic_fields + ["适宜性均值"]
    result_df = result_df[column_order]
    result_df.insert(0, "序号", range(1, len(result_df) + 1))  # 添加序号列

    # 生成CSV路径
    csv_path = os.path.join(
        output_folder,
        f"admin_suitability_{shp_filename}_{tiff_filename}.csv"
    )
    result_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"✅ CSV保存路径：{csv_path}")
    print(f"✅ 有效数据行数：{len(result_df)} | 字段：{result_df.columns.tolist()}")

    # 5. 可视化
    if visualize:
        try:
            valid_geometries = townships.loc[valid_indices, 'geometry'].reset_index(drop=True)
            plot_gdf = gpd.GeoDataFrame(
                result_df.drop(columns=['序号']),
                geometry=valid_geometries,
                crs=townships.crs
            )
            visualize_suitability(plot_gdf, shp_filename, tiff_filename, output_folder)
        except Exception as e:
            print(f"❌ 可视化失败: {str(e)}")

    return result_df


def visualize_suitability(gdf, shp_name, tiff_name, output_folder):
    """可视化适宜性分布"""
    colors = ['#f7fcf5', '#e5f5e0', '#c7e9c0', '#a1d99b', '#74c476', '#41ab5d', '#238b45', '#006d2c']
    cmap = LinearSegmentedColormap.from_list('suitability_cmap', colors, N=100)

    fig, ax = plt.subplots(figsize=(16, 12))
    vmin = gdf["适宜性均值"].min()
    vmax = gdf["适宜性均值"].max()

    plot = gdf.plot(
        column="适宜性均值",
        cmap=cmap,
        linewidth=0.3,
        edgecolor='#999999',
        ax=ax,
        legend=False,
        vmin=vmin,
        vmax=vmax
    )

    cbar = plt.colorbar(plot.collections[0], ax=ax, orientation="horizontal",
                        shrink=0.8, pad=0.05, aspect=50)
    cbar.set_label("适宜性均值", fontsize=14, labelpad=10)

    ax.set_title(f'{shp_name} - {tiff_name} 适宜性分布', fontsize=18, pad=20)
    ax.axis('off')

    stats_text = (
        f"数据概况：\n"
        f"单元数：{len(gdf)} 个\n"
        f"均值：{gdf['适宜性均值'].mean():.4f}\n"
        f"范围：{vmin:.4f} ~ {vmax:.4f}"
    )
    plt.text(0.02, 0.02, stats_text, transform=ax.transAxes,
             bbox=dict(facecolor='white', alpha=0.9, edgecolor='#dddddd'),
             fontsize=12, verticalalignment='bottom')

    plt.tight_layout()
    png_path = os.path.join(
        output_folder,
        f"admin_suitability_map_{shp_name}_{tiff_name}.png"
    )
    plt.savefig(png_path, dpi=300, bbox_inches='tight')
    print(f"✅ 地图保存为：{png_path}")
    plt.close()


if __name__ == "__main__":
    # 超参数设置
    ST_Class = "Sheng_Frame"  # 可选：Xian_Frame / Shi_Frame / Sheng_Frame
    TIFF_FOLDER = "./Data/"
    OUTPUT_FOLDER = "./results/"

    # 根据ST_Class设置路径和字段映射
    if ST_Class == "Xian_Frame":
        SHP_PATH = f"./{ST_Class}/"
        admin_field_mapping = {
            "省级类": "省级类",
            "省级": "省级",
            "地级类": "地级类",
            "地级": "地级",
            "县级类": "县级类",
            "县级": "县级",
            "地名": "地名"
        }
    elif ST_Class == "Shi_Frame":
        SHP_PATH = f"./{ST_Class}/"
        admin_field_mapping = {
            "省级类": "省级类",
            "省级": "省级",
            "地级类": "地级类",
            "地级": "地级"
        }
    elif ST_Class == "Sheng_Frame":
        SHP_PATH = f"./{ST_Class}/"
        admin_field_mapping = {
            "省类型": "省类型",
            "省": "省"
        }
    else:
        print(f"❌ 无效的ST_Class值：{ST_Class}")
        exit(1)

    # 创建结果文件夹
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # 读取shp文件
    try:
        if os.path.isdir(SHP_PATH):
            shp_files = glob.glob(os.path.join(SHP_PATH, "*.shp"))
            if not shp_files:
                raise FileNotFoundError(f"在 {SHP_PATH} 中未找到shp文件")
            shp_path = shp_files[0]
        else:
            shp_path = SHP_PATH

        shp_filename = os.path.basename(shp_path).split('.')[0]
        townships = gpd.read_file(shp_path)
        print(f"✅ 成功读取SHP文件：{shp_filename}（{len(townships)}个单元）")
        print(f"✅ 当前ST_Class：{ST_Class}，使用字段：{list(admin_field_mapping.keys())}")
    except Exception as e:
        print(f"❌ 读取SHP文件失败: {str(e)}")
        exit(1)

    # 处理所有tif文件
    tif_files = glob.glob(os.path.join(TIFF_FOLDER, "*.tif"))
    if not tif_files:
        print(f"❌ 在 {TIFF_FOLDER} 中未找到tif文件")
        exit(1)

    print(f"\n📋 找到{len(tif_files)}个tif文件，开始处理...")
    for tif_path in tif_files:
        calculate_township_suitability(
            tiff_path=tif_path,
            townships=townships,
            admin_field_mapping=admin_field_mapping,
            shp_filename=shp_filename,
            output_folder=OUTPUT_FOLDER,
            visualize=True
        )

    print(f"\n🎉 所有文件处理完成，结果保存在：{OUTPUT_FOLDER}")
