import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import warnings
import os
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import cairosvg
import io
from PIL import Image
from pypinyin import lazy_pinyin
from shapely.errors import TopologicalError
from shapely.geometry import Point
from translate import Translator

warnings.filterwarnings('ignore')


# ==============================================
# 【超参数区：所有路径、样式、尺寸统一配置】
# ==============================================
class Config:
    # --------------------- 【文件路径配置】 ---------------------
    EXCEL_PATH = r"E:\Test_Code\平衡面板.xlsx"
    SHP_PATH = r"E:\Test_Code\2023年县级\县级.shp"
    BORDER_PATH = r"E:\Test_Code\国界\Export_Output.shp"
    COASTLINE_PATH = r"E:\Test_Code\海岸线\Export_Output_2.shp"
    SVG_COMPASS_PATH = r"E:\Test_Code\指北针.svg"
    AGRI_ZONE_PATH = r"E:\Test_Code\中国东中西三大区域分布qu3\qu-sheng.shp"

    # --------------------- 【年份与绘图设置】 ---------------------
    TARGET_YEAR = 2023

    # --------------------- 【字体配置】 ---------------------
    FONT_NAME = "Times New Roman"

    # --------------------- 【颜色与边界配置】 ---------------------
    NULL_COLOR = "white"
    BORDER_COLOR = "black"
    BORDER_WIDTH = 0.1
    COAST_COLOR = "#0A93FC"
    NINE_LINE_COLOR = "red"
    NINE_LINE_WIDTH = 1.2

    # --------------------- 【出图参数】 ---------------------
    SAVE_PATH = r'E:\Test_Code\2023_Climate_Zone_Map.svg'
    DPI = 900
    FIG_SIZE = (14, 12)

    # --------------------- 【地图元素参数】 ---------------------
    COMPASS_ZOOM = 0.45
    GRID_LABEL_SIZE = 18
    LEGEND_FONT_SIZE = 18
    SCALE_TEXT_SIZE = 16

    # --------------------- 【分区配色方案】 ---------------------
    COLORS = [
        '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
        '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
        '#a8e6cf', '#ffb3b3', '#d4a5a5', '#66b3ff', '#99ff99',
        '#ff99cc', '#c9c9ff', '#ffcc99', '#ff6666', '#66ff66',
        '#6666ff', '#ff66ff', '#ffff66', '#66ffff', '#ff6600',
        '#00ff66', '#0066ff', '#6600ff', '#ff0066', '#00ffcc',
        '#ff00cc', '#cc00ff', '#cccc00', '#00cccc', '#333333',
        '#666666', '#999999', '#cccccc', '#00033', '#003300'
    ]


cfg = Config()


# ==============================================
# 自动编码读取 shp
# ==============================================
def read_shp_auto_encoding(path):
    encodings = ['utf-8', 'gbk', 'gb18030', 'gb2312']
    for enc in encodings:
        try:
            return gpd.read_file(path, encoding=enc)
        except (UnicodeDecodeError, LookupError):
            continue
    raise ValueError("所有常见编码均无法读取该 shp 文件")


# ==============================================
# 统一所有SHP为 WGS 坐标系 EPSG:4326
# ==============================================
def align_crs(*gdfs, target_crs="EPSG:4326"):
    aligned = []
    for gdf in gdfs:
        try:
            gdf = gdf.to_crs(target_crs)
        except:
            pass
        aligned.append(gdf)
    return aligned


def load_svg_compass(svg_path):
    svg_png = cairosvg.svg2png(url=svg_path)
    return Image.open(io.BytesIO(svg_png))


def auto_detect_text_col(gdf, col_type="county"):
    if col_type == "county":
        possible = ['NAME', 'name', '县', '市', '区', 'COUNTY', 'county', '地名']
    else:
        possible = ['NAME', 'name', '区划', '类型', 'zone', '区划名称', '农业区', '类型名', 'zonename', 'qu']
    for col in possible:
        if col in gdf.columns:
            return col
    return gdf.columns[0]


# ==============================================
# 翻译
# ==============================================
def auto_translate_zone(name):
    try:
        pinyin_list = lazy_pinyin(name)
        return ''.join(pinyin_list).title()
    except:
        return name


def interactive_translate(zones):
    print("\n" + "=" * 60)
    print("🔤 【翻译确认功能】")
    print("=" * 60)

    trans_map = {z: auto_translate_zone(z) for z in zones}

    print("📋 当前自动翻译结果：")
    for cn, en in trans_map.items():
        print(f"  {cn} → {en}")

    choice = input("\n是否需要手动修改翻译？(y=是，n=直接使用，默认n)：").strip().lower()
    if choice != 'y':
        print("✅ 使用自动翻译，不修改！")
        return trans_map

    print("\n✏️ 请输入自定义翻译（直接回车表示不修改）：")
    for cn in trans_map:
        current = trans_map[cn]
        user_input = input(f"  {cn}  ({current}) → ").strip()
        if user_input:
            trans_map[cn] = user_input
            print(f"  ✅ 已修改为：{user_input}")

    print("\n🎉 最终翻译结果：")
    for cn, en in trans_map.items():
        print(f"  {cn} → {en}")
    return trans_map


# ==============================================
# ✅ 精准匹配算法
# ==============================================
def accurate_zone_match(county_gdf, zone_gdf, zone_col):
    county = county_gdf.copy()
    zone = zone_gdf.copy()

    if county.crs is None:
        county.set_crs(epsg=4326, inplace=True)
    if zone.crs is None:
        zone.set_crs(epsg=4326, inplace=True)

    county_proj = county.to_crs(epsg=3857)
    zone_proj = zone.to_crs(epsg=3857)

    county_centroid = county_proj.geometry.centroid
    county_center_gdf = gpd.GeoDataFrame(county, geometry=county_centroid, crs=county_proj.crs)

    joined = gpd.sjoin(county_center_gdf, zone_proj, how='left', predicate='within')
    county["Agri_Zone"] = joined[zone_col].values

    na_mask = county["Agri_Zone"].isna()
    if na_mask.any():
        zone_points = zone_proj.copy()
        zone_points["center"] = zone_proj.geometry.centroid

        for idx in county[na_mask].index:
            pt = county_center_gdf.geometry.iloc[idx]
            zone_points["dist"] = zone_points.distance(pt)
            best = zone_points.loc[zone_points["dist"].idxmin(), zone_col]
            county.loc[idx, "Agri_Zone"] = best

    match_success_count = county["Agri_Zone"].notna().sum()
    print(f"\n✅ 空间匹配完成：成功匹配 {match_success_count} 个县级单元")
    return county


# ==============================================
# ✅ 匹配结果诊断统计
# ==============================================
def print_match_statistics(gdf):
    print("\n" + "=" * 80)
    print("📊 【空间匹配结果统计】")
    print("=" * 80)

    zone_counts = gdf["Agri_Zone"].value_counts().sort_index()
    total_zones = len(zone_counts)
    total_counties = len(gdf)
    success_count = gdf["Agri_Zone"].notna().sum()

    print(f"📍 匹配后有效区域种类：{total_zones} 类")
    print(f"📍 县级单元总数：{total_counties} 个")
    print(f"📍 成功匹配数量：{success_count} 个")
    print(f"📍 匹配成功率：{success_count / total_counties * 100:.2f}%")
    print("-" * 80)

    print("📌 各区域匹配县域数量：")
    for zone, cnt in zone_counts.items():
        pct = cnt / success_count * 100
        print(f"   🔹 {zone}：{cnt} 个 ({pct:.2f}%)")

    print("=" * 80 + "\n")


# ==============================================
# 地图元素
# ==============================================
def add_scalebar(ax, gdf):
    bounds = gdf.total_bounds
    lon_min, lon_max = bounds[0], bounds[2]
    lat_center = np.mean(bounds[[1, 3]])
    km_per_degree = 111.32 * np.cos(np.radians(lat_center))
    map_width_km = (lon_max - lon_min) * km_per_degree
    scale_total_km = 2000
    ratio = scale_total_km / map_width_km

    x0 = 0.70
    x1 = x0 + ratio * 0.8
    y = 0.08

    ax.plot([x0, x1], [y, y], lw=3, c='k', transform=ax.transAxes)
    ax.plot([x0, x0], [y, y + 0.02], lw=3, c='k', transform=ax.transAxes)
    ax.plot([x1, x1], [y, y + 0.02], lw=3, c='k', transform=ax.transAxes)
    ax.plot([(x0 + x1) / 2, (x0 + x1) / 2], [y, y + 0.015], lw=2, c='k', transform=ax.transAxes)

    ax.text(x0, y - 0.015, '0', fontsize=cfg.SCALE_TEXT_SIZE, ha='center', va='top', transform=ax.transAxes)
    ax.text((x0 + x1) / 2, y - 0.015, '1000', fontsize=cfg.SCALE_TEXT_SIZE, ha='center', va='top',
            transform=ax.transAxes)
    ax.text(x1, y - 0.015, '2000 km', fontsize=cfg.SCALE_TEXT_SIZE, ha='center', va='top', transform=ax.transAxes)


def add_grid(ax, gdf):
    ax.spines[:].set_visible(True)
    ax.spines[:].set_linewidth(1)
    ax.spines[:].set_color('black')
    bounds = gdf.total_bounds
    lon_min, lat_min, lon_max, lat_max = bounds

    lon_ticks = np.arange(np.ceil(lon_min / 5) * 5, np.floor(lon_max / 5) * 5 + 1, 5)
    lat_ticks = np.arange(np.ceil(lat_min / 5) * 5, np.floor(lat_max / 5) * 5 + 1, 5)

    for lon in lon_ticks:
        ax.axvline(lon, c='gray', ls='--', lw=0.6, alpha=0.7)
    for lat in lat_ticks:
        ax.axhline(lat, c='gray', ls='--', lw=0.6, alpha=0.7)

    ax.set_xticks(lon_ticks)
    ax.set_yticks(lat_ticks)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{x:.0f}°E"))
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f"{y:.0f}°N"))
    ax.tick_params(labelsize=cfg.GRID_LABEL_SIZE)


def add_compass(ax, img):
    ib = OffsetImage(img, zoom=cfg.COMPASS_ZOOM)
    ab = AnnotationBbox(ib, (0.08, 0.92), frameon=False, xycoords=ax.transAxes)
    ax.add_artist(ab)


# ==============================================
# 主绘图
# ==============================================
def draw_map():
    plt.rcParams['font.family'] = [cfg.FONT_NAME]
    plt.rcParams['axes.unicode_minus'] = False

    img_compass = load_svg_compass(cfg.SVG_COMPASS_PATH)
    df = pd.read_excel(cfg.EXCEL_PATH)
    county_gdf = read_shp_auto_encoding(cfg.SHP_PATH)

    border_gdf = read_shp_auto_encoding(cfg.BORDER_PATH)
    coastline_gdf = read_shp_auto_encoding(cfg.COASTLINE_PATH)
    zone_gdf = read_shp_auto_encoding(cfg.AGRI_ZONE_PATH)

    county_gdf, zone_gdf, border_gdf, coastline_gdf = align_crs(
        county_gdf, zone_gdf, border_gdf, coastline_gdf
    )

    print("=" * 60)
    print("📊 【分区shp】所有列名：")
    print(zone_gdf.columns.tolist())
    print("=" * 60)

    print("\n📌 分区SHP字段预览：")
    print_col = zone_gdf.copy().drop(columns=['geometry'], errors='ignore')
    for col in print_col.columns:
        print_col[col] = print_col[col].astype(str).apply(lambda x: x if len(str(x)) <= 10 else "EXTRM")
    print(print_col.head().to_string(show_dimensions=False))
    print("=" * 60)

    zone_col = input("请输入分区列名（如 qu）：").strip()
    while zone_col not in zone_gdf.columns:
        zone_col = input("列名不存在，请重新输入：").strip()

    original_all_zones = sorted(zone_gdf[zone_col].dropna().unique())
    print("\n" + "=" * 60)
    print("🟢 原始SHP全部区划：", len(original_all_zones), "个")
    print("=" * 60)
    for z in original_all_zones:
        print(f" - {z}")

    county_col = auto_detect_text_col(county_gdf, col_type='county')
    county_gdf['match_name'] = county_gdf[county_col].astype(str).str.strip()
    df['县'] = df['县'].astype(str).str.strip()

    df_year = df[df['year'] == cfg.TARGET_YEAR].copy()
    valid_counties = df_year['县'].dropna().unique()
    valid_count = len(valid_counties)
    print(f"\n📅 【{cfg.TARGET_YEAR}年 Excel 有效县数量】：{valid_count} 个")

    plot_gdf = accurate_zone_match(county_gdf, zone_gdf, zone_col)
    print_match_statistics(plot_gdf)

    plot_gdf['in_excel'] = plot_gdf['match_name'].isin(valid_counties)
    plot_gdf['Plot_Zone'] = plot_gdf['Agri_Zone']
    plot_gdf.loc[~plot_gdf['in_excel'], 'Plot_Zone'] = np.nan

    fig, ax = plt.subplots(1, 1, figsize=cfg.FIG_SIZE, dpi=cfg.DPI)

    plot_gdf.plot(ax=ax, facecolor=cfg.NULL_COLOR, edgecolor=cfg.BORDER_COLOR, linewidth=cfg.BORDER_WIDTH)

    plot_data = plot_gdf[plot_gdf['Plot_Zone'].notna()].copy()
    all_zones = sorted(plot_gdf['Agri_Zone'].dropna().unique())
    color_map = {z: cfg.COLORS[i % len(cfg.COLORS)] for i, z in enumerate(all_zones)}

    if not plot_data.empty:
        plot_data.plot(
            ax=ax,
            color=plot_data['Plot_Zone'].map(color_map),
            edgecolor=cfg.BORDER_COLOR,
            linewidth=cfg.BORDER_WIDTH
        )

    # ==============================================
    # ✅ 修复：只保留【图上实际画出的区划】作为图例
    # ==============================================
    plotted_zones = sorted(plot_data['Plot_Zone'].dropna().unique())  # 真正画在图上的区域
    valid_zones = plotted_zones

    trans_map = interactive_translate(valid_zones)

    legend_elements = []
    print("\n==================================================")
    print("🗺️ 最终图例（仅显示图上存在的区域）：")
    print("==================================================")
    for z in valid_zones:
        en_name = trans_map[z]
        print(f"- {z} → {en_name}")
        legend_elements.append(
            mpatches.Patch(facecolor=color_map[z], edgecolor='black', label=en_name)
        )
    legend_elements.append(mpatches.Patch(facecolor='white', edgecolor='black', label='No Data'))

    border_gdf.plot(ax=ax, color="black", linewidth=2.0, zorder=5)
    coastline_gdf.plot(ax=ax, color="#0A93FC", linewidth=1.0, zorder=5)

    add_grid(ax, plot_gdf)
    add_scalebar(ax, plot_gdf)
    add_compass(ax, img_compass)

    ax.legend(handles=legend_elements, loc='lower left', fontsize=cfg.LEGEND_FONT_SIZE, frameon=True,
              bbox_to_anchor=(0.02, 0.039))

    plt.tight_layout()
    plt.savefig(cfg.SAVE_PATH.replace('.svg', '.png'), format='png', dpi=cfg.DPI, bbox_inches='tight')
    plt.close()
    print(f"\n✅ 绘图完成：{cfg.SAVE_PATH.replace('.svg', '.png')}")


if __name__ == "__main__":
    draw_map()
