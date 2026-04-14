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
    BORDER_PATH = r"E:\Test_Code\国界\Export_Output.shp"  # 新增国界
    COASTLINE_PATH = r"E:\Test_Code\海岸线\Export_Output_2.shp"  # 新增海岸线
    SVG_COMPASS_PATH = r"E:\Test_Code\指北针.svg"
    AGRI_ZONE_PATH = r"E:\Test_Code\中国气候区划\Climate_quhua.shp"

    # --------------------- 【年份与绘图设置】 ---------------------
    TARGET_YEAR = 2023

    # --------------------- 【字体配置】 ---------------------
    FONT_NAME = "Times New Roman"

    # --------------------- 【颜色与边界配置】 ---------------------
    NULL_COLOR = "white"
    BORDER_COLOR = "black"
    BORDER_WIDTH = 0.1
    COAST_COLOR = "#0A93FC"  # 海岸线颜色
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
# 空间匹配
# ==============================================
def robust_zone_match(county_gdf, zone_gdf, zone_col):
    county = county_gdf.copy()
    zone = zone_gdf.copy()

    if county.crs != zone.crs:
        zone = zone.to_crs(county.crs)

    county = county[county.geometry.is_valid & ~county.geometry.is_empty].copy()
    zone = zone[zone.geometry.is_valid & ~zone.geometry.is_empty].copy()

    sindex = zone.sindex
    county["Agri_Zone"] = None

    for idx, county_row in county.iterrows():
        county_geom = county_row.geometry
        best_name = None
        max_score = -1

        possible_matches = zone.iloc[list(sindex.intersection(county_geom.bounds))]

        for _, zrow in possible_matches.iterrows():
            zone_geom = zrow.geometry
            zone_name = zrow[zone_col]

            try:
                center = county_geom.centroid
                if center.within(zone_geom):
                    best_name = zone_name
                    break

                inter = county_geom.intersection(zone_geom)
                score = inter.area
                dis = center.distance(zone_geom)
                score += (100 / (dis + 1))

                if score > max_score:
                    max_score = score
                    best_name = zone_name
            except:
                continue

        if best_name is None:
            try:
                center_pt = county_geom.centroid
                zone["dist"] = zone.geometry.distance(center_pt)
                best_name = zone.loc[zone["dist"].idxmin(), zone_col]
            except:
                best_name = zone[zone_col].iloc[0]

        county.at[idx, "Agri_Zone"] = best_name

    return county


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

    # 右下角，整体再右移，顺序 0 → 1000 → 2000
    x0 = 0.70
    x1 = x0 + ratio * 0.8
    y = 0.08

    ax.plot([x0, x1], [y, y], lw=3, c='k', transform=ax.transAxes)
    ax.plot([x0, x0], [y, y + 0.02], lw=3, c='k', transform=ax.transAxes)
    ax.plot([x1, x1], [y, y + 0.02], lw=3, c='k', transform=ax.transAxes)
    ax.plot([(x0 + x1) / 2, (x0 + x1) / 2], [y, y + 0.015], lw=2, c='k', transform=ax.transAxes)

    # 文字从左到右：0 — 1000 — 2000 km
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

    # 读取 国界 + 海岸线
    border_gdf = read_shp_auto_encoding(cfg.BORDER_PATH)
    coastline_gdf = read_shp_auto_encoding(cfg.COASTLINE_PATH)

    zone_gdf = read_shp_auto_encoding(cfg.AGRI_ZONE_PATH)

    # 坐标系统一
    county_gdf, zone_gdf, border_gdf, coastline_gdf = align_crs(
        county_gdf, zone_gdf, border_gdf, coastline_gdf
    )

    print("=" * 60)
    print("📊 【分区shp】所有列名：")
    print(zone_gdf.columns.tolist())
    print("=" * 60)

    print("\n📌 分区SHP字段预览（超长内容→EXTRM）：")
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
    print("🟢 原始SHP全部区划（所有颜色）：", len(original_all_zones), "个")
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

    plot_gdf = robust_zone_match(county_gdf, zone_gdf, zone_col)
    plot_gdf['in_excel'] = plot_gdf['match_name'].isin(valid_counties)

    fig, ax = plt.subplots(1, 1, figsize=cfg.FIG_SIZE, dpi=cfg.DPI)

    zone_plot = zone_gdf.copy()
    all_zones = zone_plot[zone_col].dropna().unique()
    color_map = {z: cfg.COLORS[i % len(cfg.COLORS)] for i, z in enumerate(all_zones)}
    zone_plot.plot(ax=ax, color=zone_plot[zone_col].map(color_map), edgecolor='none', alpha=0.9)

    no_data = plot_gdf[~plot_gdf['in_excel']]
    no_data.plot(ax=ax, facecolor='white', edgecolor='white', linewidth=0)

    final_used_zones = original_all_zones

    # ===================== 计算有效面积（排除白色区域） =====================
    print("\n" + "=" * 60)
    print("✅ 【最终图上彩色区域面积】白色无数据区已排除")
    print("=" * 60)

    valid_counties = plot_gdf[plot_gdf['in_excel']].copy()

    # ✅ 修复几何错误（只加这两行，不改变任何逻辑）
    valid_counties["geometry"] = valid_counties["geometry"].make_valid()
    zone_gdf["geometry"] = zone_gdf["geometry"].make_valid()

    clip_geom = valid_counties.unary_union
    clipped_zones = zone_gdf.clip(clip_geom)

    proj = clipped_zones.to_crs("+proj=aea +lat_1=25 +lat_2=47 +lat_0=0 +lon_0=105 +datum=WGS84 +units=m +no_defs")
    proj['area_km2'] = proj.geometry.area / 1e6
    area_result = proj.groupby(zone_col)['area_km2'].sum()

    # ===================== 自动删除面积=0的项，只保留有面积的图例 =====================
    valid_zones = [z for z in final_used_zones if area_result.get(z, 0.0) > 0]
    print(f"\n✅ 最终有效区划（有面积）：{len(valid_zones)} 个")
    for z in valid_zones:
        print(f" - {z:<20} 面积：{area_result.get(z, 0.0):>11.2f} km²")

    # 翻译只保留有效区划
    trans_map = interactive_translate(valid_zones)

    # 图例只生成有效区划
    legend_elements = []
    print("\n==================================================")
    print("🗺️ 最终图例（已剔除面积为0的项）：")
    print("==================================================")
    for z in valid_zones:
        en_name = trans_map[z]
        print(f"- {z} → {en_name}")
        legend_elements.append(
            mpatches.Patch(facecolor=color_map[z], edgecolor='black', label=en_name)
        )
    legend_elements.append(mpatches.Patch(facecolor='white', edgecolor='black', label='No Data'))
    # ==================================================================================

    plot_gdf.plot(ax=ax, facecolor='none', edgecolor=cfg.BORDER_COLOR, linewidth=cfg.BORDER_WIDTH)

    # ===================== 绘制国界（黑色 稍粗）+ 海岸线（0A93FC） =====================
    border_gdf.plot(ax=ax, color="black", linewidth=2.0, zorder=5)
    coastline_gdf.plot(ax=ax, color="#0A93FC", linewidth=1.0, zorder=5)

    add_grid(ax, plot_gdf)
    add_scalebar(ax, plot_gdf)
    add_compass(ax, img_compass)

    # ===================== 图例向下移动一点 =====================
    ax.legend(handles=legend_elements, loc='lower left', fontsize=cfg.LEGEND_FONT_SIZE, frameon=True,
              bbox_to_anchor=(0.02, 0.039))

    plt.tight_layout()
    plt.savefig(cfg.SAVE_PATH.replace('.svg', '.png'), format='png', dpi=cfg.DPI, bbox_inches='tight')
    plt.close()
    print(f"\n✅ 绘图完成：{cfg.SAVE_PATH.replace('.svg', '.png')}")

if __name__ == "__main__":
    draw_map()
