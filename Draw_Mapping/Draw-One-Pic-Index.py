import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
import warnings
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import cairosvg
import io
from PIL import Image
from matplotlib.colorbar import ColorbarBase
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Rectangle

warnings.filterwarnings('ignore')

# ==============================================
# 【超参数区：配置文件】
# ==============================================
class Config:
    EXCEL_PATH = r"E:\Test_Code\0626-draw\极端天气_县级年度统计——极端事件县域匹配的原始数据.xlsx"
    SHP_PATH = r"E:\Test_Code\2023年县级\县级.shp"
    BORDER_PATH = r"E:\Test_Code\国界\Export_Output.shp"
    COASTLINE_PATH = r"E:\Test_Code\海岸线\Export_Output_2.shp"
    SVG_COMPASS_PATH = r"E:\Test_Code\指北针.svg"

    KEY_FIELD = "合计次数"
    TARGET_YEARS = list(range(2009, 2025))

    FONT_NAME = "Times New Roman"
    # 起始#4973F9浅蓝，过渡到深红
    FULL_COLOR_RANGE = [
        '#4973F9','#6387FA','#7D9BFB','#97AFFC','#B1C3FD','#CBD7FE','#E5EBFF',
        '#F5E8B8','#FFE0A8','#FED69E','#FEC496','#FEB28E','#FEA088','#FD8E82',
        '#FD7C7C','#FC6A76','#F65870','#EE4868','#E63860','#DC2858','#CA1E4E',
        '#B81444','#A60A3A','#940030','#820028','#700020'
    ]
    NULL_COLOR = "white"
    BORDER_COLOR = "black"
    BORDER_WIDTH = 0.1

    NATIONAL_BORDER_COLOR = "black"
    NATIONAL_BORDER_WIDTH = 2.0
    COASTLINE_COLOR = "#0A93FC"
    COASTLINE_WIDTH = 1.0

    SAVE_PATH = r'E:\Test_Code\Extreme_Weather_County_Total_2009_2024.svg'
    DPI = 900
    FIG_SIZE = (16, 12)
    MAX_VAL = 30

cfg = Config()

# ==============================================
# 工具函数
# ==============================================
def load_svg_compass(svg_path):
    svg_png = cairosvg.svg2png(url=svg_path)
    return Image.open(io.BytesIO(svg_png))

def auto_detect_county_col(gdf):
    possible_cols = ['NAME', 'name', '县', '市', '区', 'COUNTY', 'county', '地名', '县级']
    for col in possible_cols:
        if col in gdf.columns:
            return col
    return gdf.columns[0]

def fuzzy_merge_county(gdf, data_df, key):
    gdf = gdf.copy().reset_index(drop=True)
    match_dict = {}
    data_df = data_df.copy()
    data_df['province'] = data_df['省份'].astype(str).str.strip()
    data_df['county_name'] = data_df['县级'].astype(str).str.strip()
    excel_map = dict(zip(zip(data_df['province'], data_df['county_name']), data_df[key]))
    for idx, shp_row in gdf.iterrows():
        shp_name = shp_row['match_name']
        shp_prov = shp_row['省']
        matched_val = None
        if shp_prov and (shp_prov, shp_name) in excel_map:
            matched_val = excel_map[(shp_prov, shp_name)]
        if matched_val is None and shp_prov:
            same_prov_df = data_df[data_df['province'] == shp_prov]
            for _, excel_row in same_prov_df.iterrows():
                ex_name = excel_row['county_name']
                if shp_name in ex_name or ex_name in shp_name:
                    matched_val = excel_row[key]
                    break
        if matched_val is None:
            for _, excel_row in data_df.iterrows():
                ex_name = excel_row['county_name']
                if shp_name in ex_name or ex_name in shp_name:
                    matched_val = excel_row[key]
                    break
        match_dict[shp_name] = matched_val
    gdf['TotalExtremeEvents'] = gdf['match_name'].map(match_dict)
    return gdf

# ==============================================
# 地图辅助元素
# ==============================================
def add_scalebar(ax, gdf):
    bounds = gdf.total_bounds
    lon_min, lon_max = bounds[0], bounds[2]
    lat_center = np.mean(bounds[[1, 3]])
    km_per_degree = 111.32 * np.cos(np.radians(lat_center))
    map_width_km = (lon_max - lon_min) * km_per_degree
    scale_total_km = 2000
    ratio = scale_total_km / map_width_km
    x0, x1, y = 0.05, 0.05 + ratio * 0.8, 0.07
    ax.plot([x0, x1], [y, y], lw=3, c='k', transform=ax.transAxes)
    ax.plot([x0, x0], [y, y + 0.02], lw=3, c='k', transform=ax.transAxes)
    ax.plot([x1, x1], [y, y + 0.02], lw=3, c='k', transform=ax.transAxes)
    ax.plot([(x0 + x1) / 2, (x0 + x1) / 2], [y, y + 0.015], lw=2, c='k', transform=ax.transAxes)
    # 字体放大2倍
    ax.text(x0, y - 0.02, '0', fontsize=22, ha='center', va='top', transform=ax.transAxes)
    ax.text((x0 + x1) / 2, y - 0.02, '1000', fontsize=22, ha='center', va='top', transform=ax.transAxes)
    ax.text(x1, y - 0.02, '2000 km', fontsize=22, ha='center', va='top', transform=ax.transAxes)

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
    # 坐标轴字体放大两倍
    ax.tick_params(labelsize=20)

def add_compass(ax, img):
    ib = OffsetImage(img, zoom=0.45)  # 增大指北针尺寸
    ab = AnnotationBbox(ib, (0.08, 0.92), frameon=False, xycoords=ax.transAxes)
    ax.add_artist(ab)

# ==============================================
# 主绘图函数
# ==============================================
def draw_map():
    # 全局字体整体放大两倍
    plt.rcParams['font.family'] = [cfg.FONT_NAME]
    plt.rcParams['axes.unicode_minus'] = False
    plt.rcParams['font.size'] = 24

    img_compass = load_svg_compass(cfg.SVG_COMPASS_PATH)
    df = pd.read_excel(cfg.EXCEL_PATH)
    county_gdf = gpd.read_file(cfg.SHP_PATH, encoding='utf-8')
    national_border = gpd.read_file(cfg.BORDER_PATH, encoding='utf-8')
    coastline = gpd.read_file(cfg.COASTLINE_PATH, encoding='utf-8')

    print("=" * 60)
    print("📊 数据预处理与基础统计")
    print("=" * 60)
    year_filtered_df = df[df['年份'].isin(cfg.TARGET_YEARS)].copy()
    print(f"目标年份范围：{year_filtered_df['年份'].min()}年 - {year_filtered_df['年份'].max()}年")
    print(f"原始年度记录数：{len(year_filtered_df)} 条")
    county_total_df = year_filtered_df.groupby(['省份', '县级'], as_index=False)[cfg.KEY_FIELD].sum()
    print(f"汇总后县级行政单位数：{len(county_total_df)} 个")
    print(f"【数据理论上限预设最大值】：{cfg.MAX_VAL}")
    print(f"【县域实际存在最大值】：{county_total_df[cfg.KEY_FIELD].max():.0f}")
    print(f"全周期极端事件总数平均值：{county_total_df[cfg.KEY_FIELD].mean():.2f}")
    print("=" * 60)

    county_col = "县级"
    county_gdf['match_name'] = county_gdf[county_col].astype(str).str.strip()
    if '省级' in county_gdf.columns:
        county_gdf['省'] = county_gdf['省级'].astype(str).str.strip()
    elif '省' in county_gdf.columns:
        county_gdf['省'] = county_gdf['省'].astype(str).str.strip()
    elif '省份' in county_gdf.columns:
        county_gdf['省'] = county_gdf['省份'].astype(str).str.strip()
    else:
        county_gdf['省'] = ''

    merged = fuzzy_merge_county(county_gdf, county_total_df, cfg.KEY_FIELD)
    valid_vals = merged['TotalExtremeEvents'].dropna().astype(int)
    null_cnt = merged['TotalExtremeEvents'].isna().sum()
    val_count_series = valid_vals.value_counts().sort_index()
    val_count_dict = dict(val_count_series)
    used_values = sorted(val_count_dict.keys())
    n_unique = len(used_values)

    print("\n===== 各数值县域出现次数（仅存在数据） =====")
    for val, cnt in val_count_dict.items():
        print(f"数值 {val} 次：{cnt} 个县")
    print(f"无匹配数据县域：{null_cnt} 个")
    print("=" * 60)

    # 动态区间
    breaks = [0]
    for v in used_values:
        breaks.append(v)
    breaks.append(cfg.MAX_VAL + 1)
    n_intervals = len(breaks) - 1
    full_color_len = len(cfg.FULL_COLOR_RANGE)
    sample_idx = np.linspace(0, full_color_len - 1, n_unique, dtype=int)
    used_colors = [cfg.FULL_COLOR_RANGE[i] for i in sample_idx]
    cmap = mcolors.ListedColormap(used_colors)
    norm = mcolors.BoundaryNorm(breaks, ncolors=n_intervals)

    fig, ax = plt.subplots(1, 1, figsize=cfg.FIG_SIZE, dpi=cfg.DPI)
    merged.plot(ax=ax, column='TotalExtremeEvents', cmap=cmap, norm=norm,
                edgecolor=cfg.BORDER_COLOR, linewidth=cfg.BORDER_WIDTH,
                missing_kwds={'color': cfg.NULL_COLOR})
    national_border.plot(ax=ax, color=cfg.NATIONAL_BORDER_COLOR, linewidth=cfg.NATIONAL_BORDER_WIDTH)
    coastline.plot(ax=ax, color=cfg.COASTLINE_COLOR, linewidth=cfg.COASTLINE_WIDTH)

    add_grid(ax, county_gdf)
    add_scalebar(ax, merged)
    add_compass(ax, img_compass)

    # -------------------------- 渐变图例：再缩小一半 + 靠左放置 --------------------------
    grad_cmap = LinearSegmentedColormap.from_list("grad_bar", cfg.FULL_COLOR_RANGE, N=200)
    # 宽度0.235（再次减半），x起始0.03靠左，纵向y=0.18保持上移
    cbar_ax = fig.add_axes([0.15, 0.2, 0.210, 0.03])
    cb = ColorbarBase(cbar_ax, cmap=grad_cmap, orientation='horizontal')
    # 标注刻度：2,5,10,15,20,24
    tick_target = [2,5,10,15,20,24]
    v_min, v_max = min(used_values), max(used_values)
    tick_pos = [(x - v_min) / (v_max - v_min) for x in tick_target]
    cb.set_ticks(tick_pos)
    cb.set_ticklabels([str(x) for x in tick_target])
    cb.ax.tick_params(labelsize=23)
    # No Data 白色方块放在图例右侧
    # rect = Rectangle((0.385, 0.2), 0.022, 0.03, facecolor='white', ec='black', transform=fig.transFigure)
    # fig.patches.append(rect)
    # fig.text(0.41, 0.21, "No Data", fontsize=24, va='center')

    # 无图标题（已删除）
    plt.tight_layout()
    plt.savefig(cfg.SAVE_PATH, format='svg', bbox_inches='tight')
    plt.savefig(cfg.SAVE_PATH.replace('.svg', '.png'), format='png', bbox_inches='tight', dpi=cfg.DPI)
    plt.close()
    print("\n✅ 绘图完成！文件已保存：")
    print(f"矢量图SVG: {cfg.SAVE_PATH}")
    print(f"高清图PNG: {cfg.SAVE_PATH.replace('.svg', '.png')}")

if __name__ == "__main__":
    draw_map()