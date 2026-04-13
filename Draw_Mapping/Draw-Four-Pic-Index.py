import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.patches as mpatches
import numpy as np
import warnings
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import cairosvg
import io
from PIL import Image

warnings.filterwarnings('ignore')


# ==============================================
# 【超参数区：前端/配置文件 直接修改这里】
# ==============================================
class Config:
    # === 文件路径 ===
    EXCEL_PATH = r"E:\Test_Code\平衡面板.xlsx"
    SHP_PATH = r"E:\Test_Code\2023年县级\县级.shp"
    BORDER_PATH = r"E:\Test_Code\国界\Export_Output.shp"  # 国界线
    COASTLINE_PATH = r"E:\Test_Code\海岸线\Export_Output_2.shp"  # 海岸线
    SVG_COMPASS_PATH = r"E:\Test_Code\指北针.svg"

    # === 数据字段 ===
    KEY_FIELD = "x数农融合" # 超效率
    TARGET_YEARS = [2011, 2016, 2020, 2023]  #

    # === 绘图样式 ===
    FONT_NAME = "Times New Roman"
    COLORS = ['#FFE100', '#3E26F2', '#31A354', '#FEB24C', '#FC4E2A']
    NULL_COLOR = "white"
    BORDER_COLOR = "black"
    BORDER_WIDTH = 0.1

    # 国界线样式
    NATIONAL_BORDER_COLOR = "black"
    NATIONAL_BORDER_WIDTH = 2.0

    # 海岸线样式
    COASTLINE_COLOR = "#0A93FC"
    COASTLINE_WIDTH = 1.0

    # === 输出 ===
    SAVE_PATH = r'E:\Test_Code\Four_Year_DA_Final.svg'
    DPI = 900
    FIG_SIZE = (18, 14)

    # === 分级 ===
    N_CLASSES = 5
    # 区间修正：True=自动计算，填入数字=强制覆盖
    # 格式：[True, 0.0010, True, True, True, True]
    BREAKS_OVERRIDE = [True, 0.0010, True, True, True, True]


cfg = Config()


# ==============================================
# 工具函数
# ==============================================
def load_svg_compass(svg_path):
    svg_png = cairosvg.svg2png(url=svg_path)
    return Image.open(io.BytesIO(svg_png))


def auto_detect_county_col(gdf):
    possible_cols = ['NAME', 'name', '县', '市', '区', 'COUNTY', 'county', '地名']
    for col in possible_cols:
        if col in gdf.columns:
            return col
    return gdf.columns[0]


def fuzzy_merge_county(gdf, data_df, year, key):
    gdf = gdf.copy().reset_index(drop=True)
    match_dict = {}
    for _, shp_row in gdf.iterrows():
        shp_name = shp_row['match_name']
        matched_val = None
        for _, excel_row in data_df.iterrows():
            excel_name = excel_row['县']
            eff_val = excel_row[key]
            if excel_name in shp_name or shp_name in excel_name:
                matched_val = eff_val
                break
        match_dict[shp_name] = matched_val
    gdf['SuperEfficiency'] = gdf['match_name'].map(match_dict)
    return gdf


# ==============================================
# 科学分级算法（全自动 + 支持手动覆盖断点）
# ==============================================
def get_log_quantile_breaks(series, override, n=5):
    vals = series.dropna()
    vals = vals[vals > 0]
    log_vals = np.log10(vals)
    breaks_log = np.percentile(log_vals, np.linspace(0, 100, n + 1))
    breaks = np.power(10, breaks_log)
    breaks = np.round(breaks, 4)
    breaks = np.unique(breaks)

    while len(breaks) < n + 1:
        breaks = np.append(breaks, np.round(breaks[-1] * 1.5, 4))

    if breaks[0] != 0:
        breaks = np.insert(breaks, 0, 0.0)
        breaks = breaks[:n + 1]

    # 应用断点覆盖
    final = []
    for b, o in zip(breaks, override):
        if o is True:
            final.append(round(float(b), 4))
        else:
            final.append(round(float(o), 4))
    return final


# ==============================================
# 地图元素：比例尺、经纬网、指北针
# ==============================================
def add_scalebar(ax, gdf):
    bounds = gdf.total_bounds
    lon_min, lon_max = bounds[0], bounds[2]
    lat_center = np.mean(bounds[[1, 3]])
    km_per_degree = 111.32 * np.cos(np.radians(lat_center))
    map_width_km = (lon_max - lon_min) * km_per_degree
    scale_total_km = 2000
    ratio = scale_total_km / map_width_km
    x0, x1, y = 0.05, 0.05 + ratio * 0.8, 0.08

    ax.plot([x0, x1], [y, y], lw=3, c='k', transform=ax.transAxes)
    ax.plot([x0, x0], [y, y + 0.02], lw=3, c='k', transform=ax.transAxes)
    ax.plot([x1, x1], [y, y + 0.02], lw=3, c='k', transform=ax.transAxes)
    ax.plot([(x0 + x1) / 2, (x0 + x1) / 2], [y, y + 0.015], lw=2, c='k', transform=ax.transAxes)

    ax.text(x0, y - 0.03, '0', fontsize=14, ha='center', va='top', transform=ax.transAxes)
    ax.text((x0 + x1) / 2, y - 0.03, '1000', fontsize=14, ha='center', va='top', transform=ax.transAxes)
    ax.text(x1, y - 0.03, '2000 km', fontsize=14, ha='center', va='top', transform=ax.transAxes)


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
    ax.tick_params(labelsize=11)


def add_compass(ax, img):
    ib = OffsetImage(img, zoom=0.22)
    ab = AnnotationBbox(ib, (0.08, 0.92), frameon=False, xycoords=ax.transAxes)
    ax.add_artist(ab)


# ==============================================
# 统计分级样本量
# ==============================================
def count_levels(gdf, breaks):
    eff_data = gdf['SuperEfficiency'].dropna()
    null_count = gdf['SuperEfficiency'].isna().sum()
    counts = []
    for i in range(len(breaks) - 1):
        mask = (eff_data >= breaks[i]) & (eff_data <= breaks[i + 1]) if i == 0 else \
            (eff_data > breaks[i]) & (eff_data <= breaks[i + 1])
        counts.append(mask.sum())
    return counts, null_count


# ==============================================
# 【核心：主绘图函数，完全封装】
# ==============================================
def draw_map():
    plt.rcParams['font.family'] = [cfg.FONT_NAME]
    plt.rcParams['axes.unicode_minus'] = False

    # 加载数据
    img_compass = load_svg_compass(cfg.SVG_COMPASS_PATH)
    df = pd.read_excel(cfg.EXCEL_PATH)
    county_gdf = gpd.read_file(cfg.SHP_PATH, encoding='utf-8')
    national_border = gpd.read_file(cfg.BORDER_PATH, encoding='utf-8')  # 国界
    coastline = gpd.read_file(cfg.COASTLINE_PATH, encoding='utf-8')  # 海岸线

    # 先获取列名 → 再简化！顺序修复！
    county_col = auto_detect_county_col(county_gdf)
    county_gdf['match_name'] = county_gdf[county_col].astype(str).str.strip()
    df['县'] = df['县'].astype(str).str.strip()

    # 分级（支持手动覆盖）
    breaks = get_log_quantile_breaks(df[cfg.KEY_FIELD], cfg.BREAKS_OVERRIDE, n=cfg.N_CLASSES)
    cmap = mcolors.ListedColormap(cfg.COLORS)
    norm = mcolors.BoundaryNorm(breaks, ncolors=len(cfg.COLORS))

    # 图例
    legend_elements = []
    for i in range(len(cfg.COLORS)):
        if i == 0:
            label = f"{breaks[i]:.4f} ~ {breaks[i + 1]:.4f}"
        else:
            label = f"{breaks[i] + 0.0001:.4f} ~ {breaks[i + 1]:.4f}"
        legend_elements.append(mpatches.Patch(facecolor=cfg.COLORS[i], edgecolor='black', label=label))
    legend_elements.append(mpatches.Patch(facecolor='white', edgecolor='black', label='No Data'))

    # 绘图
    fig, axes = plt.subplots(2, 2, figsize=cfg.FIG_SIZE, dpi=cfg.DPI)
    axes = axes.flatten()

    # ====================== 新增全局统计打印 ======================
    print("=" * 60)
    print("📊 全局基础统计")
    print("=" * 60)
    print(f"指标字段：{cfg.KEY_FIELD}")
    print(f"总年份：{sorted(df['year'].unique())}")
    print(f"绘图年份：{cfg.TARGET_YEARS}")
    print(f"Excel总记录数：{len(df)} 条")
    print(f"Excel含县数量：{df['县'].nunique()} 个")
    print(f"SHP地图县级数量：{len(county_gdf)} 个")
    print(f"分级断点：{breaks}")
    print("=" * 60)
    # =============================================================

    for ax, year in zip(axes, cfg.TARGET_YEARS):
        print(f"\n===== {year} =====")
        year_df = df[df['year'] == year].copy()

        # ====================== 新增单年份统计打印 ======================
        print(f"📌 {year}年 Excel记录数：{len(year_df)} 条")
        print(f"📌 {year}年 有效数值数量：{year_df[cfg.KEY_FIELD].count()} 条")
        print(f"📌 {year}年 数值缺失数：{year_df[cfg.KEY_FIELD].isna().sum()} 条")
        print(f"📌 {year}年 指标最小值：{year_df[cfg.KEY_FIELD].min():.4f}")
        print(f"📌 {year}年 指标最大值：{year_df[cfg.KEY_FIELD].max():.4f}")
        print(f"📌 {year}年 指标平均值：{year_df[cfg.KEY_FIELD].mean():.4f}")
        # ===============================================================

        merged = fuzzy_merge_county(county_gdf, year_df, year, cfg.KEY_FIELD)

        # 统计
        counts, null_cnt = count_levels(merged, breaks)
        for idx, cnt in enumerate(counts):
            print(f"分级{idx + 1}: {cnt} 个县")
        print(f"无数据: {null_cnt}")

        # 画地图
        merged.plot(ax=ax, column='SuperEfficiency', cmap=cmap, norm=norm,
                    edgecolor=cfg.BORDER_COLOR, linewidth=cfg.BORDER_WIDTH,
                    missing_kwds={'color': cfg.NULL_COLOR})

        # 绘制国界线 + 海岸线
        national_border.plot(ax=ax, color=cfg.NATIONAL_BORDER_COLOR, linewidth=cfg.NATIONAL_BORDER_WIDTH)
        coastline.plot(ax=ax, color=cfg.COASTLINE_COLOR, linewidth=cfg.COASTLINE_WIDTH)

        add_grid(ax, merged)
        add_scalebar(ax, merged)
        add_compass(ax, img_compass)

        ax.legend(handles=legend_elements, loc='lower left', fontsize=14, frameon=True, bbox_to_anchor=(0.02, 0.12))
        ax.set_title(f'{year}', fontsize=22, fontweight='bold', pad=10)

    plt.tight_layout()
    # plt.savefig(cfg.SAVE_PATH, format='svg', bbox_inches='tight')
    plt.savefig(cfg.SAVE_PATH.replace('.svg', '.png'), format='png', bbox_inches='tight', dpi=cfg.DPI)
    plt.close()
    print("\n✅ 绘图完成！")


# ==============================================
# 运行
# ==============================================
if __name__ == "__main__":
    draw_map()