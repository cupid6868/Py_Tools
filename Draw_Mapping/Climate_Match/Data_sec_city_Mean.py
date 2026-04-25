import os
import glob
import re
import pandas as pd
import numpy as np
import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib
from tqdm import tqdm
from shapely.geometry import Point
from shapely.ops import unary_union

# ===================== 全局配置参数（根据实际路径修改） =====================
# 基础数据路径
ROOT_DIR = r"D:\Data_1"  # 气象数据根目录（包含年份文件夹）
CITY_SHP_PATH = r"D:\Data_2\2023年地级\地级.shp"  # 市级边界SHP文件路径
BASELINE_YEARS = (1981, 2010)  # 气候基准期
TARGET_YEARS = (2010, 2024)  # 统计极端天数的目标期
RESULT_DIR = "02_city_weather_heatmaps_Aver"  # 结果文件夹名称
PLOT_DIR = "02_city_weather_heatmaps_Aver"  # 热力图保存文件夹

# SHP字段映射（根据你的市级SHP文件调整）
CITY_FIELD_MAP = {
    "code": ["市代码", "CODE", "ADMINCODE", "ID", "地级码"],  # 市级代码字段
    "name": ["市名称", "NAME", "CNAME", "地名", "地级市名"]  # 市级名称字段
}

# 图表配置
PLOT_CONFIG = {
    "figsize": (16, 12),
    "dpi": 300,
    "cmap": "RdYlBu_r",
    "station_size": 8,
    "station_color": "#FF0000",  # 红色标注气象站
    "station_alpha": 0.8,
    "title_fontsize": 16,
    "label_fontsize": 12,
    "legend_fontsize": 10
}

# 输出文件配置
FILES_CONFIG = {
    "baseline_sorted": "01_baseline_data_sorted.csv",
    "thresholds": "02_weather_thresholds.csv",
    "extreme_days": "03_extreme_weather_yearly_stats.csv",
    "city_panel": "04_city_weather_panel_data.csv"
}

# 极端天气指标列表
EXTREME_METRICS = [
    "extreme_low_days", "extreme_high_days",
    "extreme_rain_days", "extreme_drought_days"
]
METRIC_ALIASES = {
    "extreme_low_days": "极端低温天数",
    "extreme_high_days": "极端高温天数",
    "extreme_rain_days": "极端降雨天数",
    "extreme_drought_days": "极端干旱天数"
}


# ===================== 修复中文显示问题 =====================
def setup_chinese_font():
    """配置Matplotlib支持中文显示（兼容多系统）"""
    # 关闭字体警告
    matplotlib.rcParams['font.sans-serif'] = [
        'Microsoft YaHei',  # Windows 系统
        'SimHei',  # Windows 备选
        'PingFang SC',  # macOS 系统
        'Hiragino Sans GB',  # macOS 备选
        'DejaVu Sans'  # 兜底（无中文时）
    ]
    matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
    matplotlib.rcParams['font.family'] = 'sans-serif'
    print("✅ 中文字体配置完成")


# 初始化中文配置
setup_chinese_font()


# ===========================================================================
# 基础工具函数
# ===========================================================================
def create_directory(dir_name):
    """创建文件夹（支持多层级）"""
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
        print(f"✅ 已创建文件夹：{os.path.abspath(dir_name)}")
    else:
        print(f"ℹ️ 文件夹已存在：{os.path.abspath(dir_name)}")
    return dir_name


def sanitize_filename(filename):
    """清理文件名中的非法字符"""
    illegal_chars = r'[<>:\"/\\|?*]'
    sanitized = re.sub(illegal_chars, '_', filename)
    sanitized = re.sub('_+', '_', sanitized)
    return sanitized.strip('_')


def haversine_distance(lat1, lon1, lat2, lon2):
    """计算两点间哈弗辛距离（米）- 仅保留，本次修改未使用"""
    lat1_rad = np.radians(lat1)
    lon1_rad = np.radians(lon1)
    lat2_rad = np.radians(lat2)
    lon2_rad = np.radians(lon2)

    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1_rad) * np.cos(lat2_rad) * np.sin(dlon / 2) ** 2
    c = 2 * np.arcsin(np.sqrt(a))
    return c * 6371000  # 地球半径（米）


# ===========================================================================
# 第一步：加载和预处理气象数据（保留原有逻辑）
# ===========================================================================
def load_all_data(root_dir, start_year, end_year):
    """加载指定年份范围的所有气象数据，返回合并后的DataFrame"""
    all_dfs = []
    for year in tqdm(range(start_year, end_year + 1), desc="加载数据"):
        year_dir = os.path.join(root_dir, str(year))
        if not os.path.exists(year_dir):
            print(f"⚠️ 警告：年份文件夹 {year_dir} 不存在，跳过")
            continue

        csv_files = glob.glob(os.path.join(year_dir, "*.csv"))
        if not csv_files:
            print(f"⚠️ 警告：{year} 年文件夹内无CSV文件，跳过")
            continue

        for file in csv_files:
            try:
                df = pd.read_csv(
                    file,
                    parse_dates=["DATE"],
                    usecols=["STATION", "NAME", "LATITUDE", "LONGITUDE", "DATE", "MIN", "MAX", "PRCP"]
                )
                df = df.dropna(subset=["STATION", "NAME", "LATITUDE", "LONGITUDE", "DATE", "MIN", "MAX", "PRCP"])

                # 替换缺失值标记
                missing_marks = [9999.9, 99.99, -9999.9, -99.99]
                df["MIN"] = df["MIN"].replace(missing_marks, np.nan)
                df["MAX"] = df["MAX"].replace(missing_marks, np.nan)
                df["PRCP"] = df["PRCP"].replace(missing_marks, np.nan)

                df = df.dropna(subset=["MIN", "MAX", "PRCP"])
                df = df[(df["PRCP"] >= 0)]

                # 单位转换
                df["MIN_C"] = (df["MIN"] - 32) * 5 / 9
                df["MAX_C"] = (df["MAX"] - 32) * 5 / 9
                df["PRCP_mm"] = df["PRCP"] * 25.4

                df["year"] = df["DATE"].dt.year
                df["month_day"] = df["DATE"].dt.strftime("%m-%d")

                # 经纬度精度处理
                df["LATITUDE"] = df["LATITUDE"].round(6)
                df["LONGITUDE"] = df["LONGITUDE"].round(6)

                all_dfs.append(df)
            except Exception as e:
                print(f"❌ 读取文件 {file} 失败：{e}，跳过")

    if not all_dfs:
        raise ValueError("未加载到任何有效数据，请检查路径/文件格式/列名")

    combined_df = pd.concat(all_dfs, ignore_index=True).drop_duplicates()
    return combined_df


def sort_baseline_data(baseline_df, result_dir):
    """基准期数据升序排列"""
    baseline_sorted = baseline_df.sort_values(
        by=["STATION", "DATE"],
        ascending=True
    ).reset_index(drop=True)

    output_path = os.path.join(result_dir, FILES_CONFIG["baseline_sorted"])
    baseline_sorted.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"✅ 基准期数据已保存：{output_path}")
    return baseline_sorted


def calculate_relative_thresholds(baseline_df, result_dir):
    """计算极端天气相对阈值"""
    station_coords = baseline_df.groupby("STATION")[["NAME", "LATITUDE", "LONGITUDE"]].first().reset_index()

    threshold_df = baseline_df.groupby(["STATION", "month_day"]).agg(
        extreme_low_temp=("MIN_C", lambda x: np.percentile(x, 10)),
        extreme_high_temp=("MAX_C", lambda x: np.percentile(x, 90)),
        extreme_rain=("PRCP_mm", lambda x: np.percentile(x, 95)),
        extreme_drought=("PRCP_mm", lambda x: np.percentile(x, 5))
    ).reset_index()

    threshold_df = threshold_df.merge(station_coords, on="STATION", how="left")
    threshold_df = threshold_df[["STATION", "NAME", "LATITUDE", "LONGITUDE", "month_day",
                                 "extreme_low_temp", "extreme_high_temp", "extreme_rain", "extreme_drought"]]

    output_path = os.path.join(result_dir, FILES_CONFIG["thresholds"])
    threshold_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"✅ 阈值数据已保存：{output_path}")
    return threshold_df


def count_extreme_days_yearly(target_df, threshold_df, result_dir):
    """统计每年各气象站极端天数"""
    station_coords = target_df.groupby("STATION")[["NAME", "LATITUDE", "LONGITUDE"]].first().reset_index()

    merged_df = pd.merge(
        target_df,
        threshold_df[["STATION", "month_day", "extreme_low_temp", "extreme_high_temp",
                      "extreme_rain", "extreme_drought"]],
        on=["STATION", "month_day"],
        how="left"
    )

    # 标记极端天气
    merged_df["is_extreme_low"] = merged_df["MIN_C"] < merged_df["extreme_low_temp"]
    merged_df["is_extreme_high"] = merged_df["MAX_C"] > merged_df["extreme_high_temp"]
    merged_df["is_extreme_rain"] = merged_df["PRCP_mm"] > merged_df["extreme_rain"]
    merged_df["is_extreme_drought"] = merged_df["PRCP_mm"] < merged_df["extreme_drought"]

    # 年度统计
    extreme_stats = merged_df.groupby(["STATION", "year"]).agg(
        extreme_low_days=("is_extreme_low", "sum"),
        extreme_high_days=("is_extreme_high", "sum"),
        extreme_rain_days=("is_extreme_rain", "sum"),
        extreme_drought_days=("is_extreme_drought", "sum"),
        total_valid_days=("DATE", "count")
    ).reset_index()

    extreme_stats = extreme_stats.merge(station_coords, on="STATION", how="left")

    # 类型转换
    for metric in EXTREME_METRICS + ["total_valid_days"]:
        extreme_stats[metric] = extreme_stats[metric].astype(int)

    extreme_stats = extreme_stats[["STATION", "NAME", "LATITUDE", "LONGITUDE", "year",
                                   "extreme_low_days", "extreme_high_days", "extreme_rain_days",
                                   "extreme_drought_days", "total_valid_days"]]
    extreme_stats = extreme_stats.sort_values(["year", "STATION"]).reset_index(drop=True)

    output_path = os.path.join(result_dir, FILES_CONFIG["extreme_days"])
    extreme_stats.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"✅ 极端天数数据已保存：{output_path}")
    return extreme_stats


# ===========================================================================
# 第二步：加载市级边界数据
# ===========================================================================
def load_city_boundaries(shp_path):
    """加载市级边界数据（移除中心点计算，仅保留核心字段）"""
    print("\n===== 加载市级边界数据 =====")
    gdf = gpd.read_file(shp_path)

    # 设置坐标系
    if gdf.crs is None:
        gdf = gdf.set_crs("EPSG:4326")

    # 匹配市级代码和名称字段
    code_field = next((f for f in CITY_FIELD_MAP["code"] if f in gdf.columns), None)
    if code_field is None:
        code_field = "city_code_temp"
        gdf[code_field] = gdf.index.astype(str)

    name_field = next((f for f in CITY_FIELD_MAP["name"] if f in gdf.columns), None)
    if name_field is None:
        name_field = "city_name_temp"
        gdf[name_field] = [f"市_{i}" for i in gdf.index]

    # 整理结果（移除中心点相关字段）
    city_gdf = gdf.rename(columns={
        code_field: "city_code",
        name_field: "city_name"
    })[["city_code", "city_name", "geometry"]]

    # 统一代码为字符串类型
    city_gdf["city_code"] = city_gdf["city_code"].astype(str)

    print(f"✅ 加载完成：共{len(city_gdf)}个市级行政区")
    return city_gdf


# ===========================================================================
# 第三步：市级指标计算（修改核心：算术平均值）
# ===========================================================================
def assign_station_to_city(station_data, city_gdf):
    """将气象站点匹配到对应的地级市"""
    # 将站点数据转为GeoDataFrame
    station_geo = gpd.GeoDataFrame(
        station_data,
        geometry=gpd.points_from_xy(station_data["LONGITUDE"], station_data["LATITUDE"]),
        crs="EPSG:4326"
    )

    # 空间连接：匹配站点所属的地级市
    station_with_city = gpd.sjoin(
        station_geo,
        city_gdf[["city_code", "city_name", "geometry"]],
        how="left",
        predicate="within"
    )

    # 移除冗余列
    station_with_city = station_with_city.drop(columns=["index_right", "geometry"])

    # 统计匹配结果
    matched_count = station_with_city["city_code"].notna().sum()
    total_count = len(station_with_city)
    print(f"✅ 站点匹配完成：{matched_count}/{total_count} 个站点匹配到地级市")

    return station_with_city


def calculate_city_average(station_stats, city_gdf):
    """按地级市计算站点指标的算术平均值"""
    print("\n===== 计算地级市算术平均值 =====")

    # 1. 匹配站点到地级市
    station_with_city = assign_station_to_city(station_stats, city_gdf)

    # 2. 按城市+年份分组计算算术平均值
    city_stats = station_with_city.groupby(["city_code", "city_name", "year"]).agg(
        extreme_low_days=("extreme_low_days", "mean"),
        extreme_high_days=("extreme_high_days", "mean"),
        extreme_rain_days=("extreme_rain_days", "mean"),
        extreme_drought_days=("extreme_drought_days", "mean"),
        station_count=("STATION", "nunique")  # 记录每个城市每年的站点数量
    ).reset_index()

    # 保留两位小数
    for metric in EXTREME_METRICS:
        city_stats[metric] = city_stats[metric].round(2)

    return city_stats


def generate_city_panel_data(city_gdf, station_stats, result_dir):
    """生成市级气象面板数据（算术平均版）"""
    print("\n===== 生成市级气象面板数据 =====")

    # 计算地级市算术平均值
    city_panel_df = calculate_city_average(station_stats, city_gdf)

    if not city_panel_df.empty:
        # 保存面板数据
        output_path = os.path.join(result_dir, FILES_CONFIG["city_panel"])
        city_panel_df.to_csv(output_path, index=False, encoding="utf-8-sig")
        print(f"✅ 市级面板数据已保存：{output_path}")

        # 统计信息
        valid_cities = city_panel_df.dropna(subset=EXTREME_METRICS).shape[0]
        total_cities = len(city_panel_df["city_code"].unique())
        print(f"ℹ️ 有效市级记录数：{valid_cities} 条（覆盖 {total_cities} 个地级市）")

        return city_panel_df
    else:
        print("❌ 未生成任何市级数据")
        return pd.DataFrame()


# ===========================================================================
# 第四步：生成市域热力图（含气象站标注）
# ===========================================================================
def create_city_heatmap(year, metric, city_gdf, city_panel_df, station_stats, plot_dir):
    """生成单年份单指标的市域热力图，标注气象站（算术平均版）"""
    # 创建文件夹
    metric_alias = METRIC_ALIASES[metric]
    sanitized_alias = sanitize_filename(metric_alias)
    year_plot_dir = os.path.join(plot_dir, f"{year}年", sanitized_alias)
    create_directory(year_plot_dir)

    # 筛选数据
    year_city_df = city_panel_df[city_panel_df["year"] == year].copy()
    year_station_df = station_stats[station_stats["year"] == year].copy()

    # 合并边界与指标数据
    city_gdf["city_code_str"] = city_gdf["city_code"].astype(str)
    year_city_df["city_code_str"] = year_city_df["city_code"].astype(str)

    merged_gdf = city_gdf.merge(
        year_city_df[["city_code_str", metric]],
        left_on="city_code_str",
        right_on="city_code_str",
        how="left"
    ).dropna(subset=[metric])

    # 创建图表
    fig, ax = plt.subplots(figsize=PLOT_CONFIG["figsize"], dpi=PLOT_CONFIG["dpi"])

    # 绘制市域热力图
    merged_gdf.plot(
        column=metric,
        ax=ax,
        cmap=PLOT_CONFIG["cmap"],
        legend=True,
        legend_kwds={
            "label": metric_alias,
            "shrink": 0.8,
            "location": "right"
        },
        missing_kwds={
            "color": "lightgrey",
            "label": "无数据"
        }
    )

    # 标注气象站
    if not year_station_df.empty:
        ax.scatter(
            year_station_df["LONGITUDE"],
            year_station_df["LATITUDE"],
            s=PLOT_CONFIG["station_size"],
            c=PLOT_CONFIG["station_color"],
            alpha=PLOT_CONFIG["station_alpha"],
            label="气象站",
            zorder=5
        )

    # 图表样式设置（修改标题为算术平均）
    ax.set_title(
        f"{year}年 中国市域{metric_alias}热力图（算术平均）",
        fontsize=PLOT_CONFIG["title_fontsize"],
        pad=20
    )
    ax.set_xlabel("经度 (°E)", fontsize=PLOT_CONFIG["label_fontsize"])
    ax.set_ylabel("纬度 (°N)", fontsize=PLOT_CONFIG["label_fontsize"])
    ax.tick_params(axis="both", labelsize=PLOT_CONFIG["label_fontsize"] - 2)
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)

    # 图例设置
    if ax.get_legend():
        ax.legend(fontsize=PLOT_CONFIG["legend_fontsize"], loc="upper left")

    # 保存图片（修改文件名为算术平均）
    plot_filename = sanitize_filename(f"{year}年_{sanitized_alias}_算术平均.png")
    plot_path = os.path.join(year_plot_dir, plot_filename)
    plt.tight_layout()
    plt.savefig(plot_path, bbox_inches="tight")
    plt.close()

    return plot_path


def generate_all_heatmaps(city_gdf, city_panel_df, station_stats, plot_dir):
    """生成所有年份和指标的热力图"""
    print("\n===== 生成市域热力图 =====")
    create_directory(plot_dir)

    total_plots = (TARGET_YEARS[1] - TARGET_YEARS[0] + 1) * len(EXTREME_METRICS)
    with tqdm(total=total_plots, desc="生成热力图") as pbar:
        for year in range(TARGET_YEARS[0], TARGET_YEARS[1] + 1):
            for metric in EXTREME_METRICS:
                try:
                    create_city_heatmap(year, metric, city_gdf, city_panel_df, station_stats, plot_dir)
                    pbar.update(1)
                    pbar.set_postfix({"当前": f"{year}年-{METRIC_ALIASES[metric]}"})
                except Exception as e:
                    print(f"\n⚠️ 生成{year}年{METRIC_ALIASES[metric]}热力图失败：{e}")
                    pbar.update(1)

    print(f"✅ 所有热力图已保存至：{os.path.abspath(plot_dir)}")


# ===========================================================================
# 主函数
# ===========================================================================
def main():
    try:
        # 1. 创建文件夹
        result_dir = create_directory(RESULT_DIR)
        plot_dir = create_directory(PLOT_DIR)

        # 2. 基础数据处理（原有逻辑）
        print("\n===== 步骤1：处理基准期数据 =====")
        baseline_df = load_all_data(ROOT_DIR, *BASELINE_YEARS)
        baseline_sorted = sort_baseline_data(baseline_df, result_dir)

        print("\n===== 步骤2：计算极端天气阈值 =====")
        threshold_df = calculate_relative_thresholds(baseline_sorted, result_dir)

        print("\n===== 步骤3：统计站点极端天数 =====")
        target_df = load_all_data(ROOT_DIR, *TARGET_YEARS)
        extreme_stats = count_extreme_days_yearly(target_df, threshold_df, result_dir)

        # 3. 市级数据处理（修改后逻辑）
        print("\n===== 步骤4：加载市级边界数据 =====")
        city_gdf = load_city_boundaries(CITY_SHP_PATH)

        print("\n===== 步骤5：计算地级市算术平均指标 =====")
        city_panel_df = generate_city_panel_data(city_gdf, extreme_stats, result_dir)

        # 4. 生成热力图
        if not city_panel_df.empty:
            generate_all_heatmaps(city_gdf, city_panel_df, extreme_stats, plot_dir)

        # 输出最终统计信息
        print("\n===== 📊 最终结果统计 =====")
        print(f"📅 时间范围：{TARGET_YEARS[0]}-{TARGET_YEARS[1]} 年")
        print(f"📍 涉及站点数：{extreme_stats['STATION'].nunique()}")
        print(f"🏙️ 涉及市级行政区数：{city_gdf['city_code'].nunique()}")
        print(f"📈 有效市级记录数：{len(city_panel_df.dropna(subset=EXTREME_METRICS))}")
        print(f"\n✅ 所有任务完成！")
        print(f"📁 数据文件：{os.path.abspath(result_dir)}")
        print(f"📁 热力图文件：{os.path.abspath(plot_dir)}")

    except Exception as e:
        import traceback
        print(f"\n❌ 执行出错：{str(e)}")
        traceback.print_exc()
        return


if __name__ == "__main__":
    # 首次运行需安装依赖（取消注释）
    # !pip install pandas numpy tqdm geopandas shapely matplotlib
    main()