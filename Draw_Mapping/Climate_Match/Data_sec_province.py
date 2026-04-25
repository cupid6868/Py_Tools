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
PROVINCE_SHP_PATH = r"D:\Data_2\2023年省级\省级.shp"  # 省级边界SHP文件路径
BASELINE_YEARS = (1981, 2010)  # 气候基准期
TARGET_YEARS = (2010, 2024)  # 统计极端天数的目标期
RESULT_DIR = "02_province_weather_heatmaps"  # 结果文件夹名称
PLOT_DIR = "02_province_weather_heatmaps"  # 热力图保存文件夹

# IDW插值参数
IDW_POWER = 2  # 距离倒数的幂次（常用2）
IDW_MAX_DISTANCE = 1500000  # 最大插值距离（米，省级范围放宽至1500公里）

# SHP字段映射（根据你的省级SHP文件调整）
PROVINCE_FIELD_MAP = {
    "code": ["省代码", "CODE", "ADMINCODE", "ID", "省级码", "Province_Code"],  # 省级代码字段
    "name": ["省名称", "NAME", "CNAME", "地名", "省", "Province_Name"]  # 省级名称字段
}

# 图表配置
PLOT_CONFIG = {
    "figsize": (16, 12),
    "dpi": 300,
    "cmap": "RdYlBu_r",
    "station_size": 10,  # 省级图更大，适当调大站点标记
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
    "province_panel": "04_province_weather_panel_data.csv"
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
        'SimHei',           # Windows 备选
        'PingFang SC',      # macOS 系统
        'Hiragino Sans GB', # macOS 备选
        'DejaVu Sans'       # 兜底（无中文时）
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
    """计算两点间哈弗辛距离（米）"""
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
# 第二步：加载省级边界数据
# ===========================================================================
def load_province_boundaries(shp_path):
    """加载省级边界数据，计算省会中心点"""
    print("\n===== 加载省级边界数据 =====")
    gdf = gpd.read_file(shp_path)

    # 设置坐标系
    if gdf.crs is None:
        gdf = gdf.set_crs("EPSG:4326")

    # 匹配省级代码和名称字段
    code_field = next((f for f in PROVINCE_FIELD_MAP["code"] if f in gdf.columns), None)
    if code_field is None:
        code_field = "province_code_temp"
        gdf[code_field] = gdf.index.astype(str)

    name_field = next((f for f in PROVINCE_FIELD_MAP["name"] if f in gdf.columns), None)
    if name_field is None:
        name_field = "province_name_temp"
        gdf[name_field] = [f"省_{i}" for i in gdf.index]

    # 计算省中心点
    utm_crs = gdf.estimate_utm_crs()
    gdf_proj = gdf.to_crs(utm_crs)
    gdf_proj["centroid"] = gdf_proj.geometry.centroid
    gdf["centroid"] = gpd.GeoSeries(gdf_proj["centroid"], crs=utm_crs).to_crs(gdf.crs)

    # 提取中心点经纬度
    gdf["province_lat"] = gdf["centroid"].y
    gdf["province_lon"] = gdf["centroid"].x

    # 整理结果
    province_gdf = gdf.rename(columns={
        code_field: "province_code",
        name_field: "province_name"
    })[["province_code", "province_name", "geometry", "province_lat", "province_lon"]]

    # 统一代码为字符串类型
    province_gdf["province_code"] = province_gdf["province_code"].astype(str)

    print(f"✅ 加载完成：共{len(province_gdf)}个省级行政区")
    return province_gdf

# ===========================================================================
# 第三步：反距离加权（IDW）插值计算省级指标
# ===========================================================================
def idw_interpolation(province_gdf, station_data, year, power=2, max_distance=1500000):
    """
    反距离加权插值计算省级指标
    :param province_gdf: 省级边界GeoDataFrame
    :param station_data: 单年份站点数据
    :param year: 目标年份
    :param power: 距离幂次
    :param max_distance: 最大插值距离（米）
    :return: 省级指标DataFrame
    """
    province_results = []

    # 筛选当前年份的站点数据
    year_station_data = station_data[station_data["year"] == year].copy()
    if year_station_data.empty:
        print(f"⚠️ {year}年无站点数据，跳过")
        return pd.DataFrame()

    # 遍历每个省
    for _, province_row in tqdm(province_gdf.iterrows(), desc=f"IDW插值-{year}年", total=len(province_gdf)):
        province_code = province_row["province_code"]
        province_name = province_row["province_name"]
        province_lat = province_row["province_lat"]
        province_lon = province_row["province_lon"]

        # 计算该省到所有站点的距离
        distances = []
        values = {}

        for _, station_row in year_station_data.iterrows():
            station_lat = station_row["LATITUDE"]
            station_lon = station_row["LONGITUDE"]

            # 计算距离（米）
            dist = haversine_distance(province_lat, province_lon, station_lat, station_lon)

            # 仅使用最大距离内的站点
            if dist < max_distance and dist > 0:
                distances.append(dist)
                # 存储各指标值
                for metric in EXTREME_METRICS:
                    if metric not in values:
                        values[metric] = []
                    values[metric].append(station_row[metric])

        # 计算IDW加权平均
        province_values = {"province_code": province_code, "province_name": province_name, "year": year}

        if distances:
            # 计算权重（距离倒数的幂次）
            weights = 1 / (np.array(distances) ** power)
            weights = weights / np.sum(weights)  # 归一化

            # 对每个指标计算加权平均
            for metric in EXTREME_METRICS:
                if values[metric]:
                    weighted_avg = np.sum(np.array(values[metric]) * weights)
                    province_values[metric] = round(weighted_avg, 2)
                else:
                    province_values[metric] = np.nan
        else:
            # 无有效站点，标记为NaN
            for metric in EXTREME_METRICS:
                province_values[metric] = np.nan

        province_results.append(province_values)

    # 整理结果
    province_df = pd.DataFrame(province_results)
    return province_df


def generate_province_panel_data(province_gdf, station_stats, result_dir):
    """生成省级气象面板数据"""
    print("\n===== 生成省级气象面板数据 =====")
    all_province_data = []

    # 遍历目标年份
    for year in range(TARGET_YEARS[0], TARGET_YEARS[1] + 1):
        province_year_data = idw_interpolation(province_gdf, station_stats, year, IDW_POWER, IDW_MAX_DISTANCE)
        if not province_year_data.empty:
            all_province_data.append(province_year_data)

    # 合并所有年份数据
    if all_province_data:
        province_panel_df = pd.concat(all_province_data, ignore_index=True)
        # 排序
        province_panel_df = province_panel_df.sort_values(["year", "province_code"]).reset_index(drop=True)

        # 保存面板数据
        output_path = os.path.join(result_dir, FILES_CONFIG["province_panel"])
        province_panel_df.to_csv(output_path, index=False, encoding="utf-8-sig")
        print(f"✅ 省级面板数据已保存：{output_path}")

        # 统计信息
        valid_provinces = province_panel_df.dropna(subset=EXTREME_METRICS).shape[0]
        total_provinces = province_panel_df.shape[0]
        print(f"ℹ️ 有效省级记录数：{valid_provinces}/{total_provinces}")

        return province_panel_df
    else:
        print("❌ 未生成任何省级数据")
        return pd.DataFrame()

# ===========================================================================
# 第四步：生成省级热力图（含气象站标注）
# ===========================================================================
def create_province_heatmap(year, metric, province_gdf, province_panel_df, station_stats, plot_dir):
    """生成单年份单指标的省级热力图，标注气象站"""
    # 创建文件夹
    metric_alias = METRIC_ALIASES[metric]
    sanitized_alias = sanitize_filename(metric_alias)
    year_plot_dir = os.path.join(plot_dir, f"{year}年", sanitized_alias)
    create_directory(year_plot_dir)

    # 筛选数据
    year_province_df = province_panel_df[province_panel_df["year"] == year].copy()
    year_station_df = station_stats[station_stats["year"] == year].copy()

    # 合并边界与指标数据
    province_gdf["province_code_str"] = province_gdf["province_code"].astype(str)
    year_province_df["province_code_str"] = year_province_df["province_code"].astype(str)

    merged_gdf = province_gdf.merge(
        year_province_df[["province_code_str", metric]],
        left_on="province_code_str",
        right_on="province_code_str",
        how="left"
    ).dropna(subset=[metric])

    # 创建图表
    fig, ax = plt.subplots(figsize=PLOT_CONFIG["figsize"], dpi=PLOT_CONFIG["dpi"])

    # 绘制省级热力图
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

    # 图表样式设置
    ax.set_title(
        f"{year}年 中国省级{metric_alias}热力图（IDW插值）",
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

    # 保存图片
    plot_filename = sanitize_filename(f"{year}年_{sanitized_alias}_IDW插值.png")
    plot_path = os.path.join(year_plot_dir, plot_filename)
    plt.tight_layout()
    plt.savefig(plot_path, bbox_inches="tight")
    plt.close()

    return plot_path


def generate_all_heatmaps(province_gdf, province_panel_df, station_stats, plot_dir):
    """生成所有年份和指标的热力图"""
    print("\n===== 生成省级热力图 =====")
    create_directory(plot_dir)

    total_plots = (TARGET_YEARS[1] - TARGET_YEARS[0] + 1) * len(EXTREME_METRICS)
    with tqdm(total=total_plots, desc="生成热力图") as pbar:
        for year in range(TARGET_YEARS[0], TARGET_YEARS[1] + 1):
            for metric in EXTREME_METRICS:
                try:
                    create_province_heatmap(year, metric, province_gdf, province_panel_df, station_stats, plot_dir)
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

        # 3. 省级数据处理
        print("\n===== 步骤4：加载省级边界数据 =====")
        province_gdf = load_province_boundaries(PROVINCE_SHP_PATH)

        print("\n===== 步骤5：IDW插值生成省级面板数据 =====")
        province_panel_df = generate_province_panel_data(province_gdf, extreme_stats, result_dir)

        # 4. 生成热力图
        if not province_panel_df.empty:
            generate_all_heatmaps(province_gdf, province_panel_df, extreme_stats, plot_dir)

        # 输出最终统计信息
        print("\n===== 📊 最终结果统计 =====")
        print(f"📅 时间范围：{TARGET_YEARS[0]}-{TARGET_YEARS[1]} 年")
        print(f"📍 涉及站点数：{extreme_stats['STATION'].nunique()}")
        print(f"🏙️ 涉及省级行政区数：{province_gdf['province_code'].nunique()}")
        print(f"📈 有效省级记录数：{len(province_panel_df.dropna(subset=EXTREME_METRICS))}")
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