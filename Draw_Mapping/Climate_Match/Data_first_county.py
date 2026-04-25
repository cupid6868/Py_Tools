import os
import glob
import pandas as pd
import numpy as np
import geopandas as gpd
import matplotlib.pyplot as plt
import warnings
from tqdm import tqdm
from shapely.geometry import Point
import re

# 忽略无关警告
warnings.filterwarnings("ignore")

# ===================== 全局配置（根据需求修改） =====================
# 数据路径
ROOT_DIR = r"D:\Data_1"  # 站点年度数据根目录
SHP_PATH = r"D:\Data_2\2023年县级\县级.shp"  # 县级边界SHP文件
TARGET_START_YEAR = 2010  # 起始年份
TARGET_END_YEAR = 2024  # 结束年份
RESULT_DIR = "01_county_weather_plots"  # 结果文件夹
PLOT_DIR = "01_county_weather_plots"  # 热力图保存文件夹
# IDW参数
IDW_POWER = 2  # 距离幂次（通常取2）
MIN_STATIONS = 1  # 每个县至少需要的站点数
# 8个核心指标（含中文别名，用于图表标题，避免非法字符）
METRICS = {
    "最高温度(°C)": "年平均最高气温",
    "最低温度(°C)": "年平均最低气温",
    "极端高温天数(天)": "极端高温天数_35℃以上",
    "霜冻天数(天)": "霜冻天数_0℃以下",
    "年降水量(毫米)": "年总降水量",
    "降水强度(毫米/天)": "降水强度",
    "特强降水天数(天)": "特强降水天数_25mm以上",
    "连续干旱天数(天)": "最大连续干旱天数"
}
# SHP文件字段映射
COUNTY_FIELD_MAP = {
    "code": ["CODE", "ADMINCODE", "县级码", "行政代码", "ID"],
    "name": ["NAME", "地名", "名称", "行政区名", "CNAME"]
}
# 图表样式配置 - 核心优化：缩小站点尺寸
PLOT_CONFIG = {
    "figsize": (16, 10),  # 图表尺寸
    "dpi": 300,  # 分辨率
    "cmap": "RdYlBu_r",  # 配色方案（可换：coolwarm, viridis, YlOrRd等）
    "station_size": 8,  # 🔴 核心修改：站点尺寸从20缩小到8
    "station_color": "black",  # 站点标注颜色
    "station_alpha": 0.5,  # 🔴 核心修改：透明度从0.7降低到0.5
    "station_show": True,  # 是否显示站点（False则隐藏所有站点）
    "station_filter_distance": 5000,  # 过滤距离（米）：小于该距离只显示1个站点，避免密集重叠
    "title_fontsize": 16,  # 标题字体大小
    "label_fontsize": 12,  # 标签字体大小
    "legend_fontsize": 10,  # 图例字体大小
    "legend_position": "right",  # 图例位置：right/left/bottom/top
    "colorbar_shrink": 0.8  # 颜色条缩放比例
}


# ===========================================================================

def create_directory(dir_name):
    """创建文件夹（多层级）"""
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
        print(f"✅ 已创建文件夹：{os.path.abspath(dir_name)}")
    else:
        print(f"ℹ️ 文件夹已存在：{os.path.abspath(dir_name)}")
    return dir_name


def sanitize_filename(filename):
    """清理文件名中的非法字符"""
    # 替换Windows非法字符
    illegal_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(illegal_chars, '_', filename)
    # 移除多余下划线
    sanitized = re.sub('_+', '_', sanitized)
    return sanitized.strip('_')


# ---------------------- 工具函数：哈弗辛距离计算 ----------------------
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


# ---------------------- 新增工具函数：过滤密集站点 ----------------------
def filter_dense_stations(station_df, min_distance=5000):
    """
    过滤密集站点：保留距离大于min_distance的站点，避免重叠
    :param station_df: 站点数据DataFrame（含纬度/经度列）
    :param min_distance: 最小距离（米）
    :return: 过滤后的站点DataFrame
    """
    if len(station_df) <= 1:
        return station_df

    # 按距离筛选站点
    kept_indices = []
    for idx, row in station_df.iterrows():
        # 检查当前站点与已保留站点的距离
        keep = True
        lat1, lon1 = row["纬度(°)"], row["经度(°)"]
        for kept_idx in kept_indices:
            kept_row = station_df.loc[kept_idx]
            lat2, lon2 = kept_row["纬度(°)"], kept_row["经度(°)"]
            dist = haversine_distance(lat1, lon1, lat2, lon2)
            if dist < min_distance:
                keep = False
                break
        if keep:
            kept_indices.append(idx)

    return station_df.loc[kept_indices].copy()


# ---------------------- 第一步：站点级指标计算 ----------------------
def load_yearly_station_data(root_dir, year):
    """加载指定年份站点日数据"""
    year_dir = os.path.join(root_dir, str(year))
    if not os.path.exists(year_dir):
        print(f"警告：{year} 年文件夹不存在，跳过")
        return {}

    csv_files = glob.glob(os.path.join(year_dir, "*.csv"))
    if not csv_files:
        print(f"警告：{year} 年无CSV数据文件，跳过")
        return {}

    station_data = {}
    for file in csv_files:
        try:
            df = pd.read_csv(
                file,
                parse_dates=["DATE"],
                usecols=["STATION", "NAME", "DATE", "LATITUDE", "LONGITUDE", "MIN", "MAX", "PRCP"]
            )
            # 数据清洗
            df = df.dropna(subset=["STATION", "NAME", "DATE", "LATITUDE", "LONGITUDE", "MIN", "MAX", "PRCP"])
            df["MIN"] = df["MIN"].replace(9999.9, np.nan)
            df["MAX"] = df["MAX"].replace(9999.9, np.nan)
            df["PRCP"] = df["PRCP"].replace(99.99, np.nan)
            df = df.dropna(subset=["MIN", "MAX", "PRCP"])
            df = df[(df["PRCP"] >= 0) & (df["MIN"].notna()) & (df["MAX"].notna())]

            # 单位转换
            df["MIN_C"] = (df["MIN"] - 32) * 5 / 9
            df["MAX_C"] = (df["MAX"] - 32) * 5 / 9
            df["PRCP_mm"] = df["PRCP"] * 25.4

            # 按站点分组
            for station, station_df in df.groupby("STATION"):
                station_df = station_df.sort_values("DATE").reset_index(drop=True)
                station_name = station_df["NAME"].iloc[0].strip()
                station_lat = station_df["LATITUDE"].iloc[0]
                station_lon = station_df["LONGITUDE"].iloc[0]
                station_data[f"{year}_{station}"] = {
                    "year": year,
                    "station": station,
                    "name": station_name,
                    "latitude": station_lat,
                    "longitude": station_lon,
                    "data": station_df
                }
        except Exception as e:
            print(f"读取文件 {file} 失败：{str(e)}，跳过")

    return station_data


def calculate_drought_days(prcp_series):
    """计算连续干旱天数"""
    drought_mask = prcp_series < 1
    drought_groups = (drought_mask != drought_mask.shift()).cumsum()
    drought_lengths = drought_groups[drought_mask].value_counts()
    return drought_lengths.max() if not drought_lengths.empty else 0


def compute_station_metrics(root_dir, start_year, end_year, result_dir):
    """计算站点级年度指标并保存"""
    print("\n===== 第一步：计算站点级气象指标 =====")
    all_station_data = {}
    for year in range(start_year, end_year + 1):
        print(f"\n正在加载 {year} 年数据...")
        yearly_data = load_yearly_station_data(root_dir, year)
        all_station_data.update(yearly_data)

    if not all_station_data:
        print("❌ 错误：未加载到任何有效站点数据！")
        return pd.DataFrame()

    # 计算站点指标
    metrics_list = []
    for key, data in tqdm(all_station_data.items(), desc="计算站点指标"):
        year = data["year"]
        station = data["station"]
        station_name = data["name"]
        station_lat = data["latitude"]
        station_lon = data["longitude"]
        df = data["data"]

        # 核心指标计算
        max_temp_mean = df["MAX_C"].mean()
        min_temp_mean = df["MIN_C"].mean()
        extreme_high_days = (df["MAX_C"] > 35).sum()
        frost_days = (df["MIN_C"] < 0).sum()
        annual_precip = df["PRCP_mm"].sum()
        rain_days = (df["PRCP_mm"] > 0).sum()
        precip_intensity = annual_precip / rain_days if rain_days > 0 else 0
        heavy_rain_days = (df["PRCP_mm"] >= 25).sum()
        drought_days = calculate_drought_days(df["PRCP_mm"])

        metrics = {
            "年份": year,
            "气象站编号": station,
            "气象站名称": station_name,
            "纬度(°)": round(station_lat, 6),
            "经度(°)": round(station_lon, 6),
            "最高温度(°C)": round(max_temp_mean, 2),
            "最低温度(°C)": round(min_temp_mean, 2),
            "极端高温天数(天)": int(extreme_high_days),
            "霜冻天数(天)": int(frost_days),
            "年降水量(毫米)": round(annual_precip, 2),
            "降水强度(毫米/天)": round(precip_intensity, 2),
            "特强降水天数(天)": int(heavy_rain_days),
            "连续干旱天数(天)": int(drought_days)
        }
        metrics_list.append(metrics)

    # 保存站点数据
    metrics_df = pd.DataFrame(metrics_list)
    metrics_df = metrics_df.sort_values(["年份", "气象站编号"]).reset_index(drop=True)
    output_path = os.path.join(result_dir, "01_weather_stats_2010_2024.csv")
    metrics_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"✅ 站点级指标已保存：{os.path.abspath(output_path)}")
    return metrics_df


# ---------------------- 第二步：县级边界加载 ----------------------
def load_county_boundaries(shp_path):
    """加载县级边界SHP"""
    # 1. 读取SHP完整数据（几何+属性）
    print("\n===== 第二步：加载县级边界数据 =====")
    gdf = gpd.read_file(shp_path)

    # 修复CRS
    if gdf.crs is None:
        gdf = gdf.set_crs("EPSG:4326")
    utm_crs = gdf.estimate_utm_crs()
    gdf_proj = gdf.to_crs(utm_crs)
    gdf_proj["centroid_proj"] = gdf_proj.geometry.centroid
    gdf["centroid"] = gpd.GeoSeries(gdf_proj["centroid_proj"], crs=utm_crs).to_crs(gdf.crs)

    # 提取中心点
    gdf["county_lat"] = gdf["centroid"].y
    gdf["county_lon"] = gdf["centroid"].x

    # 匹配代码/名称字段
    code_field = None
    for field in COUNTY_FIELD_MAP["code"]:
        if field in gdf.columns:
            code_field = field
            break
    if code_field is None:
        code_field = "county_code_temp"
        gdf[code_field] = gdf.index.astype(str)

    name_field = None
    for field in COUNTY_FIELD_MAP["name"]:
        if field in gdf.columns:
            name_field = field
            break
    if name_field is None:
        gdf["county_name_temp"] = [f"县_{i}" for i in gdf.index]
        name_field = "county_name_temp"

    # 重命名
    county_gdf = gdf.rename(columns={
        code_field: "county_code",
        name_field: "county_name"
    })[["county_code", "county_name", "geometry", "county_lat", "county_lon"]]

    print(f"✅ 加载完成：共{len(county_gdf)}个县级行政区")
    return county_gdf


# ---------------------- 第三步：IDW插值与县域聚合 ----------------------
def idw_interpolate_county(county_centroid, stations_df, power=2):
    """IDW插值"""
    if len(stations_df) < MIN_STATIONS:
        return {metric: np.nan for metric in METRICS.keys()}

    county_lat, county_lon = county_centroid
    distances = []
    for _, station in stations_df.iterrows():
        dist = haversine_distance(county_lat, county_lon, station["纬度(°)"], station["经度(°)"])
        distances.append(dist)
    distances = np.array(distances)
    distances = np.where(distances == 0, 1e-6, distances)
    weights = 1 / (distances ** power)

    result = {}
    for metric in METRICS.keys():
        values = stations_df[metric].values
        weighted_sum = np.sum(values * weights)
        weight_sum = np.sum(weights)
        result[metric] = round(weighted_sum / weight_sum, 2) if weight_sum != 0 else np.nan
    return result


def aggregate_to_county(station_df, county_gdf, result_dir):
    """生成县级面板数据"""
    print("\n===== 第三步：生成县级气象面板数据 =====")
    county_results = []

    for year in tqdm(range(TARGET_START_YEAR, TARGET_END_YEAR + 1), desc="处理年度数据"):
        year_stations = station_df[station_df["年份"] == year].copy()
        if year_stations.empty:
            continue

        for _, county in county_gdf.iterrows():
            idw_result = idw_interpolate_county(
                (county["county_lat"], county["county_lon"]),
                year_stations,
                power=IDW_POWER
            )

            county_row = {
                "年份": year,
                "县域代码": county["county_code"],
                "县域名称": county["county_name"],
                "县域纬度(°)": round(county["county_lat"], 6),
                "县域经度(°)": round(county["county_lon"], 6)
            }
            county_row.update(idw_result)
            county_results.append(county_row)

    # 保存县级数据
    county_df = pd.DataFrame(county_results)
    output_path = os.path.join(result_dir, "02_county_weather_panel_2010_2024.csv")
    county_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"✅ 县级面板数据已保存：{os.path.abspath(output_path)}")
    return county_df, county_gdf


# ---------------------- 第四步：生成热力图（优化站点显示） ----------------------
def create_county_heatmap(year, metric, county_gdf, county_df, station_df, plot_dir):
    """
    生成单年份单指标的县域热力图（优化站点显示：缩小尺寸+过滤密集站点）
    """
    # 清理指标别名，避免非法字符
    metric_alias = sanitize_filename(METRICS[metric])
    year_plot_dir = os.path.join(plot_dir, f"{year}年", metric_alias)
    create_directory(year_plot_dir)

    # 筛选数据
    year_county_df = county_df[county_df["年份"] == year].copy()
    year_station_df = station_df[station_df["年份"] == year].copy()

    # 合并县级边界和指标数据
    county_gdf_with_data = county_gdf.merge(
        year_county_df[["县域代码", metric]],
        left_on="county_code",
        right_on="县域代码",
        how="left"
    )

    # 过滤无数据的县域
    county_gdf_with_data = county_gdf_with_data.dropna(subset=[metric])

    # 创建图表
    fig, ax = plt.subplots(figsize=PLOT_CONFIG["figsize"], dpi=PLOT_CONFIG["dpi"])

    # 绘制县域热力图
    im = county_gdf_with_data.plot(
        column=metric,
        ax=ax,
        cmap=PLOT_CONFIG["cmap"],
        legend=True,
        legend_kwds={
            "label": f"{METRICS[metric]}",
            "shrink": PLOT_CONFIG["colorbar_shrink"],
            "location": PLOT_CONFIG["legend_position"]
        },
        missing_kwds={
            "color": "lightgrey",  # 无数据县域颜色
            "label": "无数据"
        }
    )

    # 设置图例字体大小
    cbar = ax.get_figure().get_axes()[1]  # 获取颜色条轴
    cbar.tick_params(labelsize=PLOT_CONFIG["legend_fontsize"])

    # 标注气象站（优化：可选显示+过滤密集站点+缩小尺寸）
    if PLOT_CONFIG["station_show"] and not year_station_df.empty:
        # 过滤密集站点，避免重叠
        filtered_stations = filter_dense_stations(
            year_station_df,
            min_distance=PLOT_CONFIG["station_filter_distance"]
        )

        # 绘制站点（缩小尺寸+低透明度）
        ax.scatter(
            filtered_stations["经度(°)"],
            filtered_stations["纬度(°)"],
            s=PLOT_CONFIG["station_size"],  # 缩小后的尺寸
            c=PLOT_CONFIG["station_color"],
            alpha=PLOT_CONFIG["station_alpha"],  # 降低透明度
            label="气象站",
            zorder=5  # 确保站点在热力图上方，但不遮挡太多
        )

    # 图表样式设置
    ax.set_title(
        f"{year}年 中国县域{METRICS[metric]}热力图",
        fontsize=PLOT_CONFIG["title_fontsize"],
        pad=20
    )
    ax.set_xlabel("经度 (°E)", fontsize=PLOT_CONFIG["label_fontsize"])
    ax.set_ylabel("纬度 (°N)", fontsize=PLOT_CONFIG["label_fontsize"])
    ax.tick_params(axis="both", labelsize=PLOT_CONFIG["label_fontsize"] - 2)

    # 设置图例字体
    if ax.get_legend():
        ax.legend(fontsize=PLOT_CONFIG["legend_fontsize"], loc="upper left")

    # 移除边框
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_visible(False)

    # 保存图表（清理文件名）
    plot_filename = sanitize_filename(f"{year}年_{metric_alias}.png")
    plot_path = os.path.join(year_plot_dir, plot_filename)
    plt.tight_layout()
    plt.savefig(plot_path, bbox_inches="tight")
    plt.close()

    return plot_path


def generate_all_heatmaps(county_df, county_gdf, station_df, plot_dir):
    """生成所有年份+所有指标的热力图"""
    print("\n===== 第四步：生成县域热力图 =====")
    create_directory(plot_dir)

    # 遍历年份和指标
    total_plots = (TARGET_END_YEAR - TARGET_START_YEAR + 1) * len(METRICS)
    with tqdm(total=total_plots, desc="生成热力图") as pbar:
        for year in range(TARGET_START_YEAR, TARGET_END_YEAR + 1):
            for metric in METRICS.keys():
                try:
                    plot_path = create_county_heatmap(
                        year, metric, county_gdf, county_df, station_df, plot_dir
                    )
                    pbar.update(1)
                    pbar.set_postfix({"当前生成": f"{year}年-{METRICS[metric]}"})
                except Exception as e:
                    print(f"\n⚠️ 生成{year}年{METRICS[metric]}热力图失败：{str(e)}")
                    pbar.update(1)

    print(f"✅ 所有热力图已保存至：{os.path.abspath(plot_dir)}")


# ---------------------- 主函数 ----------------------
def main():
    try:
        # 1. 创建文件夹
        result_dir = create_directory(RESULT_DIR)

        # 2. 计算站点指标
        station_df = compute_station_metrics(ROOT_DIR, TARGET_START_YEAR, TARGET_END_YEAR, result_dir)
        if station_df.empty:
            return

        # 3. 加载县级边界
        county_gdf = load_county_boundaries(SHP_PATH)

        # 4. 生成县级面板数据
        county_df, county_gdf = aggregate_to_county(station_df, county_gdf, result_dir)

        # 5. 生成热力图
        generate_all_heatmaps(county_df, county_gdf, station_df, PLOT_DIR)

        # 输出统计信息
        print(f"\n===== 最终结果统计 =====")
        print(f"📅 时间范围：{TARGET_START_YEAR}-{TARGET_END_YEAR} 年")
        print(f"📍 涉及站点数：{station_df['气象站编号'].nunique()}")
        print(f"🏡 涉及县域数：{county_df['县域代码'].nunique()}")
        print(f"📊 有效县级记录数：{len(county_df.dropna(subset=METRICS.keys()))} / {len(county_df)}")
        print(f"📈 生成热力图数量：{(TARGET_END_YEAR - TARGET_START_YEAR + 1) * len(METRICS)}")
        print(f"\n✅ 所有任务完成！")
        print(f"📁 数据文件：{os.path.abspath(RESULT_DIR)}")
        print(f"📁 图表文件：{os.path.abspath(PLOT_DIR)}")

    except Exception as e:
        import traceback
        print(f"\n❌ 执行出错：{str(e)}")
        traceback.print_exc()
        return


if __name__ == "__main__":
    # 安装依赖（首次运行取消注释）
    # !pip install pandas numpy tqdm geopandas shapely matplotlib
    main()