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
SHP_PATH = r"D:\Data_2\2023年地级\地级.shp"  # 地级边界SHP文件
TARGET_START_YEAR = 2010  # 起始年份
TARGET_END_YEAR = 2024  # 结束年份
RESULT_DIR = "01_city_weather_plots"  # 地级结果文件夹
PLOT_DIR = "01_city_weather_plots"  # 地级热力图保存文件夹
# IDW参数
IDW_POWER = 2  # 距离幂次（通常取2，值越大近站点权重越高）
MIN_STATIONS = 1  # 每个地级市至少需要的站点数
# 8个核心指标（含中文别名，用于图表标题/文件名）
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
# SHP文件字段映射（适配常见地级SHP字段）
PREFECTURE_FIELD_MAP = {
    "code": ["CODE", "ADMINCODE", "地级码", "行政代码", "ID", "GID"],
    "name": ["NAME", "地名", "名称", "地级市名", "CNAME", "PrefName"]
}
# 图表样式配置
PLOT_CONFIG = {
    "figsize": (16, 10),  # 图表尺寸
    "dpi": 300,  # 分辨率
    "cmap": "RdYlBu_r",  # 配色方案（coolwarm/viridis/YlOrRd可选）
    "station_size": 6,  # 站点标记大小
    "station_color": "#333333",  # 站点颜色（深灰更柔和）
    "station_alpha": 0.4,  # 站点透明度
    "station_show": True,  # 是否显示站点
    "station_filter_distance": 8000,  # 过滤密集站点（米）
    "title_fontsize": 16,  # 标题字体大小
    "label_fontsize": 12,  # 坐标轴标签大小
    "legend_fontsize": 10,  # 图例字体大小
    "legend_position": "right",  # 图例位置
    "colorbar_shrink": 0.8  # 颜色条缩放比例
}


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
    sanitized = re.sub('_+', '_', sanitized)  # 合并连续下划线
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


def filter_dense_stations(station_df, min_distance=5000):
    """过滤密集站点：保留距离大于阈值的站点，避免重叠"""
    if len(station_df) <= 1:
        return station_df

    kept_indices = []
    for idx, row in station_df.iterrows():
        keep = True
        lat1, lon1 = row["纬度(°)"], row["经度(°)"]
        # 检查与已保留站点的距离
        for kept_idx in kept_indices:
            kept_row = station_df.loc[kept_idx]
            dist = haversine_distance(lat1, lon1, kept_row["纬度(°)"], kept_row["经度(°)"])
            if dist < min_distance:
                keep = False
                break
        if keep:
            kept_indices.append(idx)
    return station_df.loc[kept_indices].copy()


# ===========================================================================
# 第一步：站点级气象指标计算
# ===========================================================================
def load_yearly_station_data(root_dir, year):
    """加载指定年份的站点日数据并预处理"""
    year_dir = os.path.join(root_dir, str(year))
    if not os.path.exists(year_dir):
        print(f"⚠️ 警告：{year}年文件夹不存在，跳过")
        return {}

    csv_files = glob.glob(os.path.join(year_dir, "*.csv"))
    if not csv_files:
        print(f"⚠️ 警告：{year}年无CSV数据文件，跳过")
        return {}

    station_data = {}
    for file in csv_files:
        try:
            # 读取核心字段
            df = pd.read_csv(
                file,
                parse_dates=["DATE"],
                usecols=["STATION", "NAME", "DATE", "LATITUDE", "LONGITUDE", "MIN", "MAX", "PRCP"]
            )
            # 数据清洗
            df = df.dropna(subset=["STATION", "NAME", "DATE", "LATITUDE", "LONGITUDE", "MIN", "MAX", "PRCP"])
            df = df.replace({9999.9: np.nan, 99.99: np.nan})  # 替换缺测值
            df = df.dropna(subset=["MIN", "MAX", "PRCP"])
            df = df[(df["PRCP"] >= 0)]  # 过滤异常降水量

            # 单位转换（华氏→摄氏，英寸→毫米）
            df["MIN_C"] = (df["MIN"] - 32) * 5 / 9
            df["MAX_C"] = (df["MAX"] - 32) * 5 / 9
            df["PRCP_mm"] = df["PRCP"] * 25.4

            # 按站点分组存储
            for station, station_df in df.groupby("STATION"):
                station_df = station_df.sort_values("DATE").reset_index(drop=True)
                station_data[f"{year}_{station}"] = {
                    "year": year,
                    "station": station,
                    "name": station_df["NAME"].iloc[0].strip(),
                    "latitude": station_df["LATITUDE"].iloc[0],
                    "longitude": station_df["LONGITUDE"].iloc[0],
                    "data": station_df
                }
        except Exception as e:
            print(f"❌ 读取文件 {file} 失败：{str(e)}，跳过")

    return station_data


def calculate_drought_days(prcp_series):
    """计算最大连续干旱天数（日降水量<1mm视为干旱）"""
    drought_mask = prcp_series < 1
    drought_groups = (drought_mask != drought_mask.shift()).cumsum()
    drought_lengths = drought_groups[drought_mask].value_counts()
    return drought_lengths.max() if not drought_lengths.empty else 0


def compute_station_metrics(root_dir, start_year, end_year, result_dir):
    """计算所有年份的站点级年度指标并保存"""
    print("\n===== 第一步：计算站点级气象指标 =====")
    all_station_data = {}
    for year in range(start_year, end_year + 1):
        print(f"\n📅 正在加载 {year} 年数据...")
        yearly_data = load_yearly_station_data(root_dir, year)
        all_station_data.update(yearly_data)

    if not all_station_data:
        print("❌ 错误：未加载到任何有效站点数据！")
        return pd.DataFrame()

    # 计算年度指标
    metrics_list = []
    for key, data in tqdm(all_station_data.items(), desc="🧮 计算站点指标"):
        df = data["data"]
        metrics = {
            "年份": data["year"],
            "气象站编号": data["station"],
            "气象站名称": data["name"],
            "纬度(°)": round(data["latitude"], 6),
            "经度(°)": round(data["longitude"], 6),
            "最高温度(°C)": round(df["MAX_C"].mean(), 2),
            "最低温度(°C)": round(df["MIN_C"].mean(), 2),
            "极端高温天数(天)": int((df["MAX_C"] > 35).sum()),
            "霜冻天数(天)": int((df["MIN_C"] < 0).sum()),
            "年降水量(毫米)": round(df["PRCP_mm"].sum(), 2),
            "降水强度(毫米/天)": round(df["PRCP_mm"].sum() / (df["PRCP_mm"] > 0).sum(), 2) if (df[
                                                                                                   "PRCP_mm"] > 0).sum() > 0 else 0,
            "特强降水天数(天)": int((df["PRCP_mm"] >= 25).sum()),
            "连续干旱天数(天)": calculate_drought_days(df["PRCP_mm"])
        }
        metrics_list.append(metrics)

    # 保存结果
    metrics_df = pd.DataFrame(metrics_list)
    metrics_df = metrics_df.sort_values(["年份", "气象站编号"]).reset_index(drop=True)
    output_path = os.path.join(result_dir, "01_station_weather_stats_2010_2024.csv")
    metrics_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"✅ 站点级指标已保存：{os.path.abspath(output_path)}")
    return metrics_df


# ===========================================================================
# 第二步：加载地级边界数据
# ===========================================================================
def load_prefecture_boundaries(shp_path):
    """加载地级边界SHP，提取中心点坐标"""
    print("\n===== 第二步：加载地级边界数据 =====")
    gdf = gpd.read_file(shp_path)

    # 修复坐标系
    if gdf.crs is None:
        gdf = gdf.set_crs("EPSG:4326")  # WGS84

    # 计算几何中心点（UTM投影下更准确）
    utm_crs = gdf.estimate_utm_crs()
    gdf_proj = gdf.to_crs(utm_crs)
    gdf_proj["centroid_proj"] = gdf_proj.geometry.centroid
    gdf["centroid"] = gpd.GeoSeries(gdf_proj["centroid_proj"], crs=utm_crs).to_crs(gdf.crs)

    # 提取中心点经纬度
    gdf["pref_lat"] = gdf["centroid"].y
    gdf["pref_lon"] = gdf["centroid"].x

    # 匹配代码/名称字段
    code_field = next((f for f in PREFECTURE_FIELD_MAP["code"] if f in gdf.columns), None)
    if code_field is None:
        code_field = "pref_code_temp"
        gdf[code_field] = gdf.index.astype(str)

    name_field = next((f for f in PREFECTURE_FIELD_MAP["name"] if f in gdf.columns), None)
    if name_field is None:
        name_field = "pref_name_temp"
        gdf[name_field] = [f"地级市_{i}" for i in gdf.index]

    # 整理结果
    prefecture_gdf = gdf.rename(columns={
        code_field: "pref_code",
        name_field: "pref_name"
    })[["pref_code", "pref_name", "geometry", "pref_lat", "pref_lon"]]

    print(f"✅ 加载完成：共{len(prefecture_gdf)}个地级市")
    return prefecture_gdf


# ===========================================================================
# 第三步：IDW插值生成地级面板数据
# ===========================================================================
def idw_interpolate_prefecture(pref_centroid, stations_df, power=2):
    """
    反距离权重插值：计算地级市中心点的气象指标
    :param pref_centroid: 地级市中心点 (lat, lon)
    :param stations_df: 年度站点数据
    :param power: 距离幂次
    :return: 各指标插值结果
    """
    if len(stations_df) < MIN_STATIONS:
        return {metric: np.nan for metric in METRICS.keys()}

    pref_lat, pref_lon = pref_centroid
    # 计算所有站点到地级市的距离
    distances = np.array([
        haversine_distance(pref_lat, pref_lon, row["纬度(°)"], row["经度(°)"])
        for _, row in stations_df.iterrows()
    ])
    distances = np.where(distances == 0, 1e-6, distances)  # 避免除以0

    # 计算权重：1/距离^power
    weights = 1 / (distances ** power)

    # 加权平均计算各指标
    result = {}
    for metric in METRICS.keys():
        values = stations_df[metric].values
        weighted_sum = np.sum(values * weights)
        weight_sum = np.sum(weights)
        result[metric] = round(weighted_sum / weight_sum, 2) if weight_sum != 0 else np.nan

    return result


def aggregate_to_prefecture(station_df, prefecture_gdf, result_dir):
    """生成地级气象面板数据"""
    print("\n===== 第三步：生成地级气象面板数据 =====")
    prefecture_results = []

    for year in tqdm(range(TARGET_START_YEAR, TARGET_END_YEAR + 1), desc="📍 处理年度数据"):
        year_stations = station_df[station_df["年份"] == year].copy()
        if year_stations.empty:
            continue

        # 对每个地级市进行插值
        for _, prefecture in prefecture_gdf.iterrows():
            idw_result = idw_interpolate_prefecture(
                (prefecture["pref_lat"], prefecture["pref_lon"]),
                year_stations,
                power=IDW_POWER
            )
            # 组装结果
            prefecture_row = {
                "年份": year,
                "地级市代码": prefecture["pref_code"],
                "地级市名称": prefecture["pref_name"],
                "地级市纬度(°)": round(prefecture["pref_lat"], 6),
                "地级市经度(°)": round(prefecture["pref_lon"], 6)
            }
            prefecture_row.update(idw_result)
            prefecture_results.append(prefecture_row)

    # 保存结果
    prefecture_df = pd.DataFrame(prefecture_results)
    prefecture_df = prefecture_df.sort_values(["年份", "地级市代码"]).reset_index(drop=True)
    output_path = os.path.join(result_dir, "02_prefecture_weather_panel_2010_2024.csv")
    prefecture_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"✅ 地级面板数据已保存：{os.path.abspath(output_path)}")
    return prefecture_df


# ===========================================================================
# 第四步：生成地级热力图
# ===========================================================================
def create_prefecture_heatmap(year, metric, prefecture_gdf, prefecture_df, station_df, plot_dir):
    """生成单年份单指标的地级热力图"""
    # 创建文件夹
    metric_alias = sanitize_filename(METRICS[metric])
    year_plot_dir = os.path.join(plot_dir, f"{year}年", metric_alias)
    create_directory(year_plot_dir)

    # 筛选数据
    year_pref_df = prefecture_df[prefecture_df["年份"] == year].copy()
    year_station_df = station_df[station_df["年份"] == year].copy()

    # 合并边界与指标数据
    prefecture_gdf_with_data = prefecture_gdf.merge(
        year_pref_df[["地级市代码", metric]],
        left_on="pref_code",
        right_on="地级市代码",
        how="left"
    ).dropna(subset=[metric])

    # 创建图表
    fig, ax = plt.subplots(figsize=PLOT_CONFIG["figsize"], dpi=PLOT_CONFIG["dpi"])

    # 绘制热力图
    prefecture_gdf_with_data.plot(
        column=metric,
        ax=ax,
        cmap=PLOT_CONFIG["cmap"],
        legend=True,
        legend_kwds={
            "label": METRICS[metric],
            "shrink": PLOT_CONFIG["colorbar_shrink"],
            "location": PLOT_CONFIG["legend_position"]
        },
        missing_kwds={
            "color": "lightgrey",
            "label": "无数据"
        }
    )

    # 设置颜色条字体
    cbar = ax.get_figure().get_axes()[1]
    cbar.tick_params(labelsize=PLOT_CONFIG["legend_fontsize"])

    # 绘制站点（可选+过滤密集站点）
    if PLOT_CONFIG["station_show"] and not year_station_df.empty:
        filtered_stations = filter_dense_stations(
            year_station_df,
            min_distance=PLOT_CONFIG["station_filter_distance"]
        )
        ax.scatter(
            filtered_stations["经度(°)"],
            filtered_stations["纬度(°)"],
            s=PLOT_CONFIG["station_size"],
            c=PLOT_CONFIG["station_color"],
            alpha=PLOT_CONFIG["station_alpha"],
            label="气象站",
            zorder=5
        )

    # 图表样式
    ax.set_title(
        f"{year}年 中国地级市{METRICS[metric]}热力图",
        fontsize=PLOT_CONFIG["title_fontsize"],
        pad=20
    )
    ax.set_xlabel("经度 (°E)", fontsize=PLOT_CONFIG["label_fontsize"])
    ax.set_ylabel("纬度 (°N)", fontsize=PLOT_CONFIG["label_fontsize"])
    ax.tick_params(axis="both", labelsize=PLOT_CONFIG["label_fontsize"] - 2)
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)  # 隐藏边框

    # 图例
    if ax.get_legend():
        ax.legend(fontsize=PLOT_CONFIG["legend_fontsize"], loc="upper left")

    # 保存图片
    plot_filename = sanitize_filename(f"{year}年_{metric_alias}.png")
    plot_path = os.path.join(year_plot_dir, plot_filename)
    plt.tight_layout()
    plt.savefig(plot_path, bbox_inches="tight")
    plt.close()

    return plot_path


def generate_all_heatmaps(prefecture_df, prefecture_gdf, station_df, plot_dir):
    """生成所有年份+所有指标的热力图"""
    print("\n===== 第四步：生成地级热力图 =====")
    create_directory(plot_dir)

    total_plots = (TARGET_END_YEAR - TARGET_START_YEAR + 1) * len(METRICS)
    with tqdm(total=total_plots, desc="🎨 生成热力图") as pbar:
        for year in range(TARGET_START_YEAR, TARGET_END_YEAR + 1):
            for metric in METRICS.keys():
                try:
                    create_prefecture_heatmap(
                        year, metric, prefecture_gdf, prefecture_df, station_df, plot_dir
                    )
                    pbar.update(1)
                    pbar.set_postfix({"当前": f"{year}年-{METRICS[metric]}"})
                except Exception as e:
                    print(f"\n⚠️ 生成{year}年{METRICS[metric]}热力图失败：{str(e)}")
                    pbar.update(1)

    print(f"✅ 所有热力图已保存至：{os.path.abspath(plot_dir)}")


# ===========================================================================
# 主函数
# ===========================================================================
def main():
    try:
        # 1. 创建结果文件夹
        result_dir = create_directory(RESULT_DIR)
        plot_dir = create_directory(PLOT_DIR)

        # 2. 计算站点级指标
        station_df = compute_station_metrics(ROOT_DIR, TARGET_START_YEAR, TARGET_END_YEAR, result_dir)
        if station_df.empty:
            return

        # 3. 加载地级边界
        prefecture_gdf = load_prefecture_boundaries(SHP_PATH)

        # 4. 生成地级面板数据
        prefecture_df = aggregate_to_prefecture(station_df, prefecture_gdf, result_dir)

        # 5. 生成热力图
        generate_all_heatmaps(prefecture_df, prefecture_gdf, station_df, plot_dir)

        # 输出统计信息
        print("\n===== 📊 最终结果统计 =====")
        print(f"📅 时间范围：{TARGET_START_YEAR}-{TARGET_END_YEAR} 年")
        print(f"📍 涉及站点数：{station_df['气象站编号'].nunique()}")
        print(f"🏙️ 涉及地级市数：{prefecture_df['地级市代码'].nunique()}")
        print(f"📈 有效地级记录数：{len(prefecture_df.dropna(subset=METRICS.keys()))} / {len(prefecture_df)}")
        print(f"\n✅ 所有任务完成！")
        print(f"📁 数据文件：{os.path.abspath(result_dir)}")
        print(f"📁 图表文件：{os.path.abspath(plot_dir)}")

    except Exception as e:
        import traceback
        print(f"\n❌ 执行出错：{str(e)}")
        traceback.print_exc()
        return


if __name__ == "__main__":
    # 首次运行需安装依赖（取消注释）
    # !pip install pandas numpy tqdm geopandas shapely matplotlib
    main()