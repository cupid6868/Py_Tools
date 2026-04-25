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
# 用于直接读取DBF属性表，绕过geopandas字段匹配限制

# 忽略无关警告
warnings.filterwarnings("ignore")

# ===================== 全局配置（请重点修改此处！） =====================
# 数据路径
ROOT_DIR = r"D:\Data_1"  # 站点年度数据根目录
SHP_FOLDER = r"D:\Data_2\2023年省级"  # 省级SHP文件夹（不含文件名）
SHP_FILENAME = "省级.shp"  # SHP文件名（如：china_province.shp）
# 🔴 关键配置：请根据你的DBF文件实际字段名修改！
# 1. 先运行代码，会打印出所有字段名，再回来修改这两个值
PROVINCE_NAME_FIELD = "省"  # DBF中省份名称的字段名（如：NAME/Province/省名）
PROVINCE_CODE_FIELD = "省级码"  # DBF中省份代码的字段名（如：CODE/ADMINCODE/省代码）

# 时间范围
TARGET_START_YEAR = 2010  # 起始年份
TARGET_END_YEAR = 2024  # 结束年份

# 结果保存路径
RESULT_DIR = "01_province_weather_plots_Aver"  # 省级结果文件夹
PLOT_DIR = "01_province_weather_plots_Aver"  # 省级热力图保存文件夹

# 8个核心气象指标
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

# 图表样式配置（适配省级尺度）
PLOT_CONFIG = {
    "figsize": (18, 12),  # 图表尺寸
    "dpi": 300,  # 分辨率
    "cmap": "RdYlBu_r",  # 配色方案
    "station_size": 8,  # 站点标记大小
    "station_color": "#333333",  # 站点颜色
    "station_alpha": 0.4,  # 站点透明度
    "station_show": True,  # 是否显示站点
    "station_filter_distance": 10000,  # 过滤密集站点阈值（米）
    "title_fontsize": 18,  # 标题字体大小
    "label_fontsize": 14,  # 坐标轴标签大小
    "legend_fontsize": 12,  # 图例字体大小
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


def filter_dense_stations(station_df, min_distance=10000):
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
        print(f"📅 正在加载 {year} 年数据...")
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
# 第二步：核心修改 - 手动读取DBF属性表匹配省份名称
# ===========================================================================
def load_province_boundaries_manual(shp_folder, shp_filename, name_field, code_field):
    """
    不依赖pyshp，纯geopandas读取省级边界数据（解决名称读取失败）
    """
    print("\n===== 第二步：手动加载省级边界数据 =====")
    shp_path = os.path.join(shp_folder, shp_filename)

    # 1. 读取SHP完整数据（几何+属性）
    try:
        gdf = gpd.read_file(shp_path)
        if gdf.crs is None:
            gdf = gdf.set_crs("EPSG:4326")  # 强制设置WGS84坐标系
        print(f"✅ 成功读取SHP数据：共{len(gdf)}个省级行政区")

        # 打印所有字段名（核心！用于匹配省份名称）
        print(f"\n🔍 你的SHP文件包含以下字段：")
        for idx, field in enumerate(gdf.columns):
            print(f"   {idx + 1}. {field}")

        # 2. 匹配省份代码字段
        if code_field in gdf.columns:
            prov_code = gdf[code_field].astype(str)
        else:
            print(f"⚠️ 未找到代码字段 '{code_field}'，使用索引替代")
            prov_code = [f"PROV_{i}" for i in range(len(gdf))]

        # 3. 匹配省份名称字段
        if name_field in gdf.columns:
            prov_name = gdf[name_field].astype(str)
        else:
            print(f"⚠️ 未找到名称字段 '{name_field}'，使用默认名称")
            prov_name = [f"省份_{i}" for i in range(len(gdf))]

        # 4. 重构GeoDataFrame（仅保留核心字段）
        province_gdf = gpd.GeoDataFrame({
            "prov_code": prov_code,
            "prov_name": prov_name,
            "geometry": gdf["geometry"]
        }, crs=gdf.crs)

        # 验证匹配结果
        print(f"\n✅ 省份名称匹配结果（前5个）：")
        for _, row in province_gdf.head(5).iterrows():
            print(f"   代码：{row['prov_code']} | 名称：{row['prov_name']}")

        return province_gdf

    except Exception as e:
        print(f"❌ 读取SHP失败：{str(e)}")
        return None

# ===========================================================================
# 第三步：省级算术平均计算（核心逻辑）
# ===========================================================================
def calculate_province_average(station_df, province_gdf):
    """
    按省级行政区划匹配站点，计算省内所有站点的算术平均值
    """
    # 将站点数据转换为GeoDataFrame（用于空间匹配）
    station_geo = gpd.GeoDataFrame(
        station_df,
        geometry=gpd.points_from_xy(station_df["经度(°)"], station_df["纬度(°)"]),
        crs="EPSG:4326"
    )

    # 空间连接：匹配每个站点所属的省份（仅保留省内站点）
    station_with_province = gpd.sjoin(
        station_geo,
        province_gdf[["prov_code", "prov_name", "geometry"]],
        how="left",
        predicate="within"
    )

    # 按年份+省份分组计算算术平均值
    province_results = []
    for year in tqdm(station_df["年份"].unique(), desc="📍 计算省级算术平均"):
        year_stations = station_with_province[station_with_province["年份"] == year]

        # 按省份分组
        province_groups = year_stations.groupby(["prov_code", "prov_name"])
        for (prov_code, prov_name), group in province_groups:
            if len(group) == 0:
                continue

            # 计算各指标算术平均值
            avg_metrics = {}
            for metric in METRICS.keys():
                if group[metric].notnull().any():
                    avg_metrics[metric] = round(group[metric].mean(), 2)
                else:
                    avg_metrics[metric] = np.nan

            # 组装结果
            province_results.append({
                "年份": year,
                "省份代码": prov_code,
                "省份名称": prov_name,
                "省内站点数": len(group),
                **avg_metrics
            })

    # 补充无站点的省份（填充NaN）
    all_years = station_df["年份"].unique()
    all_provinces = province_gdf[["prov_code", "prov_name"]].values
    for year in all_years:
        for prov_code, prov_name in all_provinces:
            exists = any(
                r["年份"] == year and r["省份代码"] == prov_code
                for r in province_results
            )
            if not exists:
                province_results.append({
                    "年份": year,
                    "省份代码": prov_code,
                    "省份名称": prov_name,
                    "省内站点数": 0,
                    **{metric: np.nan for metric in METRICS.keys()}
                })

    return pd.DataFrame(province_results)


def aggregate_to_province(station_df, province_gdf, result_dir):
    """生成省级算术平均面板数据并保存"""
    print("\n===== 第三步：生成省级算术平均面板数据 =====")

    # 计算省级算术平均
    province_df = calculate_province_average(station_df, province_gdf)

    # 排序并保存
    province_df = province_df.sort_values(["年份", "省份代码"]).reset_index(drop=True)
    output_path = os.path.join(result_dir, "02_province_weather_average_2010_2024.csv")
    province_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"✅ 省级算术平均数据已保存：{os.path.abspath(output_path)}")
    return province_df


# ===========================================================================
# 第四步：生成省级热力图
# ===========================================================================
def create_province_heatmap(year, metric, province_gdf, province_df, station_df, plot_dir):
    """生成单年份单指标的省级热力图"""
    # 创建文件夹
    metric_alias = sanitize_filename(METRICS[metric])
    year_plot_dir = os.path.join(plot_dir, f"{year}年", metric_alias)
    create_directory(year_plot_dir)

    # 筛选数据
    year_prov_df = province_df[province_df["年份"] == year].copy()
    year_station_df = station_df[station_df["年份"] == year].copy()

    # 合并边界与指标数据
    province_gdf_with_data = province_gdf.merge(
        year_prov_df[["省份代码", metric, "省内站点数"]],
        left_on="prov_code",
        right_on="省份代码",
        how="left"
    ).dropna(subset=[metric])

    # 创建图表
    fig, ax = plt.subplots(figsize=PLOT_CONFIG["figsize"], dpi=PLOT_CONFIG["dpi"])

    # 绘制热力图
    province_gdf_with_data.plot(
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
        },
        edgecolor='white', linewidth=0.8
    )

    # 设置颜色条字体
    cbar = ax.get_figure().get_axes()[1]
    cbar.tick_params(labelsize=PLOT_CONFIG["legend_fontsize"])

    # 绘制省内站点
    if PLOT_CONFIG["station_show"] and not year_station_df.empty:
        station_geo = gpd.GeoDataFrame(
            year_station_df,
            geometry=gpd.points_from_xy(year_station_df["经度(°)"], year_station_df["纬度(°)"]),
            crs="EPSG:4326"
        )
        # 仅显示省份内的站点
        station_in_province = gpd.sjoin(station_geo, province_gdf, how="inner", predicate="within")
        if not station_in_province.empty:
            filtered_stations = filter_dense_stations(
                station_in_province,
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
        f"{year}年 中国省级{METRICS[metric]}热力图（算术平均）",
        fontsize=PLOT_CONFIG["title_fontsize"],
        pad=20
    )
    ax.set_xlabel("经度 (°E)", fontsize=PLOT_CONFIG["label_fontsize"])
    ax.set_ylabel("纬度 (°N)", fontsize=PLOT_CONFIG["label_fontsize"])
    ax.tick_params(axis="both", labelsize=PLOT_CONFIG["label_fontsize"] - 2)
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)

    # 图例
    if ax.get_legend():
        ax.legend(fontsize=PLOT_CONFIG["legend_fontsize"], loc="upper left")

    # 保存图片
    plot_filename = sanitize_filename(f"{year}年_{metric_alias}_算术平均.png")
    plot_path = os.path.join(year_plot_dir, plot_filename)
    plt.tight_layout()
    plt.savefig(plot_path, bbox_inches="tight")
    plt.close()

    return plot_path


def generate_all_heatmaps(province_df, province_gdf, station_df, plot_dir):
    """生成所有年份+所有指标的省级热力图"""
    print("\n===== 第四步：生成省级热力图 =====")
    create_directory(plot_dir)

    total_plots = (TARGET_END_YEAR - TARGET_START_YEAR + 1) * len(METRICS)
    with tqdm(total=total_plots, desc="🎨 生成热力图") as pbar:
        for year in range(TARGET_START_YEAR, TARGET_END_YEAR + 1):
            for metric in METRICS.keys():
                try:
                    create_province_heatmap(
                        year, metric, province_gdf, province_df, station_df, plot_dir
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

        # 3. 手动加载省级边界（核心解决名称读取问题）
        province_gdf = load_province_boundaries_manual(
            SHP_FOLDER,
            SHP_FILENAME,
            PROVINCE_NAME_FIELD,
            PROVINCE_CODE_FIELD
        )
        if province_gdf is None:
            print("❌ 省级边界数据加载失败，程序退出")
            return

        # 4. 生成省级算术平均面板数据
        province_df = aggregate_to_province(station_df, province_gdf, result_dir)

        # 5. 生成热力图
        generate_all_heatmaps(province_df, province_gdf, station_df, plot_dir)

        # 输出统计信息
        print("\n===== 📊 最终结果统计 =====")
        print(f"📅 时间范围：{TARGET_START_YEAR}-{TARGET_END_YEAR} 年")
        print(f"📍 涉及站点数：{station_df['气象站编号'].nunique()}")
        print(f"🏙️ 涉及省份数：{province_df['省份代码'].nunique()}")
        print(f"📈 有效省级记录数：{len(province_df.dropna(subset=METRICS.keys()))} / {len(province_df)}")

        # 统计各省份站点数
        station_count_by_prov = province_df.groupby("省份名称")["省内站点数"].max().sort_values(ascending=False)
        print(f"\n🔍 各省份最大站点数（前10）：")
        for prov, count in station_count_by_prov.head(10).items():
            print(f"   {prov}: {count}个站点")

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
    # !pip install pandas numpy tqdm geopandas shapely matplotlib pyshp
    main()