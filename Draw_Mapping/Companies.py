import os
import glob
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from tqdm import tqdm

# ===================== 全局配置参数（根据实际路径修改） =====================
# 气象数据根目录（包含年份文件夹）
ROOT_DIR = r"D:\Data_1"
# 企业经纬度文件路径（你的Excel）
COMPANY_PATH = r"C:\Users\18218\Desktop\PythonProject\2008-2024年企业经纬度数据.csv"
# 结果输出根文件夹
RESULT_DIR = "02_companies_result"
# 图片输出文件夹
PLOT_DIR = "02_companies_result"

COM_LON = "RegisterLongitude"
COM_LAT = "RegisterLatitude"
COM_YERA = "year"

# 气候基准期 & 目标年份
BASELINE_YEARS = (1981, 2010)
TARGET_YEARS = (2008, 2024)

# 极端天气指标
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

# 图表配置
PLOT_CONFIG = {
    "figsize": (14, 10),
    "dpi": 300,
    "station_color": "#FF3333",  # 气象站：红色
    "company_color": "#3366FF",  # 企业：蓝色
    "line_color": "#999999",  # 连线：灰色
    "station_size": 10,
    "company_size": 8,
    "line_alpha": 0.4,
    "title_fontsize": 16,
    "label_fontsize": 12
}

# 输出文件配置
FILES_CONFIG = {
    "baseline_sorted": "01_baseline_data_sorted.csv",
    "thresholds": "02_weather_thresholds.csv",
    "extreme_days": "03_station_extreme_weather_stats.csv",
}


# ===================== 修复中文显示问题 =====================
def setup_chinese_font():
    matplotlib.rcParams['font.sans-serif'] = [
        'Microsoft YaHei', 'SimHei', 'PingFang SC', 'Hiragino Sans GB', 'DejaVu Sans'
    ]
    matplotlib.rcParams['axes.unicode_minus'] = False
    matplotlib.rcParams['font.family'] = 'sans-serif'
    print("✅ 中文字体配置完成")


setup_chinese_font()


# ===========================================================================
# 基础工具函数
# ===========================================================================
def create_directory(dir_name):
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
        print(f"✅ 已创建文件夹：{os.path.abspath(dir_name)}")
    return dir_name


def sanitize_filename(filename):
    illegal_chars = r'[<>:\"/\\|?*]'
    return re.sub('_+', '_', re.sub(illegal_chars, '_', filename)).strip('_')


def haversine_distance(lat1, lon1, lat2, lon2):
    """计算两点间哈弗辛距离（米）"""
    lat1_rad, lon1_rad = np.radians(lat1), np.radians(lon1)
    lat2_rad, lon2_rad = np.radians(lat2), np.radians(lon2)
    dlat, dlon = lat2_rad - lat1_rad, lon2_rad - lon1_rad
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1_rad) * np.cos(lat2_rad) * np.sin(dlon / 2) ** 2
    return 6371000 * 2 * np.arcsin(np.sqrt(a))


# ===========================================================================
# 气象数据处理（完整保留站点指标计算）
# ===========================================================================
def load_all_data(root_dir, start_year, end_year):
    all_dfs = []
    for year in tqdm(range(start_year, end_year + 1), desc="加载气象数据"):
        year_dir = os.path.join(root_dir, str(year))
        if not os.path.exists(year_dir):
            print(f"⚠️ 跳过：{year} 年文件夹不存在")
            continue
        csv_files = glob.glob(os.path.join(year_dir, "*.csv"))
        for file in csv_files:
            try:
                df = pd.read_csv(
                    file, parse_dates=["DATE"],
                    usecols=["STATION", "NAME", "LATITUDE", "LONGITUDE", "DATE", "MIN", "MAX", "PRCP"]
                )
                df = df.dropna(subset=["STATION", "NAME", "LATITUDE", "LONGITUDE", "DATE", "MIN", "MAX", "PRCP"])
                missing = [9999.9, 99.99, -9999.9, -99.99]
                df["MIN"] = df["MIN"].replace(missing, np.nan)
                df["MAX"] = df["MAX"].replace(missing, np.nan)
                df["PRCP"] = df["PRCP"].replace(missing, np.nan)
                df = df.dropna(subset=["MIN", "MAX", "PRCP"])
                df = df[df["PRCP"] >= 0]
                df["MIN_C"] = (df["MIN"] - 32) * 5 / 9
                df["MAX_C"] = (df["MAX"] - 32) * 5 / 9
                df["PRCP_mm"] = df["PRCP"] * 25.4
                df["year"] = df["DATE"].dt.year
                df["month_day"] = df["DATE"].dt.strftime("%m-%d")
                df["LATITUDE"] = df["LATITUDE"].round(6)
                df["LONGITUDE"] = df["LONGITUDE"].round(6)
                all_dfs.append(df)
            except Exception as e:
                print(f"❌ 读取失败 {file}: {e}")
    if not all_dfs:
        raise ValueError("未加载到有效气象数据")
    return pd.concat(all_dfs, ignore_index=True).drop_duplicates()


def sort_baseline_data(baseline_df, result_dir):
    baseline_sorted = baseline_df.sort_values(["STATION", "DATE"]).reset_index(drop=True)
    path = os.path.join(result_dir, FILES_CONFIG["baseline_sorted"])
    baseline_sorted.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"✅ 基准期数据已保存")
    return baseline_sorted


def calculate_relative_thresholds(baseline_df, result_dir):
    coords = baseline_df.groupby("STATION")[["NAME", "LATITUDE", "LONGITUDE"]].first().reset_index()
    thresh = baseline_df.groupby(["STATION", "month_day"]).agg(
        extreme_low_temp=("MIN_C", lambda x: np.percentile(x, 10)),
        extreme_high_temp=("MAX_C", lambda x: np.percentile(x, 90)),
        extreme_rain=("PRCP_mm", lambda x: np.percentile(x, 95)),
        extreme_drought=("PRCP_mm", lambda x: np.percentile(x, 5))
    ).reset_index()
    thresh = thresh.merge(coords, on="STATION")
    path = os.path.join(result_dir, FILES_CONFIG["thresholds"])
    thresh.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"✅ 极端阈值已计算完成")
    return thresh


def count_extreme_days_yearly(target_df, threshold_df, result_dir):
    merged = target_df.merge(threshold_df, on=["STATION", "month_day"])
    merged["is_extreme_low"] = merged["MIN_C"] < merged["extreme_low_temp"]
    merged["is_extreme_high"] = merged["MAX_C"] > merged["extreme_high_temp"]
    merged["is_extreme_rain"] = merged["PRCP_mm"] > merged["extreme_rain"]
    merged["is_extreme_drought"] = merged["PRCP_mm"] < merged["extreme_drought"]
    stats = merged.groupby(["STATION", "year"]).agg(
        extreme_low_days=("is_extreme_low", "sum"),
        extreme_high_days=("is_extreme_high", "sum"),
        extreme_rain_days=("is_extreme_rain", "sum"),
        extreme_drought_days=("is_extreme_drought", "sum"),
        total_days=("DATE", "count")
    ).reset_index()
    coords = target_df.groupby("STATION")[["NAME", "LATITUDE", "LONGITUDE"]].first().reset_index()
    stats = stats.merge(coords, on="STATION")
    for m in EXTREME_METRICS:
        stats[m] = stats[m].astype(int)
    path = os.path.join(result_dir, FILES_CONFIG["extreme_days"])
    stats.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"✅ 站点年度极端天气指标已完成")
    return stats


# ===========================================================================
# 核心：企业 ↔ 最近气象站 匹配（严格使用【注册地经度、注册地纬度】）
# ===========================================================================
def load_company_data(path):
    """加载企业数据，严格使用【注册地经度、注册地纬度】"""
    df = pd.read_csv(path, encoding='utf-8-sig')
    print("📋 CSV文件里的真实列名：")
    print(df.columns.tolist())

    # 固定使用你提供的列名，不自动识别！
    df["company_lon"] = df[COM_LON]
    df["company_lat"] = df[COM_LAT]
    df["company_year"] = df[COM_YERA]

    # 清洗经纬度
    df["company_lon"] = pd.to_numeric(df["company_lon"], errors="coerce")
    df["company_lat"] = pd.to_numeric(df["company_lat"], errors="coerce")
    df = df.dropna(subset=["company_lon", "company_lat"])

    df["company_year"] = pd.to_numeric(df["company_year"], errors="coerce")
    df = df.dropna(subset=["company_lon", "company_lat", "company_year"])

    print(f"✅ 加载企业数据：共 {len(df)} 家有效企业（已使用注册地经纬度）")
    return df

def match_nearest_station(company_df, station_df):
    """为每家企业匹配【最近气象站】"""
    matched_list = []
    station_points = station_df[["STATION", "NAME", "LATITUDE", "LONGITUDE"]].drop_duplicates()

    for _, comp in tqdm(company_df.iterrows(), total=len(company_df), desc="匹配最近气象站"):
        clat, clon = comp["company_lat"], comp["company_lon"]
        # 计算到所有气象站距离
        distances = []
        for _, st in station_points.iterrows():
            slat, slon = st["LATITUDE"], st["LONGITUDE"]
            dist = haversine_distance(clat, clon, slat, slon)
            distances.append({
                "STATION": st["STATION"],
                "station_name": st["NAME"],
                "station_lat": slat,
                "station_lon": slon,
                "distance_m": round(dist, 2)
            })
        # 取最近
        nearest = sorted(distances, key=lambda x: x["distance_m"])[0]
        # 合并企业 + 气象站信息
        res = comp.to_dict()
        res.update(nearest)
        matched_list.append(res)

    return pd.DataFrame(matched_list)


def merge_weather_to_company(matched_df, yearly_station_df, year):
    """将年度气象指标合并到企业数据"""
    merged = matched_df.merge(
        yearly_station_df[["STATION"] + EXTREME_METRICS],
        on="STATION", how="left"
    )
    merged["match_year"] = year
    return merged


# ===========================================================================
# 可视化：气象站 + 企业 + 连线
# ===========================================================================
def plot_matching_map(matched_df, year, save_dir):
    plt.figure(figsize=PLOT_CONFIG["figsize"], dpi=PLOT_CONFIG["dpi"])
    # 气象站
    stations = matched_df[["station_lon", "station_lat"]].drop_duplicates()
    plt.scatter(
        stations["station_lon"], stations["station_lat"],
        c=PLOT_CONFIG["station_color"], s=PLOT_CONFIG["station_size"],
        alpha=0.8, label="气象站点"
    )
    # 企业（注册地位置）
    plt.scatter(
        matched_df["company_lon"], matched_df["company_lat"],
        c=PLOT_CONFIG["company_color"], s=PLOT_CONFIG["company_size"],
        alpha=0.8, label="上市企业(注册地)"
    )
    # 连线
    for _, row in matched_df.iterrows():
        plt.plot(
            [row["company_lon"], row["station_lon"]],
            [row["company_lat"], row["station_lat"]],
            c=PLOT_CONFIG["line_color"], alpha=PLOT_CONFIG["line_alpha"], linewidth=0.6
        )
    plt.title(f"{year}年 上市企业(注册地) ↔ 最近气象站 匹配分布图", fontsize=PLOT_CONFIG["title_fontsize"])
    plt.xlabel("经度 (°E)", fontsize=PLOT_CONFIG["label_fontsize"])
    plt.ylabel("纬度 (°N)", fontsize=PLOT_CONFIG["label_fontsize"])
    plt.legend(loc="best")
    plt.tight_layout()
    # 保存
    filename = f"{year}年_企业(注册地)气象站匹配分布图.png"
    plt.savefig(os.path.join(save_dir, filename), bbox_inches="tight")
    plt.close()


# ===========================================================================
# 主函数
# ===========================================================================
def main():
    try:
        # 1. 创建目录
        res_dir = create_directory(RESULT_DIR)
        plot_dir = create_directory(PLOT_DIR)
        csv_dir = create_directory(os.path.join(res_dir, "年度匹配结果"))

        # 2. 气象数据处理（站点指标完整保留）
        print("\n===== 步骤1：处理基准期气象数据 =====")
        baseline = load_all_data(ROOT_DIR, *BASELINE_YEARS)
        sort_baseline_data(baseline, res_dir)
        thresh = calculate_relative_thresholds(baseline, res_dir)

        print("\n===== 步骤2：计算站点年度极端天气指标 =====")
        target = load_all_data(ROOT_DIR, *TARGET_YEARS)
        station_yearly = count_extreme_days_yearly(target, thresh, res_dir)

        # 3. 加载企业并匹配最近气象站（注册地经纬度）
        print("\n===== 步骤3：加载企业并匹配最近气象站 =====")
        company_df = load_company_data(COMPANY_PATH)
        matched_base = match_nearest_station(company_df, station_yearly)

        # 4. 按年份输出结果 + 画图
        print("\n===== 步骤4：年度匹配结果输出与可视化 =====")
        for year in range(TARGET_YEARS[0], TARGET_YEARS[1] + 1):
            print(f"\n📅 处理 {year} 年...")
            yearly_data = station_yearly[station_yearly["year"] == year].copy()
            if yearly_data.empty:
                print(f"⚠️ {year}年无站点数据，跳过")
                continue
            # ==============================================
            # 【核心修改】只筛选当前年份的企业进行匹配
            # ==============================================
            year_company_df = company_df[company_df["company_year"] == year].copy()
            if year_company_df.empty:
                print(f"⚠️ {year}年无对应企业数据，跳过")
                continue

            # 为【当年企业】匹配最近气象站
            matched_base = match_nearest_station(year_company_df, station_yearly)

            # 合并气象指标
            final = merge_weather_to_company(matched_base, yearly_data, year)
            # 保存CSV
            csv_path = os.path.join(csv_dir, f"{year}年_企业极端天气匹配结果.csv")
            final.to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"✅ 已保存：{csv_path}")
            # 绘图
            plot_matching_map(final, year, plot_dir)

        # 完成
        print("\n===== ✅ 全部任务完成！=====")
        print(f"📁 站点气象指标：{os.path.abspath(res_dir)}")
        print(f"📁 企业年度匹配CSV：{os.path.abspath(csv_dir)}")
        print(f"📁 匹配分布图：{os.path.abspath(plot_dir)}")
        print(f"✅ 匹配依据：注册地经度、注册地纬度")

    except Exception as e:
        import traceback
        print(f"\n❌ 运行出错：{e}")
        traceback.print_exc()


if __name__ == "__main__":
    # 首次运行安装依赖
    # !pip install pandas numpy tqdm matplotlib openpyxl
    main()