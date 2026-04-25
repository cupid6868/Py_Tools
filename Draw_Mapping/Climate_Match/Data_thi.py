import os
import glob
import pandas as pd
import numpy as np
from tqdm import tqdm

# ===================== 配置参数 =====================
ROOT_DIR = r"D:\Data_1"  # 数据根目录（包含年份文件夹）
REFERENCE_PERIOD = (1990, 2005)  # 参考期：1990-2005年
TARGET_PERIOD = (2010, 2024)  # 目标期：2010-2024年
RESULT_DIR = "weather_analysis_third_results"  # 结果文件夹（与之前保持一致）
# 定义5个分项指标（P90/P10为二值化指标，R/D/W沿用原定义）
INDICATORS = {
    "P90": "MAX_C",  # 极端高温（日最高温）
    "P10": "MIN_C",  # 极端低温（日最低温）
    "R": "PRCP_mm",  # 极端降雨（日降水量）
    "D": "PRCP_mm",  # 极端干旱（日降水量）
    "W": "MAX_C"  # 极端大风（示例：暂用最高温替代，需替换为实际风速列）
}


# ===========================================================================

def create_result_dir(dir_name):
    """创建结果文件夹（复用原有函数）"""
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
        print(f"已创建结果文件夹：{os.path.abspath(dir_name)}")
    else:
        print(f"结果文件夹已存在：{os.path.abspath(dir_name)}")
    return dir_name


def load_daily_data(root_dir, start_year, end_year):
    """加载指定时段所有站点日度数据（保留原始日度信息，用于分日判断）"""
    all_dfs = []
    for year in tqdm(range(start_year, end_year + 1), desc=f"加载{start_year}-{end_year}年数据"):
        year_dir = os.path.join(root_dir, str(year))
        if not os.path.exists(year_dir):
            print(f"警告：年份文件夹 {year_dir} 不存在，跳过")
            continue

        csv_files = glob.glob(os.path.join(year_dir, "*.csv"))
        if not csv_files:
            print(f"警告：{year} 年无CSV文件，跳过")
            continue

        for file in csv_files:
            try:
                # 读取核心列（保留日度信息）
                df = pd.read_csv(
                    file,
                    parse_dates=["DATE"],
                    usecols=["STATION", "NAME", "DATE", "MIN", "MAX", "PRCP"]
                )
                # 缺失值和异常值处理
                df = df.dropna(subset=["STATION", "DATE", "MIN", "MAX", "PRCP"])
                missing_marks = [9999.9, 99.99, -9999.9, -99.99]
                df["MIN"] = df["MIN"].replace(missing_marks, np.nan)
                df["MAX"] = df["MAX"].replace(missing_marks, np.nan)
                df["PRCP"] = df["PRCP"].replace(missing_marks, np.nan)
                df = df.dropna(subset=["MIN", "MAX", "PRCP"])

                # 单位转换：华氏度→摄氏度，英寸→毫米
                df["MIN_C"] = (df["MIN"] - 32) * 5 / 9
                df["MAX_C"] = (df["MAX"] - 32) * 5 / 9
                df["PRCP_mm"] = df["PRCP"] * 25.4

                # 提取时间维度（年/月/日，核心：保留日度信息）
                df["year"] = df["DATE"].dt.year
                df["month"] = df["DATE"].dt.month
                df["day"] = df["DATE"].dt.day
                df["year_month"] = df["DATE"].dt.to_period("M")  # 年月标识

                all_dfs.append(df)
            except Exception as e:
                print(f"读取文件 {file} 失败：{e}，跳过")

    if not all_dfs:
        raise ValueError("未加载到有效数据，请检查路径/文件格式")

    combined_df = pd.concat(all_dfs, ignore_index=True)
    return combined_df


def calculate_reference_quantiles(df, reference_period):
    """计算参考期（1990-2005）各站点-各月份的气温分位数（P90/P10）"""
    # 筛选参考期数据
    ref_df = df[
        (df["year"] >= reference_period[0]) &
        (df["year"] <= reference_period[1])
        ].copy()

    # 按站点+月份分组，计算P90（最高温90分位）和P10（最低温10分位）
    ref_quantiles = {}
    for (station, month), group in ref_df.groupby(["STATION", "month"]):
        quantiles = {
            "P90_quantile": group["MAX_C"].quantile(0.9),  # 参考期同月最高温90分位
            "P10_quantile": group["MIN_C"].quantile(0.1)  # 参考期同月最低温10分位
        }
        ref_quantiles[(station, month)] = quantiles

    print(f"参考期分位数计算完成：共{len(ref_quantiles)}个站点-月份组合")
    return ref_quantiles


def calculate_p90_p10_daily(df, ref_quantiles):
    """
    计算日度P90/P10二值化指标：
    - P90：当日最高温 > 参考期同月90分位 → 1，否则0
    - P10：当日最低温 < 参考期同月10分位 → 1，否则0
    """
    daily_indicators = []

    for _, row in tqdm(df.iterrows(), desc="计算日度P90/P10指标", total=len(df)):
        station = row["STATION"]
        month = row["month"]
        quantile_key = (station, month)

        # 跳过无参考数据的站点-月份
        if quantile_key not in ref_quantiles:
            continue

        ind_row = {
            "STATION": station,
            "NAME": row["NAME"],
            "year": row["year"],
            "month": row["month"],
            "day": row["day"],
            "year_month": row["year_month"],
            "MAX_C": row["MAX_C"],
            "MIN_C": row["MIN_C"],
            "PRCP_mm": row["PRCP_mm"]
        }

        # 计算日度P90（极端高温）二值化指标
        p90_quantile = ref_quantiles[quantile_key]["P90_quantile"]
        ind_row["P90_daily"] = 1 if row["MAX_C"] > p90_quantile else 0

        # 计算日度P10（极端低温）二值化指标
        p10_quantile = ref_quantiles[quantile_key]["P10_quantile"]
        ind_row["P10_daily"] = 1 if row["MIN_C"] < p10_quantile else 0

        daily_indicators.append(ind_row)

    daily_df = pd.DataFrame(daily_indicators)

    # 按月加总构建月度P90/P10指标（核心：日度二值化结果求和）
    monthly_df = daily_df.groupby(["STATION", "NAME", "year", "month"]).agg({
        "P90_daily": "sum",  # 月度极端高温天数（P90_i,m）
        "P10_daily": "sum",  # 月度极端低温天数（P10_i,m）
        "MAX_C": "mean",  # 月度平均最高温（用于其他指标）
        "MIN_C": "mean",  # 月度平均最低温（用于其他指标）
        "PRCP_mm": "mean"  # 月度平均降水量（用于R/D指标）
    }).reset_index()

    # 重命名为统一指标名
    monthly_df.rename(columns={
        "P90_daily": "P90",
        "P10_daily": "P10"
    }, inplace=True)

    return daily_df, monthly_df


def calculate_reference_stats(monthly_df, reference_period, indicators):
    """计算参考期（1990-2005）各站点-各月份-各指标的均值和标准差（用于标准化）"""
    # 筛选参考期月度数据
    ref_df = monthly_df[
        (monthly_df["year"] >= reference_period[0]) &
        (monthly_df["year"] <= reference_period[1])
        ].copy()

    # 按站点+月份分组，计算每个指标的均值和标准差
    ref_stats = {}
    for (station, month), group in ref_df.groupby(["STATION", "month"]):
        stats = {}
        for ind_name, _ in indicators.items():
            if ind_name in group.columns:
                stats[f"{ind_name}_mean"] = group[ind_name].mean()
                stats[f"{ind_name}_std"] = group[ind_name].std() or 1e-6  # 避免除零
        ref_stats[(station, month)] = stats

    print(f"参考期均值/标准差计算完成：共{len(ref_stats)}个站点-月份组合")
    return ref_stats


def standardize_indicators(target_df, ref_stats, indicators):
    """标准化处理：(目标值 - 参考期同月均值) / 参考期同月标准差"""
    standardized_data = []

    for _, row in tqdm(target_df.iterrows(), desc="指标标准化", total=len(target_df)):
        station = row["STATION"]
        month = row["month"]
        stats_key = (station, month)

        # 跳过无参考数据的站点-月份
        if stats_key not in ref_stats:
            continue

        std_row = {
            "STATION": station,
            "NAME": row["NAME"],
            "year": row["year"],
            "month": row["month"],
            "year_month": f"{row['year']}-{row['month']:02d}"
        }

        # 逐个指标标准化（包含P90/P10二值化加总后的指标）
        for ind_name in indicators.keys():
            if ind_name not in row:
                continue
            mean = ref_stats[stats_key][f"{ind_name}_mean"]
            std = ref_stats[stats_key][f"{ind_name}_std"]
            raw_value = row[ind_name]
            std_value = (raw_value - mean) / std  # 核心标准化公式
            std_row[f"{ind_name}_std"] = std_value

        standardized_data.append(std_row)

    return pd.DataFrame(standardized_data)


def calculate_ewri(std_df):
    """计算极端天气风险指数（EWRI）：Σ(P90_std - P10_std + R_std + D_std + W_std)"""
    # 按年月分组计算全国总指数
    ewri_df = std_df.groupby("year_month").apply(
        lambda x: pd.Series({
            "year": x["year"].iloc[0],
            "month": x["month"].iloc[0],
            # 核心公式：P90_std - P10_std + R_std + D_std + W_std
            "EWRI": (
                    x["P90_std"].sum() -
                    x["P10_std"].sum() +
                    x["R_std"].sum() +
                    x["D_std"].sum() +
                    x["W_std"].sum()
            )
        })
    ).reset_index(drop=True)

    # 按年月排序
    ewri_df = ewri_df.sort_values(["year", "month"]).reset_index(drop=True)
    return ewri_df


def main():
    try:
        # 1. 创建结果文件夹
        result_dir = create_result_dir(RESULT_DIR)

        # 2. 加载全量日度数据（参考期+目标期）
        full_daily_df = load_daily_data(ROOT_DIR,
                                        min(REFERENCE_PERIOD[0], TARGET_PERIOD[0]),
                                        max(REFERENCE_PERIOD[1], TARGET_PERIOD[1]))

        # 3. 计算参考期分位数（P90=90分位，P10=10分位）
        print("\n===== 计算1990-2005年参考期气温分位数 =====")
        ref_quantiles = calculate_reference_quantiles(full_daily_df, REFERENCE_PERIOD)

        # 4. 计算日度P90/P10二值化指标 + 月度加总
        print("\n===== 计算P90/P10二值化指标（日度判断+月度加总） =====")
        daily_ind_df, monthly_ind_df = calculate_p90_p10_daily(full_daily_df, ref_quantiles)

        # 5. 筛选目标期（2010-2024）月度数据
        target_monthly_df = monthly_ind_df[
            (monthly_ind_df["year"] >= TARGET_PERIOD[0]) &
            (monthly_ind_df["year"] <= TARGET_PERIOD[1])
            ].copy()

        # 6. 计算参考期各指标均值/标准差（用于标准化）
        print("\n===== 计算参考期各指标均值和标准差 =====")
        ref_stats = calculate_reference_stats(monthly_ind_df, REFERENCE_PERIOD, INDICATORS)

        # 7. 指标标准化（包含P90/P10月度加总指标）
        print("\n===== 标准化2010-2024年各指标 =====")
        std_df = standardize_indicators(target_monthly_df, ref_stats, INDICATORS)

        # 8. 计算EWRI指数
        print("\n===== 计算极端天气风险指数（EWRI） =====")
        ewri_df = calculate_ewri(std_df)

        # 9. 保存结果（带序号的文件名）
        # 04：日度P90/P10二值化数据
        daily_output_path = os.path.join(result_dir, "04_daily_p90_p10.csv")
        # 05：月度P90/P10加总数据
        monthly_output_path = os.path.join(result_dir, "05_monthly_indicators.csv")
        # 06：标准化指标数据
        std_output_path = os.path.join(result_dir, "06_standardized_indicators.csv")
        # 07：EWRI指数数据
        ewri_output_path = os.path.join(result_dir, "07_ewri_index.csv")

        daily_ind_df.to_csv(daily_output_path, index=False, encoding="utf-8-sig")
        monthly_ind_df.to_csv(monthly_output_path, index=False, encoding="utf-8-sig")
        std_df.to_csv(std_output_path, index=False, encoding="utf-8-sig")
        ewri_df.to_csv(ewri_output_path, index=False, encoding="utf-8-sig")

        # 输出完成信息
        print(f"\n===== EWRI指数计算完成！=====")
        print(f"日度P90/P10指标：{daily_output_path}")
        print(f"月度加总指标：{monthly_output_path}")
        print(f"标准化指标：{std_output_path}")
        print(f"EWRI指数：{ewri_output_path}")
        print(f"\n月度P90/P10指标示例（前5行）：")
        print(monthly_ind_df[["STATION", "NAME", "year", "month", "P90", "P10"]].head())
        print(f"\nEWRI指数示例（前5行）：")
        print(ewri_df.head())

    except Exception as e:
        print(f"执行出错：{e}")
        return


if __name__ == "__main__":
    # 安装依赖（首次运行取消注释）
    # !pip install pandas numpy tqdm
    main()