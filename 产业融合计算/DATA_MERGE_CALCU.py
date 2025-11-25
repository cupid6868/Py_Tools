import pandas as pd
import numpy as np
import re
import os
from tqdm import tqdm
import logging
from typing import Hashable, Sequence, Any, Optional
import datetime
import multiprocessing
import time

# ===================== 日志初始化 =====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# ===================== 核心配置 =====================
# 🚨 修改为：专利CSV文件所在目录（批量处理该目录下所有.csv文件）
PATENT_CSV_DIR = r"E:\1.专利数据"  # 所有待处理的CSV文件放在这个目录下

# 其他配置保持不变
IPC_INDUSTRY_DTA = r"D:\BaiduNetdiskDownload\名师讲堂｜使用 Stata 测算数实融合水平（二）： 基于真实专利数据\使用 Stata 测算数实融合水平（二）：基于真实专利数据\国民经济分类与IPC分类号对照表.dta"
DIGITAL_INDUSTRY_DTA = r"D:\BaiduNetdiskDownload\名师讲堂｜使用 Stata 测算数实融合水平（二）： 基于真实专利数据\使用 Stata 测算数实融合水平（二）：基于真实专利数据\数字经济及其核心产业统计分类.dta"

# 输出目录优化：为每个CSV文件创建独立子目录，避免文件冲突
ROOT_OUTPUT_DIR = r"./数实融合测算结果_市级_批量处理"
os.makedirs(ROOT_OUTPUT_DIR, exist_ok=True)

CHUNK_SIZE = 100000
ENCODING = 'utf-8'
NUM_CORES = multiprocessing.cpu_count()
logger.info(f"系统检测到 {NUM_CORES} 个核心，将用于并行加速。")


# ===================== 工具函数（保持不变） =====================
def post_stata_read_clean(df: pd.DataFrame) -> pd.DataFrame:
    df_clean = df.copy()
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':
            df_clean[col] = df_clean[col].astype(str).str.strip()
            df_clean[col] = df_clean[col].replace('nan', '')
    return df_clean


def process_csv_chunk(chunk_df: pd.DataFrame) -> pd.DataFrame:
    chunk_df = post_stata_read_clean(chunk_df)
    processed_chunk = chunk_df[
        chunk_df['IPC'].astype(str).str.contains(';', na=False) &
        (chunk_df['市'] != '')
        ].copy()
    processed_chunk.rename(columns={'市': 'city'}, inplace=True)
    processed_chunk['IPC'] = processed_chunk['IPC'].astype(str).str.replace(' ', '', regex=False)
    processed_chunk['ipc'] = processed_chunk['IPC'].apply(lambda x: ';'.join(x.split(';')[:10]) if x else '')
    return processed_chunk[['ipzlid', 'city', 'ipc']]


def progress_bar(iterable, desc: str) -> tqdm:
    return tqdm(iterable, desc=desc, unit="item", ncols=80)


def unify_column_type(df: pd.DataFrame, col_name: str, target_type: type = str) -> pd.DataFrame:
    if col_name not in df.columns:
        return df
    if target_type == str:
        df[col_name] = df[col_name].astype(str).str.strip().str[:5]
        df[col_name] = df[col_name].replace('nan', '')
    else:
        df[col_name] = pd.to_numeric(df[col_name], errors='coerce')
    return df


def process_ipc_string(ipc_str):
    return [ip.strip() for ip in str(ipc_str).split(';') if ip.strip()]


def process_city_data(city_group):
    city, city_data = city_group
    city_csv_lines = []
    for _, row in city_data.iterrows():
        ipc_list = [i.strip() for i in str(row['ipc']).split(';') if i.strip()]
        ipc_list = list(set(ipc_list))
        n_classes = len(ipc_list)
        value = row['_freq']
        if n_classes >= 2:
            if n_classes == 2:
                city_csv_lines.append(f"{city},{ipc_list[0]},{ipc_list[1]},{value}\n")
                city_csv_lines.append(f"{city},{ipc_list[1]},{ipc_list[0]},{value}\n")
            else:
                for i_idx in range(n_classes):
                    for j_idx in range(n_classes):
                        city_csv_lines.append(f"{city},{ipc_list[i_idx]},{ipc_list[j_idx]},{value}\n")
    return "".join(city_csv_lines)


# ===================== 单个CSV文件处理函数（核心封装） =====================
def process_single_patent_csv(patent_csv_path, output_dir):
    """
    处理单个专利CSV文件的完整流程（从预处理到数实融合指数计算）
    :param patent_csv_path: 单个专利CSV文件路径
    :param output_dir: 该文件的专属输出目录
    """
    logger.info(f"\n{'=' * 80}")
    logger.info(f"开始处理文件：{os.path.basename(patent_csv_path)}")
    logger.info(f"文件路径：{patent_csv_path}")
    logger.info(f"输出目录：{output_dir}")
    logger.info(f"{'=' * 80}")

    # 创建该文件的专属输出子目录
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(os.path.join(output_dir, "对照表"), exist_ok=True)
    os.makedirs(os.path.join(output_dir, "行业匹配结果"), exist_ok=True)

    file_start_time = time.time()
    success = False

    try:
        # 步骤1：专利IPC数据预处理
        logger.info("\n=== 步骤1：专利IPC数据预处理（CSV分块读取优化） ====")
        start_time_step1 = time.time()
        ipcdata_path = os.path.join(output_dir, "ipcdata.csv")
        if not os.path.exists(patent_csv_path) or not patent_csv_path.lower().endswith('.csv'):
            raise FileNotFoundError(f"文件未找到或格式错误！")

        logger.info(f"🚨 使用 pd.read_csv 进行分块读取。CHUNK_SIZE: {CHUNK_SIZE}")
        all_processed_chunks = []
        total_rows_read = 0
        REQUIRED_READ_COLS = ['ipzlid', 'IPC', '市']
        iterator = pd.read_csv(
            patent_csv_path,
            chunksize=CHUNK_SIZE,
            usecols=REQUIRED_READ_COLS,
            encoding=ENCODING,
            low_memory=True
        )
        for chunk in iterator:
            processed_chunk = process_csv_chunk(chunk)
            all_processed_chunks.append(processed_chunk)
            total_rows_read += len(chunk)
            logger.info(f"  > 已读取并处理 {total_rows_read} 行...")
        patent_2012 = pd.concat(all_processed_chunks, ignore_index=True)
        patent_2012.to_csv(ipcdata_path, index=False, encoding=ENCODING)
        processed_patent_data = patent_2012
        logger.info(f"预处理完成！最终数据形状：{patent_2012.shape}, 共处理 {total_rows_read} 行。")
        end_time_step1 = time.time()
        logger.info(f"✅ 步骤1耗时：{end_time_step1 - start_time_step1:.2f} 秒")

        # 步骤2：提取所有唯一IPC分类
        logger.info("\n=== 步骤2：提取所有唯一IPC分类 (多进程加速) ====")
        start_time_step2 = time.time()
        uniq_ipc_path_csv = os.path.join(output_dir, "uniq_ipc.csv")
        if 'processed_patent_data' in locals():
            logger.info("使用步骤 1 缓存的数据进行处理。")
            patent_data_for_ipc = processed_patent_data.copy()
        else:
            logger.warning("WARN: 未找到步骤 1 缓存数据，正在从磁盘读取 ipcdata.csv。")
            patent_data_for_ipc = pd.read_csv(ipcdata_path)
        if 'ipc' not in patent_data_for_ipc.columns:
            raise KeyError("专利数据中缺少 'ipc' 列，请检查步骤 1 逻辑。")
        unique_ipc_raw = patent_data_for_ipc[patent_data_for_ipc['ipc'] != '']['ipc'].drop_duplicates().tolist()
        logger.info(f"原始唯一IPC组合数：{len(unique_ipc_raw)}")
        all_ipcs = []
        with multiprocessing.Pool(NUM_CORES) as pool:
            results = list(tqdm(pool.imap_unordered(process_ipc_string, unique_ipc_raw),
                                total=len(unique_ipc_raw),
                                desc="并行分割IPC组合",
                                unit="item",
                                ncols=80))
        for ipcs in results:
            all_ipcs.extend(ipcs)
        uniq_ipc = pd.DataFrame({'uniq_ipc': list(set(all_ipcs))})
        uniq_ipc.to_csv(uniq_ipc_path_csv, index=False, encoding=ENCODING)
        logger.info(f"提取完成！唯一IPC数量：{len(uniq_ipc)}")
        end_time_step2 = time.time()
        logger.info(f"✅ 步骤2耗时：{end_time_step2 - start_time_step2:.2f} 秒")

        # 步骤3：IPC与国民经济行业分类对照
        logger.info("\n=== 步骤3：IPC与国民经济行业分类对照（分长度保存） ====")
        start_time_step3 = time.time()
        ipc_industry_data_raw = pd.read_stata(IPC_INDUSTRY_DTA)
        ipc_industry = post_stata_read_clean(ipc_industry_data_raw).copy()
        ipc_industry.rename(columns={'国际专利分类号': 'IPC'}, inplace=True)
        ipc_industry['IPC'] = ipc_industry['IPC'].str.replace('*', '', regex=False)
        ipc_industry['IPC'] = ipc_industry['IPC'].str.replace(' ', '', regex=False)
        ipc_industry['len'] = ipc_industry['IPC'].str.len()
        len_output_dir = os.path.join(output_dir, "对照表")
        for i in progress_bar(range(3, 12), "分长度保存"):
            sub_df = ipc_industry[ipc_industry['len'] == i][['IPC', '国民经济行业代码']].copy()
            if len(sub_df) > 0:
                sub_path = os.path.join(len_output_dir, f"len{i}.csv")
                unify_column_type(sub_df, '国民经济行业代码', str)
                sub_df.to_csv(sub_path, index=False, encoding=ENCODING)
        end_time_step3 = time.time()
        logger.info(f"✅ 步骤3耗时：{end_time_step3 - start_time_step3:.2f} 秒")

        # 步骤4：全部的IPC 数据和对照表匹配
        logger.info("\n=== 步骤4：全部的IPC 数据和对照表匹配 ====")
        start_time_step4 = time.time()
        match_csv_path = os.path.join(output_dir, "专利数据与行业小类代码简易对照表.csv")
        uniq_ipc = pd.read_csv(os.path.join(output_dir, "uniq_ipc.csv"))
        if 'uniq_ipc' in uniq_ipc.columns:
            uniq_ipc['uniq_ipc'] = uniq_ipc['uniq_ipc'].astype(str)
        else:
            raise KeyError("uniq_ipc.csv 中缺少 'uniq_ipc' 列。")
        temp_match_dir = os.path.join(output_dir, "行业匹配结果")
        temp_lookup_dir = os.path.join(output_dir, "对照表")
        match_results = []
        for i in progress_bar(range(3, 12), "IPC分段匹配"):
            sub_path = os.path.join(temp_lookup_dir, f"len{i}.csv")
            if os.path.exists(sub_path):
                sub_ipc = pd.read_csv(sub_path)
                sub_ipc['IPC'] = sub_ipc['IPC'].astype(str)
                sub_ipc['国民经济行业代码'] = sub_ipc['国民经济行业代码'].astype(str).str.strip()
                current_uniq_ipc = uniq_ipc.copy()
                current_uniq_ipc['IPC_temp'] = current_uniq_ipc['uniq_ipc'].str[:i]
                matched = pd.merge(
                    current_uniq_ipc[['uniq_ipc', 'IPC_temp']],
                    sub_ipc.rename(columns={'IPC': 'IPC_temp'}),
                    on='IPC_temp',
                    how='inner'
                )
                if not matched.empty:
                    temp_csv_path = os.path.join(temp_match_dir, f"{i}.csv")
                    matched.to_csv(temp_csv_path, index=False, encoding=ENCODING)
                    match_results.append(matched[['uniq_ipc', '国民经济行业代码']].copy())
        if not match_results:
            raise RuntimeError("步骤4未能找到任何 IPC-行业匹配结果。")
        ipc_industry_match_raw = pd.concat(match_results, ignore_index=True)
        ipc_industry_match = ipc_industry_match_raw.drop_duplicates(
            subset=['uniq_ipc', '国民经济行业代码'],
            keep='first'
        ).copy()
        unify_column_type(ipc_industry_match, '国民经济行业代码')
        ipc_industry_match.to_csv(match_csv_path, index=False, encoding=ENCODING)
        logger.info(f"步骤4匹配完成！有效匹配数：{len(ipc_industry_match)}")
        end_time_step4 = time.time()
        logger.info(f"✅ 步骤4耗时：{end_time_step4 - start_time_step4:.2f} 秒")

        # 步骤5：匹配实体产业以及类型
        logger.info("\n=== 步骤5：匹配实体产业以及类型 ====")
        start_time_step5 = time.time()
        entity_csv_path = os.path.join(output_dir, "实体产业分类.csv")
        ipc_industry_data_raw = pd.read_stata(IPC_INDUSTRY_DTA)
        entity_industry = post_stata_read_clean(ipc_industry_data_raw).copy()
        entity_industry['行业门类'] = entity_industry['国民经济行业代码'].astype(str).str[:1]
        entity_industry['实体产业'] = pd.Series(dtype='object')
        entity_industry.loc[entity_industry['行业门类'] == 'C', '实体产业'] = '制造业'
        entity_industry.loc[entity_industry['行业门类'] == 'A', '实体产业'] = '农业'
        entity_industry.loc[entity_industry['行业门类'].isin(['B', 'D', 'E']), '实体产业'] = '建筑业及其他工业'
        entity_industry.loc[entity_industry['行业门类'].isin(['I', 'O']), '实体产业'] = '服务业'
        entity_industry = entity_industry[['国民经济行业代码', '实体产业']].copy()
        unify_column_type(entity_industry, '国民经济行业代码')
        entity_industry.drop_duplicates(subset=['国民经济行业代码'], keep='first', inplace=True)
        entity_industry.to_csv(entity_csv_path, index=False, encoding=ENCODING)
        logger.info(f"实体产业分类完成！分类数量：{len(entity_industry)}")
        end_time_step5 = time.time()
        logger.info(f"✅ 步骤5耗时：{end_time_step5 - start_time_step5:.2f} 秒")

        # 步骤6：数字经济核心产业分类
        logger.info("\n=== 步骤6：数字经济核心产业分类 ====")
        start_time_step6 = time.time()
        digital_csv_path = os.path.join(output_dir, "数字经济核心产业.csv")
        digital_industry_data_raw = pd.read_stata(DIGITAL_INDUSTRY_DTA)
        digital_industry = post_stata_read_clean(digital_industry_data_raw).copy()
        digital_industry = digital_industry[
            (~digital_industry['小类'].astype(str).str.startswith('05')) &
            (digital_industry['小类'] != '')
            ].copy()
        digital_industry = digital_industry[['国民经济行业代码及名称', '小类']].copy()
        digital_industry['code_list'] = digital_industry['国民经济行业代码及名称'].str.findall(r'(\d{4})')
        expanded_codes = digital_industry.explode('code_list')
        expanded_codes['数字经济产业'] = pd.Series(dtype='object')
        expanded_codes.loc[expanded_codes['小类'].astype(str).str[:2] == '01', '数字经济产业'] = '数字产品制造业'
        expanded_codes.loc[expanded_codes['小类'].astype(str).str[:2] == '02', '数字经济产业'] = '数字产品服务业'
        expanded_codes.loc[expanded_codes['小类'].astype(str).str[:2] == '03', '数字经济产业'] = '数字技术应用业'
        expanded_codes.loc[expanded_codes['小类'].astype(str).str[:2] == '04', '数字经济产业'] = '数字要素驱动业'
        digital_core = expanded_codes.dropna(subset=['code_list']).copy()
        digital_core.rename(columns={'code_list': 'industry2'}, inplace=True)
        digital_core = digital_core[['industry2', '数字经济产业']].copy()
        digital_core.drop_duplicates(inplace=True)
        unify_column_type(digital_core, 'industry2')
        digital_core.to_csv(digital_csv_path, index=False, encoding=ENCODING)
        logger.info(f"数字经济分类完成！分类数量：{len(digital_core)}")
        end_time_step6 = time.time()
        logger.info(f"✅ 步骤6耗时：{end_time_step6 - start_time_step6:.2f} 秒")

        # 步骤7：分产业（数实融合分类）
        logger.info("\n=== 步骤7：分产业（数实融合分类） - CSV 逻辑同步 ===")
        start_time_step7 = time.time()
        class_csv_path = os.path.join(output_dir, "产业数实分类2.csv")
        ipc_industry_raw = pd.read_stata(IPC_INDUSTRY_DTA)
        industry_class_df = post_stata_read_clean(ipc_industry_raw).copy()
        industry_class_df = industry_class_df[['国民经济行业代码']].drop_duplicates().copy()
        unify_column_type(industry_class_df, '国民经济行业代码')

        def extract_stata_industry2(code):
            if len(code) > 1:
                return code[1:]
            return ''

        industry_class_df['industry2'] = industry_class_df['国民经济行业代码'].apply(extract_stata_industry2)
        digital_core = pd.read_csv(os.path.join(output_dir, "数字经济核心产业.csv"))
        unify_column_type(digital_core, 'industry2')
        industry_class_df = pd.merge(
            industry_class_df, digital_core, on='industry2', how='left', indicator='_m_digital'
        )
        industry_class_df = industry_class_df[industry_class_df['_m_digital'] != 'right_only'].drop(
            columns=['_m_digital'])
        entity_industry = pd.read_csv(os.path.join(output_dir, "实体产业分类.csv"))
        unify_column_type(entity_industry, '国民经济行业代码')
        industry_class_df.drop(columns=['实体产业'], inplace=True, errors='ignore')
        industry_class_df = pd.merge(
            industry_class_df, entity_industry, on='国民经济行业代码', how='left', indicator='_m_entity'
        )
        industry_class_df.drop(columns=['industry2', '_m_entity'], inplace=True, errors='ignore')
        industry_class_df['数字经济产业'] = industry_class_df['数字经济产业'].replace('', np.nan)
        industry_class_df['实体产业'] = industry_class_df['实体产业'].replace('', np.nan)
        industry_class_df['类别'] = np.where(
            industry_class_df['数字经济产业'].notna(),
            '数字经济产业',
            industry_class_df['实体产业']
        )
        industry_class = industry_class_df.dropna(subset=['类别'])[
            ['国民经济行业代码', '类别']].drop_duplicates().copy()
        industry_class.to_csv(class_csv_path, index=False, encoding=ENCODING)
        logger.info(f"产业分类整合完成！最终分类记录数：{len(industry_class)}")
        end_time_step7 = time.time()
        logger.info(f"✅ 步骤7耗时：{end_time_step7 - start_time_step7:.2f} 秒")

        # 步骤8：分市构建IPC融合矩阵
        logger.info("\n=== 步骤8：分市构建IPC融合矩阵 (多进程加速) - 输出 CSV ===")
        start_time_step8 = time.time()
        ipcmat_csv_path = os.path.join(output_dir, "ipcmat.csv")
        ipc_data = pd.read_csv(os.path.join(output_dir, "ipcdata.csv"))
        ipc_count = ipc_data.groupby(['city', 'ipc']).size().reset_index(name='_freq')
        logger.info(f"市级-IPC组合数：{len(ipc_count)}")
        grouped_data = ipc_count.groupby('city')
        city_groups = [(city, data) for city, data in grouped_data]
        all_csv_lines = []
        with multiprocessing.Pool(NUM_CORES) as pool:
            results = list(tqdm(pool.imap_unordered(process_city_data, city_groups),
                                total=len(city_groups),
                                desc="并行生成组合矩阵",
                                unit="city",
                                ncols=80))
            all_csv_lines = "".join(results)
        res3_path = os.path.join(output_dir, "res3.csv")
        with open(res3_path, 'w', encoding=ENCODING) as fh:
            fh.write("city,ipc1,ipc2,value\n")
            fh.write(all_csv_lines)
        ipc_mat = pd.read_csv(res3_path, encoding=ENCODING)
        ipc_mat['value'] = pd.to_numeric(ipc_mat['value'], errors='coerce')
        ipc_mat = ipc_mat.groupby(['city', 'ipc1', 'ipc2'])['value'].sum().reset_index()
        ipc_mat = ipc_mat[ipc_mat['ipc1'] != ipc_mat['ipc2']].copy()
        ipc_mat.to_csv(ipcmat_csv_path, index=False, encoding=ENCODING)
        logger.info(f"IPC融合矩阵完成！形状：{ipc_mat.shape}")
        end_time_step8 = time.time()
        logger.info(f"✅ 步骤8耗时：{end_time_step8 - start_time_step8:.2f} 秒")

        # 步骤9：优化版 - 分块匹配+分块聚合（核心修改）
        logger.info("\n=== 步骤9：匹配行业小类 (分块聚合优化) - 解决内存溢出 ===")
        start_time_step9 = time.time()
        industrymat_csv_path = os.path.join(output_dir, "industrymat.csv")

        # 配置参数（优化内存占用）
        MERGE_CHUNK_SIZE = 50000  # 减小分块大小，进一步降低内存压力
        MAX_IPC_LENGTH = 15
        MIN_MATCH_LENGTH = 3
        TEMP_AGG_DIR = os.path.join(output_dir, "temp_agg_chunks")  # 临时聚合块目录
        os.makedirs(TEMP_AGG_DIR, exist_ok=True)

        # 1. 读取并预处理对照表
        logger.info("👉 9.1 读取并预处理IPC-行业对照表...")
        ipc_industry_match = pd.read_csv(os.path.join(output_dir, "专利数据与行业小类代码简易对照表.csv"))
        ipc_industry_match = unify_column_type(ipc_industry_match, '国民经济行业代码', str)
        ipc_industry_match['uniq_ipc'] = ipc_industry_match['uniq_ipc'].astype(str).str.strip()
        ipc_industry_match = ipc_industry_match[
            (ipc_industry_match['uniq_ipc'] != '') &
            (ipc_industry_match['国民经济行业代码'] != '') &
            (ipc_industry_match['uniq_ipc'].str.len() >= MIN_MATCH_LENGTH) &
            (ipc_industry_match['uniq_ipc'].str.len() <= MAX_IPC_LENGTH)
            ].copy()
        ipc_industry_match = ipc_industry_match.drop_duplicates(
            subset=['uniq_ipc', '国民经济行业代码'],
            keep='first'
        ).copy()
        logger.info(
            f"👉 9.1 对照表预处理完成：有效匹配数 {len(ipc_industry_match)}，唯一IPC数 {ipc_industry_match['uniq_ipc'].nunique()}")

        # 2. 验证步骤8输出文件
        ipcmat_read_path = os.path.join(output_dir, "ipcmat.csv")
        if not os.path.exists(ipcmat_read_path):
            raise FileNotFoundError(f"未找到步骤8输出文件：{ipcmat_read_path}")
        file_size = os.path.getsize(ipcmat_read_path) / (1024 * 1024)
        logger.info(f"👉 9.2 步骤8输出文件大小：{file_size:.2f} MB，准备分块读取+分块聚合")

        # 3. 定义带分块聚合的匹配函数
        def match_and_agg_chunk(chunk, ipc_industry_map, chunk_idx):
            """匹配+分块内聚合，直接输出聚合后的小文件"""
            # 数据清洗
            chunk['ipc1'] = chunk['ipc1'].astype(str).str.strip().str[:MAX_IPC_LENGTH]
            chunk['ipc2'] = chunk['ipc2'].astype(str).str.strip().str[:MAX_IPC_LENGTH]
            chunk['value'] = pd.to_numeric(chunk['value'], errors='coerce').fillna(0)

            # 过滤无效数据
            chunk = chunk[
                (chunk['ipc1'] != '') &
                (chunk['ipc2'] != '') &
                (chunk['value'] > 0) &
                (chunk['ipc1'].str.len() >= MIN_MATCH_LENGTH) &
                (chunk['ipc2'].str.len() >= MIN_MATCH_LENGTH)
                ].copy()

            if chunk.empty:
                return None

            # 第一次匹配：ipc1 -> industry1
            chunk = chunk.merge(
                ipc_industry_map.rename(columns={'uniq_ipc': 'ipc1', '国民经济行业代码': 'industry1'}),
                on='ipc1',
                how='left'
            )

            # 第二次匹配：ipc2 -> industry2
            chunk = chunk.merge(
                ipc_industry_map.rename(columns={'uniq_ipc': 'ipc2', '国民经济行业代码': 'industry2'}),
                on='ipc2',
                how='left'
            )

            # 过滤未匹配到行业的记录
            chunk = chunk[
                (chunk['industry1'].notna()) &
                (chunk['industry2'].notna()) &
                (chunk['industry1'] != '') &
                (chunk['industry2'] != '')
                ].copy()

            if chunk.empty:
                return None

            # 行业代码标准化
            chunk['industry1'] = chunk['industry1'].str[:5]
            chunk['industry2'] = chunk['industry2'].str[:5]

            # 排除自身融合
            chunk.loc[chunk['industry1'] == chunk['industry2'], 'value'] = 0
            chunk = chunk[chunk['value'] > 0].copy()

            if chunk.empty:
                return None

            # 分块内聚合（核心优化！）
            chunk_agg = chunk.groupby(['city', 'industry1', 'industry2'], as_index=False)['value'].sum()

            # 保存临时聚合块
            temp_agg_path = os.path.join(TEMP_AGG_DIR, f"agg_chunk_{chunk_idx}.csv")
            chunk_agg.to_csv(temp_agg_path, index=False, encoding=ENCODING)
            logger.info(f"👉 9.3 第 {chunk_idx + 1} 块：处理 {len(chunk)} 行，聚合后 {len(chunk_agg)} 行")

            return temp_agg_path

        # 4. 分块读取并处理（每个块匹配后立即聚合）
        iterator = pd.read_csv(
            ipcmat_read_path,
            chunksize=MERGE_CHUNK_SIZE,
            encoding=ENCODING,
            low_memory=True
        )

        temp_agg_paths = []
        for idx, chunk in enumerate(progress_bar(iterator, f"分块匹配+聚合 {MERGE_CHUNK_SIZE} 行/块")):
            try:
                temp_path = match_and_agg_chunk(chunk, ipc_industry_match, idx)
                if temp_path is not None:
                    temp_agg_paths.append(temp_path)
                # 释放内存
                del chunk
            except Exception as e:
                logger.error(f"👉 9.3 第 {idx + 1} 块处理出错：{str(e)}", exc_info=True)
                continue

        # 5. 检查临时聚合块
        if not temp_agg_paths:
            raise RuntimeError("步骤9未生成有效匹配结果，请检查对照表或ipcmat.csv数据")
        logger.info(f"👉 9.4 分块聚合完成：生成 {len(temp_agg_paths)} 个临时聚合块")

        # 6. 合并所有临时聚合块（最终聚合）
        logger.info("👉 9.5 合并所有临时聚合块...")
        final_agg_chunks = []
        for temp_path in progress_bar(temp_agg_paths, "读取临时聚合块"):
            chunk = pd.read_csv(temp_path, encoding=ENCODING)
            # 确保value是数值类型
            chunk['value'] = pd.to_numeric(chunk['value'], errors='coerce').astype(np.float32)
            final_agg_chunks.append(chunk)
            # 可选：删除临时文件释放磁盘空间
            os.remove(temp_path)

        # 合并并最终聚合
        industry_mat = pd.concat(final_agg_chunks, ignore_index=True)
        del final_agg_chunks  # 释放内存

        # 最终聚合（合并相同city-industry1-industry2的记录）
        industry_mat = industry_mat.groupby(['city', 'industry1', 'industry2'], as_index=False)['value'].sum()

        # 最终过滤
        industry_mat = industry_mat[
            (industry_mat['value'] > 0) &
            (industry_mat['industry1'] != industry_mat['industry2']) &
            (industry_mat['industry1'] != '') &
            (industry_mat['industry2'] != '')
            ].copy()

        # 7. 保存最终结果
        industry_mat.to_csv(industrymat_csv_path, index=False, encoding=ENCODING)

        # 删除临时目录
        os.rmdir(TEMP_AGG_DIR)

        # 输出结果统计
        logger.info(f"✅ 步骤9：行业小类匹配完成！")
        logger.info(f"  - 最终有效记录数：{len(industry_mat)}")
        logger.info(f"  - 涉及城市数：{industry_mat['city'].nunique()}")
        logger.info(f"  - 涉及行业1数：{industry_mat['industry1'].nunique()}")
        logger.info(f"  - 涉及行业2数：{industry_mat['industry2'].nunique()}")
        logger.info(f"  - 价值总和：{industry_mat['value'].sum():.2f}")
        logger.info(f"  - 前5条记录预览：")
        logger.info(f"{industry_mat.head().to_string(index=False)}")

        end_time_step9 = time.time()
        logger.info(f"✅ 步骤9耗时：{end_time_step9 - start_time_step9:.2f} 秒")

        # 步骤10：计算数实融合
        logger.info("\n=== 步骤10：计算数实融合 (优化版) - CSV 逻辑同步 ===")
        start_time_step10 = time.time()
        final_csv_path = os.path.join(output_dir, "数实融合指数_市级层面.csv")
        industry_mat = pd.read_csv(os.path.join(output_dir, "industrymat.csv"))
        industry_class = pd.read_csv(os.path.join(output_dir, "产业数实分类2.csv"))
        unify_column_type(industry_class, '国民经济行业代码')
        unify_column_type(industry_mat, 'industry1')
        unify_column_type(industry_mat, 'industry2')
        class_lookup = industry_class.set_index('国民经济行业代码')['类别'].copy()
        industry_mat['class1'] = industry_mat['industry1'].map(class_lookup)
        industry_mat['class2'] = industry_mat['industry2'].map(class_lookup)
        industry_mat = industry_mat[industry_mat['value'] != 0]
        industry_mat.dropna(subset=['class1', 'class2'], inplace=True)
        industry_mat['freq1'] = industry_mat.groupby(['city', 'class1'])['industry1'].transform('nunique')
        industry_mat['freq2'] = industry_mat.groupby(['city', 'class2'])['industry2'].transform('nunique')
        result_df = industry_mat.groupby(['city', 'class1', 'class2']).agg(
            value=('value', 'sum'),
            freq1=('freq1', 'first'),
            freq2=('freq2', 'first')
        ).reset_index()
        result_df['denom'] = result_df['freq1'] * result_df['freq2']
        result_df['RH'] = np.divide(result_df['value'], result_df['denom'],
                                    out=np.zeros_like(result_df['value'], dtype=float),
                                    where=result_df['denom'] != 0)
        result_df['RH'] = result_df['RH'].round(6)
        result_df.drop(columns=['denom'], inplace=True)
        result_df = result_df[
            (result_df['class1'] == '数字经济产业') &
            (result_df['class1'] != result_df['class2'])
            ].copy()
        if result_df.empty:
            rh_wide = pd.DataFrame({'city': []})
        else:
            result_df['class'] = result_df['class1'] + "_" + result_df['class2']
            result_df = result_df[['city', 'class', 'RH']].drop_duplicates(keep='first')
            rh_wide = result_df.pivot_table(
                index='city',
                columns='class',
                values='RH',
                fill_value=0
            ).reset_index()
            sort_col = '数字经济产业_制造业'
            if sort_col in rh_wide.columns:
                rh_wide = rh_wide.sort_values(sort_col, ascending=False)
            else:
                rh_wide = rh_wide.sort_values('city')
        rh_wide.to_csv(final_csv_path, index=False, encoding=ENCODING)
        logger.info("\n=== 数实融合指数测算完成！===")
        logger.info(f"最终结果保存在：{final_csv_path}")
        end_time_step10 = time.time()
        logger.info(f"✅ 步骤10耗时：{end_time_step10 - start_time_step10:.2f} 秒")

        success = True

    except Exception as e:
        logger.error(f"\n❌ 处理文件 {os.path.basename(patent_csv_path)} 时出错：{str(e)}", exc_info=True)
        # 清理当前文件的输出目录（可选，如需保留错误中间文件可删除）
        # import shutil
        # shutil.rmtree(output_dir, ignore_errors=True)
    finally:
        file_end_time = time.time()
        logger.info(f"\n{'=' * 80}")
        logger.info(f"文件 {os.path.basename(patent_csv_path)} 处理结束")
        logger.info(f"处理状态：{'成功' if success else '失败'}")
        logger.info(f"文件总耗时：{(file_end_time - file_start_time) / 60:.2f} 分钟")
        logger.info(f"{'=' * 80}\n")

    return success


# ===================== 主执行逻辑（批量处理） =====================
if __name__ == '__main__':
    start_time_total = time.time()
    logger.info(f"\n🚀 数实融合测算批量处理脚本启动：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 🚀\n")

    try:
        # 1. 扫描指定目录下所有CSV文件
        logger.info(f"📁 正在扫描目录：{PATENT_CSV_DIR}")
        patent_csv_files = [
            f for f in os.listdir(PATENT_CSV_DIR)
            if f.lower().endswith('.csv') and os.path.isfile(os.path.join(PATENT_CSV_DIR, f))
        ]

        if not patent_csv_files:
            logger.warning("⚠️  在指定目录中未找到任何CSV文件！")
            exit(0)

        logger.info(f"✅ 找到 {len(patent_csv_files)} 个CSV文件待处理：")
        for i, csv_file in enumerate(patent_csv_files, 1):
            logger.info(f"  {i}. {csv_file}")

        # 2. 循环处理每个CSV文件
        logger.info(f"\n🚀 开始批量处理（共 {len(patent_csv_files)} 个文件）")
        success_count = 0
        fail_count = 0
        processed_files = []

        for csv_file_name in patent_csv_files:
            # 构建完整路径
            patent_csv_path = os.path.join(PATENT_CSV_DIR, csv_file_name)

            # 创建该文件的专属输出目录（以文件名命名，避免冲突）
            file_output_dir_name = os.path.splitext(csv_file_name)[0] + "_测算结果"
            file_output_dir = os.path.join(ROOT_OUTPUT_DIR, file_output_dir_name)

            # 跳过已处理的文件（如需重新处理，删除对应输出目录即可）
            if os.path.exists(os.path.join(file_output_dir, "数实融合指数_市级层面.csv")):
                logger.info(f"\n⚠️ 文件 {csv_file_name} 已处理完成，跳过（如需重新处理请删除对应输出目录）")
                success_count += 1
                processed_files.append((csv_file_name, "已处理（跳过）"))
                continue

            # 执行单个文件处理
            if process_single_patent_csv(patent_csv_path, file_output_dir):
                success_count += 1
                processed_files.append((csv_file_name, "成功"))
            else:
                fail_count += 1
                processed_files.append((csv_file_name, "失败"))

        # 3. 输出批量处理总结报告
        logger.info("\n" + "=" * 80)
        logger.info("=== 批量处理总结报告 ===")
        logger.info(f"📊 总文件数：{len(patent_csv_files)}")
        logger.info(f"✅ 处理成功：{success_count} 个")
        logger.info(f"❌ 处理失败：{fail_count} 个")
        logger.info(f"\n📋 详细处理状态：")
        for file_name, status in processed_files:
            logger.info(f"  - {file_name}: {status}")
        logger.info(f"\n📁 所有结果文件保存在：{ROOT_OUTPUT_DIR}")
        logger.info("=" * 80)

    except Exception as e:
        logger.error(f"\n❌ 批量处理程序发生致命错误：{str(e)}", exc_info=True)
        raise

    end_time_total = time.time()
    total_minutes = (end_time_total - start_time_total) / 60
    logger.info(f"\n🎉 批量处理总运行时间：{total_minutes:.2f} 分钟（{end_time_total - start_time_total:.2f} 秒）")
    logger.info("批量处理完成！")