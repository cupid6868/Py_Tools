import pandas as pd
import os
import time
from tqdm import tqdm
import logging

# ===================== 日志初始化 =====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# ===================== 核心配置 =====================
# 🚨 DTA文件所在文件夹路径（批量处理该文件夹下所有.dta文件）
DTA_FOLDER_PATH = r"E:\1.专利数据"

# 🚨 输出CSV文件的文件夹路径（默认与DTA文件夹相同，可自行修改）
CSV_FOLDER_PATH = DTA_FOLDER_PATH  # 如需单独存放，可改为：r"E:\1.专利数据\CSV输出"

# 每次读取DTA的块大小（可根据内存调整，建议10-50万行）
READ_CHUNK_SIZE = 200000  # 内存不足可减小

# 编码格式
ENCODING = 'utf-8'

# 只需要读取的列名，减少内存占用（按需修改）
REQUIRED_READ_COLS = ['ipzlid', 'IPC', '市']


# =======================================================
# 工具函数：单个DTA文件转CSV
# =======================================================
def convert_dta_to_csv(dta_file_path, csv_file_path):
    """
    单个DTA文件分块转换为CSV文件
    :param dta_file_path: 输入DTA文件路径
    :param csv_file_path: 输出CSV文件路径
    """
    logger.info(f"\n【开始处理】{os.path.basename(dta_file_path)}")
    start_time_single = time.time()

    try:
        # 检查输入文件是否存在
        if not os.path.exists(dta_file_path):
            raise FileNotFoundError(f"DTA文件不存在：{dta_file_path}")

        # 分块读取DTA文件
        dta_iterator = pd.read_stata(
            dta_file_path,
            columns=REQUIRED_READ_COLS,
            chunksize=READ_CHUNK_SIZE
        )

        # 分块写入CSV
        first_write = True
        total_rows_written = 0

        for chunk_idx, chunk in enumerate(tqdm(
                dta_iterator, desc=f"转换 {os.path.basename(dta_file_path)}", unit="chunk", ncols=100
        )):
            # 数据清理：处理字符串列的NaN
            for col in chunk.columns:
                if chunk[col].dtype == 'object':
                    chunk[col] = chunk[col].astype(str).replace('nan', '')

            # 写入CSV（首次带表头，后续追加）
            if first_write:
                chunk.to_csv(
                    csv_file_path,
                    index=False,
                    encoding=ENCODING,
                    mode='w'
                )
                first_write = False
            else:
                chunk.to_csv(
                    csv_file_path,
                    index=False,
                    encoding=ENCODING,
                    mode='a',
                    header=False
                )

            # 统计进度
            chunk_rows = len(chunk)
            total_rows_written += chunk_rows
            logger.info(f"  已处理第 {chunk_idx + 1} 块，累计写入 {total_rows_written} 行")

        # 验证结果
        if os.path.exists(csv_file_path):
            file_size = os.path.getsize(csv_file_path) / (1024 * 1024 * 1024)  # GB
            logger.info(f"✅ 单个文件转换完成！")
            logger.info(f"  CSV文件大小：{file_size:.2f} GB")
            logger.info(f"  累计写入行数：{total_rows_written}")
            logger.info(f"  输出路径：{csv_file_path}")
            logger.info(f"  耗时：{(time.time() - start_time_single) / 60:.2f} 分钟")
            return True
        else:
            raise RuntimeError("未生成CSV文件")

    except Exception as e:
        logger.error(f"❌ 处理 {os.path.basename(dta_file_path)} 时出错：{str(e)}", exc_info=True)
        # 若转换失败，删除可能残留的不完整CSV文件
        if os.path.exists(csv_file_path):
            os.remove(csv_file_path)
            logger.warning(f"  已删除不完整的CSV文件：{csv_file_path}")
        return False


# =======================================================
# 主程序：批量处理文件夹下所有DTA文件
# =======================================================
if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("=== DTA文件批量转换为CSV程序（分块读取模式）===")
    logger.info("=" * 60)
    total_start_time = time.time()

    try:
        # 1. 检查并创建输出文件夹
        if not os.path.exists(CSV_FOLDER_PATH):
            os.makedirs(CSV_FOLDER_PATH, exist_ok=True)
            logger.info(f"📁 已创建输出文件夹：{CSV_FOLDER_PATH}")
        else:
            logger.info(f"📁 输出文件夹：{CSV_FOLDER_PATH}")

        # 2. 获取文件夹下所有.dta文件（不包含子文件夹）
        dta_files = [
            f for f in os.listdir(DTA_FOLDER_PATH)
            if f.lower().endswith('.dta') and os.path.isfile(os.path.join(DTA_FOLDER_PATH, f))
        ]

        if not dta_files:
            logger.warning("⚠️  在指定文件夹中未找到任何DTA文件！")
            exit(0)

        logger.info(f"\n📊 找到 {len(dta_files)} 个DTA文件待转换：")
        for i, dta_file in enumerate(dta_files, 1):
            logger.info(f"  {i}. {dta_file}")

        # 3. 循环处理每个DTA文件
        logger.info(f"\n🚀 开始批量转换（共 {len(dta_files)} 个文件）")
        success_count = 0
        fail_count = 0

        for dta_file_name in dta_files:
            # 构建完整路径
            dta_file_path = os.path.join(DTA_FOLDER_PATH, dta_file_name)
            # 生成同名CSV文件名（替换后缀）
            csv_file_name = os.path.splitext(dta_file_name)[0] + '.csv'
            csv_file_path = os.path.join(CSV_FOLDER_PATH, csv_file_name)

            # 跳过已存在的CSV文件（如需覆盖，可删除以下判断）
            if os.path.exists(csv_file_path):
                logger.info(f"\n⚠️ {csv_file_name} 已存在，跳过转换")
                continue

            # 执行转换
            if convert_dta_to_csv(dta_file_path, csv_file_path):
                success_count += 1
            else:
                fail_count += 1

        # 4. 输出总体统计
        logger.info("\n" + "=" * 60)
        logger.info("=== 批量转换完成 ===")
        logger.info(f"📈 总文件数：{len(dta_files)}")
        logger.info(f"✅ 成功转换：{success_count} 个")
        logger.info(f"❌ 转换失败：{fail_count} 个")
        logger.info(f"⌛ 总耗时：{(time.time() - total_start_time) / 60:.2f} 分钟")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"❌ 批量转换程序发生致命错误：{str(e)}", exc_info=True)
        raise