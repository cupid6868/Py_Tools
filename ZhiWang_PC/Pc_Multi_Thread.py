# -*- coding: utf-8 -*-
"""
CNKI 知网爬虫（多线程优化版）
- 核心优化：多线程并行提取摘要，速度提升3-5倍
- 反爬兼容：控制并发数、复用Cookies、随机延迟
- 功能保留：验证码人工处理、断点续爬、快照留存
"""

import os
import time
import random
import pickle
import traceback
import requests
from datetime import datetime
from typing import List, Dict, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup  # 用于解析摘要页

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException
)
from webdriver_manager.chrome import ChromeDriverManager


# ====================== 配置 ======================
class Config:
    KEYWORDS = "减污降碳"  # 搜索关键词
    MAX_PAGES = 5  # 最大爬取页数
    BASE_URL = "https://www.cnki.net/"
    OUTPUT_DIR = "cnki_papers"
    WAIT_TIME = 15  # 元素等待超时时间（秒）
    HEADLESS = False  # 是否无头模式（False可看到浏览器操作）

    # 多线程配置（关键优化）
    MAX_THREADS = 3  # 并发数（建议3-5，过高易触发反爬）
    MIN_DELAY = 0.8  # 基础操作延迟（秒）
    MAX_DELAY = 1.8

    # 输入模拟参数
    TYPE_MIN = 0.05  # 按键间隔最小值
    TYPE_MAX = 0.15  # 按键间隔最大值

    # 断点与存储
    CKPT_EVERY_PAGES = 1  # 每爬多少页保存一次断点
    COOKIE_PATH = os.path.join(OUTPUT_DIR, "session_cookies.pkl")
    SNAPSHOT_DIR = os.path.join(OUTPUT_DIR, "snapshots")
    LOG_PATH = os.path.join(OUTPUT_DIR, "run.log")

    # 代理配置（如需使用代理，在PROXY_POOL添加代理地址）
    USE_PROXY = False
    PROXY_POOL = []  # 格式：["http://ip:port", "https://ip:port"]

    # 浏览器指纹
    UA_MIN = 118  # Chrome版本范围
    UA_MAX = 124
    MAX_RETRY = 3  # 操作失败重试次数

    # 摘要页请求头（模拟浏览器）
    ABSTRACT_HEADERS = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }

    # 页面元素选择器
    NEXT_SELECTORS = [
        (By.CSS_SELECTOR, "a[title='下一页']"),
        (By.LINK_TEXT, "下一页"),
        (By.XPATH, "//a[contains(@class,'btn-next') or contains(text(),'下一页')]"),
    ]
    TABLE_SELECTORS = [
        (By.CSS_SELECTOR, ".result-table-list"),
        (By.CSS_SELECTOR, "table.GridTableContent"),
    ]
    TITLE_IN_COL = 1
    AUTHOR_COL_CANDIDATES = [2, 3]
    SOURCE_COL_CANDIDATES = [3, 2, 4]
    DATE_COL_CANDIDATES = [4, 5]


# ====================== 工具函数 ======================
def ensure_dirs():
    """确保输出目录存在"""
    os.makedirs(Config.OUTPUT_DIR, exist_ok=True)
    os.makedirs(Config.SNAPSHOT_DIR, exist_ok=True)


def log(msg: str):
    """日志输出（控制台+文件）"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(Config.LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def jitter(min_s=None, max_s=None):
    """随机延迟（模拟人工操作间隔）"""
    a = Config.MIN_DELAY if min_s is None else min_s
    b = Config.MAX_DELAY if max_s is None else max_s
    time.sleep(random.uniform(a, b))


def choose_proxy() -> str:
    """选择随机代理（如需使用）"""
    if Config.USE_PROXY and Config.PROXY_POOL:
        return random.choice(Config.PROXY_POOL)
    return ""


def snapshot(driver: webdriver.Chrome, tag: str = "snapshot"):
    """保存页面快照（HTML+截图）"""
    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        base = os.path.join(Config.SNAPSHOT_DIR, f"{ts}_{tag}")
        html_path = base + ".html"
        png_path = base + ".png"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.save_screenshot(png_path)
        log(f"已保存快照：{html_path} / {png_path}")
    except Exception as e:
        log(f"保存快照失败：{e}")


def backoff_sleep(retry_idx: int):
    """指数退避等待（失败重试时使用）"""
    base = 1.5
    t = (base ** retry_idx) + random.random()
    log(f"指数退避等待 {t:.2f}s ...")
    time.sleep(t)


def human_typing(el, text: str):
    """模拟人类打字（带随机间隔）"""
    for ch in text:
        el.send_keys(ch)
        time.sleep(random.uniform(Config.TYPE_MIN, Config.TYPE_MAX))


def human_scroll(driver: webdriver.Chrome):
    """模拟人类滚动页面"""
    times = random.randint(2, 5)
    for _ in range(times):
        y = random.randint(300, 900)
        driver.execute_script(f"window.scrollBy(0, {y});")
        time.sleep(random.uniform(0.4, 1.2))
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(random.uniform(0.2, 0.6))


def human_mouse_wiggle(driver: webdriver.Chrome):
    """模拟人类鼠标移动"""
    try:
        actions = ActionChains(driver)
        for _ in range(random.randint(2, 5)):
            x_off = random.randint(-30, 30)
            y_off = random.randint(-20, 20)
            actions.move_by_offset(x_off, y_off).pause(random.uniform(0.1, 0.3))
        actions.perform()
    except Exception:
        pass


def create_requests_session(driver: webdriver.Chrome) -> requests.Session:
    """创建带Cookies的requests会话（复用Selenium登录状态）"""
    session = requests.Session()
    # 添加重试机制（防止网络波动）
    retry = Retry(total=3, backoff_factor=0.5)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    # 从Selenium获取Cookies并同步到requests
    cookies = driver.get_cookies()
    for cookie in cookies:
        session.cookies.set(
            cookie["name"],
            cookie["value"],
            domain=cookie.get("domain"),
            path=cookie.get("path", "/")
        )

    # 设置随机User-Agent（与浏览器保持一致风格）
    ua_ver = random.randint(Config.UA_MIN, Config.UA_MAX)
    session.headers.update({
        "User-Agent": f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{ua_ver}.0.0.0 Safari/537.36",
        **Config.ABSTRACT_HEADERS
    })
    return session


# ====================== 浏览器初始化 ======================
def init_browser() -> webdriver.Chrome:
    """初始化Chrome浏览器（带反爬配置）"""
    options = webdriver.ChromeOptions()
    if Config.HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")

    # 随机窗口大小
    width = random.randint(1280, 1920)
    height = random.randint(800, 1080)
    options.add_argument(f"--window-size={width},{height}")

    # 随机User-Agent
    ua_ver = random.randint(Config.UA_MIN, Config.UA_MAX)
    ua = f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{ua_ver}.0.0.0 Safari/537.36"
    options.add_argument(f"user-agent={ua}")
    options.add_argument("--lang=zh-CN,zh;q=0.9")

    # 配置代理（如需）
    proxy = choose_proxy()
    if proxy:
        options.add_argument(f"--proxy-server={proxy}")
        log(f"使用代理：{proxy}")

    # 反指纹配置
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-popup-blocking")

    # 初始化驱动（手动指定路径时替换下面一行）
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    # 进一步隐藏自动化特征
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": r"""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(navigator, 'languages', {get: () => ['zh-CN', 'zh']});
        Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});
        """
    })
    return driver


# ====================== Cookies 会话管理 ======================
def save_cookies(driver: webdriver.Chrome):
    """保存Cookies（复用会话）"""
    try:
        cookies = driver.get_cookies()
        with open(Config.COOKIE_PATH, "wb") as f:
            pickle.dump(cookies, f)
        log(f"Cookies 已保存：{Config.COOKIE_PATH}")
    except Exception as e:
        log(f"保存 Cookies 失败：{e}")


def load_cookies(driver: webdriver.Chrome, url_for_scope: str):
    """加载Cookies（恢复会话）"""
    try:
        if os.path.exists(Config.COOKIE_PATH):
            with open(Config.COOKIE_PATH, "rb") as f:
                cookies = pickle.load(f)
            for ck in cookies:
                # 清理无效字段
                for k in ["sameSite", "httpOnly", "secure"]:
                    if k in ck and ck[k] is None:
                        ck.pop(k, None)
                try:
                    driver.add_cookie(ck)
                except Exception:
                    pass
            log("Cookies 已载入")
        else:
            log("未发现 Cookies 文件，跳过载入")
    except Exception as e:
        log(f"载入 Cookies 失败：{e}")


# ====================== 验证码处理 ======================
def wait_for_captcha(driver: webdriver.Chrome) -> bool:
    """检测并等待用户处理验证码"""
    hints = [
        "//iframe[contains(@src,'validate')]",
        "//iframe[contains(@src,'captcha')]",
        "//iframe[contains(@src,'verify')]",
        "//*[contains(text(),'请完成验证')]",
        "//*[contains(text(),'验证失败')]",
    ]
    for xp in hints:
        try:
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, xp)))
            log("=== 检测到验证码 ===")
            snapshot(driver, tag="captcha")
            input("请在浏览器中完成验证码，完成后回车继续...")
            return True
        except TimeoutException:
            continue
    return False


# ====================== 页面操作 ======================
def smart_get(driver: webdriver.Chrome, url: str):
    """带重试的页面访问"""
    for i in range(Config.MAX_RETRY):
        try:
            driver.get(url)
            jitter(1.0, 2.0)
            return
        except WebDriverException as e:
            log(f"打开页面失败：{e}")
            backoff_sleep(i)
    raise RuntimeError("多次尝试无法打开页面")


def get_search_elements(driver: webdriver.Chrome) -> Tuple:
    """获取搜索框和搜索按钮元素"""
    search_box_candidates = [
        (By.ID, "txt_SearchText"),
        (By.CSS_SELECTOR, "input#txt_SearchText"),
        (By.CSS_SELECTOR, "input.search-text"),
        (By.XPATH, "//input[@id='txt_SearchText' or @name='txt_SearchText']"),
    ]
    search_btn_candidates = [
        (By.ID, "btnSearch"),
        (By.CSS_SELECTOR, "#btnSearch"),
        (By.XPATH, "//input[@value='检索' or @value='搜索']"),
        (By.XPATH, "//button[contains(text(),'检索') or contains(text(),'搜索')]"),
    ]

    box = btn = None
    last_err = None
    for i in range(Config.MAX_RETRY):
        try:
            # 查找搜索框
            for by, sel in search_box_candidates:
                try:
                    box = WebDriverWait(driver, Config.WAIT_TIME).until(
                        EC.element_to_be_clickable((by, sel))
                    )
                    break
                except TimeoutException:
                    continue
            # 查找搜索按钮
            for by, sel in search_btn_candidates:
                try:
                    btn = WebDriverWait(driver, Config.WAIT_TIME).until(
                        EC.element_to_be_clickable((by, sel))
                    )
                    break
                except TimeoutException:
                    continue
            if box is None:
                raise NoSuchElementException("未找到搜索框")
            return box, btn
        except Exception as e:
            last_err = e
            snapshot(driver, tag="search_elems_fail")
            backoff_sleep(i)
    raise last_err


def perform_search(driver: webdriver.Chrome, keywords: str) -> bool:
    """执行搜索操作"""
    for i in range(Config.MAX_RETRY):
        try:
            smart_get(driver, Config.BASE_URL)
            wait_for_captcha(driver)  # 处理首页验证码

            # 尝试加载Cookies恢复会话
            try:
                load_cookies(driver, Config.BASE_URL)
                driver.refresh()
                jitter()
            except Exception:
                pass

            # 输入关键词并搜索
            box, btn = get_search_elements(driver)
            box.click()
            time.sleep(random.uniform(0.2, 0.5))
            box.clear()
            human_typing(box, keywords)  # 模拟人工输入
            jitter()
            human_mouse_wiggle(driver)  # 模拟鼠标移动
            if btn:
                btn.click()
            else:
                box.send_keys(Keys.ENTER)  # 无按钮时按回车搜索

            jitter(2.0, 3.5)
            wait_for_captcha(driver)  # 处理搜索后验证码
            save_cookies(driver)  # 保存会话
            return True
        except Exception as e:
            log(f"执行搜索失败：{e}")
            snapshot(driver, tag="perform_search_fail")
            backoff_sleep(i)
    return False


def find_results_table(driver: webdriver.Chrome):
    """查找结果列表表格"""
    for i in range(Config.MAX_RETRY):
        try:
            for by, sel in Config.TABLE_SELECTORS:
                try:
                    tbl = WebDriverWait(driver, Config.WAIT_TIME).until(
                        EC.presence_of_element_located((by, sel))
                    )
                    return tbl
                except TimeoutException:
                    continue
            raise NoSuchElementException("未找到结果列表")
        except Exception as e:
            log(f"等待结果表失败：{e}")
            backoff_sleep(i)
    return None


def fetch_abstract(session: requests.Session, href: str) -> str:
    """多线程任务：提取单篇论文摘要"""
    if not href:
        return ""
    try:
        # 随机延迟避免高频请求
        time.sleep(random.uniform(0.3, 0.8))
        response = session.get(href, timeout=10)
        response.encoding = "utf-8"
        html = response.text

        # 解析摘要（适配多种页面结构）
        soup = BeautifulSoup(html, "html.parser")
        abstract_elems = soup.select(
            "span.abstract-text, div.abstract-text, p.abstract-text, "
            ".abstract, #ChDivSummary, .summary-content"
        )
        if abstract_elems:
            return abstract_elems[0].get_text(strip=True)
        return ""
    except Exception as e:
        log(f"提取摘要失败（{href[:50]}...）：{str(e)[:50]}")
        return ""


def extract_papers_from_table(driver: webdriver.Chrome) -> List[Dict]:
    """从结果表格提取论文信息（多线程提取摘要）"""
    tbl = find_results_table(driver)
    if tbl is None:
        return []
    human_scroll(driver)  # 模拟浏览
    rows = tbl.find_elements(By.TAG_NAME, "tr")
    papers = []
    href_list = []  # 存储(索引, 链接)用于多线程

    # 第一步：提取列表页基础信息
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) < 3:
            continue  # 跳过无效行
        try:
            # 提取标题和链接
            title_cell_idx = Config.TITLE_IN_COL if len(cols) > Config.TITLE_IN_COL else 1
            title_cell = cols[title_cell_idx]
            title_a = title_cell.find_element(By.TAG_NAME, "a")
            title = title_a.text.strip()
            href = title_a.get_attribute("href") or ""

            # 提取作者
            author = ""
            for idx in Config.AUTHOR_COL_CANDIDATES:
                if idx < len(cols) and cols[idx].text.strip():
                    author = cols[idx].text.strip()
                    break

            # 提取来源（期刊/会议）
            source = ""
            for idx in Config.SOURCE_COL_CANDIDATES:
                if idx < len(cols) and cols[idx].text.strip():
                    source = cols[idx].text.strip()
                    break

            # 提取发表时间
            pub_date = ""
            for idx in Config.DATE_COL_CANDIDATES:
                if idx < len(cols) and cols[idx].text.strip():
                    pub_date = cols[idx].text.strip()
                    break

            # 暂存基础信息，摘要后续填充
            papers.append({
                "标题": title,
                "作者": author,
                "来源": source,
                "发表时间": pub_date,
                "链接": href,
                "摘要": "",
                "爬取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            if href:
                href_list.append((len(papers) - 1, href))  # 记录索引和链接
        except Exception as e:
            log(f"提取列表信息失败：{e}")
            continue

    # 第二步：多线程并行提取摘要（核心优化）
    if href_list:
        log(f"多线程提取 {len(href_list)} 篇论文摘要（并发数：{Config.MAX_THREADS}）")
        session = create_requests_session(driver)  # 复用Cookies保持会话
        with ThreadPoolExecutor(max_workers=Config.MAX_THREADS) as executor:
            # 提交所有摘要提取任务
            futures = {
                executor.submit(fetch_abstract, session, href): idx
                for idx, href in href_list
            }
            # 收集结果并填充到papers
            for future in as_completed(futures):
                idx = futures[future]
                try:
                    abstract = future.result()
                    papers[idx]["摘要"] = abstract
                except Exception as e:
                    log(f"摘要线程异常：{e}")

    return papers


def go_next_page(driver: webdriver.Chrome) -> bool:
    """翻到下一页"""
    for i in range(Config.MAX_RETRY):
        try:
            for by, sel in Config.NEXT_SELECTORS:
                try:
                    btn = WebDriverWait(driver, Config.WAIT_TIME).until(
                        EC.element_to_be_clickable((by, sel))
                    )
                    driver.execute_script("arguments[0].click();", btn)  # 避免元素被遮挡
                    jitter(1.8, 3.2)
                    wait_for_captcha(driver)  # 处理翻页后验证码
                    return True
                except TimeoutException:
                    continue
            raise NoSuchElementException("未找到下一页按钮")
        except Exception as e:
            log(f"翻页失败（重试 {i + 1}/{Config.MAX_RETRY}）：{e}")
            snapshot(driver, tag="next_page_fail")
            backoff_sleep(i)
    return False


# ====================== 数据保存 ======================
def save_checkpoint(df: pd.DataFrame):
    """保存断点数据"""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(Config.OUTPUT_DIR, f"checkpoint_{ts}.csv")
    df.to_csv(path, index=False, encoding="utf-8-sig")
    log(f"已保存断点文件：{path}")


def save_final(df: pd.DataFrame):
    """保存最终结果"""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(Config.OUTPUT_DIR, f"cnki_results_{ts}.csv")
    df.to_csv(path, index=False, encoding="utf-8-sig")
    log(f"爬取完成！共获取 {len(df)} 篇论文，已保存至：{path}")


# ====================== 主流程 ======================
def main():
    ensure_dirs()
    log("===== 知网论文爬虫（多线程版）启动 =====")
    driver = None
    all_papers: List[Dict] = []
    try:
        # 初始化浏览器并执行搜索
        driver = init_browser()
        ok = perform_search(driver, Config.KEYWORDS)
        if not ok:
            log("搜索失败，程序退出")
            return

        # 分页爬取论文
        page_idx = 0
        while page_idx < Config.MAX_PAGES:
            log(f"处理第 {page_idx + 1} 页 ...")
            papers = extract_papers_from_table(driver)
            log(f"本页抓取 {len(papers)} 条记录")
            all_papers.extend(papers)

            # 保存快照和断点
            snapshot(driver, tag=f"page_{page_idx + 1}")
            if page_idx % Config.CKPT_EVERY_PAGES == 0 and all_papers:
                save_checkpoint(pd.DataFrame(all_papers))

            # 翻到下一页
            if not go_next_page(driver):
                log("无下一页或翻页失败，结束爬取")
                break
            page_idx += 1

    except KeyboardInterrupt:
        log("用户手动中断程序")
    except Exception as e:
        log(f"运行异常：{e}")
        traceback.print_exc()
        if driver:
            snapshot(driver, tag="fatal_error")  # 保存错误快照
    finally:
        # 关闭浏览器
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    # 保存最终结果
    if all_papers:
        # 去重（按标题）
        df = pd.DataFrame(all_papers).drop_duplicates(subset=["标题"])
        save_final(df)
    else:
        log("未获取到任何论文数据")
    log("===== 爬虫结束 =====")


if __name__ == "__main__":
    main()