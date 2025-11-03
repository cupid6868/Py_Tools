# -*- coding: utf-8 -*-
"""
CNKI 知网爬虫（合规人机协同版）
- 入口：kns8s 新版搜索页
- 反爬优化：随机指纹、UA、窗口、行为模拟、指数退避、代理可选、会话复用（Cookies）
- 验证码：检测后人工处理（控制台按回车继续），不包含破解逻辑
- 健壮性：重试、快照（HTML/Screenshot）、断点续爬
"""

import os
import time
import random
import pickle
import traceback
from datetime import datetime
from typing import List, Dict, Tuple

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# ====================== 配置 ======================
class Config:
    KEYWORDS = "减污降碳"
    MAX_PAGES = 5
    BASE_URL = "https://www.cnki.net/"
    OUTPUT_DIR = "cnki_papers"
    WAIT_TIME = 15
    HEADLESS = False

    MIN_DELAY = 1.0
    MAX_DELAY = 2.8
    TYPE_MIN = 0.05
    TYPE_MAX = 0.15

    CKPT_EVERY_PAGES = 1
    COOKIE_PATH = os.path.join(OUTPUT_DIR, "session_cookies.pkl")

    SNAPSHOT_DIR = os.path.join(OUTPUT_DIR, "snapshots")
    LOG_PATH = os.path.join(OUTPUT_DIR, "run.log")

    USE_PROXY = False
    PROXY_POOL = []

    UA_MIN = 118
    UA_MAX = 124
    MAX_RETRY = 3

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
    os.makedirs(Config.OUTPUT_DIR, exist_ok=True)
    os.makedirs(Config.SNAPSHOT_DIR, exist_ok=True)

def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(Config.LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def jitter(min_s=None, max_s=None):
    a = Config.MIN_DELAY if min_s is None else min_s
    b = Config.MAX_DELAY if max_s is None else max_s
    time.sleep(random.uniform(a, b))

def choose_proxy() -> str:
    if Config.USE_PROXY and Config.PROXY_POOL:
        return random.choice(Config.PROXY_POOL)
    return ""

def snapshot(driver: webdriver.Chrome, tag: str = "snapshot"):
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
    base = 1.5
    t = (base ** retry_idx) + random.random()
    log(f"指数退避等待 {t:.2f}s ...")
    time.sleep(t)

def human_typing(el, text: str):
    for ch in text:
        el.send_keys(ch)
        time.sleep(random.uniform(Config.TYPE_MIN, Config.TYPE_MAX))

def human_scroll(driver: webdriver.Chrome):
    times = random.randint(2, 5)
    for _ in range(times):
        y = random.randint(300, 900)
        driver.execute_script(f"window.scrollBy(0, {y});")
        time.sleep(random.uniform(0.4, 1.2))
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(random.uniform(0.2, 0.6))

def human_mouse_wiggle(driver: webdriver.Chrome):
    try:
        actions = ActionChains(driver)
        for _ in range(random.randint(2, 5)):
            x_off = random.randint(-30, 30)
            y_off = random.randint(-20, 20)
            actions.move_by_offset(x_off, y_off).pause(random.uniform(0.1, 0.3))
        actions.perform()
    except Exception:
        pass

# ====================== 浏览器初始化 ======================
def init_browser() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    if Config.HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
    width = random.randint(1280, 1920)
    height = random.randint(800, 1080)
    options.add_argument(f"--window-size={width},{height}")

    ua_ver = random.randint(Config.UA_MIN, Config.UA_MAX)
    ua = f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{ua_ver}.0.0.0 Safari/537.36"
    options.add_argument(f"user-agent={ua}")
    options.add_argument("--lang=zh-CN,zh;q=0.9")

    proxy = choose_proxy()
    if proxy:
        options.add_argument(f"--proxy-server={proxy}")
        log(f"使用代理：{proxy}")

    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-popup-blocking")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": r"""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(navigator, 'languages', {get: () => ['zh-CN', 'zh']});
        Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});
        """
    })
    return driver

# ====================== Cookies 会话 ======================
def save_cookies(driver: webdriver.Chrome):
    try:
        cookies = driver.get_cookies()
        with open(Config.COOKIE_PATH, "wb") as f:
            pickle.dump(cookies, f)
        log(f"Cookies 已保存：{Config.COOKIE_PATH}")
    except Exception as e:
        log(f"保存 Cookies 失败：{e}")

def load_cookies(driver: webdriver.Chrome, url_for_scope: str):
    try:
        if os.path.exists(Config.COOKIE_PATH):
            with open(Config.COOKIE_PATH, "rb") as f:
                cookies = pickle.load(f)
            for ck in cookies:
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
            for by, sel in search_box_candidates:
                try:
                    box = WebDriverWait(driver, Config.WAIT_TIME).until(EC.element_to_be_clickable((by, sel)))
                    break
                except TimeoutException:
                    continue
            for by, sel in search_btn_candidates:
                try:
                    btn = WebDriverWait(driver, Config.WAIT_TIME).until(EC.element_to_be_clickable((by, sel)))
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
    for i in range(Config.MAX_RETRY):
        try:
            smart_get(driver, Config.BASE_URL)
            wait_for_captcha(driver)
            try:
                load_cookies(driver, Config.BASE_URL)
                driver.refresh()
                jitter()
            except Exception:
                pass
            box, btn = get_search_elements(driver)
            box.click()
            time.sleep(random.uniform(0.2, 0.5))
            box.clear()
            human_typing(box, keywords)
            jitter()
            human_mouse_wiggle(driver)
            if btn:
                btn.click()
            else:
                box.send_keys(Keys.ENTER)
            jitter(2.0, 3.5)
            wait_for_captcha(driver)
            save_cookies(driver)
            return True
        except Exception as e:
            log(f"执行搜索失败：{e}")
            snapshot(driver, tag="perform_search_fail")
            backoff_sleep(i)
    return False

def find_results_table(driver: webdriver.Chrome):
    for i in range(Config.MAX_RETRY):
        try:
            for by, sel in Config.TABLE_SELECTORS:
                try:
                    tbl = WebDriverWait(driver, Config.WAIT_TIME).until(EC.presence_of_element_located((by, sel)))
                    return tbl
                except TimeoutException:
                    continue
            raise NoSuchElementException("未找到结果列表")
        except Exception as e:
            log(f"等待结果表失败：{e}")
            backoff_sleep(i)
    return None

def extract_papers_from_table(driver: webdriver.Chrome) -> List[Dict]:
    tbl = find_results_table(driver)
    if tbl is None:
        return []
    human_scroll(driver)
    rows = tbl.find_elements(By.TAG_NAME, "tr")
    papers = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) < 3:
            continue
        try:
            title_cell_idx = Config.TITLE_IN_COL if len(cols) > Config.TITLE_IN_COL else 1
            title_cell = cols[title_cell_idx]
            title_a = title_cell.find_element(By.TAG_NAME, "a")
            title = title_a.text.strip()
            href = title_a.get_attribute("href") or ""

            author = ""
            for idx in Config.AUTHOR_COL_CANDIDATES:
                if idx < len(cols):
                    t = cols[idx].text.strip()
                    if t:
                        author = t
                        break

            source = ""
            for idx in Config.SOURCE_COL_CANDIDATES:
                if idx < len(cols):
                    t = cols[idx].text.strip()
                    if t:
                        source = t
                        break

            pub_date = ""
            for idx in Config.DATE_COL_CANDIDATES:
                if idx < len(cols):
                    t = cols[idx].text.strip()
                    if t:
                        pub_date = t
                        break

            # ========== 新增摘要抓取 ==========
            abstract = ""
            if href:
                try:
                    driver.execute_script("window.open('');")  # 新标签打开论文详情页
                    driver.switch_to.window(driver.window_handles[-1])
                    smart_get(driver, href)
                    wait_for_captcha(driver)
                    # 尝试抓取摘要文本
                    try:
                        abs_elem = driver.find_element(By.CSS_SELECTOR, "span.abstract-text, div.abstract-text, p.abstract-text")
                        abstract = abs_elem.text.strip()
                    except NoSuchElementException:
                        abstract = ""
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    jitter(0.5, 1.2)
                except Exception:
                    abstract = ""
                    if len(driver.window_handles) > 1:
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])

            papers.append({
                "标题": title,
                "作者": author,
                "来源": source,
                "发表时间": pub_date,
                "链接": href,
                "摘要": abstract,
                "爬取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
        except Exception:
            continue
    return papers


def go_next_page(driver: webdriver.Chrome) -> bool:
    for i in range(Config.MAX_RETRY):
        try:
            for by, sel in Config.NEXT_SELECTORS:
                try:
                    btn = WebDriverWait(driver, Config.WAIT_TIME).until(EC.element_to_be_clickable((by, sel)))
                    driver.execute_script("arguments[0].click();", btn)
                    jitter(1.8, 3.2)
                    wait_for_captcha(driver)
                    return True
                except TimeoutException:
                    continue
            raise NoSuchElementException("未找到下一页按钮")
        except Exception as e:
            log(f"翻页失败（重试 {i+1}/{Config.MAX_RETRY}）：{e}")
            snapshot(driver, tag="next_page_fail")
            backoff_sleep(i)
    return False

# ====================== 保存 ======================
def save_checkpoint(df: pd.DataFrame):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(Config.OUTPUT_DIR, f"checkpoint_{ts}.csv")
    df.to_csv(path, index=False, encoding="utf-8-sig")
    log(f"已保存断点文件：{path}")

def save_final(df: pd.DataFrame):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(Config.OUTPUT_DIR, f"cnki_results_{ts}.csv")
    df.to_csv(path, index=False, encoding="utf-8-sig")
    log(f"爬取完成！共获取 {len(df)} 篇论文，已保存至：{path}")

# ====================== 主流程 ======================
def main():
    ensure_dirs()
    log("===== 知网论文爬虫启动 =====")
    driver = None
    all_papers: List[Dict] = []
    try:
        driver = init_browser()
        ok = perform_search(driver, Config.KEYWORDS)
        if not ok:
            log("搜索失败，程序退出")
            return

        page_idx = 0
        while page_idx < Config.MAX_PAGES:
            log(f"处理第 {page_idx+1} 页 ...")
            papers = extract_papers_from_table(driver)
            log(f"本页抓取 {len(papers)} 条")
            all_papers.extend(papers)

            snapshot(driver, tag=f"page_{page_idx+1}")

            if page_idx % Config.CKPT_EVERY_PAGES == 0 and all_papers:
                save_checkpoint(pd.DataFrame(all_papers))

            if not go_next_page(driver):
                log("无下一页或翻页失败，结束")
                break
            page_idx += 1

    except KeyboardInterrupt:
        log("用户中断")
    except Exception as e:
        log(f"运行异常：{e}")
        traceback.print_exc()
        if driver:
            snapshot(driver, tag="fatal")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    if all_papers:
        df = pd.DataFrame(all_papers).drop_duplicates(subset=["标题"])
        save_final(df)
    else:
        log("未获取到任何论文")
    log("===== 爬虫结束 =====")

if __name__ == "__main__":
    main()
