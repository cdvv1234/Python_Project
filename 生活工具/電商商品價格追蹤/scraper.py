import time
from playwright.sync_api import sync_playwright
import random
import logging
import os
from fake_useragent import UserAgent
from config import (
    PRODUCT_URL, DISCOUNTED_PRICE_SELECTOR, TITLE_SELECTOR,
    BACKUP_DISCOUNTED_PRICE_SELECTOR, BACKUP_TITLE_SELECTOR,
    USER_AGENT, WAIT_TIMEOUT
)
from robots_checker import check_robots_txt

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[
    logging.FileHandler('scraper.log', encoding='utf-8'),
    logging.StreamHandler()
])
logger = logging.getLogger(__name__)

def save_debug_files(page, timestamp, suffix=""):
    """保存頁面 HTML 和截圖"""
    try:
        debug_file = f"debug_page_{timestamp}{suffix}.html"
        screenshot_file = f"debug_screenshot_{timestamp}{suffix}.png"
        with open(debug_file, "w", encoding="utf-8") as f:
            f.write(page.content())
        page.screenshot(path=screenshot_file)
        logger.info(f"Saved page HTML to {debug_file} and screenshot to {screenshot_file}")
    except Exception as e:
        logger.error(f"Failed to save debug files: {e}")

def scrape_product_info():
    """
    使用 Playwright 爬取商品名稱、折扣後價格和折扣前價格
    返回：(title, discounted_price, original_price) 或 (None, None, None) 如果失敗或 robots.txt 不允許
    """
    # 檢查 robots.txt
    if not check_robots_txt(PRODUCT_URL):
        logger.warning(f"Skipping scrape: {PRODUCT_URL} is disallowed by robots.txt")
        return None, None, None

    max_retries = 2
    for attempt in range(max_retries):
        try:
            with sync_playwright() as p:
                # 隨機 User-Agent
                ua = UserAgent()
                user_agent = ua.random if random.random() > 0.5 else USER_AGENT
                logger.info(f"Attempt {attempt + 1}/{max_retries} - Using User-Agent: {user_agent}")
                
                # 啟動瀏覽器
                browser = p.chromium.launch(headless=True)
                context = browser.new_context(
                    user_agent=user_agent,
                    extra_http_headers={
                        "Accept-Language": "zh-TW,zh;q=0.9",
                        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
                        "Referer": "https://24h.pchome.com.tw/",
                        "DNT": "1",
                        "Connection": "keep-alive",
                        "Upgrade-Insecure-Requests": "1",
                        "Accept-Encoding": "gzip, deflate, br",
                        "Cache-Control": "no-cache"
                    },
                    viewport={"width": 1280, "height": 720}
                )
                page = context.new_page()
                
                # 訪問商品頁面
                logger.info(f"Navigating to {PRODUCT_URL}")
                response = page.goto(PRODUCT_URL, wait_until="networkidle")
                if response:
                    logger.info(f"HTTP Status: {response.status}")
                    if response.status != 200:
                        logger.warning(f"Non-200 status code: {response.status}")
                        raise Exception(f"Failed to load page, status: {response.status}")
                
                # 等待頁面完全加載
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(40000)  # 40 秒，應對動態加載
                
                # 保存初始頁面
                save_debug_files(page, int(time.time()), "_initial")
                
                # 檢查頁面編碼和標題
                content_type = page.evaluate("document.contentType")
                charset = page.evaluate("document.characterSet")
                page_title = page.title()
                logger.info(f"Page content type: {content_type}, charset: {charset}, title: {page_title}")
                
                # 驗證商品標題
                if "IRIS OHYAMA" not in page_title and "超靜音極細保密碎紙機" not in page_title:
                    logger.warning(f"Unexpected product page: {page_title}")
                    raise Exception(f"Wrong product page loaded: {page_title}")
                
                # 提取價格（折扣價和原價都在同一個元素中）
                discounted_price = None
                original_price = None
                for selector in [DISCOUNTED_PRICE_SELECTOR, BACKUP_DISCOUNTED_PRICE_SELECTOR]:
                    logger.info(f"Waiting for price selector: {selector}")
                    try:
                        page.wait_for_selector(selector, timeout=WAIT_TIMEOUT)
                        element = page.query_selector(selector)
                        if element:
                            # 使用 JavaScript 提取價格文字
                            text = page.evaluate('element => element.textContent', element).replace("NT$", "").replace(",", "").strip()
                            logger.info(f"Raw price text: {text}")
                            # 根據 $ 符號分割價格
                            price_parts = text.split('$')[1:]  # 跳過第一個空字符串（如果有）
                            prices = []
                            for part in price_parts:
                                # 移除非數字字符並轉換為整數
                                cleaned_part = ''.join(filter(str.isdigit, part))
                                if cleaned_part:
                                    prices.append(int(cleaned_part))
                            logger.info(f"Parsed prices: {prices}")
                            if prices:
                                discounted_price = prices[0]
                                logger.info(f"Extracted discounted price: {discounted_price}")
                                if len(prices) > 1:
                                    original_price = prices[1]
                                    logger.info(f"Extracted original price: {original_price}")
                                else:
                                    logger.info("No original price found in price selector")
                                break
                    except Exception as e:
                        logger.warning(f"Selector {selector} failed: {e}")
                if not discounted_price:
                    logger.error("Failed to find discounted price")
                
                # 提取商品名稱
                title = None
                for selector in [TITLE_SELECTOR, BACKUP_TITLE_SELECTOR]:
                    logger.info(f"Waiting for title selector: {selector}")
                    try:
                        page.wait_for_selector(selector, timeout=WAIT_TIMEOUT)
                        element = page.query_selector(selector)
                        if element:
                            # 確保提取的標題使用正確編碼
                            title = page.evaluate('element => element.textContent', element).strip()
                            logger.info(f"Extracted title: {title}")
                            break
                    except Exception as e:
                        logger.warning(f"Selector {selector} failed: {e}")
                if not title:
                    logger.error("Failed to find title")
                
                # 驗證數據
                if not discounted_price or not title:
                    raise Exception("Missing required data: discounted_price or title")
                
                return title, discounted_price, original_price
        except Exception as e:
            logger.error(f"Attempt {attempt + 1}/{max_retries} failed: {e}")
            save_debug_files(page, int(time.time()), "_error") if 'page' in locals() and page and not page.is_closed() else logger.info("No page to save debug files")
    logger.error(f"All {max_retries} attempts failed")
    return None, None, None