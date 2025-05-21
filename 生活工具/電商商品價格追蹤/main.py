import schedule
import time
import signal
import sys
import logging
from scraper import scrape_product_info
from data_handler import save_to_csv
from visualizer import plot_price_trend
from config import SCHEDULE_TIME

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[
    logging.FileHandler('main.log', encoding='utf-8'),
    logging.StreamHandler()
])
logger = logging.getLogger(__name__)

def job():
    """
    執行價格追蹤任務
    """
    logger.info("Running price tracker...")
    title, discounted_price, original_price = scrape_product_info()
    if title and discounted_price:
        save_to_csv(title, discounted_price, original_price)
    else:
        logger.warning("Failed to scrape product info.")

def signal_handler(sig, frame):
    """
    處理中斷信號，優雅退出
    """
    logger.info("Interrupt received, shutting down gracefully...")
    sys.exit(0)

def main():
    # 註冊信號處理
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    # 執行一次爬蟲並儲存
    job()
    
    # 繪製價格趨勢圖
    plot_price_trend()
    
    # 設定定時任務
    schedule.every(30).seconds.do(job)  # 測試用，每 10 秒
    # schedule.every().day.at(SCHEDULE_TIME).do(job)  # 正式用，每天 9:00
    
    logger.info(f"Price tracker started. Scheduled to run daily at {SCHEDULE_TIME}. Press Ctrl+C to stop.")
    while True:
        try:
            schedule.run_pending()
            time.sleep(1)
        except KeyboardInterrupt:
            logger.info("KeyboardInterrupt caught, shutting down...")
            sys.exit(0)

if __name__ == "__main__":
    main()