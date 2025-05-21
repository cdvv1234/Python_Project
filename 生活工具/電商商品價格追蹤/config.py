# 商品 URL（PChome IRIS OHYAMA 超靜音極細保密碎紙機 P6-HCS）
PRODUCT_URL = "https://24h.pchome.com.tw/prod/DCBAK8-A900F4UZV"

# 定時設置（每天上午 9 點）
SCHEDULE_TIME = "09:00"

# CSV 文件路徑
CSV_FILE = "price_history.csv"

# 選擇器（根據 PChome 商品頁面）
DISCOUNTED_PRICE_SELECTOR = "div[class*='prodPrice']"
TITLE_SELECTOR = "#ProdBriefing > div > div > div > div.c-boxGrid__item.c-boxGrid__item--prodBriefingArea > div.o-prodMainName.o-prodMainName--prodNick"

# 備用選擇器（基於 HTML 結構）
BACKUP_DISCOUNTED_PRICE_SELECTOR = "div.c-blockCombine--priceGray div.o-prodPrice__price.o-prodPrice__price--xxl700Primary, div.o-prodPrice__price--xxl700Primary"
BACKUP_TITLE_SELECTOR = "h1.o-prodMainName__grayDarkest"

# User-Agent（模擬真實瀏覽器）
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"

# 增加等待時間（毫秒）
WAIT_TIMEOUT = 40000  # 40 秒