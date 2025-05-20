import requests
import time
import json
import os
from bs4 import BeautifulSoup
from datetime import datetime
import threading
import pystray
from PIL import Image
import sys

# 配置
TARGET_AUTHORS = ["apparition10", "lotterywin", "bvbin10242"]
MIN_COMMENTS = 50
CHECK_INTERVAL = 120  # 2分鐘（秒）
MAX_CHECK_INTERVAL = 1800  # 30分鐘（秒）
BASE_URL = "https://www.ptt.cc"
BOARD_URL = f"{BASE_URL}/bbs/SportLottery/index.html"
DATA_FILE = "tracked_posts.json"

# LINE Messaging API 配置
LINE_CHANNEL_ID = ""
LINE_CHANNEL_SECRET = ""
LINE_CHANNEL_ACCESS_TOKEN = ""
LINE_USER_ID = ""

# 初始化執行緒控制
running = threading.Event()
running.set()  # 預設為運行狀態
exit_event = threading.Event()  # 用於結束程式

# 初始化追蹤文章
def load_tracked_posts():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"posts": []}

def save_tracked_posts(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# 獲取PTT Cookie（年齡驗證）
def get_ptt_cookies():
    response = requests.get(BASE_URL)
    cookies = response.cookies
    if 'over18' not in cookies:
        cookies.set('over18', '1')  # 繞過年齡限制
    return cookies

# 發送LINE訊息
def send_line_message(message):
    if not LINE_CHANNEL_ACCESS_TOKEN or not LINE_USER_ID:
        print("未設置LINE Channel Access Token或User ID，跳過通知。")
        return
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"
    }
    
    # 拆分訊息以避免超過LINE長度限制
    message_chunks = [message[i:i+2000] for i in range(0, len(message), 2000)]
    
    for chunk in message_chunks:
        payload = {
            "to": LINE_USER_ID,
            "messages": [
                {
                    "type": "text",
                    "text": chunk
                }
            ]
        }
        
        response = requests.post(
            "https://api.line.me/v2/bot/message/push",
            headers=headers,
            json=payload
        )
        
        if response.status_code == 200:
            print("LINE訊息發送成功")
        else:
            print(f"LINE訊息發送失敗: {response.status_code} {response.text}")

# 獲取並解析頁面
def fetch_page(url, cookies):
    try:
        response = requests.get(url, cookies=cookies)
        response.raise_for_status()
        return BeautifulSoup(response.text, 'html.parser')
    except Exception as e:
        print(f"獲取頁面 {url} 時錯誤: {e}")
        return None

# 提取文章資訊
def extract_post_info(post_element):
    try:
        title_element = post_element.select_one('.title a')
        if not title_element:
            return None
        
        title = title_element.text.strip()
        link = title_element.get('href')
        
        # 排除LIVE、活動和公告文章
        if any(tag in title for tag in ["[LIVE]", "[活動]", "[公告]"]):
            return None
        
        # 提取作者
        meta_elements = post_element.select('.meta .author')
        author = meta_elements[0].text.strip() if meta_elements else "Unknown"
        
        # 提取推文數
        nrec = post_element.select_one('.nrec')
        comment_count = 0
        if nrec and nrec.text:
            if nrec.text == '爆':
                comment_count = 100
            elif nrec.text.isdigit():
                comment_count = int(nrec.text)
        
        return {
            "title": title,
            "author": author,
            "link": BASE_URL + link if link else None,
            "comment_count": comment_count,
            "discovered_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    except Exception as e:
        print(f"提取文章資訊時錯誤: {e}")
        return None

# 檢查文章是否符合條件
def post_meets_criteria(post):
    if not post:
        return False
    
    author_match = post["author"] in TARGET_AUTHORS
    comment_match = post["comment_count"] >= MIN_COMMENTS and "[LIVE]" not in post["title"]
    
    return author_match or comment_match

# 掃描多個頁面
def scan_pages(cookies, num_pages=3):
    new_posts = []
    current_url = BOARD_URL
    
    for _ in range(num_pages):
        soup = fetch_page(current_url, cookies)
        if not soup:
            continue
        
        posts = soup.select('.r-ent')
        for post_element in posts:
            post_info = extract_post_info(post_element)
            if post_info and post_meets_criteria(post_info):
                new_posts.append(post_info)
        
        prev_link = soup.select_one('.btn.wide:nth-of-type(2)')
        if prev_link and 'href' in prev_link.attrs:
            current_url = BASE_URL + prev_link['href']
        else:
            break
    
    return new_posts

# 主迴圈（在獨立執行緒中運行）
def main_loop():
    print(f"PTT SportLottery 追蹤器啟動於 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"追蹤作者: {', '.join(TARGET_AUTHORS)}")
    print(f"追蹤推文數大於等於 {MIN_COMMENTS} 的文章（排除LIVE文章）")
    print(f"檢查間隔: {CHECK_INTERVAL}-{MAX_CHECK_INTERVAL} 秒")
    
    tracked_data = load_tracked_posts()
    tracked_urls = [post["link"] for post in tracked_data["posts"]]
    cookies = get_ptt_cookies()
    
    while not exit_event.is_set():
        if running.is_set():
            print(f"\n於 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 檢查新文章...")
            
            new_posts = scan_pages(cookies)
            truly_new_posts = [post for post in new_posts if post["link"] not in tracked_urls]
            
            if truly_new_posts:
                tracked_data["posts"].extend(truly_new_posts)
                tracked_urls.extend([post["link"] for post in truly_new_posts])
                save_tracked_posts(tracked_data)
                
                notification = "\n🏆 新的 PTT SportLottery 文章 🏆\n\n"
                for post in truly_new_posts:
                    notification += f"📌 {post['title']}\n"
                    notification += f"👤 {post['author']}\n"
                    notification += f"💬 {post['comment_count']} 推文\n"
                    notification += f"🔗 {post['link']}\n\n"
                
                print(notification)
                send_line_message(notification)
                print(f"找到 {len(truly_new_posts)} 篇新文章")
            else:
                print("未找到新文章")
        
        # 等待下一次檢查或檢查退出信號
        for _ in range(CHECK_INTERVAL):
            if exit_event.is_set():
                break
            time.sleep(1)

# 系統托盤設置
def create_tray_icon():
    # 加載圖標
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        image = Image.open(icon_path)
    except Exception as e:
        print(f"無法加載圖標: {e}，使用預設圖標")
        image = Image.new('RGB', (64, 64), color='blue')  # 預設藍色方塊
    
    # 定義托盤選單
    def on_pause(icon, item):
        running.clear()
        print("程式已暫停")
    
    def on_resume(icon, item):
        running.set()
        print("程式已繼續")
    
    def on_exit(icon, item):
        print("結束程式")
        exit_event.set()  # 設置退出信號
        icon.stop()  # 停止托盤圖標
    
    menu = (
        pystray.MenuItem("繼續執行", on_resume, enabled=lambda item: not running.is_set()),
        pystray.MenuItem("暫停", on_pause, enabled=lambda item: running.is_set()),
        pystray.MenuItem("結束程式", on_exit)
    )
    
    # 創建托盤圖標
    icon = pystray.Icon("PTT Tracker", image, "PTT SportLottery Tracker", menu)
    return icon

# 主程式
def main():
    # 啟動主迴圈執行緒
    main_thread = threading.Thread(target=main_loop, daemon=True)
    main_thread.start()
    
    # 創建並運行系統托盤
    tray_icon = create_tray_icon()
    tray_icon.run()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("程式被使用者中斷")
        exit_event.set()
    except Exception as e:
        print(f"主程式錯誤: {e}")
        send_line_message(f"⚠️ PTT Tracker 錯誤: {str(e)}")
        exit_event.set()
