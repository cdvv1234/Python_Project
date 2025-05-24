import requests
import time
import json
import os
from bs4 import BeautifulSoup
from datetime import datetime
import threading
import pystray
from PIL import Image
import tkinter as tk
from tkinter import messagebox

# 配置文件路徑
CONFIG_FILE = "config.json"
DATA_FILE = "tracked_posts.json"

# 預設配置
DEFAULT_CONFIG = {
    "target_authors": ["apparition10", "lotterywin", "bvbin10242","qbb741000","binbinyolee"],
    "min_comments": 50,
    "check_interval": 600
}

# LINE Messaging API 配置,
LINE_CHANNEL_ID = ""
LINE_CHANNEL_SECRET = ""
LINE_CHANNEL_ACCESS_TOKEN = ""
LINE_USER_ID = ""

# 全局配置變數
config = {}

# 初始化執行緒控制
running = threading.Event()
running.set()
exit_event = threading.Event()

# 載入或初始化配置文件
def load_config():
    global config
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
    else:
        config = DEFAULT_CONFIG.copy()
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    return config

# 保存配置
def save_config():
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

# 修改配置（使用單一視窗，多行輸入框，按鈕同一行）
def modify_config(icon):
    root = tk.Tk()
    root.title("修改配置")
    root.geometry("400x300")
    root.resizable(False, False)

    try:
        # 追蹤作者（多行輸入框）
        tk.Label(root, text="追蹤作者（每行一個或用逗號分隔）：").pack(pady=5)
        authors_text = tk.Text(root, height=5, width=40)
        authors_text.pack(pady=5)
        authors_text.insert(tk.END, "\n".join(config["target_authors"]))

        # 最低推文數
        tk.Label(root, text="最低推文數：").pack(pady=5)
        comments_entry = tk.Entry(root, width=40)
        comments_entry.pack(pady=5)
        comments_entry.insert(0, str(config["min_comments"]))

        # 檢查間隔
        tk.Label(root, text="檢查間隔（秒）：").pack(pady=5)
        interval_entry = tk.Entry(root, width=40)
        interval_entry.pack(pady=5)
        interval_entry.insert(0, str(config["check_interval"]))

        # 按鈕框架（確認和取消按鈕同一行）
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)

        def save_changes():
            try:
                authors_input = authors_text.get("1.0", tk.END).strip()
                authors = [a.strip() for a in authors_input.replace(",", "\n").split("\n") if a.strip()]
                if not authors:
                    raise ValueError("請至少輸入一個作者")
                config["target_authors"] = authors

                min_comments = comments_entry.get().strip()
                if not min_comments.isdigit() or int(min_comments) < 0:
                    raise ValueError("推文數必須是非負整數")
                config["min_comments"] = int(min_comments)

                check_interval = interval_entry.get().strip()
                if not check_interval.isdigit() or int(check_interval) < 1:
                    raise ValueError("檢查間隔必須是大於0的整數")
                config["check_interval"] = int(check_interval)

                save_config()
                print(f"配置已更新: {config}")
                messagebox.showinfo("成功", "配置已更新！")
                root.destroy()
            except ValueError as e:
                messagebox.showerror("錯誤", str(e))

        tk.Button(button_frame, text="確認", command=save_changes).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=root.destroy).pack(side=tk.LEFT, padx=10)

        root.mainloop()

    except Exception as e:
        messagebox.showerror("錯誤", f"修改配置時發生錯誤: {e}")
        root.destroy()

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
    response = requests.get("https://www.ptt.cc")
    cookies = response.cookies
    if 'over18' not in cookies:
        cookies.set('over18', '1')
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
    
    message_chunks = [message[i:i+2000] for i in range(0, len(message), 2000)]
    
    for chunk in message_chunks:
        payload = {
            "to": LINE_USER_ID,
            "messages": [{"type": "text", "text": chunk}]
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
        
        if any(tag in title for tag in ["[LIVE]", "[活動]", "[公告]", "[實況]"]):
            return None
        
        meta_elements = post_element.select('.meta .author')
        author = meta_elements[0].text.strip() if meta_elements else "Unknown"
        
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
            "link": "https://www.ptt.cc" + link if link else None,
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
    
    author_match = post["author"] in config["target_authors"]
    comment_match = post["comment_count"] >= config["min_comments"] and "[LIVE]" not in post["title"]
    
    return author_match or comment_match

# 掃描多個頁面
def scan_pages(cookies, num_pages=3):
    new_posts = []
    current_url = "https://www.ptt.cc/bbs/SportLottery/index.html"
    
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
            current_url = "https://www.ptt.cc" + prev_link['href']
        else:
            break
    
    return new_posts

# 主迴圈
def main_loop():
    load_config()
    print(f"PTT SportLottery 追蹤器啟動於 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"追蹤作者: {', '.join(config['target_authors'])}")
    print(f"追蹤推文數大於等於 {config['min_comments']} 的文章（排除LIVE文章）")
    print(f"檢查間隔: {config['check_interval']} 秒")
    
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
        
        for _ in range(config["check_interval"]):
            if exit_event.is_set():
                break
            time.sleep(1)

# 系統托盤設置
def create_tray_icon():
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        image = Image.open(icon_path)
    except Exception as e:
        print(f"無法加載圖標: {e}，使用預設圖標")
        image = Image.new('RGB', (64, 64), color='blue')
    
    def on_pause(icon, item):
        running.clear()
        print("程式已暫停")
    
    def on_resume(icon, item):
        running.set()
        print("程式已繼續")
    
    def on_exit(icon, item):
        print("結束程式")
        exit_event.set()
        icon.stop()
    
    def on_modify_config(icon, item):
        modify_config(icon)
    
    menu = (
        pystray.MenuItem("繼續執行", on_resume, enabled=lambda item: not running.is_set()),
        pystray.MenuItem("暫停", on_pause, enabled=lambda item: running.is_set()),
        pystray.MenuItem("修改配置", on_modify_config),
        pystray.MenuItem("結束程式", on_exit)
    )
    
    icon = pystray.Icon("PTT Tracker", image, "PTT SportLottery Tracker", menu)
    return icon

# 主程式
def main():
    main_thread = threading.Thread(target=main_loop, daemon=True)
    main_thread.start()
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
