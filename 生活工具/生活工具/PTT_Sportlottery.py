import requests
import time
import json
import os
from bs4 import BeautifulSoup
from datetime import datetime

# Configuration
TARGET_AUTHORS = ["apparition10", "lotterywin", "bvbin10242"]
MIN_COMMENTS = 50
CHECK_INTERVAL = 120  # 2 minutes in seconds (can be adjusted to your preference)
MAX_CHECK_INTERVAL = 1800  # 30 minutes
BASE_URL = "https://www.ptt.cc"
BOARD_URL = f"{BASE_URL}/bbs/SportLottery/index.html"
DATA_FILE = "tracked_posts.json"

# LINE Messaging API 設定
LINE_CHANNEL_ID = ""
LINE_CHANNEL_SECRET = ""
LINE_CHANNEL_ACCESS_TOKEN = ""  # 你需要填寫這個值
LINE_USER_ID = ""  # 你需要填寫接收訊息的使用者ID

# Initialize tracked posts
def load_tracked_posts():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"posts": []}

def save_tracked_posts(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# Function to get cookies (PTT age verification)
def get_ptt_cookies():
    response = requests.get(BASE_URL)
    cookies = response.cookies
    if 'over18' not in cookies:
        cookies.set('over18', '1')  # Set over18 cookie to bypass age verification
    return cookies

# Function to send LINE message using Messaging API
def send_line_message(message):
    if not LINE_CHANNEL_ACCESS_TOKEN or not LINE_USER_ID:
        print("LINE Channel Access Token or User ID not set. Skipping notification.")
        return
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"
    }
    
    # 將訊息拆分成多個部分，確保不超過LINE的訊息長度限制
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
            print("LINE Message sent successfully")
        else:
            print(f"Failed to send LINE Message: {response.status_code} {response.text}")
            print(f"Response: {response.json()}")

# Function to fetch and parse a page
def fetch_page(url, cookies):
    try:
        response = requests.get(url, cookies=cookies)
        response.raise_for_status()
        return BeautifulSoup(response.text, 'html.parser')
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return None

# Function to extract post information
def extract_post_info(post_element):
    try:
        # Extract the title and remove "[公告]" tags
        title_element = post_element.select_one('.title a')
        if not title_element:
            return None
        
        title = title_element.text.strip()
        link = title_element.get('href')
        
        # Skip LIVE posts
        if "[LIVE]" in title:
            return None
        
        # Extract author
        meta_elements = post_element.select('.meta .author')
        author = meta_elements[0].text.strip() if meta_elements else "Unknown"
        
        # Extract comment count from nrec element
        nrec = post_element.select_one('.nrec')
        comment_count = 0
        if nrec and nrec.text:
            if nrec.text == '爆':
                comment_count = 100  # "爆" usually means 100+ comments
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
        print(f"Error extracting post info: {e}")
        return None

# Function to check if a post meets our criteria
def post_meets_criteria(post):
    if not post:
        return False
    
    # Check if the author is one we're tracking
    author_match = post["author"] in TARGET_AUTHORS
    
    # Check if the post has enough comments and is not a LIVE post
    comment_match = post["comment_count"] >= MIN_COMMENTS and "[LIVE]" not in post["title"]
    
    return author_match or comment_match

# Function to scan multiple pages
def scan_pages(cookies, num_pages=3):
    new_posts = []
    current_url = BOARD_URL
    
    for _ in range(num_pages):
        soup = fetch_page(current_url, cookies)
        if not soup:
            continue
        
        # Get all posts on the page
        posts = soup.select('.r-ent')
        for post_element in posts:
            post_info = extract_post_info(post_element)
            if post_info and post_meets_criteria(post_info):
                new_posts.append(post_info)
        
        # Get the link to the previous page
        prev_link = soup.select_one('.btn.wide:nth-of-type(2)')
        if prev_link and 'href' in prev_link.attrs:
            current_url = BASE_URL + prev_link['href']
        else:
            break
    
    return new_posts

# Main function
def main():
    print(f"PTT SportLottery Tracker started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Tracking authors: {', '.join(TARGET_AUTHORS)}")
    print(f"Tracking posts with {MIN_COMMENTS}+ comments (excluding LIVE posts)")
    print(f"Check interval: {CHECK_INTERVAL}-{MAX_CHECK_INTERVAL} seconds")
    
    tracked_data = load_tracked_posts()
    tracked_urls = [post["link"] for post in tracked_data["posts"]]
    cookies = get_ptt_cookies()
    
    try:
        while True:
            print(f"\nChecking for new posts at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}...")
            
            # Scan the first few pages
            new_posts = scan_pages(cookies)
            
            # Filter out posts we've already tracked
            truly_new_posts = [post for post in new_posts if post["link"] not in tracked_urls]
            
            if truly_new_posts:
                # Update our tracking data
                tracked_data["posts"].extend(truly_new_posts)
                tracked_urls.extend([post["link"] for post in truly_new_posts])
                save_tracked_posts(tracked_data)
                
                # Prepare notification
                notification = "\n🏆 新的 PTT SportLottery 文章 🏆\n\n"
                for post in truly_new_posts:
                    notification += f"📌 {post['title']}\n"
                    notification += f"👤 {post['author']}\n"
                    notification += f"💬 {post['comment_count']} 推文\n"
                    notification += f"🔗 {post['link']}\n\n"
                
                print(notification)
                send_line_message(notification)
                print(f"Found {len(truly_new_posts)} new posts")
            else:
                print("No new posts found")
            
            # Wait for the next check
            time.sleep(CHECK_INTERVAL)
    
    except KeyboardInterrupt:
        print("\nTracker stopped by user")
    except Exception as e:
        print(f"Error in main loop: {e}")
        send_line_message(f"⚠️ PTT Tracker 錯誤: {str(e)}")

if __name__ == "__main__":
    main()
