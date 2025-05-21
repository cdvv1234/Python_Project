import requests
from urllib.parse import urlparse
from fake_useragent import UserAgent

def check_robots_txt(url):
    """Check if a URL is allowed by robots.txt."""
    try:
        # 隨機 User-Agent
        ua = UserAgent()
        headers = {"User-Agent": ua.random}
        
        # 構建 robots.txt URL
        parsed_url = urlparse(url)
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
        robots_url = f"{base_url}/robots.txt"
        
        # 發送請求
        response = requests.get(robots_url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # 檢查路徑是否被禁止
        path = parsed_url.path or "/"
        for line in response.text.splitlines():
            line = line.strip()
            if line.startswith("Disallow:"):
                disallow_path = line.split(":", 1)[1].strip()
                if disallow_path and path.startswith(disallow_path):
                    print(f"Warning: {url} is disallowed by robots.txt")
                    return False
        
        print(f"{url} is allowed by robots.txt")
        return True
    
    except Exception as e:
        print(f"Error checking robots.txt for {url}: {e}")
        return True  # 若無 robots.txt 或請求失敗，假設允許