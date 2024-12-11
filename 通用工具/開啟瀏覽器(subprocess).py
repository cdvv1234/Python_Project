import subprocess
import pyautogui
import time

# 啟動 Chrome 瀏覽器並開啟登入頁面
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
login_url = "https://www.example.com/login"
subprocess.Popen([chrome_path, "--incognito", login_url])

# 等待頁面加載
time.sleep(5)

# 填入帳號密碼
# 假設帳號和密碼欄位的位置如下，請根據實際情況調整座標
# 帳號欄位
pyautogui.click(x=100, y=200)  # 假設帳號欄位的座標是 (100, 200)
pyautogui.write("your_username")  # 填寫帳號

# 密碼欄位
pyautogui.click(x=100, y=250)  # 假設密碼欄位的座標是 (100, 250)
pyautogui.write("your_password")  # 填寫密碼

# 按下 Enter 鍵登入
pyautogui.press("enter")  # 假設登入按鈕的焦點在密碼欄位並可以直接按 Enter

# 等待輸入 2FA 驗證碼
input("請手動輸入驗證碼，並按 Enter 鍵繼續...")

# 登入完成後，開始打開更多分頁
# 假設要打開的網站清單
urls = [
    "https://www.example.com/dashboard",
    "https://www.example.com/settings",
    "https://www.example.com/profile"
]

# 等待一些時間，確保登入完成
time.sleep(3)

# 使用 subprocess 開啟新的分頁
for url in urls:
    subprocess.Popen([chrome_path, "--incognito", url])
    time.sleep(3)  # 等待頁面加載

# 程式結束，瀏覽器會保持開啟狀態
print("已經完成所有操作！")
