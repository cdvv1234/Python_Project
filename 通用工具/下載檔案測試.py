import tkinter as tk
from tkinter import filedialog
from playwright.sync_api import sync_playwright

def select_save_location(filename):
    """彈出儲存檔案的路徑選擇器，讓使用者選擇儲存位置"""
    root = tk.Tk()
    root.withdraw()  # 隱藏主視窗
    root.attributes('-topmost', True)  # 讓視窗置頂

    # 讓使用者選擇儲存檔案位置，不設置預設副檔名
    file_path = filedialog.asksaveasfilename(initialfile=filename, title="選擇儲存檔案的位置")
    
    root.lift()  # 確保選擇框顯示在最上層
    root.attributes('-topmost', True)  # 讓選擇框保持置頂

    if not file_path:
        print("未選擇儲存位置，取消下載")
        return None
    return file_path

def monitor_and_download():
    with sync_playwright() as p:
        # 啟動瀏覽器
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)  # 設定為接受下載
        page = context.new_page()

        # 打開目標網站
        page.goto("https://sample-videos.com/download-sample-pdf.php")

        # 監控網頁上的下載請求
        def on_download(download):
            print(f"檔案下載: {download.url}")
            # 取得檔案的名稱，從 URL 或下載標頭中推測
            filename = download.suggested_filename or download.url.split("/")[-1]

            # 嘗試根據 URL 來推測副檔名（如果沒有 suggested_filename）
            if '.' not in filename:
                # 如果 filename 沒有副檔名，推測 URL 類型（這裡我們假設是二進位檔案）
                filename = filename + '.bin'

            # 選擇儲存位置
            file_path = select_save_location(filename)
            if file_path:
                download.save_as(file_path)
                print(f"檔案已下載並儲存於: {file_path}")

        # 註冊下載事件
        page.on("download", on_download)

        print("請手動點擊下載按鈕，或等待檔案下載請求被捕捉...")
        
        # 等待用戶操作，並捕獲下載
        page.wait_for_timeout(20000)  # 等待 20 秒鐘進行操作，捕捉下載請求

        # 讓用戶有足夠時間進行下載測試
        print("請繼續進行下載測試，瀏覽器將保持開啟。")
        page.wait_for_timeout(20000)  # 進一步的等待時間，讓使用者進行多次測試

monitor_and_download()
