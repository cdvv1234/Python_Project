import asyncio
import json
from tkinter import Tk, Frame, Checkbutton, IntVar, Button, messagebox
from threading import Thread
from playwright.async_api import async_playwright
import sys
import os

# 從 JSON 載入站台資訊
def load_sites():
    # 獲取程式執行的目錄（無論是 .py 還是 .exe）
    if getattr(sys, 'frozen', False):  # PyInstaller 打包後
        base_path = os.path.dirname(sys.executable)  # .exe 所在目錄
    else:  # 開發環境
        base_path = os.path.dirname(os.path.abspath(__file__))  # 腳本所在目錄

    json_path = os.path.join(base_path, "sites.json")

    # 檢查檔案是否存在
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"找不到 {json_path} 檔案，請確保 JSON 檔案與程式在同一目錄下。")

    # 讀取 JSON 檔案內容
    with open(json_path, "r", encoding="utf-8") as file:
        return json.load(file)

# 使用 Playwright 登入選中的站台，並在單一瀏覽器中開啟不同的分頁
async def login_to_sites(selected_sites, browser):
    tasks = []

    async def login(site, context):
        try:
            page = await context.new_page()  # 在同一個上下文中創建新分頁
            await page.goto(site["url"])

            # 填寫帳號與密碼
            await page.fill(site["username_selector"], site["credentials"]["username"])
            await page.fill(site["password_selector"], site["credentials"]["password"])

            # 選擇語言
            await page.click(site["language_dropdown_selector"])
            await page.click(site["language_option_selector"])

            # 動態選擇按鈕名稱
            login_button_text = site["language_to_login_button"][site["selected_language"]]
            login_button_selector = f"button.v-btn span:has-text('{login_button_text}')"

            # 等待並點擊登入按鈕
            await page.wait_for_selector(login_button_selector, state="visible")
            await page.click(login_button_selector)

            print(f"Logged into {site['name']}")
        except Exception as e:
            print(f"Failed to log into {site['name']}: {e}")

    # 使用相同的上下文對每個選中的站台執行登入
    context = await browser.new_context()  # 在同一個瀏覽器中創建一個上下文
    for site in selected_sites:
        tasks.append(login(site, context))

    await asyncio.gather(*tasks)
    print("All selected sites are logged in.")


# 在 GUI 按鈕觸發時使用 threading 執行非同步操作
def on_login():
    Thread(target=lambda: asyncio.run(perform_login())).start()


# 處理登入邏輯
async def perform_login():
    selected_sites = [
        sites[i] for i, var in enumerate(vars) if var.get() == 1
    ]
    if not selected_sites:
        messagebox.showwarning("警告", "請選擇至少一個網站進行登入！")
        return

    # 手動啟動 Playwright，並確保瀏覽器保持開啟
    playwright = await async_playwright().start()
    browser = await playwright.chromium.launch(headless=False)  # 開啟顯示瀏覽器
    await login_to_sites(selected_sites, browser)  # 使用同一個瀏覽器實例登入所有站台
    print("All selected sites are logged in. Please manually close the browser when done.")

    # 保持 Playwright 進程運行直到手動關閉瀏覽器
    await monitor_browser(browser)  # 監控瀏覽器是否關閉，並在關閉後清理

# 監控瀏覽器是否關閉
async def monitor_browser(browser):
    try:
        # 監控 browser 關閉事件
        await browser.contexts[0].page.close()  # 確保有活動頁面
        await browser.close()  # 關閉瀏覽器
    except Exception as e:
        print(f"Error while monitoring browser: {e}")
    finally:
        print("Browser closed. Cleaning up.")

# 建立 GUI 介面
def create_gui():
    global sites, vars
    sites = load_sites()

    # GUI 介面
    root = Tk()
    root.title("選擇網站進行登入")

    # 建立核選方塊
    frame = Frame(root)
    frame.pack(pady=10)
    vars = []
    for i, site in enumerate(sites):
        var = IntVar()
        vars.append(var)
        chk = Checkbutton(frame, text=site["name"], variable=var)
        chk.grid(row=i // 6, column=i % 6, padx=5, pady=5)

    # 登入按鈕
    login_button = Button(root, text="登入選中的網站", command=on_login)
    login_button.pack(pady=10)

    # 結束按鈕
    exit_button = Button(root, text="結束程式", command=root.quit)
    exit_button.pack(pady=10)

    # 在 tkinter 窗口關閉時清理
    def on_exit():
        asyncio.run(cleanup())
        root.quit()

    root.protocol("WM_DELETE_WINDOW", on_exit)

    root.mainloop()


# 清理資源，關閉 Playwright 進程
async def cleanup():
    print("Cleaning up and closing the browser.")
    # 根據需求，在這裡可以手動釋放資源或進行其他清理操作


# 啟動程式
if __name__ == "__main__":
    create_gui()
