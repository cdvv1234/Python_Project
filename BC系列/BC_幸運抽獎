import asyncio
import tkinter as tk
from tkinter import filedialog, messagebox
from playwright.async_api import async_playwright
import pandas as pd
from datetime import datetime

# 儲存抓取的所有資料
all_data = []
# 記錄標題是否已經加入過
header_added = False
page = None  # 用來儲存頁面對象

async def scrape_table_and_save():
    global all_data, header_added

    # 開始抓取資料
    while True:
        try:
            # 確認表格是否存在
            table_selector = "#userDetailTable"
            await page.wait_for_selector(table_selector, timeout=2000)  # 等待表格出現

            # 抓取表格資料
            rows = await page.locator(f"{table_selector} tbody tr").all()
            page_data = []

            for row in rows:
                columns = await row.locator("td").all_text_contents()
                page_data.append(columns)

            # 如果是第一次抓取，加入標題
            if not header_added:
                columns = ["用户名", "日期", "异动时间", "抽奖次数", "亏损金额", "领取金额", "兑换状态", "功能", "操作者"]
                all_data.append(columns)  # 只在第一次抓取時添加標題
                header_added = True

            # 添加當前頁面的資料
            all_data.extend(page_data)

            # 確認是否有下一頁按鈕並可點擊
            next_button = page.locator("#userDetailTable_next")
            next_button_class = await next_button.evaluate("el => el.className")  # 使用 evaluate 來取得 class

            # 檢查按鈕是否禁用（class 包含 "disabled"）
            if "disabled" in next_button_class:
                break  # 如果是最後一頁，退出抓取
            
            # 點擊下一頁按鈕
            await next_button.click()

            # 等待頁面加載
            await asyncio.sleep(2)
        except Exception as e:
            messagebox.showerror("錯誤", f"抓取資料時發生錯誤：{e}")
            break

    # 如果沒有抓到資料
    if not all_data:
        messagebox.showinfo("提示", "未抓取到任何資料。")
        return

    # 將資料存為 DataFrame 並儲存為 Excel
    df = pd.DataFrame(all_data)
    file_name = f"幸運抽獎_{datetime.now().strftime('%Y%m%d')}.xlsx"

    # 提示使用者選擇保存位置
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=file_name)
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("完成", f"資料已成功保存至：{save_path}")
    else:
        messagebox.showinfo("取消", "未保存資料。")

async def open_website_and_scrape():
    global page

    async with async_playwright() as p:
        # 啟動瀏覽器
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # 自動導航到指定網站
        target_url = ""  # 替換為目標網站 URL
        await page.goto(target_url)

        # 提示使用者登入
        messagebox.showinfo("操作提示", "請確認到達幸運抽獎頁面。")

        # 等待使用者按下 "開始抓取" 按鈕後繼續抓取資料
        await scraping_event.wait()  # 等待抓取事件觸發

        # 開始抓取資料
        await scrape_table_and_save()

# 用來控制是否可以開始抓取資料
scraping_event = asyncio.Event()

def start_scraping_button():
    # 設置抓取事件為 True，並開始運行網站開啟和抓取的協程
    scraping_event.set()
    
    # 使用 root.after 來觸發 asyncio.run() 運行協程，避免衝突
    root.after(0, asyncio.run, open_website_and_scrape())  # 在 Tkinter 事件循環中運行協程

# Tkinter 介面
root = tk.Tk()
root.title("網頁表格抓取工具")
root.geometry("400x300")

# 開啟目標網站並開始抓取的按鈕
open_button = tk.Button(root, text="開啟網站並開始抓取", command=start_scraping_button, width=20, height=2)
open_button.pack(pady=20)

root.mainloop()
