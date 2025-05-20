import asyncio
import tkinter as tk
from tkinter import font as tkFont, filedialog, messagebox
from playwright.async_api import async_playwright
import pandas as pd
from datetime import datetime

# 站点配置
sites = [
    {"name": "XX", "url": "XX"},
]

def get_selected_sites():
    selected_sites = []
    root = tk.Tk()  # Create Tkinter window to select sites
    root.title("選擇站台")

    # 設定視窗大小
    root.geometry("400x300")

    font_style = tkFont.Font(size=16)  # 調整字體大小

    var_list = []
    for index, site in enumerate(sites):
        var = tk.BooleanVar(value=True)  # 預設勾選
        chk = tk.Checkbutton(root, text=site["name"], variable=var, font=font_style)
        chk.grid(row=index // 4, column=index % 4, sticky='nsew', padx=5, pady=1)

        var_list.append(var)

    for i in range(4):
        root.grid_columnconfigure(i, weight=1)
    for i in range((len(sites) + 3) // 4):
        root.grid_rowconfigure(i, weight=1)

    def on_submit():
        for var, site in zip(var_list, sites):
            if var.get():
                selected_sites.append(site)
        root.quit()  # Close the site selection window

    submit_button = tk.Button(root, text="提交", command=on_submit, font=font_style)
    submit_button.grid(row=(len(sites) + 3) // 4, column=0, columnspan=4, pady=10)

    root.mainloop()  # Open the site selection window
    return selected_sites

# 儲存抓取的所有資料
all_data = []
header_added = False
page = None  # 用來儲存頁面對象

async def scrape_table_and_save(site):
    global all_data, header_added

    # 重置 header_added 為 False，這樣每個站台都會重新添加標題
    header_added = False

    while True:
        try:
            table_selector = "#eventLogTables"
            await page.wait_for_selector(table_selector, timeout=2000)

            rows = await page.locator(f"{table_selector} tbody tr").all()
            page_data = []

            for row in rows:
                columns = await row.locator("td").all_text_contents()
                page_data.append(columns)

            if not header_added:
                columns = ["时间", "事件类型", "资讯类型", "资讯描述", "状态", "功能", "操作者","详情"]
                all_data.append(columns)
                header_added = True

            all_data.extend(page_data)

            next_button = page.locator("#eventLogTables_next")
            next_button_class = await next_button.evaluate("el => el.className")

            if "disabled" in next_button_class:
                break

            await next_button.click()
            await asyncio.sleep(2)
        except Exception as e:
            messagebox.showerror("錯誤", f"抓取資料時發生錯誤：{e}")
            break

    if not all_data:
        messagebox.showinfo("提示", "未抓取到任何資料。")
        return

    file_name = f"事件紀錄_{site['name']}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=file_name)

    if save_path:
        df = pd.DataFrame(all_data)
        df.to_excel(save_path, index=False)
        messagebox.showinfo("完成", f"資料已成功保存至：{save_path}")
        all_data.clear()  # 清空 all_data
    else:
        messagebox.showinfo("取消", "未保存資料。")

async def open_website_and_scrape_for_site(site):
    global page
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        await page.goto(site["url"])
        messagebox.showinfo("操作提示", f"請確認到達 {site['name']} 的事件紀錄頁面。")

        # 在這裡設置 scraping_event，告訴程式可以繼續抓取
        scraping_event.set()

        # 等待並開始抓取
        await scrape_table_and_save(site)

async def scrape_all_sites(selected_sites):
    for site in selected_sites:
        await open_website_and_scrape_for_site(site)

def start_scraping_button():
    selected_sites = get_selected_sites()
    if not selected_sites:
        messagebox.showinfo("提示", "未選擇任何站台。")
        return

    # 依序抓取所有選中的站台
    asyncio.run(scrape_all_sites(selected_sites))

scraping_event = asyncio.Event()

# No need to create an additional Tk window here
selected_sites = get_selected_sites()  # 等待使用者選擇站台
if selected_sites:
    asyncio.run(scrape_all_sites(selected_sites))  # 開始抓取選中的站台
