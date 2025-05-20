import tkinter as tk
from tkinter import font as tkFont
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import os

# 站台配置
sites = [
    {"type": "第一類站台", "name": "", "url": "", "handler": "type_1"},
    {"type": "第二類站台", "name": "", "url": "", "handler": "type_2"},
    {"type": "第三類站台", "name": "", "url": "", "handler": "type_3"}
]

# 通用工具
def export_to_excel_append(data, filename, sheet_name):
    """
    將數據追加保存為 Excel 文件。
    """
    df = pd.DataFrame(data).reset_index(drop=True)  # 重設索引，防止輸出索引列
    if os.path.exists(filename):
        with pd.ExcelWriter(filename, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            start_row = writer.sheets[sheet_name].max_row
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=start_row)
    else:
        with pd.ExcelWriter(filename, mode="w", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def add_columns_type_1_and_3(data, site_name, calc_date):
    """
    第一類和第三類站台：新增欄位 - 月份、上下月、平台
    """
    updated_data = []
    date_obj = calc_date
    month = date_obj.strftime("%y/%m")
    day = date_obj.day
    up_or_down = "上月" if day <= 15 else "下月"
    for row in data:
        updated_data.append([month, up_or_down, site_name] + row)
    return updated_data

def add_columns_type_2(data, calc_date):
    """
    第二類站台：新增欄位 - 計算日期、月份、天數
    """
    updated_data = []
    date_obj = calc_date
    month = date_obj.month
    day = date_obj.day
    if day <= 10:
        day_count = 10
    elif day <= 20:
        day_count = 20
    else:
        day_count = 30
    for row in data:
        updated_data.append([calc_date.strftime("%Y/%m/%d"), month, day_count] + row)
    return updated_data


# 抓取邏輯

def handle_site_type_1(page, site_name, start_date, end_date, processed_dates, base_url, original_url):
    """
    第一類站台處理邏輯，並新增欄位。
    """

    try:
        # 等待 iframe 加載
        page.wait_for_selector("iframe", timeout=10000)
        iframe = page.query_selector("iframe")
        if iframe:
            frame = iframe.content_frame()
            if not frame:
                print("Unable to access iframe.")
                return
        else:
            frame = page

        details_links = []
        valid_dates = []

        # 抓取表格資料
        data = fetch_table_data(frame, "table.table-box")
        print(f"Fetched table data (first 5 rows): {data[:5]}")  # 打印前5筆表格資料，幫助調試

        # 清空 valid_dates 和 details_links
        valid_dates.clear()
        details_links.clear()

        for row in data:
            try:
                row_start_date = datetime.strptime(row[0].strip(), "%Y/%m/%d")
            except ValueError:
                continue  # 無效的日期格式則跳過

            if start_date <= row_start_date <= end_date and row_start_date not in processed_dates:
                valid_dates.append(row_start_date)
                details_button = frame.query_selector(f"tr:has-text('{row[0]}') a:has-text('詳細')")
                if details_button:
                    details_href = details_button.get_attribute('href')
                    if details_href:
                        details_links.append((row_start_date, details_href))

        print(f"Valid dates for scraping: {valid_dates}")  # 打印篩選後的日期
        print(f"Corresponding links for scraping: {[link for _, link in details_links]}")  # 打印篩選後的超連結



        for calc_date, details_href in details_links:
            full_url = f"{original_url}{details_href}"
            print(f"Navigating to: {full_url}")  # 打印出詳細頁面的 URL
            frame.goto(full_url)
            frame.wait_for_load_state("load")  # 確保頁面已完全加載
            frame.wait_for_timeout(2000)
            details_data = fetch_table_data(frame, "table.table-box")
            updated_data = add_columns_type_1_and_3(details_data, site_name, calc_date)
            header = ["月份", "上下月", "平台", "直属帐号", "代理一级帐号", "总投注量", "总投注人數", "上半月盈亏", "总盈亏", "总分紅", "百分比", "状态"]
            if not os.path.exists(f"自動分紅_{site_name}_{datetime.now().strftime('%Y_%m_%d')}.xlsx"):
                updated_data.insert(0, header)  # 如果檔案不存在，新增標題列
            export_to_excel_append(updated_data, f"自動分紅_{site_name}_{datetime.now().strftime('%Y_%m_%d')}.xlsx", "詳情數據")
            processed_dates.add(calc_date)
    
    except Exception as e:
        print(f"An error occurred: {e}")



def handle_site_type_2(page, site_name, start_date, end_date, processed_dates, base_url, original_url):
    """
    第二類站台處理邏輯，並新增欄位。
    """
    details_links = []
    valid_dates = []
    select_show_all(page, "select[name='ActivityTable_length']")
    data = fetch_table_data(page, "table#ActivityTable")
    print(f"Fetched table data (first 5 rows): {data[:5]}")  # 打印前5筆表格資料，幫助調試

    # 清空 valid_dates 和 details_links
    valid_dates.clear()
    details_links.clear()

    for row in data:
        try:
            row_calc_date = datetime.strptime(row[0].strip(), "%Y/%m/%d")
        except ValueError:
            continue  # 無效的日期格式則跳過

        # 日期篩選及數值檢查
        if start_date <= row_calc_date <= end_date and row_calc_date not in processed_dates and row[1] != '0' and row[2] != '0':
            valid_dates.append(row_calc_date)
            
            # 使用 Playwright 獲取按鈕的 href 屬性
            details_button = page.query_selector(f"tr:has-text('{row[0]}') a:has-text('詳細')")
            if details_button:
                details_href = details_button.get_attribute('href')
                if details_href:
                    details_links.append((row_calc_date, details_href))

    print(f"Valid dates for scraping: {valid_dates}")  # 打印篩選後的日期
    print(f"Corresponding links for scraping: {[link for _, link in details_links]}")  # 打印篩選後的超連結

    # 根據篩選出的日期進行抓取
    for calc_date, details_href in details_links:
        full_url = f"{original_url}{details_href}"
        print(f"Navigating to: {full_url}")  # 打印出詳細頁面的 URL
        page.goto(full_url)
        page.wait_for_load_state("load")  # 確保頁面已完全加載
        page.wait_for_timeout(2000)

        # 新增下拉選單選擇“全部”
        select_show_all(page, "select[name='ActivityTable_length']")
        
        details_data = fetch_table_data(page, "table#ActivityTable")
        updated_data = add_columns_type_2(details_data, calc_date)

        header = ["計算日期", "月份", "天數", "帳號", "投注量", "盈虧", "奖励", "狀態", "備註"]
        excel_filename = f"自動分紅_{site_name}_{datetime.now().strftime('%Y_%m_%d')}.xlsx"
        
        if not os.path.exists(excel_filename):
            updated_data.insert(0, header)  # 如果檔案不存在，新增標題列

        export_to_excel_append(updated_data, excel_filename, "總數據")
        processed_dates.add(calc_date)


def handle_site_type_3(page, site_name, start_date, end_date, processed_dates, base_url, original_url):
    """
    第三類站台處理邏輯，並新增欄位。
    """
    details_links = []
    valid_dates = []
    # 先選擇全部資料以便抓取完整的表格
    select_show_all(page, "select[name='AutoBonusTable_length']")
    data = fetch_table_data(page, "table#AutoBonusTable")
    print(f"Fetched table data (first 5 rows): {data[:5]}")  # 打印前5筆表格資料，幫助調試

    # 清空 valid_dates 和 details_links
    valid_dates.clear()
    details_links.clear()

    for row in data:
        try:
            row_start_date = datetime.strptime(row[0].strip(), "%Y/%m/%d")
        except ValueError:
            continue  # 無效的日期格式則跳過

        if start_date <= row_start_date <= end_date and row_start_date not in processed_dates:
            valid_dates.append(row_start_date)
            details_button = page.query_selector(f"tr:has-text('{row[0]}') a:has-text('详情')")
            if details_button:
                details_href = details_button.get_attribute('href')
                if details_href:
                    details_links.append((row_start_date, details_href))

    print(f"Valid dates for scraping: {valid_dates}")  # 打印篩選後的日期
    print(f"Corresponding links for scraping: {[link for _, link in details_links]}")  # 打印篩選後的超連結

    for calc_date, details_href in details_links:
        full_url = f"{original_url}{details_href}"
        print(f"Navigating to: {full_url}")  # 打印出詳細頁面的 URL
        page.goto(full_url)
        page.wait_for_load_state("load")  # 確保頁面已完全加載
        page.wait_for_timeout(2000)

        # 新增下拉選單選擇“全部”
        select_show_all(page, "select[name='AutoBonusRecordTable_length']")

        # 抓取詳細頁面中的數據
        details_data = fetch_table_data(page, "table#AutoBonusRecordTable")
        if not details_data:
            print("No data fetched from details page.")
            continue

        updated_data = add_columns_type_1_and_3(details_data, site_name, calc_date)

        header = ["月份", "上下月", "平台", "直属帐号",  "总投注量", "总投注人數", "总盈亏", "总分紅", "百分比", "发放状态"]
        excel_filename = f"自動分紅_{site_name}_{datetime.now().strftime('%Y_%m_%d')}.xlsx"

        if not os.path.exists(excel_filename):
            updated_data.insert(0, header)  # 如果檔案不存在，新增標題列

        export_to_excel_append(updated_data, excel_filename, "詳情數據")
        processed_dates.add(calc_date)



# 抓取數據示例
def fetch_table_data(page, table_selector):
    """
    抓取表格數據，返回數據列表。
    """
    rows = page.query_selector_all(f"{table_selector} tbody tr")
    data = []
    for row in rows:
        cols = row.query_selector_all("td")
        data.append([col.inner_text() for col in cols])
    return data

def select_show_all(page, dropdown_selector):
    """
    選擇下拉框中的 "顯示全部"。
    """
    page.select_option(dropdown_selector, "-1")
    page.wait_for_timeout(2000)

# GUI 選擇站台
def get_selected_sites():
    selected_sites = []
    start_date = None
    end_date = None

    def on_submit():
        nonlocal selected_sites, start_date, end_date
        try:
            start_date = start_date_entry.get()
            end_date = end_date_entry.get()
            if start_date > end_date:
                messagebox.showerror("日期錯誤", "起始日期不能晚於結束日期！")
                return
        except ValueError:
            messagebox.showerror("日期格式錯誤", "請使用正確的日期格式：YYYY/MM/DD")
            return

        selected_sites = [site for var, site in var_list if var.get()]
        root.destroy()

    root = tk.Tk()
    root.title("選擇站台")
    root.geometry("600x600")
    font_style = tkFont.Font(size=12)

    # 日期選擇部分
    tk.Label(root, text="起始日期 (YYYY/MM/DD):", font=font_style).grid(row=0, column=0, sticky="w", pady=5)
    start_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/MM/dd', font=font_style)
    start_date_entry.grid(row=0, column=1, pady=5, padx=5)

    tk.Label(root, text="結束日期 (YYYY/MM/DD):", font=font_style).grid(row=1, column=0, sticky="w", pady=5)
    end_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/MM/dd', font=font_style)
    end_date_entry.grid(row=1, column=1, pady=5, padx=5)

    site_by_type = {}
    for site in sites:
        site_by_type.setdefault(site["type"], []).append(site)

    var_list = []
    row = 2
    for site_type, site_list in site_by_type.items():
        tk.Label(root, text=site_type, font=tkFont.Font(size=14, weight="bold")).grid(row=row, column=0, columnspan=4, sticky="w", pady=5)
        row += 1
        for index, site in enumerate(site_list):
            var = tk.BooleanVar(value=True)
            chk = tk.Checkbutton(root, text=site["name"], variable=var, font=font_style)
            chk.grid(row=row + index // 4, column=index % 4, sticky="w", pady=5, padx=15)
            var_list.append((var, site))
        row += (len(site_list) + 3) // 4  # Move to the next row after every full set of columns

    submit_button = tk.Button(root, text="確定", command=on_submit, font=font_style)
    submit_button.grid(row=row, column=0, columnspan=4, pady=10)

    root.mainloop()
    return selected_sites, start_date, end_date

# 主邏輯
def start_fetch():
    selected_sites, start_date_str, end_date_str = get_selected_sites()
    start_date = datetime.strptime(start_date_str, "%Y/%m/%d")
    end_date = datetime.strptime(end_date_str, "%Y/%m/%d")
    processed_dates = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)

        for site in selected_sites:
            context = browser.new_context()
            page = context.new_page()
            page.goto(site["url"])

            processed_dates = set()

            # 記錄當前頁面 URL 作為基礎 URL
            base_url = f"{site['url'].split('/Account')[0]}"
            if site["type"] == "第一類站台":
                base_url += "/AutoBonusManagement/AutoBonusCalculationList"
            elif site["type"] == "第二類站台":
                base_url += "/AutoSupportBonusRecord"
            elif site["type"] == "第三類站台":
                base_url += "/AutoBonusRecord"

            # 原始網址，用於拼接詳情按鈕的超連結
            original_url = f"{site['url'].split('/Account')[0]}"

            # 添加開始抓取按鈕，等待用戶手動登錄和導航到指定頁面
            def on_start_scraping():
                handler_function = globals().get(f"handle_site_{site['handler']}")
                if handler_function:
                    handler_function(page, site["name"], start_date, end_date, processed_dates, base_url, original_url)
                root.destroy()

            root = tk.Tk()
            root.title("手動登錄完成後開始抓取")
            root.geometry("300x150")
            font_style = tkFont.Font(size=12)

            start_button = tk.Button(root, text="開始抓取", command=on_start_scraping, font=font_style)
            start_button.pack(expand=True, pady=20)

            root.mainloop()
            context.close()
            
        browser.close()

if __name__ == "__main__":
    start_fetch()
