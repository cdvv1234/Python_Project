import tkinter as tk
from tkinter import font as tkFont, filedialog, messagebox
from tkcalendar import DateEntry
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime, timedelta
import time
from tkinter import ttk
import re

# 通用的文字清理函數
def clean_text(text):
    """清理文字欄位，移除多餘空行並將多個空格替換為單個空格"""
    if isinstance(text, str):
        # 移除多餘的空行和換行符
        text = re.sub(r'\n+', ' ', text)
        # 將多個空格替換為單個空格
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    return text

# 站點配置（與 main.py 一致）
sites = [
    {"name": "TC", "url": ""},
]

def select_dates(parent):
    selected_data = {"start_date": "", "end_date": ""}
    window = tk.Toplevel(parent)
    window.title("選擇日期與時間")
    window.geometry("400x350")
    font_style = tkFont.Font(size=10)

    yesterday = datetime.now() - timedelta(days=1)
    today = datetime.now()

    tk.Label(window, text="開始日期與時間 (YYYY/MM/DD HH:MM)：", font=font_style).pack(pady=5)
    start_frame = tk.Frame(window)
    start_frame.pack(pady=5)
    
    start_date_entry = DateEntry(start_frame, width=12, font=font_style, date_pattern='yyyy/mm/dd')
    start_date_entry.set_date(yesterday)
    start_date_entry.pack(side=tk.LEFT, padx=5)
    
    start_hour_combo = ttk.Combobox(start_frame, width=4, values=[f"{i:02d}" for i in range(24)], font=font_style)
    start_hour_combo.pack(side=tk.LEFT, padx=5)
    start_hour_combo.set("03")
    
    tk.Label(start_frame, text=":", font=font_style).pack(side=tk.LEFT)
    
    start_minute_combo = ttk.Combobox(start_frame, width=4, values=[f"{i:02d}" for i in range(60)], font=font_style)
    start_minute_combo.pack(side=tk.LEFT, padx=5)
    start_minute_combo.set("00")
    
    start_manual_entry = tk.Entry(start_frame, width=16, font=font_style)
    start_manual_entry.pack(side=tk.LEFT, padx=5)
    
    def update_start_manual():
        date = start_date_entry.get()
        hour = start_hour_combo.get()
        minute = start_minute_combo.get()
        start_manual_entry.delete(0, tk.END)
        start_manual_entry.insert(0, f"{date} {hour}:{minute}")
    
    start_date_entry.bind("<<DateEntrySelected>>", lambda e: update_start_manual())
    start_hour_combo.bind("<<ComboboxSelected>>", lambda e: update_start_manual())
    start_minute_combo.bind("<<ComboboxSelected>>", lambda e: update_start_manual())
 
    tk.Label(window, text="結束日期與時間 (YYYY/MM/DD HH:MM)：", font=font_style).pack(pady=5)
    end_frame = tk.Frame(window)
    end_frame.pack(pady=5)
    
    end_date_entry = DateEntry(end_frame, width=12, font=font_style, date_pattern='yyyy/mm/dd')
    end_date_entry.set_date(today)
    end_date_entry.pack(side=tk.LEFT, padx=5)
    
    end_hour_combo = ttk.Combobox(end_frame, width=4, values=[f"{i:02d}" for i in range(24)], font=font_style)
    end_hour_combo.pack(side=tk.LEFT, padx=5)
    end_hour_combo.set("03")
    
    tk.Label(end_frame, text=":", font=font_style).pack(side=tk.LEFT)
    
    end_minute_combo = ttk.Combobox(end_frame, width=4, values=[f"{i:02d}" for i in range(60)], font=font_style)
    end_minute_combo.pack(side=tk.LEFT, padx=5)
    end_minute_combo.set("00")

    end_manual_entry = tk.Entry(end_frame, width=16, font=font_style)
    end_manual_entry.pack(side=tk.LEFT, padx=5)
    
    def update_end_manual():
        date = end_date_entry.get()
        hour = end_hour_combo.get()
        minute = end_minute_combo.get()
        end_manual_entry.delete(0, tk.END)
        end_manual_entry.insert(0, f"{date} {hour}:{minute}")
    
    end_date_entry.bind("<<DateEntrySelected>>", lambda e: update_end_manual())
    end_hour_combo.bind("<<ComboboxSelected>>", lambda e: update_end_manual())
    end_minute_combo.bind("<<ComboboxSelected>>", lambda e: update_end_manual())

    def confirm_selection():
        selected_data["start_date"] = start_manual_entry.get()
        selected_data["end_date"] = end_manual_entry.get()
        
        try:
            datetime.strptime(selected_data["start_date"], '%Y/%m/%d %H:%M')
            datetime.strptime(selected_data["end_date"], '%Y/%m/%d %H:%M')
            window.destroy()
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的日期格式 (YYYY/MM/DD HH:MM)")

    submit_button = tk.Button(window, text="確定", command=confirm_selection, font=font_style)
    submit_button.pack(pady=10)

    update_start_manual()
    update_end_manual()

    window.wait_window()
    return selected_data

def scrape_table(site, page):
    all_data = []
    header_added = False

    while True:
        try:
            table_selector = "#eventLogTables"
            page.wait_for_selector(table_selector, state="visible", timeout=2000)
            rows = page.locator(f"{table_selector} tbody tr").all()
            page_data = []

            for row in rows:
                columns = row.locator("td").all_text_contents()
                # 清理每一列的文字
                cleaned_columns = [clean_text(col) for col in columns]
                page_data.append(cleaned_columns)

            if not header_added:
                if site["name"].startswith("TC") or site["name"].startswith("TF"):
                    columns = ["时间", "事件类型", "资讯描述", "状态", "功能", "操作者", "详情"]
                    all_data.append(columns)
                    header_added = True
                else:
                    columns = ["时间", "事件类型", "资讯类型", "资讯描述", "状态", "功能", "操作者", "详情"]
                    all_data.append(columns)
                    header_added = True

            all_data.extend(page_data)

            next_button = page.locator("#eventLogTables_next > a")
            parent_li = page.locator("#eventLogTables_next")
            if (parent_li.count() and "disabled" in parent_li.get_attribute("class", timeout=5000)) or \
               next_button.get_attribute("disabled", timeout=5000) or not next_button.is_visible():
                break

            next_button.click()
            time.sleep(2)
        except Exception as e:
            print(f"抓取資料時發生錯誤：{e}")
            break

    return all_data

def save_results(all_data_dict):
    if not all_data_dict:
        messagebox.showinfo("提示", "未抓取到任何資料。")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        initialfile=f"事件紀錄_{timestamp}.xlsx"
    )

    if save_path:
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            for site_name, data in all_data_dict.items():
                if data:
                    df = pd.DataFrame(data[1:], columns=data[0])
                    df = df.applymap(clean_text)  # 再次確保資料乾淨
                    df.to_excel(writer, sheet_name=site_name, index=False)
        messagebox.showinfo("完成", f"資料已保存至 {save_path}")

def process_site(site, page, start_date, end_date):
    is_tc_tf = site["name"] in ["TC", "TF"]
    event_log_url = f"{site['url'].rstrip('/')}/EventLog"

    max_retries = 2
    for attempt in range(max_retries):
        try:
            page.goto(event_log_url, wait_until="domcontentloaded", timeout=30000)
            final_url = page.url
            if final_url.rstrip('/') == site["url"].rstrip('/'):
                print(f"站台 {site['name']} 的 EventLog 重定向到首頁: {final_url}")
                page.screenshot(path=f"error_{site['name']}_eventlog_navigation.png")
                return []
            break
        except Exception as e:
            print(f"站台 {site['name']} 導航到 {event_log_url} 失敗 (嘗試 {attempt + 1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                page.screenshot(path=f"error_{site['name']}_eventlog_navigation.png")
                return []
            time.sleep(2)

    try:
        page.wait_for_selector("#StartTime", state="visible", timeout=10000)
        if is_tc_tf:
            page.type("#StartTime", start_date)
            page.type("#EndTime", end_date)
            print(f"站台 {site['name']} StartTime={start_date}, EndTime={end_date}")
        else:
            page.fill("#StartTime", start_date)
            page.fill("#EndTime", end_date)
            print(f"站台 {site['name']} StartTime={start_date}, EndTime={end_date}")
    except Exception as e:
        print(f"站台 {site['name']} 填入日期失敗: {e}")
        page.screenshot(path=f"error_{site['name']}_eventlog_input.png")
        return []

    search_selector = (
        "#KoEventLog > div.row.clearfix > div > div > div > div.panel-collapse.collapse.in > div > div > form > div.form-actions > button"
        if is_tc_tf else
        "#KoEventLog > div.card.shadow-sm.mb-4 > div.card-body > form > div.form-group.row > div > button"
    )

    try:
        page.wait_for_selector(search_selector, state="visible", timeout=20000)
        page.click(search_selector)
        time.sleep(2)
    except Exception as e:
        print(f"站台 {site['name']} 點擊查詢按鈕失敗: {e}")
        page.screenshot(path=f"error_{site['name']}_eventlog_search.png")
        return []

    return scrape_table(site, page)

def run_program_1(selected_sites, pages):
    if not selected_sites:
        messagebox.showinfo("提示", "未選擇任何站台。")
        return

    root = tk.Tk()
    root.withdraw()
    selected_data = select_dates(root)
    root.destroy()

    if not selected_data["start_date"] or not selected_data["end_date"]:
        print("未輸入日期，程式結束。")
        return

    start_date = selected_data["start_date"]
    end_date = selected_data["end_date"]

    all_data = {}
    for site, page in zip(selected_sites, pages):
        site_data = process_site(site, page, start_date, end_date)
        if site_data:
            all_data[site["name"]] = site_data

    save_results(all_data)
    print("run_program_1 執行完畢")
