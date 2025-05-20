import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from tkinter import font as tkFont
import pandas as pd
from datetime import datetime, timedelta
import time

# 支援的站台
sites = [
    {"name": "CJ", "url": ""}
]

def select_dates(parent):
    selected_data = {"start_date": "", "end_date": ""}
    window = tk.Toplevel(parent)
    window.title("選擇日期")
    window.geometry("400x300")
    font_style = tkFont.Font(size=10)

    day_before_yesterday = datetime.now() - timedelta(days=2)

    tk.Label(window, text="開始日期 (YYYY/MM/DD)：", font=font_style).pack(pady=5)
    start_date_entry = DateEntry(window, width=12, font=font_style, date_pattern='yyyy/mm/dd')
    start_date_entry.set_date(day_before_yesterday)
    start_date_entry.pack(pady=5)

    tk.Label(window, text="結束日期 (YYYY/MM/DD)：", font=font_style).pack(pady=5)
    end_date_entry = DateEntry(window, width=12, font=font_style, date_pattern='yyyy/mm/dd')
    end_date_entry.set_date(day_before_yesterday)
    end_date_entry.pack(pady=5)

    def confirm_selection():
        selected_data["start_date"] = start_date_entry.get()
        selected_data["end_date"] = end_date_entry.get()
        
        try:
            datetime.strptime(selected_data["start_date"], '%Y/%m/%d')
            datetime.strptime(selected_data["end_date"], '%Y/%m/%d')
            window.destroy()
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的日期格式 (YYYY/MM/DD)")

    submit_button = tk.Button(window, text="確定", command=confirm_selection, font=font_style)
    submit_button.pack(pady=10)

    window.wait_window()
    return selected_data

def scrape_lucky_draw(page):
    all_data = []
    header_added = False

    while True:
        try:
            table_selector = "#userDetailTable"
            page.wait_for_selector(table_selector, state="visible", timeout=2000)
            rows = page.locator(f"{table_selector} tbody tr").all()
            page_data = []

            for row in rows:
                columns = row.locator("td").all_text_contents()
                page_data.append(columns)

            if not header_added:
                columns = page.locator(f"{table_selector} thead th").all_text_contents()
                all_data.append(columns)
                header_added = True

            all_data.extend(page_data)

            next_button = page.locator("#userDetailTable_next")
            next_button_class = next_button.get_attribute("class")
            if "disabled" in next_button_class:
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
        initialfile=f"幸運抽獎_{timestamp}.xlsx"
    )

    if save_path:
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            for site_name, data in all_data_dict.items():
                if data:
                    df = pd.DataFrame(data[1:], columns=data[0])
                    df.to_excel(writer, sheet_name=site_name, index=False)
        messagebox.showinfo("完成", f"資料已保存至 {save_path}")

def set_date(page, date_field_selector, date_str):
    try:
        # 使用 evaluate，將 selector 和 date_str 包裝成列表傳遞
        page.evaluate("""
            (args) => {
                const selector = args[0];
                const value = args[1];
                const input = document.querySelector(selector);
                if (input) {
                    input.value = value;  // 設置日期值
                    const event = new Event('change', { bubbles: true });  // 觸發變更事件
                    input.dispatchEvent(event);
                }
            }
        """, [date_field_selector, date_str])  # 把兩個參數打包成列表傳遞

        # 驗證日期是否正確設置
        selected_date = page.input_value(date_field_selector)
        expected_date = date_str
        if selected_date != expected_date:
            print(f"日期設置失敗，期望 {expected_date}，實際 {selected_date}")
            page.screenshot(path=f"error_{date_field_selector.replace('#', '')}_date_selection.png")
            return False
        return True
    except Exception as e:
        print(f"設置日期失敗: {e}")
        page.screenshot(path=f"error_{date_field_selector.replace('#', '')}_date_selection.png")
        return False


def process_site(site, page, start_date, end_date):
    lucky_draw_url = f"{site['url'].rstrip('/')}/LuckyDrawActivity"

    max_retries = 2
    for attempt in range(max_retries):
        try:
            page.goto(lucky_draw_url, wait_until="domcontentloaded", timeout=30000)
            final_url = page.url
            if final_url.rstrip('/') == site["url"].rstrip('/'):
                print(f"站台 {site['name']} 的 LuckyDrawActivity 重定向到首頁: {final_url}")
                page.screenshot(path=f"error_{site['name']}_luckydraw_navigation.png")
                return []
            break
        except Exception as e:
            print(f"站台 {site['name']} 導航到 {lucky_draw_url} 失敗 (嘗試 {attempt + 1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                page.screenshot(path=f"error_{site['name']}_luckydraw_navigation.png")
                return []
            time.sleep(2)

    try:
        summary_button_selector = "#content > div.container-fluid > div > div.card-header.py-3 > div > div:nth-child(2) > a"
        page.wait_for_selector(summary_button_selector, state="visible", timeout=10000)
        page.click(summary_button_selector)
        print(f"站台 {site['name']} 已點擊「领取统计」按鈕")
        time.sleep(2)
    except Exception as e:
        print(f"站台 {site['name']} 點擊「领取统计」按鈕失敗: {e}")
        page.screenshot(path=f"error_{site['name']}_luckydraw_summary_button.png")
        return []

    # 設置起始日期
    if not set_date(page, "#StartDate", start_date):
        return []

    # 設置結束日期
    if not set_date(page, "#EndDate", end_date):
        return []

    try:
        search_button_selector = "#SummaryForm > div.form-group.row > div > button"
        page.wait_for_selector(search_button_selector, state="visible", timeout=20000)
        page.click(search_button_selector)
        time.sleep(2)
    except Exception as e:
        print(f"站台 {site['name']} 點擊查詢按鈕失敗: {e}")
        page.screenshot(path=f"error_{site['name']}_luckydraw_search.png")
        return []

    return scrape_lucky_draw(page)

def run_program_3(selected_sites, pages):
    supported_sites = [site for site in selected_sites if site["name"] in [s["name"] for s in sites]]
    if not supported_sites:
        messagebox.showinfo("提示", "未選擇任何支援的站台（CJ）。")
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
        if site["name"] not in [s["name"] for s in sites]:
            continue
        site_data = process_site(site, page, start_date, end_date)
        if site_data:
            all_data[site["name"]] = site_data

    save_results(all_data)
    print("run_program_3 執行完畢")
