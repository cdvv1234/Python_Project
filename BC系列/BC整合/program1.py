import tkinter as tk
from tkinter import font as tkFont, filedialog, messagebox
from tkcalendar import DateEntry
from tkinter import ttk
import pandas as pd
import asyncio
import re
from datetime import datetime, timedelta

# --- 通用的文字清理函數 ---
def clean_text(text):
    if isinstance(text, str):
        text = re.sub(r'\n+', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    return text

# --- 日期選擇視窗 ---
def select_dates(parent):
    selected_data = {"start_date": "", "end_date": ""}
    window = tk.Toplevel(parent)
    window.title("選擇日期與時間 - 事件紀錄")
    window.geometry("400x350")
    window.grab_set()
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
    
    def update_start_manual(*args):
        date = start_date_entry.get()
        hour = start_hour_combo.get()
        minute = start_minute_combo.get()
        start_manual_entry.delete(0, tk.END)
        start_manual_entry.insert(0, f"{date} {hour}:{minute}")
    
    start_date_entry.bind("<<DateEntrySelected>>", update_start_manual)
    start_hour_combo.bind("<<ComboboxSelected>>", update_start_manual)
    start_minute_combo.bind("<<ComboboxSelected>>", update_start_manual)

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
    
    def update_end_manual(*args):
        date = end_date_entry.get()
        hour = end_hour_combo.get()
        minute = end_minute_combo.get()
        end_manual_entry.delete(0, tk.END)
        end_manual_entry.insert(0, f"{date} {hour}:{minute}")
    
    end_date_entry.bind("<<DateEntrySelected>>", update_end_manual)
    end_hour_combo.bind("<<ComboboxSelected>>", update_end_manual)
    end_minute_combo.bind("<<ComboboxSelected>>", update_end_manual)

    def confirm_selection():
        try:
            datetime.strptime(start_manual_entry.get(), '%Y/%m/%d %H:%M')
            datetime.strptime(end_manual_entry.get(), '%Y/%m/%d %H:%M')
            selected_data["start_date"] = start_manual_entry.get()
            selected_data["end_date"] = end_manual_entry.get()
            window.destroy()
        except ValueError:
            messagebox.showerror("錯誤", "日期格式不正確")

    tk.Button(window, text="確定", command=confirm_selection, font=font_style).pack(pady=10)
    update_start_manual()
    update_end_manual()
    window.wait_window()
    return selected_data

# --- 異步抓取表格 ---
async def scrape_table_async(site, page):
    all_data = []
    header_added = False
    table_selector = "#eventLogTables"

    while True:
        try:
            await page.wait_for_selector(table_selector, state="visible", timeout=5000)
            rows_data = await page.evaluate("""(sel) => {
                const rows = Array.from(document.querySelectorAll(sel + ' tbody tr'));
                return rows.map(tr => Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim()));
            }""", table_selector)

            if not rows_data or "表中数据为空" in (rows_data[0][0] if rows_data else ""):
                break

            if not header_added:
                if site["name"].startswith("TC") or site["name"].startswith("TF"):
                    columns = ["时间", "事件类型", "资讯描述", "状态", "功能", "操作者", "详情"]
                else:
                    columns = ["时间", "事件类型", "资讯類型", "资讯描述", "状态", "功能", "操作者", "详情"]
                all_data.append(columns)
                header_added = True

            for row in rows_data:
                all_data.append([clean_text(col) for col in row])

            next_button = page.locator("#eventLogTables_next > a")
            parent_li = page.locator("#eventLogTables_next")
            is_visible = await next_button.is_visible()
            class_attr = await parent_li.get_attribute("class") or ""
            
            if not is_visible or "disabled" in class_attr:
                break

            await next_button.click()
            await asyncio.sleep(2)
        except Exception as e:
            print(f"[{site['name']}] 翻頁時出錯: {e}")
            break
    return all_data

# --- 單一站台處理程序 ---
async def process_site_async(site, page, start_date, end_date):
    is_tc_tf = site["name"] in ["TC", "TF"]
    event_log_url = f"{site['url'].rstrip('/')}/EventLog"
    for attempt in range(2):
        try:
            await page.goto(event_log_url, wait_until="domcontentloaded", timeout=30000)
            if page.url.rstrip('/') == site["url"].rstrip('/'):
                await asyncio.sleep(1)
                continue
            break
        except Exception as e:
            if attempt == 1: return []
            await asyncio.sleep(2)

    try:
        await page.wait_for_selector("#StartTime", state="visible", timeout=10000)
        if is_tc_tf:
            await page.fill("#StartTime", "")
            await page.type("#StartTime", start_date)
            await page.fill("#EndTime", "")
            await page.type("#EndTime", end_date)
        else:
            await page.fill("#StartTime", start_date)
            await page.fill("#EndTime", end_date)
    except:
        return []

    search_selector = "button:has-text('查询')" if is_tc_tf else "#KoEventLog button:has-text('查询'), #KoEventLog button:has-text('Search')"
    try:
        btn = page.locator(search_selector).first
        await btn.click()
        await asyncio.sleep(2)
    except:
        return []

    return await scrape_table_async(site, page)

# --- Program 1 入口函式 ---
def run_program_1(root, selected_sites, pages, cb):
    if not selected_sites:
        messagebox.showinfo("提示", "未選擇任何站台。")
        if cb: root.after(0, cb)
        return

    selected_data = select_dates(root)
    if not selected_data["start_date"] or not selected_data["end_date"]:
        if cb: root.after(0, cb)
        return

    start_date = selected_data["start_date"]
    end_date = selected_data["end_date"]
    all_data_dict = {}

    # 定義輸出的工作表順序
    SHEET_ORDER = ["TC", "TF", "TS", "SY", "FL", "WX", "XC", "XH", "CJ", "CY"]

    async def _async_process():
        try:
            tasks = []
            for site, page in zip(selected_sites, pages):
                async def worker(s, p):
                    res = await process_site_async(s, p, start_date, end_date)
                    if res: all_data_dict[s["name"]] = res
                tasks.append(worker(site, page))
            
            await asyncio.gather(*tasks)

            if all_data_dict:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx")],
                    initialfile=f"事件紀錄_{datetime.now().strftime('%m%d_%H%M')}.xlsx"
                )
                if save_path:
                    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                        # --- 重要調整：依照預設順序寫入工作表 ---
                        for site_name in SHEET_ORDER:
                            if site_name in all_data_dict:
                                data = all_data_dict[site_name]
                                df = pd.DataFrame(data[1:], columns=data[0])
                                df.to_excel(writer, sheet_name=site_name, index=False)
                        
                        # 如果有不在 SHEET_ORDER 裡的站台（預防萬一），最後補上
                        for site_name, data in all_data_dict.items():
                            if site_name not in SHEET_ORDER:
                                df = pd.DataFrame(data[1:], columns=data[0])
                                df.to_excel(writer, sheet_name=site_name, index=False)
                                
                    messagebox.showinfo("完成", f"存檔成功：{save_path}")
            else:
                messagebox.showinfo("提示", "查無任何數據")
        except Exception as e:
            messagebox.showerror("錯誤", f"抓取中斷: {e}")
        finally:
            if cb: root.after(0, cb)

    from __main__ import app
    app.run_async(_async_process())