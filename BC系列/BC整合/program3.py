import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from tkinter import font as tkFont
import pandas as pd
import asyncio
import re
from datetime import datetime, timedelta

# 支援的站台 (保留原過濾邏輯)
SUPPORTED_SITE_NAMES = ["CJ", "CY"]

# --- 日期選擇視窗 ---
def select_dates(parent):
    selected_data = {"start_date": "", "end_date": ""}
    window = tk.Toplevel(parent)
    window.title("選擇日期 - 幸運抽獎")
    window.geometry("400x300")
    window.grab_set()
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
        try:
            sd = start_date_entry.get()
            ed = end_date_entry.get()
            datetime.strptime(sd, '%Y/%m/%d')
            datetime.strptime(ed, '%Y/%m/%d')
            selected_data["start_date"] = sd
            selected_data["end_date"] = ed
            window.destroy()
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的日期格式 (YYYY/MM/DD)")

    tk.Button(window, text="確定", command=confirm_selection, font=font_style).pack(pady=10)
    window.wait_window()
    return selected_data

# --- 異步抓取抽獎資料 ---
async def scrape_lucky_draw_async(page):
    all_data = []
    header_added = False
    table_selector = "#userDetailTable"

    while True:
        try:
            await page.wait_for_selector(table_selector, state="visible", timeout=5000)
            if not header_added:
                headers = await page.locator(f"{table_selector} thead th").all_text_contents()
                all_data.append([h.strip() for h in headers])
                header_added = True

            rows = await page.locator(f"{table_selector} tbody tr").all()
            for row in rows:
                cols = await row.locator("td").all_text_contents()
                all_data.append([c.strip() for c in cols])

            next_button = page.locator("#userDetailTable_next")
            is_visible = await next_button.is_visible()
            class_attr = await next_button.get_attribute("class") or ""
            
            if not is_visible or "disabled" in class_attr:
                break

            await next_button.click()
            await asyncio.sleep(2)
        except Exception as e:
            print(f"抓取表格時發生錯誤：{e}")
            break
    return all_data

# --- 異步設置日期 ---
async def set_date_async(page, selector, date_str):
    try:
        await page.evaluate("""
            (args) => {
                const [sel, val] = args;
                const input = document.querySelector(sel);
                if (input) {
                    input.value = val;
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }
        """, [selector, date_str])
        return True
    except:
        return False

# --- 單一站台處理程序 ---
async def process_site_async(site, page, start_date, end_date):
    lucky_draw_url = f"{site['url'].rstrip('/')}/LuckyDrawActivity"
    for attempt in range(2):
        try:
            await page.goto(lucky_draw_url, wait_until="domcontentloaded", timeout=30000)
            if page.url.rstrip('/') == site['url'].rstrip('/'): continue
            break
        except:
            if attempt == 1: return []
            await asyncio.sleep(2)

    try:
        summary_btn = page.locator("a:has-text('领取统计')").first
        await page.wait_for_selector("a:has-text('领取统计')", timeout=10000)
        await summary_btn.click()
        await asyncio.sleep(2)
    except: return []

    if not await set_date_async(page, "#StartDate", start_date): return []
    if not await set_date_async(page, "#EndDate", end_date): return []

    try:
        search_btn = page.locator("#SummaryForm button:has-text('查询')").first
        await search_btn.click()
        await asyncio.sleep(2)
    except: return []

    return await scrape_lucky_draw_async(page)

# --- Program 3 入口函式 ---
def run_program_3(root, selected_sites, pages, cb):
    supported_sites = [s for s in selected_sites if s["name"] in SUPPORTED_SITE_NAMES]
    if not supported_sites:
        messagebox.showinfo("提示", f"未選擇任何支援的站台 ({', '.join(SUPPORTED_SITE_NAMES)})。")
        if cb: root.after(0, cb)
        return

    selected_data = select_dates(root)
    if not selected_data["start_date"] or not selected_data["end_date"]:
        if cb: root.after(0, cb)
        return

    start_date = selected_data["start_date"]
    end_date = selected_data["end_date"]
    all_results = {}

    # 定義輸出的標準順序
    SHEET_ORDER = ["TC", "TF", "TS", "SY", "FL", "WX", "XC", "XH", "CJ", "CY"]

    async def _async_process():
        try:
            tasks = []
            for site, page in zip(selected_sites, pages):
                if site["name"] not in SUPPORTED_SITE_NAMES:
                    continue
                async def worker(s, p):
                    data = await process_site_async(s, p, start_date, end_date)
                    if data: all_results[s["name"]] = data
                tasks.append(worker(site, page))

            await asyncio.gather(*tasks)

            if all_results:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx")],
                    initialfile=f"幸運抽獎_{datetime.now().strftime('%m%d_%H%M')}.xlsx"
                )
                if save_path:
                    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                        # --- 依照預設順序寫入工作表 ---
                        for site_name in SHEET_ORDER:
                            if site_name in all_results:
                                data = all_results[site_name]
                                df = pd.DataFrame(data[1:], columns=data[0])
                                df.to_excel(writer, sheet_name=site_name, index=False)
                        
                        # 如果有不在 SHEET_ORDER 裡的站台，補在後面
                        for site_name, data in all_results.items():
                            if site_name not in SHEET_ORDER:
                                df = pd.DataFrame(data[1:], columns=data[0])
                                df.to_excel(writer, sheet_name=site_name, index=False)
                                
                    messagebox.showinfo("完成", f"存檔成功：{save_path}")
            else:
                messagebox.showinfo("提示", "查無數據。")
        except Exception as e:
            messagebox.showerror("程式錯誤", str(e))
        finally:
            if cb: root.after(0, cb)

    from __main__ import app
    app.run_async(_async_process())