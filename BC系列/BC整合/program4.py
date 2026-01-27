import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import font as tkFont
from tkinter import ttk
from tkcalendar import DateEntry
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

# --- 日期選擇視窗 (Synchronous UI) ---
def select_dates(parent):
    selected_data = {"start_date": "", "end_date": ""}
    window = tk.Toplevel(parent)
    window.title("選擇日期 - 審單DATA")
    window.geometry("400x300")
    window.grab_set()  # 鎖定視窗，強制使用者操作
    font_style = tkFont.Font(size=10)

    today = datetime.now()
    day_before_yesterday = today - timedelta(days=2)
    yesterday = today - timedelta(days=1)

    tk.Label(window, text="開始日期與時間 (YYYY/MM/DD HH:MM)：", font=font_style).pack(pady=5)
    start_frame = tk.Frame(window)
    start_frame.pack(pady=5)
    
    start_date_entry = DateEntry(start_frame, width=12, font=font_style, date_pattern='yyyy/mm/dd')
    start_date_entry.set_date(day_before_yesterday)
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
    end_date_entry.set_date(yesterday)
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

# --- 路徑設定 (與原 logic 一致) ---
tc_tf_page_paths = [
    {"path": "WithdrawRiskControl", "channel": "草"},
    {"path": "WithdrawRiskControl/VcpIndex", "channel": "U"},
    {"path": "ElectronicPurseWithdrawExamination/RiskIndex", "channel": "GB"}
]
other_page_paths = [
    {"path": "WithdrawExamination/RiskIndex", "channel": "草"},
    {"path": "USDTWithdrawExamination/RiskIndex", "channel": "U"},
    {"path": "ElectronicPurseWithdrawExamination/RiskIndex", "channel": "GB"}
]

# --- 異步抓取單一站台邏輯 ---
async def scrape_site_page_async(site, page, page_path, channel, start_date, end_date, is_tc_tf=False):
    results = []
    full_url = f"{site['url'].rstrip('/')}/{page_path}"
    try:
        await page.goto(full_url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2)
        
        # 處理彈窗
        try:
            close_btn = page.locator("#Warning-Dialog-CloseBtn")
            if await close_btn.is_visible(timeout=2000): await close_btn.click()
        except: pass

        # 填寫日期並查詢 (依據 site 類型切換選擇器)
        if is_tc_tf:
            await page.fill("#StartCreateTime", "")
            await page.fill("#EndCreateTime", "")
            await page.fill("#StartConfirmTime", start_date)
            await page.fill("#EndConfirmTime", end_date)
            await page.click("#search > span")
        else:
            await page.fill("#StartCreateTime", "")
            await page.fill("#EndCreateTime", "")
            await page.fill("#StartRiskControlConfirmTime", start_date)
            await page.fill("#EndRiskControlConfirmTime", end_date)
            search_btn = page.locator("button:has-text('查询'), button:has-text('Search'), #queryForm button").first
            await search_btn.click()

        await asyncio.sleep(3)

        # 翻頁抓取表格內容
        while True:
            try:
                await page.wait_for_selector("#RequestTable", timeout=5000)
            except: break

            table_data = await page.evaluate("""() => {
                const headers = Array.from(document.querySelectorAll('#RequestTable thead th')).map(th => th.innerText.trim());
                const rows = Array.from(document.querySelectorAll('#RequestTable tbody tr')).map(tr => 
                    Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim())
                );
                return { headers, rows };
            }""")

            if not table_data['rows'] or "表中数据为空" in (table_data['rows'][0][0] if table_data['rows'] else ""):
                break

            for row in table_data['rows']:
                if len(row) >= 4:
                    results.append({"headers": table_data['headers'], "data": row, "channel": channel, "page_path": page_path})

            next_btn = page.locator("#RequestTable_next > a")
            parent_li = page.locator("#RequestTable_next")
            if not await next_btn.is_visible() or "disabled" in (await parent_li.get_attribute("class") or ""):
                break
            await next_btn.click()
            await asyncio.sleep(2)

    except Exception as e:
        print(f"Error scraping {site['name']}: {e}")
    return results

# --- 數據處理 (與原 logic 一致) ---
def process_data_logic(site_name, raw_results):
    processed = []
    target_cols = ["状态", "操作者", "申请日期", "确认日期"]
    variants = {
        "状态": ["状态", "状态码", "审核状态", "處理狀態", "審核狀態"],
        "操作者": ["操作者", "操作人员", "审核人", "經辦人", "處理人員"],
        "申请日期": ["申请日期", "申请时间", "创建时间", "申請日期", "申請時間"],
        "确认日期": ["确认日期", "确认时间", "處理時間", "確認日期", "審核時間"]
    }

    for item in raw_results:
        headers = [clean_text(h) for h in item["headers"]]
        data = [clean_text(d) for d in item["data"]]
        
        idx_map = {}
        for col, v_list in variants.items():
            for v in v_list:
                if v in headers:
                    idx_map[col] = headers.index(v)
                    break
        
        # TC/TF U 渠道特殊補位邏輯
        if item["page_path"] == "WithdrawRiskControl/VcpIndex" and site_name in ["TC", "TF"]:
            if "状态" in idx_map:
                s_idx = idx_map["状态"]
                idx_map["操作者"] = s_idx + 1
                idx_map["申请日期"] = s_idx + 2
                idx_map["确认日期"] = s_idx + 3

        if "状态" in idx_map and "申请日期" in idx_map:
            row_dict = {"平台": site_name, "渠道": item["channel"]}
            for col in target_cols:
                idx = idx_map.get(col)
                row_dict[col] = data[idx] if idx is not None and idx < len(data) else ""
            processed.append(row_dict)
    return processed

# --- 主並行任務 ---
async def site_worker(site, page, start_date, end_date, final_dict):
    all_raw = []
    paths = tc_tf_page_paths if site["name"] in ["TC", "TF"] else other_page_paths
    is_tc = site["name"] in ["TC", "TF"]
    
    for p_info in paths:
        raw = await scrape_site_page_async(site, page, p_info["path"], p_info["channel"], start_date, end_date, is_tc)
        all_raw.extend(raw)
    
    final_dict[site["name"]] = process_data_logic(site["name"], all_raw)

# --- Program 4 入口函式 ---
def run_program_4(root, selected_sites, pages, cb):
    # 1. 彈出日期選擇 (同步 UI)
    selected_data = select_dates(root)
    
    # 2. 如果使用者關閉視窗或沒選日期，必須呼叫 cb() 恢復主 UI 並退出
    if not selected_data["start_date"]:
        print("使用者取消日期選擇")
        if cb: root.after(0, cb)
        return

    start_date = selected_data["start_date"]
    end_date = selected_data["end_date"]
    all_site_data = {}

    async def _async_wrapper():
        try:
            # 3. 並行執行所有站台
            tasks = [site_worker(s, p, start_date, end_date, all_site_data) for s, p in zip(selected_sites, pages)]
            await asyncio.gather(*tasks)

            # 4. 存檔 (同步 UI 操作)
            if any(all_site_data.values()):
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=f"審單DATA_{datetime.now().strftime('%Y%m%d')}.xlsx"
                )
                if save_path:
                    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                        for site_name, rows in all_site_data.items():
                            if rows:
                                df = pd.DataFrame(rows)
                                df.to_excel(writer, index=False, sheet_name=site_name)
                    messagebox.showinfo("完成", f"存檔成功：{save_path}")
            else:
                messagebox.showinfo("提示", "所選區間查無數據")

        except Exception as e:
            messagebox.showerror("程式錯誤", str(e))
        finally:
            # 5. 無論成功或失敗，最後一定要解鎖主介面
            if cb: root.after(0, cb)

    # 透過 MainApp 的非同步循環執行
    from __main__ import app
    app.run_async(_async_wrapper())