import tkinter as tk
from tkinter import messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime, timedelta
import asyncio
import os

def run_program_10(root, selected_sites, pages, callback):
    def create_input_window():
        input_window = tk.Toplevel(root)
        input_window.title("彩票遊戲統計 (穩定強化版)")
        input_window.geometry("400x300")
        input_window.grab_set()

        # UI 介面
        tk.Label(input_window, text="選擇查詢日期範圍：", font=("Arial", 10, "bold")).pack(pady=10)
        frame_date = tk.Frame(input_window)
        frame_date.pack(pady=5)
        start_date_entry = DateEntry(frame_date, width=12, date_pattern='yyyy-mm-dd')
        start_date_entry.pack(side="left", padx=5)
        tk.Label(frame_date, text="至").pack(side="left")
        end_date_entry = DateEntry(frame_date, width=12, date_pattern='yyyy-mm-dd')
        end_date_entry.pack(side="left", padx=5)

        # --- 核心抓取函數 (改為回傳數據) ---
        async def fetch_site_stats(site, page, s_d, e_d):
            site_data = []
            try:
                # 1. 決定路徑與表格 ID
                if site["name"] in ["TC", "TF"]:
                    url_path = "LotteryStatistics"
                    table_selector = "#RequestTable"
                else:
                    url_path = "LotteryGameStatistics"
                    table_selector = "#LotteryTable"

                print(f"[{site['name']}] 開始導航...")
                await page.goto(f"{site['url']}{url_path}", wait_until="networkidle", timeout=60000)
                
                # 2. 處理日期填寫
                search_date_str = s_d.replace('-', '/')
                if site["name"] in ["TC", "TF"]:
                    s_dt = datetime.strptime(s_d, '%Y-%m-%d').strftime('%Y/%m/%d 03:00')
                    e_dt = (datetime.strptime(e_d, '%Y-%m-%d') + timedelta(days=1)).strftime('%Y/%m/%d 03:00')
                    await page.fill('#StartTime', "")
                    await page.fill('#EndTime', "")
                    await page.fill('#StartTime', s_dt)
                    await page.fill('#EndTime', e_dt)
                else:
                    await page.fill('#StartTime', s_d)
                    await page.fill('#EndTime', e_d)

                # 3. 隱藏日期選擇器 (使用您提供的精確代碼)
                await page.evaluate("""
                    const xdsoft = document.querySelectorAll('.xdsoft_datetimepicker');
                    xdsoft.forEach(el => { el.style.display = 'none'; el.style.visibility = 'hidden'; });
                    const ui_datepicker = document.getElementById('ui-datepicker-div');
                    if (ui_datepicker) {
                        ui_datepicker.style.display = 'none';
                        ui_datepicker.style.visibility = 'hidden';
                    }
                """)
                
                # 4. 點擊查詢
                await page.wait_for_selector('#querybutton', timeout=10000)
                await page.click('#querybutton')
                print(f"[{site['name']}] 已點擊查詢，等待數據更新...")
                
                # 5. 等待查詢條件更新 (比對 #searchCondition)
                try:
                    await page.wait_for_function(
                        "(args) => { const el = document.querySelector('#searchCondition'); return el && el.innerText.includes(args.dateStr); }",
                        arg={"dateStr": search_date_str},
                        timeout=25000
                    )
                except Exception as e:
                    print(f"[{site['name']}] 警告：查詢條件比對逾時，嘗試直接抓取...")

                # 6. 等待表格行出現 (確保 AJAX 完成)
                await asyncio.sleep(2) # 給予緩衝時間
                await page.wait_for_selector(f"{table_selector} tbody tr", timeout=15000)

                # 7. 抓取表格內容
                rows_data = await page.evaluate(f"""(sel) => {{
                    const rows = Array.from(document.querySelectorAll(sel + ' tbody tr'));
                    // 過濾掉明顯不是數據行的行 (例如只有1欄的"無數據"提示)
                    return rows.filter(r => r.cells.length > 5)
                               .map(r => Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()));
                }}""", table_selector)

                if rows_data:
                    period = f"{s_d.replace('-', '/')}~{e_d.replace('-', '/')}"
                    for row in rows_data:
                        # 根據站台類型存入數據，並標記來源
                        site_data.append({"type": "TC_TF" if site["name"] in ["TC", "TF"] else "Others", 
                                          "content": [site['name'], period] + row})
                    print(f"[{site['name']}] 成功抓取 {len(rows_data)} 筆資料")
                else:
                    print(f"[{site['name']}] 查無數據 (表格為空)")

            except Exception as e:
                print(f"[{site['name']}] 發生異常: {str(e)}")
            
            return site_data

        # --- 主要處理與匯總 ---
        async def _async_process():
            s_d, e_d = start_date_entry.get(), end_date_entry.get()
            
            # 使用列表收集所有任務的結果
            tasks = [fetch_site_stats(site, page, s_d, e_d) for site, page in zip(selected_sites, pages)]
            all_raw_data = await asyncio.gather(*tasks)
            
            # 展開結果並分類
            final_tc_tf = []
            final_others = []
            
            for site_list in all_raw_data:
                for item in site_list:
                    if item["type"] == "TC_TF":
                        # TC/TF 標題 8 欄數據
                        if len(item["content"]) >= 10: # 2(平台期間)+8
                            final_tc_tf.append(item["content"][:10])
                    else:
                        # Others 標題 10 欄數據
                        if len(item["content"]) >= 12: # 2(平台期間)+10
                            final_others.append(item["content"][:12])

            root.after(0, lambda: finalize_save(final_tc_tf, final_others))

        def finalize_save(tc_tf_data, others_data):
            if not tc_tf_data and not others_data:
                messagebox.showinfo("完成", "所有站台查詢結束，未抓取到有效數據。")
                input_window.destroy()
                if callback: callback()
                return

            timestamp = datetime.now().strftime('%Y%m%d_%H%M')
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"彩票遊戲統計_{timestamp}.xlsx")

            if save_path:
                try:
                    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                        if tc_tf_data:
                            cols_tc = ['平台', '期間', '彩种', '投注人数', '投注单数', '投注', '奖金', '返点', '盈亏', 'RTP']
                            pd.DataFrame(tc_tf_data, columns=cols_tc).to_excel(writer, sheet_name="TC_TF", index=False)
                        
                        if others_data:
                            cols_oth = ['平台', '期間', '彩种', '投注人数', '投注单数', '投注', '奖金', '返點', '代理服務費', '平台服務費', '盈虧', 'RTP']
                            pd.DataFrame(others_data, columns=cols_oth).to_excel(writer, sheet_name="Others", index=False)
                    messagebox.showinfo("成功", f"統計報表已匯出，共計抓取：\nTC/TF: {len(tc_tf_data)} 筆\n其餘站台: {len(others_data)} 筆")
                except Exception as e:
                    messagebox.showerror("匯出錯誤", f"Excel 儲存失敗: {e}")

            input_window.destroy()
            if callback: callback()

        def start_task():
            import __main__
            if hasattr(__main__, 'app'):
                __main__.app.run_async(_async_process())

        tk.Button(input_window, text="開始執行查詢", command=start_task, bg="#2196F3", fg="white", font=("Arial", 11, "bold"), height=2).pack(pady=25)

    create_input_window()