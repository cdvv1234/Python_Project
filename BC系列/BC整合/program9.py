import tkinter as tk
from tkinter import messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime
import asyncio
import time

def run_program_9(root, selected_sites, pages, callback):
    # 這裡多接收一個 callback 參數
    
    def create_input_window():
        input_window = tk.Toplevel(root)
        input_window.title("彩種玩法統計 (多站並行版)")
        input_window.geometry("500x350")
        input_window.grab_set()

        tk.Label(input_window, text="彩種玩法統計", font=("Helvetica", 16, "bold")).pack(pady=20)

        tk.Label(input_window, text="開始日期 (YYYY/MM/DD)：").pack(anchor="w", padx=50)
        start_date_entry = DateEntry(input_window, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
        start_date_entry.pack(pady=5)

        tk.Label(input_window, text="結束日期 (YYYY/MM/DD)：").pack(anchor="w", padx=50, pady=(20,0))
        end_date_entry = DateEntry(input_window, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
        end_date_entry.pack(pady=5)

        async def fetch_site_data(site, page, start_date_str, end_date_str):
            site_name = site['name']
            is_tc_tf = site_name in ["TC", "TF"]
            data_list = []

            try:
                # 導航 (保留您的原始 BetTypeStatistics)
                navigated = False
                for attempt in range(1, 4):
                    try:
                        await page.goto(f"{site['url']}BetTypeStatistics", wait_until="networkidle", timeout=30000)
                        await asyncio.sleep(2)
                        if await page.locator('#StartTime').count() > 0:
                            navigated = True
                            break
                    except:
                        if attempt < 3: await asyncio.sleep(2)

                if not navigated:
                    print(f"站台 {site_name} 導航失敗")
                    return site_name, []

                # 填寫日期
                await page.fill('#StartTime', "")
                await page.fill('#EndTime', "")
                await page.fill('#StartTime', start_date_str)
                await page.fill('#EndTime', end_date_str)
                
                # --- 核心修正：同時隱藏兩種日期選擇器 ---
                await page.evaluate("""
                    // 1. 隱藏其他站點使用的 xdsoft 選擇器
                    const xdsoft = document.querySelectorAll('.xdsoft_datetimepicker');
                    xdsoft.forEach(el => { el.style.display = 'none'; el.style.visibility = 'hidden'; });

                    // 2. 隱藏 TC/TF 使用的 jQuery UI 選擇器
                    const ui_datepicker = document.getElementById('ui-datepicker-div');
                    if (ui_datepicker) {
                        ui_datepicker.style.display = 'none';
                        ui_datepicker.style.visibility = 'hidden';
                    }
                """)
                await asyncio.sleep(1)

                # 下拉選單處理
                if is_tc_tf:
                    dropdown_btn = '#LottoGame'
                    option_selector = 'ul.dropdown-menu li.LottoGame'
                    await page.click(dropdown_btn)
                    await page.wait_for_selector(option_selector, state="visible", timeout=10000)
                else:
                    dropdown_btn = '#LottoGameID'
                    option_selector = 'select#LottoGameID option.LottoGame'
                    await page.wait_for_selector(option_selector, state="attached", timeout=10000)

                options = await page.locator(option_selector).all()
                filtered_info = []
                for opt in options:
                    val = await opt.get_attribute('value')
                    txt = await opt.inner_text()
                    if val and val != "0":
                        filtered_info.append({"val": val, "txt": txt.strip()})

                for info in filtered_info:
                    try:
                        await page.click(dropdown_btn)
                        if is_tc_tf:
                            btn = page.locator('#LottoGame')
                            if await btn.get_attribute('aria-expanded') == "false": await btn.click()
                            await page.locator(f'ul.dropdown-menu li.LottoGame[value="{info["val"]}"]').click(force=True)
                        else:
                            await page.select_option('#LottoGameID', value=info["val"])

                        query_btn = '#search' if is_tc_tf else '#activeSearch'
                        await page.click(query_btn, force=True)
                        await asyncio.sleep(2.5)

                        while True:
                            table_id = '#LogList' if is_tc_tf else '#BetTypeStatisticTable'
                            if await page.locator(f"{table_id} td.dataTables_empty").is_visible(): break
                            
                            rows = page.locator(f"{table_id} tbody tr")
                            count = await rows.count()
                            for r in range(count):
                                cells = await rows.nth(r).locator('td').all_inner_texts()
                                if is_tc_tf and len(cells) >= 9:
                                    data_list.append([site_name, f"{start_date_str}~{end_date_str}", info["txt"]] + [c.strip() for c in cells[:9]])
                                elif not is_tc_tf and len(cells) >= 8:
                                    data_list.append([site_name, f"{start_date_str}~{end_date_str}", info["txt"]] + [c.strip() for c in cells[:8]])

                            next_btn_sel = '#LogList_next' if is_tc_tf else '#BetTypeStatisticTable_next'
                            cls = await page.locator(next_btn_sel).get_attribute('class')
                            if "disabled" in (cls or ""): break
                            await page.locator(next_btn_sel).locator('a').click()
                            await asyncio.sleep(2)
                    except Exception as e:
                        print(f"站台 {site_name} 彩種 {info['txt']} 錯誤: {e}")

                return site_name, data_list
            except Exception as e:
                print(f"站台 {site_name} 發生嚴重錯誤: {e}")
                return site_name, []

        async def _async_process_statistics():
            try:
                start_date_str = start_date_entry.get_date().strftime('%Y/%m/%d')
                end_date_str = end_date_entry.get_date().strftime('%Y/%m/%d')
                tasks = [fetch_site_data(site, page, start_date_str, end_date_str) for site, page in zip(selected_sites, pages)]
                results = await asyncio.gather(*tasks)
                all_data = {site_name: data for site_name, data in results if data}
                root.after(0, lambda: finalize_ui(all_data))
            except Exception as e:
                print(f"非同步執行出錯: {e}")
                if callback: callback() # 出錯時解鎖

        def finalize_ui(all_data):
            if all_data:
                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"彩種玩法統計_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
                if save_path:
                    cols = ['平台', '期間', '彩种', '彩类', '玩法', '人數', '投注筆數', '投注金額', '獎金', '盈虧', 'RTP']
                    with pd.ExcelWriter(save_path) as writer:
                        for s_name, data in all_data.items():
                            cur_cols = cols.copy()
                            if s_name in ["TC", "TF"]: cur_cols.insert(9, '返點')
                            pd.DataFrame(data, columns=cur_cols).to_excel(writer, sheet_name=s_name, index=False)
                    messagebox.showinfo("完成", "儲存成功")
            
            # 判斷是否繼續
            if messagebox.askyesno("繼續", "是否繼續執行？"):
                input_window.destroy()
                create_input_window()
            else:
                input_window.destroy()
                if callback: callback() # 點擊「否」正式結束，解鎖按鈕

        def start_task():
            # 這裡不關閉視窗，依據您的原始邏輯
            from __main__ import app
            app.run_async(_async_process_statistics())

        tk.Button(input_window, text="開始執行", command=start_task, bg="#4CAF50", fg="white", font=("Helvetica", 14), height=2).pack(pady=30)
        
        # 處理使用者直接關閉日期視窗的情況
        def on_close():
            input_window.destroy()
            if callback: callback()
        input_window.protocol("WM_DELETE_WINDOW", on_close)

    create_input_window()