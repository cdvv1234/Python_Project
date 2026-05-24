import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime, timedelta
import asyncio
import os

# --- 核心遞迴抓取邏輯 ---
async def fetch_category_recursive(page, site, account, s_d, e_d, superior, cat_name, cat_val, all_results, should_recursive):
    """
    對特定帳號進行抓取。根據 should_recursive 決定是否繼續向下抓取下級。
    """
    try:
        # 1. 導航到報表頁面
        current_url = page.url
        target_path = "TeamProfitReport" if site["name"] in ["TC", "TF"] else "ProfitReport"
        if target_path not in current_url:
            await page.goto(f"{site['url']}{target_path}", wait_until="networkidle")
            await asyncio.sleep(1)

        # 2. 執行查詢邏輯
        sub_table_id = "" 

        if site["name"] in ["TC", "TF"]:
            # --- TC/TF 邏輯 ---
            start_dt = datetime.strptime(s_d, '%Y-%m-%d')
            end_dt = datetime.strptime(e_d, '%Y-%m-%d') + timedelta(days=1)
            
            await page.fill('#StartTime', start_dt.strftime('%Y/%m/%d 03:00'))
            await page.fill('#EndTime', end_dt.strftime('%Y/%m/%d 03:00'))
            await page.locator('#GameTypeId').click()
            await page.locator(f'li.GameType[value="{cat_val}"] a').click()
            
            await page.fill('#LoginID', account)
            await page.click('#querybutton')
            await asyncio.sleep(2.5) 
            
            tab_sel = '#TeamProfitTable' if cat_name == "彩票" else '#TeamProfitTableForExternalGame'
            data = await page.evaluate(f"""(sel) => {{ 
                const row = Array.from(document.querySelectorAll(sel + ' tbody tr')).find(r => r.innerText.includes('总和') || r.innerText.includes('總和'));
                return row ? Array.from(row.querySelectorAll('td')).map(c => c.innerText.trim()) : null;
            }}""", tab_sel)
            
            if data:
                cols = ['总投注', '总奖金', '总返點', '总活動', '总盈虧', '总充值', '总提款', '总紅利'] if cat_name == "彩票" else ['项目', '总投注', '有效投注', '总奖金', '总活動', '总盈虧']
                df_data = [[site['name'], superior, account] + data[:8]] if cat_name == "彩票" else [[site['name'], cat_name.replace("電子",""), superior, account] + data[:6]]
                header = ['平台', '直屬', '帳號'] + cols if cat_name == "彩票" else ['平台', '類別', '直屬', '帳號'] + cols
                all_results.setdefault(f"{site['name']}_{cat_name}", []).append(pd.DataFrame(df_data, columns=header))
            
            sub_table_id = "#TeamMemberProfitTable"

        else:
            # --- 一般站台邏輯 ---
            if cat_name == "彩票":
                await page.locator('#lottery-tab').click()
                await page.fill('#StartTime', s_d)
                await page.fill('#EndTime', e_d)
                await page.fill('#LoginId', account)
                await page.locator('button[data-bind*="SearchClick"]').first.click()
                await asyncio.sleep(2)
                
                data = await page.evaluate("""() => {
                    const row = document.querySelector('tbody[data-bind*="PersonalStatistics"] tr');
                    return row ? Array.from(row.querySelectorAll('td')).map(c => c.innerText.trim()) : null;
                }""")
                if data:
                    cols = ['总投注', '总奖金', '总返點', '总活動', '总盈虧', '总充值', '总提款', '总紅利', '平台服務費']
                    df_data = [[site['name'], superior, account] + data]
                    all_results.setdefault(f"{site['name']}_彩票", []).append(pd.DataFrame(df_data, columns=['平台', '直屬', '帳號'] + cols[:len(data)]))
                
                sub_table_id = "#MemberDataTable"
            else:
                await page.locator('#external-tab').click()
                await page.fill('#External #StartTime', s_d)
                await page.fill('#External #EndTime', e_d)
                await page.fill('#External #LoginId', account)
                await page.select_option('#External #Category', cat_val)
                await page.click('#External #SearchForm button.btn-primary')
                await asyncio.sleep(2)
                
                data = await page.evaluate("""() => { 
                    const body = document.querySelector('#External tbody[data-bind*="PersonalStatistics"]');
                    if(!body) return null;
                    const sumRow = Array.from(body.querySelectorAll('tr')).find(r => r.innerText.includes('总和'));
                    return sumRow ? Array.from(sumRow.querySelectorAll('td')).map(c => c.innerText.trim()) : null; 
                }""")
                if data:
                    cols = ['项目', '总投注', '有效投注', '总奖金', '总盈虧', '总返點', '总活動', '結果']
                    df_data = [[site['name'], cat_name.replace("電子",""), superior, account] + data]
                    all_results.setdefault(f"{site['name']}_{cat_name}", []).append(pd.DataFrame(df_data, columns=['平台', '類別', '直屬', '帳號'] + cols[:len(data)]))
                
                sub_table_id = "#MemberDataTable"

        # --- 判斷是否要繼續抓下級 (遞迴關鍵) ---
        if not should_recursive:
            return # 使用者未勾選「往下查詢」，在此停止

        # 3. 抓取下級名單
        subordinate_accounts = []
        try:
            await asyncio.sleep(1) 
            rows = await page.query_selector_all(f"{sub_table_id} tbody tr")
            for r in rows:
                acc_link = await r.query_selector("a.searchNest")
                if not acc_link:
                    acc_link = await r.query_selector("td:nth-child(1) > a")
                if not acc_link:
                    acc_link = await r.query_selector("td:nth-child(1)")

                if acc_link:
                    txt = (await acc_link.inner_text()).strip()
                    if txt and txt != account and not any(skip in txt for skip in ["無數據", "总和", "總和", "查询无记录", "查無數據"]):
                        subordinate_accounts.append(txt)
            subordinate_accounts = list(dict.fromkeys(subordinate_accounts))
        except: pass

        # 4. 遞迴下級
        for sub_acc in subordinate_accounts:
            await fetch_category_recursive(page, site, sub_acc, s_d, e_d, superior, cat_name, cat_val, all_results, should_recursive)

    except Exception as e:
        print(f"[{cat_name}] 帳號 {account} 處理中斷: {e}")

# --- 站台任務分配 ---
async def process_site_task(site, page, excel_data, s_d, e_d, selected_cats, category_map, all_results, should_recursive):
    if site["name"] not in excel_data: return
    df = excel_data[site["name"]]
    start_accounts = df.iloc[:, 0].dropna().astype(str).str.strip().tolist() if df.shape[1] == 1 else df.iloc[:, 1].dropna().astype(str).str.strip().tolist()
    
    for cat_name in selected_cats:
        cat_val = category_map[cat_name]
        for acc in start_accounts:
            await fetch_category_recursive(page, site, acc, s_d, e_d, acc, cat_name, cat_val, all_results, should_recursive)

# --- UI 介面 ---
def run_program_5(root, selected_sites, pages, callback):
    def create_input_window():
        input_window = tk.Toplevel(root)
        input_window.title("直屬及下級盈虧 (遞迴開關版)")
        input_window.geometry("400x600")
        input_window.grab_set()

        tk.Label(input_window, text="1. 選擇 Excel 檔案：", font=("Arial", 10, "bold")).pack(pady=5)
        excel_path = tk.StringVar()
        tk.Button(input_window, text="瀏覽檔案", command=lambda: excel_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).pack(pady=5)
        tk.Label(input_window, textvariable=excel_path, wraplength=350, fg="blue").pack(pady=2)

        tk.Label(input_window, text="2. 設定日期：", font=("Arial", 10, "bold")).pack(pady=5)
        start_date_entry = DateEntry(input_window, width=12, date_pattern='yyyy-mm-dd')
        start_date_entry.pack(pady=2)
        end_date_entry = DateEntry(input_window, width=12, date_pattern='yyyy-mm-dd')
        end_date_entry.pack(pady=2)

        # --- 新增：往下查詢開關 ---
        tk.Label(input_window, text="3. 查詢範圍設定：", font=("Arial", 10, "bold")).pack(pady=5)
        is_downward_var = tk.BooleanVar(value=True) # 預設打勾
        tk.Checkbutton(input_window, text="往下查詢 (遞迴抓取下級數據)", variable=is_downward_var, fg="red").pack(pady=2)

        tk.Label(input_window, text="4. 勾選類別：", font=("Arial", 10, "bold")).pack(pady=5)
        check_vars = {cat: tk.BooleanVar(value=(cat == "彩票")) for cat in ["彩票", "真人電子", "體育", "棋牌"]}
        for label in check_vars:
            tk.Checkbutton(input_window, text=label, variable=check_vars[label]).pack(anchor="w", padx=140)

        all_results = {}

        async def _async_process():
            try:
                f_path = excel_path.get()
                if not f_path:
                    messagebox.showwarning("錯誤", "未選擇檔案")
                    return
                
                excel_data = pd.read_excel(f_path, sheet_name=None)
                s_d, e_d = start_date_entry.get(), end_date_entry.get()
                selected_cats = [c for c, v in check_vars.items() if v.get()]
                category_map = {"彩票": "0", "真人電子": "1", "體育": "2", "棋牌": "3"}
                should_recursive = is_downward_var.get()

                tasks = [process_site_task(site, page, excel_data, s_d, e_d, selected_cats, category_map, all_results, should_recursive) 
                         for site, page in zip(selected_sites, pages)]
                await asyncio.gather(*tasks)
                
                if all_results:
                    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"盈虧報表_{datetime.now().strftime('%m%d_%H%M')}.xlsx")
                    if save_path:
                        # 排序順序定義
                        cat_order = ["彩票", "真人電子", "體育", "棋牌"]
                        site_order = ["TC", "TF", "TS", "SY", "FL", "WX", "XC", "XH", "CJ", "CY", "YD"]
                        
                        ordered_keys = [f"{s}_{c}" for c in cat_order for s in site_order if f"{s}_{c}" in all_results]

                        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                            for key in ordered_keys:
                                pd.concat(all_results[key], ignore_index=True).to_excel(writer, sheet_name=key[:31], index=False)
                        messagebox.showinfo("完成", f"存檔成功：\n{save_path}")
                else:
                    messagebox.showinfo("提示", "未抓取到任何數據")
            except Exception as e:
                messagebox.showerror("程式故障", str(e))
            finally:
                if input_window.winfo_exists(): input_window.destroy()
                if callback: root.after(0, callback)

        def start():
            import __main__
            if hasattr(__main__, 'app'): __main__.app.run_async(_async_process())

        tk.Button(input_window, text="開始並行查詢", command=start, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), height=2).pack(pady=25)

    create_input_window()