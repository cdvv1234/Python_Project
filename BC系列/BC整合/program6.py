import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime, timedelta
import asyncio
import os
import time

def run_program_6(root, selected_sites, pages, callback):
    def create_input_window():
        input_window = tk.Toplevel(root)
        input_window.title("招商觀察 (日期BUG修復+排序版)")
        input_window.geometry("400x550")
        input_window.grab_set()

        # Excel 檔案選擇
        tk.Label(input_window, text="選擇包含帳號的 Excel 檔案：", font=("Arial", 10, "bold")).pack(pady=5)
        excel_path = tk.StringVar()
        tk.Button(input_window, text="瀏覽", command=lambda: excel_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).pack(pady=5)
        tk.Label(input_window, textvariable=excel_path, wraplength=350, fg="grey").pack(pady=5)

        # 日期選擇
        tk.Label(input_window, text="日期範圍：", font=("Arial", 10, "bold")).pack(pady=5)
        start_date_entry = DateEntry(input_window, width=12, date_pattern='yyyy-mm-dd')
        start_date_entry.pack(pady=5)
        end_date_entry = DateEntry(input_window, width=12, date_pattern='yyyy-mm-dd')
        end_date_entry.pack(pady=5)

        # 核選方塊
        tk.Label(input_window, text="選擇查詢項目：", font=("Arial", 10, "bold")).pack(pady=5)
        check_vars = {cat: tk.BooleanVar(value=True) for cat in ["彩票", "真人電子", "體育", "棋牌"]}
        for label in check_vars:
            tk.Checkbutton(input_window, text=label, variable=check_vars[label]).pack(anchor="w", padx=120)

        all_results = {}

        # --- [1] 彩票欄位判定 (8/9欄) ---
        async def fetch_data_lottery(page, site, account):
            try:
                await page.fill('#LoginId', account)
                await page.locator('button[data-bind*="SearchClick"]').first.click()
                await asyncio.sleep(1.5)
                await page.wait_for_selector('tbody[data-bind="with: TeamStatistics"]', timeout=15000)
                table_data = await page.evaluate("() => Array.from(document.querySelectorAll('tbody[data-bind=\"with: TeamStatistics\"] tr')).map(r => Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()))")
                if table_data:
                    actual_cols = len(table_data[0])
                    if actual_cols == 8:
                        cols = ['总投注', '总奖金', '总返点', '总活動', '总盈亏', '总充值', '总提款', '总红利']
                    elif actual_cols == 9:
                        cols = ['总投注', '总奖金', '总返点', '总活動', '总盈虧', '总充值', '总提款', '总紅利', '平台服务费']
                    else:
                        cols = [f"列_{i+1}" for i in range(actual_cols)]
                    df_data = [[site['name'], account] + row for row in table_data]
                    df = pd.DataFrame(df_data, columns=['平台', '帳號'] + cols)
                    all_results.setdefault(f"{site['name']}_彩票", []).append(df)
            except: pass

        # --- [2] 其餘項目欄位判定 (6/8欄) ---
        async def fetch_data_external(page, site, account, s_d, e_d, cat, cat_val):
            try:
                await page.locator('#external-tab').click()
                await page.fill('#External #LoginId', account)
                await page.fill('#External #StartTime', s_d)
                await page.fill('#External #EndTime', e_d)
                await page.select_option('#External #Category', cat_val)
                await page.click('#External #SearchForm button.btn-primary')
                await asyncio.sleep(1.5)
                data = await page.evaluate("""() => { 
                    const row = Array.from(document.querySelectorAll('#External .table-responsive table tbody tr')).find(r => r.innerText.includes('总和')); 
                    return row ? Array.from(row.querySelectorAll('td')).map(c => c.innerText.trim()) : []; 
                }""")
                if data:
                    actual_cols = len(data)
                    if actual_cols == 6:
                        cols = ['项目', '总投注', '有效投注', '总奖金', '总活動', '总盈亏']
                    elif actual_cols == 8:
                        cols = ['项目', '总投注', '有效投注', '总奖金', '总盈虧', '总返點', '总活動', '结果']
                    else:
                        cols = [f"列_{i+1}" for i in range(actual_cols)]
                    display_cat = cat.replace("電子", "")
                    df_data = [[site['name'], display_cat, account] + data]
                    df = pd.DataFrame(df_data, columns=['平台', '類別', '帳號'] + cols)
                    all_results.setdefault(f"{site['name']}_{cat}", []).append(df)
            except: pass

        # --- [3] TC/TF 抓取函數 (日期填寫修復 + 大小寫不敏感) ---
        async def fetch_data_tc_tf(page, site, account, category, category_value):
            try:
                await page.locator('#GameTypeId').click()
                await page.locator(f'li.GameType[value="{category_value}"] a').click()
                await page.fill('#LoginID', account)
                await page.click('#querybutton')
                res_sel = '#teamResult' if category == "彩票" else '#teamResultForExternalGame'
                await page.wait_for_function("(args) => { const el = document.querySelector(args.sel); return el && el.innerText.toLowerCase().includes(args.acc.toLowerCase()); }", arg={"sel": res_sel, "acc": account}, timeout=15000)
                tab_sel = '#TeamProfitTable' if category == "彩票" else '#TeamProfitTableForExternalGame'
                if category == "彩票":
                    data = await page.evaluate(f"() => Array.from(document.querySelectorAll('{tab_sel} tbody tr')).map(r => Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()))")
                    if data:
                        actual_cols = len(data[0])
                        cols = ['总投注', '总奖金', '总返點', '总活動', '总盈虧', '总充值', '总提款', '总紅利']
                        df_data = [[site['name'], account] + r for r in data]
                        df = pd.DataFrame(df_data, columns=['平台', '帳號'] + cols[:actual_cols])
                        all_results.setdefault(f"{site['name']}_彩票", []).append(df)
                else:
                    data = await page.evaluate("""(sel) => { 
                        const row = Array.from(document.querySelectorAll(sel + ' tbody tr')).find(r => r.innerText.includes('总和')); 
                        return row ? Array.from(row.querySelectorAll('td')).map(c => c.innerText.trim()) : null; 
                    }""", tab_sel)
                    if data:
                        display_cat = category.replace("電子", "")
                        df_data = [[site['name'], display_cat, account] + data[:6]]
                        cols = ['平台', '類別', '帳號', '项目', '总投注', '有效投注', '总奖金', '总活動', '总盈亏']
                        df = pd.DataFrame(df_data, columns=cols)
                        all_results.setdefault(f"{site['name']}_{category}", []).append(df)
            except: pass

        # --- [4] 站台處理任務 (包含日期填寫修復邏輯) ---
        async def process_site(site, page, excel_data, s_d, e_d, category_map):
            try:
                url_suffix = "TeamProfitReport" if site["name"] in ["TC", "TF"] else "ProfitReport"
                await page.goto(f"{site['url']}{url_suffix}", wait_until="networkidle")
                await page.wait_for_load_state("domcontentloaded")
                
                # 填入時間邏輯修正 (比照原始碼)
                await page.wait_for_selector('#StartTime', timeout=30000)
                await page.wait_for_selector('#EndTime', timeout=30000)

                if site["name"] in ["TC", "TF"]:
                    # 格式: 2026/01/10 03:00 且 結束日+1
                    start_dt_obj = datetime.strptime(s_d, '%Y-%m-%d')
                    end_dt_obj = datetime.strptime(e_d, '%Y-%m-%d') + timedelta(days=1)
                    start_time_str = start_dt_obj.strftime('%Y/%m/%d 03:00')
                    end_time_str = end_dt_obj.strftime('%Y/%m/%d 03:00')
                    
                    # 先清空再填入
                    await page.fill('#StartTime', "")
                    await page.fill('#EndTime', "")
                    await page.fill('#StartTime', start_time_str)
                    await page.fill('#EndTime', end_time_str)
                else:
                    # 一般站台直接填 yyyy-mm-dd
                    await page.fill('#StartTime', s_d)
                    await page.fill('#EndTime', e_d)

                accounts = []
                if site["name"] in excel_data:
                    df = excel_data[site["name"]]
                    accounts = df.iloc[:, 0].dropna().astype(str).str.strip().tolist() if df.shape[1] == 1 else df.iloc[:, 1].dropna().astype(str).str.strip().tolist()

                for cat, var in check_vars.items():
                    if var.get():
                        for acc in accounts:
                            if site["name"] in ["TC", "TF"]:
                                await fetch_data_tc_tf(page, site, acc, cat, category_map[cat])
                            elif cat == "彩票":
                                await fetch_data_lottery(page, site, acc)
                            else:
                                await fetch_data_external(page, site, acc, s_d, e_d, cat, category_map[cat])
            except Exception as e:
                print(f"站台 {site['name']} 導航出錯: {e}")

        # --- [5] 主要調度 ---
        async def _async_process():
            try:
                excel_file = excel_path.get()
                if not excel_file: return
                excel_data = pd.read_excel(excel_file, sheet_name=None)
                s_d, e_d = start_date_entry.get(), end_date_entry.get()
                category_map = {"彩票": "0", "真人電子": "1", "體育": "2", "棋牌": "3"}

                tasks = [process_site(site, page, excel_data, s_d, e_d, category_map) for site, page in zip(selected_sites, pages)]
                await asyncio.gather(*tasks)
                root.after(0, finalize_save)
            except Exception as e:
                print(f"全局並行錯誤: {e}")

        def finalize_save():
            if all_results:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M')
                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"招商觀察_{timestamp}.xlsx")
                if save_path:
                    # 排序順序定義
                    cat_order = ["彩票", "真人電子", "體育", "棋牌"]
                    site_order = ["TC", "TF", "TS", "SY", "FL", "WX", "XC", "XH", "CJ", "CY"]
                    
                    ordered_keys = [f"{s}_{c}" for c in cat_order for s in site_order if f"{s}_{c}" in all_results]

                    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                        for key in ordered_keys:
                            pd.concat(all_results[key], ignore_index=True).to_excel(writer, sheet_name=key[:31], index=False)
                    messagebox.showinfo("成功", "抓取完成並排序儲存")
            input_window.destroy()
            if callback: callback()

        def start_task():
            import __main__
            if hasattr(__main__, 'app'): __main__.app.run_async(_async_process())

        tk.Button(input_window, text="確認查詢", command=start_task, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), height=2).pack(pady=20)

    create_input_window()