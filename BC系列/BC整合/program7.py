import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
from datetime import datetime
import asyncio
from openpyxl.utils import get_column_letter
from collections import defaultdict


def run_program_7(root, selected_sites, pages, callback=None):
    def create_input_window():
        input_window = tk.Toplevel(root)
        input_window.title("帳戶管理抓取")
        input_window.geometry("440x240")
        input_window.transient(root)
        input_window.grab_set()

        tk.Label(input_window, text="選擇包含帳號的 Excel 檔案：", font=("Arial", 10)).pack(pady=10)

        excel_path = tk.StringVar()
        tk.Button(input_window, text="瀏覽 Excel 檔案", 
                  command=lambda: excel_path.set(
                      filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]))
                  ).pack(pady=5)

        tk.Label(input_window, textvariable=excel_path, wraplength=400, fg="gray", justify="left").pack(pady=5)

        all_results = defaultdict(list)

# ====================== 單一帳號查詢（TC/TF + 其他站台皆已優化） ======================
        async def fetch_data(page, site, account):
            try:
                account = str(account).strip()
                if not account:
                    return

                # 清空並填入帳號
                if site["name"] in ["TC", "TF"]:
                    await page.fill('#SearchStr', '')
                else:
                    await page.fill('#Account', '')
                await asyncio.sleep(0.8)

                if site["name"] in ["TC", "TF"]:
                    await page.fill('#SearchStr', account)
                else:
                    await page.fill('#Account', account)

                # 點擊查詢
                if site["name"] in ["TC", "TF"]:
                    await page.click('#Searchbtn')
                else:
                    await page.click('#searchBtn')

                await asyncio.sleep(1.5)

                # ====================== TC / TF（維持原本優化） ======================
                if site["name"] in ["TC", "TF"]:
                    try:
                        await page.wait_for_selector('#playersTables_processing', state='hidden', timeout=15000)
                    except:
                        pass
                    await asyncio.sleep(1.2)

                    table_data = await page.evaluate("""
                        () => Array.from(document.querySelectorAll('#playersTables tbody tr'))
                                  .map(row => Array.from(row.querySelectorAll('td')).map(td => td.innerText.trim()))
                    """)

                    account_found = any(row and len(row) > 0 and row[0] == account for row in table_data)

                    if not account_found:
                        print(f"[{site['name']}] 帳號 {account} 未在表格中找到，跳過")
                        return

                # ====================== 其他站台（新增重試機制） ======================
                else:
                    found = False
                    for attempt in range(4):   # 最多重試 3 次
                        await page.wait_for_selector('#playersTables', timeout=30000)

                        table_data = await page.evaluate("""
                            () => Array.from(document.querySelectorAll('#playersTables tbody tr'))
                                      .map(row => Array.from(row.querySelectorAll('td')).map(td => td.innerText.trim()))
                        """)

                        # 使用你指定的 selector 比對第一欄帳號
                        account_elements = await page.query_selector_all("#playersTables > tbody > tr > td.sorting_1 > span")
                        current_accounts = [await el.inner_text() for el in account_elements]

                        if account in current_accounts:
                            found = True
                            print(f"[{site['name']}] 帳號 {account} 比對成功（第 {attempt+1} 次）")
                            break
                        else:
                            print(f"[{site['name']}] 帳號 {account} 第 {attempt+1} 次比對失敗，等待 1 秒後重試...")
                            await asyncio.sleep(1)

                    if not found:
                        print(f"[{site['name']}] 帳號 {account} 重試 3 次後仍未找到，跳過")
                        return

                    # 抓取完整表格資料
                    table_data = await page.evaluate("""
                        () => Array.from(document.querySelectorAll('#playersTables tbody tr'))
                                  .map(row => Array.from(row.querySelectorAll('td')).map(td => td.innerText.trim()))
                    """)

                # ====================== 儲存資料（通用） ======================
                if table_data:
                    actual_columns = len(table_data[0]) if table_data else 0

                    if site["name"] in ["TC", "TF"]:
                        columns = ['帐号', '昵称', '下级', '时时彩奖金', '类型', '余额', '团队余额', '诚信率 (%)', '创建时间', '最后登录时间', '创建来源', '登录', '状态', '功能']
                    else:
                        columns = ['帐号', '昵称', '下级', '时时彩奖金', '类型', '余额', '团队余额', '创建时间', '最后登录时间', '最后下注时间', '创建來源', '登录', '状态', '功能']

                    used_columns = columns[:actual_columns] if actual_columns <= len(columns) else [f"列_{i+1}" for i in range(actual_columns)]
                    
                    df = pd.DataFrame(table_data, columns=used_columns)
                    all_results[site["name"]].append(df)
                    print(f"[{site['name']}] 帳號 {account} 成功抓取 {len(table_data)} 筆")

            except Exception as e:
                print(f"[{site['name']}] 帳號 {account} 錯誤: {e}")
                
        # ====================== 主處理流程 ======================
        async def _async_process():
            excel_file = excel_path.get()
            if not excel_file:
                root.after(0, lambda: messagebox.showerror("錯誤", "請選擇 Excel 檔案！"))
                return

            all_results.clear()

            # 1. 並行導航
            async def navigate(site, page):
                for attempt in range(5):
                    try:
                        if site["name"] in ["TC", "TF"]:
                            await page.goto(f"{site['url']}AccountManagement/Players", wait_until="networkidle")
                        else:
                            await page.goto(f"{site['url']}Player", wait_until="networkidle")
                        await page.wait_for_load_state("domcontentloaded")
                        await asyncio.sleep(1)
                        return True
                    except Exception as e:
                        print(f"[{site['name']}] 導航失敗 (嘗試 {attempt+1}/5): {e}")
                        await asyncio.sleep(1.5)
                return False

            nav_tasks = [navigate(site, page) for site, page in zip(selected_sites, pages)]
            await asyncio.gather(*nav_tasks)

            # 2. 讀取 Excel 帳號
            accounts_dict = {}
            try:
                excel_data = pd.read_excel(excel_file, sheet_name=None)
                for site in selected_sites:
                    site_name = site["name"]
                    if site_name in excel_data:
                        df = excel_data[site_name]
                        col_idx = 0 if df.shape[1] == 1 else 1
                        accounts_dict[site_name] = df.iloc[:, col_idx].dropna().tolist()
                    else:
                        accounts_dict[site_name] = []
            except Exception as e:
                root.after(0, lambda: messagebox.showerror("錯誤", f"讀取 Excel 失敗: {e}"))
                return

            # 3. 並行執行各站台查詢
            site_tasks = []
            for site, page in zip(selected_sites, pages):
                accounts = accounts_dict.get(site["name"], [])
                if not accounts:
                    continue

                async def process_one_site(site=site, page=page, accounts=accounts):
                    for account in accounts:
                        if account:
                            await fetch_data(page, site, str(account).strip())
                site_tasks.append(process_one_site())

            if site_tasks:
                await asyncio.gather(*site_tasks)

            root.after(0, finalize_save)

        # ====================== 儲存結果（已加入兩點修改） ======================
        def finalize_save():
            if not all_results:
                messagebox.showinfo("提示", "未查詢到任何資料。")
            else:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                default_filename = f"帳戶管理抓取_{timestamp}.xlsx"
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=default_filename,
                    filetypes=[("Excel files", "*.xlsx")],
                    title="另存為 Excel 檔案"
                )
                if file_path:
                    try:
                        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                            # === 固定工作表順序 ===
                            site_order = ["TC", "TF", "TS", "SY", "FL", "WX", "XC", "XH", "CJ", "CY", "YD"]

                            for site_name in site_order:
                                if site_name in all_results and all_results[site_name]:
                                    dfs = all_results[site_name]
                                    combined_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

                                    if not combined_df.empty:
                                        # === 新增最左側「平台」欄位 ===
                                        combined_df.insert(0, "平台", site_name)

                                        combined_df.to_excel(writer, sheet_name=site_name, index=False)

                                        # 調整欄位寬度（包含新欄位）
                                        worksheet = writer.sheets[site_name]
                                        for col_idx, column in enumerate(combined_df.columns, 1):
                                            max_length = 0
                                            header_length = sum(2 if ord(char) > 127 else 1 for char in str(column))
                                            max_length = max(max_length, header_length)
                                            for value in combined_df[column].astype(str):
                                                value_length = sum(2 if ord(char) > 127 else 1 for char in str(value))
                                                max_length = max(max_length, value_length)
                                            adjusted_width = min(max(max_length * 1.1, 10), 50)
                                            worksheet.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

                        messagebox.showinfo("成功", f"查詢完成！\n結果已儲存至：\n{file_path}")
                    except Exception as e:
                        messagebox.showerror("錯誤", f"匯出 Excel 失敗: {e}")

            # 是否繼續
            if messagebox.askyesno("繼續查詢", "是否要繼續查詢其他帳號？"):
                input_window.destroy()
                create_input_window()
            else:
                input_window.destroy()
                if callback:
                    callback()

        # ====================== 開始按鈕 ======================
        def start_task():
            try:
                from __main__ import app
                app.run_async(_async_process())
            except Exception as e:
                messagebox.showerror("錯誤", f"無法啟動任務: {e}")

        tk.Button(input_window, text="開始查詢", 
                  command=start_task,
                  bg="#2196F3", fg="white", font=("Arial", 11, "bold"), height=2, width=20).pack(pady=20)

    create_input_window()