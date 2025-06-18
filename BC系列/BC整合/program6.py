import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime, timedelta
import time
from playwright.sync_api import sync_playwright

def run_program_6(root, selected_sites, pages):
    def create_input_window():
        # 創建輸入視窗
        input_window = tk.Toplevel(root)
        input_window.title("招商觀察")
        input_window.geometry("400x300")
        input_window.transient(root)  # 設置為依附於主視窗
        input_window.grab_set()       # 確保輸入視窗獲得焦點

        # Excel 檔案選擇
        tk.Label(input_window, text="選擇包含帳號的 Excel 檔案：").pack(pady=5)
        excel_path = tk.StringVar()
        tk.Button(input_window, text="瀏覽", command=lambda: excel_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]))).pack(pady=5)
        tk.Label(input_window, textvariable=excel_path).pack(pady=5)

        # 日期選擇
        today = datetime.today()

        tk.Label(input_window, text="選擇查詢起始日期：").pack(pady=5)
        start_date_entry = DateEntry(input_window, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        start_date_entry.set_date(today)
        start_date_entry.pack(pady=5)

        tk.Label(input_window, text="選擇查詢結束日期：").pack(pady=5)
        end_date_entry = DateEntry(input_window, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        end_date_entry.set_date(today)
        end_date_entry.pack(pady=5)

        # 儲存查詢結果
        all_results = {}

        def fetch_data_lottery(page, site, account, start_time, end_time):
            try:
                page.wait_for_selector('#LoginId', state="visible", timeout=30000)
                page.fill('#LoginId', account)
                page.locator('button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]').first.click()
                time.sleep(1)
                page.wait_for_selector('tbody[data-bind="with: TeamStatistics"]', timeout=50000)
                table_data = page.evaluate("""
                    () => {
                        const rows = Array.from(document.querySelectorAll('tbody[data-bind="with: TeamStatistics"] tr'));
                        return rows.map(row => {
                            const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                            return cells;
                        });
                    }
                """)
                if table_data:
                    actual_columns = len(table_data[0])
                    print(f"站台 {site['name']}，帳號 {account} 表格列數: 實際 {actual_columns} 列")
                    print(f"抓取到第一筆資料: {table_data[0]}")
                    if actual_columns == 7:
                        columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款']
                    elif actual_columns == 8:
                        columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
                    else:
                        columns = [f"列_{i+1}" for i in range(actual_columns)]
                    df_data = [[account] + row for row in table_data]
                    df = pd.DataFrame(df_data, columns=['帳號'] + columns)
                    if site["name"] not in all_results:
                        all_results[site["name"]] = []
                    all_results[site["name"]].append(df)
                else:
                    print(f"站台 {site['name']}，帳號 {account} 彩票查詢沒有抓取到資料。")
            except Exception as e:
                print(f"處理站台 {site['name']}，帳號 {account} 彩票查詢時出错: {e}")

        def fetch_data_external(page, site, account, start_time, end_time):
            try:
                page.locator('#external-tab').click()
                page.wait_for_selector('#External', state="visible", timeout=30000)
                page.wait_for_load_state("networkidle")
                time.sleep(2)

                page.wait_for_selector('#External #LoginId', state="visible", timeout=30000)
                print(f"找到 #External #LoginId 元素: {page.locator('#External #LoginId').count()} 個")
                page.fill('#External #LoginId', account)
                page.wait_for_selector('#External #StartTime', state="visible", timeout=30000)
                print(f"填充 #External #StartTime: {start_time}")
                page.fill('#External #StartTime', start_time)
                page.wait_for_selector('#External #EndTime', state="visible", timeout=30000)
                print(f"填充 #External #EndTime: {end_time}")
                page.fill('#External #EndTime', end_time)

                page.wait_for_selector('#External #SearchForm > div:nth-child(5) > div > button.btn.btn-primary', state="visible", timeout=30000)
                print("點擊查詢按鈕")
                page.click('#External #SearchForm > div:nth-child(5) > div > button.btn.btn-primary')
                time.sleep(2)

                page.wait_for_function("document.querySelector('#External .table-responsive table tbody tr') !== null", timeout=50000)
                external_data = page.evaluate("""
                    () => {
                        const rows = Array.from(document.querySelectorAll('#External .table-responsive table tbody tr'));
                        return rows.filter(row => {
                            const td = row.querySelector('td[data-bind="text: GameTypeName"]');
                            return td && td.innerText.trim() === "总和";
                        }).map(row => {
                            const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                            return cells;
                        })[0] || [];
                    }
                """)
                if external_data and len(external_data) > 0:
                    actual_columns = len(external_data)
                    print(f"站台 {site['name']}，帳號 {account} 外接電子表格列數: 實際 {actual_columns} 列")
                    print(f"抓取到第一筆外接電子資料: {external_data}")
                    if actual_columns == 8:
                        external_columns = ['项目', '总投注', '有效投注', '总奖金', '总盈亏', '总返点', '总活动', '结果']
                    else:
                        external_columns = [f"列_{i+1}" for i in range(actual_columns)]
                    df_data = [[account] + external_data]
                    external_df = pd.DataFrame(df_data, columns=['帳號'] + external_columns)
                    if f"{site['name']}_真人電子" not in all_results:
                        all_results[f"{site['name']}_真人電子"] = []
                    all_results[f"{site['name']}_真人電子"].append(external_df)
                else:
                    print(f"站台 {site['name']}，帳號 {account} 外接電子查詢沒有抓取到資料。")
            except Exception as e:
                print(f"處理站台 {site['name']}，帳號 {account} 外接電子查詢時出错: {e}")

        def process_accounts():
            excel_file = excel_path.get()
            if not excel_file:
                messagebox.showerror("錯誤", "請選擇 Excel 檔案！")
                return

            start_date = start_date_entry.get()
            end_date = end_date_entry.get()

            if not start_date or not end_date:
                messagebox.showerror("錯誤", "請選擇查詢日期！")
                return

            # 將字符串日期轉換為 datetime 對象並格式化
            start_time = datetime.strptime(start_date, '%Y-%m-%d').strftime('%Y/%m/%d')
            end_time = datetime.strptime(end_date, '%Y-%m-%d').strftime('%Y/%m/%d')

            # 清除舊的結果
            all_results.clear()

            # 從 Excel 讀取帳號
            accounts_dict = {}
            try:
                excel_data = pd.read_excel(excel_file, sheet_name=None)
                for site_name in [s["name"] for s in selected_sites]:
                    if site_name in excel_data:
                        df = excel_data[site_name]
                        print(f"工作表 {site_name} 資料: {df.to_string(index=True)}")
                        print(f"工作表 {site_name} 列名: {df.columns.tolist()}")
                        if df.shape[1] == 1:
                            accounts = df['查詢帳號'].iloc[0:].dropna().tolist()
                            print(f"提取前帳號 {site_name}: {df['查詢帳號'].iloc[0:].tolist()}")
                            print(f"提取後帳號 {site_name}: {accounts}")
                            if accounts:
                                accounts_dict[site_name] = accounts
                            else:
                                accounts_dict[site_name] = []
                                print(f"警告：工作表 {site_name} 無有效帳號資料，跳過帳號讀取。")
                        elif df.shape[1] > 1:
                            accounts_dict[site_name] = df.iloc[:, 1].dropna().tolist()
                        else:
                            accounts_dict[site_name] = []
                            print(f"警告：工作表 {site_name} 格式無效，跳過帳號讀取。")
                    else:
                        accounts_dict[site_name] = []
                        print(f"警告：工作表 {site_name} 未在 Excel 中找到，跳過該站台。")
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取 Excel 失敗: {e}")
                return

            # 填入時間並導航到指定頁面
            for site, page in zip(selected_sites, pages):
                max_retries = 5
                for attempt in range(max_retries):
                    try:
                        page.goto(f"{site['url']}", wait_until="networkidle")
                        page.wait_for_load_state("domcontentloaded")
                        time.sleep(1)  # 增加等待時間
                        if site["name"] in ["TC", "TF"]:
                            start_datetime = datetime.strptime(start_date, '%Y-%m-%d')
                            end_datetime = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)
                            start_time = start_datetime.strftime('%Y/%m/%d 03:00')
                            end_time = end_datetime.strftime('%Y/%m/%d 03:00')
                            page.goto(f"{site['url']}TeamProfitReport", wait_until="networkidle")
                        else:
                            start_time = datetime.strptime(start_date, '%Y-%m-%d').strftime('%Y/%m/%d')
                            end_time = datetime.strptime(end_date, '%Y-%m-%d').strftime('%Y/%m/%d')
                            page.goto(f"{site['url']}ProfitReport", wait_until="networkidle")
                        page.wait_for_load_state("networkidle")
                        time.sleep(2)
                        break
                    except Exception as e:
                        print(f"導航站台 {site['name']} 嘗試 {attempt + 1}/{max_retries} 時出错: {e}")
                        if attempt == max_retries - 1:
                            continue
                        time.sleep(1)  # 重試間隔

                page.wait_for_selector('#StartTime', timeout=30000)
                page.wait_for_selector('#EndTime', timeout=30000)
                if site["name"] in ["TC", "TF"]:
                    page.fill('#StartTime', start_time)
                    page.fill('#EndTime', end_time)
                    page.click('span.input-group-addon.add-on:has(span.glyphicon-calendar)', timeout=5000)
                    page.fill('#StartTime', start_time)
                    page.click('span.input-group-addon.add-on:has(span.glyphicon-calendar)', timeout=5000)
                    page.fill('#EndTime', end_time)
                else:
                    page.fill('#StartTime', start_time)
                    page.fill('#EndTime', end_time)

            # 先執行所有帳號的彩票查詢
            for site, page in zip(selected_sites, pages):
                accounts = accounts_dict.get(site["name"], [])
                if site["name"] in ["TC", "TF"]:
                    for i, account in enumerate(accounts):
                        if i > 0:
                            page.fill('#LoginID', '')  # 清空帳號輸入框
                            time.sleep(1)  # 等待 1 秒
                        page.fill('#LoginID', account)
                        page.click('#querybutton')
                        page.wait_for_selector('#TeamProfitTable', timeout=50000)
                        time.sleep(1)
                        table_data = page.evaluate("""
                            () => {
                                const rows = Array.from(document.querySelectorAll('#TeamProfitTable tbody tr'));
                                return rows.map(row => {
                                    const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                                    return cells;
                                });
                            }
                        """)
                        if table_data:
                            actual_columns = len(table_data[0])
                            print(f"站台 {site['name']}，帳號 {account} 表格列數: 實際 {actual_columns} 列")
                            print(f"抓取到第一筆資料: {table_data[0]}")
                            if actual_columns == 7:
                                columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款']
                            elif actual_columns == 8:
                                columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
                            else:
                                columns = [f"列_{i+1}" for i in range(actual_columns)]
                            df_data = [[account] + row for row in table_data]
                            df = pd.DataFrame(df_data, columns=['帳號'] + columns)
                            if site["name"] not in all_results:
                                all_results[site["name"]] = []
                            all_results[site["name"]].append(df)
                        else:
                            print(f"站台 {site['name']}，帳號 {account} 彩票查詢沒有抓取到資料。")
                else:
                    for account in accounts:
                        fetch_data_lottery(page, site, account, start_time, end_time)

            # 再執行所有帳號的外接電子查詢
            for site, page in zip(selected_sites, pages):
                accounts = accounts_dict.get(site["name"], [])
                if site["name"] in ["TC", "TF"]:
                    for i, account in enumerate(accounts):
                        if i > 0:
                            page.fill('#LoginID', '')  # 清空帳號輸入框
                            time.sleep(1)  # 等待 1 秒
                        page.goto(f"{site['url']}TeamProfitReport", wait_until="networkidle")
                        page.wait_for_load_state("networkidle")
                        time.sleep(2)
                        page.locator('#GameTypeId').click()
                        page.locator('a[href="#"]:has-text("真人电子")').click()
                        page.wait_for_load_state("networkidle")
                        time.sleep(2)
                        page.fill('#LoginID', account)
                        page.click('#querybutton')
                        page.wait_for_selector('#TeamProfitReportForExternalGame', timeout=50000)
                        time.sleep(1)
                        external_data = page.evaluate("""
                            () => {
                                const rows = Array.from(document.querySelectorAll('#TeamProfitReportForExternalGame tbody tr'));
                                return rows.filter(row => {
                                    const span = row.querySelector('span[data-bind="text: Project"]');
                                    return span && span.innerText.trim() === "总和";
                                }).map(row => {
                                    const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                                    return cells;
                                })[0] || [];
                            }
                        """)
                        if external_data:
                            actual_columns = len(external_data)
                            print(f"站台 {site['name']}，帳號 {account} 外接電子表格列數: 實際 {actual_columns} 列")
                            print(f"抓取到第一筆外接電子資料: {external_data}")
                            if actual_columns == 5:
                                external_columns = ['项目', '总投注', '有效投注', '总奖金', '总活动']
                            elif actual_columns == 6:
                                external_columns = ['项目', '总投注', '有效投注', '总奖金', '总活动', '总盈亏']
                            else:
                                external_columns = [f"列_{i+1}" for i in range(actual_columns)]
                            df_data = [[account] + external_data]
                            external_df = pd.DataFrame(df_data, columns=['帳號'] + external_columns)
                            if f"{site['name']}_真人電子" not in all_results:
                                all_results[f"{site['name']}_真人電子"] = []
                            all_results[f"{site['name']}_真人電子"].append(external_df)
                        else:
                            print(f"站台 {site['name']}，帳號 {account} 外接電子查詢沒有抓取到資料。")
                else:
                    for account in accounts:
                        fetch_data_external(page, site, account, start_time, end_time)

            # 將結果匯出為單一 Excel 檔案
            if all_results:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                default_filename = f"招商觀察_{timestamp}.xlsx"
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=default_filename,
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    title="另存為 Excel 檔案"
                )
                if file_path:
                    try:
                        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                            for site_name, dfs in all_results.items():
                                combined_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                                if not combined_df.empty:
                                    combined_df.to_excel(writer, sheet_name=site_name, index=False)
                        messagebox.showinfo("成功", f"查詢完成，結果已存至 {file_path}")
                    except Exception as e:
                        print(f"匯出 Excel 時出错: {e}")
                        messagebox.showerror("錯誤", f"匯出 Excel 失敗: {e}")
                else:
                    messagebox.showinfo("提示", "已取消儲存 Excel 檔案。")
            else:
                messagebox.showinfo("提示", "未查詢到任何資料。")

            # 詢問是否繼續查詢
            if messagebox.askyesno("繼續查詢", "是否繼續查詢？"):
                input_window.destroy()
                create_input_window()
            else:
                input_window.destroy()

        # 確認按鈕
        tk.Button(input_window, text="確認查詢", command=process_accounts).pack(pady=10)

    # 首次創建輸入視窗
    create_input_window()

# 註冊 program6 到主程式（需在 main.py 中新增按鈕）
if __name__ == "__main__":
    root = tk.Tk()
    run_program_6(root, [], [])
    root.mainloop()