import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime, timedelta
import time
from playwright.sync_api import sync_playwright
import os

def run_program_6(root, selected_sites, pages):
    def create_input_window():
        # 創建輸入視窗
        input_window = tk.Toplevel(root)
        input_window.title("招商觀察")
        input_window.geometry("400x400")
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

        # 核選方塊
        tk.Label(input_window, text="選擇查詢項目：").pack(pady=5)
        check_vars = {
            "彩票": tk.BooleanVar(value=True),
            "真人電子": tk.BooleanVar(value=True),
            "體育": tk.BooleanVar(value=True),
            "棋牌": tk.BooleanVar(value=True)
        }
        for label in check_vars:
            tk.Checkbutton(input_window, text=label, variable=check_vars[label]).pack(anchor="w", padx=20)

        # 儲存查詢結果
        all_results = {}

        def save_emergency_backup():
            """將當前已收集的資料保存為緊急備份 Excel"""
            if all_results:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_filename = f"招商觀察_緊急備份_{timestamp}.xlsx"
                backup_filepath = os.path.join(os.getcwd(), backup_filename)
                try:
                    with pd.ExcelWriter(backup_filepath, engine="openpyxl") as writer:
                        for site_name, dfs in all_results.items():
                            combined_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                            if not combined_df.empty:
                                combined_df.to_excel(writer, sheet_name=site_name, index=False)
                    print(f"緊急備份已保存至: {backup_filepath}")
                except Exception as e:
                    print(f"緊急備份保存失敗: {e}")
            else:
                print("無數據可保存為緊急備份")

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
                    print(f"站台 {site['name']}，帳號 {account}，項目 彩票 表格列數: 實際 {actual_columns} 列")
                    print(f"抓取到第一筆資料: {table_data[0]}")
                    if actual_columns == 8:
                        columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
                    elif actual_columns == 9:
                        columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利', '平台服务费']
                    else:
                        columns = [f"列_{i+1}" for i in range(actual_columns)]
                    df_data = [[account] + row for row in table_data]
                    df = pd.DataFrame(df_data, columns=['帳號'] + columns)
                    sheet_name = f"{site['name']}_彩票"
                    if sheet_name not in all_results:
                        all_results[sheet_name] = []
                    all_results[sheet_name].append(df)
                else:
                    print(f"站台 {site['name']}，帳號 {account}，項目 彩票 查詢沒有抓取到資料。")
            except Exception as e:
                print(f"站台 {site['name']}，帳號 {account}，項目 彩票 查詢時出错: {e}")
                save_emergency_backup()

        def fetch_data_external(page, site, account, start_time, end_time, category, category_value):
            try:
                page.wait_for_selector('#external-tab', state="visible", timeout=30000)
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

                # 選擇項目
                page.select_option('#External #Category', category_value)
                page.wait_for_selector('#External #SearchForm > div:nth-child(5) > div > button.btn.btn-primary', state="visible", timeout=30000)
                print(f"站台 {site['name']}，帳號 {account}，項目 {category} 點擊查詢按鈕")
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
                    print(f"站台 {site['name']}，帳號 {account}，項目 {category} 表格列數: 實際 {actual_columns} 列")
                    print(f"抓取到第一筆外接資料: {external_data}")
                    if actual_columns == 6:
                        external_columns = ['项目', '总投注', '有效投注', '总奖金', '总活动', '总盈亏']
                    elif actual_columns == 8:
                        external_columns = ['项目', '总投注', '有效投注', '总奖金', '总盈亏', '总返点', '总活动', '结果']
                    else:
                        external_columns = [f"列_{i+1}" for i in range(actual_columns)]
                    df_data = [[account] + external_data[:len(external_columns)]]
                    external_df = pd.DataFrame(df_data, columns=['帳號'] + external_columns)
                    sheet_name = f"{site['name']}_{category}"
                    if sheet_name not in all_results:
                        all_results[sheet_name] = []
                    all_results[sheet_name].append(external_df)
                else:
                    print(f"站台 {site['name']}，帳號 {account}，項目 {category} 查詢沒有抓取到資料。")
            except Exception as e:
                print(f"站台 {site['name']}，帳號 {account}，項目 {category} 查詢時出错: {e}")
                save_emergency_backup()

        def fetch_data_tc_tf(page, site, account, start_time, end_time, category, category_value):
            try:
                # 等待下拉選單按鈕可見
                page.wait_for_selector('#GameTypeId', state="visible", timeout=30000)
                # 點擊下拉選單按鈕以展開選項
                page.locator('#GameTypeId').click()
                # 等待下拉選單可見
                page.wait_for_selector('div.btn-group ul.dropdown-menu', state="visible", timeout=30000)
                # 輸出下拉選單選項以診斷
                options = page.evaluate("""
                    () => {
                        const items = Array.from(document.querySelectorAll('div.btn-group ul.dropdown-menu li.GameType'));
                        return items.map(item => ({
                            text: item.querySelector('a').innerText.trim(),
                            value: item.getAttribute('value')
                        }));
                    }
                """)
                print(f"站台 {site['name']}，#GameTypeId 下拉選單選項: {options}")
                # 檢查項目是否在下拉選單中
                if not any(option['value'] == category_value for option in options):
                    print(f"站台 {site['name']}，項目 {category} 在下拉選單中不存在，跳過查詢")
                    return
                # 選擇項目
                page.locator(f'div.btn-group ul.dropdown-menu li.GameType[value="{category_value}"] a').click()
                page.wait_for_load_state("networkidle")
                time.sleep(1)
                page.fill('#LoginID', '')
                page.fill('#LoginID', account)
                page.click('#querybutton')
                # 根據項目選擇不同的結果元素和表格
                result_selector = '#teamResult' if category == "彩票" else '#teamResultForExternalGame'
                table_selector = '#TeamProfitTable' if category == "彩票" else '#TeamProfitTableForExternalGame'
                # 等待查詢結果包含帳號
                page.wait_for_function(
                    f"""
                    (account) => {{
                        const resultElement = document.querySelector('{result_selector}');
                        return resultElement && resultElement.innerText.includes(account);
                    }}
                    """,
                    arg=account,
                    timeout=50000
                )
                time.sleep(1)  # 額外等待確保表格載入
                # 根據項目選擇抓取邏輯
                if category == "彩票":
                    # 彩票抓取所有行
                    table_data = page.evaluate(f"""
                        () => {{
                            const rows = Array.from(document.querySelectorAll('{table_selector} tbody tr'));
                            return rows.map(row => {{
                                const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                                return cells;
                            }});
                        }}
                    """)
                    if table_data:
                        actual_columns = len(table_data[0])
                        print(f"站台 {site['name']}，帳號 {account}，項目 {category} 表格列數: 實際 {actual_columns} 列")
                        print(f"抓取到第一筆資料: {table_data[0]}")
                        if actual_columns == 7:
                            columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款']
                        elif actual_columns == 8:
                            columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
                        else:
                            columns = [f"列_{i+1}" for i in range(actual_columns)]
                    else:
                        print(f"站台 {site['name']}，帳號 {account}，項目 {category} 查詢沒有抓取到資料。")
                        return
                else:
                    # 真人電子、體育、棋牌只抓取“总和”行
                    table_data = page.evaluate(f"""
                        () => {{
                            const rows = Array.from(document.querySelectorAll('{table_selector} tbody tr'));
                            return rows.filter(row => {{
                                const span = row.querySelector('span[data-bind="text: Project"]');
                                return span && span.innerText.trim() === "总和";
                            }}).map(row => {{
                                const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                                return cells;
                            }})[0] || [];
                        }}
                    """)
                    if table_data:
                        table_data = [table_data]  # 轉為與彩票格式一致（列表的列表）
                        actual_columns = len(table_data[0])
                        print(f"站台 {site['name']}，帳號 {account}，項目 {category} 表格列數: 實際 {actual_columns} 列")
                        print(f"抓取到第一筆資料: {table_data[0]}")
                        # 固定使用要求的標題
                        columns = ['项目', '总投注', '有效投注', '总奖金', '总活动', '总盈亏']
                        # 調整數據以匹配標題
                        adjusted_data = []
                        for row in table_data:
                            if len(row) >= len(columns):
                                adjusted_row = row[:len(columns)]  # 取前 6 列
                            else:
                                adjusted_row = row + [''] * (len(columns) - len(row))  # 補空值
                            adjusted_data.append([account] + adjusted_row)
                        df = pd.DataFrame(adjusted_data, columns=['帳號'] + columns)
                        sheet_name = f"{site['name']}_{category}"
                        if sheet_name not in all_results:
                            all_results[sheet_name] = []
                        all_results[sheet_name].append(df)
                        return
                    else:
                        print(f"站台 {site['name']}，帳號 {account}，項目 {category} 查詢沒有抓取到資料。")
                        return
                df_data = [[account] + row for row in table_data]
                df = pd.DataFrame(df_data, columns=['帳號'] + columns)
                sheet_name = f"{site['name']}_{category}"
                if sheet_name not in all_results:
                    all_results[sheet_name] = []
                all_results[sheet_name].append(df)
            except Exception as e:
                print(f"站台 {site['name']}，帳號 {account}，項目 {category} 查詢時出错: {e}")
                save_emergency_backup()

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

            # 檢查是否有至少一個項目被勾選
            if not any(check_vars[category].get() for category in check_vars):
                messagebox.showerror("錯誤", "請至少選擇一個查詢項目！")
                return

            # 清除舊的結果
            all_results.clear()

            # 導航到指定頁面並填入時間
            for site, page in zip(selected_sites, pages):
                max_retries = 5
                for attempt in range(max_retries):
                    try:
                        if site["name"] in ["TC", "TF"]:
                            page.goto(f"{site['url']}TeamProfitReport", wait_until="networkidle")
                        else:
                            page.goto(f"{site['url']}ProfitReport", wait_until="networkidle")
                        page.wait_for_load_state("domcontentloaded")
                        time.sleep(2)
                        # 填入時間
                        page.wait_for_selector('#StartTime', timeout=30000)
                        page.wait_for_selector('#EndTime', timeout=30000)
                        if site["name"] in ["TC", "TF"]:
                            start_datetime = datetime.strptime(start_date, '%Y-%m-%d')
                            end_datetime = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)
                            start_time = start_datetime.strftime('%Y/%m/%d 03:00')
                            end_time = end_datetime.strftime('%Y/%m/%d 03:00')
                            page.fill('#StartTime', "")
                            page.fill('#EndTime', "")
                            page.fill('#StartTime', start_time)
                            page.fill('#EndTime', end_time)
                            print(f"站台 {site['name']} 設置時間: {start_time}, {end_time}")
                        else:
                            page.fill('#StartTime', start_date)
                            page.fill('#EndTime', end_date)
                        break
                    except Exception as e:
                        print(f"導航站台 {site['name']} 嘗試 {attempt + 1}/{max_retries} 時出错: {e}")
                        if attempt == max_retries - 1:
                            save_emergency_backup()
                            continue
                        time.sleep(1)

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
                print(f"讀取 Excel 失敗: {e}")
                save_emergency_backup()
                messagebox.showerror("錯誤", f"讀取 Excel 失敗: {e}")
                return

            # 定義 category 到 value 的映射
            category_map = {
                "彩票": "0",
                "真人電子": "1",
                "體育": "2",
                "棋牌": "3"
            }

            # 執行查詢
            try:
                for site, page in zip(selected_sites, pages):
                    accounts = accounts_dict.get(site["name"], [])
                    for category in check_vars:
                        if check_vars[category].get():  # 僅處理勾選的項目
                            for i, account in enumerate(accounts):
                                if site["name"] in ["TC", "TF"]:
                                    fetch_data_tc_tf(page, site, account, start_date, end_date, category, category_map[category])
                                else:
                                    if category == "彩票":
                                        fetch_data_lottery(page, site, account, start_date, end_date)
                                    else:
                                        fetch_data_external(page, site, account, start_date, end_date, category, category_map[category])
            except Exception as e:
                print(f"處理帳號時發生錯誤: {e}")
                save_emergency_backup()

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