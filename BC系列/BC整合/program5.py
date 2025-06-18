import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime, timedelta
import time
from playwright.sync_api import sync_playwright

def run_program_5(root, selected_sites, pages):
    def create_input_window():
        # 創建輸入視窗
        input_window = tk.Toplevel(root)
        input_window.title("直屬及下級盈虧查詢")
        input_window.geometry("400x350")

        # Excel 檔案選擇
        tk.Label(input_window, text="選擇包含帳號的 Excel 檔案：").pack(pady=5)
        excel_path = tk.StringVar()
        def select_excel():
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if path:
                excel_path.set(path)
            input_window.lift()  # 確保視窗保持在前台
        tk.Button(input_window, text="瀏覽", command=select_excel).pack(pady=5)
        tk.Label(input_window, textvariable=excel_path).pack(pady=5)

        # 日期選擇
        # 預設開始日期和結束日期都為今天
        today = datetime.today()

        tk.Label(input_window, text="選擇查詢起始日期：").pack(pady=5)
        start_date_entry = DateEntry(input_window, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        start_date_entry.set_date(today)  # 設置預設值為今天
        start_date_entry.pack(pady=5)

        tk.Label(input_window, text="選擇查詢結束日期：").pack(pady=5)
        end_date_entry = DateEntry(input_window, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        end_date_entry.set_date(today)  # 設置預設值為今天
        end_date_entry.pack(pady=5)

        # 儲存查詢結果
        all_results = {}

        def fetch_data(page, site, account, start_date, end_date, superior, is_first_query):
            # 根據站台類型導航並填入日期
            if is_first_query:
                max_retries = 5
                for attempt in range(max_retries):
                    try:
                        if site["name"] in ["TC", "TF"]:
                            page.goto(f"{site['url']}TeamProfitReport", wait_until="networkidle")
                        else:
                            page.goto(f"{site['url']}ProfitReport", wait_until="networkidle")
                        page.wait_for_load_state("domcontentloaded")
                        time.sleep(1)  # 增加等待時間
                        break
                    except Exception as e:
                        print(f"導航站台 {site['name']} 嘗試 {attempt + 1}/{max_retries} 時出错: {e}")
                        if attempt == max_retries - 1:
                            continue
                        time.sleep(1)  # 重試間隔

                # 填入日期
                if site["name"] in ["TC", "TF"]:
                    start_datetime = datetime.strptime(start_date, '%Y-%m-%d')
                    end_datetime = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)
                    start_date = start_datetime.strftime('%Y/%m/%d 03:00')
                    end_date = end_datetime.strftime('%Y/%m/%d 03:00')
                    page.fill('#StartTime', start_date)
                    page.fill('#EndTime', end_date)
                    page.click('span.input-group-addon.add-on:has(span.glyphicon-calendar)', timeout=5000)
                    page.fill('#StartTime', start_date)
                    page.click('span.input-group-addon.add-on:has(span.glyphicon-calendar)', timeout=5000)
                    page.fill('#EndTime', end_date)
                else:
                    page.fill('#StartTime', start_date)
                    page.fill('#EndTime', end_date)
            else:
                # 清除帳號欄位
                if site["name"] in ["TC", "TF"]:
                    page.fill("#LoginID", "")
                else:
                    page.fill("#LoginId", "")

            # 填入帳號並查詢
            if site["name"] in ["TC", "TF"]:
                page.fill("#LoginID", account)
            else:
                page.fill("#LoginId", account)
            if site["name"] in ["TC", "TF"]:
                page.click("#querybutton")
            else:
                page.click("#SearchForm > div:nth-child(5) > div > button.btn.btn-primary")
            page.wait_for_load_state("networkidle")
            time.sleep(2)  # 等待查詢結果載入

            # 抓取個人報表結果
            if site["name"] in ["TC", "TF"]:
                personal_table_selector = "#ProfitTable > tbody"
            else:
                personal_table_selector = "#Lottery > div:nth-child(2) > div.card-body > div.text-note > div:nth-child(5) > table > tbody"
            personal_rows = page.query_selector_all(f"{personal_table_selector} >> tr")
            if personal_rows:
                for row in personal_rows:
                    cols = row.query_selector_all("td")
                    if cols:  # 確保有列
                        actual_columns = len(cols)
                        # 動態生成列名和數據
                        columns = ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利', '平台服务费']
                        used_columns = columns[:actual_columns] if actual_columns <= len(columns) else [f"列_{i+1}" for i in range(actual_columns)]
                        result = {"直屬": superior, "帳號": account}
                        for i, col in enumerate(cols):
                            if i < len(used_columns):
                                result[used_columns[i]] = float(col.inner_text().replace(",", "") or 0)
                        if f"{site['name']}_個人報表" not in all_results:
                            all_results[f"{site['name']}_個人報表"] = []
                        all_results[f"{site['name']}_個人報表"].append(result)

            # 檢查是否有下級 table
            if site["name"] in ["TC", "TF"]:
                subordinate_table = page.query_selector("#TeamMemberProfitTable")
                subordinate_accounts = []
                if subordinate_table:
                    rows = page.query_selector_all("#TeamMemberProfitTable > tbody > tr")
                    for row in rows:
                        account_cell = row.query_selector("td:nth-child(1) > a")
                        if account_cell:
                            subordinate_account = account_cell.inner_text().strip()
                            if subordinate_account:
                                subordinate_accounts.append(subordinate_account)
            else:
                subordinate_table = page.query_selector("#MemberDataTable")
                subordinate_accounts = []
                if subordinate_table:
                    rows = page.query_selector_all("#MemberDataTable > tbody > tr")
                    for row in rows:
                        account_cell = row.query_selector("td:nth-child(1)")
                        if account_cell:
                            subordinate_account = account_cell.inner_text().strip()
                            if subordinate_account:
                                subordinate_accounts.append(subordinate_account)

            # 遞迴查詢下級帳號
            for subordinate_account in subordinate_accounts:
                fetch_data(page, site, subordinate_account, start_date, end_date, superior, False)

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

            all_results.clear()  # 清空之前的結果
            for site, page in zip(selected_sites, pages):
                accounts = accounts_dict.get(site["name"], [])
                is_first_query = True  # 標記是否為第一次查詢
                for account in accounts:
                    fetch_data(page, site, account, start_date, end_date, account, is_first_query)
                    is_first_query = False  # 後續查詢設為 False

            # 將結果存入 Excel，使用另存為對話框並分工作頁
            if all_results:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                default_filename = f"直屬及下級盈虧_{timestamp}.xlsx"
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=default_filename,
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    title="另存為 Excel 檔案"
                )
                if file_path:
                    try:
                        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                            for site_name, results in all_results.items():
                                if results:  # 確保有數據
                                    df = pd.DataFrame(results)
                                    sheet_name = site_name.replace("_個人報表", "")  # 移除 "_個人報表" 後綴作為工作頁名
                                    df.to_excel(writer, sheet_name=sheet_name, index=False)
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
                input_window.destroy()  # 關閉當前視窗
                create_input_window()  # 重新打開輸入視窗
            else:
                input_window.destroy()  # 關閉視窗

        # 確認按鈕
        tk.Button(input_window, text="確認查詢", command=process_accounts).pack(pady=10)

    # 首次創建輸入視窗
    create_input_window()

# 註冊 program5 到主程式（需在 main.py 中新增按鈕）
if __name__ == "__main__":
    root = tk.Tk()
    run_program_5(root, [], [])
    root.mainloop()