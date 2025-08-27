import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
from datetime import datetime
import time
from openpyxl.utils import get_column_letter

def run_program_7(root, selected_sites, pages):
    def create_input_window():
        # 創建輸入視窗
        input_window = tk.Toplevel(root)
        input_window.title("帳戶管理抓取")
        input_window.geometry("400x200")
        input_window.transient(root)  # 設置為依附於主視窗
        input_window.grab_set()       # 確保輸入視窗獲得焦點

        # Excel 檔案選擇
        tk.Label(input_window, text="選擇包含帳號的 Excel 檔案：").pack(pady=5)
        excel_path = tk.StringVar()
        tk.Button(input_window, text="瀏覽", command=lambda: excel_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]))).pack(pady=5)
        tk.Label(input_window, textvariable=excel_path).pack(pady=5)

        # 儲存查詢結果
        all_results = {}

        def fetch_data(page, site, account):
            try:
                # 清空搜尋框（非首次查詢）
                if site["name"] in ["TC", "TF"]:
                    page.fill('#SearchStr', '')
                else:
                    page.fill('#Account', '')
                time.sleep(1)  # 等待清空

                # 填入帳號
                if site["name"] in ["TC", "TF"]:
                    page.fill('#SearchStr', account)
                else:
                    page.fill('#Account', account)

                # 點擊搜尋按鈕
                if site["name"] in ["TC", "TF"]:
                    page.click('#Searchbtn')
                else:
                    page.click('#searchBtn')
                time.sleep(2)  # 等待查詢結果載入
                page.wait_for_selector('#playersTables', timeout=50000)

                # 抓取表格資料
                table_data = page.evaluate("""
                    () => {
                        const rows = Array.from(document.querySelectorAll('#playersTables tbody tr'));
                        return rows.map(row => {
                            const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText.trim());
                            return cells;
                        });
                    }
                """)
                if table_data:
                    actual_columns = len(table_data[0])
                    print(f"站台 {site['name']}，帳號 {account} 表格列數: 實際 {actual_columns} 列")
                    print(f"抓取到資料: {table_data}")
                    if site["name"] in ["TC", "TF"]:
                        columns = ['帐号', '昵称', '下级', '时时彩奖金', '类型', '余额', '团队余额', '诚信率 (%)', '创建时间', '最后登录时间', '创建来源', '登录', '状态', '功能']
                    else:
                        columns = ['帐号', '昵称', '下级', '时时彩奖金', '类型', '余额', '团队余额', '创建时间', '最后登录时间', '最后下注时间', '创建來源', '登录', '状态', '功能']
                    # 使用實際列數調整 columns
                    used_columns = columns[:actual_columns] if actual_columns <= len(columns) else [f"列_{i+1}" for i in range(actual_columns)]
                    df_data = [row for row in table_data]
                    df = pd.DataFrame(df_data, columns=used_columns)
                    if site["name"] not in all_results:
                        all_results[site["name"]] = []
                    all_results[site["name"]].append(df)
                else:
                    print(f"站台 {site['name']}，帳號 {account} 查詢沒有抓取到資料。")
            except Exception as e:
                print(f"處理站台 {site['name']}，帳號 {account} 查詢時出错: {e}")

        def process_accounts():
            excel_file = excel_path.get()
            if not excel_file:
                messagebox.showerror("錯誤", "請選擇 Excel 檔案！")
                return

            # 清空舊的結果
            all_results.clear()

            # 導航到指定頁面
            for site, page in zip(selected_sites, pages):
                max_retries = 5
                for attempt in range(max_retries):
                    try:
                        if site["name"] in ["TC", "TF"]:
                            page.goto(f"{site['url']}AccountManagement/Players", wait_until="networkidle")
                        else:
                            page.goto(f"{site['url']}Player", wait_until="networkidle")
                        page.wait_for_load_state("domcontentloaded")
                        time.sleep(1)  # 增加等待時間
                        break
                    except Exception as e:
                        print(f"導航站台 {site['name']} 嘗試 {attempt + 1}/{max_retries} 時出错: {e}")
                        if attempt == max_retries - 1:
                            continue
                        time.sleep(1)  # 重試間隔

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

            # 執行所有帳號的查詢
            for site, page in zip(selected_sites, pages):
                accounts = accounts_dict.get(site["name"], [])
                for i, account in enumerate(accounts):
                    fetch_data(page, site, account)

            # 將結果匯出為單一 Excel 檔案，並調整欄位寬度
            if all_results:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                default_filename = f"帳戶管理抓取_{timestamp}.xlsx"
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
                                    # 調整欄位寬度根據最長內容
                                    worksheet = writer.sheets[site_name]
                                    for col_idx, column in enumerate(combined_df.columns, 1):
                                        max_length = 0
                                        # 計算標題長度
                                        header_length = sum(2 if ord(char) > 127 else 1 for char in str(column))
                                        max_length = max(max_length, header_length)
                                        # 計算資料長度
                                        for value in combined_df[column].astype(str):
                                            value_length = sum(2 if ord(char) > 127 else 1 for char in str(value))
                                            max_length = max(max_length, value_length)
                                        # 設置欄位寬度，添加padding並限制範圍
                                        adjusted_width = min(max(max_length * 1.1, 10), 50)  # 乘以1.1增加一點padding
                                        worksheet.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
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

# 註冊 program7 到主程式（需在 main.py 中新增按鈕）
if __name__ == "__main__":
    root = tk.Tk()
    run_program_7(root, [], [])
    root.mainloop()