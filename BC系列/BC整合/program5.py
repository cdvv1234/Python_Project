import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime
import time

def run_program_5(root, selected_sites, pages):
    def create_input_window():
        # 創建輸入視窗
        input_window = tk.Toplevel(root)
        input_window.title("直屬及下級盈虧查詢")
        input_window.geometry("400x350")

        # 帳號輸入框
        tk.Label(input_window, text="請輸入要查詢的帳號（每行一個）：").pack(pady=5)
        account_text = tk.Text(input_window, height=5, width=30)
        account_text.pack(pady=5)

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
        all_results = []

        def fetch_data(page, site, account, start_date, end_date, superior, is_first_query):
            # 如果是第一次查詢，導航並填入日期
            if is_first_query:
                page.goto(site["url"] + "ProfitReport")
                page.wait_for_load_state("networkidle")
                # 填入起始日期和結束日期
                page.fill("#StartTime", start_date)
                page.fill("#EndTime", end_date)
            else:
                # 清除帳號欄位
                page.fill("#LoginId", "")

            # 填入帳號並查詢
            page.fill("#LoginId", account)
            page.click("#SearchForm > div:nth-child(5) > div > button.btn.btn-primary")
            page.wait_for_load_state("networkidle")
            time.sleep(2)  # 等待查詢結果載入

            # 抓取個人報表結果
            personal_table_selector = "#Lottery > div:nth-child(2) > div.card-body > div.text-note > div:nth-child(5) > table > tbody"
            personal_rows = page.query_selector_all(f"{personal_table_selector} >> tr")
            if personal_rows:
                for row in personal_rows:
                    cols = row.query_selector_all("td")
                    if len(cols) >= 9:  # 確保有足夠的欄位
                        result = {
                            "直屬": superior,
                            "帳號": account,
                            "总投注": float(cols[0].inner_text().replace(",", "") or 0),
                            "总奖金": float(cols[1].inner_text().replace(",", "") or 0),
                            "总返点": float(cols[2].inner_text().replace(",", "") or 0),
                            "总活动": float(cols[3].inner_text().replace(",", "") or 0),
                            "总盈亏": float(cols[4].inner_text().replace(",", "") or 0),
                            "总充值": float(cols[5].inner_text().replace(",", "") or 0),
                            "总提款": float(cols[6].inner_text().replace(",", "") or 0),
                            "总红利": float(cols[7].inner_text().replace(",", "") or 0),
                            "平台服务费": float(cols[8].inner_text().replace(",", "") or 0)
                        }
                        all_results.append(result)

            # 檢查是否有下級 table
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
            accounts = account_text.get("1.0", tk.END).strip().split("\n")
            accounts = [acc.strip() for acc in accounts if acc.strip()]  # 移除空行
            start_date = start_date_entry.get()
            end_date = end_date_entry.get()

            if not accounts:
                messagebox.showerror("錯誤", "請輸入至少一個帳號！")
                return

            all_results.clear()  # 清空之前的結果
            for site, page in zip(selected_sites, pages):
                is_first_query = True  # 標記是否為第一次查詢
                for account in accounts:
                    fetch_data(page, site, account, start_date, end_date, account, is_first_query)
                    is_first_query = False  # 後續查詢設為 False

            # 將結果存入 Excel，使用另存為對話框
            if all_results:
                df = pd.DataFrame(all_results)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                default_filename = f"profit_report_{timestamp}.xlsx"
                
                # 跳出另存為對話框
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=default_filename,
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    title="另存為 Excel 檔案"
                )
                
                if file_path:  # 使用者選擇了儲存路徑
                    df.to_excel(file_path, index=False)
                    messagebox.showinfo("成功", f"查詢完成，結果已存至 {file_path}")
                else:  # 使用者取消儲存
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