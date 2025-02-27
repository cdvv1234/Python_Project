import tkinter as tk
from tkinter import filedialog, messagebox
from playwright.sync_api import sync_playwright
import pandas as pd
import time
from datetime import timedelta, datetime 

# 站點配置
sites = [
    {
        "name": "CJ",
        "url": "",
        "input_selector": "#LoginId",
        "button_selector": '#Search',
        "table_selector": '#SearchResult tbody',
        "columns": []
    }
]

# 讀取帳號、領取日期、報表日期(起)、報表日期(迄)
def read_account_date_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df[['帐号', '领取日期', '起始日期', '迄止日期']]

# 用來儲存每個帳號的投注與盈亏資料
previous_bet_amount = None  # 追蹤上一個帳號的 bet_amount
previous_profit_loss = None  # 追蹤上一個帳號的 profit_loss

# 查詢邏輯，將結果一次性儲存
def scrape_cj_site(site, accounts, start_time, end_time, page):
    global previous_bet_amount, previous_profit_loss

    results = []  # 用來存儲所有查詢結果
    processed_accounts = set()  # 用來記錄已經處理過的帳號+時間區間，防止重複

    for query in accounts:  
        # 生成一個 unique key，包含帳號和時間區間
        account_time_key = f"{query}_{start_time}_{end_time}"

        # 如果該帳號和時間區間已經處理過，則跳過
        if account_time_key in processed_accounts:
            print(f"帳號 {query} 時間區間 {start_time} - {end_time} 已經處理過，跳過這次查詢。")
            continue  # 跳過該帳號的處理

        try:
            # 將 Timestamp 轉換為字串格式
            start_time_str = (start_time + timedelta(seconds=30)).strftime('%Y/%m/%d %H:%M:%S') if isinstance(start_time, (pd.Timestamp, datetime)) else start_time
            end_time_str = end_time.strftime('%Y/%m/%d %H:%M:%S') if isinstance(end_time, (pd.Timestamp, datetime)) else end_time

            # 填入帳號、報表日期(起)、報表日期(迄)
            page.fill('#LoginId', query)  # 填入新帳號
            page.fill('#UpdateStartTime', start_time_str)  # 使用轉換後的字串
            page.fill('#UpdateEndTime', end_time_str)     # 使用轉換後的字串

            # 等待 #BetRecordSearchRange 可見
            page.wait_for_selector('#BetRecordSearchRange', state='visible', timeout=10000)

            # 等待查詢按鈕不再禁用
            page.wait_for_selector('#Search:not([disabled])', state='visible')  # 確保查詢按鈕不再禁用

            # 點擊查詢按鈕
            page.click('#Search', force=True)

            time.sleep(2)

            # 等待表格載入
            page.wait_for_selector('#SearchResult tbody', timeout=10000)

            # 檢查是否有警告對話框需要關閉
            if page.locator('#Warning-Dialog-CloseBtn').is_visible():
                page.click('#Warning-Dialog-CloseBtn')  # 點擊關閉按鈕關閉對話框
                print(f"帳號: {query} - 關閉警告對話框")

                # 重新按一次查詢按鈕，確保資料載入正確
                # 等待查詢按鈕不再禁用
                page.wait_for_selector('#Search:not([disabled])', state='visible')  # 確保查詢按鈕不再禁用
                page.click('#Search', force=True)
                print(f"帳號: {query} - 重新點擊查詢按鈕")

                # 等待資料表格重新載入
                page.wait_for_selector('#SearchResult tbody', timeout=10000)

            # 判斷表格中是否有包含 "查无资料" 的單元格
            empty_table = page.locator('#SearchResult > tbody > tr > td').first.inner_text()
            if "查无资料" in empty_table:
                bet_amount = '0'
                profit_loss = '0'
                print(f"帳號: {query} - 查無資料，設定投注金額和盈虧為 0")
            else:
                # 初始讀取投注和盈亏資料
                bet_amount = page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(11)").inner_text()
                profit_loss = page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(13)").inner_text()

                # 若查不到資料，則預設為 0
                bet_amount = bet_amount if bet_amount else '0'
                profit_loss = profit_loss if profit_loss else '0'

            print(f"檢查資料: {bet_amount}, {previous_bet_amount}, {profit_loss}, {previous_profit_loss}")

           # 確保第一次查詢時不進行重試邏輯，直接儲存資料
            if previous_bet_amount is None and previous_profit_loss is None:
                print(f"帳號: {query} 為第一次處理，直接儲存資料。")
            else:
                retries = 0
                while bet_amount == previous_bet_amount and profit_loss == previous_profit_loss:
                    retries += 1
                    if retries > 3:
                        print(f"帳號: {query} 查詢結果與上一帳號相同，重試次數達到上限，資料可能未更新")
                        break
                    print(f"帳號: {query} 查詢結果與上一帳號相同，進行重試...")
                    time.sleep(2)
                    page.click('#Search', force=True)  # 重新點擊查詢按鈕
                    page.wait_for_selector('#SearchResult tbody', timeout=10000)
                    bet_amount = page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(11)").inner_text()
                    profit_loss = page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(13)").inner_text()
                    bet_amount = bet_amount if bet_amount else '0'
                    profit_loss = profit_loss if profit_loss else '0'

            # 儲存結果到列表
            results.append([query, bet_amount, profit_loss, start_time, end_time_str])

            # 印出帳號、日期、投注、盈亏
            print(f"帳號: {query}, 起始日期: {start_time_str}, 迄止日期: {end_time_str}, 投注: {bet_amount}, 盈亏: {profit_loss}")

            # 記錄該帳號和時間區間已處理
            processed_accounts.add(account_time_key)

            # 更新上一個帳號的資料
            previous_bet_amount = bet_amount
            previous_profit_loss = profit_loss  # 這裡確保更新為當前的 bet_amount 和 profit_loss
            print(f"更新資料: {previous_bet_amount}, {previous_profit_loss}")

        except Exception as e:
            print(f"處理帳號 {query} 時發生錯誤: {e}")

    return results

# 儲存所有結果至 Excel，並合併原始資料與查詢結果
def save_results_to_excel(results, file_path, sheet_name, account_data):
    # 將查詢結果與原始帳號資料合併
    result_df = pd.DataFrame(results, columns=['帐号', '投注', '盈亏', '起始日期', '迄止日期'])

    # 確保 account_data 和 result_df 中的 '起始日期' 和 '迄止日期' 都是 datetime64[ns] 類型
    account_data['起始日期'] = pd.to_datetime(account_data['起始日期'], errors='coerce')
    result_df['起始日期'] = pd.to_datetime(result_df['起始日期'], errors='coerce')

    # 在合併時根據 '帐号', '起始日期', 和 '迄止日期' 來合併，這樣可以避免時間不同的帳號被重複新增
    df_merged = pd.merge(account_data, result_df, on=['帐号', '起始日期'], how='left')

    # 設定預設檔名為原檔案名稱加上 "_updated"
    new_file_name = file_path.replace('.xlsx', '_updated.xlsx')

    # 顯示儲存檔案對話框，允許使用者修改檔名
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=new_file_name)
    
    if save_path:  # 若使用者選擇了檔案位置
        with pd.ExcelWriter(save_path, engine='openpyxl', mode='w') as writer:
            df_merged.to_excel(writer, index=False, sheet_name=sheet_name)
        messagebox.showinfo("完成", f"資料已成功保存至：{save_path}")
    else:
        messagebox.showinfo("取消", "未保存資料。")
        return

# 提示用戶選擇文件
def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="選擇帳號日期資料的 Excel 檔案", filetypes=[("Excel Files", "*.xlsx")])
    return file_path

# 訊息框顯示函數
def show_messagebox():
    root = tk.Tk()
    root.withdraw()  # 隱藏主視窗
    messagebox.askokcancel("開始查詢", "即將開始進行查詢，是否繼續？")

# -------------------------------
# 主程式，採用迴圈替代遞迴（第 6 點修改）
# -------------------------------
def main(browser, page):
    continue_query = True
    while continue_query:
        file_path = select_excel_file()
        if not file_path:
            print("未選擇檔案，程式結束。")
            return

        account_data = read_account_date_from_excel(file_path)
        print(account_data)

        # 顯示訊息框確認是否開始查詢
        response = messagebox.askokcancel("開始查詢", "即將開始進行查詢，是否繼續？")
        if not response:
            print("使用者取消查詢。")
            return

        all_results = []
        for site in sites:
            for _, row in account_data.iterrows():
                accounts = [row['帐号']]
                start_time = row['起始日期']
                end_time = row['迄止日期']
                results = scrape_cj_site(site, accounts, start_time, end_time, page)
                all_results.extend(results)

        if all_results:
            save_results_to_excel(all_results, file_path, sites[0]["name"], account_data)

        print("所有查詢已完成，結果已儲存。")
        continue_query = messagebox.askyesno("繼續查詢", "查詢已完成，是否繼續查詢？")

    print("程式結束。")
    browser.close()

# -------------------------------
# 7. 統一程式進入點，移除多餘呼叫
# -------------------------------
if __name__ == "__main__":
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(sites[0]["url"])
        main(browser, page)
