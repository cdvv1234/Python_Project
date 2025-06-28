import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import font as tkFont
import pandas as pd
import time
from datetime import timedelta, datetime

def read_account_date_from_excel(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df[['帐号', '领取日期', '起始日期', '迄止日期']]

def scrape_cj_site(site, accounts, start_time, end_time, page):
    results = []
    processed_accounts = set()
    previous_bet_amount = None
    previous_profit_loss = None

    for query in accounts:
        account_time_key = f"{query}_{start_time}_{end_time}"
        if account_time_key in processed_accounts:
            continue

        try:
            start_time_str = start_time.strftime('%Y/%m/%d %H:%M:%S') if isinstance(start_time, (pd.Timestamp, datetime)) else start_time
            end_time_str = end_time.strftime('%Y/%m/%d %H:%M:%S') if isinstance(end_time, (pd.Timestamp, datetime)) else end_time

            if site["name"].startswith("TC") or site["name"].startswith("TF"):
                iframe = page.locator("iframe[name=\"mainFrame\"]").content_frame
                if not iframe:
                    print(f"未能找到指定的 iframe 'mainFrame'，請檢查選擇器。")
                    continue

                iframe.locator("#LoginID").fill("")
                iframe.locator("#LoginID").fill(query)
                iframe.locator("#UpdateStartTime").fill(start_time_str)
                iframe.locator("#UpdateEndTime").fill(end_time_str)
                iframe.locator('#BetRecordSearchRangeTitle').wait_for(state="visible", timeout=10000)

                iframe.locator('#querybutton:not([disabled])').wait_for(state="visible", timeout=10000)
                iframe.get_by_role("button", name="查询").click()
                time.sleep(2)
                iframe.locator('#BetRecordTable').wait_for(state="attached", timeout=10000)



                empty_table = page.evaluate("""
                    () => {
                        const iframe = document.querySelector("iframe[name='mainFrame']");
                        if (!iframe || !iframe.contentDocument) return '';
                        const td = iframe.contentDocument.querySelector('#BetRecordTable tbody tr td');
                        return td ? td.innerText.trim() : '';
                    }
                """)
                print(f"檢查表格是否為空: {empty_table}")
                if "表中数据为空" in empty_table:
                    bet_amount, profit_loss = '0', '0'
                else:
                    total_bet_amount = 0.0
                    total_profit_loss = 0.0

                    while True:
                        iframe.locator('#BetRecordTable').wait_for(state="attached", timeout=10000)
                        rows_data = page.evaluate("""
                            () => {
                                const iframe = document.querySelector("iframe[name='mainFrame']");
                                if (!iframe || !iframe.contentDocument) return [];
                                const rows = Array.from(iframe.contentDocument.querySelectorAll('#BetRecordTable tbody tr'));
                                return rows.map(row => {
                                    const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText.trim());
                                    return {
                                        bet_amount: cells[10] || '0',
                                        profit_loss: cells[12] || '0'
                                    };
                                });
                            }
                        """)
                        print(f"提取到的行數據: {rows_data}")

                        for row_data in rows_data:
                            try:
                                bet_amount = float(row_data['bet_amount'].replace(',', '')) if row_data['bet_amount'] else 0.0
                                profit_loss = float(row_data['profit_loss'].replace(',', '')) if row_data['profit_loss'] else 0.0
                            except ValueError:
                                print(f"無法將投注金額 {row_data['bet_amount']} 或盈亏 {row_data['profit_loss']} 轉換為數字，跳過此行")
                                continue
                            total_bet_amount += bet_amount
                            total_profit_loss += profit_loss

                        next_button = iframe.locator('#BetRecordTable_next')
                        print(f"下一頁按鈕類名: {next_button.get_attribute('class')}")
                        if "disabled" in next_button.get_attribute("class"):
                            break

                        next_button.locator('a').click()
                        time.sleep(2.5)

                    bet_amount = str(total_bet_amount)
                    profit_loss = str(total_profit_loss)

            else:
                page.fill('#LoginId', query)
                page.fill('#UpdateStartTime', start_time_str)
                page.fill('#UpdateEndTime', end_time_str)
                page.click('#Search', force=True)
                time.sleep(2.5)
                page.wait_for_selector('#SearchResult tbody', timeout=10000)

                if page.locator('#Warning-Dialog-CloseBtn').is_visible():
                    page.click('#Warning-Dialog-CloseBtn')
                    page.click('#Search', force=True)
                    page.wait_for_selector('#SearchResult tbody', timeout=10000)

                empty_table = page.locator('#SearchResult > tbody > tr > td').first.inner_text()
                bet_amount, profit_loss = ('0', '0') if "查无资料" in empty_table else (
                    page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(11)").inner_text(),
                    page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(13)").inner_text()
                )

                # 強化重試邏輯並打印結果
                retries = 0
                max_retries = 5
                while (previous_bet_amount is not None and previous_profit_loss is not None and
                       bet_amount == previous_bet_amount and profit_loss == previous_profit_loss) and retries < max_retries:
                    print(f"重試 {retries + 1}/{max_retries}，帳號: {query}，當前結果 - 投注: {bet_amount}, 盈亏: {profit_loss}")
                    page.fill('#LoginId', query)
                    page.fill('#UpdateStartTime', start_time_str)
                    page.fill('#UpdateEndTime', end_time_str)
                    page.click('#Search', force=True)
                    time.sleep(2.5)
                    page.wait_for_selector('#SearchResult tbody', timeout=10000)

                    if page.locator('#Warning-Dialog-CloseBtn').is_visible():
                        page.click('#Warning-Dialog-CloseBtn')
                        page.click('#Search', force=True)
                        page.wait_for_selector('#SearchResult tbody', timeout=10000)

                    empty_table = page.locator('#SearchResult > tbody > tr > td').first.inner_text()
                    bet_amount, profit_loss = ('0', '0') if "查无资料" in empty_table else (
                        page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(11)").inner_text(),
                        page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(13)").inner_text()
                    )
                    print(f"重試後結果 - 帳號: {query}，投注: {bet_amount}, 盈亏: {profit_loss}")
                    retries += 1

            results.append([query, bet_amount, profit_loss, start_time, end_time_str])
            processed_accounts.add(account_time_key)
            previous_bet_amount = bet_amount
            previous_profit_loss = profit_loss

        except Exception as e:
            print(f"處理帳號 {query} 時發生錯誤: {e}")

    return results

def save_results_to_excel(file_path, all_results_dict, account_data_dict):
    # 獲取當前時間並格式化為 yyyymmddhhmm
    current_time = datetime.now().strftime('%Y%m%d%H%M')
    new_file_name = f"投注紀錄查詢_{current_time}.xlsx"
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=new_file_name, title="另存查詢結果")
    if save_path:
        with pd.ExcelWriter(save_path, engine='openpyxl', mode='w') as writer:
            for site_name, results in all_results_dict.items():
                if results:
                    result_df = pd.DataFrame(results, columns=['帐号', '投注', '盈亏', '起始日期', '迄止日期'])
                    result_df = result_df.drop_duplicates(subset=['帐号', '起始日期'], keep='first')
                    
                    account_data = account_data_dict.get(site_name, pd.DataFrame())  # 處理可能的空數據
                    
                    account_data['起始日期'] = pd.to_datetime(account_data['起始日期'], errors='coerce')
                    result_df['起始日期'] = pd.to_datetime(result_df['起始日期'], errors='coerce')
                    
                    df_merged = pd.merge(account_data, result_df, on=['帐号', '起始日期'], how='left', suffixes=('', '_drop'))
                    df_merged = df_merged.drop(columns=['迄止日期_drop'] if '迄止日期_drop' in df_merged.columns else [])
                    
                    df_merged['投注'] = df_merged['投注'].fillna('0')
                    df_merged['盈亏'] = df_merged['盈亏'].fillna('0')
                    
                    print(f"Final output for {site_name}:")
                    print(str(df_merged))
                    
                    df_merged.to_excel(writer, index=False, sheet_name=site_name)
        messagebox.showinfo("完成", f"資料已成功保存至：{save_path}")

def run_program_2(selected_sites, pages):
    tk.Tk().withdraw()
    file_path = filedialog.askopenfilename(title="選擇帳號日期資料的 Excel 檔案", filetypes=[("Excel Files", "*.xlsx")])
    if not file_path:
        print("未選擇檔案，程式結束。")
        return

    excel_file = pd.ExcelFile(file_path)
    sheet_names = excel_file.sheet_names
    valid_sites = [site for site in selected_sites if any(sheet_name.startswith(site["name"]) for sheet_name in sheet_names)]

    if not valid_sites:
        print("選擇的站台與 Excel 工作表名稱不匹配，程式結束。")
        return

    continue_query = True
    while continue_query:
        all_results_dict = {}
        for site in valid_sites:
            # 找到所有以當前站台名稱開頭的工作表
            site_sheets = [sheet for sheet in sheet_names if sheet.startswith(site["name"])]
            if not site_sheets:
                print(f"未找到 {site['name']} 相關工作表，跳過。")
                continue

            for sheet_name in site_sheets:
                account_data = read_account_date_from_excel(file_path, sheet_name)
                if account_data.empty:
                    print(f"工作表 {sheet_name} 無有效帳號數據，跳過。")
                    continue

                # 獲取對應的頁面（假設頁面已手動導航）
                page = next((p for s, p in zip(selected_sites, pages) if s["name"] == site["name"]), None)
                if not page:
                    print(f"未找到 {site['name']} 的頁面，跳過。")
                    continue

                # 清除前一次查詢的表單數據
                if site["name"].startswith("TC") or site["name"].startswith("TF"):
                    page.locator("iframe[name=\"mainFrame\"]").content_frame.locator("#LoginID").fill("")
                    page.locator("iframe[name=\"mainFrame\"]").content_frame.locator("#UpdateStartTime").fill("")
                    page.locator("iframe[name=\"mainFrame\"]").content_frame.locator("#UpdateEndTime").fill("")
                else:
                    page.fill('#LoginId', '')
                    page.fill('#UpdateStartTime', '')
                    page.fill('#UpdateEndTime', '')
                time.sleep(2.5)

                site_results = []
                for _, row in account_data.iterrows():
                    accounts = [row['帐号']]
                    start_time = row['起始日期']
                    end_time = row['迄止日期']
                    results = scrape_cj_site(site, accounts, start_time, end_time, page)
                    site_results.extend(results)

                all_results_dict[sheet_name] = site_results  # 按工作表名稱分開儲存結果

        if any(all_results_dict.values()):
            save_results_to_excel(file_path, all_results_dict, {sheet: read_account_date_from_excel(file_path, sheet) for sheet in all_results_dict.keys()})

        tk.Tk().withdraw()
        continue_query = messagebox.askyesno("繼續查詢", "查詢已完成，是否繼續查詢？")

        if continue_query:
            file_path = filedialog.askopenfilename(title="選擇帳號日期資料的 Excel 檔案", filetypes=[("Excel Files", "*.xlsx")])
            if not file_path:
                print("未選擇檔案，程式結束。")
                break
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            valid_sites = [site for site in selected_sites if any(sheet_name.startswith(site["name"]) for sheet_name in sheet_names)]
            if not valid_sites:
                print("選擇的站台與 Excel 工作表名稱不匹配，程式結束。")
                break

    print("run_program_2 執行完畢")
