import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import asyncio
from datetime import datetime

# 保持原有的讀取邏輯
def read_account_date_from_excel(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df[['帐号', '领取日期', '起始日期', '迄止日期']]

# 改為非同步版本的抓取邏輯
async def scrape_cj_site_async(site, accounts, start_time, end_time, page):
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

            is_tc_tf = site["name"] in ["TC", "TF"]

            # 1. 每次查詢前清除「帳號與更新時間」欄位 (不影響狀態勾選)
            if is_tc_tf:
                await page.locator("#LoginID").fill("")
                await page.locator("#UpdateStartTime").fill("")
                await page.locator("#UpdateEndTime").fill("")
            else:
                await page.fill('#LoginId', '')
                await page.fill('#UpdateStartTime', '')
                await page.fill('#UpdateEndTime', '')
            await asyncio.sleep(0.5)

            # 2. 執行查詢
            if is_tc_tf:
                await page.locator("#LoginID").fill(query)
                await page.locator("#UpdateStartTime").fill(start_time_str)
                await page.locator("#UpdateEndTime").fill(end_time_str)
                await page.click('#querybutton', force=True)
                await asyncio.sleep(3.5)

                if await page.locator("td:has-text('表中数据为空')").is_visible(timeout=5000):
                    bet_amount, profit_loss = '0', '0'
                else:
                    bet_amount = (await page.locator('#TotalBet').inner_text(timeout=10000)).strip().replace(',', '')
                    profit_loss = (await page.locator('#TotalEarn').inner_text(timeout=10000)).strip().replace(',', '')
            else:
                await page.fill('#LoginId', query)
                await page.fill('#UpdateStartTime', start_time_str)
                await page.fill('#UpdateEndTime', end_time_str)
                await page.click('#Search', force=True)
                await asyncio.sleep(3.5)
                
                # 處理警告視窗
                if await page.locator('#Warning-Dialog-CloseBtn').is_visible():
                    await page.click('#Warning-Dialog-CloseBtn')
                    await page.click('#Search', force=True)
                    await asyncio.sleep(2)

                await page.wait_for_selector('#SearchResult tbody', timeout=10000)
                empty_table = await page.locator('#SearchResult > tbody > tr > td').first.inner_text()
                if "查无资料" in empty_table:
                    bet_amount, profit_loss = '0', '0'
                else:
                    bet_amount = await page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(11)").inner_text()
                    profit_loss = await page.locator("#SearchResult > tfoot > tr:nth-child(2) > th:nth-child(13)").inner_text()

            results.append([query, bet_amount, profit_loss, start_time, end_time_str])
            processed_accounts.add(account_time_key)
            previous_bet_amount = bet_amount
            previous_profit_loss = profit_loss

        except Exception as e:
            print(f"處理帳號 {query} 時發生錯誤: {e}")

    return results

# 新增：專門處理狀態設定與日期清除的函數 (含重試機制)
async def setup_site_state_async(site, page):
    for attempt in range(1, 4):
        try:
            await page.goto(f"{site['url']}BetRecord", wait_until="networkidle", timeout=30000)
            await asyncio.sleep(3)

            if site["name"] in ["TC", "TF"]:
                # TC/TF 邏輯
                state_container = page.locator("div.col-12.col-md-1.dropdown_list.LOTTERY > dropdown-multiple-checkbox")
                await state_container.locator("input[readonly][data-bind*='click: controlOptions']").click()
                await page.wait_for_selector("div.col-12.col-md-1.dropdown_list.LOTTERY > dropdown-multiple-checkbox ul li", timeout=10000)
                
                all_check = state_container.locator("input[data-bind*='checked: checkAll']")
                if await all_check.is_checked(): await all_check.uncheck()
                
                await state_container.locator("span:text-is('已中奖')").locator("..").locator("input[type='checkbox']").check()
                await state_container.locator("span:text-is('未中奖')").locator("..").locator("input[type='checkbox']").check()
                
                # 清除日期 (還原原本邏輯)
                await page.locator("#StartTime").fill("")
                await page.locator("#EndTime").fill("")
                await page.click('body')
            else:
                # 其他站台邏輯
                await page.click('#StateSelectedList')
                await page.wait_for_function("document.querySelector('#StateDropDownMultiple').style.display !== 'none'", timeout=10000)
                
                if await page.locator('#StateClickAll').is_checked(): await page.locator('#StateClickAll').uncheck()
                
                await page.locator("#StateDropDownMultiple label:has(span:text-is('未中奖')) input[type='checkbox']").check()
                await page.locator("#StateDropDownMultiple label:has(span:text-is('中奖')) input[type='checkbox']").check()
                
                # 清除日期
                await page.fill('#StartTime', "")
                await page.fill('#EndTime', "")
                await page.click('body')

            print(f"站台 {site['name']} 狀態與日期初始化完成")
            return True
        except Exception as e:
            print(f"站台 {site['name']} 第 {attempt} 次初始化失敗: {e}")
            if attempt == 3: return False
            await asyncio.sleep(2)
    return False

# 封裝任務
async def process_site_task(site, page, sheet_names, file_path):
    site_results_dict = {}
    site_sheets = [sheet for sheet in sheet_names if sheet.startswith(site["name"])]
    if not site_sheets: return site_results_dict

    # 執行初始化 (狀態+清除日期)
    success = await setup_site_state_async(site, page)
    if not success:
        print(f"站台 {site['name']} 初始化最終失敗，跳過此站台")
        return site_results_dict

    for sheet_name in site_sheets:
        account_data = read_account_date_from_excel(file_path, sheet_name)
        if account_data.empty: continue
        
        sheet_results = []
        for _, row in account_data.iterrows():
            res = await scrape_cj_site_async(site, [row['帐号']], row['起始日期'], row['迄止日期'], page)
            sheet_results.extend(res)
        
        site_results_dict[sheet_name] = sheet_results
    return site_results_dict

# 儲存 Excel
def save_results_to_excel(file_path, all_results_dict, account_data_dict):
    current_time = datetime.now().strftime('%Y%m%d%H%M')
    new_file_name = f"投注紀錄查詢_{current_time}.xlsx"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", 
        filetypes=[("Excel Files", "*.xlsx")], 
        initialfile=new_file_name, 
        title="另存查詢結果"
    )
    
    if save_path:
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            for sheet_name, results in all_results_dict.items():
                if results:
                    # 1. 將查詢結果轉為 DataFrame
                    result_df = pd.DataFrame(results, columns=['帐号', '投注', '盈亏', '起始日期', '迄止日期'])
                    
                    # 2. 獲取原始 Excel 數據
                    account_data = account_data_dict.get(sheet_name, pd.DataFrame()).copy()
                    
                    # 3. 統一日期格式以利合併 (避免字串與時間物件比對失敗)
                    account_data['起始日期'] = pd.to_datetime(account_data['起始日期'], errors='coerce')
                    result_df['起始日期'] = pd.to_datetime(result_df['起始日期'], errors='coerce')
                    
                    # 4. 去除查詢結果中重複的帳號+日期，並「丟棄」迄止日期，避免產生 _x, _y
                    # 我們保留原始 account_data 裡的迄止日期即可
                    result_clean = result_df.drop_duplicates(subset=['帐号', '起始日期']).drop(columns=['迄止日期'])
                    
                    # 5. 合併資料
                    df_merged = pd.merge(
                        account_data, 
                        result_clean, 
                        on=['帐号', '起始日期'], 
                        how='left'
                    )
                    
                    # 6. 填補空值並確保欄位順序 (視需求調整)
                    df_merged['投注'] = df_merged['投注'].fillna('0')
                    df_merged['盈亏'] = df_merged['盈亏'].fillna('0')
                    
                    # 寫入 Excel
                    df_merged.to_excel(writer, index=False, sheet_name=sheet_name)
                    
        messagebox.showinfo("完成", f"資料已成功保存至：{save_path}")

# 主入口
def run_program_2(root, selected_sites, pages, callback):
    file_path = filedialog.askopenfilename(title="選擇帳號日期資料的 Excel 檔案", filetypes=[("Excel Files", "*.xlsx")])
    if not file_path:
        if callback: callback()
        return

    excel_file = pd.ExcelFile(file_path)
    sheet_names = excel_file.sheet_names
    task_pairs = [(site, page) for site, page in zip(selected_sites, pages) if any(sn.startswith(site["name"]) for sn in sheet_names)]

    if not task_pairs:
        messagebox.showwarning("錯誤", "未找到匹配的工作表")
        if callback: callback()
        return

    async def main_logic():
        try:
            # 並行執行所有站台任務
            tasks = [process_site_task(site, page, sheet_names, file_path) for site, page in task_pairs]
            results_list = await asyncio.gather(*tasks)

            all_results_dict = {}
            for r_dict in results_list: all_results_dict.update(r_dict)

            if all_results_dict:
                save_results_to_excel(file_path, all_results_dict, 
                                     {sn: read_account_date_from_excel(file_path, sn) for sn in all_results_dict.keys()})
            
            # 詢問是否繼續
            if messagebox.askyesno("繼續查詢", "查詢已完成，是否繼續查詢？"):
                # 如果選「是」，遞迴呼叫自己，不恢復按鈕
                run_program_2(root, selected_sites, pages, callback)
            else:
                # 如果選「否」，執行 callback 恢復 Main 按鈕並關閉
                if callback: callback()

        except Exception as e:
            print(f"主程序錯誤: {e}")
            if callback: callback()
        finally:
            # 確保不會卡住
            print("非同步任務結束")

    # 啟動非同步
    import __main__
    if hasattr(__main__, 'app'):
        __main__.app.run_async(main_logic())
    else:
        asyncio.run(main_logic())