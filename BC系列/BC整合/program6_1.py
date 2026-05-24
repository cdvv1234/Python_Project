import pandas as pd
from datetime import datetime, timedelta
import asyncio
from tkinter import messagebox, filedialog

# ====================== 輔助函數 ======================
async def set_date_async(page, selector, date_str):
    try:
        await page.evaluate("""
            (args) => {
                const [sel, val] = args;
                const input = document.querySelector(sel);
                if (input) {
                    input.value = val;
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    input.dispatchEvent(new Event('blur', { bubbles: true }));
                }
            }
        """, [selector, date_str])
        await asyncio.sleep(0.5)
        return True
    except:
        return False


# ====================== 核心執行邏輯 ======================
async def run_program_6_1_logic(root, selected_sites, pages, excel_path, start_date, end_date, callback):
    all_lottery = {}      # 存放彩票資料
    chart_dict = {}       # 存放團隊統計資料 (key = (站台, 帳號))

    # ====================== 1. 彩票抓取 ======================
    async def fetch_lottery(page, site, account):
        try:
            url_suffix = "TeamProfitReport" if site["name"] in ["TC", "TF"] else "ProfitReport"
            await page.goto(f"{site['url']}{url_suffix}", wait_until="networkidle")
            await page.wait_for_load_state("domcontentloaded")

            if site["name"] in ["TC", "TF"]:
                start_dt = datetime.strptime(start_date, '%Y-%m-%d')
                end_dt = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)
                await page.fill('#StartTime', "")
                await page.fill('#EndTime', "")
                await page.fill('#StartTime', start_dt.strftime('%Y/%m/%d 03:00'))
                await page.fill('#EndTime', end_dt.strftime('%Y/%m/%d 03:00'))
            else:
                await set_date_async(page, '#StartTime', start_date)
                await set_date_async(page, '#EndTime', end_date)

            if site["name"] in ["TC", "TF"]:
                await page.fill('#LoginID', account)
                await page.click('#querybutton')
                await page.wait_for_selector('#TeamProfitTable', timeout=15000)
                data = await page.evaluate("""
                    () => Array.from(document.querySelectorAll('#TeamProfitTable tbody tr'))
                              .map(r => Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()))
                """)
                if data:
                    actual_cols = len(data[0])
                    cols = ['总投注', '总奖金', '总返點', '总活動', '总盈虧', '总充值', '总提款', '总紅利'][:actual_cols]
                    df_data = [[site['name'], account] + row for row in data]
                    return pd.DataFrame(df_data, columns=['平台', '帳號'] + cols)
            else:
                await page.fill('#LoginId', account)
                await page.locator('button[data-bind*="SearchClick"]').first.click()
                await asyncio.sleep(1.5)
                await page.wait_for_selector('tbody[data-bind="with: TeamStatistics"]', timeout=15000)
                table_data = await page.evaluate("""
                    () => Array.from(document.querySelectorAll('tbody[data-bind="with: TeamStatistics"] tr'))
                              .map(r => Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()))
                """)
                if table_data:
                    actual_cols = len(table_data[0])
                    if actual_cols == 8:
                        cols = ['总投注', '总奖金', '总返点', '总活動', '总盈亏', '总充值', '总提款', '总红利']
                    elif actual_cols == 9:
                        cols = ['总投注', '总奖金', '总返点', '总活動', '总盈虧', '总充值', '总提款', '总紅利', '平台服务费']
                    else:
                        cols = [f"列_{i+1}" for i in range(actual_cols)]
                    df_data = [[site['name'], account] + row for row in table_data]
                    return pd.DataFrame(df_data, columns=['平台', '帳號'] + cols)
        except Exception as e:
            print(f"[{site['name']}] 彩票抓取失敗 {account}: {e}")
        return None

    # ====================== 2. 團隊統計圖表（TC/TF 不填投注量） ======================
    async def fetch_team_chart_for_site(page, site, s_d, e_d, accounts):
        try:
            url = f"{site['url']}StatisticsChart/TeamInfo" if site["name"] in ["TC", "TF"] else f"{site['url']}TeamStatisticsChart"
            await page.goto(url, wait_until="networkidle", timeout=60000)
            await asyncio.sleep(2)

            # 填日期（只做一次）
            if site["name"] in ["TC", "TF"]:
                start_str = s_d.replace('-', '/')
                end_str = e_d.replace('-', '/')
                await page.fill('#StartTime', start_str)
                await page.fill('#EndTime', end_str)
            else:
                await set_date_async(page, '#StartTime', s_d)
                await set_date_async(page, '#EndTime', e_d)

            # 填投注量（只有非 TC/TF 才填）
            if site["name"] not in ["TC", "TF"]:
                bet_sel = '#BetAmountTH'
                await page.fill(bet_sel, "0.001")
                await asyncio.sleep(1)

            # 依序查詢每個帳號
            for acc in accounts:
                login_sel = '#LoginID' if site["name"] in ["TC", "TF"] else '#LoginID'
                await page.fill(login_sel, acc)

                if site["name"] in ["TC", "TF"]:
                    await page.click('#Search')
                else:
                    await page.click("#SearchForm > div:nth-child(4) > div > button")

                await asyncio.sleep(2.5)

                # === 修正重點：更穩定的提示窗處理 ===
                try:
                    # 最多等 2 秒看是否有提示窗
                    warning_btn = await page.wait_for_selector("#Warning-Dialog-CloseBtn", timeout=2000)
                    if warning_btn:
                        await warning_btn.click()
                        await asyncio.sleep(1.2)
                        chart_dict[(site["name"], acc)] = {'投注人數': '0', '開戶人數': '0'}
                        continue
                except:
                    # 超過 2 秒還沒出現提示窗 = 正常有資料
                    pass

                # 正常抓取圖表數據
                bet = await page.evaluate('() => document.querySelector("#chart > div.jqplot-point-label.jqplot-series-1.jqplot-point-0")?.innerText.trim() || "0"')
                open_num = await page.evaluate('() => document.querySelector("#chart > div.jqplot-point-label.jqplot-series-0.jqplot-point-0")?.innerText.trim() || "0"')

                chart_dict[(site["name"], acc)] = {'投注人數': bet, '開戶人數': open_num}

        except Exception as e:
            print(f"[{site['name']}] 團隊統計圖表失敗 {acc if 'acc' in locals() else ''}: {e}")

    # ====================== 主流程 ======================
    try:
        excel_data = pd.read_excel(excel_path, sheet_name=None)

        tasks = []
        for site, page in zip(selected_sites, pages):
            async def process_site(site=site, page=page):
                accounts = []
                if site["name"] in excel_data:
                    df = excel_data[site["name"]]
                    col_idx = 0 if df.shape[1] == 1 else 1
                    accounts = df.iloc[:, col_idx].dropna().astype(str).str.strip().tolist()

                # 彩票部分
                lottery_dfs = []
                for acc in accounts:
                    df = await fetch_lottery(page, site, acc)
                    if df is not None and not df.empty:
                        lottery_dfs.append(df)
                if lottery_dfs:
                    all_lottery[site["name"]] = pd.concat(lottery_dfs, ignore_index=True)

                # 團隊統計部分（每個站台只導航一次）
                if accounts:
                    await fetch_team_chart_for_site(page, site, start_date, end_date, accounts)

            tasks.append(process_site())

        await asyncio.gather(*tasks)

    except Exception as e:
        root.after(0, lambda: messagebox.showerror("錯誤", f"執行失敗: {e}"))
        if callback: callback()
        return

# ====================== 合併並只保留 7 個欄位 + 新增總表 ======================
    def finalize_save():
        if not all_lottery:
            messagebox.showinfo("提示", "未抓取到任何資料。")
            if callback: callback()
            return

        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"阿恩团队统计表_{timestamp}.xlsx"
        )

        if save_path:
            try:
                site_order = ["TC", "TF", "TS", "SY", "FL", "WX", "XC", "XH", "CJ", "CY", "YD"]

                with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                    all_data = []   # 用來存放總表資料

                    # 先寫入各站台的明細表（按照指定順序）
                    for site_name in site_order:
                        if site_name in all_lottery:
                            lottery_df = all_lottery[site_name]

                            # 合併團隊統計資料
                            chart_data = []
                            for _, row in lottery_df.iterrows():
                                acc = row['帳號']
                                result = chart_dict.get((site_name, acc), {'投注人數': '0', '開戶人數': '0'})
                                chart_data.append([result['投注人數'], result['開戶人數']])

                            chart_df = pd.DataFrame(chart_data, columns=['投注人數', '開戶人數'])

                            final_df = pd.concat([lottery_df.reset_index(drop=True), chart_df.reset_index(drop=True)], axis=1)

                            # 只保留指定的 7 個欄位
                            keep_columns = ['平台', '帳號', '总投注', '投注人數', '開戶人數', '总充值', '总提款']
                            final_df = final_df[keep_columns]

                            final_df.to_excel(writer, sheet_name=f"{site_name}_彩票", index=False)

                            # 收集資料給總表
                            all_data.append(final_df)

                    # ====================== 新增總表（放在最前面） ======================
                    if all_data:
                        total_df = pd.concat(all_data, ignore_index=True)
                        # 總表按照站台順序排序
                        total_df = total_df.sort_values(by=['平台'], 
                                                       key=lambda x: x.map({s: i for i, s in enumerate(site_order)}))
                        total_df.to_excel(writer, sheet_name="總表", index=False)

                messagebox.showinfo("成功", f"抓取完成！\n已儲存至：\n{save_path}\n\n第一頁為總表 + 各站台明細")
            except Exception as e:
                messagebox.showerror("錯誤", f"儲存失敗: {e}")

        if callback:
            callback()

    root.after(0, finalize_save)


# ====================== 給 program6 呼叫的入口 ======================
def run_program_6_1(root, selected_sites, pages, excel_path, start_date, end_date, callback):
    from __main__ import app
    app.run_async(run_program_6_1_logic(root, selected_sites, pages, excel_path, start_date, end_date, callback))