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


# ====================== 下級名單抓取（其餘站台加強定位版） ======================
async def fetch_lower_members_for_category(page, site, upper_account, category):
    try:
        page_num = 1
        if site["name"] in ["TC", "TF"]:
            # ==================== TC/TF 不做任何改動 ====================
            if category == "彩票":
                table_sel = "#TeamMemberProfitTable"
                next_sel = "#TeamMemberProfitTable_next"
            else:
                table_sel = "#TeamMemberProfitTableForExternalGame"
                next_sel = "#TeamMemberProfitTableForExternalGame_next"

        else:
            # ==================== 其餘站台 - 精準定位 ====================
            if category == "彩票":
                # 彩票專屬區塊
                await page.locator('#lottery-tab').click()
                await asyncio.sleep(1.2)
                table_sel = "#Lottery #MemberDataTable"
                next_sel = "#Lottery #MemberDataTable_next"
            else:
                # 外接遊戲（真人電子、體育、棋牌）
                await page.locator('#external-tab').click()
                await asyncio.sleep(1.2)
                table_sel = "#External #MemberDataTable"
                next_sel = "#External #MemberDataTable_next"

        # 檢查表格是否存在
        if await page.locator(table_sel).count() == 0:
            return None

        await page.wait_for_selector(table_sel, timeout=15000)
        await asyncio.sleep(1.5)

        all_rows = []
        while True:
            rows = await page.evaluate(f"""
                () => Array.from(document.querySelectorAll('{table_sel} tbody tr'))
                          .map(r => Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()))
            """)

            for row in rows:
                if row and len(row) > 1:
                    print(f"[{site['name']}] 正在抓取帳號: {upper_account} | 頁碼: {page_num} | Row內容: {row}")
                    if any(skip in str(row[0]) for skip in ["总和", "總和", "无数据", "查無數據", "表中数据为空"]):
                        continue
                    all_rows.append(row)

            # 翻頁
            next_li = await page.query_selector(next_sel)
            if not next_li:
                break

            li_class = await next_li.get_attribute("class") or ""
            if "disabled" in li_class:
                break

            clicked = await page.evaluate(f"""
                () => {{
                    const next = document.querySelector('{next_sel} a');
                    if (next) {{
                        next.click();
                        return true;
                    }}
                    return false;
                }}
            """)
            if not clicked:
                break

            page_num += 1
            await asyncio.sleep(1.2)

        if not all_rows:
            return None

        # 👉 修改這裡：依照類別，補齊「類別」欄位，讓上下級欄位長度完全一致
        if category == "彩票":
            cols = ['平台', '上級帳號'] + [f"欄位{i+1}" for i in range(len(all_rows[0]))]
            data = [[site["name"], upper_account] + row for row in all_rows]
        else:
            # 👉 解決問題 2：外接遊戲手動補上「類別」，並插入空位給「项目」
            display_cat = category.replace("電子", "")
            data = []
            for row in all_rows:
                if len(row) > 1:
                    # 原本的 row 是 [下級帳號, 總投注, 有效投注...]
                    # 我們在 index 1 插入一個空字串 ""，變成 [下級帳號, "", 總投注, 有效投注...]
                    aligned_row = [row[0], ""] + row[1:]
                    data.append([site["name"], upper_account, display_cat] + aligned_row)
            
            # 因為多塞了一個 ""，所以欄位數量要以 len(all_rows[0]) + 1 計算
            cols = ['平台', '上級帳號', '類別'] + [f"欄位{i+1}" for i in range(len(all_rows[0]) + 1)]

        return pd.DataFrame(data, columns=cols)

    except Exception as e:
        print(f"[{site['name']}] {category} 下級名單抓取跳過 {upper_account}: {e}")
        return None

# ====================== 原有抓取函數（完整複製自 program6） ======================
async def fetch_data_lottery(page, site, account):
    try:
        await page.fill('#LoginId', account)
        await page.locator('button[data-bind*="SearchClick"]').first.click()
        await page.wait_for_selector('tbody[data-bind="with: TeamStatistics"]', timeout=15000)
        table_data = await page.evaluate("() => Array.from(document.querySelectorAll('tbody[data-bind=\"with: TeamStatistics\"] tr')).map(r => Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()))")
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
        return None
    except:
        return None


async def fetch_data_external(page, site, account, s_d, e_d, cat, cat_val):
    try:
        await page.locator('#external-tab').click()
        await page.fill('#External #LoginId', account)
        await page.fill('#External #StartTime', s_d)
        await page.fill('#External #EndTime', e_d)
        await page.select_option('#External #Category', cat_val)
        await page.click('#External #SearchForm button.btn-primary')
        await asyncio.sleep(1.5)
        data = await page.evaluate("""() => { 
            const row = Array.from(document.querySelectorAll('#External .table-responsive table tbody tr')).find(r => r.innerText.includes('总和')); 
            return row ? Array.from(row.querySelectorAll('td')).map(c => c.innerText.trim()) : []; 
        }""")
        if data:
            actual_cols = len(data)
            if actual_cols == 6:
                cols = ['项目', '总投注', '有效投注', '总奖金', '总活動', '总盈亏']
            elif actual_cols == 8:
                cols = ['项目', '总投注', '有效投注', '总奖金', '总盈虧', '总返點', '总活動', '结果']
            else:
                cols = [f"列_{i+1}" for i in range(actual_cols)]
            display_cat = cat.replace("電子", "")
            df_data = [[site['name'], display_cat, account] + data]
            return pd.DataFrame(df_data, columns=['平台', '類別', '帳號'] + cols)
        return None
    except:
        return None


async def fetch_data_tc_tf(page, site, account, category, category_value):
    try:
        await page.locator('#GameTypeId').click()
        await page.locator(f'li.GameType[value="{category_value}"] a').click()
        await page.fill('#LoginID', account)
        await page.click('#querybutton')

        res_sel = '#teamResult' if category == "彩票" else '#teamResultForExternalGame'
        await page.wait_for_function("(args) => { const el = document.querySelector(args.sel); return el && el.innerText.toLowerCase().includes(args.acc.toLowerCase()); }", arg={"sel": res_sel, "acc": account}, timeout=15000)
        tab_sel = '#TeamProfitTable' if category == "彩票" else '#TeamProfitTableForExternalGame'
        
        # 👉 解決問題 1：外接遊戲切換時，強制等待表格內出現對應的文字
        if category != "彩票":
            wait_text = ""
            if category == "真人電子": wait_text = "电子"
            elif category == "體育": wait_text = "体育"
            elif category == "棋牌": wait_text = "棋牌"
            
            if wait_text:
                try:
                    await page.wait_for_function(
                        "(args) => { const el = document.querySelector(args.sel); return el && el.innerText.includes(args.text); }",
                        arg={"sel": tab_sel, "text": wait_text},
                        timeout=15000
                    )
                except Exception as e:
                    print(f"[{site['name']}] 等待表格文字 '{wait_text}' 超時或查無資料")
        
        if category == "彩票":
            data = await page.evaluate(f"() => Array.from(document.querySelectorAll('{tab_sel} tbody tr')).map(r => Array.from(r.querySelectorAll('td')).map(c => c.innerText.trim()))")
            if data:
                actual_cols = len(data[0])
                cols = ['总投注', '总奖金', '总返點', '总活動', '总盈虧', '总充值', '总提款', '总紅利']
                df_data = [[site['name'], account] + r for r in data]
                return pd.DataFrame(df_data, columns=['平台', '帳號'] + cols[:actual_cols])
        else:
            data = await page.evaluate("""(sel) => { 
                const row = Array.from(document.querySelectorAll(sel + ' tbody tr')).find(r => r.innerText.includes('总和')); 
                return row ? Array.from(row.querySelectorAll('td')).map(c => c.innerText.trim()) : null; 
            }""", tab_sel)
            if data:
                display_cat = category.replace("電子", "")
                df_data = [[site['name'], display_cat, account] + data[:6]]
                cols = ['平台', '類別', '帳號', '项目', '总投注', '有效投注', '总奖金', '总活動', '总盈亏']
                return pd.DataFrame(df_data, columns=cols)
        return None
    except:
        return None


# ====================== 核心執行邏輯 ======================
async def run_program_6_2_logic(root, selected_sites, pages, excel_path, start_date, end_date, check_vars, callback):
    all_results = {}

    async def process_site(site, page, excel_data, s_d, e_d, category_map):
        try:
            # 1. 導航與初始化
            url_suffix = "TeamProfitReport" if site["name"] in ["TC", "TF"] else "ProfitReport"
            await page.goto(f"{site['url']}{url_suffix}", wait_until="networkidle")
            await page.wait_for_load_state("domcontentloaded")

            # 處理日期設定
            if site["name"] in ["TC", "TF"]:
                start_dt = datetime.strptime(s_d, '%Y-%m-%d')
                end_dt = datetime.strptime(e_d, '%Y-%m-%d') + timedelta(days=1)
                await page.fill('#StartTime', start_dt.strftime('%Y/%m/%d 03:00'))
                await page.fill('#EndTime', end_dt.strftime('%Y/%m/%d 03:00'))
            else:
                await set_date_async(page, '#StartTime', s_d)
                await set_date_async(page, '#EndTime', e_d)

            # 獲取帳號列表
            accounts = []
            if site["name"] in excel_data:
                df = excel_data[site["name"]]
                col_idx = 0 if df.shape[1] == 1 else 1
                accounts = df.iloc[:, col_idx].dropna().astype(str).str.strip().tolist()

            # 2. 類別順序：強制將彩票放最前，確保執行順序
            categories_order = ["彩票", "真人電子", "體育", "棋牌"]

            for cat in categories_order:
                if cat not in check_vars or not check_vars[cat].get():
                    continue

                # --- A. TC/TF 邏輯區 (完全分離，保持穩定) ---
                if site["name"] in ["TC", "TF"]:
                    for acc in accounts:
                        df = await fetch_data_tc_tf(page, site, acc, cat, category_map[cat])
                        if df is not None:
                            df.insert(1, '上級帳號', acc)
                            all_results.setdefault(f"{site['name']}_{cat}", []).append(df)
                        
                        lower_df = await fetch_lower_members_for_category(page, site, acc, cat)
                        if lower_df is not None and not lower_df.empty:
                            all_results.setdefault(f"{site['name']}_{cat}", []).append(lower_df)

                # --- B. 其他站台邏輯區 (分批批次處理) ---
                else:
                    # 在該類別下，執行一次性切換
                    if cat == "彩票":
                        await page.locator('#lottery-tab').click()
                    else:
                        await page.locator('#external-tab').click()
                    await asyncio.sleep(1.0) # 給予切換緩衝

                    for acc in accounts:
                        # 在此 Tab 內僅進行清空與查詢
                        try:
                            await page.fill('#LoginId', '') 
                            await asyncio.sleep(0.3)
                        except: pass
                        
                        # 執行 fetch_data_lottery 或 fetch_data_external
                        # 注意：這些函數內不應再有重複點擊 Tab 的行為
                        if cat == "彩票":
                            df = await fetch_data_lottery(page, site, acc)
                        else:
                            df = await fetch_data_external(page, site, acc, s_d, e_d, cat, category_map[cat])
                        
                        if df is not None:
                            if '上級帳號' not in df.columns:
                            # 將「上級帳號」插入到 index 1 的位置
                                df.insert(1, '上級帳號', acc)
                            all_results.setdefault(f"{site['name']}_{cat}", []).append(df)

                            # 將數值轉為 float，排除項目文字欄位
                            numeric_data = []
                            for col in df.columns[3:]: # 假設從第 4 欄開始是數值
                                val = str(df.iloc[0][col]).replace(',', '')
                                try:
                                    numeric_data.append(float(val))
                                except:
                                    numeric_data.append(0.0)
                            
                            # 如果整列數值加總大於 0，才去抓下級
                            if sum(numeric_data) > 0:
                                lower_df = await fetch_lower_members_for_category(page, site, acc, cat)
                                if lower_df is not None and not lower_df.empty:
                                    # 【核心修正 1】：確保下級名單欄位與基準完全一致
                                    # 強制設定下級名單的欄位名稱，使其與上方統計資料對齊
                                    lower_df.columns = df.columns 
                                    all_results.setdefault(f"{site['name']}_{cat}", []).append(lower_df)
                            else:
                                print(f"[{site['name']}] 帳號 {acc} 該類別無數據，跳過抓取下級。")

        except Exception as e:
            print(f"[{site['name']}] 處理失敗: {e}")

    async def _async_process():
        try:
            excel_data = pd.read_excel(excel_path, sheet_name=None)
            s_d = start_date
            e_d = end_date
            category_map = {"彩票": "0", "真人電子": "1", "體育": "2", "棋牌": "3"}

            tasks = [process_site(site, page, excel_data, s_d, e_d, category_map) for site, page in zip(selected_sites, pages)]
            await asyncio.gather(*tasks)

            root.after(0, lambda: finalize_save(all_results, callback))

        except Exception as e:
            root.after(0, lambda: messagebox.showerror("錯誤", f"執行失敗: {e}"))
            if callback: callback()

    from __main__ import app
    app.run_async(_async_process())


def finalize_save(all_results, callback):
    if not all_results:
        messagebox.showinfo("提示", "未抓取到任何資料。")
        if callback: callback()
        return

    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=f"招商觀察_含下級名單_{timestamp}.xlsx"
    )

    if save_path:
        try:
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                site_order = ["TC", "TF", "TS", "SY", "YD", "FL", "WX", "XC", "XH", "CJ", "CY"]
                cat_order = ["彩票", "真人電子", "體育", "棋牌"]
                
                for site_name in site_order:
                    for cat in cat_order:
                        key = f"{site_name}_{cat}"
                        if key in all_results and all_results[key]:
                            dfs = all_results[key]
                            
                            # 👉 1. 找出「上級資料」的真實表頭當作基準 (排除名稱含"欄位"的下級表頭)
                            base_columns = None
                            for df in dfs:
                                if not any("欄位" in str(c) for c in df.columns):
                                    base_columns = df.columns
                                    break
                            
                            # 如果這組全部都只有下級資料 (極端情況)，就拿第一個當基準
                            if base_columns is None:
                                base_columns = dfs[0].columns
                                
                            aligned_dfs = []
                            for df in dfs:
                                # 👉 2. 強制對齊欄位名稱，避免 pd.concat 往右新增
                                if len(df.columns) == len(base_columns):
                                    df.columns = base_columns
                                elif len(df.columns) > len(base_columns):
                                    # 如果下級欄位多於上級(例如網頁多了一欄操作按鈕)，截斷多餘部分
                                    df = df.iloc[:, :len(base_columns)]
                                    df.columns = base_columns
                                else:
                                    # 如果下級欄位少於上級，依序對齊前面欄位 (防錯機制)
                                    new_cols = list(base_columns[:len(df.columns)])
                                    df.columns = new_cols
                                    
                                aligned_dfs.append(df)
                                
                            # 👉 3. 欄位名稱完全一致後再合併匯出
                            pd.concat(aligned_dfs, ignore_index=True).to_excel(writer, sheet_name=key[:31], index=False)
                            
            messagebox.showinfo("成功", f"抓取完成！\n已儲存至：\n{save_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗: {e}")

    if callback:
        callback()


def run_program_6_2(root, selected_sites, pages, excel_path, start_date, end_date, check_vars, callback):
    from __main__ import app
    app.run_async(run_program_6_2_logic(root, selected_sites, pages, excel_path, start_date, end_date, check_vars, callback))