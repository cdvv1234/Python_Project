import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, timedelta
import time
import os
from tkcalendar import DateEntry
from playwright.sync_api import sync_playwright, TimeoutError 

# --- 輔助函式：合併與清理 Excel (多工作表版本) ---

def combine_and_clean_excel(all_site_files, start_dt, end_dt):
    """
    接收按站台分組的下載檔案列表 (all_site_files)，將每個站台的數據合併到一個工作表，
    最終輸出一個包含多個站台工作表的 Excel 檔案。
    """
    
    # ----------------------------------------------------------------
    # 步驟 0: 定義標準欄位映射和最終輸出欄位
    # ----------------------------------------------------------------
    
    # 1. 最終輸出的標準**簡體**欄位 (用於內部處理和最終輸出標題)
    FINAL_EXPORT_COLS_SIMPLIFIED = [
        '交易号', '用户名', '彩种玩法', '期号', '下注时间', '报表日期', 
        '投注注数', '倍数', '单注金额', '投注金额', '中奖金额', '盈亏', '状态', 'IP'
    ]

    # 2. 涵蓋所有可能變體的映射表 (Key: 下載檔案中的可能欄位名, Value: 標準簡體欄位名)
    # 通用映射（其他站台）
    COL_MAPPING_COMMON = {
        '交易號': '交易号', '交易号': '交易号', '交易号.1': '交易号',
        '用户名': '用户名', '帳號': '用户名',
        '彩种玩法': '彩种玩法', '彩種玩法': '彩种玩法',
        '期號': '期号', '期号': '期号',
        '下注時間': '下注时间', '下注时间': '下注时间',
        '報表日期': '报表日期', '报表日期': '报表日期',
        '投注注数': '投注注数', '投注注數': '投注注数',
        '倍数': '倍数', '倍數': '倍数',
        '单注金额': '单注金额', '單注金額': '单注金额',
        '投注金额': '投注金额', '投注金額': '投注金额',
        '中奖金额': '中奖金额', '中獎金額': '中奖金额',
        '盈亏': '盈亏', '盈虧': '盈亏',
        '狀態': '状态', '状态': '状态',
        'IP': 'IP',
        '投注內容': '__TEMP_DISCARD__', 
        '投注内容': '__TEMP_DISCARD__',
        '投注內容_2': '__TEMP_DISCARD__', 
        '投注内容.1': '__TEMP_DISCARD__', 
    }

    # 3. TC/TF 專用映射（標題不同）
    COL_MAPPING_TC_TF = {
        '帐号': '用户名',
        '玩法': '玩法',  # 臨時保留，用於合併彩种玩法
        '中奖金额': '中奖金额',
        '交易号': '交易号',
        '位置': '__TEMP_DISCARD__',
        '下注时间': '下注时间',
        '报表日期': '报表日期',
        '开奖号码': '__TEMP_DISCARD__',
        '彩种': '彩种',  # 臨時保留，用於合併彩种玩法
        '投注总额': '投注金额',
        '期号': '期号',
        '状态': '状态',
        '單注金额': '单注金额',
        '投注注数': '投注注数',
        '销售返点': '__TEMP_DISCARD__',
        '盈亏': '盈亏',
        '倍数模式': '倍数',
        '一倍奖金': '__TEMP_DISCARD__',
        'IP': 'IP',
        '备注': '__TEMP_DISCARD__',
        '投注号码': '__TEMP_DISCARD__',
    }

    # 1. 選擇儲存路徑並匯出
    start_time_str = start_dt.strftime('%Y%m%d%H%M')
    end_time_str = end_dt.strftime('%Y%m%d%H%M')
    default_filename = f"合併投注紀錄_多站點_{start_time_str}_{end_time_str}.xlsx"
    
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=default_filename,
        filetypes=[("Excel files", "*.xlsx")],
        title="儲存合併後的投注紀錄 Excel 檔案 (多工作表)"
    )
    
    if not save_path:
        messagebox.showinfo("提示", "已取消儲存最終檔案。")
        return

    # 2. 使用 pd.ExcelWriter 處理多工作表
    successful_sites = 0 
    try:
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            
            for site_name, all_downloaded_files in all_site_files.items():
                
                if not all_downloaded_files:
                    print(f"站台 {site_name} 沒有下載檔案，跳過處理。")
                    continue
                
                print(f"--- 正在處理站台 {site_name} 的 {len(all_downloaded_files)} 個檔案 ---")
                
                all_dfs = [] 
                is_tc_tf = site_name in ["TC", "TF"]

                for file_path in all_downloaded_files:
                    try:
                        df = pd.read_excel(file_path, header=0, engine='openpyxl')
                        df.dropna(how='all', inplace=True) 

                        # 欄位名稱清理
                        df.columns = df.columns.str.strip().str.replace('\n', '', regex=False)
                        
                        if is_tc_tf:
                            # TC/TF 特殊處理
                            rename_map = {col: COL_MAPPING_TC_TF.get(col.strip(), col) for col in df.columns}
                            df.rename(columns=rename_map, inplace=True)

                            # 合併「彩种」與「玩法」為「彩种玩法」(彩种(玩法))
                            if '彩种' in df.columns and '玩法' in df.columns:
                                df['彩种玩法'] = df['彩种'].astype(str) + '(' + df['玩法'].astype(str) + ')'
                                df.drop(columns=['彩种', '玩法'], inplace=True)
                            elif '彩种' in df.columns:
                                df['彩种玩法'] = df['彩种']
                                df.drop(columns=['彩种'], inplace=True)
                            elif '玩法' in df.columns:
                                df['彩种玩法'] = df['玩法']
                                df.drop(columns=['玩法'], inplace=True)

                            # 倍数處理
                            if '倍数' in df.columns:
                                df['倍数'] = df['倍数'].astype(str)

                        else:
                            rename_map = {}
                            for col in df.columns:
                                stripped_col = col.strip()
                                standard_name = COL_MAPPING_COMMON.get(stripped_col)
                                if standard_name and standard_name != '__TEMP_DISCARD__':
                                    rename_map[col] = standard_name
                                elif stripped_col.startswith('Unnamed:'):
                                    continue
                                elif stripped_col in FINAL_EXPORT_COLS_SIMPLIFIED:
                                    rename_map[col] = stripped_col
                            df.rename(columns=rename_map, inplace=True)

                        # 統一保留最終欄位
                        cols_to_keep = [col for col in FINAL_EXPORT_COLS_SIMPLIFIED if col in df.columns]
                        df = df[cols_to_keep]

                        # 長數字轉字串
                        for col in ['交易号', '期号']:
                            if col in df.columns:
                                df[col] = df[col].astype(str).replace(r'\.0$', '', regex=True)

                        # 補齊缺失欄位
                        for col in FINAL_EXPORT_COLS_SIMPLIFIED:
                            if col not in df.columns:
                                df[col] = ''

                        df = df[FINAL_EXPORT_COLS_SIMPLIFIED]
                        all_dfs.append(df)

                    except Exception as e:
                        print(f"處理站台 {site_name} 的檔案 {os.path.basename(file_path)} 時發生錯誤: {e}")
                        continue

                if not all_dfs:
                    print(f"站台 {site_name} 沒有任何檔案成功處理，跳過。")
                    continue

                combined_df = pd.concat(all_dfs, ignore_index=True)

                if '交易号' in combined_df.columns:
                    combined_df.dropna(subset=['交易号'], inplace=True)
                    combined_df.drop_duplicates(subset=['交易号'], keep='first', inplace=True)

                export_df = combined_df[FINAL_EXPORT_COLS_SIMPLIFIED]
                sheet_name = site_name.strip()
                export_df.to_excel(writer, sheet_name=sheet_name, index=False)
                successful_sites += 1

        if successful_sites == 0:
            messagebox.showerror("錯誤", "所有站台均未成功產生數據")
        else:
            messagebox.showinfo("成功", f"合併完成！共處理 {successful_sites} 個站台，檔案已儲存至：\n{save_path}")

    except Exception as e:
        messagebox.showerror("錯誤", f"匯出失敗：{e}")

# --- 主要程式碼：run_program_8 ---

def run_program_8(root, selected_sites, pages):
    def create_input_window():
        input_window = tk.Toplevel(root)
        input_window.title("投注紀錄全抓取（匯出下載）")
        input_window.geometry("500x550")

        tk.Label(input_window, text="投注紀錄全抓取", font=("Helvetica", 16, "bold")).pack(pady=20)

        # === 起始時間選擇 ===
        tk.Label(input_window, text="開始日期與時間 (YYYY/MM/DD HH:MM:SS)：").pack(anchor="w", padx=50)
        start_frame = tk.Frame(input_window)
        start_frame.pack(pady=5)

        today = datetime.today()
        day_before_yesterday = today - timedelta(days=1)

        start_date_entry = DateEntry(start_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
        start_date_entry.set_date(day_before_yesterday)
        start_date_entry.pack(side=tk.LEFT, padx=5)

        start_hour_combo = ttk.Combobox(start_frame, width=4, values=[f"{i:02d}" for i in range(24)])
        start_hour_combo.pack(side=tk.LEFT, padx=5)
        start_hour_combo.set("03")

        tk.Label(start_frame, text=":").pack(side=tk.LEFT)

        start_minute_combo = ttk.Combobox(start_frame, width=4, values=[f"{i:02d}" for i in range(60)])
        start_minute_combo.pack(side=tk.LEFT, padx=5)
        start_minute_combo.set("00")

        tk.Label(start_frame, text=":").pack(side=tk.LEFT)

        start_second_combo = ttk.Combobox(start_frame, width=4, values=[f"{i:02d}" for i in range(60)])
        start_second_combo.pack(side=tk.LEFT, padx=5)
        start_second_combo.set("00")

        start_manual_entry = tk.Entry(start_frame, width=20)
        start_manual_entry.pack(side=tk.LEFT, padx=10)

        def update_start_manual():
            date = start_date_entry.get()
            h = start_hour_combo.get()
            m = start_minute_combo.get()
            s = start_second_combo.get()
            start_manual_entry.delete(0, tk.END)
            start_manual_entry.insert(0, f"{date} {h}:{m}:{s}")

        start_date_entry.bind("<<DateEntrySelected>>", lambda e: update_start_manual())
        start_hour_combo.bind("<<ComboboxSelected>>", lambda e: update_start_manual())
        start_minute_combo.bind("<<ComboboxSelected>>", lambda e: update_start_manual())
        start_second_combo.bind("<<ComboboxSelected>>", lambda e: update_start_manual())

        # === 結束時間選擇 ===
        tk.Label(input_window, text="結束日期與時間 (YYYY/MM/DD HH:MM:SS)：").pack(anchor="w", padx=50, pady=(20,0))
        end_frame = tk.Frame(input_window)
        end_frame.pack(pady=5)

        yesterday = datetime.today() - timedelta(days=0)

        end_date_entry = DateEntry(end_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
        end_date_entry.set_date(yesterday)
        end_date_entry.pack(side=tk.LEFT, padx=5)

        end_hour_combo = ttk.Combobox(end_frame, width=4, values=[f"{i:02d}" for i in range(24)])
        end_hour_combo.pack(side=tk.LEFT, padx=5)
        end_hour_combo.set("03")

        tk.Label(end_frame, text=":").pack(side=tk.LEFT)

        end_minute_combo = ttk.Combobox(end_frame, width=4, values=[f"{i:02d}" for i in range(60)])
        end_minute_combo.pack(side=tk.LEFT, padx=5)
        end_minute_combo.set("00")

        tk.Label(end_frame, text=":").pack(side=tk.LEFT)

        end_second_combo = ttk.Combobox(end_frame, width=4, values=[f"{i:02d}" for i in range(60)])
        end_second_combo.pack(side=tk.LEFT, padx=5)
        end_second_combo.set("00")

        end_manual_entry = tk.Entry(end_frame, width=20)
        end_manual_entry.pack(side=tk.LEFT, padx=10)

        def update_end_manual():
            date = end_date_entry.get()
            h = end_hour_combo.get()
            m = end_minute_combo.get()
            s = end_second_combo.get()
            end_manual_entry.delete(0, tk.END)
            end_manual_entry.insert(0, f"{date} {h}:{m}:{s}")

        end_date_entry.bind("<<DateEntrySelected>>", lambda e: update_end_manual())
        end_hour_combo.bind("<<ComboboxSelected>>", lambda e: update_end_manual())
        end_minute_combo.bind("<<ComboboxSelected>>", lambda e: update_end_manual())
        end_second_combo.bind("<<ComboboxSelected>>", lambda e: update_end_manual())

        update_start_manual()
        update_end_manual()

        DOWNLOAD_TIMEOUT_MS = 180000  # 3 分鐘

        def process_records():
            start_str = start_manual_entry.get().strip()
            end_str = end_manual_entry.get().strip()

            try:
                start_dt = datetime.strptime(start_str, '%Y/%m/%d %H:%M:%S')
                end_dt = datetime.strptime(end_str, '%Y/%m/%d %H:%M:%S')
            except ValueError:
                messagebox.showerror("錯誤", "日期時間格式錯誤，請確認格式為 YYYY/MM/DD HH:MM:SS")
                return

            if end_dt <= start_dt:
                messagebox.showerror("錯誤", "結束時間必須大於起始時間")
                return

           # 時段分割邏輯（CJ、CY 分成 24 段，其餘站台分成 6 段）
            if any(site['name'] in ["CJ", "CY"] for site in selected_sites):
                NUM_SEGMENTS = 24
            else:
                NUM_SEGMENTS = 6

            total_seconds = int((end_dt - start_dt).total_seconds())
            if total_seconds <= 0:
                messagebox.showerror("錯誤", "時間範圍無效")
                return

            segment_seconds = total_seconds // NUM_SEGMENTS
            segments = []
            current = start_dt

            for i in range(NUM_SEGMENTS):
                if i == NUM_SEGMENTS - 1:
                    seg_end = end_dt
                else:
                    seg_end = current + timedelta(seconds=segment_seconds)
                    # 對齊到整小時（分秒歸零）
                    seg_end = seg_end.replace(minute=0, second=0, microsecond=0)

                segments.append((
                    current.strftime('%Y/%m/%d %H:%M:%S'),
                    seg_end.strftime('%Y/%m/%d %H:%M:%S')
                ))
                current = seg_end + timedelta(seconds=1)  # 下一段從下一秒開始，避免重疊

                if current > end_dt:
                    break
            
            all_site_files = {}  # {site_name: [download_path1, download_path2, ...]}

            for site, page in zip(selected_sites, pages):
                site_name = site['name']
                is_tc_tf = site_name in ["TC", "TF"]
                site_downloads = []
                all_site_files[site_name] = site_downloads

                # === 導航到 BetRecord 頁面，加入重試機制（至少 3 次）===
                navigated = False
                for attempt in range(1, 4):
                    try:
                        page.goto(f"{site['url']}BetRecord", wait_until="networkidle", timeout=30000)
                        time.sleep(3)
                        
                        # 檢查關鍵元素是否存在，確認是否真的導航成功
                        if is_tc_tf:
                            # TC/TF 檢查 #UpdateStartTime 或 #BetRecordbtn
                            if page.locator('#UpdateStartTime').count() > 0 or page.locator('#BetRecordbtn').count() > 0:
                                navigated = True
                                print(f"站台 {site_name} 第 {attempt} 次嘗試導航成功")
                                break
                        else:
                            # 其他站台檢查 #UpdateStartTime 或 #downloadToExcel
                            if page.locator('#UpdateStartTime').count() > 0 or page.locator('#downloadToExcel').count() > 0:
                                navigated = True
                                print(f"站台 {site_name} 第 {attempt} 次嘗試導航成功")
                                break
                    except Exception as e:
                        print(f"站台 {site_name} 第 {attempt} 次導航失敗: {e}")
                        if attempt < 3:
                            time.sleep(2)
                        else:
                            page.screenshot(path=f"error_{site_name}_navigation_failed.png")
                
                if not navigated:
                    print(f"站台 {site_name} 導航失敗，跳過此站台")
                    continue

                # 所有操作都在主 page 上進行（TC/TF 不再使用 iframe）
                frame = page

                # 狀態設定
                try:
                    if is_tc_tf:
                        # 點擊 TC/TF 的狀態下拉選單（精確限定 LOTTERY 容器）
                        state_container = page.locator("div.col-12.col-md-1.dropdown_list.LOTTERY > dropdown-multiple-checkbox")
                        state_input = state_container.locator("input[readonly][data-bind*='click: controlOptions']")
                        state_input.click()

                        # 等待下拉選單內的選項出現（更穩定的等待方式）
                        page.wait_for_selector("div.col-12.col-md-1.dropdown_list.LOTTERY > dropdown-multiple-checkbox ul li", timeout=10000)

                        # 取消「所有狀態」
                        all_check = state_container.locator("input[data-bind*='checked: checkAll']")
                        if all_check.is_checked():
                            all_check.uncheck()

                        # 正確勾選「已中奖」和「未中奖」（直接找 span 文字後面的 input）
                        state_container.locator("span:text-is('已中奖')").locator("..").locator("input[type='checkbox']").check()
                        state_container.locator("span:text-is('未中奖')").locator("..").locator("input[type='checkbox']").check()

                        # 點擊頁面關閉下拉
                        page.click('body')
                    else:
                        frame.click('#StateSelectedList')
                        frame.wait_for_function("document.querySelector('#StateDropDownMultiple').style.display !== 'none'", timeout=10000)

                        if frame.locator('#StateClickAll').is_checked():
                            frame.locator('#StateClickAll').uncheck()

                        frame.locator("#StateDropDownMultiple label:has(span:text-is('未中奖')) input[type='checkbox']").check()
                        frame.locator("#StateDropDownMultiple label:has(span:text-is('中奖')) input[type='checkbox']").check()

                        frame.click('body')

                    print(f"站台 {site_name} 狀態已設定完成。")
                    time.sleep(1)

                except Exception as e:
                    print(f"站台 {site_name} 狀態設定失敗: {e}")
                    page.screenshot(path=f"error_{site_name}_state_setting.png")
                    continue 

                for idx, (seg_start, seg_end) in enumerate(segments):
                    try:
                        # 清除 + 填入日期
                        frame.fill('#StartTime', "")
                        frame.fill('#EndTime', "")
                        frame.fill('#UpdateStartTime', seg_start)
                        frame.fill('#UpdateEndTime', seg_end)
                        
                        frame.click('body', position={'x': 10, 'y': 10}, force=True)
                        print(f"站台 {site_name} 時段 {idx+1}：{seg_start} ~ {seg_end}")

                        try:
                            frame.wait_for_selector('.xdsoft_datetimepicker', state='hidden', timeout=5000)
                        except:
                            frame.click('body', force=True)
                            time.sleep(1)

                        # 查詢按鈕（TC/TF 在 iframe 內）
                        query_btn = '#querybutton' if is_tc_tf else '#Search'
                        frame.click(query_btn, force=True)
                        time.sleep(4)

                        # 匯出 Excel 按鈕
                        download_btn = '#BetRecordbtn' if is_tc_tf else '#downloadToExcel'

                        download = None
                        download_page = None
                        
                        with page.context.expect_event('page') as new_page_info:
                            with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as download_info:
                                frame.click(download_btn, no_wait_after=True) 
                        
                        download_page = new_page_info.value
                        download = download_info.value 
                        
                        # --- 自訂檔案命名邏輯 (使用站台名稱) ---
                        def format_time_for_filename(time_str):
                            dt = datetime.strptime(time_str, '%Y/%m/%d %H:%M:%S')
                            return dt.strftime('%y%m%d%H%M')

                        start_filename = format_time_for_filename(seg_start)
                        end_filename = format_time_for_filename(seg_end)
                        
                        new_filename = f"{site_name}_{start_filename}_{end_filename}.xlsx"
                        download_path = os.path.join(os.path.expanduser("~/Downloads"), new_filename)
                        
                        download.save_as(download_path)
                        print(f"站台 {site_name} 時段 {idx+1} 下載完成: {download_path}")
                        site_downloads.append(download_path)
                        
                        # 下載完成後，立即關閉新開的分頁
                        if download_page and not download_page.is_closed():
                            download_page.close()
                            print(f"   --> 下載分頁已關閉，回到原始投注紀錄頁面。")
                        
                        time.sleep(2)
                        
                    except TimeoutError as e:
                        print(f"站台 {site_name} 時段 {idx+1} 下載超時")
                        page.screenshot(path=f"error_{site_name}_timeout_segment_{idx+1}.png")
                        continue
                    except Exception as e:
                        print(f"站台 {site_name} 時段 {idx+1} 錯誤: {e}")
                        page.screenshot(path=f"error_{site_name}_segment_{idx+1}.png")

            if all_site_files:
                combine_and_clean_excel(all_site_files, start_dt, end_dt)
            else:
                messagebox.showinfo("提示", "無下載檔案，無法合併。")

            if messagebox.askyesno("繼續", "是否再次執行？"):
                input_window.destroy()
                create_input_window()

        tk.Button(input_window, text="開始執行", command=process_records, bg="#4CAF50", fg="white", font=("Helvetica", 14), height=2).pack(pady=30)

    create_input_window()