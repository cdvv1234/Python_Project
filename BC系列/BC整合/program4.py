import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import font as tkFont
from tkinter import ttk
from tkcalendar import DateEntry
from playwright.sync_api import TimeoutError
import pandas as pd
import time
from datetime import datetime, timedelta
import re

# 通用的文字清理函數
def clean_text(text):
    """清理文字欄位，移除多餘空行並將多個空格替換為單個空格"""
    if isinstance(text, str):
        # 移除多餘的空行和換行符
        text = re.sub(r'\n+', ' ', text)
        # 將多個空格替換為單個空格
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    return text

def select_dates(parent):
    selected_data = {"start_date": "", "end_date": ""}
    window = tk.Toplevel(parent)
    window.title("選擇日期")
    window.geometry("400x300")
    font_style = tkFont.Font(size=10)

    today = datetime.now()
    day_before_yesterday = today - timedelta(days=2)
    yesterday = today - timedelta(days=1)

    tk.Label(window, text="開始日期與時間 (YYYY/MM/DD HH:MM)：", font=font_style).pack(pady=5)
    start_frame = tk.Frame(window)
    start_frame.pack(pady=5)
    
    start_date_entry = DateEntry(start_frame, width=12, font=font_style, date_pattern='yyyy/mm/dd')
    start_date_entry.set_date(day_before_yesterday)
    start_date_entry.pack(side=tk.LEFT, padx=5)
    
    start_hour_combo = ttk.Combobox(start_frame, width=4, values=[f"{i:02d}" for i in range(24)], font=font_style)
    start_hour_combo.pack(side=tk.LEFT, padx=5)
    start_hour_combo.set("03")
    
    tk.Label(start_frame, text=":", font=font_style).pack(side=tk.LEFT)
    
    start_minute_combo = ttk.Combobox(start_frame, width=4, values=[f"{i:02d}" for i in range(0, 60, 1)], font=font_style)
    start_minute_combo.pack(side=tk.LEFT, padx=5)
    start_minute_combo.set("00")
    
    start_manual_entry = tk.Entry(start_frame, width=16, font=font_style)
    start_manual_entry.pack(side=tk.LEFT, padx=5)
    
    def update_start_manual():
        date = start_date_entry.get()
        hour = start_hour_combo.get()
        minute = start_minute_combo.get()
        start_manual_entry.delete(0, tk.END)
        start_manual_entry.insert(0, f"{date} {hour}:{minute}")
    
    start_date_entry.bind("<<DateEntrySelected>>", lambda e: update_start_manual())
    start_hour_combo.bind("<<ComboboxSelected>>", lambda e: update_start_manual())
    start_minute_combo.bind("<<ComboboxSelected>>", lambda e: update_start_manual())

    tk.Label(window, text="結束日期與時間 (YYYY/MM/DD HH:MM)：", font=font_style).pack(pady=5)
    end_frame = tk.Frame(window)
    end_frame.pack(pady=5)
    
    end_date_entry = DateEntry(end_frame, width=12, font=font_style, date_pattern='yyyy/mm/dd')
    end_date_entry.set_date(yesterday)
    end_date_entry.pack(side=tk.LEFT, padx=5)
    
    end_hour_combo = ttk.Combobox(end_frame, width=4, values=[f"{i:02d}" for i in range(24)], font=font_style)
    end_hour_combo.pack(side=tk.LEFT, padx=5)
    end_hour_combo.set("03")
    
    tk.Label(end_frame, text=":", font=font_style).pack(side=tk.LEFT)
    
    end_minute_combo = ttk.Combobox(end_frame, width=4, values=[f"{i:02d}" for i in range(0, 60, 1)], font=font_style)
    end_minute_combo.pack(side=tk.LEFT, padx=5)
    end_minute_combo.set("00")
    
    end_manual_entry = tk.Entry(end_frame, width=16, font=font_style)
    end_manual_entry.pack(side=tk.LEFT, padx=5)
    
    def update_end_manual():
        date = end_date_entry.get()
        hour = end_hour_combo.get()
        minute = end_minute_combo.get()
        end_manual_entry.delete(0, tk.END)
        end_manual_entry.insert(0, f"{date} {hour}:{minute}")
    
    end_date_entry.bind("<<DateEntrySelected>>", lambda e: update_end_manual())
    end_hour_combo.bind("<<ComboboxSelected>>", lambda e: update_end_manual())
    end_minute_combo.bind("<<ComboboxSelected>>", lambda e: update_end_manual())

    def confirm_selection():
        selected_data["start_date"] = start_manual_entry.get()
        selected_data["end_date"] = end_manual_entry.get()
        
        try:
            datetime.strptime(selected_data["start_date"], '%Y/%m/%d %H:%M')
            datetime.strptime(selected_data["end_date"], '%Y/%m/%d %H:%M')
            window.destroy()
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的日期格式 (YYYY/MM/DD HH:MM)")

    submit_button = tk.Button(window, text="確定", command=confirm_selection, font=font_style)
    submit_button.pack(pady=10)

    update_start_manual()
    update_end_manual()

    window.wait_window()
    return selected_data

# TC 和 TF 的頁面路徑與對應渠道
tc_tf_page_paths = [
    {"path": "WithdrawRiskControl", "channel": "草"},
    {"path": "WithdrawRiskControl/VcpIndex", "channel": "U"},
    {"path": "ElectronicPurseWithdrawExamination/RiskIndex", "channel": "GB"}
]

# 其他站台（包括 CJ）的頁面路徑與對應渠道
other_page_paths = [
    {"path": "WithdrawExamination/RiskIndex", "channel": "草"},
    {"path": "USDTWithdrawExamination/RiskIndex", "channel": "U"},
    {"path": "ElectronicPurseWithdrawExamination/RiskIndex", "channel": "GB"}
]

def scrape_site_page(site, page, page_path, channel, start_date, end_date, is_tc_tf=False):
    results = []
    full_url = f"{site['url'].rstrip('/')}/{page_path}"

    max_retries = 2
    for attempt in range(max_retries):
        try:
            page.goto(full_url, wait_until="domcontentloaded", timeout=30000)
            time.sleep(3)  # 等待頁面完全加載
            final_url = page.url
            if final_url.rstrip('/') == site["url"].rstrip('/'):
                print(f"站台 {site['name']} 的 {page_path} 重定向到首頁: {final_url}")
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_navigation.png")
                return results
            break
        except TimeoutError as e:
            print(f"站台 {site['name']} 導航到 {full_url} 超時 (嘗試 {attempt + 1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_navigation.png")
                return results
        except Exception as e:
            print(f"站台 {site['name']} 導航到 {full_url} 失敗 (嘗試 {attempt + 1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_navigation.png")
                return results
            time.sleep(2)

    try:
        time.sleep(2)  # 等待頁面元素加載
        iframe_elements = page.locator("iframe").all()
        if iframe_elements:
            print(f"站台 {site['name']} 找到的 iframe: {[elem.get_attribute('name') or elem.get_attribute('src') for elem in iframe_elements]}")
        else:
            print(f"站台 {site['name']} 未找到任何 iframe，直接在主頁面操作")

        try:
            close_button = page.locator("#Warning-Dialog-CloseBtn")
            if close_button.is_visible(timeout=5000):
                close_button.click()
                print(f"站台 {site['name']} 關閉彈窗")
                time.sleep(2)  # 等待彈窗關閉
        except Exception as e:
            print(f"站台 {site['name']} 檢查彈窗失敗: {e}")

        if is_tc_tf:
            try:
                page.wait_for_selector("#StartCreateTime", state="visible", timeout=10000)
                time.sleep(1)  # 等待輸入框完全加載
                page.fill("#StartCreateTime", "")
                page.fill("#EndCreateTime", "")
                time.sleep(1)  # 等待清空操作完成
            except Exception as e:
                print(f"站台 {site['name']} 清空申請日期失敗: {e}")
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_input.png")

            try:
                page.wait_for_selector("#StartConfirmTime", state="visible", timeout=10000)
                time.sleep(1)  # 等待輸入框完全加載
                page.fill("#StartConfirmTime", start_date)
                page.fill("#EndConfirmTime", end_date)
                time.sleep(1)  # 等待填入操作完成
                start_value = page.input_value("#StartConfirmTime")
                end_value = page.input_value("#EndConfirmTime")
                if start_value != start_date or end_value != end_date:
                    print(f"站台 {site['name']} 日期填入失敗: StartConfirmTime={start_value}, EndConfirmTime={end_value}")
                    page.click("#StartConfirmTime")
                    time.sleep(0.5)  # 等待點擊響應
                    page.type("#StartConfirmTime", start_date)
                    page.click("#EndConfirmTime")
                    time.sleep(0.5)  # 等待點擊響應
                    page.type("#EndConfirmTime", end_date)
                    time.sleep(1)  # 等待輸入完成
            except Exception as e:
                print(f"站台 {site['name']} 填入確認日期失敗: {e}")
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_input.png")

            try:
                page.wait_for_selector("#search > span", state="visible", timeout=20000)
                time.sleep(1)  # 等待按鈕完全加載
                page.click("#search > span")
                time.sleep(3)  # 等待查詢結果加載
            except Exception as e:
                print(f"站台 {site['name']} 點擊查詢按鈕失敗: {e}")
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_search.png")

        else:
            try:
                page.wait_for_selector("#StartCreateTime", state="visible", timeout=10000)
                time.sleep(1)  # 等待輸入框完全加載
                page.fill("#StartCreateTime", "")
                page.fill("#EndCreateTime", "")
                time.sleep(1)  # 等待清空操作完成
            except Exception as e:
                print(f"站台 {site['name']} 清空申請日期失敗: {e}")
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_input.png")

            try:
                page.wait_for_selector("#StartRiskControlConfirmTime", state="visible", timeout=10000)
                time.sleep(1)  # 等待輸入框完全加載
                page.fill("#StartRiskControlConfirmTime", start_date)
                page.fill("#EndRiskControlConfirmTime", end_date)
                time.sleep(1)  # 等待填入操作完成
                start_value = page.input_value("#StartRiskControlConfirmTime")
                end_value = page.input_value("#EndRiskControlConfirmTime")
                if start_value != start_date or end_value != end_date:
                    print(f"站台 {site['name']} 日期填入失敗: StartRiskControlConfirmTime={start_value}, EndRiskControlConfirmTime={end_value}")
                    page.click("#StartRiskControlConfirmTime")
                    time.sleep(0.5)  # 等待點擊響應
                    page.type("#StartRiskControlConfirmTime", start_date)
                    page.click("#EndRiskControlConfirmTime")
                    time.sleep(0.5)  # 等待點擊響應
                    page.type("#EndRiskControlConfirmTime", end_date)
                    time.sleep(1)  # 等待輸入完成
            except Exception as e:
                print(f"站台 {site['name']} 填入確認日期失敗: {e}")
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_input.png")

            search_button_selectors = [
                "button:contains('查询')",
                "button:contains('Search')",
                "#queryForm > div.form-group.row > div > button"
            ]
            search_button = None
            for selector in search_button_selectors:
                try:
                    search_button = page.locator(selector)
                    search_button.wait_for(state="visible", timeout=30000)
                    break
                except:
                    continue

            if not search_button:
                print(f"站台 {site['name']} 的頁面 {page_path} 未找到查詢按鈕，選擇器: {search_button_selectors}")
                buttons = page.evaluate("""
                    () => {
                        const btns = Array.from(document.querySelectorAll('button'));
                        return btns.map(btn => ({
                            text: btn.innerText.trim(),
                            class: btn.className,
                            id: btn.id
                        }));
                    }
                """)
                print(f"站台 {site['name']} 的 button 元素: {buttons}")
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_search.png")
                return results

            time.sleep(1)  # 等待按鈕完全加載
            search_button.click()
            time.sleep(3)  # 等待查詢結果加載

        while True:
            try:
                page.wait_for_selector("#RequestTable", state="visible", timeout=20000)
                time.sleep(2)  # 等待表格完全加載
            except Exception as e:
                print(f"站台 {site['name']} 等待表格失敗: {e}")
                page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}_table.png")
                break

            empty_table = page.evaluate("""
                () => {
                    const td = document.querySelector('#RequestTable tbody tr td');
                    return td ? td.innerText.trim() : '';
                }
            """)
            if "表中数据为空" in empty_table:
                print(f"站台 {site['name']} 表格為空: {empty_table}")
                break

            headers = page.evaluate("""
                () => {
                    const header_cells = Array.from(document.querySelectorAll('#RequestTable thead th'));
                    return header_cells.map(cell => cell.innerText.trim().replace(/\\n/g, ' '));
                }
            """)
            # 清理表頭
            cleaned_headers = [clean_text(header) for header in headers]
            print(f"站台 {site['name']} 的頁面 {page_path} 表頭: {cleaned_headers}")

            rows_data = page.evaluate("""
                () => {
                    const rows = Array.from(document.querySelectorAll('#RequestTable tbody tr'));
                    return rows.map(row => {
                        const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText.trim());
                        return cells;
                    });
                }
            """)
            # 清理每一行數據
            cleaned_rows_data = [[clean_text(cell) for cell in row] for row in rows_data]
            print(f"站台 {site['name']} 的頁面 {page_path} 提取的數據行數: {len(cleaned_rows_data)}")

            for row in cleaned_rows_data:
                if len(row) < 4:
                    print(f"站台 {site['name']} 的頁面 {page_path} 發現無效數據行: {row}")
                    continue
                results.append({"headers": cleaned_headers, "data": row, "channel": channel, "page_path": page_path})
                print(f"站台 {site['name']} 的頁面 {page_path} 數據行: {row}")

            next_button = page.locator("#RequestTable_next > a")
            parent_li = page.locator("#RequestTable_next")
            if (parent_li.count() and "disabled" in parent_li.get_attribute("class", timeout=5000)) or \
               next_button.get_attribute("disabled", timeout=5000) or not next_button.is_visible():
                break

            time.sleep(1)  # 等待按鈕完全加載
            next_button.click()
            time.sleep(2)  # 等待下一頁加載

    except Exception as e:
        print(f"處理 {site['name']} 的 {page_path} 時發生錯誤: {e}")
        page.screenshot(path=f"error_{site['name']}_{page_path.replace('/', '_')}.png")

    return results

def process_site_data(site_name, raw_data, channel, page=None):
    processed_data = []
    target_columns = ["状态", "操作者", "申请日期", "确认日期"]
    column_variants = {
        "状态": ["状态", "状态码", "状态名称", "审核状态", "處理狀態", "審核狀態", "審查狀態"],
        "操作者": ["操作者", "操作人员", "操作人", "审核人", "經辦人", "審核人員", "處理人員"],
        "申请日期": ["申请日期", "申请时间", "提交日期", "提交时间", "创建时间", "申請日期", "申請時間", "創建時間"],
        "确认日期": ["确认日期", "确认时间", "审核日期", "审核时间", "处理时间", "確認日期", "確認時間", "審核時間", "處理時間"]
    }

    for row in raw_data:
        headers = row["headers"]
        data = row["data"]
        channel = row["channel"]
        page_path = row["page_path"]

        print(f"處理站台 {site_name} 的頁面 {page_path}，渠道 {channel}，表頭: {headers}")
        print(f"數據行: {data}")

        column_indices = {}
        matched_headers = {}
        for col in target_columns:
            for variant in column_variants[col]:
                for i, header in enumerate(headers):
                    if header == variant:
                        column_indices[col] = i
                        matched_headers[col] = header
                        break
                if col in column_indices:
                    break

        if page_path == "WithdrawRiskControl/VcpIndex" and site_name in ["TC", "TF"]:
            print(f"站台 {site_name} 的 U 渠道初始匹配表頭: {matched_headers}")
            print(f"初始欄位索引: {column_indices}")
            if "状态" in column_indices and "申请日期" in column_indices:
                status_index = column_indices["状态"]
                apply_date_index = column_indices["申请日期"]
                if apply_date_index >= status_index + 1:
                    column_indices["操作者"] = status_index + 1
                    matched_headers["操作者"] = f"隱藏欄位（索引 {status_index + 1}）"
                    print(f"站台 {site_name} 的 U 渠道推斷「操作者」索引: {column_indices['操作者']}，數據: {data[column_indices['操作者']]}")

                    operator_index = column_indices["操作者"]
                    if "申请日期" in column_indices and column_indices["申请日期"] >= operator_index:
                        column_indices["申请日期"] = operator_index + 1
                    if "确认日期" in column_indices and column_indices["确认日期"] >= operator_index:
                        column_indices["确认日期"] = operator_index + 2
                    print(f"調整後欄位索引: {column_indices}")
                else:
                    print(f"站台 {site_name} 的 U 渠道索引異常，狀態索引: {status_index}，申請日期索引: {apply_date_index}")
                    if page is not None:
                        page.screenshot(path=f"error_{site_name}_VcpIndex_columns.png")
                    continue
            else:
                print(f"站台 {site_name} 的 U 渠道缺少「狀態」或「申請日期」，表頭: {headers}")
                if page is not None:
                    page.screenshot(path=f"error_{site_name}_VcpIndex_columns.png")
                continue

        missing_columns = [col for col in target_columns if col not in column_indices and col != "操作者"]
        if missing_columns:
            print(f"站台 {site_name} 的頁面 {page_path} 缺少關鍵欄位: {missing_columns}，表頭: {headers}")
            if page is not None:
                page.screenshot(path=f"error_{site_name}_{page_path.replace('/', '_')}_columns.png")
            continue

        max_index = max(column_indices.values(), default=0) if column_indices else 0
        if len(data) <= max_index:
            print(f"站台 {site_name} 的頁面 {page_path} 數據長度不足: {data}，預期最小長度: {max_index + 1}")
            continue

        processed_row = {
            "平台": site_name,
            "渠道": channel
        }
        for col in target_columns:
            processed_row[col] = data[column_indices[col]] if col in column_indices and len(data) > column_indices[col] else ""
        processed_data.append(processed_row)

        print(f"站台 {site_name} 的頁面 {page_path} 最終匹配表頭: {matched_headers}")
        print(f"處理後數據: {processed_row}")

    return processed_data

def save_results_to_excel(all_results_dict):
    timestamp = datetime.now().strftime("%Y%m%d")
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        initialfile=f"results_{timestamp}.xlsx",
        title="另存查詢結果"
    )
    if save_path:
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            for site_name, results in all_results_dict.items():
                if results:
                    df = pd.DataFrame(results)
                    df = df[["平台", "渠道", "状态", "操作者", "申请日期", "确认日期"]]
                    df = df.applymap(clean_text)  # 確保最終輸出乾淨
                    df.to_excel(writer, index=False, sheet_name=site_name)
        messagebox.showinfo("完成", f"資料已成功保存至：{save_path}")

def run_program(root, selected_sites, pages):
    selected_data = select_dates(root)
    if not selected_data["start_date"] or not selected_data["end_date"]:
        print("未輸入日期，程式結束。")
        return

    start_date = selected_data["start_date"]
    end_date = selected_data["end_date"]

    all_results_dict = {}
    for site, page in zip(selected_sites, pages):
        site_results = []
        page_paths = tc_tf_page_paths if site["name"] in ["TC", "TF"] else other_page_paths
        is_tc_tf = site["name"] in ["TC", "TF"]
        
        for page_info in page_paths:
            raw_data = scrape_site_page(
                site, page, page_info["path"], page_info["channel"],
                start_date, end_date, is_tc_tf=is_tc_tf
            )
            processed_data = process_site_data(
                site["name"], raw_data, page_info["channel"], page=page
            )
            site_results.extend(processed_data)
        all_results_dict[site["name"]] = site_results

    if any(all_results_dict.values()):
        save_results_to_excel(all_results_dict)

    print("run_program 執行完畢（program4）")
