import tkinter as tk
from tkinter import messagebox, filedialog
from tkcalendar import DateEntry
import asyncio
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict
import re

# 全域變數 - 讓 TS、SY、YD 共用一個分頁
special_chat_page = None
special_scraped = False
special_lock = None

# ==========================================
# 評分規則配置
# ==========================================
SCORE_RULES = {
    "numbers": {f"{i}分": i for i in range(1, 6)},
    "chinese": {"一分": 1, "二分": 2, "三分": 3, "四分": 4, "五分": 5}
}

def calculate_score(messages, user_name):
    """只判斷有 IP 的用戶訊息，並加強 numbers 規則"""
    final_score = ""
    for msg in messages:
        if user_name not in msg['sender']:
            continue

        content = msg['content'].strip()

        if re.match(r'^\s*[1-5]\s*$', content):
            final_score = int(content)
            continue

        for k, v in SCORE_RULES["numbers"].items():
            if re.search(rf'(?<!\w){k}(?!\w)', content):
                final_score = v
                break

        for k, v in SCORE_RULES["chinese"].items():
            if k in content:
                final_score = v
                break

    return final_score


# ==========================================
# 核心爬蟲邏輯
# ==========================================
async def scrape_site_process(site, page, start_datetime, end_datetime, username=None, password=None):
    site_name = site["name"]
    is_tc_tf = site_name in ["TC", "TF"]
    is_special = site_name in ["TS", "SY", "YD"]

    list_data = []
    detail_data = []

    try:
 # ==================== TS / SY / YD 使用新分頁登入 + 完整抓取邏輯 ====================
        if is_special:
            global special_chat_page, special_scraped, special_lock

            # 建立鎖，確保只開一個分頁
            if special_lock is None:
                special_lock = asyncio.Lock()

            async with special_lock:
                if special_chat_page is None:
                    context = page.context
                    chat_page = await context.new_page()
                    special_chat_page = chat_page

                    login_url = ""   # 如果 YD 是不同網域請再告訴我
                    print(f"[{site_name}] 另開新分頁登入 → {login_url}")
                    await chat_page.goto(login_url, wait_until="networkidle", timeout=60000)

                    await chat_page.wait_for_selector("input[placeholder='帐号']", timeout=30000)

                    await chat_page.fill("input[placeholder='帐号']", username)
                    await chat_page.fill("input[placeholder='密码']", password)

                    await chat_page.click("button:has-text('登录')")
                    await asyncio.sleep(4)

                    chat_record_url = ""
                    print(f"[{site_name}] 登入完成，導航至聊天記錄頁 → {chat_record_url}")
                    await chat_page.goto(chat_record_url, wait_until="networkidle", timeout=60000)
                    await asyncio.sleep(2)
                else:
                    chat_page = special_chat_page
                    print(f"[{site_name}] 重用已登入的分頁（TS/SY/YD 共用一個分頁）")

            # ====================== 抓取流程（只執行一次） ======================
            if not special_scraped:
                special_scraped = True

                # 1. 填日期
                await chat_page.evaluate("""
                    (args) => {
                        const startInput = document.querySelector("input[placeholder='起']");
                        const endInput = document.querySelector("input[placeholder='止']");
                        if (startInput) {
                            startInput.focus();
                            startInput.value = args[0];
                            startInput.dispatchEvent(new Event('input', { bubbles: true }));
                            startInput.dispatchEvent(new Event('change', { bubbles: true }));
                            startInput.dispatchEvent(new Event('blur', { bubbles: true }));
                        }
                        if (endInput) {
                            endInput.focus();
                            endInput.value = args[1];
                            endInput.dispatchEvent(new Event('input', { bubbles: true }));
                            endInput.dispatchEvent(new Event('change', { bubbles: true }));
                            endInput.dispatchEvent(new Event('blur', { bubbles: true }));
                        }
                    }
                """, [start_datetime, end_datetime])

                await asyncio.sleep(1.5)

                # 2. 點擊查詢
                await chat_page.click("button:has-text('搜索')")
                await asyncio.sleep(3)

                # 3. 主列表抓取 + 分頁 + 詳情（保持你原本的邏輯）
                while True:
                    rows = await chat_page.query_selector_all("#app > div > div.layout-main-container > div > div.p-card.p-component.mt-5 > div > div > div > div.p-datatable.p-component.p-datatable-hoverable.p-datatable-scrollable > div.p-datatable-table-container > table tbody tr")
                    if not rows:
                        break

                    for row in rows:
                        cols = await row.query_selector_all("td")
                        if len(cols) < 7:
                            continue

                        s_t = await cols[0].inner_text()
                        e_t = await cols[1].inner_text()
                        merchant_code = (await cols[2].inner_text()).strip().upper()   # 自動轉 TS / SY / YD
                        cs_n = await cols[3].inner_text()
                        u_n = await cols[4].inner_text()

                        detail_btn = await cols[6].query_selector("button")
                        if detail_btn:
                            await detail_btn.click()
                            await asyncio.sleep(1.2)
                            await chat_page.wait_for_selector("body > div.p-dialog-mask .p-datatable-table-container table", timeout=10000)

                            chat_history = []

                            while True:
                                detail_rows = await chat_page.query_selector_all("body > div.p-dialog-mask.p-overlay-mask.p-overlay-mask-enter > div > div.p-dialog-content > div > div.p-datatable.p-component.p-datatable-hoverable.p-datatable-scrollable > div.p-datatable-table-container > table tbody tr")
                                for d_row in detail_rows:
                                    d_cols = await d_row.query_selector_all("td")
                                    if len(d_cols) < 3:
                                        continue
                                    d_t = await d_cols[0].inner_text()
                                    d_s = await d_cols[1].inner_text()

                                    # 圖片與表情符號處理
                                    content_cell = d_cols[2]
                                    d_c = await content_cell.inner_text()

                                    img = await content_cell.query_selector("img")
                                    if img:
                                        src = await img.get_attribute("src")
                                        if src:
                                            d_c = src if src.startswith("http") else f"https://chatbe.cywdsd2505.com{src}"
                                    else:
                                        emotion = await content_cell.query_selector("span.chat-emotion")
                                        if emotion:
                                            class_name = await emotion.get_attribute("class")
                                            if class_name:
                                                d_c = class_name.strip()

                                    ip = await d_cols[3].inner_text() if len(d_cols) > 3 else ""

                                    detail_line = f"[{d_t}] {d_s}: {d_c.strip()}"

                                    detail_data.append({
                                        "站台名稱": merchant_code,
                                        "開始時間": s_t.strip(),
                                        "結束時間": e_t.strip(),
                                        "客服名稱": cs_n.strip(),
                                        "用戶名稱": u_n.strip(),
                                        "詳細內容": detail_line,
                                        "IP": ip.strip()
                                    })

                                    if ip.strip():
                                        chat_history.append({"sender": d_s, "content": d_c})

                                # 詳情彈窗下一頁
                                next_detail = await chat_page.query_selector("body > div.p-dialog-mask.p-overlay-mask.p-overlay-mask-enter button.p-paginator-next:not(.p-disabled)")
                                if next_detail:
                                    await next_detail.click()
                                    await asyncio.sleep(1.5)
                                else:
                                    break

                            score = calculate_score(chat_history, u_n)

                            list_data.append({
                                "站台名稱": merchant_code,
                                "開始時間": s_t.strip(),
                                "結束時間": e_t.strip(),
                                "客服名稱": cs_n.strip(),
                                "用戶名稱": u_n.strip(),
                                "評分": score
                            })

                            await chat_page.keyboard.press("Escape")
                            await asyncio.sleep(1)
                            await chat_page.keyboard.press("Escape")
                            await asyncio.sleep(0.5)

                    # 主列表下一頁
                    next_btn = await chat_page.query_selector("#app > div > div.layout-main-container > div > div.p-card.p-component.mt-5 button.p-paginator-next:not(.p-disabled)")
                    if next_btn:
                        await next_btn.click()
                        await asyncio.sleep(2)
                    else:
                        break

                print(f"[{site_name}] 完成 → 清單 {len(list_data)} 筆，詳細 {len(detail_data)} 筆（含IP）")

            else:
                print(f"[{site_name}] 已由另一站台抓取完畢，跳過重複抓取")

            # 新分頁不會關閉
            return list_data, detail_data

        # ==================== TC/TF 與其他站台（完全維持原本邏輯） ====================
        if is_tc_tf:
            target_url = f"{site['url']}HistoryChat"
            main_table = "#LogList"
            next_btn_selector = "#LogList_next:not(.disabled) > a"
            start_sel = "#start"
            end_sel = "#end"
            search_btn = "#form > div.form-actions > input"
            detail_table = "#LogList"
        else:
            target_url = f"{site['url']}ChatMessageRecord/CustomerServiceChatRecord"
            main_table = "#sessionResult"
            next_btn_selector = "li.next:not(.disabled) a"
            start_sel = "#StartTime"
            end_sel = "#EndTime"
            search_btn = "#SearchForm > div:nth-child(4) > div > button"
            detail_table = "#messageResult"

        print(f"[{site_name}] 導向列表頁 → {target_url}")
        await page.goto(target_url, wait_until="networkidle", timeout=60000)

        await page.fill(start_sel, "")
        await page.fill(end_sel, "")
        await page.fill(start_sel, start_datetime)
        await page.fill(end_sel, end_datetime)
        await page.click(search_btn)
        await asyncio.sleep(2)

        # TC/TF 收集連結 + 其他站台彈窗翻頁的原本邏輯（保持不變）
        if is_tc_tf:
            # ...（你的原本 TC/TF 程式碼）
            collected = []
            while True:
                rows = await page.query_selector_all(f"{main_table} tbody tr")
                if not rows:
                    break

                for row in rows:
                    cols = await row.query_selector_all("td")
                    if len(cols) < 6:
                        continue

                    s_t = await cols[0].inner_text()
                    e_t = await cols[1].inner_text()
                    cs_n = await cols[2].inner_text()
                    u_n = await cols[3].inner_text()

                    detail_link = await cols[5].query_selector("a.btn")
                    if detail_link:
                        href = await detail_link.get_attribute("href")
                        if href:
                            detail_url = href if href.startswith("http") else f"{site['url']}{href.lstrip('/')}"
                            collected.append({
                                "s_t": s_t.strip(),
                                "e_t": e_t.strip(),
                                "cs_n": cs_n.strip(),
                                "u_n": u_n.strip(),
                                "detail_url": detail_url
                            })

                next_btn = await page.query_selector(next_btn_selector)
                if next_btn:
                    await next_btn.click()
                    await asyncio.sleep(1.5)
                else:
                    break

            # TC/TF 後續抓取詳情...
            for item in collected:
                await page.goto(item["detail_url"], wait_until="networkidle")
                await asyncio.sleep(1)

                chat_history = []
                detail_rows = await page.query_selector_all(f"{detail_table} tbody tr")

                for d_row in detail_rows:
                    d_cols = await d_row.query_selector_all("td")
                    if len(d_cols) < 3:
                        continue
                    d_t = await d_cols[0].inner_text()
                    d_s = await d_cols[1].inner_text()
                    d_c = await d_cols[2].inner_text()
                    ip = await d_cols[3].inner_text() if len(d_cols) > 3 else ""

                    detail_line = f"[{d_t}] {d_s}: {d_c.strip()}"

                    detail_data.append({
                        "站台名稱": site_name,
                        "開始時間": item["s_t"],
                        "結束時間": item["e_t"],
                        "客服名稱": item["cs_n"],
                        "用戶名稱": item["u_n"],
                        "詳細內容": detail_line,
                        "IP": ip.strip()
                    })

                    if ip.strip():
                        chat_history.append({"sender": d_s, "content": d_c})

                score = calculate_score(chat_history, item["u_n"])

                list_data.append({
                    "站台名稱": site_name,
                    "開始時間": item["s_t"],
                    "結束時間": item["e_t"],
                    "客服名稱": item["cs_n"],
                    "用戶名稱": item["u_n"],
                    "評分": score
                })

        else:
            # 其他站台原本的彈窗 + 翻頁邏輯（完全不動）
            while True:
                rows = await page.query_selector_all(f"{main_table} tbody tr")
                if not rows:
                    break

                for row in rows:
                    cols = await row.query_selector_all("td")
                    if len(cols) < 6:
                        continue

                    s_t = await cols[0].inner_text()
                    e_t = await cols[1].inner_text()
                    cs_n = await cols[2].inner_text()
                    u_n = await cols[3].inner_text()

                    detail_btn = await cols[5].query_selector("button")
                    if detail_btn:
                        await detail_btn.click()
                        await asyncio.sleep(0.8)

                        try:
                            await page.wait_for_selector(f"{detail_table} tbody tr", timeout=8000)
                        except:
                            print(f"[{site_name}] 詳情彈窗載入逾時，跳過本筆")
                            await page.keyboard.press("Escape")
                            await asyncio.sleep(1)
                            continue

                        chat_history = []

                        # === 詳情視窗內翻頁 + 特殊內容處理 ===
                        while True:
                            detail_rows = await page.query_selector_all(f"{detail_table} tbody tr")

                            for d_row in detail_rows:
                                d_cols = await d_row.query_selector_all("td")
                                if len(d_cols) < 3:
                                    continue

                                d_t = await d_cols[0].inner_text()
                                d_s = await d_cols[1].inner_text()

                                content_cell = d_cols[2]
                                d_c = await content_cell.inner_text()

                                # 圖片連結
                                link = await content_cell.query_selector("a[target='_blank']")
                                if link:
                                    href = await link.get_attribute("href")
                                    if href:
                                        d_c = href if href.startswith("http") else f"{site['url'].rstrip('/')}{href}"
                                else:
                                    # 表情符號
                                    emotion = await content_cell.query_selector("span.chat-emotion")
                                    if emotion:
                                        class_name = await emotion.get_attribute("class")
                                        if class_name:
                                            d_c = class_name.strip()

                                ip = await d_cols[3].inner_text() if len(d_cols) > 3 else ""

                                detail_line = f"[{d_t}] {d_s}: {d_c.strip()}"

                                detail_data.append({
                                    "站台名稱": site_name,
                                    "開始時間": s_t.strip(),
                                    "結束時間": e_t.strip(),
                                    "客服名稱": cs_n.strip(),
                                    "用戶名稱": u_n.strip(),
                                    "詳細內容": detail_line,
                                    "IP": ip.strip()
                                })

                                if ip.strip():
                                    chat_history.append({"sender": d_s, "content": d_c})

                            # 詳情下一頁
                            next_detail_btn = await page.query_selector("#messageResult_next > a")
                            if next_detail_btn:
                                is_disabled = await next_detail_btn.evaluate(
                                    "el => el.parentElement.classList.contains('disabled')"
                                )
                                if is_disabled:
                                    break
                                await next_detail_btn.click()
                                await asyncio.sleep(1.5)
                            else:
                                break

                        score = calculate_score(chat_history, u_n)

                        list_data.append({
                            "站台名稱": site_name,
                            "開始時間": s_t.strip(),
                            "結束時間": e_t.strip(),
                            "客服名稱": cs_n.strip(),
                            "用戶名稱": u_n.strip(),
                            "評分": score
                        })

                        # ====================== 加強關閉彈窗 ======================
                        try:
                            await page.keyboard.press("Escape")
                            await asyncio.sleep(0.8)
                            await page.keyboard.press("Escape")
                            await asyncio.sleep(0.6)
                            # 額外保險：點擊彈窗外背景區域
                            await page.click("body", position={"x": 10, "y": 10})
                            await asyncio.sleep(0.5)
                        except:
                            pass

                        # 確認彈窗是否真的關閉
                        try:
                            await page.wait_for_selector("body > div.p-dialog-mask", state="hidden", timeout=3000)
                        except:
                            print(f"[{site_name}] 彈窗關閉可能失敗，已強制繼續下一筆")

                # 主列表翻頁
                next_btn = await page.query_selector(next_btn_selector)
                if next_btn:
                    await next_btn.click()
                    await asyncio.sleep(1.5)
                else:
                    break

        print(f"[{site_name}] 完成 → 清單 {len(list_data)} 筆，詳細 {len(detail_data)} 筆（含IP）")

    except Exception as e:
        print(f"[{site_name}] 發生錯誤: {e}")

    return list_data, detail_data


# ==========================================
# UI 介面與任務啟動
# ==========================================
def run_program_11(root, selected_sites, pages, callback):
    def create_input_window():
        input_window = tk.Toplevel(root)
        input_window.title("Program 11 - 客服評分抓取")
        input_window.geometry("520x400")
        input_window.grab_set()

        tk.Label(input_window, text="選擇查詢區間", font=("Arial", 12, "bold")).pack(pady=15)

        frame_input = tk.Frame(input_window)
        frame_input.pack(pady=10, padx=20)

        tk.Label(frame_input, text="開始日期：", font=("Arial", 10)).grid(row=0, column=0, sticky="e", pady=6)
        start_date_ent = DateEntry(frame_input, date_pattern='yyyy-mm-dd', width=14)
        start_date_ent.grid(row=0, column=1, padx=8)

        tk.Label(frame_input, text="開始時間：", font=("Arial", 10)).grid(row=1, column=0, sticky="e", pady=6)
        start_time_var = tk.StringVar(value="03:00:00")
        start_time_entry = tk.Entry(frame_input, textvariable=start_time_var, width=12, font=("Arial", 10))
        start_time_entry.grid(row=1, column=1, padx=8, sticky="w")

        tk.Label(frame_input, text="結束日期：", font=("Arial", 10)).grid(row=0, column=2, sticky="e", pady=6, padx=(30,0))
        end_date_ent = DateEntry(frame_input, date_pattern='yyyy-mm-dd', width=14)
        end_date_ent.grid(row=0, column=3, padx=8)

        tk.Label(frame_input, text="結束時間：", font=("Arial", 10)).grid(row=1, column=2, sticky="e", pady=6, padx=(30,0))
        end_time_var = tk.StringVar(value="03:00:00")
        end_time_entry = tk.Entry(frame_input, textvariable=end_time_var, width=12, font=("Arial", 10))
        end_time_entry.grid(row=1, column=3, padx=8, sticky="w")

        today = datetime.now()
        yesterday = today - timedelta(days=1)
        start_date_ent.set_date(yesterday)
        end_date_ent.set_date(today)

        # 在日期輸入框下方加入
        special_site_var = tk.BooleanVar(value=True) # 預設勾選
        tk.Checkbutton(input_window, text="開BC-客服整合系統(TS/SY/YD)", 
                       variable=special_site_var, font=("Arial", 10)).pack(pady=10)

        tk.Label(input_window, 
                 text="時間格式：HH:mm:ss\n預設為 03:00:00",
                 font=("Arial", 10), fg="blue").pack(pady=12)

        async def _async_process():
            start_d = start_date_ent.get()
            end_d = end_date_ent.get()
            start_t = start_time_var.get().strip()
            end_t = end_time_var.get().strip()

            start_datetime = f"{start_d} {start_t}"
            end_datetime = f"{end_d} {end_t}"

            print(f"查詢時間範圍：{start_datetime} ~ {end_datetime}")

            tasks = []
            # 取得用戶是否勾選了抓取特殊站台
            do_special = special_site_var.get()

            for site, page in zip(selected_sites, pages):
                site_name = site["name"]
                
                # 如果站台屬於特殊群組
                if site_name in ["TS", "SY", "YD"]:
                    if do_special:
                        tasks.append(scrape_site_process(site, page, start_datetime, end_datetime, 
                                                        username="", 
                                                        password=""))
                    else:
                        print(f"[{site_name}] 選擇跳過抓取")
                else:
                    # 普通站台照常執行
                    tasks.append(scrape_site_process(site, page, start_datetime, end_datetime))
            try:
                all_results = await asyncio.gather(*tasks)
                list_results = []
                detail_results = []
                for list_d, detail_d in all_results:
                    list_results.extend(list_d)
                    detail_results.extend(detail_d)

                root.after(0, lambda: finalize_save(list_results, detail_results))

            except Exception as e:
                root.after(0, lambda: (
                    messagebox.showerror("錯誤", f"執行出錯: {str(e)}"),
                    input_window.destroy(),
                    callback() if callback else None
                ))

        def finalize_save(list_data, detail_data):
            if not list_data and not detail_data:
                messagebox.showinfo("完成", "所選區間內無聊天紀錄。")
                input_window.destroy()
                if callback:
                    callback()
                return

            list_groups = defaultdict(list)
            for row in list_data:
                list_groups[row["站台名稱"]].append(row)

            detail_groups = defaultdict(list)
            for row in detail_data:
                detail_groups[row["站台名稱"]].append(row)

            timestamp = datetime.now().strftime('%m%d%H%M')
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"P11_客服結果_{timestamp}.xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

            if save_path:
                try:
                    with pd.ExcelWriter(save_path) as writer:
                        for site_name, data in list_groups.items():
                            pd.DataFrame(data).to_excel(writer, sheet_name=f"{site_name}_客服聊天列表", index=False)
                        for site_name, data in detail_groups.items():
                            pd.DataFrame(data).to_excel(writer, sheet_name=f"{site_name}_詳細對話內容", index=False)
# ====================== 修改完成訊息 ======================
                    msg = f"已儲存至 {save_path}\n\n"
                    msg += "=== 各站台聊天列表筆數 ===\n"
                    
                    # 按照常用順序排序顯示
                    site_order = ["TC", "TF", "TS", "SY", "YD", "FL", "WX", "XC", "XH", "CJ", "CY"]
                    for site in site_order:
                        count = len(list_groups.get(site, []))
                        if count > 0:
                            msg += f"{site}: {count} 筆聊天列表\n"
                    
                    # 顯示其他站台（如果有）
                    for site, data in list_groups.items():
                        if site not in site_order and len(data) > 0:
                            msg += f"{site}: {len(data)} 筆聊天列表\n"

                    messagebox.showinfo("完成", msg)

                except Exception as e:
                    messagebox.showerror("儲存錯誤", f"Excel 儲存失敗: {e}")

            input_window.destroy()
            if callback:
                callback()

        def start_task():
            from __main__ import app
            app.run_async(_async_process())

        tk.Button(
            input_window,
            text="開始抓取數據",
            command=start_task,
            bg="#2196F3",
            fg="white",
            font=("Arial", 12, "bold"),
            height=2,
            width=24
        ).pack(pady=25)

    create_input_window()
