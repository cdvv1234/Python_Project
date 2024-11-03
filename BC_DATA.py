import tkinter as tk
from tkinter import simpledialog
from tkinter import font as tkFont
from playwright.sync_api import sync_playwright
import pandas as pd
import time
import os

# 站点配置
sites = [
    {
        #
        "name": "",
        "url": "",
        "login_selector": "iframe[name='mainFrame']",
        "input_selector": "#LoginID",
        "button_selector": "button[name='查询']",
        "table_selector": "#TeamProfitTable",
        "columns": ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
    },
    {   
        #
        "name": "",
        "url": "",
        "login_selector": "iframe[name='mainFrame']",
        "input_selector": "#LoginID",
        "button_selector": "button[name='查询']",
        "table_selector": "#TeamProfitTable",
        "columns": ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
    },
    {  
        #
        "name": "",
        "url": "",
        "login_selector": "body",
        "input_selector": "#LoginId",
        "button_selector": 'button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]',
        "table_selector": 'tbody[data-bind="with: TeamStatistics"]',
        "columns": ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利', '平台服务费']
    },
    {   
        #
        "name": "",
        "url": "",
        "login_selector": "body",
        "input_selector": "#LoginId",
        "button_selector": 'button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]',
        "table_selector": 'tbody[data-bind="with: TeamStatistics"]',
        "columns": ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
    },
    {
        #
        "name": "",
        "url": "", 
        "login_selector": "body",
        "input_selector": "#LoginId",
        "button_selector": 'button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]',
        "table_selector": 'tbody[data-bind="with: TeamStatistics"]',
        "columns": ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
    },
    {
        #
        "name": "",
        "url": "",  
        "login_selector": "body",
        "input_selector": "#LoginId",
        "button_selector": 'button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]',
        "table_selector": 'tbody[data-bind="with: TeamStatistics"]',
        "columns": ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
    },
    {
        #
        "name": "",
        "url": "",  
        "login_selector": "body",
        "input_selector": "#LoginId",
        "button_selector": 'button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]',
        "table_selector": 'tbody[data-bind="with: TeamStatistics"]',
        "columns": ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利']
    },
    {
        #
        "name": "",
        "url": "",  
        "login_selector": "body",
        "input_selector": "#LoginId",
        "button_selector": 'button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]',
        "table_selector": 'tbody[data-bind="with: TeamStatistics"]',
        "columns": ['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利', '平台服务费']
    }
]

def get_selected_sites():
    selected_sites = []
    root = tk.Tk()
    root.title("BC")

    # 設定視窗大小
    root.geometry("400x300")

    # 設定字體大小
    font_style = tkFont.Font(size=16)  # 調整字體大小

    var_list = []
    for index, site in enumerate(sites):
        var = tk.BooleanVar(value=True)  # 預設勾選
        chk = tk.Checkbutton(root, text=site["name"], variable=var, font=font_style)  # 設定字體
        # 使用 grid 方法將選項排成四個一排並置中，調整 pady 來減少間距
        chk.grid(row=index // 4, column=index % 4, sticky='nsew', padx=5, pady=1)  # 減小 pady

        var_list.append(var)

    # 調整列和行的權重，讓每個選項和按鈕置中
    for i in range(4):
        root.grid_columnconfigure(i, weight=1)
    for i in range((len(sites) + 3) // 4):  # 動態調整行數
        root.grid_rowconfigure(i, weight=1)

    def on_submit():
        for var, site in zip(var_list, sites):
            if var.get():
                selected_sites.append(site)
        root.destroy()

    submit_button = tk.Button(root, text="提交", command=on_submit, font=font_style)  # 設定字體
    submit_button.grid(row=(len(sites) + 3) // 4, column=0, columnspan=4, pady=10)  # 按鈕在最下方，橫跨四列並加點間距

    root.mainloop()
    return selected_sites

def get_accounts_from_user(site_name):
    accounts = []
    root = tk.Tk()
    root.title("帳號列表")
    
    label = tk.Label(root, text=f"請貼上 '{site_name}' 的帳號列表（每個帳號一行）:")
    label.pack()

    text_box = tk.Text(root, height=15, width=50)
    text_box.pack()

    def on_submit():
        nonlocal accounts
        user_input = text_box.get("1.0", "end-1c")
        accounts = [account.strip() for account in user_input.splitlines() if account.strip()]
        root.destroy()

    submit_button = tk.Button(root, text="提交", command=on_submit)
    submit_button.pack()

    root.mainloop()
    return accounts


# 初始化结果存储
all_results = []

def scrape_site(site, accounts):
    results = []

    def on_continue():
    # 關閉登入提示窗口
        root.destroy()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        page.goto(site["url"])
        
        # 創建一個新的窗口來顯示提示
        root = tk.Tk()
        root.title(f"{site['name']} - 進行中")

        label = tk.Label(root, text=f"請在 '{site['name']}' 至並選好日期後，點擊開始查詢。")
        label.pack(pady=10)

        continue_button = tk.Button(root, text="開始查詢", command=on_continue)
        continue_button.pack(pady=10)

        root.mainloop()  # 等待用戶點擊按鈕

        for query in accounts:
            try:
                # 处理 iframe 站点
                if site["name"] in ["TC","TF"]:  # 将所有需要处理 iframe 的站点放入列表
                    iframe = page.locator("iframe[name=\"mainFrame\"]").content_frame
                    if not iframe:
                        print(f"未能找到指定的 iframe '{site['login_selector']}'，请检查选择器。")
                        continue
                    
                    iframe.locator("#LoginID").wait_for()  # 等待输入框可用
                    iframe.locator("#LoginID").fill("")  # 清空输入框
                    iframe.locator("#LoginID").fill(query)  # 填入帐号
                    iframe.get_by_role("button", name="查询").click()  # 点击查询按钮

                    iframe.locator('#TeamProfitTable').wait_for(timeout=50000)  # 等待表格加载
                    time.sleep(1)

                    # 使用 page.evaluate() 提取整个表格的数据
                    table_data = page.evaluate("""
                        () => {
                            const iframe = document.querySelector("iframe[name='mainFrame']");
                            const iframeDocument = iframe.contentDocument || iframe.contentWindow.document;
                            const rows = Array.from(iframeDocument.querySelectorAll('#TeamProfitTable tbody tr'));
                            return rows.map(row => {
                                const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                                return cells;
                            });
                        }
                    """)

                else:  # 处理其他站点
                    page.fill('#LoginId', query)  # 填入当前帐号
                    page.click('button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]')  # 点击查询按钮
                    time.sleep(1)  # 等待1秒
                    page.wait_for_selector('tbody[data-bind="with: TeamStatistics"]', timeout=50000)  # 等待表格加载

                    # 提取表格数据
                    table_data = page.evaluate("""
                        () => {
                            const rows = Array.from(document.querySelectorAll('tbody[data-bind="with: TeamStatistics"] tr'));
                            return rows.map(row => {
                                const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                                return cells;
                            });
                        }
                    """)

                if table_data:
                    results.extend(table_data)
                else:
                    print("没有抓取到数据，请检查选择器或数据加载状态。")
            
            except Exception as e:
                print(f"处理帐号 '{query}' 时出错: {e}")

        # 将结果添加到所有结果中
        all_results.extend(results)

    return results

selected_sites = get_selected_sites()  # 获取用户选择的站点
if not selected_sites:
    print("没有选择任何站点，程序结束。")
else:
    for site in selected_sites:
        accounts = get_accounts_from_user(site["name"])
        if not accounts:
            print(f"没有输入帐号，跳过 '{site['name']}'。")
            continue

        results = scrape_site(site, accounts)

        file_name = f'查询结果_{site["name"]}.xlsx'
        
        if not os.path.exists(file_name):
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                df = pd.DataFrame(columns=site["columns"])
                df.to_excel(writer, index=False)

        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df = pd.DataFrame(results, columns=site["columns"])
            sheet_name = f'查询{len(all_results) // len(accounts) + 1}'
            df.to_excel(writer, index=False, header=True, sheet_name=sheet_name)

        print(f"查询结果已保存至 '{file_name}' 的工作表 '{sheet_name}'。")

print("查询已结束，结果已保存。")