from playwright.sync_api import sync_playwright
import pandas as pd
import time
import os

# 帐号列表
queries = [
  
]

# 初始化结果存储
all_results = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    page.goto('')  # 替换为你的页面网址
    print("请手动登入，然后按 Enter 继续...")
    input()

    while True:
        try:
            results = []

            for query in queries:
                # 清空输入框并填入当前帐号
                page.fill('#LoginId', query)  # 填入当前值
                page.click('button[data-bind="click: SearchClick, attr: { disabled: SearchLock }"]')

                # 等待1.5秒
                time.sleep(1)

                # 等待统计数据显示
                page.wait_for_selector('tbody[data-bind="with: TeamStatistics"]', timeout=50000)

                # 使用 page.evaluate() 提取整个表格的数据
                table_data = page.evaluate("""
                    () => {
                        const rows = Array.from(document.querySelectorAll('tbody[data-bind="with: TeamStatistics"] tr'));
                        return rows.map(row => {
                            const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerText);
                            // 将总盈亏乘以 -1
                            cells[4] = (-1 * parseFloat(cells[4])).toFixed(3);  // 总盈亏在索引4
                            return cells;
                        });
                    }
                """)

                # 将结果添加到列表中
                results.extend(table_data)

            # 将结果添加到所有结果中
            all_results.extend(results)

            # 检查 Excel 文件是否存在
            file_name = '查询结果.xlsx'
            if not os.path.exists(file_name):
                # 如果文件不存在，创建新的文件并写入标题
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    df = pd.DataFrame(columns=['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利'])
                    df.to_excel(writer, index=False)

            # 追加写入数据
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                # 将当前查询结果写入新的工作表
                df = pd.DataFrame(results, columns=['总投注', '总奖金', '总返点', '总活动', '总盈亏', '总充值', '总提款', '总红利'])
                sheet_name = f'查询{len(all_results) // len(queries) + 1}'
                df.to_excel(writer, index=False, header=True, sheet_name=sheet_name)

            print(f"查询结果已保存至 '{file_name}' 的工作表 '{sheet_name}'。")

            # 询问用户是否继续查询新的日期
            continue_query = input("是否继续查询新的日期？（y/n）: ")
            if continue_query.lower() != 'y':
                break  # 结束循环

        except Exception as e:
            print(f"发生错误: {e}")

    page.close()  # 确保页面关闭
    browser.close()  # 确保浏览器关闭

    print("查询已结束，结果已保存。")
