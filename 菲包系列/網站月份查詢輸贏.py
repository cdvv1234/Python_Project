from playwright.sync_api import sync_playwright
import pandas as pd

def extract_flat_headers(header_rows):
    """Extract and flatten multi-level headers."""
    flat_headers = []
    for header_row in header_rows:
        row_headers = []
        for th in header_row.query_selector_all('th'):
            colspan = int(th.get_attribute('colspan') or 1)
            text = th.inner_text().strip()
            row_headers.extend([text] * colspan)
        flat_headers.append(row_headers)
    
    # Flatten the headers by combining all rows' headers into a single list
    # assuming headers are to be combined into one row.
    if flat_headers:
        return flat_headers[0] + flat_headers[1]  # Adjust if you need to combine differently
    return []

def run(playwright):
    browser = playwright.firefox.launch(headless=False)
    page = browser.new_page()

    try:
        page.goto('##')
        page.wait_for_load_state('networkidle')

        page.fill('input#input-22', '##')  # 填寫用戶名
        page.fill('input#input-26', '##')  # 填寫密碼
        page.click('.v-select__selections')
        page.wait_for_selector('.v-menu__content.menuable__content__active', state='visible')
        page.click('text=简体中文')
        page.wait_for_selector('.v-select__selection >> text=简体中文')
        page.click('button:has-text("登入")')
        page.wait_for_load_state('networkidle')

        page.click('button.v-app-bar__nav-icon')
        page.click('div.v-list-group__header:has-text("报表")')
        page.click('a:has-text("##")')

        # 開始日期的選擇
        page.get_by_label("开始日期时间").click()
        page.locator('div.v-date-picker-header').locator('button[aria-label="前一个月"]').first.click()
        page.get_by_role("button", name="31").click()
        page.get_by_role("button", name="保存").click()

        # 等待日期選擇器關閉
        page.wait_for_function(
            "document.querySelector('[aria-expanded=\"false\"]').getAttribute('aria-expanded') === 'false'"
        )

        # 結束日期的選擇
        page.get_by_label("结束日期时间").click()
        page.locator('div.v-date-picker-header').locator('button[aria-label="前一个月"]').nth(1).click()
        page.get_by_role("button", name="31").click()
        page.get_by_role("button", name="保存").click()

        # 等待日期選擇器關閉
        page.wait_for_function(
            "document.querySelector('[aria-expanded=\"false\"]').getAttribute('aria-expanded') === 'false'"
        )

        # 點按搜尋按鈕
        page.click('button.addBtn')
        page.wait_for_selector('tr[data-v-32aa03bf]')

        # 抓取表格數據
        thead_element = page.query_selector('thead')
        rows = page.query_selector_all('tr[data-v-32aa03bf]')

        if thead_element and rows:
            # 解析標頭
            header_rows = thead_element.query_selector_all('tr')
            flat_headers = extract_flat_headers(header_rows)

            # 解析行數據
            rows_data = []
            for row in rows:
                cells = [cell.inner_text().strip() for cell in row.query_selector_all('td, th')]
                rows_data.append(cells)

            # 創建 DataFrame，確保列數匹配
            max_columns = max(len(row) for row in rows_data)
            if len(flat_headers) < max_columns:
                flat_headers.extend([''] * (max_columns - len(flat_headers)))
            df = pd.DataFrame(rows_data, columns=flat_headers[:max_columns])

            # 將 DataFrame 輸出到 Excel
            df.to_excel('report.xlsx', index=False)

            print("數據已成功導出到 'report.xlsx'")
        else:
            print("未找到表格標頭或行")

    except Exception as e:
        print(f"發生錯誤: {e}")

    finally:
        # 保持瀏覽器打開以便觀察
        # browser.close()
        pass

with sync_playwright() as playwright:
    run(playwright)
