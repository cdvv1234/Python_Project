from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# 設定 geckodriver 路徑
service = Service(executable_path='C:/TEST/geckodriver.exe')

# 初始化 Firefox 驅動
driver = webdriver.Firefox(service=service)

try:
    # 打開網站
    driver.get("#")

    # 定位帳號輸入框並輸入帳號
    username_input = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "input-22"))
    )
    username_input.send_keys("#")

    # 定位密碼輸入框並輸入密碼
    password_input = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "input-26"))
    )
    password_input.send_keys("#")

    # 等待並點擊語系選擇框
    language_select = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.v-select'))
    )
    language_select.click()

    # 選擇「簡體中文」
    option = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@role='listbox']//div[contains(@class, 'v-list-item') and .//div[contains(@class, 'v-list-item__title') and text()='简体中文']]"))
    )
    option.click()

    # 定位登入按鈕並點擊
    login_button = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.v-btn.v-btn--text.theme--light.v-size--x-large.dialogTitleBg--text'))
    )
    login_button.click()

    # 等待幾秒鐘以便觀察結果，或執行其他操作
    time.sleep(5)

    # 登入後導航到指定頁面
    driver.get('https://sy.dggo.net/#/WinLostByProduct')

    # 確保頁面加載完成
    WebDriverWait(driver, 30).until(
        EC.title_contains("WinLostByProduct")
    )

    # 等待搜尋按鈕可見
    search_button = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'button.addBtn'))
    )
    # 滾動到元素
    driver.execute_script("arguments[0].scrollIntoView();", search_button)
    # 等待按鈕可點擊
    search_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.addBtn'))
    )
    # 點擊搜尋按鈕
    search_button.click()

    # 獲取總計值
    total_row = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//tr[contains(., '总计 :')]")))
    total_values = total_row.find_elements(By.TAG_NAME, 'td')

    for value in total_values:
        print(value.text)

except Exception as e:
    print(f"發生錯誤: {e}")

finally:
    # 保持瀏覽器窗口開啟
    pass
