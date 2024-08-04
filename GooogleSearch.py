from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Edge()

# 打開 google 網頁
driver.get("https://www.google.com")

wait = WebDriverWait(driver, 5)  #設置等待時間為5秒

element = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "gLFyf")))
element.send_keys("Python")#輸入要搜尋的項目

button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "gNO89b")))
button.click()#點擊搜尋

while True:
    pass  # 添加無限循環，保持程序運行狀態
