from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By

# 設定 geckodriver 路徑（假設 geckodriver 在系統 PATH 中）
service = Service(executable_path='C:/TEST/geckodriver.exe')

# 初始化 Firefox 驅動
driver = webdriver.Firefox(service=service)

# 打開 Google
driver.get("你的網址")

# 定位帳號輸入框並輸入帳號
username_input = driver.find_element(By.ID, "LoginID")
username_input.send_keys("你的帳號")

# 定位密碼輸入框並輸入密碼
password_input = driver.find_element(By.ID, "Password")
password_input.send_keys("你的密碼")

# 定位登入按鈕並點擊
login_button = driver.find_element(By.CSS_SELECTOR, "input.btn.btn-primary.btn-user.btn-block")
login_button.click()

# 可以選擇等待幾秒鐘以便觀察結果，或執行其他操作
import time
time.sleep(5)