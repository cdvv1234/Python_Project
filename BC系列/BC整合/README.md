概述：
整併之前程式包括事件紀錄、投注與盈虧記錄、幸運抽獎統計和審單資料抓取。

工具：playwright

安裝：

Python 3.8 或更高版本
與 Playwright 兼容的網頁瀏覽器（Chromium、Firefox 或 WebKit）

使用 pip 安裝所需的 Python 套件：
pip install pandas openpyxl tkcalendar playwright

安裝 Playwright 瀏覽器：
playwright install

圖形界面操作：

選擇網站：從列表中選擇一個或多個網站（例如 TC、TF、CJ）。
輸入帳號密碼：輸入用於網站訪問的用戶名和密碼。
選擇程式：選擇要運行的程式（1、2、3 或 4）：
程式 1：事件紀錄
程式 2：投注與盈虧
程式 3：幸運抽獎統計
程式 4：審單資料抓取

輸出：

結果儲存為 Excel 檔案，文字欄位已清理（無多餘換行或多個空格）。
例如：欄位「允许出款\n\n第一次出款」將儲存為「允许出款 第一次出款」。


錯誤處理：

若發生錯誤（例如導航失敗、找不到元素），程式會記錄詳細資訊並在專案目錄中儲存截圖（例如 error_TC_eventlog_navigation.png）。


檔案結構：
BC整合/
├── main.py               # 主腳本，包含圖形界面和網站登入邏輯
├── program1.py           # 抓取事件紀錄
├── program2.py           # 抓取投注與盈虧資料
├── program3.py           # 抓取幸運抽獎統計
├── program4.py           # 抓取出款審核資料
├── README.md             # 專案說明文件
├── error_*.png           # 錯誤截圖（運行時生成）
└── *.xlsx                # 輸出 Excel 檔案（運行時生成）


Web Scraper for Site Data Extraction
Overview
It uses the Playwright library for browser automation and supports scraping different types of data, including event logs, 
betting and profit/loss records, lucky draw statistics, and withdrawal audit data. 

Installation
Prerequisites

Python 3.8 or higher
A compatible web browser (Chromium, Firefox, or WebKit) for Playwright

Dependencies
Install the required Python packages using pip:
pip install pandas openpyxl tkcalendar playwright

Install Playwright browsers:
playwright install

GUI Interaction:

Select Sites: Choose one or more sites from the list (e.g., TC, TF, CJ).
Enter Credentials: Input your username and password for site access.
Choose Program: Select a program (1, 2, 3, or 4) to run:
Program 1: Event logs
Program 2: Betting and profit/loss
Program 3: Lucky draw statistics
Program 4: Withdrawal audit data

Specify Dates: Enter start and end dates (and times, where applicable) for data extraction.

Output:

Results are saved as Excel files with cleaned text (no excessive newlines or multiple spaces).
Example: A field like 允许出款\n\n第一次出款 will be saved as 允许出款 第一次出款.


Error Handling:

If errors occur (e.g., navigation failures, missing elements), the program logs details and saves screenshots in the project directory (e.g., error_TC_eventlog_navigation.png).


File Structure
BC整合/
├── main.py               # Main script with GUI and site login logic
├── program1.py           # Scrapes event logs
├── program2.py           # Scrapes betting and profit/loss data
├── program3.py           # Scrapes lucky draw statistics
├── program4.py           # Scrapes withdrawal audit data
├── README.md             # Project documentation
├── error_*.png           # Screenshots of errors (generated during runtime)
└── *.xlsx                # Output Excel files (generated during runtime)
