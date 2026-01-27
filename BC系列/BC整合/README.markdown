# 網站資料抓取工具

## 概述

本專案是一個基於 Python 的網頁爬蟲工具，整合之前所做的程式，包括事件紀錄、投注與盈虧記錄、幸運抽獎統計和審單資料的抓取功能。使用 Playwright 進行瀏覽器自動化操作，支持從多個網站提取資料並儲存為 Excel 檔案。該工具提供 Tkinter 圖形界面，方便用戶選擇網站、輸入帳號密碼和指定日期範圍，並包含文字清理功能，移除多餘換行符和多個空格。

## 工具

- **Playwright**：用於網頁瀏覽器自動化。
- **Pandas** 和 **Openpyxl**：處理和儲存 Excel 檔案。
- **Tkcalendar**：提供日期選擇界面。
- **Tkinter**：構建圖形用戶界面。

## 安裝

### 前置條件

- Python 3.8 或更高版本
- 與 Playwright 兼容的網頁瀏覽器（Chromium、Firefox 或 WebKit）

### 依賴項

使用 pip 安裝所需套件：

```bash
pip install pandas openpyxl tkcalendar playwright
```

安裝 Playwright 瀏覽器：

```bash
playwright install
```

### 設置

1. 將專案檔案下載或複製到本地。
2. 確保所有依賴項已安裝。
3. 將 `main.py`、`program1.py`、`program2.py`、`program3.py` 和 `program4.py` 放置在同一目錄。

## 圖形界面操作

1. **選擇網站**：從列表中選擇網站（例如 TC、TF、CJ）。
2. **輸入帳號密碼**：輸入網站訪問所需的用戶名和密碼。
3. **選擇程式**：選擇要運行的程式：
   - **程式 1**：事件紀錄
   - **程式 2**：投注與盈虧資料
   - **程式 3**：幸運抽獎統計
   - **程式 4**：審單資料
   - **程式 5**：個人盈虧(含直屬及其下級帳號)
   - **程式 6**：個人盈虧資料(含彩票與電子)
   - **程式 7**：帳戶管理抓取
   - **程式 8**：投注紀錄全抓取
   - **程式 9**：彩種玩法統計
4. **指定日期**：輸入資料抓取的開始和結束日期（部分程式需包含時間）。

## 輸出

- 結果儲存為 Excel 檔案，文字欄位已清理（無多餘換行或多個空格）。
- 例如：欄位「允许出款\n\n第一次出款」將儲存為「允许出款 第一次出款」。

---

# Web Scraper for Site Data Extraction

## Overview

This project is a Python-based web scraping tool that consolidates functionalities for extracting event logs, betting and profit/loss records, lucky draw statistics, and withdrawal audit data. It uses the Playwright library for browser automation, supports data extraction from multiple sites, and saves results as Excel files. The tool includes a Tkinter GUI for selecting sites, entering credentials, and specifying date ranges, with text cleaning to remove excessive newlines and spaces.

## Tools

- **Playwright**: For browser automation.
- **Pandas** and **Openpyxl**: For Excel file handling.
- **Tkcalendar**: For date selection interface.
- **Tkinter**: For building the GUI.

## Installation

### Prerequisites

- Python 3.8 or higher
- A compatible web browser (Chromium, Firefox, or WebKit) for Playwright

### Dependencies

Install required packages using pip:

```bash
pip install pandas openpyxl tkcalendar playwright
```

Install Playwright browsers:

```bash
playwright install
```

### Setup

1. Download or clone the project files.
2. Ensure all dependencies are installed.
3. Place `main.py`, `program1.py`, `program2.py`, `program3.py`, and `program4.py` in the same directory.

