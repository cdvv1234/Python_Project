# PTT SportLottery 追蹤器

## 專案簡介
PTT SportLottery 追蹤器是一個 Python 腳本，用於監控台灣 PTT 論壇的 SportLottery 看板。它會自動檢查新發表的文章，篩選特定作者或推文數達到指定門檻的文章（排除 LIVE、活動和公告文章），並透過 LINE Messaging API 發送通知。程式支援系統托盤功能，允許用戶暫停、繼續或結束程式運行。

## 功能
- **文章監控**：掃描 SportLottery 看板的最新文章，檢查是否符合條件（特定作者或推文數 ≥ 50）。
- **條件篩選**：追蹤指定作者（預設：`apparition10`, `lotterywin`, `bvbin10242`）或推文數大於等於 50 的文章，排除標題包含 `[LIVE]`、`[活動]` 或 `[公告]` 的文章。
- **LINE 通知**：當發現符合條件的文章時，透過 LINE 發送包含標題、作者、推文數和連結的通知。
- **系統托盤**：提供系統托盤圖標，支援暫停、繼續和結束程式的操作。
- **資料持久化**：將已追蹤的文章儲存至本地的 JSON 檔案，避免重複通知。
- **靈活檢查間隔**：每 30 分鐘檢查一次新文章。

## 依賴項
- Python 3.6 或以上
- 必要 Python 套件：
  ```bash
  requests
  beautifulsoup4
  pystray
  Pillow
  ```
- 可選：自訂系統托盤圖標（`icon.ico`）

## 安裝
1. **克隆或下載程式碼**：
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. **安裝依賴項**：
   使用 pip 安裝必要的 Python 套件：
   ```bash
   pip install requests beautifulsoup4 pystray Pillow
   ```

3. **設置 LINE Messaging API**：
   - 註冊 LINE 官方帳號並獲取以下資訊：
     - `LINE_CHANNEL_ID`
     - `LINE_CHANNEL_SECRET`
     - `LINE_CHANNEL_ACCESS_TOKEN`
     - `LINE_USER_ID`
   - 在程式碼中更新以下變數：
     ```python
     LINE_CHANNEL_ID = "your_channel_id"
     LINE_CHANNEL_SECRET = "your_channel_secret"
     LINE_CHANNEL_ACCESS_TOKEN = "your_access_token"
     LINE_USER_ID = "your_user_id"
     ```

4. **（可選）自訂系統托盤圖標**：
   - 將自訂的 `icon.ico` 檔案放置於程式碼所在的目錄。
   - 若無自訂圖標，程式將使用預設的藍色方塊圖標。

## 使用方法
1. **配置追蹤條件**：
   - 修改程式碼中的 `TARGET_AUTHORS` 來指定要追蹤的作者：
     ```python
     TARGET_AUTHORS = ["apparition10", "lotterywin", "bvbin10242"]
     ```
   - 調整 `MIN_COMMENTS` 來設置最低推文數門檻（預設為 50）：
     ```python
     MIN_COMMENTS = 50
     ```
   - 設定檢查間隔（單位：秒，預設為 120 秒）：
     ```python
     CHECK_INTERVAL = 120
     ```

2. **運行程式**：
   ```bash
   python ptt_tracker.py
   ```
   - 程式啟動後將在背景運行，檢查 SportLottery 看板的新文章。
   - 符合條件的文章將觸發 LINE 通知，並將資訊儲存至 `tracked_posts.json`。

3. **系統托盤操作**：
   - 程式運行時，系統托盤中會顯示一個圖標（預設為藍色方塊或自訂圖標）。
   - 右鍵點擊圖標可選擇：
     - **繼續執行**：恢復文章監控。
     - **暫停**：暫停文章監控。
     - **結束程式**：完全退出程式。

4. **停止程式**：
   - 使用系統托盤的「結束程式」選項，或按 `Ctrl+C` 終止終端機中的執行。

## 檔案結構
```
ptt_tracker/
├── ptt_tracker.py        # 主程式碼
├── tracked_posts.json    # 儲存已追蹤文章的資料檔案
├── icon.ico              # （可選）自訂系統托盤圖標
└── README.md             # 本說明文件
```

## 注意事項
- **LINE API 配置**：確保正確設置 LINE Messaging API 的認證資訊，否則通知功能將不可用。
- **錯誤處理**：程式包含基本的錯誤處理，若發生問題（如網路錯誤），會在控制台輸出錯誤訊息，並嘗試透過 LINE 通知錯誤。
- **資料儲存**：已追蹤的文章儲存於 `tracked_posts.json`，確保程式結束後不會重複通知。
- **執行緒安全**：程式使用多執行緒運行主迴圈和系統托盤，確保穩定運行。

## 故障排除
- **LINE 通知未收到**：
  - 檢查 `LINE_CHANNEL_ACCESS_TOKEN` 和 `LINE_USER_ID` 是否正確。
  - 確認網路連線是否正常。
- **程式無回應**：
  - 檢查控制台輸出是否有錯誤訊息。
  - 確保依賴項已正確安裝。
- **圖標未顯示**：
  - 確認 `icon.ico` 是否存在於程式目錄，或移除自訂圖標以使用預設圖標。

---

# PTT SportLottery Tracker

## Overview
The PTT SportLottery Tracker is a Python script designed to monitor the SportLottery board on Taiwan's PTT forum. It automatically checks for new posts by specific authors or with a minimum number of comments (excluding LIVE, event, and announcement posts) and sends notifications via the LINE Messaging API. The script includes a system tray icon for pausing, resuming, or exiting the program.

## Features
- **Post Monitoring**: Scans the SportLottery board for new posts that meet criteria (specific authors or ≥50 comments).
- **Criteria Filtering**: Tracks posts by specified authors (default: `apparition10`, `lotterywin`, `bvbin10242`) or with at least 50 comments, excluding titles with `[LIVE]`, `[活動]`, or `[公告]`.
- **LINE Notifications**: Sends notifications via LINE with post title, author, comment count, and link.
- **System Tray**: Provides a system tray icon for pausing, resuming, or exiting the program.
- **Data Persistence**: Stores tracked posts in a local JSON file to avoid duplicate notifications.
- **Flexible Interval**: Checks for new posts every 30 minutes.

## Dependencies
- Python 3.6 or higher
- Required Python packages:
  ```bash
  requests
  beautifulsoup4
  pystray
  Pillow
  ```
- Optional: Custom system tray icon (`icon.ico`)

## Installation
1. **Clone or Download the Code**:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. **Install Dependencies**:
   Install required Python packages using pip:
   ```bash
   pip install requests beautifulsoup4 pystray Pillow
   ```

3. **Set Up LINE Messaging API**:
   - Register a LINE Official Account and obtain:
     - `LINE_CHANNEL_ID`
     - `LINE_CHANNEL_SECRET`
     - `LINE_CHANNEL_ACCESS_TOKEN`
     - `LINE_USER_ID`
   - Update the following variables in the code:
     ```python
     LINE_CHANNEL_ID = "your_channel_id"
     LINE_CHANNEL_SECRET = "your_channel_secret"
     LINE_CHANNEL_ACCESS_TOKEN = "your_access_token"
     LINE_USER_ID = "your_user_id"
     ```

4. **(Optional) Custom System Tray Icon**:
   - Place a custom `icon.ico` file in the same directory as the script.
   - If no custom icon is provided, a default blue square icon is used.

## Usage
1. **Configure Tracking Criteria**:
   - Modify `TARGET_AUTHORS` to specify authors to track:
     ```python
     TARGET_AUTHORS = ["apparition10", "lotterywin", "bvbin10242"]
     ```
   - Adjust `MIN_COMMENTS` to set the minimum comment threshold (default: 50):
     ```python
     MIN_COMMENTS = 50
     ```
   - Set the check interval (in seconds, default: 120):
     ```python
     CHECK_INTERVAL = 120
     ```

2. **Run the Script**:
   ```bash
   python ptt_tracker.py
   ```
   - The script runs in the background, monitoring the SportLottery board.
   - Matching posts trigger LINE notifications and are saved to `tracked_posts.json`.

3. **System Tray Operations**:
   - A system tray icon (default: blue square or custom icon) appears when the script runs.
   - Right-click the icon to:
     - **Resume**: Resume post monitoring.
     - **Pause**: Pause post monitoring.
     - **Exit**: Terminate the program.

4. **Stop the Script**:
   - Use the system tray’s “Exit” option or press `Ctrl+C` in the terminal.

## File Structure
```
ptt_tracker/
├── ptt_tracker.py        # Main script
├── tracked_posts.json    # Stores tracked post data
├── icon.ico              # (Optional) Custom system tray icon
└── README.md             # This documentation
```

## Notes
- **LINE API Setup**: Ensure LINE Messaging API credentials are correctly configured, or notifications will not work.
- **Error Handling**: Includes basic error handling; errors (e.g., network issues) are logged to the console and notified via LINE if possible.
- **Data Storage**: Tracked posts are saved in `tracked_posts.json` to prevent duplicate notifications.
- **Thread Safety**: Uses threading for the main loop and system tray, ensuring stable operation.
