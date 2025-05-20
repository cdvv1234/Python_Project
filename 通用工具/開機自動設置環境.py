import pyautogui
import time
import pyperclip
import os
import tkinter as tk
from tkinter import messagebox
import json

# Step 1: 創建資料夾
def create_folders(selected_groups, folder_groups):
    base_path = r"C:\Users\Administrator\Downloads"  # 基本路徑為 Downloads

    # 創建 done 和 test 資料夾（位於 Downloads 下）
    for folder_name in ["done", "test"]:
        folder_path = os.path.join(base_path, folder_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"資料夾已創建: {folder_path}")
        else:
            print(f"資料夾已存在: {folder_path}")
    
    # 在 test 資料夾下創建選擇的資料夾
    test_folder_path = os.path.join(base_path, "test")
    for group in selected_groups:
        folder_names = folder_groups[group]
        for folder_name in folder_names:
            folder_path = os.path.join(test_folder_path, folder_name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                print(f"資料夾已創建: {folder_path}")
            else:
                print(f"資料夾已存在: {folder_path}")
    
    messagebox.showinfo("完成", "資料夾創建完成！")

# 載入或創建 folder_groups
def load_folder_groups():
    json_file = "folder_groups.json"
    default_groups = {
        "1. - 1、2、4、5": ["01", "02", "04", "05"],
        "2. - 6、8、9、12": ["06", "08", "09", "12"],
        "3. - 14、17、19、22": ["14", "17", "19", "22"],
        "4. - 27、28、31、32": ["27", "28", "31", "32"],
        "5. - 33、34、35、36": ["33", "34", "35", "36"],
        "6.- 37、38、ICT02、TG1、TG2": ["37", "38", "ICT02", "TG1", "TG2"]
    }
    
    # 檢查 JSON 檔案是否存在
    if not os.path.exists(json_file):
        # 若不存在，創建預設 JSON 檔案
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(default_groups, f, ensure_ascii=False, indent=4)
        print(f"已創建預設 {json_file}")
        return default_groups
    
    # 載入 JSON 檔案
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            folder_groups = json.load(f)
        print(f"已載入 {json_file}")
        return folder_groups
    except json.JSONDecodeError:
        print(f"無法解析 {json_file}，使用預設值")
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(default_groups, f, ensure_ascii=False, indent=4)
        return default_groups

# Step 2: 開啟「網路和網際網路」設定
def open_vpn_settings():
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.write('ms-settings:network-vpn')
    pyautogui.press('enter')

# Step 3: 模擬點擊新增 VPN
def click_add_vpn():
    time.sleep(2)  # 等待設定頁面加載
    pyautogui.click(1572,264)  # 這裡需要確定新增按鈕的位置

# Step 4: 輸入 VPN 資訊
def fill_vpn_info(vpn_name, server_address, acc, psk):
    time.sleep(1)
    
    pyautogui.press('tab')
    pyperclip.copy(vpn_name)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyperclip.copy(server_address)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyperclip.copy(acc)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyperclip.copy(psk)
    pyautogui.hotkey('ctrl', 'v')

# Step 5: 儲存 VPN 設定
def save_vpn():
    pyautogui.click(855,931)  # 點擊儲存按鈕

# Step 6: 回到新增 VPN 頁面開始的步驟
def reset_to_start():
    pyautogui.press('esc')
    time.sleep(1)

# Step 7: 新增多組 VPN 連線
def add_vpns():
    vpn_list = [
        {"name": "", "server": "", "account": "", "password": ""},
        {"name": "", "server": "", "account": "", "password": ""},
        {"name": "", "server": "", "account": "", "password": ""}
    ]
    
    open_vpn_settings()

    for vpn in vpn_list:
        click_add_vpn()
        fill_vpn_info(vpn["name"], vpn["server"], vpn["account"], vpn["password"])
        save_vpn()
        reset_to_start()
        time.sleep(1)

    messagebox.showinfo("完成", "VPN 已成功新增！")

# UI 佈局
def create_ui():
    folder_groups = load_folder_groups()
    window = tk.Tk()
    window.title("選擇要創建的資料夾和新增VPN")
    window.geometry("400x350")

    selected_groups = []

    def toggle_selection(group_name):
        if group_name in selected_groups:
            selected_groups.remove(group_name)
        else:
            selected_groups.append(group_name)

    # 顯示選項
    for group_name in folder_groups:
        checkbox = tk.Checkbutton(window, text=group_name, command=lambda name=group_name: toggle_selection(name))
        checkbox.pack(anchor="w")

    # 創建資料夾按鈕
    def on_create_folders():
        if selected_groups:
            create_folders(selected_groups, folder_groups)
        else:
            messagebox.showwarning("警告", "請選擇至少一個資料夾！")

    btn_create_folders = tk.Button(window, text="創建資料夾", command=on_create_folders, width=20)
    btn_create_folders.pack(pady=10)

    # 新增VPN按鈕
    btn_add_vpns = tk.Button(window, text="新增 VPN", command=add_vpns, width=20)
    btn_add_vpns.pack(pady=10)

    window.mainloop()

# 主程序
def main():
    create_ui()

if __name__ == "__main__":
    main()
