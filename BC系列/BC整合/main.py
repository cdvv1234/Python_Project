import tkinter as tk
from tkinter import messagebox
import program1
import program2
import program3
import program4
import program5
import program6
from playwright.sync_api import sync_playwright
from tkinter import font as tkFont
import time

# 所有站台列表
sites = [
    {"name": "TC", "url": ""}
]

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("程式選擇器")
        self.root.geometry("600x400")
        self.browser = None
        self.pages = []
        self.selected_sites = []
        self.playwright = None
        self.open_site_button = None
        self.close_site_button = None
        self.btn1 = None
        self.btn2 = None
        self.btn3 = None
        self.btn4 = None
        self.btn5 = None
        self.btn6 = None
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.create_main_interface()

    def create_main_interface(self):
        title_label = tk.Label(self.root, text="請選擇要執行的程式", font=("Arial", 16))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        function_frame = tk.LabelFrame(self.root, text="功能列", font=("Arial", 12))
        function_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        self.btn1 = tk.Button(function_frame, text="事件紀錄抓取", command=lambda: program1.run_program_1(self.selected_sites, self.pages), width=20, height=2, state="disabled")
        self.btn1.pack(pady=5)

        self.btn2 = tk.Button(function_frame, text="投注與盈亏抓取", command=lambda: program2.run_program_2(self.selected_sites, self.pages), width=20, height=2, state="disabled")
        self.btn2.pack(pady=5)

        self.btn3 = tk.Button(function_frame, text="幸運抽獎抓取", command=lambda: program3.run_program_3(self.selected_sites, self.pages), width=20, height=2, state="disabled")
        self.btn3.pack(pady=5)

        self.btn4 = tk.Button(function_frame, text="審單DATA抓取", command=lambda: program4.run_program(self.root, self.selected_sites, self.pages), width=20, height=2, state="disabled")
        self.btn4.pack(pady=5)

        self.btn5 = tk.Button(function_frame, text="直屬及下級盈虧查詢", command=lambda: program5.run_program_5(self.root, self.selected_sites, self.pages), width=20, height=2, state="disabled")
        self.btn5.pack(pady=5)

        self.btn6 = tk.Button(function_frame, text="招商觀察", command=lambda: program6.run_program_6(self.root, self.selected_sites, self.pages), width=20, height=2, state="disabled")
        self.btn6.pack(pady=5)

        site_frame = tk.Frame(self.root)
        site_frame.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")

        open_frame = tk.LabelFrame(site_frame, text="開啟站台", font=("Arial", 12))
        open_frame.pack(pady=5)
        self.open_site_button = tk.Button(open_frame, text="開啟站台", command=self.select_sites, width=15, height=2)
        self.open_site_button.pack()

        close_frame = tk.LabelFrame(site_frame, text="關閉站台", font=("Arial", 12))
        close_frame.pack(pady=5)
        self.close_site_button = tk.Button(close_frame, text="關閉站台", command=self.close_browsers, width=15, height=2, state="disabled")
        self.close_site_button.pack()

        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=3)
        self.root.grid_columnconfigure(1, weight=1)

        print("主程式啟動，主視窗已創建")

    def select_sites(self):
        window = tk.Toplevel(self.root)
        window.title("選擇站台")
        window.geometry("600x400")
        custom_font = tkFont.Font(family="Arial", size=12)

        # 勾選站台部分，調整為新的佈局
        checkbox_vars = []
        site_layout = [
            [("TC", 0, 0), ("SY", 0, 1), ("XC", 0, 2), ("CY", 0, 3)],
            [("TF", 1, 0), ("FL", 1, 1), ("XH", 1, 2)],
            [("TS", 2, 0), ("WX", 2, 1), ("CJ", 2, 2)]
        ]
        for row_data in site_layout:
            for site_name, row, col in row_data:
                var = tk.BooleanVar(value=False)
                chk = tk.Checkbutton(window, text=site_name, variable=var, font=custom_font)
                chk.grid(row=row, column=col, padx=20, pady=10, sticky="nsew")
                checkbox_vars.append((site_name, var))

        # 設置列權重以均勻分配空間
        for i in range(4):
            window.grid_columnconfigure(i, weight=1)
        for i in range(3):
            window.grid_rowconfigure(i, weight=1)

        # 帳號和密碼輸入框，進一步縮小間距，並添加預設值
        account_var = tk.StringVar(value="")  # 預設帳號
        account_label = tk.Label(window, text="帳號：", font=custom_font)
        account_label.grid(row=3, column=0, padx=2, pady=5, sticky="e")
        account_entry = tk.Entry(window, width=30, font=custom_font, textvariable=account_var)
        account_entry.grid(row=3, column=1, columnspan=3, padx=2, pady=5)

        password_var = tk.StringVar(value="")  # 預設密碼
        password_label = tk.Label(window, text="密碼：", font=custom_font)
        password_label.grid(row=4, column=0, padx=2, pady=5, sticky="e")
        password_entry = tk.Entry(window, width=30, font=custom_font, show="*", textvariable=password_var)
        password_entry.grid(row=4, column=1, columnspan=3, padx=2, pady=5)

        def confirm_selection():
            # 獲取帳號和密碼的值
            username = account_var.get()
            password = password_var.get()
            
            # 獲取選擇的站台
            self.selected_sites = [site for site in sites if next((var.get() for name, var in checkbox_vars if name == site["name"]), False)]
            
            # 銷毀視窗
            window.destroy()
            
            # 檢查是否選擇了站台
            if not self.selected_sites:
                messagebox.showinfo("提示", "未選擇任何站台。")
            else:
                self.open_browsers_with_login(username, password)

        def cancel_selection():
            window.destroy()

        # 確定和取消按鈕置中
        button_frame = tk.Frame(window)
        button_frame.grid(row=5, column=0, columnspan=4, pady=20)

        submit_button = tk.Button(button_frame, text="確定", command=confirm_selection, font=custom_font, width=10)
        submit_button.pack(side="left", padx=10)

        cancel_button = tk.Button(button_frame, text="取消", command=cancel_selection, font=custom_font, width=10)
        cancel_button.pack(side="left", padx=10)

        # 調整權重以向左移動
        window.grid_columnconfigure(0, weight=2)
        window.grid_columnconfigure(1, weight=1)
        window.grid_columnconfigure(2, weight=1)
        window.grid_columnconfigure(3, weight=1)

        window.wait_window()

    def open_browsers_with_login(self, username, password):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=False)
        self.pages = []
        for site in self.selected_sites:
            page = self.browser.new_page()
            page.goto(site["url"])
            # 根據站台類型選擇不同的選擇器進行自動登入
            try:
                # 帳號和密碼的選擇器對於所有站台相同
                page.wait_for_selector("#LoginID", state="visible", timeout=5000)
                page.fill("#LoginID", username)
                page.wait_for_selector("#Password", state="visible", timeout=5000)
                page.fill("#Password", password)

                # 根據站台類型選擇不同的登入按鈕
                if site["name"] in ["TC", "TF"]:
                    login_button_selector = "body > div > div > div > div > div.panel-body > form > fieldset > input"
                else:
                    login_button_selector = "body > div > div > div > div > div > div > div > div > form > input.btn.btn-primary.btn-user.btn-block"

                page.wait_for_selector(login_button_selector, state="visible", timeout=5000)
                page.click(login_button_selector)
                time.sleep(2)  # 等待登入完成
            except Exception as e:
                print(f"站台 {site['name']} 登入失敗: {e}")
            self.pages.append(page)

        self.btn1.config(state="normal")
        self.btn2.config(state="normal")
        self.btn3.config(state="normal")
        self.btn4.config(state="normal")
        self.btn5.config(state="normal")
        self.btn6.config(state="normal")
        self.close_site_button.config(state="normal")
        self.open_site_button.config(state="disabled")

        tk.Tk().withdraw()
        messagebox.showinfo("準備分頁", "請在所有分頁完成登入，然後點擊確定繼續。")

    def close_browsers(self):
        if self.browser:
            self.browser.close()
            self.browser = None
        if self.playwright:
            self.playwright.stop()
            self.playwright = None
        self.pages = []
        self.selected_sites = []

        self.btn1.config(state="disabled")
        self.btn2.config(state="disabled")
        self.btn3.config(state="disabled")
        self.btn4.config(state="disabled")
        self.btn5.config(state="disabled")
        self.btn6.config(state="disabled")
        self.close_site_button.config(state="disabled")
        self.open_site_button.config(state="normal")

        print("瀏覽器已關閉")

    def on_closing(self):
        print("關閉主視窗，清理資源...")
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        self.root.quit()
        self.root.destroy()
        print("主程式結束，主視窗已銷毀")

if __name__ == "__main__":
    try:    
        root = tk.Tk()
        app = MainApp(root)
        root.mainloop()
    except Exception as e:
        print(f"程式發生錯誤: {e}")
        raise
