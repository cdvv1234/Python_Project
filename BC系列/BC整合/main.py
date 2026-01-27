import tkinter as tk
from tkinter import messagebox, ttk
import program1, program2, program3, program4, program5, program6, program7, program8, program9, program10
from playwright.async_api import async_playwright
import asyncio
import threading
import time

# 所有站台列表
sites = [
    {"name": "TC", "url": "https://beh.tcssc9.net/"},
    {"name": "TF", "url": "https://be.jf988.net/"},
    {"name": "TS", "url": "https://be-h.tssc99.net/"},
    {"name": "SY", "url": "https://be-h.syss00.com/"},
    {"name": "FL", "url": "https://be.flyl77.net/"},
    {"name": "WX", "url": "https://be-s.wxyl77.com/"},
    {"name": "XC", "url": "https://be-s.xcwin77.com/"},
    {"name": "XH", "url": "https://be.xhxhyl33.com/"},
    {"name": "CJ", "url": "https://be.cjcjyl11.com/"},
    {"name": "CY", "url": "https://be.cywdsd2505.com/"}
]

# 多組預設帳號
ACCOUNT_GROUPS = {
    "data001": {"username": "data001", "password": "Data20251113"},
    "data002": {"username": "data002", "password": "Data20251113"},
}

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("程式選擇器 (全並行防錯版)")
        self.root.geometry("800x500")
        
        # 核心狀態控制
        self.is_running = False 

        # 初始化非同步事件循環
        self.loop = asyncio.new_event_loop()
        self.thread = threading.Thread(target=self._run_async_loop, daemon=True)
        self.thread.start()

        self.playwright = None
        self.browser = None
        self.context = None
        self.pages = []
        self.selected_sites = []
        self.btn_list = []
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.create_main_interface()

    def _run_async_loop(self):
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def run_async(self, coro):
        return asyncio.run_coroutine_threadsafe(coro, self.loop)

    def set_buttons_state(self, state):
        """統一設定功能按鈕與關閉站台按鈕的狀態"""
        for btn in self.btn_list:
            btn.config(state=state)
        self.close_site_button.config(state=state)

    def safe_run_program(self, program_func):
        """防止重複執行的安全包裝器"""
        if self.is_running:
            messagebox.showwarning("警告", "已有功能正在執行中，請等候完成。")
            return
        
        # 鎖定 UI 並標記為執行中
        self.is_running = True
        self.set_buttons_state("disabled")
        
        # 執行子程式，並傳入「完成後的回呼函式 (callback)」
        # 之後每個 program 的入口 function 都要接收這個 callback
        program_func(self.on_program_complete)

    def on_program_complete(self):
        """當子程式結束時，由子程式呼叫此函式解鎖 UI"""
        self.is_running = False
        # 確保在主線程更新 UI
        self.root.after(0, lambda: self.set_buttons_state("normal"))

    def create_main_interface(self):
        title_label = tk.Label(self.root, text="請選擇要執行的程式", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)

        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=30, pady=10)

        # 左側：功能列
        function_frame = tk.LabelFrame(main_frame, text="功能列", font=("Arial", 14), padx=20, pady=20)
        function_frame.pack(side=tk.LEFT, fill="y", padx=(0, 50))

        # 功能按鈕定義
        # 傳入 self.on_program_complete 作為最後一個參數
        buttons = [
            ("事件紀錄抓取", lambda: self.safe_run_program(lambda cb: program1.run_program_1(self.root, self.selected_sites, self.pages, cb))),
            ("投注與盈亏抓取", lambda: self.safe_run_program(lambda cb: program2.run_program_2(self.root, self.selected_sites, self.pages, cb))),
            ("幸運抽獎抓取", lambda: self.safe_run_program(lambda cb: program3.run_program_3(self.root, self.selected_sites, self.pages, cb))),
            ("審單DATA抓取", lambda: self.safe_run_program(lambda cb: program4.run_program_4(self.root, self.selected_sites, self.pages, cb))),
            ("直屬及下級盈虧查詢", lambda: self.safe_run_program(lambda cb: program5.run_program_5(self.root, self.selected_sites, self.pages, cb))),
            ("招商觀察", lambda: self.safe_run_program(lambda cb: program6.run_program_6(self.root, self.selected_sites, self.pages, cb))),
            ("帳戶管理抓取", lambda: self.safe_run_program(lambda cb: program7.run_program_7(self.root, self.selected_sites, self.pages, cb))),
            ("投注紀錄全抓取", lambda: self.safe_run_program(lambda cb: program8.run_program_8(self.root, self.selected_sites, self.pages, cb))),
            ("彩種玩法統計", lambda: self.safe_run_program(lambda cb: program9.run_program_9(self.root, self.selected_sites, self.pages, cb))),
            ("彩種統計", lambda: self.safe_run_program(lambda cb: program10.run_program_10(self.root, self.selected_sites, self.pages, cb))),
        ]

        for i, (text, cmd) in enumerate(buttons):
            btn = tk.Button(function_frame, text=text, command=cmd, width=22, height=2, font=("Arial", 12), state="disabled")
            btn.grid(row=i // 2, column=i % 2, padx=10, pady=10)
            self.btn_list.append(btn)

        # 右側：開啟/關閉站台
        control_frame = tk.Frame(main_frame)
        control_frame.pack(side=tk.RIGHT, fill="y")
        tk.Frame(control_frame, height=60).pack()

        self.open_site_button = tk.Button(control_frame, text="開啟站台", command=self.select_and_open_sites, bg="#4CAF50", fg="white", font=("Arial", 12), width=22, height=2)
        self.open_site_button.pack(pady=10)

        self.close_site_button = tk.Button(control_frame, text="關閉站台", command=lambda: self.run_async(self.close_browsers()), bg="#f44336", fg="white", font=("Arial", 12), width=22, height=2, state="disabled")
        self.close_site_button.pack(pady=10)

    def select_and_open_sites(self):
        window = tk.Toplevel(self.root)
        window.title("選擇要開啟的站台與登入帳號")
        window.geometry("600x680")
        window.grab_set()

        tk.Label(window, text="請選擇要開啟的站台：", font=("Arial", 14)).pack(pady=20)
        site_grid = tk.Frame(window)
        site_grid.pack(pady=20)

        site_order = [["TC", "FL", "XH"], ["TF", "WX", "CJ"], ["TS", "XC", "CY"], ["SY", "", ""]]
        self.site_vars = {}
        for r, row_sites in enumerate(site_order):
            for c, name in enumerate(row_sites):
                if name:
                    var = tk.BooleanVar()
                    chk = tk.Checkbutton(site_grid, text=name, variable=var, font=("Arial", 14))
                    chk.grid(row=r, column=c, padx=50, pady=12, sticky="w")
                    self.site_vars[name] = var

        credential_frame = tk.LabelFrame(window, text="登入帳號設定", font=("Arial", 12), padx=20, pady=15)
        credential_frame.pack(pady=30, fill="x", padx=40)

        tk.Label(credential_frame, text="帳號群組：", font=("Arial", 12)).grid(row=0, column=0, pady=10)
        account_group_var = tk.StringVar(value=list(ACCOUNT_GROUPS.keys())[0])
        account_combo = ttk.Combobox(credential_frame, textvariable=account_group_var, values=list(ACCOUNT_GROUPS.keys()), state="readonly", width=25, font=("Arial", 12))
        account_combo.grid(row=0, column=1, padx=20)

        tk.Label(credential_frame, text="帳號：", font=("Arial", 12)).grid(row=1, column=0)
        username_entry = tk.Entry(credential_frame, width=30, font=("Arial", 12))
        username_entry.grid(row=1, column=1, pady=10)
        
        tk.Label(credential_frame, text="密碼：", font=("Arial", 12)).grid(row=2, column=0)
        password_entry = tk.Entry(credential_frame, width=30, font=("Arial", 12), show="*")
        password_entry.grid(row=2, column=1, pady=10)

        def fill_creds():
            selected = account_group_var.get()
            username_entry.delete(0, tk.END)
            username_entry.insert(0, ACCOUNT_GROUPS[selected]["username"])
            password_entry.delete(0, tk.END)
            password_entry.insert(0, ACCOUNT_GROUPS[selected]["password"])
        
        fill_creds()
        account_combo.bind("<<ComboboxSelected>>", lambda e: fill_creds())

        def confirm():
            selected = [s for s in sites if self.site_vars.get(s["name"]) and self.site_vars[s["name"]].get()]
            if not selected:
                messagebox.showwarning("警告", "請至少選擇一個站台！")
                return
            self.selected_sites = selected
            self.run_async(self.open_browsers_with_login(username_entry.get(), password_entry.get()))
            window.destroy()

        btn_frame = tk.Frame(window)
        btn_frame.pack(pady=30)
        tk.Button(btn_frame, text="確定", command=confirm, bg="#4CAF50", fg="white", font=("Arial", 14), width=12).pack(side=tk.LEFT, padx=20)
        tk.Button(btn_frame, text="取消", command=window.destroy, font=("Arial", 14), width=12).pack(side=tk.LEFT, padx=20)

    async def open_browsers_with_login(self, username, password):
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(headless=False)
        self.context = await self.browser.new_context()
        self.pages = []
        
        for site in self.selected_sites:
            page = await self.context.new_page()
            try:
                await page.goto(site["url"])
                await page.wait_for_selector("#LoginID", timeout=10000)
                await page.fill("#LoginID", username)
                await page.fill("#Password", password)
                
                # 自動判斷登入按鈕
                login_btn = "input[type='submit'], input.btn-success" if site["name"] in ["TC", "TF"] else "input.btn.btn-primary.btn-user.btn-block"
                if await page.locator(login_btn).count() > 0:
                    await page.click(login_btn)
            except:
                pass
            self.pages.append(page)
        
        self.root.after(0, self._enable_ui_after_open)

    def _enable_ui_after_open(self):
        self.set_buttons_state("normal")
        self.open_site_button.config(state="disabled")
        self.is_running = False # 初始狀態解鎖
        messagebox.showinfo("就緒", "分頁已開啟，請點擊功能按鈕。")

    async def close_browsers(self):
        if self.browser: await self.browser.close()
        if self.playwright: await self.playwright.stop()
        self.pages = []
        self.is_running = False
        self.root.after(0, self._disable_ui_after_close)

    def _disable_ui_after_close(self):
        self.set_buttons_state("disabled")
        self.open_site_button.config(state="normal")

    def on_closing(self):
        self.run_async(self.close_browsers())
        self.root.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()