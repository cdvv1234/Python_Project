import asyncio
import json
import os
import signal
import sys
from tkinter import Tk, Frame, Checkbutton, IntVar, Button, filedialog
from threading import Thread
from playwright.async_api import async_playwright
from asyncio import Event

class App:
    def __init__(self):
        self.context = None
        self.playwright = None
        self.browser = None
        self.sites = self.load_sites()
        self.vars = []
        self.root = None
        self.running_tasks = []
        self.stop_event = Event()  # 在類的初始化時添加

    def load_sites(self):
        """載入站台資訊"""
        if getattr(sys, "frozen", False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        json_path = os.path.join(base_path, "sites.json")
        if not os.path.exists(json_path):
            raise FileNotFoundError(f"找不到 {json_path} 檔案，請確保 JSON 檔案與程式在同一目錄下。")

        with open(json_path, "r", encoding="utf-8") as file:
            return json.load(file)

    async def login_to_sites(self, selected_sites):
        """登入選中的站台"""
        tasks = []
        pages = []

        async def login(site):
            try:
                page = await self.context.new_page()
                await page.goto(site["url"])
                await page.fill(site["username_selector"], site["credentials"]["username"])
                await page.fill(site["password_selector"], site["credentials"]["password"])
                await page.click(site["language_dropdown_selector"])
                await page.click(site["language_option_selector"])
                login_button_text = site["language_to_login_button"][site["selected_language"]]
                login_button_selector = f"button.v-btn span:has-text('{login_button_text}')"
                await page.wait_for_selector(login_button_selector, state="visible")
                await page.click(login_button_selector)
                print(f"Logged into {site['name']}")
                pages.append(page)
            except Exception as e:
                print(f"Failed to log into {site['name']}: {e}")

        for site in selected_sites:
            tasks.append(login(site))

        await asyncio.gather(*tasks)
        print("All selected sites are logged in.")
        return pages

    async def perform_login(self):
        """執行登入流程"""
        selected_sites = [self.sites[i] for i, var in enumerate(self.vars) if var.get() == 1]
        if not selected_sites:
            print("請選擇至少一個網站進行登入！")
            return

        # 啟動瀏覽器和上下文
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(headless=False, args=["--start-maximized"])
        self.context = await self.browser.new_context(accept_downloads=True, no_viewport=True)

        try:
            pages = await self.login_to_sites(selected_sites)
            print("登入完成，現在你可以手動點擊下載按鈕。")

            for page in pages:
                page.on("download", self.on_download)

            print("請手動點擊下載按鈕，等待下載請求被捕捉...")
            print("等待用戶操作結束...（按結束程式以退出）")

            # 等待 stop_event 被觸發
            await self.stop_event.wait()

        except asyncio.CancelledError:
            print("等待被取消，準備結束程式。")
        except Exception as e:
            print(f"登入過程中發生錯誤: {e}")
        finally:
            print("用戶操作結束，準備清理資源...")


    async def on_download(self, download):
        """處理下載事件"""
        try:
            print(f"檔案下載: {download.url}")
            filename = download.suggested_filename or download.url.split("/")[-1]
            file_path = self.select_save_location(filename)
            if file_path:
                await asyncio.wait_for(download.save_as(file_path), timeout=30)
                print(f"檔案已下載並儲存於: {file_path}")
            else:
                print("下載被取消")
        except asyncio.TimeoutError:
            print(f"下載 {download.url} 時間過長，已超時")
        except Exception as e:
            print(f"處理下載事件時發生錯誤: {e}")

    def select_save_location(self, filename):
        """選擇儲存位置"""
        self.root.lift()
        self.root.attributes("-topmost", True)
        file_path = filedialog.asksaveasfilename(initialfile=filename, title="選擇儲存檔案的位置", filetypes=[("All Files", "*.*")])
        self.root.attributes("-topmost", False)
        return file_path

    def create_gui(self):
        """建立 GUI"""
        self.root = Tk()
        self.root.title("選擇網站進行登入")
        frame = Frame(self.root)
        frame.pack(pady=10)

        for i, site in enumerate(self.sites):
            var = IntVar()
            self.vars.append(var)
            chk = Checkbutton(frame, text=site["name"], variable=var)
            chk.grid(row=i // 6, column=i % 6, padx=5, pady=5)

        login_button = Button(self.root, text="登入選中的網站", command=self.on_login)
        login_button.pack(pady=10)
        exit_button = Button(self.root, text="結束程式", command=self.on_exit)
        exit_button.pack(pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.root.mainloop()

    def on_login(self):
        """處理登入按鈕點擊"""
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)

        if loop.is_running():
            asyncio.create_task(self.perform_login())
        else:
            Thread(target=lambda: asyncio.run(self.perform_login())).start()
    


    def on_exit(self):
        """處理結束按鈕點擊"""
        def force_exit():
            print("強制退出程式...")
            os.kill(os.getpid(), signal.SIGTERM)
        
        async def cleanup():
            try:
                print("清理完成，強制退出...")
                force_exit()
            #     print("正在清理資源...")
            #     if self.stop_event.is_set():
            #         return

            #     self.stop_event.set()

            #     # 關閉所有頁面和上下文
            #     if self.context:
            #         for page in self.context.pages:
            #             try:
            #                 await asyncio.wait_for(page.close(), timeout=5)  # 限制單個頁面關閉時間
            #             except asyncio.TimeoutError:
            #                 print(f"關閉頁面 {page.url} 時超時，跳過此步驟。")
            #         try:
            #             await asyncio.wait_for(self.context.close(), timeout=10)  # 限制上下文關閉時間
            #         except asyncio.TimeoutError:
            #             print("關閉上下文超時，跳過此步驟。")
            #         self.context = None

            #     # 關閉瀏覽器
            #     if self.browser:
            #         try:
            #             await asyncio.wait_for(self.browser.close(), timeout=10)
            #         except asyncio.TimeoutError:
            #             print("關閉瀏覽器超時，跳過此步驟。")
            #         finally:
            #             self.browser = None

            #     # 停止 Playwright
            #     if self.playwright:
            #         try:
            #             await asyncio.wait_for(self.playwright.stop(), timeout=5)
            #         except asyncio.TimeoutError:
            #             print("停止 Playwright 超時，跳過此步驟。")
            #         self.playwright = None

            #     print("資源清理完成，程式即將退出。")

            except Exception as e:
                print(f"清理過程中發生錯誤: {e}")

        def run_cleanup():
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(cleanup())
            loop.close()

        # 啟動清理流程
        cleanup_thread = Thread(target=run_cleanup)
        cleanup_thread.start()
        cleanup_thread.join()  # 確保清理完成後再退出 GUI
        self.root.quit()

if __name__ == "__main__":
    app = App()
    app.create_gui()
