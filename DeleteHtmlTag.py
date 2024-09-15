import tkinter as tk
from bs4 import BeautifulSoup

# 創建應用程式窗口
window = tk.Tk()
window.title("HTML 標籤過濾器")

# 設定窗口尺寸
window.geometry("600x400")

# 創建標題標籤
title_label = tk.Label(window, text="請貼上 HTML 內容：")
title_label.pack()

# 創建 HTML 輸入區域
html_input = tk.Text(window, height=10)
html_input.pack(fill="x")

# 創建提取按鈕，並將其放置在輸入區域之下
extract_button = tk.Button(window, text="提取純文本", command=lambda: extract_text())
extract_button.pack(pady=10)

# 創建顯示結果的區域標籤
result_label = tk.Label(window, text="提取出的純文本內容：")
result_label.pack()

# 創建框架來放置結果框和滾動條
result_frame = tk.Frame(window)
result_frame.pack(fill="both", expand=True)

# 創建滾動條
scrollbar = tk.Scrollbar(result_frame)
scrollbar.pack(side="right", fill="y")

# 創建結果顯示框，並連接滾動條
result_output = tk.Text(result_frame, wrap="word", yscrollcommand=scrollbar.set)
result_output.pack(fill="both", expand=True)

# 設置滾動條控制結果框的滾動
scrollbar.config(command=result_output.yview)

# 定義函數進行排版處理並去除多餘空白行


def format_text(text):
    lines = text.splitlines()  # 將文本按行分割
    formatted_text = ""

    for line in lines:
        line = line.strip()  # 去除行首行尾的空白字符
        if line:  # 只保留非空行
            # 如果是數字，將其加上 "數值結果" 標記
            if any(char.isdigit() for char in line):
                formatted_text += f"{line}\n"
            else:
                formatted_text += f"{line}\n"

    return formatted_text

# 定義按鈕的功能：解析HTML並顯示排版後的純文本


def extract_text():
    html_content = html_input.get("1.0", "end")  # 獲取輸入框中的HTML內容
    soup = BeautifulSoup(html_content, 'html.parser')  # 使用BeautifulSoup解析HTML
    pure_text = soup.get_text(separator="\n")  # 使用分行符號提取純文本
    formatted_text = format_text(pure_text)  # 對提取的文本進行排版處理並去除多餘空行
    result_output.delete("1.0", "end")  # 清空結果顯示區
    result_output.insert("1.0", formatted_text)  # 將排版後的純文本顯示在結果框中

# 運行主循環


window.mainloop()
