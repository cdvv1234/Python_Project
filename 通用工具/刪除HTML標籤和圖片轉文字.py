import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageEnhance, ImageFilter
import requests
from io import BytesIO
import pytesseract
from bs4 import BeautifulSoup

# 如果 Tesseract OCR 沒有安裝在預設路徑，則需要指定路徑
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Windows 路徑

# 影像預處理
def preprocess_image(image):
    image = image.convert('L')  # 轉換為灰階
    enhancer = ImageEnhance.Contrast(image)  # 增加對比度
    image = enhancer.enhance(2.0)
    image = image.point(lambda p: p > 128 and 255)  # 二值化處理
    image = image.filter(ImageFilter.MedianFilter())  # 去噪
    return image

# 圖片文字識別
def image_to_text(image):
    lang = 'chi_sim+chi_tra'
    text = pytesseract.image_to_string(image, lang=lang)
    return text

def process_image_from_url():
    url = url_entry.get()
    if not url:
        messagebox.showerror("錯誤", "請輸入圖片 URL。")
        return
    try:
        response = requests.get(url)
        image = Image.open(BytesIO(response.content))
        image = preprocess_image(image)
        text = image_to_text(image)
        result_text.delete("1.0", "end")
        result_text.insert("1.0", text)
    except Exception as e:
        messagebox.showerror("錯誤", f"處理圖片時出錯：{e}")

def upload_image():
    file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")])
    if not file_path:
        return
    try:
        image = Image.open(file_path)
        image = preprocess_image(image)
        text = image_to_text(image)
        result_text.delete("1.0", "end")
        result_text.insert("1.0", text)
    except Exception as e:
        messagebox.showerror("錯誤", f"處理圖片時出錯：{e}")

# HTML 標籤解析
def format_text(text):
    lines = text.splitlines()
    formatted_text = ""
    for line in lines:
        line = line.strip()
        if line:
            if any(char.isdigit() for char in line):
                formatted_text += f"{line}\n"
            else:
                formatted_text += f"{line}\n"
    return formatted_text

def extract_text():
    html_content = html_input.get("1.0", "end")
    soup = BeautifulSoup(html_content, 'html.parser')
    pure_text = soup.get_text(separator="\n")
    formatted_text = format_text(pure_text)
    html_result_output.delete("1.0", "end")
    html_result_output.insert("1.0", formatted_text)

# 創建應用程式窗口
window = tk.Tk()
window.title("多功能處理工具")
window.geometry("800x600")

# 創建功能選擇區域
def show_image_tab():
    html_tab.pack_forget()
    image_tab.pack(fill="both", expand=True)
    # 確保結果顯示框只顯示圖片結果
    result_text.delete("1.0", "end")

def show_html_tab():
    image_tab.pack_forget()
    html_tab.pack(fill="both", expand=True)
    # 確保結果顯示框只顯示 HTML 結果
    html_result_output.delete("1.0", "end")

# 創建選擇功能的按鈕區域
button_frame = tk.Frame(window)
button_frame.pack(side="left", fill="y")

image_button = tk.Button(button_frame, text="圖片文字識別", command=show_image_tab)
image_button.pack(pady=5, fill="x")

html_button = tk.Button(button_frame, text="HTML 標籤過濾器", command=show_html_tab)
html_button.pack(pady=5, fill="x")

# 圖片文字識別區域
image_tab = tk.Frame(window)

url_label = tk.Label(image_tab, text="圖片 URL：")
url_label.pack()

url_entry = tk.Entry(image_tab, width=80)
url_entry.pack(pady=5)

url_button = tk.Button(image_tab, text="處理圖片 URL", command=process_image_from_url)
url_button.pack(pady=5)

upload_button = tk.Button(image_tab, text="上傳本地圖片", command=upload_image)
upload_button.pack(pady=5)

result_label = tk.Label(image_tab, text="識別結果：")
result_label.pack(pady=5)

result_frame = tk.Frame(image_tab)
result_frame.pack(fill="both", expand=True)

scrollbar = tk.Scrollbar(result_frame)
scrollbar.pack(side="right", fill="y")

result_text = tk.Text(result_frame, wrap="word", yscrollcommand=scrollbar.set)
result_text.pack(fill="both", expand=True)

scrollbar.config(command=result_text.yview)

# HTML 標籤解析區域
html_tab = tk.Frame(window)

html_label = tk.Label(html_tab, text="請貼上 HTML 內容：")
html_label.pack()

html_input = tk.Text(html_tab, height=10)
html_input.pack(fill="x")

html_extract_button = tk.Button(html_tab, text="提取純文本", command=extract_text)
html_extract_button.pack(pady=10)

html_result_label = tk.Label(html_tab, text="提取出的純文本內容：")
html_result_label.pack()

html_result_frame = tk.Frame(html_tab)
html_result_frame.pack(fill="both", expand=True)

html_scrollbar = tk.Scrollbar(html_result_frame)
html_scrollbar.pack(side="right", fill="y")

html_result_output = tk.Text(html_result_frame, wrap="word", yscrollcommand=html_scrollbar.set)
html_result_output.pack(fill="both", expand=True)

html_scrollbar.config(command=html_result_output.yview)

# 設置預設顯示為圖片處理選項卡
show_image_tab()

# 運行主循環
window.mainloop()
