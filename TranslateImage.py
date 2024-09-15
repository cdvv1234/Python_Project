import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageEnhance, ImageFilter
import requests
from io import BytesIO
import pytesseract

# 如果 Tesseract OCR 沒有安裝在預設路徑，則需要指定路徑
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Windows 路徑

def preprocess_image(image):
    # 轉換為灰階
    image = image.convert('L')
    
    # 增加對比度
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    
    # 二值化處理
    image = image.point(lambda p: p > 128 and 255)
    
    # 去噪
    image = image.filter(ImageFilter.MedianFilter())
    
    return image

def image_to_text(image):
    # 指定簡體和繁體中文語言包進行 OCR 辨識
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

# 創建應用程式窗口
window = tk.Tk()
window.title("圖片文字識別")
window.geometry("800x600")

# 創建 URL 輸入區域
url_label = tk.Label(window, text="圖片 URL：")
url_label.pack()
url_entry = tk.Entry(window, width=80)
url_entry.pack(pady=5)

# 創建處理 URL 按鈕
url_button = tk.Button(window, text="處理圖片 URL", command=process_image_from_url)
url_button.pack(pady=5)

# 創建上傳圖片按鈕
upload_button = tk.Button(window, text="上傳本地圖片", command=upload_image)
upload_button.pack(pady=5)

# 創建結果顯示區域
result_label = tk.Label(window, text="識別結果：")
result_label.pack(pady=5)

# 創建結果顯示框和滾動條
result_frame = tk.Frame(window)
result_frame.pack(fill="both", expand=True)

scrollbar = tk.Scrollbar(result_frame)
scrollbar.pack(side="right", fill="y")

result_text = tk.Text(result_frame, wrap="word", yscrollcommand=scrollbar.set)
result_text.pack(fill="both", expand=True)

scrollbar.config(command=result_text.yview)

# 運行主循環
window.mainloop()
