import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from config import CSV_FILE
import os

def plot_price_trend():
    """
    繪製價格趨勢圖
    """
    if not os.path.exists(CSV_FILE):
        print(f"CSV file {CSV_FILE} does not exist.")
        return

    # 讀取 CSV
    df = pd.read_csv(CSV_FILE)
    if df.empty:
        print("No data in CSV file.")
        return

    # 轉換時間戳為日期格式
    df['Timestamp'] = pd.to_datetime(df['Timestamp'])
    
    # 設置中文字型
    font_path = "C:/Windows/Fonts/msjh.ttc"  # 微軟正黑體，Windows 預設路徑
    if not os.path.exists(font_path):
        font_path = "NotoSansCJK-Regular.ttc"  # 備用：思源黑體，需下載
    if os.path.exists(font_path):
        font_prop = fm.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
    else:
        print("Warning: Chinese font not found. Using default font.")
        plt.rcParams['font.family'] = 'sans-serif'

    # 繪製圖表
    plt.figure(figsize=(10, 6))
    plt.plot(df['Timestamp'], df['DiscountedPrice'], marker='o', label='Discounted Price')
    if 'OriginalPrice' in df.columns and not df['OriginalPrice'].isna().all():
        plt.plot(df['Timestamp'], df['OriginalPrice'], marker='x', label='Original Price')
    
    # 設置標題和標籤
    plt.title(f"Price Trend of {df['Title'].iloc[0]}", fontproperties=font_prop if os.path.exists(font_path) else None)
    plt.xlabel("Date", fontproperties=font_prop if os.path.exists(font_path) else None)
    plt.ylabel("Price (NT$)", fontproperties=font_prop if os.path.exists(font_path) else None)
    plt.legend(prop=font_prop if os.path.exists(font_path) else None)
    plt.grid(True)
    
    # 旋轉 x 軸標籤
    plt.xticks(rotation=45)
    
    # 調整佈局並保存
    plt.tight_layout()
    plt.savefig('price_trend.png')
    plt.close()