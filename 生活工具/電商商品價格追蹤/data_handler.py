import pandas as pd
import os
from datetime import datetime
from config import CSV_FILE

def save_to_csv(title, discounted_price, original_price):
    """
    將爬取的數據保存到 CSV 文件
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data = {
        "Timestamp": [timestamp],
        "Title": [title],
        "DiscountedPrice": [discounted_price],
        "OriginalPrice": [original_price]
    }
    df = pd.DataFrame(data)
    
    if os.path.exists(CSV_FILE):
        df_existing = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
        df = pd.concat([df_existing, df], ignore_index=True)
    
    df.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
    print(f"Saved: {title} - Discounted: NT${discounted_price}, Original: NT${original_price} at {timestamp}")