import tkinter as tk
from tkinter import ttk
from bs4 import BeautifulSoup

def extract_content():
    html_content = text_box.get("1.0", tk.END).strip()
    if not html_content:
        return

    # 清空之前的提取結果
    output_text.delete('1.0', tk.END)

    # 定義通用的選擇器規則
    selectors = {
        "category": ".cate-name .title",
        "item": ".menu-item",
        "mark": ".hot, .new",
        "bank": ".bankBox",
        "bank_title": ".banktitle",
        "cate_item": ".cateWrapper .cate-item",
        "method_type": ".cateWrapper .method-type",
        "lottery_section": ".i-navitem-box",
        "lottery_title": ".i-navitem-h3",
        "lottery_list": ".i-navitem-list .cursor a",
        "recharge_gateway": "#recharge-gateway-ul .recharge-gateway-item a",
        "recharge_hot": ".hot-item img",
        "game_group": ".i-r-nav .game-group .ng-star-inserted a",
        "lottery_menu": ".lottery-menu-row",
        "lottery_menu_title": ".lottery-menu-row__title span",
        "lottery_menu_button": ".lottery-menu-row__content .lottery-tag-wrapper button span",
        "lottery_tag_new": ".lottery-tag-new",
        "lottery_tag_hot": ".lottery-tag-popular",
        "vr_menu_row": ".menu-row",
        "vr_menu_title": ".menu-row__title span",
        "vr_menu_button": ".menu-row__content button span"
    }

    # 使用BeautifulSoup解析HTML
    soup = BeautifulSoup(html_content, 'lxml')

    try:
        # 處理彩票內容
        categories = soup.select(selectors["category"])
        for category in categories:
            category_title = category.get_text(strip=True)
            items = category.find_next_sibling('ul').select(selectors["item"])
            for item in items:
                item_text = item.get_text(strip=True)
                mark = item.select_one(selectors["mark"])
                mark_text = mark.get_text(strip=True) if mark else ''
                output_text.insert(tk.END, f"{category_title}\n{item_text}\n{mark_text}\n")

        # 處理銀行選項內容
        banks = soup.select(selectors["bank"])
        for bank in banks:
            bank_title = bank.select_one(selectors["bank_title"])
            bank_text = bank_title.get_text(strip=True) if bank_title else '無標題'
            output_text.insert(tk.END, f"{bank_text}\n")

        # 處理類別項目
        cate_items = soup.select(selectors["cate_item"])
        for item in cate_items:
            item_text = item.get_text(strip=True)
            output_text.insert(tk.END, f"類別項目: {item_text}\n")

        # 處理方法類型
        method_type = soup.select_one(selectors["method_type"])
        method_text = method_type.get_text(strip=True) if method_type else '無方法類型'
        output_text.insert(tk.END, f"方法類型: {method_text}\n")

        # 處理彩票類型和其選項
        lottery_sections = soup.select(selectors["lottery_section"])
        for section in lottery_sections:
            section_title = section.select_one(selectors["lottery_title"]).get_text(strip=True)
            output_text.insert(tk.END, f"\n彩票類型: {section_title}\n")
            lottery_items = section.select(selectors["lottery_list"])
            for item in lottery_items:
                item_text = item.get_text(strip=True)
                output_text.insert(tk.END, f"  - {item_text}\n")

        # 處理充值方式
        recharge_gateways = soup.select(selectors["recharge_gateway"])
        for gateway in recharge_gateways:
            gateway_text = gateway.get_text(strip=True)
            hot_item = gateway.find_next_sibling(selectors["recharge_hot"])
            hot_text = hot_item.get_text(strip=True) if hot_item else ''
            output_text.insert(tk.END, f"充值方式: {gateway_text}\n{hot_text}\n")

        # 處理遊戲組別
        game_groups = soup.select(selectors["game_group"])
        for group in game_groups:
            group_text = group.get_text(strip=True)
            output_text.insert(tk.END, f"遊戲組別: {group_text}\n")

        # 處理彩票選單
        lottery_menus = soup.select(selectors["lottery_menu"])
        for menu in lottery_menus:
            menu_title = menu.select_one(selectors["lottery_menu_title"]).get_text(strip=True)
            output_text.insert(tk.END, f"\n彩票選單: {menu_title}\n")
            buttons = menu.select(selectors["lottery_menu_button"])
            for button in buttons:
                button_text = button.get_text(strip=True)
                output_text.insert(tk.END, f"  - {button_text}\n")
                # 處理新標籤
                if button.find_parent(selectors["lottery_tag_new"]):
                    output_text.insert(tk.END, "    新\n")
                # 處理熱門標籤
                if button.find_parent(selectors["lottery_tag_hot"]):
                    output_text.insert(tk.END, "    熱門\n")

        # 處理VR娛樂
        vr_menus = soup.select(selectors["vr_menu_row"])
        for menu in vr_menus:
            menu_title = menu.select_one(selectors["vr_menu_title"]).get_text(strip=True)
            output_text.insert(tk.END, f"\nVR娛樂: {menu_title}\n")
            buttons = menu.select(selectors["vr_menu_button"])
            for button in buttons:
                button_text = button.get_text(strip=True)
                output_text.insert(tk.END, f"  - {button_text}\n")
                # 處理熱門標籤
                if button.find_parent(selectors["lottery_tag_hot"]):
                    output_text.insert(tk.END, "    熱門\n")

    except Exception as e:
        output_text.insert(tk.END, f"錯誤: {str(e)}\n")

# 創建主窗口
root = tk.Tk()
root.title("HTML內容提取器")

# 創建文本框用於輸入HTML內容
text_box = tk.Text(root, wrap=tk.WORD, height=15, width=100)
text_box.pack(padx=10, pady=10)

# 創建提取按鈕
extract_button = tk.Button(root, text="提取內容", command=extract_content)
extract_button.pack(pady=5)

# 創建文本小部件來顯示提取的內容
output_text = tk.Text(root, wrap=tk.NONE, height=25, width=100)
output_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# 使文本小部件可複製
output_text.config(state=tk.NORMAL)

root.mainloop()
