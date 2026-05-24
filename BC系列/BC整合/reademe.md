****虛擬環境執行步驟
python -m venv venv
Set-ExecutionPolicy RemoteSigned
.\venv\Scripts\activate
deactivate

****安裝套件
pip install playwright pandas openpyxl tkcalendar auto-py-to-exe
playwright install chromium

****安裝&產出需求元件文件
pip install -r requirements.txt
pip freeze > requirements.txt

****playwright打包注意事項
C:\Users\Administrator\AppData\Local\ms-playwright
C:\Users\Administrator\AppData\Local\Programs\Python\Python313\Lib\site-packages\playwright\driver\package\
新增資料夾「.local-browsers」，並將playwright瀏覽器的資料夾複製進去一份

****修改為並行執行版本
已完成：main、1、2、3、4、5、6、7、9、10、11
未完成：8
program_1：事件紀錄抓取 
program_2：投注與盈亏抓取 
program_3：幸運抽獎抓取 
program_4：審單DATA抓取 
program_5：直屬及下級盈虧查詢 
program_6：招商觀察
program6_1：阿思團隊統計表，放在招商觀察下，彩票抓取部分可再優化
program_7：帳戶管理抓取
program_8：投注紀錄全抓取
program_9：彩種玩法統計 
program_10：彩種統計 
program_11：客服聊天抓取

****新站新增
1.main.py：新增對應的站台
2.各支program：excel轉出部分，新增對應站台名稱