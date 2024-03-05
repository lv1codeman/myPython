# Introduction
利用selenium爬蟲開課查詢的結果，透過BeautifulSoup4解析後存入.csv檔案
# Feature
1. 將課程名稱中英文、教學大綱分開存在不同欄位
2. 授課教師欄位以逗號分隔
# Usage
1. 下載[Python](https://www.python.org/downloads/)  
  安裝時勾選add Python to PATH
2. 開啟vscode terminal安裝使用到的模組
```
pip install pandas
```
```
pip install selenium
```
```
pip install BeautifulSoup4
```
3. 執行程式  
```
python course01.py
```
4. 資料夾內的1122_course01.csv即為結果
> [!Note] 
> 檔名開頭為當前學年度學期
# Useful vscode extensions
- Text Pastry  
  Generate increment numbers.
- Black Formatter  
  python formatter
- Python  
  python language highlight
  
