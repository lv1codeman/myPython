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
```
pip install openpyxl
```
3. 執行程式  
```
python course01.py
```
4. 資料夾內的開課查詢_[date].xlsx即為結果  
> [!Note] 
> [date]為當前日期
# Useful vscode extensions
- Text Pastry  
  Generate increment numbers.
- Black Formatter  
  python formatter
- Python  
  python language highlight
  
## 使用pyinstaller打包成執行檔
1. 安裝pyinstaller
```
pip install pyinstaller
```
2. 執行封裝
```
pyinstaller -F .\course01.py
```
3. 產生資料夾build、dist，其中dist內即有.exe檔可執行
> [!Note]
> 使用course01.exe時須將chrome webdriver放在同一資料夾內