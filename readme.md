# Introduction
main.py - 主程式  
interface.py - GUI介面  
util.py - 功能包  
search_courses.py - 主查詢功能，獨立執行時以CMD介面接收輸入  
preprocess.py - 預先讀取網頁上下拉選單的選項，在main.py設定到GUI介面上
# Feature
利用selenium爬蟲開課查詢的結果，透過BeautifulSoup4解析後存入.csv檔案，再將.csv檔轉存為.xlsx檔。  
1. 將課程名稱中英文、教學大綱分開存在不同欄位
2. 授課教師欄位以逗號分隔
# Usage
1. 下載[Python](https://www.python.org/downloads/)  
  安裝時勾選add Python to PATH
2. 開啟vscode terminal安裝使用到的Library  
3. 執行程式  
```
python main.py
```
4. 資料夾內的.xlsx檔即為結果  
> [!Note] 
> [date]為當前日期  

# 需安裝的Library
Webdriver
```
pip install selenium
```
For scrape information from web pages
```
pip install BeautifulSoup4
```
Excel module
```
pip install openpyxl
```
GUI
```
pip install PyQt5
```
Merge excel
```
pip install pandas
```


# Useful vscode extensions
- Text Pastry  
  Generate increment numbers.
- Black Formatter  
  python formatter
- Python  
  python language highlight
  

# 建立GUI介面
1. 下載[Qt Designer](https://build-system.fman.io/qt-designer-download)  
2. 用Qt Designer繪製GUI介面，另存為.ui檔案  
3. 執行以下指令生成GUI介面的.py檔案
```
pyuic5 -x [filename].ui -o [filename].py
```
```
pyuic5 -x interface.ui -o interface.py
```
4. 可於[filename].py編輯GUI介面，建議將配置寫在main.py，例如按鈕的動作、下拉選單設定預設值...等。
## 使用pyinstaller打包成.exe執行檔
[參考文件](https://medium.com/pyladies-taiwan/python-%E5%B0%87python%E6%89%93%E5%8C%85%E6%88%90exe%E6%AA%94-32a4bacbe351)
1. 安裝pyinstaller
```
pip install pyinstaller
```
2. 執行封裝  
-F 打包成一個.exe檔  
-w 執行.exe時不顯示console  
-n [filename].exe 命名.exe檔  
-i="[filename].ico" 設置.exe檔的icon 
--splash opening.png 設置程式開啟時的等候圖示
```
pyinstaller -F -w -n course_search.exe --splash opening.png main.py
```
3. 產生資料夾build、dist，其中dist內即有.exe檔可執行
> [!Note]
> 執行時須將chrome webdriver放在同一資料夾內