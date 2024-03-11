# 使用pyinstaller打包成.exe執行檔
# https://medium.com/pyladies-taiwan/python-%E5%B0%87python%E6%89%93%E5%8C%85%E6%88%90exe%E6%AA%94-32a4bacbe351


# import requests
# from bs4 import BeautifulSoup

# response = requests.get("http://webap0.ncue.edu.tw/deanv2/other/ob010")
# response.encoding = 'utf-8'
# soup = BeautifulSoup(response.text, "html.parser")
# result = soup.find(id='result')
# print(result)

# 載入需要的套件
import numpy as np
from util import get_num_column_dict
from util import is_contain_chinese
from util import getSyllabusColumns

from datetime import datetime
import openpyxl
import pandas as pd
import os
# import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import re
import csv
import sys


# 改搜尋資料改這裡
if __name__ == "__main__":
    while True:
        try:
            print("請輸入搜尋條件或輸入q離開...")

            year = input("學年度: ")
            if year == "q":
                sys.exit(0)
            if re.match(r"^(9[5-9]|1[0-4][0-9]|150)$", year) is None:
                raise ValueError("[必填欄位] 學年度錯誤(例: 112)")

            semester = input("學期: ")
            if semester == "q":
                sys.exit(0)
            if re.match(r"^[1234]{1}$", semester) is None:
                raise ValueError("[必填欄位] 學期錯誤(例: 1)")

            crsid = input("課程代碼(為空則查詢全部): ")
            if crsid == "q":
                sys.exit(0)
            if re.match(r"^$|([A-Z0-9]{2}\d{3})$", crsid) is None:
                raise ValueError("[非必填欄位] 課程代碼僅限英數字")
        except ValueError as err:
            print("輸入有誤，請檢查輸入的值")
            print("錯誤訊息: ", err.args)
            os.system("pause")
            continue
        else:
            break
else:
    runSearch(112, 2)
