from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import QtGui

import sys
import interface
from search_courses import search_courses
from preprocess import get_preload_data
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
from selenium.webdriver.chrome.options import Options
import logging
import numpy as np
import time

options = Options()
driver = webdriver.Chrome(options=options)
driver.get("http://webapt.ncue.edu.tw/DEANV2/Other/ob010")
print("END Chrome")
selectyear = Select(driver.find_element(By.ID, "ddl_yms_year"))
selectyear.select_by_value('111')
res = {}
time.sleep(1)
# Wait for the class dropdown to become available
try:
    WebDriverWait(driver, 30).until(
        EC.text_to_be_present_in_element_value((By.ID, "ddl_yms_year"), '111'))
except TimeoutException:
    print("Timeout while waiting for dropdown value to change")

html = driver.page_source
soup = BeautifulSoup(html, "html.parser")
class_list = soup.find(id="ddl_scj_cls_id").select("option")
classes = np.array([[0, 0]])
