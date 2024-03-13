import numpy as np
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import logging


def get_preload_data(dataType):
    if dataType == "classes":
        print("start Chrome")

        logging.getLogger('selenium').setLevel(logging.WARNING)

        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_experimental_option(
            'excludeSwitches', ['enable-logging'])

        driver = webdriver.Chrome(options=chrome_options)
        driver.get("http://webapt.ncue.edu.tw/DEANV2/Other/ob010")

        html = driver.page_source
        driver.close()  # 關閉瀏覽器
        print("END Chrome")
        soup = BeautifulSoup(html, "html.parser")
        classeslist = soup.find(id="ddl_scj_cls_id")
        res = classeslist.select("option")

        classes = np.array([[0, 0]])
        for item in res:
            classes = np.concatenate(
                (classes, [[item.get("value"), item.get_text()]]), axis=0
            )
        classes = np.delete(classes, 0, 0)

        return classes
    elif dataType == "courseType":
        crsType = {}
        data = list()
        data.append("素養通識-生活藝能及應用")
        data.append("校必(英)")
        data.append("跨學院通識(教)")
        data.append("軍訓選修")
        data.append("產業必修")
        data.append("社會學科")
        data.append("人文學科或自然學科")
        data.append("組選修")
        data.append("校必(國)")
        data.append("校必(通識)")
        data.append("組必修")
        data.append("跨學院通識(社體)")
        data.append("跨學院通識(管)")
        data.append("跨學院通識(文)")
        data.append("跨學院通識(科技)")
        data.append("系選修")
        data.append("師培選修")
        data.append("體育選修")
        data.append("跨學院通識(工)")
        data.append("語文選修")
        data.append("跨學院通識(理)")
        data.append("系必修")
        data.append("外系選修")
        data.append("素養通識-文化美學與文明")
        data.append("師培必修")
        data.append("自由選修")
        data.append("校必(體)")

        for i in range(len(data)):
            crsType[i] = data[i]

        return crsType


if __name__ == "__main__":
    res = get_preload_data("classes")

    for i in range(0, len(res)):
        if res[i][1] == "輔一甲":
            print(res[i][0])

    # print(res[20][1])
    # res = get_preload_data("courseType")
    # print(res[1])
