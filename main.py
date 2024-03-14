from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import QtGui

import sys
import interface
from search_courses import search_courses
from preprocess import get_preload_data

# TODO:
# 未製作的查詢：開課名稱、老師姓名、星期、遠距課程、全英語課程
# 未製作的非原生查詢：課程性質


class myMainWindow(QMainWindow, interface.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)


def pbtn_reset_click():
    preloader()


def close():
    print("close")
    app.quit()


def pbtn_search_click():
    msg = QMessageBox()
    msg.setWindowTitle("執行結果")
    msg.setText("查詢完成，請檢視檔案。")

    crsclassID = ""
    for i in range(0, len(classes)):
        if classes[i][1] == window.cb_crsclass.currentText():
            crsclassID = classes[i][0]

    # 更改鼠標為讀取圖示
    QApplication.setOverrideCursor(Qt.WaitCursor)

    res = search_courses(
        window.cb_year.currentText(),
        window.cb_semester.currentText(),
        window.le_crsid.text(),
        crsclassID,
        window.cb_crossclass.currentText(),
    )
    if res == 0:
        # 還原鼠標
        QApplication.restoreOverrideCursor()
        msg.exec_()


def preloader(isFirstRun=False):
    if isFirstRun:
        # 寫入下拉選單選項
        window.cb_semester.addItems(["1", "2"])
        window.cb_crossclass.addItems(["", "限本系", "可跨班系", "限本班"])

        # 網頁讀取下拉選單選項
        # 設置學年度清單
        res = get_preload_data("preload_select_area")
        for i in res["years"]:
            window.cb_year.addItem(i)

        # 設置開課班級清單
        # 查詢按鈕按下時需要提供開課班級列表來比對使用者選的班級是哪個班級代碼
        # 所以要把classes寫成全域變數提供def pbtn_search_click()使用
        global classes
        classes = res["classes"]
        for i in range(0, len(classes)):
            window.cb_crsclass.addItem(classes[i][1])

    # 設定預設值
    # TODO:學年度的預設值選第二項，以後要根據當前時間更改預設值
    window.cb_year.setCurrentIndex(1)
    window.cb_semester.setCurrentIndex(1)
    window.cb_crossclass.setCurrentIndex(0)
    window.cb_crsclass.setCurrentIndex(0)
    window.le_crsid.setText("")
    window.le_crsnm.setText("")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = myMainWindow()
    window.setWindowIcon(QtGui.QIcon("sixer-small.jpg"))
    window.setWindowTitle("開課查詢")
    window.actionexit.triggered.connect(close)

    window.pbtn_search.clicked.connect(pbtn_search_click)
    window.pbtn_reset.clicked.connect(pbtn_reset_click)

    preloader(True)
    try:
        import pyi_splash

        pyi_splash.close()
    except ImportError:
        pass

    # set validator
    reg = QRegExp(r"^$|([A-Z0-9]{2}\d{3})$")
    regVal = QRegExpValidator()
    regVal.setRegExp(reg)
    window.le_crsid.setValidator(regVal)

    window.show()
    sys.exit(app.exec_())
