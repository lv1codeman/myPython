from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import QtGui

import sys
import interface
from search_courses import search_courses
from preprocess import get_preload_data


class myMainWindow(QMainWindow, interface.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)


def pbtn_search_click():
    msg = QMessageBox()
    msg.setWindowTitle("執行結果")
    msg.setText("查詢完成，請檢視檔案。")
    crsclassID = ""

    for i in range(0, len(classes)):
        if classes[i][1] == window.cb_crsclass.currentText():
            crsclassID = classes[i][0]

    res = search_courses(
        window.cb_year.currentText(),
        window.cb_semester.currentText(),
        window.le_crsid.text(),
        crsclassID,
        window.cb_crossclass.currentText(),
    )
    if res == 0:
        msg.exec_()


def pbtn_reset_click():
    preloader()


def close():
    print("close")
    app.quit()


def preloader(default=False):
    if default:
        # 設置學年度清單
        for i in reversed(range(95, 113)):
            window.cb_year.addItem(str(i))
        window.cb_semester.addItems(["1", "2"])
        # 設置學期清單
        window.cb_crossclass.addItems(["", "限本系", "可跨班系", "限本班"])
        # 設置開課班級清單(讀取網頁帶入)
        global classes
        classes = get_preload_data("classes")
        for i in range(0, len(classes)):
            window.cb_crsclass.addItem(classes[i][1])

    # 設定預設值
    window.cb_year.setCurrentIndex(0)
    window.cb_semester.setCurrentIndex(1)
    window.cb_crossclass.setCurrentIndex(0)
    window.cb_crsclass.setCurrentIndex(0)
    window.le_crsid.setText("")
    window.le_crsnm.setText("")


if __name__ == "__main__":
    # revert from 01
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
