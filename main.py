from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

import sys
import interface
from search_courses import search_courses


class myMainWindow(QMainWindow, interface.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)


def pbtn_search_click():
    msg = QMessageBox()
    msg.setWindowTitle("執行結果")
    msg.setText("查詢完成，請檢視檔案。")

    res = search_courses(window.cb_year.currentText(),
                         window.cb_semester.currentText(), window.le_crsid.text())
    if res == 0:
        x = msg.exec_()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = myMainWindow()
    window.menubar
    for i in reversed(range(95, 113)):
        window.cb_year.addItem(str(i))
    window.cb_semester.addItems(["1", "2"])

    window.pbtn_search.clicked.connect(pbtn_search_click)

    window.show()
    sys.exit(app.exec_())
