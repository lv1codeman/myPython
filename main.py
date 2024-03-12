from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

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

    for i in range(0, len(classes)):
        if classes[i][1] == window.cb_crsclass.currentText():
            crsclassID = classes[i][0]

    res = search_courses(
        window.cb_year.currentText(),
        window.cb_semester.currentText(),
        window.le_crsid.text(),
        crsclassID,
        window.cb_crossclass.currentText()
    )
    if res == 0:
        msg.exec_()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = myMainWindow()
    window.menubar
    for i in reversed(range(95, 113)):
        window.cb_year.addItem(str(i))
    window.cb_semester.addItems(["1", "2"])
    window.cb_semester.setCurrentIndex(1)
    window.pbtn_search.clicked.connect(pbtn_search_click)

    window.cb_crossclass.addItems(["", "限本系", "可跨班系", "限本班"])
    window.cb_crossclass.setCurrentIndex(0)

    classes = get_preload_data("classes")

    for i in range(0, len(classes)):
        window.cb_crsclass.addItem(classes[i][1])
    window.cb_crsclass.adjustSize()

    window.show()
    sys.exit(app.exec_())
