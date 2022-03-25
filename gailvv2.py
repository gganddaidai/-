from PySide2.QtWidgets import QTableWidgetItem, QApplication, QMessageBox  # , QInputDialog, QTableWidgetItem
from PySide2.QtUiTools import QUiLoader
from PySide2.QtGui import QIcon
from random import uniform
from xlwt import Workbook


class mwin:
    def __init__(self):
        self.ui = QUiLoader().load('untitled.ui')
        self.ui.pushButton.clicked.connect(self.send)
        self.ui.pushButton_2.clicked.connect(self.dao)
        self.snum = 0
        self.head = []

    def send(self):
        money = self.ui.lineEdit.text()
        money = int(money)
        num = self.ui.lineEdit_2.text()
        num = int(num)
        self.snum = num
        ci = self.ui.lineEdit_3.text()
        ci = int(ci)
        self.ui.tableWidget.setRowCount(ci)
        self.ui.tableWidget.setColumnCount(num+1)
        i = 0
        self.head = []
        while(i < num):
            self.head = self.head + ['第'+str(i+1)+'个']
            i = i+1
        self.head = self.head + ['手气最佳']
        self.ui.tableWidget.setHorizontalHeaderLabels(self.head)
        i = 0
        while(ci > i):
            lnum = num
            lmoney = money
            all = []
            while(lnum > 1):
                re = round(uniform(0.01, lmoney*2/lnum), 2)
                lmoney = lmoney - re
                self.ui.tableWidget.setItem(i, num - lnum, QTableWidgetItem(str(re)))
                lnum = lnum - 1
                all = all + [re]
            lmoney = round(lmoney, 2)
            self.ui.tableWidget.setItem(i, num - lnum, QTableWidgetItem(str(lmoney)))
            all = all + [lmoney]
            ma = all.index(max(all))+1
            self.ui.tableWidget.setItem(i, num - lnum+1, QTableWidgetItem(str(ma)))
            i = i+1

    def dao(self):
        xl = Workbook(encoding='utf-8')
        sheet = xl.add_sheet('1', cell_overwrite_ok=False)
        i = 0
        print(self.head)
        for ea in self.head:
            sheet.write(0, i, ea)
            i = i+1
        rowcount = self.ui.tableWidget.rowCount()
        n = 1
        while(n <= rowcount):
            i = 0
            while(i <= self.snum):
                mess = self.ui.tableWidget.item(n-1, i).text()
                mess = float(mess)
                sheet.write(n, i, mess)
                i = i+1
            n = n+1
        xl.save('红包模拟.xls')
        QMessageBox.information(self.ui, '导出成功', '确认')


app = QApplication([])
app.setWindowIcon(QIcon('hb.png'))
mmin = mwin()
mmin.ui.show()
app.exec_()
