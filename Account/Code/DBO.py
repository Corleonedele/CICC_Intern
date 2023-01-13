#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

class DBO(QTabWidget):
    def __init__(self, parent = None):
        super(DBO, self).__init__()
        self.initTab()
        self.setWindowTitle("DB Operator")
        self.setGeometry(300, 300, 600, 300)
        self.show()
		
    def initTab(self):
        self.申购 = QWidget()
        self.追加 = QWidget()
        self.调减 = QWidget()
        self.提取 = QWidget()
        self.现金分红 = QWidget()
        self.分红再投 = QWidget()
        
        self.addTab(self.申购, "申购")
        self.addTab(self.追加, "追加")
        self.addTab(self.调减, "调减")
        self.addTab(self.提取, "提取")
        self.addTab(self.现金分红, "现金分红")
        self.addTab(self.分红再投, "分红再投")

        self.申购GUI()
        self.追加GUI()
        self.调减GUI()
        self.提取GUI()
        self.现金分红GUI()
        self.分红再投GUI()


    def 申购GUI(self):
        layout = QFormLayout()
        layout.addRow("产品名称", QLineEdit())
        layout.addRow("产品代码",QLineEdit())
        self.setTabText(0, "申购")
        self.申购.setLayout(layout)
            
    def 追加GUI(self):
        layout = QFormLayout()
        layout.addRow("产品名称",QLineEdit())
        layout.addRow("产品代码",QLineEdit())
        self.setTabText(1, "追加")
        self.追加.setLayout(layout)

    def 调减GUI(self):
        layout = QFormLayout()
        layout.addRow("产品名称",QLineEdit())
        layout.addRow("产品代码",QLineEdit())
        self.setTabText(2, "调减")
        self.申购.setLayout(layout)

    def 提取GUI(self):
        layout = QFormLayout()
        layout.addRow("产品名称",QLineEdit())
        layout.addRow("产品代码",QLineEdit())
        self.setTabText(3, "提取")
        self.调减.setLayout(layout)

    def 现金分红GUI(self):
        layout = QFormLayout()
        layout.addRow("产品名称",QLineEdit())
        layout.addRow("产品代码",QLineEdit())
        self.setTabText(4, "现金分红")
        self.现金分红.setLayout(layout)

    def 分红再投GUI(self):
        layout = QFormLayout()
        layout.addRow("产品名称",QLineEdit())
        layout.addRow("产品代码",QLineEdit())
        self.setTabText(5, "分红再投")
        self.分红再投.setLayout(layout)

def main():
    app = QApplication(sys.argv)
    ex = DBO()
    sys.exit(app.exec_())
	
if __name__ == '__main__':
    main()