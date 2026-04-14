# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'shipments_window.ui'
##
## Created by: Qt User Interface Compiler version 6.11.0
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QMainWindow, QMenuBar, QPushButton,
    QSizePolicy, QStatusBar, QTableWidget, QTableWidgetItem,
    QTextBrowser, QVBoxLayout, QWidget)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(1280, 720)
        MainWindow.setMinimumSize(QSize(1280, 720))
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.verticalLayout = QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.manMenu_frame = QFrame(self.centralwidget)
        self.manMenu_frame.setObjectName(u"manMenu_frame")
        self.manMenu_frame.setMinimumSize(QSize(0, 50))
        self.manMenu_frame.setMaximumSize(QSize(1280, 50))
        self.manMenu_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.manMenu_frame.setFrameShadow(QFrame.Shadow.Raised)
        self.horizontalLayout = QHBoxLayout(self.manMenu_frame)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.dateUpdated_textBrowser = QTextBrowser(self.manMenu_frame)
        self.dateUpdated_textBrowser.setObjectName(u"dateUpdated_textBrowser")
        self.dateUpdated_textBrowser.setMaximumSize(QSize(150, 50))

        self.horizontalLayout.addWidget(self.dateUpdated_textBrowser)

        self.newShipment_label = QLabel(self.manMenu_frame)
        self.newShipment_label.setObjectName(u"newShipment_label")

        self.horizontalLayout.addWidget(self.newShipment_label)

        self.newMelbourne_pushButton = QPushButton(self.manMenu_frame)
        self.newMelbourne_pushButton.setObjectName(u"newMelbourne_pushButton")

        self.horizontalLayout.addWidget(self.newMelbourne_pushButton)

        self.newSydney_pushButton = QPushButton(self.manMenu_frame)
        self.newSydney_pushButton.setObjectName(u"newSydney_pushButton")

        self.horizontalLayout.addWidget(self.newSydney_pushButton)

        self.newSaba_pushButton = QPushButton(self.manMenu_frame)
        self.newSaba_pushButton.setObjectName(u"newSaba_pushButton")

        self.horizontalLayout.addWidget(self.newSaba_pushButton)


        self.verticalLayout.addWidget(self.manMenu_frame)

        self.mainshipping_table = QTableWidget(self.centralwidget)
        self.mainshipping_table.setObjectName(u"mainshipping_table")
        self.mainshipping_table.setMinimumSize(QSize(0, 250))

        self.verticalLayout.addWidget(self.mainshipping_table)

        self.sydneyShipping_table = QTableWidget(self.centralwidget)
        self.sydneyShipping_table.setObjectName(u"sydneyShipping_table")
        self.sydneyShipping_table.setMaximumSize(QSize(16777215, 165))

        self.verticalLayout.addWidget(self.sydneyShipping_table)

        self.sabaShipping_table = QTableWidget(self.centralwidget)
        self.sabaShipping_table.setObjectName(u"sabaShipping_table")
        self.sabaShipping_table.setMaximumSize(QSize(16777215, 168))

        self.verticalLayout.addWidget(self.sabaShipping_table)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 1280, 33))
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.newShipment_label.setText(QCoreApplication.translate("MainWindow", u"Add New Shipment", None))
        self.newMelbourne_pushButton.setText(QCoreApplication.translate("MainWindow", u"Melbourne", None))
        self.newSydney_pushButton.setText(QCoreApplication.translate("MainWindow", u"Sydney", None))
        self.newSaba_pushButton.setText(QCoreApplication.translate("MainWindow", u"SABA", None))
    # retranslateUi

