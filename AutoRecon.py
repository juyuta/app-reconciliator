import re
import os
import sys
import time
import sqlite3
import logging
import platform
import datetime
import warnings
import numpy as np
import pandas as pd
from os import path
import win32com.client as w3c
from win32com.client import Dispatch
from xlsxwriter.workbook import Workbook
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QBrush, QColor, QCursor, QFont, QIcon, QPalette, QPixmap, QMovie
from PyQt5.QtCore import QCoreApplication, QMetaObject, QRect, QThread, QSize, QUrl, Qt, QDir, pyqtSlot, pyqtSignal
from PyQt5.QtWidgets import QApplication, QFileDialog, QGraphicsDropShadowEffect, QScrollArea, QComboBox, QToolButton, QWidget, QMessageBox, QPushButton, QScrollBar, QSizePolicy, QListWidget, QFrame, QSpacerItem, QStackedWidget, QTextBrowser, QSizeGrip, QLabel, QMainWindow, QGridLayout

warnings.filterwarnings("ignore")
logging.basicConfig(filename='Log\debug.log', level=logging.DEBUG, format='%(asctime)s:%(levelname)s:%(message)s')

GLOBAL_STATE = 0

class MainWindow(QMainWindow):
    ### MainWindow TO INITIALIZE THE USER INTERFACE ###
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        def moveWindow(event):
            # Allows moving of the window with a configured title bar
            if UIFunctions.returnStatus == 1:
                UIFunctions.maximize_restore(self)

            if event.buttons() == Qt.LeftButton:
                self.move(self.pos() + event.globalPos() - self.dragPos)
                self.dragPos = event.globalPos()
                event.accept()

        self.ui.name.mouseMoveEvent = moveWindow

        UIFunctions.uiDefinitions(self)

        self.show()

    def mousePressEvent(self, event):
        self.dragPos = event.globalPos()

class UIFunctions(MainWindow):
    ### UIFunctions HELPS WITH REMOVING THE DEFAULT WINDOW LAYOUT + ALLOW MINIMIZE/MAXIMIZE/CLOSE BUTTONS TO WORK AS IT SHOULD ###
    # ==> MAXIMIZE RESTORE FUNCTION
    def maximize_restore(self):
        global GLOBAL_STATE
        status = GLOBAL_STATE

        # IF NOT MAXIMIZED
        if status == 0:
            self.showMaximized()

            # SET GLOBAL TO 1
            GLOBAL_STATE = 1

            # IF MAXIMIZED REMOVE MARGINS AND BORDER RADIUS
            self.ui.body_layout.setContentsMargins(0, 0, 0, 0)
        #     self.ui.drop_shadow_frame.setStyleSheet(
        #         "border-radius: 0px; background-color: rgb(0, 26, 51);")
            self.ui.maximize_button.setIcon(self.ui.icon9)

        else:
            GLOBAL_STATE = 0
            self.showNormal()
            self.resize(self.width(), self.height())
        #     self.ui.drop_shadow_frame.setStyleSheet(
        #         "border-radius: 10px; background-color: rgb(0, 26, 51);")
            self.ui.maximize_button.setIcon(self.ui.icon2)

    # ==> UI DEFINITIONS
    def uiDefinitions(self):

        # REMOVE TITLE BAR
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        # SET DROPSHADOW WINDOW
        self.shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        self.shadow.setBlurRadius(20)
        self.shadow.setXOffset(0)
        self.shadow.setYOffset(0)
        self.shadow.setColor(QColor(0, 0, 0, 100))

        # MAXIMIZE / RESTORE
        self.ui.maximize_button.clicked.connect(
            lambda: UIFunctions.maximize_restore(self))

        # MINIMIZE
        self.ui.minimize_button.clicked.connect(lambda: self.showMinimized())

        # CLOSE
        self.ui.close_button.clicked.connect(lambda: self.close())

        # ==> CREATE SIZE GRIP TO RESIZE WINDOW
        # self.sizegrip = QSizeGrip(self.ui.frame_grip)
        # self.sizegrip.setStyleSheet(
        #     "QSizeGrip { width: 10px; height: 10px; margin: 5px } QSizeGrip:hover { background-color: rgb(50, 42, 94) }")
        # self.sizegrip.setToolTip("Resize Window")

    # RETURN STATUS IF WINDOWS IS MAXIMIZED OR RESTORED
    def returnStatus(self):
        return GLOBAL_STATE

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        ### INITIALIZE GUI ###
        if MainWindow.objectName():
            MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1200, 900)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(1200, 900))
        MainWindow.setAutoFillBackground(False)
        self.central_widget = QtWidgets.QWidget(MainWindow)
        self.central_widget.setMinimumSize(QtCore.QSize(1200, 900))
        self.central_widget.setStyleSheet('''
                                        QScrollbar:vertical {
                                            border: none;
                                            width: 14px;
                                            margin: 15px 0 15px 0;
                                            border-radius: 7px;
                                        }    
                                        QScrollBar::handle:vertical {
                                            background-color: rgb(80, 80, 122);
                                            min-height: 30px;
                                            border-radius: 7px;
                                        }
                                        QScrollBar::handle:vertical:hover {
                                            background-color: rgb(200, 200, 200);
                                        }
                                        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                                            background: none;
                                        }
                                        QScrollBar::sub-line:vertical {
                                            border:none;
                                        }
                                        QScrollBar::add-line:vertical {
                                            border: none;
                                        }
                                        ''')
        self.central_widget.setObjectName("central_widget")
        self.centralwidget_layout = QtWidgets.QVBoxLayout(self.central_widget)
        self.centralwidget_layout.setContentsMargins(0, 0, 0, 0)
        self.centralwidget_layout.setSpacing(0)
        self.centralwidget_layout.setObjectName("centralwidget_layout")

        # DECLARE ALL FONTS
        font = QtGui.QFont()
        font.setFamily("Leelawadee")
        font1 = QtGui.QFont("Leelawadee", 15, 75) # Font type, font size, weight sizes
        font2 = QtGui.QFont("Leelawadee", 14, 75)
        font3 = QtGui.QFont("Leelawadee", 10, 75)
        font4 = QtGui.QFont("Leelawadee", 10)
        font5 = QtGui.QFont("Leelawadee", 9, 75)

        # DECLARE ICONS
        self.icon1 = QtGui.QIcon()
        self.icon1.addPixmap(QtGui.QPixmap(r"icons\minus-sign.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon2 = QtGui.QIcon()
        self.icon2.addPixmap(QtGui.QPixmap(r"icons\resize.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon3 = QtGui.QIcon()
        self.icon3.addPixmap(QtGui.QPixmap(r"icons\cancel.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon4 = QtGui.QIcon()
        self.icon4.addPixmap(QtGui.QPixmap(r"icons\next.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon5 = QtGui.QIcon()
        self.icon5.addPixmap(QtGui.QPixmap(r"icons\upload.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon6 = QtGui.QIcon()
        self.icon6.addPixmap(QtGui.QPixmap(r"icons\checklist.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon7 = QtGui.QIcon()
        self.icon7.addPixmap(QtGui.QPixmap(r"icons\delete.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon8 = QtGui.QIcon()
        self.icon8.addPixmap(QtGui.QPixmap(r"icons\reload.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon9 = QtGui.QIcon()
        self.icon9.addPixmap(QtGui.QPixmap(r"icons\restore.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon10 = QtGui.QIcon()
        self.icon10.addPixmap(QtGui.QPixmap(r"icons\back.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)

        self.header = QtWidgets.QFrame(self.central_widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.header.sizePolicy().hasHeightForWidth())
        self.header.setSizePolicy(sizePolicy)
        self.header.setMinimumSize(QtCore.QSize(0, 75))
        self.header.setStyleSheet("background-color: rgb(49, 54, 59)")
        self.header.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.header.setFrameShadow(QtWidgets.QFrame.Plain)
        self.header.setObjectName("header")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.header)
        self.horizontalLayout.setSpacing(9)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.horizontalLayout.setContentsMargins(25,0,15,0)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        self.name = QtWidgets.QLabel(self.header)
        self.name.setFont(font1)
        self.name.setStyleSheet("color: rgb(222, 222, 222)")
        self.name.setTextFormat(QtCore.Qt.MarkdownText)
        self.name.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.name.setIndent(-1)
        self.name.setObjectName("name")
        self.horizontalLayout.addWidget(self.name)
        self.minimize_button = QtWidgets.QPushButton(self.header)
        self.minimize_button.setFixedSize(QtCore.QSize(50, 50))
        self.minimize_button.setStyleSheet("border: none")
        self.minimize_button.setIcon(self.icon1)
        self.minimize_button.setIconSize(QtCore.QSize(20, 20))
        self.minimize_button.setObjectName("minimize_button")
        self.minimize_button.setStyleSheet("QPushButton {\n"
                                          "    border: none;\n"
                                          "}\n"
                                          "QPushButton::hover {\n"
                                          "    background-color: rgb(79, 84, 89)\n"
                                          "}")
        self.horizontalLayout.addWidget(self.minimize_button)
        self.maximize_button = QtWidgets.QPushButton(self.header)
        self.maximize_button.setFixedSize(QtCore.QSize(50, 50))
        self.maximize_button.setStyleSheet("QPushButton {\n"
                                          "    border: none;\n"
                                          "}\n"
                                          "QPushButton::hover {\n"
                                          "    background-color: rgb(79, 84, 89)\n"
                                          "}")        
        self.maximize_button.setIcon(self.icon2)
        self.maximize_button.setIconSize(QtCore.QSize(20, 20))
        self.maximize_button.setObjectName("maximize_button")
        self.horizontalLayout.addWidget(self.maximize_button)
        self.close_button = QtWidgets.QPushButton(self.header)
        self.close_button.setFixedSize(QtCore.QSize(50, 50))
        self.close_button.setStyleSheet("QPushButton {\n"
                                          "    border: none;\n"
                                          "}\n"
                                          "QPushButton::hover {\n"
                                          "    background-color: rgb(79, 84, 89)\n"
                                          "}")
        self.close_button.setIcon(self.icon3)
        self.close_button.setIconSize(QtCore.QSize(20, 20))
        self.close_button.setObjectName("close_button")
        self.horizontalLayout.addWidget(self.close_button)
        self.centralwidget_layout.addWidget(self.header)
        self.body = QtWidgets.QFrame(self.central_widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.body.sizePolicy().hasHeightForWidth())
        self.body.setSizePolicy(sizePolicy)
        self.body.setMinimumSize(QtCore.QSize(0, 0))
        self.body.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.body.setFrameShadow(QtWidgets.QFrame.Plain)
        self.body.setLineWidth(1)
        self.body.setObjectName("body")
        self.body_layout = QtWidgets.QHBoxLayout(self.body)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        self.body_layout.setSpacing(0)
        self.body_layout.setObjectName("body_layout")
        self.left_bar = QtWidgets.QFrame(self.body)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.left_bar.sizePolicy().hasHeightForWidth())
        self.left_bar.setSizePolicy(sizePolicy)
        self.left_bar.setMinimumSize(QtCore.QSize(250, 0))
        self.left_bar.setMaximumSize(QtCore.QSize(400, 16777215))
        self.left_bar.setStyleSheet("background-color: rgb(255, 255, 255)")
        self.left_bar.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.left_bar.setFrameShadow(QtWidgets.QFrame.Plain)
        self.left_bar.setObjectName("left_bar")
        self.left_bar_layout = QtWidgets.QVBoxLayout(self.left_bar)
        self.left_bar_layout.setContentsMargins(0, 0, 0, 0)
        self.left_bar_layout.setSpacing(0)
        self.left_bar_layout.setObjectName("left_bar_layout")
        self.tutorial_tab = QtWidgets.QPushButton(self.left_bar)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.tutorial_tab.sizePolicy().hasHeightForWidth())
        self.tutorial_tab.setSizePolicy(sizePolicy)
        self.tutorial_tab.setMinimumSize(QtCore.QSize(0, 75))
        self.tutorial_tab.setMaximumSize(QtCore.QSize(16777215, 200))
        self.tutorial_tab.setFont(font3)
        self.tutorial_tab.setStyleSheet("border: none; background-color: rgb(246, 246, 246)")
        self.tutorial_tab.setObjectName("tutorial_tab")
        self.left_bar_layout.addWidget(self.tutorial_tab, 0, QtCore.Qt.AlignTop)
        self.upload_tab = QtWidgets.QPushButton(self.left_bar)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.upload_tab.sizePolicy().hasHeightForWidth())
        self.upload_tab.setSizePolicy(sizePolicy)
        self.upload_tab.setMinimumSize(QtCore.QSize(0, 75))
        self.upload_tab.setMaximumSize(QtCore.QSize(16777215, 200))
        self.upload_tab.setFont(font3)
        self.upload_tab.setStyleSheet("border: none")
        self.upload_tab.setObjectName("upload_tab")
        self.left_bar_layout.addWidget(self.upload_tab, 0, QtCore.Qt.AlignTop)
        self.rule_tab = QtWidgets.QPushButton(self.left_bar)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.rule_tab.sizePolicy().hasHeightForWidth())
        self.rule_tab.setSizePolicy(sizePolicy)
        self.rule_tab.setMinimumSize(QtCore.QSize(0, 75))
        self.rule_tab.setMaximumSize(QtCore.QSize(16777215, 200))
        self.rule_tab.setFont(font3)
        self.rule_tab.setStyleSheet("border: none")
        self.rule_tab.setObjectName("rule_tab")
        self.left_bar_layout.addWidget(self.rule_tab, 0, QtCore.Qt.AlignTop)
        self.recon_tab = QtWidgets.QPushButton(self.left_bar)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.recon_tab.sizePolicy().hasHeightForWidth())
        self.recon_tab.setSizePolicy(sizePolicy)
        self.recon_tab.setMinimumSize(QtCore.QSize(0, 75))
        self.recon_tab.setMaximumSize(QtCore.QSize(16777215, 200))
        self.recon_tab.setFont(font3)
        self.recon_tab.setStyleSheet("border: none")
        self.recon_tab.setObjectName("recon_tab")
        self.left_bar_layout.addWidget(self.recon_tab, 0, QtCore.Qt.AlignTop)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.left_bar_layout.addItem(spacerItem)
        self.version_label = QtWidgets.QLabel(self.left_bar)
        self.version_label.setFont(font)
        self.version_label.setAlignment(QtCore.Qt.AlignCenter)
        self.version_label.setObjectName("version_label")
        self.left_bar_layout.addWidget(self.version_label)
        self.body_layout.addWidget(self.left_bar)
        self.screen = QtWidgets.QFrame(self.body)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.screen.sizePolicy().hasHeightForWidth())
        self.screen.setSizePolicy(sizePolicy)
        self.screen.setStyleSheet("background-color: rgb(246, 246, 246)")
        self.screen.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.screen.setFrameShadow(QtWidgets.QFrame.Plain)
        self.screen.setObjectName("screen")
        self.screen_layout = QtWidgets.QHBoxLayout(self.screen)
        self.screen_layout.setObjectName("screen_layout")
        self.stacked_Widget = QtWidgets.QStackedWidget(self.screen)
        self.stacked_Widget.setStyleSheet("")
        self.stacked_Widget.setObjectName("stacked_Widget")
        self.page_1 = QtWidgets.QWidget()
        self.page_1.setObjectName("page_1")
        self.page_1_layout = QtWidgets.QVBoxLayout(self.page_1)
        self.page_1_layout.setContentsMargins(11, 11, 11, 11)
        self.page_1_layout.setSpacing(20)
        self.page_1_layout.setObjectName("page_1_layout")
        self.page_1_header = QtWidgets.QFrame(self.page_1)
        self.page_1_header.setMinimumSize(QtCore.QSize(0, 50))
        self.page_1_header.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_1_header.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_1_header.setObjectName("page_1_header")        
        self.page_1_header_layout = QtWidgets.QVBoxLayout(self.page_1_header)
        self.page_1_header_layout.setContentsMargins(0, 0, 0, 0)
        self.page_1_header_layout.setObjectName("page_1_header_layout")
        self.page_1_header_text = QtWidgets.QLabel(self.page_1_header)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.page_1_header_text.sizePolicy().hasHeightForWidth())
        self.page_1_header_text.setSizePolicy(sizePolicy)
        self.page_1_header_text.setMinimumSize(QtCore.QSize(0, 0))
        self.page_1_header_text.setFont(font2)
        self.page_1_header_text.setStyleSheet("")
        self.page_1_header_text.setTextFormat(QtCore.Qt.AutoText)
        self.page_1_header_text.setAlignment(QtCore.Qt.AlignCenter)
        self.page_1_header_text.setWordWrap(True)
        self.page_1_header_text.setObjectName("page_1_header_text")
        self.page_1_header_layout.addWidget(self.page_1_header_text)
        self.page_1_layout.addWidget(self.page_1_header)
        self.page_1_body = QtWidgets.QFrame(self.page_1)       
        self.page_1_body.setMinimumSize(QtCore.QSize(0, 300))
        self.page_1_body.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_1_body.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_1_body.setLineWidth(0)
        self.page_1_body.setObjectName("page_1_body")
        self.page_1_body_layout = QtWidgets.QVBoxLayout(self.page_1_body)
        self.page_1_body_layout.setContentsMargins(0, 0, 0, 0)
        self.page_1_body_layout.setSpacing(0)
        self.page_1_body_layout.setObjectName("page_1_body_layout")
        self.page_1_scrollArea = QtWidgets.QScrollArea(self.page_1_body)
        self.page_1_scrollArea.setFont(font)
        self.page_1_scrollArea.setAcceptDrops(False)
        self.page_1_scrollArea.setStyleSheet("QScrollBar:vertical {\n"
                                            "border:none;\n"
                                            "background: rgb(49, 54, 59)"
                                            "}")
        self.page_1_scrollArea.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.page_1_scrollArea.setFrameShadow(QtWidgets.QFrame.Plain)
        self.page_1_scrollArea.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.page_1_scrollArea.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.page_1_scrollArea.setWidgetResizable(True)
        self.page_1_scrollArea.setObjectName("page_1_scrollArea")
        self.page_1_scrollAreaWidgetContents = QtWidgets.QWidget()
        self.page_1_scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 905, 1210))
        self.page_1_scrollAreaWidgetContents.setObjectName("page_1_scrollAreaWidgetContents")
        self.verticalLayout_18 = QtWidgets.QVBoxLayout(self.page_1_scrollAreaWidgetContents)
        self.verticalLayout_18.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_18.setObjectName("verticalLayout_18")
        self.page_1_body_text = QtWidgets.QLabel(self.page_1_scrollAreaWidgetContents)
        self.page_1_body_text.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.page_1_body_text.sizePolicy().hasHeightForWidth())
        self.page_1_body_text.setSizePolicy(sizePolicy)
        self.page_1_body_text.setFont(font4)
        self.page_1_body_text.setContentsMargins(5,5,5,5)
        self.page_1_body_text.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.page_1_body_text.setFrameShadow(QtWidgets.QFrame.Plain)
        self.page_1_body_text.setText("")
        self.page_1_body_text.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignTop)
        self.page_1_body_text.setWordWrap(True)
        self.page_1_body_text.setIndent(0)
        self.page_1_body_text.setObjectName("page_1_body_text")
        self.verticalLayout_18.addWidget(self.page_1_body_text)
        self.page_1_scrollArea.setWidget(self.page_1_scrollAreaWidgetContents)
        self.page_1_body_layout.addWidget(self.page_1_scrollArea)
        self.page_1_layout.addWidget(self.page_1_body)
        self.page_1_next = QtWidgets.QFrame(self.page_1)
        self.page_1_next.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_1_next.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_1_next.setObjectName("page_1_next")
        self.page_1_next_layout = QtWidgets.QVBoxLayout(self.page_1_next)
        self.page_1_next_layout.setObjectName("page_1_next_layout")
        self.page_1_next_button = QtWidgets.QPushButton(self.page_1_next)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.page_1_next_button.sizePolicy().hasHeightForWidth())
        self.page_1_next_button.setSizePolicy(sizePolicy)
        self.page_1_next_button.setMinimumSize(QtCore.QSize(90, 35))
        self.page_1_next_button.setFont(font3)
        self.page_1_next_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.page_1_next_button.setLocale(QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.UnitedStates))
        self.page_1_next_button.setIcon(self.icon4)
        self.page_1_next_button.setIconSize(QtCore.QSize(15, 15))
        self.page_1_next_button.setObjectName("page_1_next_button")
        self.page_1_next_layout.addWidget(self.page_1_next_button, 0, QtCore.Qt.AlignRight)
        self.page_1_layout.addWidget(self.page_1_next)
        self.stacked_Widget.addWidget(self.page_1)
        self.page_2 = QtWidgets.QWidget()
        self.page_2.setObjectName("page_2")
        self.page_2_layout = QtWidgets.QVBoxLayout(self.page_2)
        self.page_2_layout.setContentsMargins(11, 11, 11, 11)
        self.page_2_layout.setSpacing(20)
        self.page_2_layout.setObjectName("page_2_layout")
        self.page_2_header = QtWidgets.QFrame(self.page_2)
        self.page_2_header.setMinimumSize(QtCore.QSize(0, 50))
        self.page_2_header.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_2_header.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_2_header.setObjectName("page_2_header")
        self.page_2_header_layout = QtWidgets.QVBoxLayout(self.page_2_header)
        self.page_2_header_layout.setContentsMargins(0, 0, 0, 0)
        self.page_2_header_layout.setSpacing(7)
        self.page_2_header_layout.setObjectName("page_2_header_layout")
        self.page_2_header_text = QtWidgets.QLabel(self.page_2_header)
        self.page_2_header_text.setMinimumSize(QtCore.QSize(0, 50))
        self.page_2_header_text.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.page_2_header_text.setFont(font2)
        self.page_2_header_text.setAlignment(QtCore.Qt.AlignCenter)
        self.page_2_header_text.setWordWrap(True)
        self.page_2_header_text.setObjectName("page_2_header_text")
        self.page_2_header_layout.addWidget(self.page_2_header_text, 0, QtCore.Qt.AlignHCenter|QtCore.Qt.AlignTop)
        self.page_2_layout.addWidget(self.page_2_header, 0, QtCore.Qt.AlignTop)
        self.page_2_upload_files = QtWidgets.QFrame(self.page_2)
        self.page_2_upload_files.setStyleSheet("background-color: rgb(237, 237, 237); border-radius: 7px;")
        self.page_2_upload_files.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_2_upload_files.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_2_upload_files.setObjectName("page_2_upload_files")
        self.page_2_upload_files_2 = QtWidgets.QVBoxLayout(self.page_2_upload_files)
        self.page_2_upload_files_2.setObjectName("page_2_upload_files_2")
        self.page_2_body = QtWidgets.QLabel(self.page_2_upload_files)
        self.page_2_body.setFont(font4)
        self.page_2_body.setObjectName("page_2_body")
        self.page_2_upload_files_2.addWidget(self.page_2_body, 0, QtCore.Qt.AlignHCenter)
        self.upload = QtWidgets.QFrame(self.page_2_upload_files)
        self.upload.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.upload.setFrameShadow(QtWidgets.QFrame.Raised)
        self.upload.setObjectName("upload")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.upload)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setSpacing(0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.source = QtWidgets.QFrame(self.upload)
        self.source.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.source.setFrameShadow(QtWidgets.QFrame.Raised)
        self.source.setObjectName("source")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.source)
        self.verticalLayout_4.setContentsMargins(100, -1, 100, -1)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.upload_source_button = QtWidgets.QPushButton(self.source)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.upload_source_button.sizePolicy().hasHeightForWidth())
        self.upload_source_button.setSizePolicy(sizePolicy)
        self.upload_source_button.setMinimumSize(QtCore.QSize(300, 35))
        self.upload_source_button.setFont(font3)
        self.upload_source_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.upload_source_button.setIcon(self.icon5)
        self.upload_source_button.setObjectName("upload_source_button")
        self.verticalLayout_4.addWidget(self.upload_source_button, 0, QtCore.Qt.AlignHCenter)
        self.source_file_name = QtWidgets.QLabel(self.source)
        self.source_file_name.setEnabled(False)
        self.source_file_name.setMaximumSize(QtCore.QSize(230, 100))
        self.source_file_name.setText("")
        self.source_file_name.setWordWrap(False)
        self.source_file_name.setObjectName("source_file_name")
        self.verticalLayout_4.addWidget(self.source_file_name, 0, QtCore.Qt.AlignHCenter)
        self.horizontalLayout_2.addWidget(self.source)
        self.target = QtWidgets.QFrame(self.upload)
        self.target.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.target.setFrameShadow(QtWidgets.QFrame.Raised)
        self.target.setObjectName("target")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.target)
        self.verticalLayout_5.setContentsMargins(100, -1, 100, -1)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.upload_target_button = QtWidgets.QPushButton(self.target)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.upload_source_button.sizePolicy().hasHeightForWidth())
        self.upload_target_button.setSizePolicy(sizePolicy)        
        self.upload_target_button.setMinimumSize(QtCore.QSize(300, 35))
        self.upload_target_button.setFont(font3)
        self.upload_target_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.upload_target_button.setIcon(self.icon5)
        self.upload_target_button.setObjectName("upload_target_button")
        self.verticalLayout_5.addWidget(self.upload_target_button, 0, QtCore.Qt.AlignHCenter)
        self.target_file_name = QtWidgets.QLabel(self.target)
        self.target_file_name.setEnabled(False)
        self.target_file_name.setMaximumSize(QtCore.QSize(230, 100))
        self.target_file_name.setText("")
        self.target_file_name.setWordWrap(False)
        self.target_file_name.setObjectName("target_file_name")
        self.verticalLayout_5.addWidget(self.target_file_name, 0, QtCore.Qt.AlignHCenter)
        self.horizontalLayout_2.addWidget(self.target)
        self.page_2_upload_files_2.addWidget(self.upload)
        self.page_2_layout.addWidget(self.page_2_upload_files)
        self.page_2_prevalidation = QtWidgets.QFrame(self.page_2)
        self.page_2_prevalidation.setStyleSheet("background-color: rgb(237, 237, 237); border-radius: 7px;")
        self.page_2_prevalidation.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_2_prevalidation.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_2_prevalidation.setObjectName("page_2_prevalidation")
        self.page_2_prevalidation_layout = QtWidgets.QVBoxLayout(self.page_2_prevalidation)
        self.page_2_prevalidation_layout.setContentsMargins(100, 11, 100, -1)
        self.page_2_prevalidation_layout.setObjectName("page_2_prevalidation_layout")
        self.page_2_body_2 = QtWidgets.QLabel(self.page_2_prevalidation)
        self.page_2_body_2.setFont(font4)
        self.page_2_body_2.setLineWidth(0)
        self.page_2_body_2.setObjectName("page_2_body_2")
        self.page_2_prevalidation_layout.addWidget(self.page_2_body_2, 0, QtCore.Qt.AlignHCenter)
        self.prevalidate_button = QtWidgets.QPushButton(self.page_2_prevalidation)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.prevalidate_button.sizePolicy().hasHeightForWidth())
        self.prevalidate_button.setSizePolicy(sizePolicy)
        self.prevalidate_button.setMinimumSize(QtCore.QSize(300, 35))
        self.prevalidate_button.setFont(font3)
        self.prevalidate_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.prevalidate_button.setIcon(self.icon6)
        self.prevalidate_button.setObjectName("prevalidate_button")
        self.page_2_prevalidation_layout.addWidget(self.prevalidate_button, 0, QtCore.Qt.AlignHCenter)
        self.page_2_layout.addWidget(self.page_2_prevalidation)
        self.page_2_output = QtWidgets.QFrame(self.page_2)
        self.page_2_output.setMinimumSize(QtCore.QSize(0, 200))
        self.page_2_output.setStyleSheet("background-color: rgb(237, 237, 237); border-radius: 7px;")
        self.page_2_output.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_2_output.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_2_output.setObjectName("page_2_output")
        self.page_2_output_layout = QtWidgets.QVBoxLayout(self.page_2_output)
        self.page_2_output_layout.setSpacing(7)
        self.page_2_output_layout.setObjectName("page_2_output_layout")
        self.page_2_output_label = QtWidgets.QLabel(self.page_2_output)
        self.page_2_output_label.setFont(font4)
        self.page_2_output_label.setAlignment(QtCore.Qt.AlignCenter)
        self.page_2_output_label.setObjectName("page_2_output_label")
        self.page_2_output_layout.addWidget(self.page_2_output_label)
        self.page_2_scrollArea = QtWidgets.QScrollArea(self.page_1_body)
        self.page_2_scrollArea.setFont(font)
        self.page_2_scrollArea.setAcceptDrops(False)
        self.page_2_scrollArea.setStyleSheet("QScrollBar:vertical {\n"
                                            "border:none;\n"
                                            "background: rgb(49, 54, 59)"
                                            "}")
        self.page_2_scrollArea.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.page_2_scrollArea.setFrameShadow(QtWidgets.QFrame.Plain)
        self.page_2_scrollArea.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.page_2_scrollArea.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.page_2_scrollArea.setWidgetResizable(True)
        self.page_2_scrollArea.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self.page_2_scrollArea.setObjectName("page_2_scrollArea")
        self.page_2_output_layout.addWidget(self.page_2_scrollArea)
        self.page_2_output_list = QtWidgets.QLabel(self.page_2_output)
        self.page_2_output_list.setFont(font3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.page_2_output_list.sizePolicy().hasHeightForWidth())
        self.page_2_output_list.setSizePolicy(sizePolicy)
        self.page_2_output_list.setStyleSheet("background-color:white")
        self.page_2_output_list.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.page_2_output_list.setWordWrap(True)
        self.page_2_output_list.setObjectName("page_2_text_browser")
        self.page_2_output_list.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self.page_2_scrollArea.setWidget(self.page_2_output_list)
        self.page_2_layout.addWidget(self.page_2_output)
        self.page_2_next = QtWidgets.QFrame(self.page_2)
        self.page_2_next.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_2_next.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_2_next.setObjectName("page_2_next")
        self.page_2_next_layout = QtWidgets.QHBoxLayout(self.page_2_next)
        self.page_2_next_layout.setObjectName("page_2_next_layout")
        self.page_2_back_button = QtWidgets.QPushButton(self.page_2_next)
        self.page_2_back_button.setMinimumSize(QtCore.QSize(90, 35))
        self.page_2_back_button.setFont(font3)
        self.page_2_back_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")        
        self.page_2_back_button.setIcon(self.icon10)
        self.page_2_back_button.setIconSize(QtCore.QSize(15, 15))
        self.page_2_back_button.setObjectName("page_2_back_button")
        self.page_2_next_layout.addWidget(self.page_2_back_button, 0, QtCore.Qt.AlignLeft)
        self.page_2_next_button = QtWidgets.QPushButton(self.page_2_next)
        self.page_2_next_button.setMinimumSize(QtCore.QSize(90, 35))
        self.page_2_next_button.setFont(font3)
        self.page_2_next_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")        
        self.page_2_next_button.setIcon(self.icon4)
        self.page_2_next_button.setIconSize(QtCore.QSize(15, 15))
        self.page_2_next_button.setObjectName("page_2_next_button")
        self.page_2_next_layout.addWidget(self.page_2_next_button, 0, QtCore.Qt.AlignRight)
        self.page_2_layout.addWidget(self.page_2_next)
        self.stacked_Widget.addWidget(self.page_2)
        self.page_3 = QtWidgets.QWidget()
        self.page_3.setObjectName("page_3")
        self.page_3_layout = QtWidgets.QVBoxLayout(self.page_3)
        self.page_3_layout.setContentsMargins(11, 11, 11, 11)
        self.page_3_layout.setSpacing(20)
        self.page_3_layout.setObjectName("page_3_layout")
        self.page_3_header = QtWidgets.QFrame(self.page_3)
        self.page_3_header.setMinimumSize(QtCore.QSize(0, 50))
        self.page_3_header.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_3_header.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_3_header.setObjectName("page_3_header")
        self.page_3_header_layout = QtWidgets.QVBoxLayout(self.page_3_header)
        self.page_3_header_layout.setContentsMargins(0, 0, 0, 0)
        self.page_3_header_layout.setSpacing(7)
        self.page_3_header_layout.setObjectName("page_3_header_layout")
        self.page_3_header_text = QtWidgets.QLabel(self.page_3_header)
        self.page_3_header_text.setFont(font2)
        self.page_3_header_text.setAlignment(QtCore.Qt.AlignCenter)
        self.page_3_header_text.setWordWrap(True)
        self.page_3_header_text.setObjectName("page_3_header_text")
        self.page_3_header_layout.addWidget(self.page_3_header_text)
        self.page_3_layout.addWidget(self.page_3_header)
        self.page_3_intro = QtWidgets.QFrame(self.page_3)
        self.page_3_intro.setStyleSheet("background-color: rgb(237, 237, 237); border-radius: 7px;")
        self.page_3_intro.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_3_intro.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_3_intro.setObjectName("page_3_intro")
        self.page_3_intro_layout = QtWidgets.QVBoxLayout(self.page_3_intro)
        self.page_3_intro_layout.setObjectName("page_3_intro_layout")
        self.page_3_intro_text = QtWidgets.QLabel(self.page_3_intro)
        self.page_3_intro_text.setFont(font4)
        self.page_3_intro_text.setObjectName("page_3_intro_text")
        self.page_3_intro_layout.addWidget(self.page_3_intro_text)
        self.page_3_layout.addWidget(self.page_3_intro)
        self.page_3_dropdown = QtWidgets.QFrame(self.page_3)
        self.page_3_dropdown.setStyleSheet("background-color: rgb(237, 237, 237); border-radius: 7px;")
        self.page_3_dropdown.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_3_dropdown.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_3_dropdown.setObjectName("page_3_dropdown")
        self.page_3_dropdown_layout = QtWidgets.QHBoxLayout(self.page_3_dropdown)
        self.page_3_dropdown_layout.setObjectName("page_3_dropdown_layout")
        self.source_column = QtWidgets.QFrame(self.page_3_dropdown)
        self.source_column.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.source_column.setFrameShadow(QtWidgets.QFrame.Raised)
        self.source_column.setObjectName("source_column")
        self.source_column_layout = QtWidgets.QVBoxLayout(self.source_column)
        self.source_column_layout.setObjectName("source_column_layout")
        self.source_column_label = QtWidgets.QLabel(self.source_column)
        self.source_column_label.setFont(font4)
        self.source_column_label.setObjectName("source_column_label")
        self.source_column_layout.addWidget(self.source_column_label)
        self.source_column_dropdown = QtWidgets.QComboBox(self.source_column)
        self.source_column_dropdown.setStyleSheet("background-color: white; border-radius: 5px")
        self.source_column_dropdown.setEditable(False)
        self.source_column_dropdown.setFont(font4)
        self.source_column_dropdown.setObjectName("source_column_dropdown")
        self.source_column_layout.addWidget(self.source_column_dropdown)
        self.page_3_dropdown_layout.addWidget(self.source_column)
        self.operator_column = QtWidgets.QFrame(self.page_3_dropdown)
        self.operator_column.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.operator_column.setFrameShadow(QtWidgets.QFrame.Raised)
        self.operator_column.setObjectName("operator_column")
        self.operator_column_layout = QtWidgets.QVBoxLayout(self.operator_column)
        self.operator_column_layout.setObjectName("operator_column_layout")
        self.operator_label = QtWidgets.QLabel(self.operator_column)
        self.operator_label.setFont(font4)
        self.operator_label.setObjectName("operator_label")
        self.operator_column_layout.addWidget(self.operator_label)
        self.operator_dropdown = QtWidgets.QComboBox(self.operator_column)
        self.operator_dropdown.setStyleSheet("background-color: white; border-radius: 5px")
        self.operator_dropdown.setEditable(False)
        self.operator_dropdown.setFont(font4)
        self.operator_dropdown.setObjectName("operator_dropdown")
        self.operator_column_layout.addWidget(self.operator_dropdown)
        self.page_3_dropdown_layout.addWidget(self.operator_column)
        self.target_column = QtWidgets.QFrame(self.page_3_dropdown)
        self.target_column.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.target_column.setFrameShadow(QtWidgets.QFrame.Raised)
        self.target_column.setObjectName("target_column")
        self.target_column_layout = QtWidgets.QVBoxLayout(self.target_column)
        self.target_column_layout.setObjectName("target_column_layout")
        self.target_column_label = QtWidgets.QLabel(self.target_column)
        self.target_column_label.setFont(font4)
        self.target_column_label.setObjectName("target_column_label")
        self.target_column_layout.addWidget(self.target_column_label)
        self.target_column_dropdown = QtWidgets.QComboBox(self.target_column)
        self.target_column_dropdown.setStyleSheet("background-color: white; border-radius: 5px")
        self.target_column_dropdown.setEditable(False)
        self.target_column_dropdown.setFont(font4)
        self.target_column_dropdown.setObjectName("target_column_dropdown")
        self.target_column_layout.addWidget(self.target_column_dropdown)
        self.page_3_dropdown_layout.addWidget(self.target_column)
        self.page_3_layout.addWidget(self.page_3_dropdown)
        self.page_3_rule_output_viewer = QtWidgets.QFrame(self.page_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.page_3_rule_output_viewer.sizePolicy().hasHeightForWidth())
        self.page_3_rule_output_viewer.setSizePolicy(sizePolicy)
        self.page_3_rule_output_viewer.setMinimumSize(QtCore.QSize(0, 50))
        self.page_3_rule_output_viewer.setStyleSheet("background-color: rgb(237, 237, 237); border-radius: 7px;")
        self.page_3_rule_output_viewer.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_3_rule_output_viewer.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_3_rule_output_viewer.setFont(font4)
        self.page_3_rule_output_viewer.setObjectName("page_3_rule_output_viewer")
        self.page_3_rule_output_viewer_layout = QtWidgets.QVBoxLayout(self.page_3_rule_output_viewer)
        self.page_3_rule_output_viewer_layout.setContentsMargins(11, 11, 11, 11)
        self.page_3_rule_output_viewer_layout.setSpacing(7)
        self.page_3_rule_output_viewer_layout.setObjectName("page_3_rule_output_viewer_layout")
        self.rule_viewer = QtWidgets.QFrame(self.page_3_rule_output_viewer)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.rule_viewer.sizePolicy().hasHeightForWidth())
        self.rule_viewer.setSizePolicy(sizePolicy)
        self.rule_viewer.setMinimumSize(QtCore.QSize(0, 100))
        self.rule_viewer.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.rule_viewer.setFrameShadow(QtWidgets.QFrame.Raised)
        self.rule_viewer.setObjectName("rule_viewer")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.rule_viewer)
        self.horizontalLayout_3.setContentsMargins(11, 11, -1, -1)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.construct_rule_label = QtWidgets.QLabel(self.rule_viewer)
        self.construct_rule_label.setFont(font4)
        self.construct_rule_label.setObjectName("construct_rule_label")
        self.horizontalLayout_3.addWidget(self.construct_rule_label)
        self.construct_rule_browser = QtWidgets.QTextEdit(self.rule_viewer)
        self.construct_rule_browser.setMaximumSize(QtCore.QSize(16777215, 30))
        self.construct_rule_browser.setStyleSheet("background-color: white; border-radius: 5px")
        self.construct_rule_browser.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.construct_rule_browser.setFont(font4)
        self.construct_rule_browser.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.construct_rule_browser.setObjectName("construct_rule_browser")
        self.horizontalLayout_3.addWidget(self.construct_rule_browser)
        self.clear_rule_construct_button = QtWidgets.QPushButton(self.rule_viewer)
        self.clear_rule_construct_button.setMinimumSize(QtCore.QSize(90, 35))
        self.clear_rule_construct_button.setBaseSize(QtCore.QSize(0, 0))
        self.clear_rule_construct_button.setFont(font3)
        self.clear_rule_construct_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.clear_rule_construct_button.setIcon(self.icon7)
        self.clear_rule_construct_button.setObjectName("clear_rule_construct_button")
        self.horizontalLayout_3.addWidget(self.clear_rule_construct_button)
        self.validate_button = QtWidgets.QPushButton(self.rule_viewer)
        self.validate_button.setMinimumSize(QtCore.QSize(90, 35))
        self.validate_button.setFont(font3)
        self.validate_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.validate_button.setIcon(self.icon6)
        self.validate_button.setObjectName("validate_button")
        self.horizontalLayout_3.addWidget(self.validate_button)
        self.page_3_rule_output_viewer_layout.addWidget(self.rule_viewer)
        self.page_3_layout.addWidget(self.page_3_rule_output_viewer)
        self.page_3_validated_rule_list = QtWidgets.QFrame(self.page_3)
        self.page_3_validated_rule_list.setStyleSheet("background-color: rgb(237, 237, 237); border-radius: 7px;")
        self.page_3_validated_rule_list.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_3_validated_rule_list.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_3_validated_rule_list.setObjectName("page_3_validated_rule_list")
        self.page_3_validated_rule_list_layout = QtWidgets.QHBoxLayout(self.page_3_validated_rule_list)
        self.page_3_validated_rule_list_layout.setObjectName("page_3_validated_rule_list_layout")
        self.label = QtWidgets.QLabel(self.page_3_validated_rule_list)
        self.label.setFont(font4)
        self.label.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label.setObjectName("label")
        self.page_3_validated_rule_list_layout.addWidget(self.label)
        self.listWidget = QtWidgets.QListWidget(self.page_3_validated_rule_list)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.listWidget.sizePolicy().hasHeightForWidth())
        self.listWidget.setSizePolicy(sizePolicy)
        self.listWidget.setStyleSheet("background-color: white; border-radius: 5px")
        self.listWidget.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.listWidget.setWordWrap(True)
        self.listWidget.setObjectName("listWidget")
        self.page_3_validated_rule_list_layout.addWidget(self.listWidget)
        self.delete_upload_rule = QtWidgets.QFrame(self.page_3_validated_rule_list)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.delete_upload_rule.sizePolicy().hasHeightForWidth())
        self.delete_upload_rule.setSizePolicy(sizePolicy)
        self.delete_upload_rule.setMinimumSize(QtCore.QSize(50, 0))
        self.delete_upload_rule.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.delete_upload_rule.setFrameShadow(QtWidgets.QFrame.Raised)
        self.delete_upload_rule.setObjectName("delete_upload_rule")
        self.verticalLayout_10 = QtWidgets.QVBoxLayout(self.delete_upload_rule)
        self.verticalLayout_10.setObjectName("verticalLayout_10")
        self.upload_rules_button = QtWidgets.QPushButton(self.delete_upload_rule)
        self.upload_rules_button.setMinimumSize(QtCore.QSize(90, 35))
        self.upload_rules_button.setFont(font3)
        self.upload_rules_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.upload_rules_button.setIcon(self.icon5)
        self.upload_rules_button.setObjectName("upload_rules_button")
        self.verticalLayout_10.addWidget(self.upload_rules_button)
        self.delete_rule_button = QtWidgets.QPushButton(self.delete_upload_rule)
        self.delete_rule_button.setMinimumSize(QtCore.QSize(90, 35))
        self.delete_rule_button.setFont(font3)
        self.delete_rule_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.delete_rule_button.setIcon(self.icon7)
        self.delete_rule_button.setObjectName("delete_rule_button")
        self.verticalLayout_10.addWidget(self.delete_rule_button)
        self.page_3_validated_rule_list_layout.addWidget(self.delete_upload_rule)
        self.page_3_layout.addWidget(self.page_3_validated_rule_list)
        self.page_3_next = QtWidgets.QFrame(self.page_3)
        self.page_3_next.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_3_next.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_3_next.setObjectName("page_3_next")
        self.page_3_next_layout = QtWidgets.QHBoxLayout(self.page_3_next)
        self.page_3_next_layout.setObjectName("page_3_next_layout")
        self.page_3_back_button = QtWidgets.QPushButton(self.page_3_next)
        self.page_3_back_button.setMinimumSize(QtCore.QSize(90, 35))
        self.page_3_back_button.setFont(font3)
        self.page_3_back_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.page_3_back_button.setIcon(self.icon10)
        self.page_3_back_button.setIconSize(QtCore.QSize(15, 15))
        self.page_3_back_button.setObjectName("page_3_next_button")
        self.page_3_next_layout.addWidget(self.page_3_back_button, 0, QtCore.Qt.AlignLeft)
        self.reconcile_button = QtWidgets.QPushButton(self.page_3_next)
        self.reconcile_button.setMinimumSize(QtCore.QSize(90, 35))
        self.reconcile_button.setFont(font3)
        self.reconcile_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.reconcile_button.setIcon(self.icon4)
        self.reconcile_button.setIconSize(QtCore.QSize(15, 15))
        self.reconcile_button.setObjectName("page_3_next_button")
        self.page_3_next_layout.addWidget(self.reconcile_button, 0, QtCore.Qt.AlignRight)
        self.page_3_layout.addWidget(self.page_3_next)
        self.stacked_Widget.addWidget(self.page_3)
        self.page_4 = QtWidgets.QWidget()
        self.page_4.setObjectName("page_4")
        self.page_4_layout = QtWidgets.QVBoxLayout(self.page_4)
        self.page_4_layout.setContentsMargins(11, 11, 11, 11)
        self.page_4_layout.setSpacing(20)
        self.page_4_layout.setObjectName("page_4_layout")
        self.page_4_header = QtWidgets.QFrame(self.page_4)
        self.page_4_header.setMinimumSize(QtCore.QSize(0, 50))
        self.page_4_header.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_4_header.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_4_header.setObjectName("page_4_header")
        self.page_4_header_layout = QtWidgets.QVBoxLayout(self.page_4_header)
        self.page_4_header_layout.setContentsMargins(0, 0, 0, 0)
        self.page_4_header_layout.setSpacing(0)
        self.page_4_header_layout.setObjectName("page_4_header_layout")
        self.page_4_header_text = QtWidgets.QLabel(self.page_4_header)
        self.page_4_header_text.setMinimumSize(QtCore.QSize(0, 50))
        self.page_4_header_text.setFont(font2)
        self.page_4_header_text.setAlignment(QtCore.Qt.AlignCenter)
        self.page_4_header_text.setObjectName("page_4_header_text")
        self.page_4_header_layout.addWidget(self.page_4_header_text)
        self.page_4_layout.addWidget(self.page_4_header)
        self.page_4_loading = QtWidgets.QFrame(self.page_4)
        self.page_4_loading.setMinimumSize(QtCore.QSize(0, 0))
        self.page_4_loading.setStyleSheet("background-color: rgb(237, 237, 237);border-radius: 7px;")
        self.page_4_loading.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_4_loading.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_4_loading.setObjectName("page_4_loading")
        self.page_4_loading_layout = QtWidgets.QVBoxLayout(self.page_4_loading)
        self.page_4_loading_layout.setObjectName("page_4_loading_layout")
        self.progress_bar_text = QtWidgets.QLabel(self.page_4_loading)
        self.progress_bar_text.setFont(font3)
        self.progress_bar_text.setAlignment(QtCore.Qt.AlignCenter)
        self.progress_bar_text.setObjectName("progress_bar_text")
        self.page_4_loading_layout.addWidget(self.progress_bar_text)
        self.progressBar = QtWidgets.QProgressBar(self.page_4_loading)
        self.progressBar.setProperty("value", 24)
        self.progressBar.setTextVisible(False)
        self.progressBar.setOrientation(QtCore.Qt.Horizontal)
        self.progressBar.setObjectName("progressBar")
        self.page_4_loading_layout.addWidget(self.progressBar)
        self.page_4_layout.addWidget(self.page_4_loading)
        self.page_4_list = QtWidgets.QFrame(self.page_4)
        self.page_4_list.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_4_list.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_4_list.setObjectName("page_4_list")
        self.page_4_list_layout = QtWidgets.QVBoxLayout(self.page_4_list)
        self.page_4_list_layout.setContentsMargins(0, 0, 0, 0)
        self.page_4_list_layout.setSpacing(0)
        self.page_4_list_layout.setObjectName("page_4_list_layout")
        self.listView = QtWidgets.QLabel(self.page_4_list)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.left_bar.sizePolicy().hasHeightForWidth())
        self.listView.setSizePolicy(sizePolicy)
        self.listView.setFont(font3)
        self.listView.setStyleSheet("background-color: white;border-radius: 7px")
        self.listView.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self.listView.setWordWrap(True)
        self.listView.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.listView.setFrameShadow(QtWidgets.QFrame.Plain)
        self.listView.setObjectName("listView")
        self.page_4_list_layout.addWidget(self.listView)
        self.page_4_layout.addWidget(self.page_4_list)
        self.page_4_restart = QtWidgets.QFrame(self.page_4)
        self.page_4_restart.setMinimumSize(QtCore.QSize(0, 0))
        self.page_4_restart.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.page_4_restart.setFrameShadow(QtWidgets.QFrame.Raised)
        self.page_4_restart.setObjectName("page_4_restart")
        self.page_4_restart_layout = QtWidgets.QHBoxLayout(self.page_4_restart)
        self.page_4_restart_layout.setContentsMargins(11, 11, 11, 11)
        self.page_4_restart_layout.setSpacing(7)
        self.page_4_restart_layout.setObjectName("page_4_restart_layout")
        self.checkBox = QtWidgets.QCheckBox(self.page_4_restart)
        self.checkBox.setFont(font5)
        self.checkBox.setChecked(True)
        self.checkBox.setObjectName("checkBox")
        self.page_4_restart_layout.addWidget(self.checkBox)
        self.checkBox_2 = QtWidgets.QCheckBox(self.page_4_restart)
        self.checkBox_2.setFont(font5)
        self.checkBox_2.setChecked(True)
        self.checkBox_2.setObjectName("checkBox_2")
        self.page_4_restart_layout.addWidget(self.checkBox_2)
        self.restart_button = QtWidgets.QPushButton(self.page_4_restart)
        self.restart_button.setMinimumSize(QtCore.QSize(90, 35))
        self.restart_button.setFont(font3)
        self.restart_button.setStyleSheet("QPushButton {border-radius: 5px; background-color: rgb(49, 54, 59); color: white}\n"
                                                "QPushButton::Hover {background-color: rgb(79, 84, 89)}")
        self.restart_button.setIcon(self.icon8)
        self.restart_button.setObjectName("restart_button")
        self.page_4_restart_layout.addWidget(self.restart_button, 0, QtCore.Qt.AlignRight)
        self.page_4_layout.addWidget(self.page_4_restart, 0, QtCore.Qt.AlignRight)
        self.stacked_Widget.addWidget(self.page_4)
        self.screen_layout.addWidget(self.stacked_Widget)
        self.body_layout.addWidget(self.screen)
        self.centralwidget_layout.addWidget(self.body)
        MainWindow.setCentralWidget(self.central_widget)

        self.retranslateUi(MainWindow)
        self.stacked_Widget.setCurrentIndex(0)
        self.operator_dropdown.setCurrentIndex(-1)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        ### 2ND PART TO GUI INITIALIZATION ###
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowIcon(QtGui.QIcon(r"icons\mercedes-benz.png"))
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.name.setText(_translate("MainWindow", "AutoRecon"))
        self.tutorial_tab.setText(_translate("MainWindow", "Read Me"))
        self.upload_tab.setText(_translate("MainWindow", "Upload Files"))
        self.rule_tab.setText(_translate("MainWindow", "Construct Rules"))
        self.recon_tab.setText(_translate("MainWindow", "Reconciliation"))
        self.version_label.setText(_translate("MainWindow", f"")) # CAN PUT VERSION NO E.G Version 1.0
        self.page_1_header_text.setText(_translate("MainWindow", "Read Me"))
        self.page_1_body_text.setText(_translate("MainWindow", "Take note of the pointers below before proceeding: \n\n"
                                                "1. Ensure your column names are only in alphanumeric and underscore, while date-related columns are read as 'Date' in excel. However, the output for date columns will be in 'YYYY-MM-DD' format for now.\n\n" 
                                                "2. The UI is resizable by clicking on the maximize button.\n\n"
                                                "3. Upload Source and Target excel files by clicking the respective buttons. Ensure that there are no duplicate column names present in each file and that both files should not be empty.\n\n"
                                                "4. Do note that column names may be amended to a more appropriate format (such as removing leading spaces and replacing '.' with '_'). Leading & Trailing spaces in each cell will be removed as well.\n\n"
                                                "5. You may encounter a scenario / error where the tool says similar terms such as 'Unnamed:_16', this means that the 17th column (in this case) is included even though it is blank. Please delete the column for fix.\n\n"
                                                "6. Once both files are uploaded, click on the 'Pre-Validation' button to validate the datasets for errors and warnings. Errors will be shown in console & you cannot proceed to the next step. Warnings will be stored into an excel file & you can be proceed upon acknowledging the message.\n\n"
                                                "7. Configuring pre-validation warning is flexible and users can add new SQL statement(s) into file 'prevalidation.sql' found in 'SQL' folder for program to execute.\n\n"
                                                "8. To construct the reconciliation rules, please select the columns & operators found in the dropdowns.\n\n"
                                                "9. After setting a desired rule, click on the 'Validate Rule' button to ensure that selected rule is error-free, which will be reflected below. In case of error, a popup will list down the reason(s).\n\n"
                                                "10. You can remove unwanted validated rule in the list by selecting the rule & clicking on the 'Delete' button.\n\n"
                                                "11. Click 'Reconcile' button to generate a reconciled excel file with the output, source, target and rules data.\n\n"
                                                "12. Click 'Restart' button after reconciliation output file has been generated to start the process all over again.\n\n"
                                                "13. In order to re-use the rules applied in the previous reconciled output file, keep the 'Rules' tab located in output file (and delete the remaining worksheets) as a single sheet. Then, in the next iteration, click on 'Rules' button to upload it.\n\n"
                                                "14. For better experience, please ensure that your Window Display Scaling is 125%.\n\n"
                                                "15. Since the first column for both excel files are the data drivers / primary keys, there is a possibility of cross join in the event these fields are not unique.\n\n"
                                                "16. Remove any blank column(s) or column(s) without any data."))
        self.page_1_next_button.setText(_translate("MainWindow", "Next"))
        self.page_2_header_text.setText(_translate("MainWindow", "UPLOAD FILES"))
        self.page_2_body.setText(_translate("MainWindow", "<html><head/><body><p>Click on the respective buttons to select &amp; upload the files</p></body></html>"))
        self.upload_source_button.setText(_translate("MainWindow", "Source"))
        self.upload_target_button.setText(_translate("MainWindow", "Target"))
        self.page_2_body_2.setText(_translate("MainWindow", "Once the files are uploaded successfully, click on the button below"))
        self.prevalidate_button.setText(_translate("MainWindow", "Pre-Validate"))
        self.prevalidate_button.setEnabled(False)
        self.page_2_output_label.setText(_translate("MainWindow", "Output Viewer"))
        self.page_2_back_button.setText(_translate("MainWindow", "Back"))
        self.page_2_next_button.setText(_translate("MainWindow", "Next"))
        self.page_3_header_text.setText(_translate("MainWindow", "SELECT OR UPLOAD RULES"))
        self.page_3_intro_text.setText(_translate("MainWindow", "Click on the dropdowns below to set your rules.\n"
                                                "You can upload it as your pre-set rules"))
        self.source_column_label.setText(_translate("MainWindow", "Source Column"))
        self.target_column_label.setText(_translate("MainWindow", "Target Column"))
        self.operator_label.setText(_translate("MainWindow", "Operator"))
        self.construct_rule_label.setText(_translate("MainWindow", "Output:"))
        self.clear_rule_construct_button.setText(_translate("MainWindow", "Clear"))
        self.validate_button.setText(_translate("MainWindow", "Validate"))
        self.label.setText(_translate("MainWindow", "Validated:"))
        self.upload_rules_button.setText(_translate("MainWindow", "Upload"))
        self.delete_rule_button.setText(_translate("MainWindow", "Delete"))
        self.page_3_back_button.setText(_translate("MainWindow", "Back"))
        self.reconcile_button.setText(_translate("MainWindow", "Reconcile"))
        self.page_4_header_text.setText(_translate("MainWindow", "SUMMARY"))
        self.progress_bar_text.setText(_translate("MainWindow", "Loading... Please wait..."))
        self.checkBox.setText(_translate("MainWindow", "Remove Source File"))
        self.checkBox_2.setText(_translate("MainWindow", "Remove Target File"))
        self.restart_button.setText(_translate("MainWindow", "Restart"))

        # Connecting buttons
        self.page_1_next_button.clicked.connect(self.nextCurrentIndex)
        self.page_2_next_button.clicked.connect(self.nextCurrentIndex)
        self.reconcile_button.clicked.connect(self.nextCurrentIndex)
        self.restart_button.clicked.connect(self.nextCurrentIndex)
        self.page_2_back_button.clicked.connect(self.prevCurrentIndex)
        self.page_3_back_button.clicked.connect(self.prevCurrentIndex)
        self.upload_source_button.clicked.connect(self.uploadSource)
        self.upload_target_button.clicked.connect(self.uploadTarget)
        self.prevalidate_button.clicked.connect(self.prevalidate)
        self.source_column_dropdown.currentTextChanged.connect(self.concatSourceDD)
        self.target_column_dropdown.currentTextChanged.connect(self.concatTargetDD)
        self.operator_dropdown.currentTextChanged.connect(self.concatOpDD)
        self.clear_rule_construct_button.clicked.connect(self.clearConstructingRule)
        self.validate_button.clicked.connect(self.ruleValidation)
        self.upload_rules_button.clicked.connect(self.uploadRules)
        self.delete_rule_button.clicked.connect(self.ruleDeleting)

        self.page_2_next_button.setDisabled(True)
        self.reconcile_button.setDisabled(True)
        self.restart_button.setDisabled(True)

    def prevCurrentIndex(self):
        ### FUNCTION FOR 'BACK' BUTTONS ###
        rgb246 = "border: none; background-color: rgb(246, 246, 246)"
        rgb250 = "border: none; background-color: rgb(255, 255, 255)"

        if self.stacked_Widget.currentIndex() == 1:
            self.stacked_Widget.setCurrentIndex(0)
            self.tutorial_tab.setStyleSheet(rgb246)
            self.upload_tab.setStyleSheet(rgb250)

        elif self.stacked_Widget.currentIndex() == 2:
            self.stacked_Widget.setCurrentIndex(1)
            self.upload_tab.setStyleSheet(rgb246)
            self.rule_tab.setStyleSheet(rgb250)

    def nextCurrentIndex(self):
        ### FUNCTION FOR 'NEXT' & 'RECONCILE' BUTTONS ###
        rgb246 = "border: none; background-color: rgb(246, 246, 246)"
        rgb250 = "border: none; background-color: rgb(255, 255, 255)"

        if self.stacked_Widget.currentIndex() == 0:
            self.stacked_Widget.setCurrentIndex(1)
            self.upload_tab.setStyleSheet(rgb246)
            self.tutorial_tab.setStyleSheet(rgb250)

        elif self.stacked_Widget.currentIndex() == 1:
            self.stacked_Widget.setCurrentIndex(2)
            self.rule_tab.setStyleSheet(rgb246)
            self.upload_tab.setStyleSheet(rgb250)

        elif self.stacked_Widget.currentIndex() == 2:
            self.stacked_Widget.setCurrentIndex(3)
            self.recon_tab.setStyleSheet(rgb246)
            self.rule_tab.setStyleSheet(rgb250)
            Ui_MainWindow.reconcile(self)

        elif self.stacked_Widget.currentIndex() == 3:
            self.stacked_Widget.setCurrentIndex(1)
            self.upload_tab.setStyleSheet(rgb246)
            self.recon_tab.setStyleSheet(rgb250)
            Ui_MainWindow.restart(self)

    def loadingSpinner(self, int):
        ### START LOADING SPINNER WHEN QTHREAD IS WORKING & END WHEN QTHREAD IS DONE WORKING ###
        if int == 1:
            self.loader = LoadingScreen(self)
        if int == 2:
            self.loader.movie.stop()
            self.loader.close()

    def uploadSource(self):
        ### UPLOADING SOURCE FILE ###
        
        fileName = QFileDialog.getOpenFileName(None, 'Open File', 'c;\\', "Excel files (*.xls *.xlsx)")  # 'c:\\'
        if fileName != ('', ''):
            splitList = re.split('/', fileName[0])
            start = time.time()
            df = pd.read_excel(fileName[0], header=0, sheet_name=0)
            end = time.time()
            dict = {"dataframe": df, "start": start, "end": end, "splitList": splitList}
            Ui_MainWindow.uploadedSource(self, dict)

    def uploadedSource(self, dict):
        ### ONCE SELECTED FILE HAS BEEN READ INTO PANDAS DATAFRAME IN THE WORKER THREAD, FOLLOWING STEPS ARE TO CONCLUDE UPLOADING STATUS & CHANGES 
        try:
            self.df = dict["dataframe"]
            if self.df.empty is True:
                now = datetime.datetime.now()
                self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: File is empty, please upload again.\n\n".format(now.strftime("%I:%M %p")))
                self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
            else:
                self.source_file_name.setText(dict["splitList"][-1])
                self.source_file_name.setToolTip(dict["splitList"][-1])
                self.source_file_name.adjustSize()
                for x in self.df:
                    if self.df[x].isnull().all() == True:
                        reply = self.dtypePopup(self, x)
                        now = datetime.datetime.now()
                        if reply == 0:
                            self.df[x] = self.df[x].replace(
                                np.nan, 0, regex=True)
                            self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {} is converted to Numeric!\n\n".format(now.strftime("%I:%M %p"), x))
                            self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
                        elif reply == 1:
                            self.df[x] = self.df[x].replace(
                                np.nan, '', regex=True)
                            self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {} is converted to Text!\n\n".format(now.strftime("%I:%M %p"), x))
                            self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
                        else:
                            self.df[x] = self.df[x].replace(
                                np.nan, datetime.date(2099, 1, 1), regex=True)
                            self.df[x] = self.df[x].astype('datetime64[ns]')
                            self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {} is converted to Date!\n\n".format(now.strftime("%I:%M %p"), x))
                            self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())

                convertSC = "Please take note of changes in column names:\n\n"
                before = str(self.df.columns.values)
                withoutSC = []
                # Substitute special characters, period, spaces and so on to underscore, therefore reducing chances of errors when passing data into sqlite
                for x in self.df.columns:
                    x = re.sub(r"^\W+|^\d+|^ +", "", x)
                    x = re.sub(r"\.+|\s+", "_", x)
                    withoutSC.append(x)
                self.df.columns = withoutSC
                # As we're helping user to tweak column names, we have to let user know about the change
                if before != str(self.df.columns.values):
                    for x in self.df.columns.values:
                        convertSC += "{}\n".format(x)
                    convertSC = convertSC[:len(convertSC)-1]
                else:
                    # Eventually convertSC will concatenate into the message. If there's no change then the empty string will not show anything.
                    convertSC = ""
                
        except:
            logging.exception("Got exception at uploading source file")
            raise
        
        # If both dataframes exist (there might be a possibility user uploads target already but wants to re-upload source)
        if self.df is not None and self.df.empty is False:
            
            try:
                if self.df1.any:
                    now = datetime.datetime.now()
                    self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: Ensure selected files are correct before clicking on Pre-Validate.\n{}\nDuration to read Source file: {}s\n\n".format(now.strftime("%I:%M %p"), convertSC, round(dict["end"]-dict["start"], 2)))
                    self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
                    self.upload_source_button.setDisabled(True)
                    self.prevalidate_button.setEnabled(True)
            except AttributeError:
                # Happens when there's only source
                self.upload_target_button.setEnabled(True)
                x = datetime.datetime.now()
                self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {}\n\nDuration taken to upload Source file: {}s. Please upload target file.\n\n".format(x.strftime("%I:%M %p"), convertSC, round(dict["end"]-dict["start"], 2)))
                self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())

    def uploadTarget(self):
        ### UPLOADING TARGET FILE ###
        
        fileName = QFileDialog.getOpenFileName(None, 'Open File', 'c;\\', "Excel files (*.xls *.xlsx)")  # 'c:\\'
        if fileName != ('', ''):
            splitList = re.split('/', fileName[0])
            start = time.time()
            df = pd.read_excel(fileName[0], header=0, sheet_name=0)
            end = time.time()
            dict = {"dataframe": df, "start": start, "end": end, "splitList": splitList}
            Ui_MainWindow.uploadedTarget(self, dict)

    def uploadedTarget(self, dict):
        ### ONCE SELECTED FILE HAS BEEN READ INTO PANDAS DATAFRAME IN THE WORKER THREAD, FOLLOWING STEPS ARE TO CONCLUDE UPLOADING STATUS & CHANGES 
        try:
            self.df1 = dict["dataframe"]
            if self.df1.empty is True:
                now = datetime.datetime.now()
                self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: File is empty, please upload again.\n\n".format(now.strftime("%I:%M %p")))
                self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
            else:
                self.target_file_name.setText(dict["splitList"][-1])
                self.target_file_name.setToolTip(dict["splitList"][-1])
                self.target_file_name.adjustSize()
                for x in self.df1:
                    if self.df1[x].isnull().all() == True:
                        reply = Ui_MainWindow.dtypePopup(self, x)
                        now = datetime.datetime.now()
                        if reply == 0:
                            self.df1[x] = self.df1[x].replace(
                                np.nan, 0, regex=True)
                            self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {} is converted to Numeric!\n\n".format(now.strftime("%I:%M %p"), x))
                            self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
                        elif reply == 1:
                            self.df1[x] = self.df1[x].replace(
                                np.nan, '', regex=True)
                            self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {} is converted to Text!\n\n".format(now.strftime("%I:%M %p"), x))
                            self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
                        else:
                            self.df1[x] = self.df1[x].replace(
                                np.nan, datetime.date(2099, 1, 1), regex=True)
                            self.df1[x] = self.df1[x].astype('datetime64[ns]')
                            self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {} is converted to Date!\n\n".format(now.strftime("%I:%M %p"), x))
                            self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())

                convertSC1 = "Please take note of changes in column names:\n\n"
                before = str(self.df1.columns.values)
                withoutSC1 = []
                for x in self.df1.columns:
                    x = re.sub(r"^\W+|^\d+|^ +", "", x)
                    x = re.sub(r"\.+|\s+", "_", x)
                    withoutSC1.append(x)
                self.df1.columns = withoutSC1
                if before != str(self.df1.columns.values):
                    for x in self.df1.columns.values:
                        convertSC1 += "{}\n".format(x)
                    convertSC1 = convertSC1[:len(convertSC1)-1]
                else:
                    convertSC1 = ""
        except:
            logging.exception("Got exception at uploading target file")
            raise
        if self.df1 is not None and self.df1.empty is False:
            self.prevalidate_button.setEnabled(True)
            now = datetime.datetime.now()
            self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {}\n\nDuration taken to read Target file: {}s. Ensure selected files are correct before clicking the Pre-Validate button.\n\n".format(now.strftime("%I:%M %p"), convertSC1, round(dict["end"]-dict["start"], 2)))
            self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())

    def prevalidate(self):
        ### ASSIGNS A WORKER THREAD TO PROCESS PRE-VALIDATION CHECK TO PREVENT UI FROM FREEZING ###
        self.worker = WorkerThread1(self.df, self.df1)
        self.worker.start()
        self.worker.worker_loading.connect(self.loadingSpinner)
        self.worker.worker_complete.connect(self.prevalidated)

    def prevalidated(self, dict):
        ### RETURNS DICTIONARY FROM WORKER THREAD WITH RELEVANT DATA (DATAFRAME, ERRORMESSAGE, WARNINGCOUNTS, PREVALIDATION TIME) & DETERMINES THE CORRESPONDING OUTCOME
        self.df = dict["df"]
        self.df1 = dict["df1"]
        self.src_changecolA = dict["src_changecolA"]
        self.tgt_changecolA = dict["tgt_changecolA"]
        now = datetime.datetime.now()
        
        if "errorMessage" in dict:
            self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: {}\n\nDuration to handle errors: {}s\n\n".format(now.strftime("%I:%M %p"), dict["errorMessage"], round(dict["time"], 2)))
            self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
        else:
            if dict["warningCount"] != 0:
                reply = Ui_MainWindow.prevalidatePopup(self)
                if reply == 0:  # 0 represents the first button which is "Yes"
                    self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: Pre-validation pass successfully with the acknowledgement of prompted warnings & no errors. Please continue by clicking Next.\n\nDuration to handle warnings: {}s\n\n".format(now.strftime("%I:%M %p"), round(dict["time"], 2)))
                    self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
                    self.page_2_next_button.setEnabled(True)
                    Ui_MainWindow.populateDropdowns(self)
                else:
                    self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "You have selected 'Not to Proceed Further' due to warnings, please pre-validate again if you want to proceed with the same files. Or else, please fix and re-upload files again.\n\n")
                    self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
            else:
                self.page_2_output_list.setText(self.page_2_output_list.text().lstrip('\n').lstrip(' ') + "{}: Pre-validation pass successfully. No errors/warnings\n\nDuration to handle warnings: {}s\n\n".format(now.strftime("%I:%M %p"), round(dict["time"], 2)))
                self.page_2_scrollArea.verticalScrollBar().setValue(self.page_2_scrollArea.verticalScrollBar().maximum())
                self.page_2_next_button.setEnabled(True)
                Ui_MainWindow.populateDropdowns(self)

    @ staticmethod
    def populateDropdowns(self):
        ### ONCE PRE-VALIDATED, ALL COLUMN NAMES & OPERATORS WILL BE APPENDED INTO DROPDOWNS ###
        try:
            # Establish connection with SQLite3 to create a database file & store under folder "Database"
            conn = sqlite3.connect('Database\Autorecon.db')
            c = conn.cursor()

            # self.srcCol & self.tgtcol will store respective column names & its datatype
            self.srcCol = {}
            self.tgtCol = {}

            # similar to 'DROP TABLE IF EXISTS' concept
            self.source_column_dropdown.clear()
            self.target_column_dropdown.clear()
            self.operator_dropdown.clear()

            # Appending column names & operators to respective dropdown 
            c.execute('''pragma table_info('SOURCE')''')
            for col, row in zip(self.df.columns, c.fetchall()): 
                self.srcCol[col] = row[2]
                self.source_column_dropdown.addItem("src."+ str(col))

            c.execute('''pragma table_info('TARGET')''')
            for col, row in zip(self.df1.columns, c.fetchall()):
                self.tgtCol[col] = row[2]
                self.target_column_dropdown.addItem("tgt."+ str(col))

            self.comparator = ["=", "!=", "<", "<=", ">", ">=", "+", "-"]
            for op in self.comparator:
                self.operator_dropdown.addItem(op)

            # Set current index for the 3 dropdowns at Construct Rule tab
            self.source_column_dropdown.setCurrentIndex(-1)
            self.target_column_dropdown.setCurrentIndex(-1)
            self.operator_dropdown.setCurrentIndex(-1)
            self.construct_rule_browser.clear()

        except:
            logging.exception("Got exception at populating dropdowns")
            raise

    def concatSourceDD(self, value):
        ### CONCAT FUNCTIONS TO APPEND THE VALUE FROM DROPDOWNS TO THE FIELD WIDGET BELOW IT ###
        self.construct_rule_browser.setText(self.construct_rule_browser.toPlainText() + value + " ")
        self.source_column_dropdown.setCurrentIndex(-1)

    def concatTargetDD(self, value):
        ### CONCAT FUNCTIONS TO APPEND THE VALUE FROM DROPDOWNS TO THE FIELD WIDGET BELOW IT ###
        self.construct_rule_browser.setText(self.construct_rule_browser.toPlainText() + value + " ")
        self.target_column_dropdown.setCurrentIndex(-1)

    def concatOpDD(self, value):
        ### CONCAT FUNCTIONS TO APPEND THE VALUE FROM DROPDOWNS TO THE FIELD WIDGET BELOW IT ###
        self.construct_rule_browser.setText(self.construct_rule_browser.toPlainText() + value + " ")
        self.operator_dropdown.setCurrentIndex(-1)

    def clearConstructingRule(self):
        ### CONNECTED TO CLEAR BUTTON IN CASE WHEN USER SELECTED WRONG COLUMNS/OPERATORS WHEN CONSTRUCTING THE RULE
        self.construct_rule_browser.clear()

    def ruleValidation(self):
        ### VALIDATE SELECTED RULE TO ENSURE SANITY ###
        # TEXT Data Type (only use = or !=)
        # DATE Data Type (Use =, !=, <, <=, >, >=)
        # INT Data Type (use all operators)
        # Enable comparison of more than 2 columns (Src.NFA_Amt = Src.Financed_Amt - Src.Body_Funding_Amt)
        # Should only have 1 equal comparator (or equivalents) in a rule
        try:
            # Split the string retrieved from the field widget into pieces and remove ''
            ruleBlocks = re.split(' ', self.construct_rule_browser.toPlainText())
            while '' in ruleBlocks:
                ruleBlocks.remove('')

            errorMessage = []

            ### SYNTAX LOGIC ###

            # Need to have at least 3 parts to form a rule
            if len(ruleBlocks) < 3:
                errorMessage.append('Rule needs to have at least two columns with operator for comparison.')

            # Due to structure of the recon rule, position of column & operator should be alternating
            for x in ruleBlocks[1::2]: #sliced list should only contain operators
                if x not in self.comparator:
                    errorMessage.append('Operator should be placed in even index position.') # syntax wrong, try again

            for x in ruleBlocks[::2]: #sliced list should only contain column headers
                if x in self.comparator:
                    errorMessage.append('Column should be placed in odd index position.') # syntax wrong, try again

            if ruleBlocks[-1] in self.comparator:
                errorMessage.append('Operator should not be at the end of the rule.') # syntax wrong, try again

            ### DATATYPE LOGIC ###
            
            dtypeList = []
            for x in ruleBlocks[::2]:
                if x.split('.')[0] == 'src':
                    dtypeList.append(self.srcCol[x.split('.')[1]]) #record datatype of selected cols
                elif x.split('.')[0] == 'tgt':
                    dtypeList.append(self.tgtCol[x.split('.')[1]]) #record datatype of selected cols

            dtypeList = list(dict.fromkeys(dtypeList)) #remove duplicating values
            if len(dtypeList) > 1 and dtypeList != ['REAL','INTEGER'] and dtypeList != ['INTEGER','REAL']:
                errorMessage.append(f'Unable to compare different datatypes.')

            if dtypeList == ["TEXT"]:
                if len(ruleBlocks[1::2]) > 1: 
                    errorMessage.append('Text-type rules can only have 1 operator.') 
                for x in ruleBlocks[1::2]:
                    if x not in ["=", "!="]:
                        errorMessage.append('Text-type rules applicable only to "=" & "!=".') 

            elif dtypeList == ["TIMESTAMP"]:
                if len(ruleBlocks[1::2]) > 1:
                    errorMessage.append('Date-type rules can only have 1 operator.')
                for x in ruleBlocks[1::2]:
                    if x not in ["=", "!=", '<', "<=", ">", ">="]:
                        errorMessage.append('Date-type rules appplicable only to operators that return True/False.')

            else: # For dtypeList == INTEGER OR REAL
                counter = 0
                for x in ruleBlocks[1::2]:
                    if x in ["=", "!=", '<', "<=", ">", ">="]:
                        counter += 1
                if counter > 1:
                    errorMessage.append('Integer-type rules can only use 1 operator that returns True/False, rest should be "+" or "-".')

            ### CONCLUSION ###  

            if errorMessage == []:
                Ui_MainWindow.ruleAppending(self)
            else:
                self.construct_rule_browser.clear()
                errorString = ''
                for x in errorMessage:
                    errorString += f"{x}\n"
                errorString = errorString[:-2] 
                msg = QMessageBox()
                msg.setWindowTitle("Rule Validation Error")
                msg.setText(errorString)
                msg.addButton("Ok", QMessageBox.YesRole)
                x = msg.exec_()
                return x

        except:
            logging.exception("Got exception at rule validation")
            raise

    @ staticmethod
    def ruleAppending(self):
        ### ADD VALIDATED RULE INTO THE LIST WIDGET FOR VISUAL FEEDBACK & RECONCILE ###
        try:
            text = self.construct_rule_browser.toPlainText()[:-2].replace("  ", " ")
            self.listWidget.addItem(text)
            self.reconcile_button.setEnabled(True)
            self.construct_rule_browser.clear()
        except:
            logging.exception("Got exception at rule appending")
            raise

    def ruleDeleting(self):
        ### DELETE RULE(S) IN THE LIST WIDGET ###
        try:
            self.listWidget.takeItem(self.listWidget.currentRow())
        except:
            logging.exception("Got exception at rule deleting")
            raise

    def uploadRules(self): 
        try:
            ### READS THE "RULES" TAB FROM EXCEL TO APPLY THE SAME SET OF RULES QUICKER FOR NEXT ITERATION ###
            rfName = QFileDialog.getOpenFileName(None, 'Open File', 'c;\\', "Excel files (*.xls *.xlsx)")  # 'c:\\'
            if rfName != ('', ''):
                self.ruleBank = pd.read_excel(
                    rfName[0], header=0, sheet_name="RULES", engine='openpyxl')
                
                # Add the prefix src. & tgt. to the keys 
                self.srcCol = {f"src.{k}":v for (k,v) in self.srcCol.items()}
                self.tgtCol = {f"tgt.{k}":v for (k,v) in self.tgtCol.items()}

                # Merge the two dictionaries to one
                colList = {}
                colList.update(self.srcCol)
                colList.update(self.tgtCol)

                columnDontExist = "" #Column Name does not exist in Database:\n\n"
                differentDtype = "" #Unacceptable comparison between the two columns:\n\n"
                wrongSyntax = ""
                acceptedRules = {}

                for x in range(len(self.ruleBank)):
                    count = 0 
                    ruleBank = self.ruleBank.iloc[x].values.tolist()
                    ruleBank = [x.replace("'", "") for x in ruleBank[1:]]
                    ruleBank = [x for x in ruleBank if x]
                    ### SYNTAX LOGIC ###
                    
                    # Need to have at least 3 parts to form a rule
                    if len(ruleBank) < 3:
                        wrongSyntax += f'Uploaded Rule #{x+1}: Do not fit the criteria of at having at least two columns and an operator.'
                        count += 1

                    # Due to structure of the recon rule, position of column & operator should be alternating
                    for y in ruleBank[1::2]: #sliced list should only contain operators
                        if y not in self.comparator:
                            wrongSyntax += f'Uploaded Rule #{x+1}: Operator should be placed in even index position.' # syntax wrong, try again
                            count += 1

                    for y in ruleBank[::2]: #sliced list should only contain column headers
                        if y in self.comparator:
                            wrongSyntax += f'Uploaded Rule #{x+1}: Column should be placed in odd index position.' # syntax wrong, try again
                            count += 1

                    ### COLUMN NAMING LOGIC ###

                    d = {j:i for i,j in enumerate(ruleBank)}
                    ruleBankwoComp  = sorted(list((set(ruleBank) - set(self.comparator))), key = lambda x: d[x])
                    for y in ruleBankwoComp:
                        if y not in colList:
                            columnDontExist += f'Uploaded Rule #{x+1}: {y} does not exist'
                            count += 1
                            del colList[y]
                    
                    ### DATATYPE LOGIC ###

                    dtypeList = []
                    for y in ruleBank[::2]:
                        dtypeList.append(colList[y])
                    dtypeList = list(dict.fromkeys(dtypeList)) #remove duplicating values

                    if len(dtypeList) > 1:
                        differentDtype += f'Uploaded Rule #{x+1}: Unable to compare different datatypes.'
                        count += 1

                    if dtypeList == ["TEXT"]:
                        if len(ruleBank[1::2]) > 1: 
                            differentDtype += f'Uploaded Rule #{x+1}: Text-type rules can only have 1 operator.'
                            count += 1 

                        for y in ruleBank[1::2]:
                            if y not in ["=", "!="]:
                                differentDtype += f'Uploaded Rule #{x+1}: Text-type rules applicable only to "=" & "!=".'
                                count += 1

                    elif dtypeList == ["TIMESTAMP"]:
                        if len(ruleBank[1::2]) > 1:
                            differentDtype += f'Uploaded Rule #{x+1}: Date-type rules can only have 1 operator.'
                            count += 1

                        for y in ruleBank[1::2]:
                            if y not in ["=", "!=", '<', "<=", ">", ">="]:
                                differentDtype += f'Uploaded Rule #{x+1}: Date-type rules appplicable only to operators that return True/False.'
                                count += 1

                    else: # For dtypeList == INTEGER OR REAL
                        counter = 0
                        for y in ruleBank[1::2]:
                            if y in ["=", "!=", '<', "<=", ">", ">="]:
                                counter += 1
                        if counter > 1:
                            differentDtype += f'Uploaded Rule #{x+1}: Integer-type rules can only use 1 operator that returns True/False, rest should be "+" or "-".'
                            count += 1
                    
                    ### APPEND VALID RULE(S) ###

                    if count != 0:
                        print("Has error validating")
                    else:
                        acceptedRules[x+1] = " ".join(ruleBank)

                    print(f"iteration {x+1} done")

                for a in acceptedRules:
                    # print(f"acceptedRules[a]: {acceptedRules[a]}")
                    self.listWidget.addItem(acceptedRules[a])
                    
                if columnDontExist != "" or differentDtype != "" or wrongSyntax != "":
                    msg = QMessageBox()
                    msg.setWindowTitle("Uploaded Rule(s) Validation Error")
                    msg.setText(columnDontExist + differentDtype + wrongSyntax)
                    msg.addButton("Ok", QMessageBox.YesRole)
                    msg.exec_()
                
                self.delete_rule_button.setEnabled(True)
                self.reconcile_button.setEnabled(True)

        except:
            logging.exception("Got exception at uploading rules")
            raise
        
    @ staticmethod
    def reconcile(self):
        ### GENERATES AN EXCEL FILE WITH THE COMPARISON & OTHER DATA AS OUTPUT ###
        try:
            # Map the progression of this function to the progress bar
            progress = 0
            total = 15
            self.progressBar.setValue(0)
            
            # QApplication.processEvents()

            # Create a workbook for output
            t1 = time.time()
            if path.exists('Autorecon.xlsx'):
                xl = w3c.Dispatch('Excel.Application')
                wb = xl.Workbooks.Open('{}/Autorecon.xlsx'.format(str(os.getcwd()).replace("\\", "/")))
                wb.Close(True)

            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Contains the original list of rule(s)
            rules = []
            for index in range(self.listWidget.count()):
                rules.append(self.listWidget.item(index).text())
            
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Rename first column for source & target to "ColumnA" placeholder for SQL writing
            items = [w.replace(str("src." + self.src_changecolA), "src.ColumnA") for w in rules]
            items = [w.replace(str("tgt." + self.tgt_changecolA), "tgt.ColumnA") for w in items]

            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Transform the data in the above list to a dictionary
            count = 1
            ruleDict = {} 
            for item in items:
                itempart = item.split(" ")
                ruleDict[str(count)] = itempart
                count += 1
            
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Declaring the SQL command for each rule that users want
            count = 1
            createtable = "CREATE TABLE BIG_TABLE AS\n"
            selectStatement = "SELECT DISTINCT src.ColumnA src_ColumnA, tgt.ColumnA tgt_ColumnA, "
            for key in ruleDict:
                for value in ruleDict[key]:
                    if "." in value:
                        z = value.replace(".", "_")
                        selectStatement += "{} R{}_{}, ".format(value, count, z)

                x = ['<', '<=', '=', '!=', '>', '>=']
                result = [i for i in ruleDict[key] if i.startswith(tuple(x))]
                if result != []:
                    # if x in ''.join(ruleDict[key]):
                    selectStatement += "CASE WHEN "
                    for value in ruleDict[key]:
                        if value not in ['<', '<=', '=', '!=', '>', '>=', '+', '-']:
                            selectStatement += f"IFNULL({value},\"\") "
                        else: 
                            selectStatement += f"{value} "
                    selectStatement += f"THEN 'YES' ELSE 'NO' END '{' '.join(ruleDict[key])}', "
                else:
                    for value in ruleDict[key]:
                        if value not in ['<', '<=', '=', '!=', '>', '>=', '+', '-']:
                            selectStatement += f"IFNULL({value}, 0) "
                        else:
                            selectStatement += f"{value} "
                    selectStatement += f"'{' '.join(ruleDict[key])}', "
                count += 1
            selectStatement = selectStatement[:-2]
            
            SQLStatement = f"{createtable}{selectStatement}\nFROM source src LEFT JOIN target tgt USING(ColumnA)\nUNION ALL\n{selectStatement}\nFROM target tgt LEFT JOIN source src USING(ColumnA) WHERE src.ColumnA IS NULL"

            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Record down the SQL statement just in case
            logging.info(f"OUTPUT SQL QUERY\n{SQLStatement}\n\n")

            conn = sqlite3.connect('Database\Autorecon.db')
            c = conn.cursor()
            c.execute('''DROP TABLE IF EXISTS R_RULES_TBL''')

            # Since we've given autonomy for users to select more complex rules, the creating of rule table's columns will vary for each use
            createRuleTableStr = "CREATE TABLE R_RULES_TBL (SN REAL, "
            for x in range(len(ruleDict[max(ruleDict, key=lambda k: len(ruleDict[k]))])):
                x += 1
                createRuleTableStr += "VAL{} TEXT, ".format(x)
            createRuleTableStr = createRuleTableStr[:-2] + ");"
            c.executescript(createRuleTableStr)

            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Constructing the sql statement for capturing the rule(s) for re-use purpose
            conn.commit()
            abc = ["VAL{}".format(x+1) for x in range(len(ruleDict[max(ruleDict, key=lambda k: len(ruleDict[k]))]))]
            insertRuleTableStr = "INSERT INTO R_RULES_TBL (SN, {}) VALUES".format(', '.join(abc))
            for i, rule in enumerate(rules):
                insertRuleTableStr += " ({}, '{}".format(i+1, rule.replace(" ", "', '"))
                if len(rule.split(" ")) < len(ruleDict[max(ruleDict, key=lambda k: len(ruleDict[k]))]):
                    insertRuleTableStr += "'"
                    for x in range(len(ruleDict[max(ruleDict, key=lambda k: len(ruleDict[k]))]) - len(rule.split(" "))):
                        insertRuleTableStr += ", ''"
                    insertRuleTableStr += "),"
                else:
                    insertRuleTableStr += "'),"
            insertRuleTableStr = insertRuleTableStr[:-1]
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            logging.info(f"INSERT RULE TABLE QUERY\n{insertRuleTableStr}\n\n")
            c.executescript(insertRuleTableStr)
            
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            c.execute('''DROP TABLE IF EXISTS BIG_TABLE''')
            c.execute('''{}'''.format(SQLStatement))
            
            # Declaring the excel file & the 4 sheets attached to it
            workbook = Workbook('AutoRecon.xlsx')
            worksheet = workbook.add_worksheet(name="RECON_OUTPUT")
            worksheet1 = workbook.add_worksheet(name="SOURCE")
            worksheet2 = workbook.add_worksheet(name="TARGET")
            worksheet3 = workbook.add_worksheet(name="RULES")
            
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Appending output data into 1st sheet of excel
            c.execute('''PRAGMA table_info('BIG_TABLE')''')
            colList = list(row[1] for row in c.fetchall())

            colList = [value.replace("src.ColumnA", str("src." + self.src_changecolA)) for value in colList]
            colList = [value.replace("tgt.ColumnA", str("tgt." + self.tgt_changecolA)) for value in colList]

            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            #worksheet.write_row(0,0,colList)
            worksheet.write(0,0,str("src_" + self.src_changecolA))
            worksheet.write(0,1,str("tgt_" + self.tgt_changecolA))
            worksheet.write_row(0, 2, colList[2:])
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            #Populating output data into 1st sheet of excel
            c.execute('''SELECT * FROM BIG_TABLE ''')
            count = 1
            for x in c.fetchall(): #the values
                worksheet.write_row(count, 0, list(x))
                count += 1
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Populating source data into 2nd sheet of excel
            c.execute('''PRAGMA table_info('SOURCE')''')
            colList = list(row[1] for row in c.fetchall())
            worksheet1.write(0,0,self.src_changecolA)
            worksheet1.write_row(0, 1, colList[1:])
            c.execute('''SELECT * FROM source''')
            count = 1
            for x in c.fetchall():
                worksheet1.write_row(count, 0, list(x))
                count += 1
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Populating target data into 3rd sheet of excel
            c.execute('''PRAGMA table_info('TARGET')''')
            colList = list(row[1] for row in c.fetchall())
            worksheet2.write(0,0,self.tgt_changecolA)
            worksheet2.write_row(0, 1, colList[1:])
            c.execute('''SELECT * FROM target''')
            count = 1
            for x in c.fetchall():
                worksheet2.write_row(count, 0, list(x))
                count += 1
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Populating rules data into 4th sheet of excel 
            c.execute('''PRAGMA table_info('R_RULES_TBL')''')
            colList = list(row[1] for row in c.fetchall())
            worksheet3.write_row(0, 0, colList)
            c.execute('''SELECT * FROM R_RULES_TBL''')
            for i, x in enumerate(c.fetchall()):
                worksheet3.write_row(i+1, 0, ["'{}".format(y) for y in x])
            workbook.close()
            progress += 1
            self.progressBar.setValue(int((progress/total)*100))
            
            # Allow user to click to restart / start new round
            self.restart_button.setEnabled(True)
            
            self.progress_bar_text.setText("Done!")
            t2 = time.time()
            self.listView.setText("Reconciliation is successful! The output file is stored in the same directory as this executable file.\n\nSummary:\nSource file: {}\nTarget file: {}\nRule(s):\n{}\nOutput file: AutoRecon.xlsx\n\nSession ended...\nDuration to reconcile: {}s".format(self.source_file_name.text(), self.target_file_name.text(), '\n'.join(rules), round(t2-t1, 2)))

            for i in range(self.listWidget.count()):
                read = self.listWidget.item(i).text().split(" ")
                if read[0] == "src.ColumnA":
                    read[0] = "src." + self.src_changecolA
                if read[0] == "tgt.ColumnA":
                    read[0] = "tgt." + self.tgt_changecolA
                if read[2] == "tgt.ColumnA":
                    read[2] = "tgt." + self.tgt_changecolA
                if read[2] == "src.ColumnA":
                    read[2] = "src." + self.src_changecolA
                self.listWidget.takeItem(i)
                self.listWidget.addItem(" ".join(read))

        except:
            logging.exception("Got exception at reconcile")
            raise

        conn.close()

    @ staticmethod
    def restart(self):
        ### RESTART FUNCTION SO USERS DO NOT HAVE TO CLOSE APP TO RE-RUN IT ###
        try:
            lambda: self.stacked_Widget.setCurrentIndex(1)
            self.upload_source_button.setEnabled(True)
            self.upload_target_button.setEnabled(True)
            self.prevalidate_button.setDisabled(True)
            self.page_2_next_button.setDisabled(True)
            self.operator_dropdown.clear()
            self.source_column_dropdown.clear()
            self.target_column_dropdown.clear()
            self.listWidget.clear()
            self.page_2_output_list.clear()
            self.progress_bar_text.setText("Loading... Please wait...")
            self.listView.clear()

            if self.checkBox.isChecked() is not True and self.checkBox_2.isChecked() is not True:
                self.prevalidate_button.setEnabled(True)
            if self.checkBox.isChecked():
                self.df = None
                self.source_file_name.clear()
            if self.checkBox_2.isChecked():
                self.df1 = None
                self.target_file_name.clear()

        except:
            logging.exception("Got exception at restart")
            raise

    @staticmethod
    def dtypePopup(self, x):
        ### FOR EMPTY COLUMN(S), USER SHALL STATE THE DATA TYPE TO PROCEED ###
        msg = QMessageBox()
        msg.setText("Column '{}' currently is empty. Please state the data type for this column to manually declare it to database.".format(x))
        msg.addButton(QPushButton("Numeric"), QMessageBox.YesRole)
        msg.addButton(QPushButton("Text"), QMessageBox.NoRole)
        msg.addButton(QPushButton("Date"), QMessageBox.RejectRole)
        msg.setWindowTitle("Ambiguous Datatype")
        msg.setIcon(QMessageBox.Question)
        x = msg.exec_()
        return x

    @staticmethod
    def prevalidatePopup(self):
        ### STORES WARNING MESSAGE(S) DURING PRE-VALIDATION TO EXCEL & PROMPT USER ###
        msg = QMessageBox()
        msg.setWindowTitle("Warning Message")
        path = str(QDir.currentPath()) + \
            r"/Warning Message/warningMessage.xlsx"
        url = bytearray(QUrl.fromLocalFile(path).toEncoded()).decode()
        msg.setText("<a href={}>{}</a> Do you still want to proceed?".format(url,
                                               "Click here for warning details;"))
        msg.addButton(QPushButton("Yes"), QMessageBox.YesRole)
        msg.addButton(QPushButton("No"), QMessageBox.NoRole)
        msg.setIcon(QMessageBox.Question)
        x = msg.exec_()
        return x
    
class WorkerThread(QThread):
    ### CREATE A WORKER THREAD TO WORK ON THE PROCESSING OF FILE UPLOADS / BASICALLY ENSURES THE UI THREAD DOES NOT FREEZE ### 
    
    worker_complete = pyqtSignal(dict)
    worker_loading = pyqtSignal(int)
    splitList = ''
    
    def __init__(self, fileName):
        self.fileName = fileName

    def run(self):
        ### UPLOADING SOURCE FILE ###
        try:
            
            #fileName = QFileDialog.getOpenFileName(None, 'Open File', 'c;\\', "Excel files (*.xls *.xlsx)")  # 'c:\\'
            if self.fileName != ('', ''):
                splitList = re.split('/', self.fileName[0])
                self.worker_loading.emit(1)
                # label1 will show the filename with splitList[-1]
                start = time.time()
                # Pandas will read the excel & convert it into dataframe
                df = pd.read_excel(self.fileName[0], header=0, sheet_name=0, engine='openpyxl')
                end = time.time()
                self.worker_loading.emit(2)
                self.worker_complete.emit({"dataframe": df, "start": start, "end": end, "splitList": splitList})
        except:
            logging.exception("Got exception at uploading file")

class WorkerThread1(QThread):
    ### CREATE ANOTHER THREAD TO WORK ON THE ERROR & WARNING HANDLING FOR PRE-VALIDATION ###

    worker_complete = pyqtSignal(dict)
    worker_loading = pyqtSignal(int)

    def __init__(self, df, df1):
        ### PASS ARGUMENTS (BOTH SRC & TGT DATAFRAME FROM UI_MAINWINDOW CLASS)
        QThread.__init__(self)
        self.df = df
        self.df1 = df1

    def run(self):
        try:
            ### CATCHES ERROR SUCH AS SPECIAL CHAR IN COLUMN NAMES, DUPLICATED COLUMN NAMES ###
            ### CATCHES WARNING SUCH AS DUPLICATED KEYS, COLUMN(S) WITH MIXED DATATYPE ###
            self.worker_loading.emit(1)

            # errorMessage (10 characters) will be a string to show user what are the errors when being pre-validated
            errorMessage = "--ERROR--\n"
            t1 = time.time()

            # Always change first column (aka ColumnA) to string
            self.df.iloc[:,0] = self.df.iloc[:,0].astype(str)
            self.df1.iloc[:,0] = self.df1.iloc[:,0].astype(str)

            # Assign first column name to variable as it needs to be converted to ColumnA for sql use
            self.src_changecolA = self.df.columns[0]
            self.tgt_changecolA = self.df1.columns[0]

            # These are columns with datetime value
            dateCol = [col for col in self.df.select_dtypes(include=[np.datetime64]).columns]
            dateCol1 = [col for col in self.df1.select_dtypes(include=[np.datetime64]).columns]

            # Convert columns with nullable integers back to int32
            # Also strip off empty spaces found in text
            for col in self.df.columns:
                dt = self.df[col].dtype
                if dt == float:
                    if self.df[col].fillna(0).apply(float.is_integer).all():
                        self.df[col] = self.df[col].astype('Int64')
                elif dt == object:
                    self.df[col] = self.df[col].str.strip()
            for col in self.df1.columns:
                dt = self.df1[col].dtype
                if dt == float:
                    if self.df1[col].fillna(0).apply(float.is_integer).all():
                        self.df1[col] = self.df1[col].astype('Int64')
                elif dt == object:
                    self.df1[col] = self.df1[col].str.strip()

            # Check for special characters that are within each column name
            specialChar = "Source Column Name(s) with special characters:\n"
            for x in self.df.columns:
                result = re.search(r"\W", x)
                if result:
                    specialChar += x + "\n"
            if len(specialChar) != 47:
                errorMessage += specialChar + "\n"
            specialChar1 = "Target Column Name(s) with special characters:\n"
            for x in self.df1.columns:
                result = re.search(r"\W", x)
                if result:
                    specialChar1 += x + "\n"
            if len(specialChar1) != 47:
                errorMessage += specialChar1 + "\n"

            # Check if there are column Names with more than n number of characters
            if self.df.columns.str.len().max() > 100:
                errorMessage += "Source Column Name(s) with more than 100 characters:\n"
                for x in self.self.df.columns.values:
                    if len(x) > 100:
                        errorMessage += x + "\n"
                errorMessage += "\n"
            if self.df1.columns.str.len().max() > 100:
                errorMessage += "Target Column Name(s) with more than 100 characters:\n"
                for x in self.df1.columns.values:
                    if len(x) > 100:
                        errorMessage += x + "\n"
                errorMessage += "\n"

            # Convert columns with None & Int back to Int64
            for col in self.df.columns:
                dt = self.df[col].dtype
                if dt == int or dt == float:
                    if col not in dateCol: #dateCol = list of columns that are date
                        self.df[col] = self.df[col].astype('Int64', errors='ignore')
            for col in self.df1.columns:
                dt = self.df1[col].dtype
                if dt == int or dt == float:
                    if col in dateCol1: #dateCol = list of columns that are date
                        self.df1[col] = self.df1[col].astype('Int64', errors='ignore')

            # Revert column name back (read_excel automatically renamed duplicates with ".1")
            self.df.columns = self.df.columns.str.split('.').str[0]
            self.df1.columns = self.df1.columns.str.split('.').str[0]

            # Check for duplicated Column Names
            duplicatedColumn = str(self.df.columns[self.df.columns.duplicated()])
            duplicatedColumn1 = str(self.df1.columns[self.df1.columns.duplicated()])
            duplicatedColumn = duplicatedColumn.split('[')[1].split(']')[0]
            duplicatedColumn1 = duplicatedColumn1.split('[')[1].split(']')[0]

            if duplicatedColumn != "":
                errorMessage += "Duplicated Column Name(s) in Source:\n" + \
                    duplicatedColumn + "\n"
            if duplicatedColumn1 != "":
                errorMessage += "Duplicated Column Name(s) in Target:\n" + \
                    duplicatedColumn1 + "\n"

            t2 = time.time()
            if len(errorMessage) != 10:
                errorMessage += "Please fix the error(s) and reupload the files."
                now = datetime.datetime.now()
                self.worker_loading.emit(2)
                self.worker_complete.emit({"df": self.df, "df1": self.df1, "errorMessage": errorMessage, "time": t2-t1})
            else:
                warningCount = 0

                # Check if datatype is consistent for each column
                # Dictionary to capture column names with mixed datatype
                inconsistentDict = {'Mixed Column(s) with words & numbers in Source:': [], 'Mixed Column(s) with words & numbers in Target:': []}

                # Going through each column & comparing each value to sieve out weird datatype
                for col in self.df.columns: #source
                    weird = (self.df[[col]].applymap(type) !=
                            self.df[[col]].iloc[0].apply(type)).any(axis=1)
                    if len(self.df[weird]) > 0:
                        inconsistentDict['Mixed Column(s) with words & numbers in Source:'].append(col)
                for col in self.df1.columns: #target
                    weird = (self.df1[[col]].applymap(type) != self.df1[[
                        col]].iloc[0].apply(type)).any(axis=1)
                    if len(self.df1[weird]) > 0:
                        inconsistentDict['Mixed Column(s) with words & numbers in Target:'].append(col)

                # Just add to the warningcount when there's any warning
                if inconsistentDict["Mixed Column(s) with words & numbers in Source:"] != []:
                    warningCount += 1
                if inconsistentDict["Mixed Column(s) with words & numbers in Target:"] != []:
                    warningCount += 1

                # Create copies of dataframe to change first column into ColumnA; For pre-validation
                self.df.changed = self.df.copy() 
                self.df1.changed = self.df1.copy()

                # Let SQL perform the remaining warning handlings
                # sqlFormat(1) are strings to describe name & type of each column to create table with
                sqlFormat = ""
                sqlFormat1 = ""

                for col in self.df.changed.columns:
                    dt = self.df.changed[col].dtype
                    if dt == object: #FOR DATATYPE = OBJECT (aka string)
                        if sqlFormat == "":
                            self.df.changed.rename(columns={col: "ColumnA"}, inplace=True)
                        sqlFormat += col + " text, "
                    elif col in dateCol: #FOR DATATYPE = TIMESTAMP (aka date)
                        sqlFormat += col + " date, "
                    else:
                    # elif dt == int or dt == float: #FOR DATATYPE = INT (aka number)
                        sqlFormat += col + " real, "
                sqlFormat = sqlFormat[:-2]

                for col in self.df1.changed.columns:
                    dt = self.df1.changed[col].dtype
                    if dt == object: #FOR DATATYPE = OBJECT (aka string)
                        if sqlFormat1 == "":
                            self.df1.changed.rename(columns={col: "ColumnA"}, inplace=True)
                        sqlFormat1 += col + " text, "
                    elif col in dateCol1: #FOR DATATYPE = TIMESTAMP (aka date)
                        sqlFormat1 += col + " date, "
                    else:
                    # elif dt == int or dt == float: #FOR DATATYPE = INT (aka number)
                        sqlFormat1 += col + " real, "
                sqlFormat1 = sqlFormat1[:-2]

                # Connecting to existing / new database file
                conn = sqlite3.connect('Database\Autorecon.db')
                c = conn.cursor()
                c.execute("DROP TABLE IF EXISTS SOURCE")
                c.execute("DROP TABLE IF EXISTS TARGET")
                c.execute('''CREATE TABLE SOURCE
                            ({})'''.format(sqlFormat))
                c.execute('''CREATE TABLE TARGET
                            ({})'''.format(sqlFormat1))
                conn.commit()

                # Import dataset from self.df & self.df1 to SQL TABLE source & target
                self.df.changed.to_sql('SOURCE', conn, if_exists='replace', index=False)
                self.df1.changed.to_sql('TARGET', conn, if_exists='replace', index=False)
                logging.info(f"SOURCE & TARGET DATAFRAME CONVERTED TO DATABASE TABLE\n\n")

                # pragma table_info shows the name of each column & its datatype
                c.execute('''pragma table_info('SOURCE')''')
                colnamedtype = {}

                # Pairing the Column Name & its datatype together (E.g. 'ColumnA' : 'INTEGER')
                for row in c.fetchall():
                    colnamedtype[row[1]] = row[2]

                # Column that is TIMESTAMP (Both date & time) shall only have the date in the database
                for x in colnamedtype:
                    if colnamedtype[x] == "TIMESTAMP":
                        c.execute('''UPDATE SOURCE SET {} = DATE({});'''.format(x, x))

                c.execute('''pragma table_info('TARGET')''')
                colnamedtype1 = {}
                for row in c.fetchall():
                    colnamedtype1[row[1]] = row[2]
                for x in colnamedtype1:
                    if colnamedtype1[x] == "TIMESTAMP":
                        c.execute('''UPDATE TARGET SET {} = DATE({});'''.format(x, x))

                # Prevalidate with SQL script
                with open("SQL\prevalidation.sql", 'r') as s:
                    sql_script = s.read()
                    sqlcommands = sql_script.split(";")
                    sqlcommands.pop()
                    
                    try:
                        c.execute('''CREATE INDEX IDX_COLUMNA_SOURCE ON SOURCE(COLUMNA)''')
                        c.execute('''CREATE INDEX IDX_COLUMNA_TARGET ON TARGET(COLUMNA)''')
                        c.execute('''DROP TABLE IF EXISTS R_PREVALIDATION_OUTPUT_TBL''')
                        c.execute('''CREATE TABLE R_PREVALIDATION_OUTPUT_TBL (UNIQUE_ID INT, DESC TEXT, VAL TEXT)''')

                        for i, sqlcommand in enumerate(sqlcommands):
                            c.executescript(sqlcommand)
                            logging.info(f"PREVALIDATION CHECK #{i+1}: {sqlcommand}\n\n")
                        
                        conn.commit()
                        s.close()
                        c.execute('''SELECT DESC, VAL FROM R_PREVALIDATION_OUTPUT_TBL''')

                        workbook = Workbook('Warning Message\warningMessage.xlsx')
                        worksheet = workbook.add_worksheet(name="descriptions")
                        cell_format = workbook.add_format()
                        cell_format.set_bold()

                        # Add header
                        worksheet.write(0, 0, "Description", cell_format)
                        worksheet.write(0, 1, "Value", cell_format)
                        counter = 1
                        for row in c.fetchall():
                            worksheet.write_row(counter, 0, row)
                            warningCount += 1
                            counter += 1

                        for ab in inconsistentDict:
                            for x in inconsistentDict[ab]:
                                worksheet.write(counter, 0, ab)
                                worksheet.write(counter, 1, x)
                                counter += 1

                        workbook.close()

                    except:
                        logging.exception("Got exception when executing SQL")
                        raise

                t2 = time.time()

                # Any warning messages caught will be added to warningCount, hence 0 = no warning at all
                self.worker_loading.emit(2)
                self.worker_complete.emit({"df": self.df, "df1": self.df1, "warningCount": warningCount, "time": t2-t1, "src_changecolA": self.src_changecolA, "tgt_changecolA": self.tgt_changecolA})
        except:
            logging.exception("Got exception when running this function")
            raise

class LoadingScreen(QWidget):
    ### LOADING ANIMATION ON A WIDGET TO DISPLAY ON FRONTEND FOR USER'S FEEDBACK WHENEVER THERE IS A FUNCTION THAT TAKES TIME TO EXECUTE ###
    def __init__(self, parent):
        super().__init__()
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | self.windowFlags() | QtCore.Qt.Window | QtCore.Qt.WindowStaysOnTopHint) #QtCore.Qt.WindowStaysOnTopHint
        self.setFixedSize(200, 200)
        self.activateWindow()
        label = QLabel(self)
        self.movie = QMovie(r"icons\loading.gif")
        label.setMovie(self.movie)
        self.movie.start()
        self.show()

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        sys.exit(app.exec_())
    except:
        logging.exception('Exception')
        raise
    # app = QtWidgets.QApplication(sys.argv)
    # QMainWindow = QtWidgets.QMainWindow()
    # ui = Ui_MainWindow()
    # ui.setupUi(MainWindow)
    # QMainWindow.show()
