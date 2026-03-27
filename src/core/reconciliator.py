import re
import os
import sys
import time
import sqlite3
import datetime
import warnings
import traceback
import pandas as pd
from os import path
import win32com.client as w3c
from xlsxwriter.workbook import Workbook
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QColor, QMovie
from PyQt5.QtCore import QUrl, Qt
from PyQt5.QtWidgets import (QApplication, QFileDialog, QGraphicsDropShadowEffect,
                              QWidget, QMessageBox, QPushButton, QLabel, QMainWindow)

from config.settings import setup_logging, get_logger
from config.constants import (
    APP_NAME, APP_VERSION, REQUIRED_DIRS,
    DATABASE_PATH, ICON_PATHS,
)
from workers.threads import PrevalidationWorker

# ---- Application-level setup ----
logger = get_logger("reconciliator")

# Create all required directories
for dir_path in REQUIRED_DIRS.values():
    os.makedirs(dir_path, exist_ok=True)

warnings.filterwarnings("ignore")

GLOBAL_STATE = 0


def show_error_dialog(title, message):
    """Show a non-crashing error dialog to the user."""
    try:
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setIcon(QMessageBox.Critical)
        msg.setText(str(message))
        msg.addButton("Ok", QMessageBox.AcceptRole)
        msg.exec_()
    except Exception:
        logger.exception("Failed to show error dialog")


def global_exception_handler(exc_type, exc_value, exc_tb):
    """Global exception handler to prevent silent crashes."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))
    error_text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    show_error_dialog("Unexpected Error", f"An unexpected error occurred:\n\n{error_text}")

class MainWindow(QMainWindow):
    """Main application window."""
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        def moveWindow(event):
            try:
                if UIFunctions.returnStatus == 1:
                    UIFunctions.maximize_restore(self)
                if event.buttons() == Qt.LeftButton:
                    self.move(self.pos() + event.globalPos() - self.dragPos)
                    self.dragPos = event.globalPos()
                    event.accept()
            except Exception as e:
                logger.error(f"Error moving window: {e}")

        self.ui.name.mouseMoveEvent = moveWindow
        UIFunctions.uiDefinitions(self)
        self.show()

    def mousePressEvent(self, event):
        self.dragPos = event.globalPos()


class UIFunctions(MainWindow):
    """Window control functions (minimize, maximize, close)."""

    def maximize_restore(self):
        global GLOBAL_STATE
        status = GLOBAL_STATE
        try:
            if status == 0:
                self.showMaximized()
                GLOBAL_STATE = 1
                self.ui.body_layout.setContentsMargins(0, 0, 0, 0)
                self.ui.maximize_button.setIcon(self.ui.icon9)
            else:
                GLOBAL_STATE = 0
                self.showNormal()
                self.resize(self.width(), self.height())
                self.ui.maximize_button.setIcon(self.ui.icon2)
        except Exception as e:
            logger.error(f"Error toggling maximize: {e}")

    def uiDefinitions(self):
        try:
            self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
            self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
            self.shadow = QGraphicsDropShadowEffect(self)
            self.shadow.setBlurRadius(20)
            self.shadow.setXOffset(0)
            self.shadow.setYOffset(0)
            self.shadow.setColor(QColor(0, 0, 0, 100))
            self.ui.maximize_button.clicked.connect(lambda: UIFunctions.maximize_restore(self))
            self.ui.minimize_button.clicked.connect(lambda: self.showMinimized())
            self.ui.close_button.clicked.connect(lambda: self.close())
        except Exception as e:
            logger.error(f"Error setting up UI definitions: {e}")

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
        self.icon1.addPixmap(QtGui.QPixmap(ICON_PATHS["minus"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon2 = QtGui.QIcon()
        self.icon2.addPixmap(QtGui.QPixmap(ICON_PATHS["resize"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon3 = QtGui.QIcon()
        self.icon3.addPixmap(QtGui.QPixmap(ICON_PATHS["cancel"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon4 = QtGui.QIcon()
        self.icon4.addPixmap(QtGui.QPixmap(ICON_PATHS["next"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon5 = QtGui.QIcon()
        self.icon5.addPixmap(QtGui.QPixmap(ICON_PATHS["upload"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon6 = QtGui.QIcon()
        self.icon6.addPixmap(QtGui.QPixmap(ICON_PATHS["checklist"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon7 = QtGui.QIcon()
        self.icon7.addPixmap(QtGui.QPixmap(ICON_PATHS["delete"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon8 = QtGui.QIcon()
        self.icon8.addPixmap(QtGui.QPixmap(ICON_PATHS["reload"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon9 = QtGui.QIcon()
        self.icon9.addPixmap(QtGui.QPixmap(ICON_PATHS["restore"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon10 = QtGui.QIcon()
        self.icon10.addPixmap(QtGui.QPixmap(ICON_PATHS["back"]), QtGui.QIcon.Normal, QtGui.QIcon.Off)

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
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", APP_NAME))
        self.name.setText(_translate("MainWindow", APP_NAME.lower()))
        self.tutorial_tab.setText(_translate("MainWindow", "Read Me"))
        self.upload_tab.setText(_translate("MainWindow", "Upload Files"))
        self.rule_tab.setText(_translate("MainWindow", "Construct Rules"))
        self.recon_tab.setText(_translate("MainWindow", "Reconciliation"))
        self.version_label.setText(_translate("MainWindow", f"v{APP_VERSION}"))
        self.page_1_header_text.setText(_translate("MainWindow", "Read Me"))
        self.page_1_body_text.setText(_translate("MainWindow", "Take note of the pointers below before proceeding: \n\n"
                                                "1. Ensure your column names are only in alphanumeric and underscore, while date-related columns are read as 'Date' in excel. However, the output for date columns will be in 'YYYY-MM-DD' format for now.\n\n" 
                                                "2. Upload Source and Target excel files by clicking the respective buttons. Ensure that there are no duplicate column names present in each file and that both files should not be empty.\n\n"
                                                "3. Do note that column names may be amended to a more appropriate format (such as removing leading spaces and replacing '.' with '_'). Leading & Trailing spaces in each cell will be removed as well.\n\n"
                                                "4. You may encounter a scenario / error where the tool says similar terms such as 'Unnamed:_16', this means that the 17th column (in this case) is included even though it is blank. Please delete the column for fix.\n\n"
                                                "5. Once both files are uploaded, click on the 'Pre-Validation' button to validate the datasets for errors and warnings. Errors will be shown in console & you cannot proceed to the next step. Warnings will be stored into an excel file & you can be proceed upon acknowledging the message.\n\n"
                                                "6. Configuring pre-validation warning is flexible and users can add new SQL statement(s) into file 'prevalidation.sql' found in 'SQL' folder for program to execute.\n\n"
                                                "7. To construct the reconciliation rules, please select the columns & operators found in the dropdowns.\n\n"
                                                "8. After setting a desired rule, click on the 'Validate Rule' button to ensure that selected rule is error-free, which will be reflected below. In case of error, a popup will list down the reason(s).\n\n"
                                                "9. You can remove unwanted validated rule in the list by selecting the rule & clicking on the 'Delete' button.\n\n"
                                                "10. Click 'Reconcile' button to generate a reconciled excel file with the output, source, target and rules data.\n\n"
                                                "11. Click 'Restart' button after reconciliation output file has been generated to start the process all over again.\n\n"
                                                "12. In order to re-use the rules applied in the previous reconciled output file, keep the 'Rules' tab located in output file (and delete the remaining worksheets) as a single sheet. Then, in the next iteration, click on 'Rules' button to upload it.\n\n"
                                                "13. For better experience, please ensure that your Window Display Scaling is 125%.\n\n"
                                                "14. Since the first column for both excel files are the data drivers / primary keys, there is a possibility of cross join in the event these fields are not unique.\n\n"
                                                "15. Remove any blank column(s) or column(s) without any data."))
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
            self.reconcile()

        elif self.stacked_Widget.currentIndex() == 3:
            self.stacked_Widget.setCurrentIndex(1)
            self.upload_tab.setStyleSheet(rgb246)
            self.recon_tab.setStyleSheet(rgb250)
            self.restart()

    def loadingSpinner(self, state):
        """Start/stop loading spinner for background operations."""
        try:
            if state == 1:
                self._loading_screen = LoadingScreen(None)
            elif state == 2:
                if hasattr(self, '_loading_screen') and self._loading_screen:
                    self._loading_screen.movie.stop()
                    self._loading_screen.close()
                    self._loading_screen = None
        except Exception as e:
            logger.error(f"Error with loading spinner: {e}")

    # ---- File Upload (shared logic for source & target) ----

    def _upload_file(self, file_type):
        """Shared upload handler for both source and target files."""
        try:
            fileName = QFileDialog.getOpenFileName(
                None, 'Open File', '', "Excel files (*.xls *.xlsx)")
            if fileName == ('', ''):
                return

            file_path = fileName[0]
            splitList = re.split('/', file_path)
            start = time.time()
            df = pd.read_excel(file_path, header=0, sheet_name=0, engine='openpyxl')
            end = time.time()

            result = {
                "dataframe": df, "start": start, "end": end, "splitList": splitList
            }

            if file_type == "source":
                self._process_uploaded_file(result, is_source=True)
            else:
                self._process_uploaded_file(result, is_source=False)

        except Exception as e:
            logger.exception(f"Error uploading {file_type} file")
            show_error_dialog("Upload Error", f"Failed to upload {file_type} file:\n{e}")

    def uploadSource(self):
        self._upload_file("source")

    def uploadTarget(self):
        self._upload_file("target")

    def _process_uploaded_file(self, result, is_source=True):
        """Process an uploaded file (source or target) and update UI."""
        try:
            df = result["dataframe"]
            label = "Source" if is_source else "Target"

            if df.empty:
                self._append_output("File is empty, please upload again.")
                return

            # Set filename in UI
            fname = result["splitList"][-1]
            name_widget = self.source_file_name if is_source else self.target_file_name
            name_widget.setText(fname)
            name_widget.setToolTip(fname)
            name_widget.adjustSize()

            # Handle fully-null columns — ask user for datatype
            for col in df.columns:
                if df[col].isnull().all():
                    reply = self._dtype_popup(col)
                    if reply == 0:
                        df[col] = df[col].fillna(0)
                        self._append_output(f"{col} is converted to Numeric!")
                    elif reply == 1:
                        df[col] = df[col].fillna('')
                        self._append_output(f"{col} is converted to Text!")
                    else:
                        df[col] = df[col].fillna(datetime.date(2099, 1, 1))
                        df[col] = df[col].astype('datetime64[ns]')
                        self._append_output(f"{col} is converted to Date!")

            # Sanitize column names (remove special chars, replace dots/spaces with _)
            change_msg = ""
            before = str(df.columns.values)
            cleaned_cols = []
            for col_name in df.columns:
                col_name = re.sub(r"^\W+|^\d+|^ +", "", col_name)
                col_name = re.sub(r"\.+|\s+", "_", col_name)
                cleaned_cols.append(col_name)
            df.columns = cleaned_cols

            if before != str(df.columns.values):
                change_msg = "Please take note of changes in column names:\n" + "\n".join(df.columns.values)

            # Store dataframe
            if is_source:
                self.df = df
            else:
                self.df1 = df

            duration = round(result["end"] - result["start"], 2)

            # Check if both files are uploaded
            both_ready = self._both_files_ready()
            if both_ready:
                self._append_output(
                    f"{change_msg}\nEnsure selected files are correct before clicking on Pre-Validate.\n"
                    f"Duration to read {label} file: {duration}s")
                if is_source:
                    self.upload_source_button.setDisabled(True)
                else:
                    self.upload_target_button.setDisabled(True)
                self.prevalidate_button.setEnabled(True)
            else:
                self._append_output(
                    f"{change_msg}\nDuration taken to upload {label} file: {duration}s. "
                    f"Please upload {'target' if is_source else 'source'} file.")

        except Exception as e:
            logger.exception("Error processing uploaded file")
            show_error_dialog("Processing Error", f"Error processing file:\n{e}")

    def _both_files_ready(self):
        """Check if both source and target dataframes are loaded."""
        try:
            return (hasattr(self, 'df') and self.df is not None and not self.df.empty
                    and hasattr(self, 'df1') and self.df1 is not None and not self.df1.empty)
        except AttributeError:
            return False

    def _append_output(self, message):
        """Append a timestamped message to the output viewer."""
        try:
            now = datetime.datetime.now().strftime("%I:%M %p")
            current = self.page_2_output_list.text().lstrip('\n').lstrip(' ')
            self.page_2_output_list.setText(f"{current}{now}: {message}\n\n")
            self.page_2_scrollArea.verticalScrollBar().setValue(
                self.page_2_scrollArea.verticalScrollBar().maximum())
        except Exception as e:
            logger.error(f"Error appending output: {e}")

    # ---- Prevalidation ----

    def prevalidate(self):
        """Start prevalidation in a worker thread."""
        try:
            self.worker = PrevalidationWorker(self.df, self.df1)
            self.worker.start()
            self.worker.worker_loading.connect(self.loadingSpinner)
            self.worker.worker_complete.connect(self.prevalidated)
        except Exception as e:
            logger.exception("Error starting prevalidation")
            show_error_dialog("Prevalidation Error", f"Failed to start prevalidation:\n{e}")

    def prevalidated(self, result):
        """Handle prevalidation results from worker thread."""
        try:
            self.df = result["df"]
            self.df1 = result["df1"]
            self.src_changecolA = result["src_changecolA"]
            self.tgt_changecolA = result["tgt_changecolA"]
            duration = round(result["time"], 2)

            if "errorMessage" in result:
                self._append_output(f"{result['errorMessage']}\n\nDuration to handle errors: {duration}s")
            else:
                if result["warningCount"] != 0:
                    reply = self._prevalidate_popup()
                    if reply == 0:
                        self._append_output(
                            f"Pre-validation pass successfully with the acknowledgement of prompted warnings "
                            f"& no errors. Please continue by clicking Next.\n\nDuration to handle warnings: {duration}s")
                        self.page_2_next_button.setEnabled(True)
                        self._populate_dropdowns()
                    else:
                        self._append_output(
                            "You have selected 'Not to Proceed Further' due to warnings, "
                            "please pre-validate again if you want to proceed with the same files. "
                            "Or else, please fix and re-upload files again.")
                else:
                    self._append_output(
                        f"Pre-validation pass successfully. No errors/warnings\n\n"
                        f"Duration to handle warnings: {duration}s")
                    self.page_2_next_button.setEnabled(True)
                    self._populate_dropdowns()

        except Exception as e:
            logger.exception("Error handling prevalidation results")
            show_error_dialog("Prevalidation Error", f"Error processing results:\n{e}")

    def _populate_dropdowns(self):
        """Populate column/operator dropdowns after successful prevalidation."""
        try:
            conn = sqlite3.connect(DATABASE_PATH)
            c = conn.cursor()

            self.srcCol = {}
            self.tgtCol = {}

            self.source_column_dropdown.clear()
            self.target_column_dropdown.clear()
            self.operator_dropdown.clear()

            c.execute("PRAGMA table_info('SOURCE')")
            for col, row in zip(self.df.columns, c.fetchall()):
                self.srcCol[col] = row[2]
                self.source_column_dropdown.addItem("src." + str(col))

            c.execute("PRAGMA table_info('TARGET')")
            for col, row in zip(self.df1.columns, c.fetchall()):
                self.tgtCol[col] = row[2]
                self.target_column_dropdown.addItem("tgt." + str(col))

            self.comparator = ["=", "!=", "<", "<=", ">", ">=", "+", "-"]
            for op in self.comparator:
                self.operator_dropdown.addItem(op)

            self.source_column_dropdown.setCurrentIndex(-1)
            self.target_column_dropdown.setCurrentIndex(-1)
            self.operator_dropdown.setCurrentIndex(-1)
            self.construct_rule_browser.clear()
            conn.close()

        except Exception as e:
            logger.exception("Error populating dropdowns")
            show_error_dialog("Dropdown Error", f"Failed to populate dropdowns:\n{e}")

    # ---- Dropdown concat ----

    def concatSourceDD(self, value):
        try:
            self.construct_rule_browser.setText(self.construct_rule_browser.toPlainText() + value + " ")
            self.source_column_dropdown.setCurrentIndex(-1)
        except Exception as e:
            logger.error(f"Error in source dropdown: {e}")

    def concatTargetDD(self, value):
        try:
            self.construct_rule_browser.setText(self.construct_rule_browser.toPlainText() + value + " ")
            self.target_column_dropdown.setCurrentIndex(-1)
        except Exception as e:
            logger.error(f"Error in target dropdown: {e}")

    def concatOpDD(self, value):
        try:
            self.construct_rule_browser.setText(self.construct_rule_browser.toPlainText() + value + " ")
            self.operator_dropdown.setCurrentIndex(-1)
        except Exception as e:
            logger.error(f"Error in operator dropdown: {e}")

    def clearConstructingRule(self):
        try:
            self.construct_rule_browser.clear()
        except Exception as e:
            logger.error(f"Error clearing rule: {e}")

    # ---- Rule Validation ----

    def ruleValidation(self):
        """Validate the constructed rule for syntax and datatype compatibility."""
        try:
            ruleBlocks = re.split(' ', self.construct_rule_browser.toPlainText())
            while '' in ruleBlocks:
                ruleBlocks.remove('')

            errorMessage = []

            # Syntax: need at least 3 parts
            if len(ruleBlocks) < 3:
                errorMessage.append('Rule needs to have at least two columns with operator for comparison.')

            for x in ruleBlocks[1::2]:
                if x not in self.comparator:
                    errorMessage.append('Operator should be placed in even index position.')

            for x in ruleBlocks[::2]:
                if x in self.comparator:
                    errorMessage.append('Column should be placed in odd index position.')

            if ruleBlocks and ruleBlocks[-1] in self.comparator:
                errorMessage.append('Operator should not be at the end of the rule.')

            # Datatype compatibility
            dtypeList = []
            for x in ruleBlocks[::2]:
                parts = x.split('.')
                if len(parts) >= 2:
                    if parts[0] == 'src':
                        dtypeList.append(self.srcCol.get(parts[1], ''))
                    elif parts[0] == 'tgt':
                        dtypeList.append(self.tgtCol.get(parts[1], ''))

            dtypeList = list(dict.fromkeys(dtypeList))
            if len(dtypeList) > 1 and sorted(dtypeList) != ['INTEGER', 'REAL']:
                errorMessage.append('Unable to compare different datatypes.')

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
                        errorMessage.append('Date-type rules applicable only to operators that return True/False.')

            else:  # INTEGER or REAL
                counter = sum(1 for x in ruleBlocks[1::2] if x in ["=", "!=", '<', "<=", ">", ">="])
                if counter > 1:
                    errorMessage.append('Integer-type rules can only use 1 operator that returns True/False, rest should be "+" or "-".')

            if not errorMessage:
                self._rule_appending()
            else:
                self.construct_rule_browser.clear()
                msg = QMessageBox()
                msg.setWindowTitle("Rule Validation Error")
                msg.setText("\n".join(errorMessage))
                msg.addButton("Ok", QMessageBox.YesRole)
                msg.exec_()

        except Exception as e:
            logger.exception("Error validating rule")
            show_error_dialog("Validation Error", f"Error validating rule:\n{e}")

    def _rule_appending(self):
        """Add validated rule to the list widget."""
        try:
            text = self.construct_rule_browser.toPlainText().rstrip().replace("  ", " ")
            self.listWidget.addItem(text)
            self.reconcile_button.setEnabled(True)
            self.construct_rule_browser.clear()
        except Exception:
            logger.exception("Error appending rule")

    def ruleDeleting(self):
        """Delete selected rule from the list widget."""
        try:
            self.listWidget.takeItem(self.listWidget.currentRow())
        except Exception:
            logger.exception("Error deleting rule")

    def uploadRules(self):
        """Upload rules from a previous reconciliation output file."""
        try:
            rfName = QFileDialog.getOpenFileName(
                None, 'Open File', '', "Excel files (*.xls *.xlsx)")
            if rfName == ('', ''):
                return

            ruleBank = pd.read_excel(rfName[0], header=0, sheet_name="RULES", engine='openpyxl')

            # Add prefix to column keys
            srcColPrefixed = {f"src.{k}": v for k, v in self.srcCol.items()}
            tgtColPrefixed = {f"tgt.{k}": v for k, v in self.tgtCol.items()}
            colList = {**srcColPrefixed, **tgtColPrefixed}

            columnDontExist = ""
            differentDtype = ""
            wrongSyntax = ""
            acceptedRules = {}

            for x in range(len(ruleBank)):
                count = 0
                rule_row = ruleBank.iloc[x].values.tolist()
                rule_row = [str(v).replace("'", "") for v in rule_row[1:]]
                rule_row = [v for v in rule_row if v and v != 'nan']

                # Syntax validation
                if len(rule_row) < 3:
                    wrongSyntax += f'Uploaded Rule #{x+1}: Does not fit criteria of at least two columns and an operator.\n'
                    count += 1

                for y in rule_row[1::2]:
                    if y not in self.comparator:
                        wrongSyntax += f'Uploaded Rule #{x+1}: Operator should be placed in even index position.\n'
                        count += 1

                for y in rule_row[::2]:
                    if y in self.comparator:
                        wrongSyntax += f'Uploaded Rule #{x+1}: Column should be placed in odd index position.\n'
                        count += 1

                # Column existence
                ruleBankwoComp = [v for v in rule_row if v not in self.comparator]
                for y in ruleBankwoComp:
                    if y not in colList:
                        columnDontExist += f'Uploaded Rule #{x+1}: {y} does not exist.\n'
                        count += 1

                # Datatype validation (only if all columns exist)
                if count == 0:
                    dtypeList = []
                    for y in rule_row[::2]:
                        if y in colList:
                            dtypeList.append(colList[y])
                    dtypeList = list(dict.fromkeys(dtypeList))

                    if len(dtypeList) > 1:
                        differentDtype += f'Uploaded Rule #{x+1}: Unable to compare different datatypes.\n'
                        count += 1

                    if dtypeList == ["TEXT"]:
                        if len(rule_row[1::2]) > 1:
                            differentDtype += f'Uploaded Rule #{x+1}: Text-type rules can only have 1 operator.\n'
                            count += 1
                        for y in rule_row[1::2]:
                            if y not in ["=", "!="]:
                                differentDtype += f'Uploaded Rule #{x+1}: Text-type rules applicable only to "=" & "!=".\n'
                                count += 1

                    elif dtypeList == ["TIMESTAMP"]:
                        if len(rule_row[1::2]) > 1:
                            differentDtype += f'Uploaded Rule #{x+1}: Date-type rules can only have 1 operator.\n'
                            count += 1
                        for y in rule_row[1::2]:
                            if y not in ["=", "!=", '<', "<=", ">", ">="]:
                                differentDtype += f'Uploaded Rule #{x+1}: Date-type rules applicable only to True/False operators.\n'
                                count += 1
                    else:
                        counter = sum(1 for y in rule_row[1::2] if y in ["=", "!=", '<', "<=", ">", ">="])
                        if counter > 1:
                            differentDtype += f'Uploaded Rule #{x+1}: Integer-type rules can only use 1 True/False operator.\n'
                            count += 1

                if count == 0:
                    acceptedRules[x+1] = " ".join(rule_row)

            for key in acceptedRules:
                self.listWidget.addItem(acceptedRules[key])

            all_errors = columnDontExist + differentDtype + wrongSyntax
            if all_errors:
                msg = QMessageBox()
                msg.setWindowTitle("Uploaded Rule(s) Validation Error")
                msg.setText(all_errors)
                msg.addButton("Ok", QMessageBox.YesRole)
                msg.exec_()

            if acceptedRules:
                self.delete_rule_button.setEnabled(True)
                self.reconcile_button.setEnabled(True)

        except Exception as e:
            logger.exception("Error uploading rules")
            show_error_dialog("Upload Rules Error", f"Failed to upload rules:\n{e}")

    # ---- Reconciliation ----

    def reconcile(self):
        """Generate the reconciliation output Excel file."""
        conn = None
        try:
            progress = 0
            total = 15
            self.progressBar.setValue(0)

            t1 = time.time()
            output_path = os.path.join(os.getcwd(), 'reconciliator.xlsx')
            if path.exists(output_path):
                try:
                    xl = w3c.Dispatch('Excel.Application')
                    wb = xl.Workbooks.Open(output_path.replace("\\", "/"))
                    wb.Close(True)
                except Exception as e:
                    logger.warning(f"Could not close existing Excel file: {e}")

            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            # Collect rules from list widget
            rules = [self.listWidget.item(i).text() for i in range(self.listWidget.count())]
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            # Replace first column references with ColumnA placeholder
            items = [w.replace(f"src.{self.src_changecolA}", "src.ColumnA") for w in rules]
            items = [w.replace(f"tgt.{self.tgt_changecolA}", "tgt.ColumnA") for w in items]
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            # Build rule dictionary
            ruleDict = {}
            for i, item in enumerate(items, start=1):
                ruleDict[str(i)] = item.split(" ")
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            # Build SQL statement
            createtable = "CREATE TABLE BIG_TABLE AS\n"
            selectStatement = "SELECT DISTINCT src.ColumnA src_ColumnA, tgt.ColumnA tgt_ColumnA, "
            count = 1
            comparison_ops = ['<', '<=', '=', '!=', '>', '>=']

            for key in ruleDict:
                for value in ruleDict[key]:
                    if "." in value:
                        z = value.replace(".", "_")
                        selectStatement += f"{value} R{count}_{z}, "

                result = [i for i in ruleDict[key] if i in comparison_ops]
                if result:
                    selectStatement += "CASE WHEN "
                    for value in ruleDict[key]:
                        if value not in comparison_ops + ['+', '-']:
                            selectStatement += f'IFNULL({value},"") '
                        else:
                            selectStatement += f"{value} "
                    selectStatement += f"THEN 'YES' ELSE 'NO' END '{' '.join(ruleDict[key])}', "
                else:
                    for value in ruleDict[key]:
                        if value not in comparison_ops + ['+', '-']:
                            selectStatement += f"IFNULL({value}, 0) "
                        else:
                            selectStatement += f"{value} "
                    selectStatement += f"'{' '.join(ruleDict[key])}', "
                count += 1
            selectStatement = selectStatement[:-2]

            SQLStatement = (
                f"{createtable}{selectStatement}\n"
                f"FROM source src LEFT JOIN target tgt USING(ColumnA)\n"
                f"UNION ALL\n{selectStatement}\n"
                f"FROM target tgt LEFT JOIN source src USING(ColumnA) WHERE src.ColumnA IS NULL"
            )
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            logger.info(f"OUTPUT SQL QUERY\n{SQLStatement}")

            conn = sqlite3.connect(DATABASE_PATH)
            c = conn.cursor()
            c.execute("DROP TABLE IF EXISTS R_RULES_TBL")

            # Create rules table with dynamic column count
            max_rule_len = len(ruleDict[max(ruleDict, key=lambda k: len(ruleDict[k]))])
            createRuleTableStr = "CREATE TABLE R_RULES_TBL (SN REAL, "
            for x in range(max_rule_len):
                createRuleTableStr += f"VAL{x + 1} TEXT, "
            createRuleTableStr = createRuleTableStr[:-2] + ");"
            c.executescript(createRuleTableStr)
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            conn.commit()
            abc = [f"VAL{x + 1}" for x in range(max_rule_len)]
            insertRuleTableStr = f"INSERT INTO R_RULES_TBL (SN, {', '.join(abc)}) VALUES"
            for i, rule in enumerate(rules):
                parts = rule.split(" ")
                insertRuleTableStr += f" ({i + 1}, '{rule.replace(' ', chr(39) + ', ' + chr(39))}"
                if len(parts) < max_rule_len:
                    insertRuleTableStr += "'"
                    for _ in range(max_rule_len - len(parts)):
                        insertRuleTableStr += ", ''"
                    insertRuleTableStr += "),"
                else:
                    insertRuleTableStr += "'),"
            insertRuleTableStr = insertRuleTableStr[:-1]
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            logger.info(f"INSERT RULE TABLE QUERY\n{insertRuleTableStr}")
            c.executescript(insertRuleTableStr)
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            c.execute("DROP TABLE IF EXISTS BIG_TABLE")
            c.execute(SQLStatement)

            # Create output workbook
            workbook = Workbook(output_path)
            worksheet = workbook.add_worksheet(name="RECON_OUTPUT")
            worksheet1 = workbook.add_worksheet(name="SOURCE")
            worksheet2 = workbook.add_worksheet(name="TARGET")
            worksheet3 = workbook.add_worksheet(name="RULES")
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            # Write output sheet
            c.execute("PRAGMA table_info('BIG_TABLE')")
            colList = [row[1] for row in c.fetchall()]
            colList = [v.replace("src.ColumnA", f"src.{self.src_changecolA}") for v in colList]
            colList = [v.replace("tgt.ColumnA", f"tgt.{self.tgt_changecolA}") for v in colList]
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            worksheet.write(0, 0, f"src_{self.src_changecolA}")
            worksheet.write(0, 1, f"tgt_{self.tgt_changecolA}")
            worksheet.write_row(0, 2, colList[2:])
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            c.execute("SELECT * FROM BIG_TABLE")
            for row_num, row_data in enumerate(c.fetchall(), start=1):
                worksheet.write_row(row_num, 0, list(row_data))
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            # Write source sheet
            c.execute("PRAGMA table_info('SOURCE')")
            colList = [row[1] for row in c.fetchall()]
            worksheet1.write(0, 0, self.src_changecolA)
            worksheet1.write_row(0, 1, colList[1:])
            c.execute("SELECT * FROM source")
            for row_num, row_data in enumerate(c.fetchall(), start=1):
                worksheet1.write_row(row_num, 0, list(row_data))
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            # Write target sheet
            c.execute("PRAGMA table_info('TARGET')")
            colList = [row[1] for row in c.fetchall()]
            worksheet2.write(0, 0, self.tgt_changecolA)
            worksheet2.write_row(0, 1, colList[1:])
            c.execute("SELECT * FROM target")
            for row_num, row_data in enumerate(c.fetchall(), start=1):
                worksheet2.write_row(row_num, 0, list(row_data))
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            # Write rules sheet
            c.execute("PRAGMA table_info('R_RULES_TBL')")
            colList = [row[1] for row in c.fetchall()]
            worksheet3.write_row(0, 0, colList)
            c.execute("SELECT * FROM R_RULES_TBL")
            for i, row_data in enumerate(c.fetchall()):
                worksheet3.write_row(i + 1, 0, [f"'{y}" for y in row_data])
            workbook.close()
            progress += 1
            self.progressBar.setValue(int((progress / total) * 100))

            self.restart_button.setEnabled(True)
            self.progress_bar_text.setText("Done!")
            t2 = time.time()
            self.listView.setText(
                f"Reconciliation is successful! The output file is stored in the same directory as this executable file.\n\n"
                f"Summary:\nSource file: {self.source_file_name.text()}\n"
                f"Target file: {self.target_file_name.text()}\n"
                f"Rule(s):\n{chr(10).join(rules)}\n"
                f"Output file: reconciliator.xlsx\n\n"
                f"Session ended...\nDuration to reconcile: {round(t2 - t1, 2)}s")

            # Fix ColumnA references in rule display
            for i in range(self.listWidget.count()):
                read = self.listWidget.item(i).text().split(" ")
                for idx in [0, 2] if len(read) > 2 else [0]:
                    if idx < len(read):
                        if read[idx] == "src.ColumnA":
                            read[idx] = f"src.{self.src_changecolA}"
                        if read[idx] == "tgt.ColumnA":
                            read[idx] = f"tgt.{self.tgt_changecolA}"
                self.listWidget.takeItem(i)
                self.listWidget.addItem(" ".join(read))

            logger.info(f"Reconciliation completed in {round(t2 - t1, 2)}s")

        except Exception as e:
            logger.exception("Error during reconciliation")
            show_error_dialog("Reconciliation Error", f"An error occurred during reconciliation:\n{e}")
        finally:
            if conn:
                conn.close()

    def restart(self):
        """Reset the application for a new reconciliation run."""
        try:
            self.stacked_Widget.setCurrentIndex(1)
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

            if not self.checkBox.isChecked() and not self.checkBox_2.isChecked():
                self.prevalidate_button.setEnabled(True)
            if self.checkBox.isChecked():
                self.df = None
                self.source_file_name.clear()
            if self.checkBox_2.isChecked():
                self.df1 = None
                self.target_file_name.clear()

            logger.info("Application restarted for new session")

        except Exception as e:
            logger.exception("Error during restart")
            show_error_dialog("Restart Error", f"Error during restart:\n{e}")

    # ---- Popup Dialogs ----

    def _dtype_popup(self, column_name):
        """Ask user to specify datatype for an empty column."""
        msg = QMessageBox()
        msg.setText(f"Column '{column_name}' is empty. Please state the data type.")
        msg.addButton(QPushButton("Numeric"), QMessageBox.YesRole)
        msg.addButton(QPushButton("Text"), QMessageBox.NoRole)
        msg.addButton(QPushButton("Date"), QMessageBox.RejectRole)
        msg.setWindowTitle("Ambiguous Datatype")
        msg.setIcon(QMessageBox.Question)
        return msg.exec_()

    def _prevalidate_popup(self):
        """Show warning popup during prevalidation with link to details."""
        msg = QMessageBox()
        msg.setWindowTitle("Warning Message")
        warning_path = os.path.join(REQUIRED_DIRS["warnings"], "warningMessage.xlsx")
        url = bytearray(QUrl.fromLocalFile(warning_path).toEncoded()).decode()
        msg.setText(f'<a href="{url}">Click here for warning details;</a> Do you still want to proceed?')
        msg.addButton(QPushButton("Yes"), QMessageBox.YesRole)
        msg.addButton(QPushButton("No"), QMessageBox.NoRole)
        msg.setIcon(QMessageBox.Question)
        return msg.exec_()


class LoadingScreen(QWidget):
    """Loading animation overlay shown during long-running operations."""
    def __init__(self, parent):
        super().__init__()
        try:
            self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
            self.setWindowFlags(
                QtCore.Qt.FramelessWindowHint | QtCore.Qt.Window | QtCore.Qt.WindowStaysOnTopHint)
            self.setFixedSize(200, 200)
            self.activateWindow()
            label = QLabel(self)
            self.movie = QMovie(ICON_PATHS["loading"])
            label.setMovie(self.movie)
            self.movie.start()
            self.show()
        except Exception as e:
            logger.error(f"Error creating loading screen: {e}")


if __name__ == "__main__":
    # Install global exception handler to prevent silent crashes
    sys.excepthook = global_exception_handler

    # Set up logging
    setup_logging()
    logger.info(f"Starting {APP_NAME} v{APP_VERSION}")

    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        sys.exit(app.exec_())
    except Exception as e:
        logger.critical(f"Fatal application error: {e}", exc_info=True)
        show_error_dialog("Fatal Error", f"Application failed to start:\n{e}")
        sys.exit(1)
