#! python3
import datetime
import os
import sys
import time
from typing import Any, Tuple

import pandas as pd
import PyQt5
from googleapiclient.discovery import Resource, build
from oauth2client.service_account import ServiceAccountCredentials
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

if hasattr(QtCore.Qt, "AA_EnableHighDpiScaling"):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, "AA_UseHighDpiPixmaps"):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)


# Main Window for GUI
class Ui_mainWindow(object):
    saveReportToPath = ""
    GBCScheduleOutput = False
    SFCScheduleOutput = False
    uploadGBCSchedule = True
    GBCGDriveFolderId = ""
    uploadSFCSchedule = True
    SFCGDriveFolderId = ""
    attachGBCSchedule = True
    GBCCalendarId = ""
    attachSFCSchedule = True
    SFCCalendarId = ""
    center = {
        "Golden Bear Center": {
            "campus": "Berkeley - CA0001",
            "building": (
                "UC Berkeley Extension Golden Bear Center, 1995 University Ave. - GBC"
            ),
        },
        "San Francisco Center": {
            "campus": "San Francisco - CA0003",
            "building": "San Francisco Campus, 160 Spear St. - SFCAMPUS",
        },
    }

    centerReverse = {
        "GBC - UC Berkeley Extension Golden Bear Center, 1995 University Ave.": {
            "name": "GBC"
        },
        "SFCAMPUS - San Francisco Campus, 160 Spear St.": {"name": "SFC"},
    }

    def __init__(self) -> None:
        super().__init__()

        # Initialize settings from config.ini file, otherwise set default
        self.settings = QtCore.QSettings("config.ini", QtCore.QSettings.IniFormat)
        self.saveReportToPath = self.settings.value(
            "saveReportToPath", os.path.dirname(os.path.abspath(__file__)), type=str
        )
        self.GBCScheduleOutput = self.settings.value(
            "GBCScheduleOutput", False, type=bool
        )
        self.SFCScheduleOutput = self.settings.value(
            "SFCScheduleOutput", False, type=bool
        )
        self.uploadGBCSchedule = self.settings.value(
            "uploadGBCSchedule", True, type=bool
        )
        self.GBCGDriveFolderId = self.settings.value("GBCGDriveFolderId", "", type=str)
        self.uploadSFCSchedule = self.settings.value(
            "uploadSFCSchedule", True, type=bool
        )
        self.SFCGDriveFolderId = self.settings.value("SFCGDriveFolderId", "", type=str)
        self.attachGBCSchedule = self.settings.value(
            "attachGBCSchedule", True, type=bool
        )
        self.GBCCalendarId = self.settings.value("GBCCalendarId", "", type=str)
        self.attachSFCSchedule = self.settings.value(
            "attachSFCSchedule", True, type=bool
        )
        self.SFCCalendarId = self.settings.value("SFCCalendarId", "", type=str)

    def setupUi(self, mainWindow: QtWidgets.QWidget) -> None:
        mainWindow.setObjectName("mainWindow")
        mainWindow.setWindowModality(QtCore.Qt.NonModal)
        mainWindow.setEnabled(True)
        mainWindow.resize(547, 222)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Maximum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(mainWindow.sizePolicy().hasHeightForWidth())
        mainWindow.setSizePolicy(sizePolicy)
        mainWindow.setBaseSize(QtCore.QSize(430, 400))
        self.mainWindowLayout = QtWidgets.QVBoxLayout(mainWindow)
        self.mainWindowLayout.setObjectName("mainWindowLayout")
        self.genReportBox = QtWidgets.QGroupBox(mainWindow)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.MinimumExpanding
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.genReportBox.sizePolicy().hasHeightForWidth())
        self.genReportBox.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.genReportBox.setFont(font)
        self.genReportBox.setCheckable(False)
        self.genReportBox.setChecked(False)
        self.genReportBox.setObjectName("genReportBox")
        self.genReportLayout = QtWidgets.QVBoxLayout(self.genReportBox)
        self.genReportLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.genReportLayout.setObjectName("genReportLayout")
        self.dateLayout = QtWidgets.QHBoxLayout()
        self.dateLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.dateLayout.setObjectName("dateLayout")
        self.startDateLabel = QtWidgets.QLabel(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.startDateLabel.sizePolicy().hasHeightForWidth()
        )
        self.startDateLabel.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.startDateLabel.setFont(font)
        self.startDateLabel.setObjectName("startDateLabel")
        self.dateLayout.addWidget(self.startDateLabel)
        self.selectStartDate = QtWidgets.QDateEdit(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.selectStartDate.sizePolicy().hasHeightForWidth()
        )
        self.selectStartDate.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.selectStartDate.setFont(font)
        self.selectStartDate.setFrame(True)
        self.selectStartDate.setReadOnly(False)
        self.selectStartDate.setProperty("showGroupSeparator", False)
        self.selectStartDate.setCalendarPopup(True)
        self.selectStartDate.setDate(QtCore.QDate.currentDate())
        self.startDate = str(QtCore.QDate.currentDate().toPyDate())
        self.selectStartDate.setObjectName("selectStartDate")
        self.dateLayout.addWidget(self.selectStartDate)
        self.endDateLabel = QtWidgets.QLabel(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.endDateLabel.sizePolicy().hasHeightForWidth())
        self.endDateLabel.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.endDateLabel.setFont(font)
        self.endDateLabel.setObjectName("endDateLabel")
        self.dateLayout.addWidget(self.endDateLabel)
        self.selectEndDate = QtWidgets.QDateEdit(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.selectEndDate.sizePolicy().hasHeightForWidth()
        )
        self.selectEndDate.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.selectEndDate.setFont(font)
        self.selectEndDate.setFocusPolicy(QtCore.Qt.WheelFocus)
        self.selectEndDate.setReadOnly(False)
        self.selectEndDate.setCalendarPopup(True)
        self.selectEndDate.setDate(QtCore.QDate.currentDate().addMonths(12))
        self.endDate = str(QtCore.QDate.currentDate().addMonths(12).toPyDate())
        self.selectEndDate.setObjectName("selectEndDate")
        self.dateLayout.addWidget(self.selectEndDate)
        self.genReportLayout.addLayout(self.dateLayout)
        self.saveReportPathLayout = QtWidgets.QHBoxLayout()
        self.saveReportPathLayout.setObjectName("saveReportPathLayout")
        self.saveReportPathLabel = QtWidgets.QLabel(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.saveReportPathLabel.sizePolicy().hasHeightForWidth()
        )
        self.saveReportPathLabel.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.saveReportPathLabel.setFont(font)
        self.saveReportPathLabel.setObjectName("saveReportPathLabel")
        self.saveReportPathLayout.addWidget(self.saveReportPathLabel)
        self.selectSaveReportPath = QtWidgets.QLineEdit(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.selectSaveReportPath.sizePolicy().hasHeightForWidth()
        )
        self.selectSaveReportPath.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.selectSaveReportPath.setFont(font)
        self.selectSaveReportPath.setReadOnly(True)
        self.selectSaveReportPath.setObjectName("selectSaveReportPath")
        self.saveReportPathLayout.addWidget(self.selectSaveReportPath)
        self.browseSaveReportButton = QtWidgets.QToolButton(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.browseSaveReportButton.sizePolicy().hasHeightForWidth()
        )
        self.browseSaveReportButton.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.browseSaveReportButton.setFont(font)
        self.browseSaveReportButton.setCheckable(False)
        self.browseSaveReportButton.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.browseSaveReportButton.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        self.browseSaveReportButton.setObjectName("browseSaveReportButton")
        self.saveReportPathLayout.addWidget(self.browseSaveReportButton)
        self.genReportLayout.addLayout(self.saveReportPathLayout)
        self.locationLayout = QtWidgets.QHBoxLayout()
        self.locationLayout.setObjectName("locationLayout")
        self.label = QtWidgets.QLabel(self.genReportBox)
        self.label.setObjectName("label")
        self.locationLayout.addWidget(self.label)
        self.GBCcheckBox = QtWidgets.QCheckBox(self.genReportBox)
        font = QtGui.QFont()
        font.setUnderline(False)
        self.GBCcheckBox.setFont(font)
        self.GBCcheckBox.setChecked(self.GBCScheduleOutput)
        self.GBCcheckBox.setObjectName("GBCcheckBox")
        self.locationLayout.addWidget(self.GBCcheckBox)
        spacerItem = QtWidgets.QSpacerItem(
            40, 20, QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum
        )
        self.locationLayout.addItem(spacerItem)
        self.SFCcheckBox = QtWidgets.QCheckBox(self.genReportBox)
        font = QtGui.QFont()
        font.setUnderline(False)
        self.SFCcheckBox.setFont(font)
        self.SFCcheckBox.setChecked(self.SFCScheduleOutput)
        self.SFCcheckBox.setObjectName("SFCcheckBox")
        self.locationLayout.addWidget(self.SFCcheckBox)
        spacerItem1 = QtWidgets.QSpacerItem(
            40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum
        )
        self.locationLayout.addItem(spacerItem1)
        self.genReportLayout.addLayout(self.locationLayout)
        self.mainWindowLayout.addWidget(self.genReportBox)
        self.startExitLayout = QtWidgets.QHBoxLayout()
        self.startExitLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.startExitLayout.setObjectName("startExitLayout")
        spacerItem2 = QtWidgets.QSpacerItem(
            40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum
        )
        self.startExitLayout.addItem(spacerItem2)
        self.StartButton = QtWidgets.QPushButton(mainWindow)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.StartButton.sizePolicy().hasHeightForWidth())
        self.StartButton.setSizePolicy(sizePolicy)
        self.StartButton.setMinimumSize(QtCore.QSize(175, 50))
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.StartButton.setFont(font)
        self.StartButton.setObjectName("StartButton")
        self.startExitLayout.addWidget(self.StartButton)
        spacerItem3 = QtWidgets.QSpacerItem(
            40,
            20,
            QtWidgets.QSizePolicy.MinimumExpanding,
            QtWidgets.QSizePolicy.Minimum,
        )
        self.startExitLayout.addItem(spacerItem3)
        self.exitButton = QtWidgets.QPushButton(mainWindow)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.exitButton.sizePolicy().hasHeightForWidth())
        self.exitButton.setSizePolicy(sizePolicy)
        self.exitButton.setMinimumSize(QtCore.QSize(175, 50))
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.exitButton.setFont(font)
        self.exitButton.setObjectName("exitButton")
        self.startExitLayout.addWidget(self.exitButton)
        spacerItem4 = QtWidgets.QSpacerItem(
            40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum
        )
        self.startExitLayout.addItem(spacerItem4)
        self.mainWindowLayout.addLayout(self.startExitLayout)

        self.retranslateUi(mainWindow)
        QtCore.QMetaObject.connectSlotsByName(mainWindow)

    def retranslateUi(self, mainWindow: QtWidgets.QWidget) -> None:
        _translate = QtCore.QCoreApplication.translate
        mainWindow.setWindowTitle(_translate("mainWindow", "Auto Schedule"))
        self.genReportBox.setTitle(_translate("mainWindow", "Generate Destiny Report"))
        self.startDateLabel.setText(_translate("mainWindow", "Start Date:"))
        self.endDateLabel.setText(_translate("mainWindow", "End Date:"))
        self.saveReportPathLabel.setText(_translate("mainWindow", "Save Path:"))
        self.selectSaveReportPath.setText(
            _translate("mainWindow", self.saveReportToPath)
        )
        self.browseSaveReportButton.setText(_translate("mainWindow", "Browse"))
        self.label.setText(_translate("mainWindow", "Location: "))
        self.GBCcheckBox.setText(_translate("mainWindow", "GBC"))
        self.SFCcheckBox.setText(_translate("mainWindow", "SFC"))
        self.StartButton.setText(_translate("mainWindow", "Start"))
        self.exitButton.setText(_translate("mainWindow", "Exit"))

        self.selectStartDate.dateChanged.connect(self.startDateChanged)
        self.selectEndDate.dateChanged.connect(self.endDateChanged)
        self.browseSaveReportButton.clicked.connect(self.saveReportDirectory)
        self.GBCcheckBox.toggled.connect(self.GBCcheckBoxState)
        self.SFCcheckBox.toggled.connect(self.SFCcheckBoxState)
        self.exitButton.clicked.connect(self.exitApp)
        self.StartButton.clicked.connect(self.startApp)

    def startDateChanged(self) -> None:
        self.startDate = str(self.selectStartDate.date().toPyDate())
        self.endDate = str(self.selectStartDate.date().addMonths(12).toPyDate())
        self.selectEndDate.setDate(self.selectStartDate.date().addMonths(12))

    def endDateChanged(self) -> None:
        self.endDate = str(self.selectEndDate.date().toPyDate())

    def saveReportDirectory(self) -> None:
        path = os.path.normpath(
            QFileDialog.getExistingDirectory(
                None, "Save Destiny Report to", self.saveReportToPath
            )
        )
        if path and path != ".":
            self.saveReportToPath = path
            self.selectSaveReportPath.setText(self.saveReportToPath)

    def GBCcheckBoxState(self) -> None:
        if self.GBCcheckBox.isChecked():
            self.GBCScheduleOutput = True
        else:
            self.GBCScheduleOutput = False

    def SFCcheckBoxState(self) -> None:
        if self.SFCcheckBox.isChecked():
            self.SFCScheduleOutput = True
        else:
            self.SFCScheduleOutput = False

    def exitApp(self) -> None:
        reply = QMessageBox.question(
            None,
            "Exit",
            "Are you sure you want to exit?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            # Save settings state to config.ini file
            self.settings.setValue("saveReportToPath", self.saveReportToPath)
            self.settings.setValue("GBCScheduleOutput", self.GBCScheduleOutput)
            self.settings.setValue("SFCScheduleOutput", self.SFCScheduleOutput)
            self.settings.setValue("uploadGBCSchedule", self.uploadGBCSchedule)
            self.settings.setValue("GBCGDriveFolderId", self.GBCGDriveFolderId)
            self.settings.setValue("uploadSFCSchedule", self.uploadSFCSchedule)
            self.settings.setValue("SFCGDriveFolderId", self.SFCGDriveFolderId)
            self.settings.setValue("attachGBCSchedule", self.attachGBCSchedule)
            self.settings.setValue("GBCCalendarId", self.GBCCalendarId)
            self.settings.setValue("attachSFCSchedule", self.attachSFCSchedule)
            self.settings.setValue("SFCCalendarId", self.SFCCalendarId)
            sys.exit()
        else:
            pass

    def startApp(self) -> None:
        # Invalid input checks
        result = 1
        if self.endDate < self.startDate:
            QMessageBox.warning(
                None, "Invalid date range", "Please select a valid date range."
            )
            return
        elif self.saveReportToPath == "":
            QMessageBox.warning(
                None,
                "Save location error",
                "Please select where you want to save the report to.",
            )
            return
        else:
            if os.path.isdir(self.saveReportToPath):
                result = self.genReportFunction(self.startDate, self.endDate)
            else:
                QMessageBox.warning(
                    None,
                    "Save location error",
                    (
                        "The directory you've selected does not exist. "
                        "Please select where you want to save the report to."
                    ),
                )
                return
        if result == 0:
            QMessageBox.warning(None, "Error!!!", "An unexpected error occured.")
        else:
            QMessageBox.warning(None, "Done", "Done creating signs.")

    def genReportFunction(self, startDate: str, endDate: str) -> bool:
        # Set Chrome defaults to automate download
        chrome_options = Options()
        chrome_options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": self.saveReportToPath,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.endabled": True,
            },
        )

        locationList = []
        if self.GBCScheduleOutput:
            locationList.append("Golden Bear Center")
        if self.SFCScheduleOutput:
            locationList.append("San Francisco Center")

        # Delete old report if it exists
        if os.path.exists(f"{self.saveReportToPath}\\SectionScheduleDailySummary.xls"):
            os.remove(f"{self.saveReportToPath}\\SectionScheduleDailySummary.xls")
        if os.path.exists(
            f"{self.saveReportToPath}\\SectionScheduleDailySummary (1).xls"
        ):
            os.remove(f"{self.saveReportToPath}\\SectionScheduleDailySummary (1).xls")
        if os.path.exists(
            f"{self.saveReportToPath}\\SectionScheduleDailySummary (2).xls"
        ):
            os.remove(f"{self.saveReportToPath}\\SectionScheduleDailySummary (2).xls")
        if os.path.exists(
            f"{self.saveReportToPath}\\SectionScheduleDailySummary (3).xls"
        ):
            os.remove(f"{self.saveReportToPath}\\SectionScheduleDailySummary (3).xls")

        reportPath = []
        browser = webdriver.Chrome(
            executable_path=ChromeDriverManager().install(), options=chrome_options
        )
        browser.get("https://berkeleysv.destinysolutions.com")
        WebDriverWait(browser, 3600).until(
            EC.presence_of_element_located((By.ID, "main-area-body"))
        )
        for i in range(len(locationList)):
            # Download Destiny Report
            browser.get(
                "https://berkeleysv.destinysolutions.com/srs/reporting/sectionScheduleDailySummary.do?method=load"  # noqa: E501
            )
            startDateElm = browser.find_element_by_id("startDateRecordString")
            startDateElm.send_keys(startDate)
            endDateElm = browser.find_element_by_id("endDateRecordString")
            endDateElm.send_keys(endDate)
            campusElm = browser.find_element_by_name("scheduleBlock.campusId")
            campusElm.send_keys(self.center[locationList[i]]["campus"])
            buildingElm = browser.find_element_by_name("scheduleBlock.buildingId")
            buildingElm.send_keys(self.center[locationList[i]]["building"])
            outputTypeElm = browser.find_element_by_name("outputType")
            outputTypeElm.send_keys("Output to XLS (Export)")
            generateReportElm = browser.find_element_by_id("processReport")
            generateReportElm.click()
            if i == 0:
                while not os.path.exists(
                    f"{self.saveReportToPath}\\SectionScheduleDailySummary.xls"
                ):
                    time.sleep(1)
                reportPath.append(
                    f"{self.saveReportToPath}\\SectionScheduleDailySummary.xls"
                )
            elif i == 1:
                while not os.path.exists(
                    f"{self.saveReportToPath}\\SectionScheduleDailySummary (1).xls"
                ):
                    time.sleep(1)
                reportPath.append(
                    f"{self.saveReportToPath}\\SectionScheduleDailySummary (1).xls"
                )
            elif i == 2:
                while not os.path.exists(
                    f"{self.saveReportToPath}\\SectionScheduleDailySummary (2).xls"
                ):
                    time.sleep(1)
                reportPath.append(
                    f"{self.saveReportToPath}\\SectionScheduleDailySummary (2).xls"
                )
            else:
                while not os.path.exists(
                    f"{self.saveReportToPath}\\SectionScheduleDailySummary (3).xls"
                ):
                    time.sleep(1)
                reportPath.append(
                    f"{self.saveReportToPath}\\SectionScheduleDailySummary (3).xls"
                )
        browser.quit()

        for rp in reportPath:
            self.createSchedule(rp)
        return True

    def createSchedule(self, reportPath: str) -> None:
        # Read in courses from Excel
        # 1     B   Date
        # 3     D   Type
        # 4     E   Start Time
        # 6     G   End Time
        # 9     J   Section Number
        # 11    L   Section Title
        # 12    M   Instructor
        # 13    N   Building
        # 15    P   Room
        # 16    Q   Configuration
        # 17    R   Technology
        # 18    S   Section Size
        # 20    U   Notes
        # 22    W   Approval Status

        # Read into Pandas dataframe for relevant columns
        schedule = pd.read_excel(
            reportPath,
            header=6,
            skipfooter=1,
            usecols=[1, 15, 18, 4, 6, 11, 12, 9, 17, 20, 13, 22],
            parse_dates=["Start Time", "End Time"],
        )
        schedule = schedule[schedule["Approval Status"] == "Final Approval"].copy()

        # Determine if the Destiny report does not have any classes
        if schedule.empty:
            print(f"No classes found in {reportPath}")
        else:  # Not empty, determine location and template to use
            location = self.centerReverse[schedule.iloc[0][6]]["name"]
            try:
                gauth = GoogleAuth()
                scope = [
                    "https://www.googleapis.com/auth/drive",
                    "https://www.googleapis.com/auth/calendar",
                ]
                gauth.credentials = ServiceAccountCredentials.from_json_keyfile_name(
                    "service_file.json", scope
                )
                drive = GoogleDrive(gauth)
                service = build("calendar", "v3", credentials=gauth.credentials)
            except Exception as error:
                print(f"[Error] {error}")

            if location == "SFC":
                fileName, date = self.SFCSchedule(schedule, location)
                print(f"uploadSFCSchedule: {self.uploadSFCSchedule}")
                print(f"attachSFCSchedule: {self.attachSFCSchedule}")
                if self.uploadSFCSchedule:
                    self.uploadToGoogleDrive(drive, fileName, self.SFCGDriveFolderId)
                if self.attachSFCSchedule:
                    self.createGoogleCalendarEvent(
                        drive,
                        fileName,
                        self.SFCGDriveFolderId,
                        service,
                        self.SFCCalendarId,
                        date,
                    )
            else:
                fileName, date = self.GBCSchedule(schedule, location)
                if self.uploadGBCSchedule:
                    self.uploadToGoogleDrive(drive, fileName, self.GBCGDriveFolderId)
                if self.attachGBCSchedule:
                    self.createGoogleCalendarEvent(
                        drive,
                        fileName,
                        self.GBCGDriveFolderId,
                        service,
                        self.GBCCalendarId,
                        date,
                    )

    def createGoogleCalendarEvent(
        self,
        drive: GoogleDrive,
        fileName: str,
        parentFolderId: str,
        service: Resource,
        calendarId: str,
        date: str,
    ) -> None:
        """Function to create a Google Calendar Event and attach Google Sheets to:
        Args:
            drive (obj): Google Drive authenticated class object.
            filenName (str): File name of the attachment in Google Drive.
            parentFolderId (str): Google Drive parent folder ID
            service (obj): Google Calendar service to interact with the API.
            calenderId (str): Google Calendar ID to create the event in.
            date (str): Date of the event.
        """
        # Find the file's metadata in Google Drive under the parent folder.
        timeout = 0
        while timeout < 30:
            file = self.getFile(drive, fileName[:-5], parentFolderId)
            if file:
                break
            timeout += timeout if timeout else 2
            # print(f"[Debug] Checking Google Drive folder again in {timeout} seconds.")
            time.sleep(timeout)

        if not file:
            print(f"[Warning] Could not find {fileName} in Google Drive.")
            return None

        # Find whether an event already exists.
        event_start = (
            datetime.datetime.strptime(f"{date}", "%Y-%m-%d").isoformat() + "Z"
        )
        events_result = (
            service.events()
            .list(
                calendarId=calendarId,
                timeMin=event_start,
                maxResults=100,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )
        events = events_result.get("items", [])

        existingEvent = {}
        for event in events:
            if event["summary"] == fileName[:-5]:
                existingEvent = event
                break

        # Create the event body payload, including attachment information.
        body = {
            "summary": fileName[:-5],
            "start": {
                "date": date,
                "timeZone": "America/Los_Angeles",
            },
            "end": {
                "date": (
                    datetime.datetime.strptime(date, "%Y-%m-%d")
                    + datetime.timedelta(days=1)
                ).strftime("%Y-%m-%d"),
                "timeZone": "America/Los_Angeles",
            },
            "reminders": {
                "useDefault": True,
            },
            "attachments": [
                {
                    "fileUrl": f"https://drive.google.com/open?id={file['id']}",
                    # "title": file['title'],
                    "title": fileName[:-5],
                    "mimeType": file["mimeType"],
                    "iconLink": file["iconLink"],
                    "fileId": file["id"],
                }
            ],
        }

        # If the event already exists, update the existing event.
        if existingEvent:
            event = (
                service.events()
                .update(
                    calendarId=calendarId,
                    eventId=existingEvent["id"],
                    body=body,
                    supportsAttachments=True,
                )
                .execute()
            )
        else:  # Event does not exist. Create a new event.
            event = (
                service.events()
                .insert(calendarId=calendarId, body=body, supportsAttachments=True)
                .execute()
            )

    def getFile(self, drive: GoogleDrive, fileName: str, parentFolderId: str) -> Any:
        # Retrieve the file's metadata under the parent folder. Return file.
        file_list = drive.ListFile(
            {"q": f"'{parentFolderId}' in parents and trashed=false"}
        ).GetList()
        for file in file_list:
            if file["title"] == fileName:
                return file
        return {}

    def uploadToGoogleDrive(
        self,
        drive: GoogleDrive,
        fileName: str,
        folderId: str = "",
        folderName: str = "",
    ) -> int:
        if not fileName:
            print("[Error] No file name given. Nothing to upload.")
            return 0

        # Find the folder ID in Google Drive if not given.
        if folderName and not folderId:
            folders = drive.ListFile(
                {
                    "q": (
                        f"title='{folderName}' "
                        f"and mimeType='application/vnd.google-apps.folder' "
                        f"and trashed=false"
                    )
                }
            ).GetList()
            for folder in folders:
                if folder["title"] == folderName:
                    folderId = folder["id"]

        if folderId:
            # Find the file's metadata in Google Drive under the parent folder.
            fileId = self.getFile(drive, fileName[:-5], folderId).get("id", "")
            if fileId:
                # A file with the same name already exists. Replace the existing file.
                file = drive.CreateFile(
                    metadata={
                        "title": fileName,
                        "id": fileId,
                        "parents": [{"id": folderId}],
                    }
                )
            else:
                # A file with the same name does not already exists. Create a new file.
                file = drive.CreateFile(
                    metadata={"title": fileName, "parents": [{"id": folderId}]}
                )

            # Set file to upload and convert to Google Doc type.
            file.SetContentFile(f"{self.saveReportToPath}\\{fileName}")
            file.Upload({"convert": True})
        else:
            return 0

        return 1

    def GBCSchedule(self, schedule: pd.DataFrame, location: str) -> Tuple[str, str]:
        # Sort the schedule.
        sortedSchedule = schedule.sort_values(by=["Date", "Start Time", "Room"])
        sortedSchedule["Start Time"] = sortedSchedule["Start Time"].dt.strftime(
            "%I:%M %p"
        )
        sortedSchedule["End Time"] = sortedSchedule["End Time"].dt.strftime("%I:%M %p")
        sortedSchedule = sortedSchedule.fillna("")
        dateList = pd.to_datetime(sortedSchedule["Date"].unique())

        # Create the Excel file to write to.
        fileName = (
            f"{location} Schedule {dateList[0].strftime('%Y-%m-%d')} "
            f"{dateList[0].strftime('%A')}.xlsx"
        )
        writer = pd.ExcelWriter(
            f"{self.saveReportToPath}\\{fileName}", engine="xlsxwriter"
        )
        workbook = writer.book

        # Set the Excel workbook formatting
        worksheet = workbook.add_worksheet(location)
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column("K:XFD", None, None, {"hidden": True})
        worksheet.freeze_panes(2, 0)
        worksheet.set_landscape()  # Page orientation as landscape.
        worksheet.hide_gridlines(0)  # Don’t hide gridlines.
        worksheet.fit_to_pages(1, 0)  # Fit to 1x1 pages.
        worksheet.center_horizontally()
        worksheet.center_vertically()
        worksheet.set_paper(1)  # Set paper size to 8.5" x 11"
        worksheet.set_margins(left=0.25, right=0.25, top=0.25, bottom=0.25)
        worksheet.set_header("", {"margin": 0})
        worksheet.set_footer("", {"margin": 0})

        worksheet.set_column("A:A", 29)  # Column A (Date) width
        worksheet.set_column("B:B", 7.29)  # Column B (Room) width
        worksheet.set_column("C:C", 3.86)  # Column C (Size) width
        worksheet.set_column("D:D", 9.29)  # Column D (Start Date) width
        worksheet.set_column("E:E", 8.43)  # Column E (End Date) width
        worksheet.set_column("F:F", 46.29)  # Column F (Section Title) width
        worksheet.set_column("G:G", 12.57)  # Column G (Instructor) width
        worksheet.set_column("H:H", 16)  # Column H (Section Number) width
        worksheet.set_column("I:I", 18.57)  # Column I (Technology) width
        worksheet.set_column("J:J", 64)  # Column J (Notes) width

        locationFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#000000",
                "bg_color": "#FFC000",
            }
        )

        genFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#C00000",
                "bg_color": "#FDE9D9",
            }
        )

        headerFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "bottom": 2,
                "bottom_color": "#000000",
            }
        )

        bodyFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": False,
                "valign": "top",
                "text_wrap": False,
                "font_color": "#000000",
            }
        )

        daySeparator = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#000000",
                "bg_color": "#00B050",
            }
        )

        roomFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "align": "right",
                "font_color": "#C00000",
            }
        )

        roomFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#C00000",
            }
        )

        laptopReadyFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "italic": True,
                "text_wrap": False,
                "font_color": "#0070C0",
            }
        )

        instructorFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": False,
                "text_wrap": False,
                "font_color": "#C00000",
            }
        )

        # AM
        amFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": False,
                "text_wrap": False,
                "font_color": "#FF0000",
                "bg_color": "#DAEEF3",
            }
        )

        # Computer Lab
        labFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#FF0000",
                "bg_color": "#FFFF00",
            }
        )

        worksheet.write(0, 0, f"{sortedSchedule.iloc[0][6]}", locationFormat)
        worksheet.write(0, 1, "", locationFormat)
        worksheet.write(0, 2, "", locationFormat)
        worksheet.write(0, 3, "", locationFormat)
        worksheet.write(0, 4, "", locationFormat)

        worksheet.write(
            0,
            5,
            (
                f"Report generated as of {dateList[0].strftime('%A')}, "
                f"{dateList[0].strftime('%B %d, %Y').replace(' 0', ' ')}"
            ),
            genFormat,
        )
        for col_num, value in enumerate(
            [
                "Date",
                "Room",
                "Size",
                "Start Time",
                "End Time",
                "Section Title",
                "Instructor",
                "Section Number",
                "Technology",
                "Notes",
            ]
        ):
            worksheet.write(1, col_num, value, headerFormat)
        excelRow = 2
        # Loop through each day
        for i in range(0, len(dateList)):
            singleDaySched = sortedSchedule.loc[
                sortedSchedule["Date"] == dateList[i], :
            ]
            morningBlock = singleDaySched.loc[
                singleDaySched["Start Time"].astype("datetime64") < "12:00:00", :
            ]
            afternoonBlock = singleDaySched.loc[
                (singleDaySched["Start Time"].astype("datetime64") >= "12:00:00")
                & (singleDaySched["Start Time"].astype("datetime64") < "17:00:00"),
                :,
            ]
            eveningBlock = singleDaySched.loc[
                singleDaySched["Start Time"].astype("datetime64") >= "17:00:00", :
            ]

            worksheet.write(
                excelRow,
                0,
                (
                    f"{dateList[i].strftime('%A')}, "
                    f"{dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}"
                ),
                daySeparator,
            )
            worksheet.write_number(excelRow, 3, len(singleDaySched.index), amFormat)
            excelRow += 1
            if not morningBlock.empty:
                for idx, row in morningBlock.iterrows():
                    worksheet.write(
                        excelRow,
                        0,
                        (
                            f"{dateList[i].strftime('%A')}, "
                            f"{dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}"
                        ),
                        bodyFormat,
                    )
                    if location == "GBC" and row["Room"] == "Classroom 201":
                        worksheet.write_number(
                            excelRow,
                            1,
                            int(row["Room"].replace("Classroom ", "")),
                            labFormat,
                        )
                    else:
                        if (
                            "Conference Room" in row["Room"]
                            or "Conference Room" in row["Room"]
                        ):
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Conference Room ", "CR"),
                                roomFormat,
                            )
                        elif row["Room"].replace("Classroom ", "").isdigit():
                            worksheet.write_number(
                                excelRow,
                                1,
                                int(row["Room"].replace("Classroom ", "")),
                                roomFormat,
                            )
                        else:
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Classroom ", "").lstrip("0"),
                                roomFormat,
                            )
                    worksheet.write(excelRow, 2, row["Section Size"], bodyFormat)
                    if "AM" in row["Start Time"]:
                        worksheet.write(
                            excelRow, 3, row["Start Time"].lstrip("0"), amFormat
                        )
                    else:
                        worksheet.write(
                            excelRow, 3, row["Start Time"].lstrip("0"), bodyFormat
                        )
                    if "AM" in row["End Time"]:
                        worksheet.write(
                            excelRow, 4, row["End Time"].lstrip("0"), amFormat
                        )
                    else:
                        worksheet.write(
                            excelRow, 4, row["End Time"].lstrip("0"), bodyFormat
                        )
                    if "Boot Camp" in row["Section Title"]:
                        worksheet.write(
                            excelRow, 5, row["Section Title"], laptopReadyFormat
                        )
                    else:
                        worksheet.write(excelRow, 5, row["Section Title"], bodyFormat)
                    if row["Instructor"] == "Instructor To Be Announced":
                        worksheet.write(excelRow, 6, "TBA", instructorFormat)
                    elif not pd.isnull(row["Instructor"]):
                        worksheet.write(
                            excelRow, 6, row["Instructor"], instructorFormat
                        )
                    worksheet.write(excelRow, 7, row["Section Number"], bodyFormat)
                    worksheet.write(excelRow, 8, row["Technology"], bodyFormat)
                    worksheet.write(excelRow, 9, row["Notes"], bodyFormat)
                    excelRow += 1

            if not afternoonBlock.empty:
                for idx, row in afternoonBlock.iterrows():
                    worksheet.write(
                        excelRow,
                        0,
                        (
                            f"{dateList[i].strftime('%A')}, "
                            f"{dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}"
                        ),
                        bodyFormat,
                    )
                    if location == "GBC" and row["Room"] == "Classroom 201":
                        worksheet.write_number(
                            excelRow,
                            1,
                            int(row["Room"].replace("Classroom ", "")),
                            labFormat,
                        )
                    else:
                        if (
                            "Conference Room" in row["Room"]
                            or "Conference Room" in row["Room"]
                        ):
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Conference Room ", "CR"),
                                roomFormat,
                            )
                        elif row["Room"].replace("Classroom ", "").isdigit():
                            worksheet.write_number(
                                excelRow,
                                1,
                                int(row["Room"].replace("Classroom ", "")),
                                roomFormat,
                            )
                        else:
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Classroom ", "").lstrip("0"),
                                roomFormat,
                            )
                    worksheet.write(excelRow, 2, row["Section Size"], bodyFormat)
                    worksheet.write(
                        excelRow, 3, row["Start Time"].lstrip("0"), bodyFormat
                    )
                    worksheet.write(
                        excelRow, 4, row["End Time"].lstrip("0"), bodyFormat
                    )
                    if "Boot Camp" in row["Section Title"]:
                        worksheet.write(
                            excelRow, 5, row["Section Title"], laptopReadyFormat
                        )
                    else:
                        worksheet.write(excelRow, 5, row["Section Title"], bodyFormat)
                    if row["Instructor"] == "Instructor To Be Announced":
                        worksheet.write(excelRow, 6, "TBA", instructorFormat)
                    elif not pd.isnull(row["Instructor"]):
                        worksheet.write(
                            excelRow, 6, row["Instructor"], instructorFormat
                        )
                    worksheet.write(excelRow, 7, row["Section Number"], bodyFormat)
                    worksheet.write(excelRow, 8, row["Technology"], bodyFormat)
                    worksheet.write(excelRow, 9, row["Notes"], bodyFormat)
                    excelRow += 1

            if not eveningBlock.empty:
                for idx, row in eveningBlock.iterrows():
                    worksheet.write(
                        excelRow,
                        0,
                        (
                            f"{dateList[i].strftime('%A')}, "
                            f"{dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}"
                        ),
                        bodyFormat,
                    )
                    if location == "GBC" and row["Room"] == "Classroom 201":
                        worksheet.write_number(
                            excelRow,
                            1,
                            int(row["Room"].replace("Classroom ", "")),
                            labFormat,
                        )
                    else:
                        if (
                            "Conference Room" in row["Room"]
                            or "Conference Room" in row["Room"]
                        ):
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Conference Room ", "CR"),
                                roomFormat,
                            )
                        elif row["Room"].replace("Classroom ", "").isdigit():
                            worksheet.write_number(
                                excelRow,
                                1,
                                int(row["Room"].replace("Classroom ", "")),
                                roomFormat,
                            )
                        else:
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Classroom ", "").lstrip("0"),
                                roomFormat,
                            )
                    worksheet.write(excelRow, 2, row["Section Size"], bodyFormat)
                    worksheet.write(
                        excelRow, 3, row["Start Time"].lstrip("0"), bodyFormat
                    )
                    worksheet.write(
                        excelRow, 4, row["End Time"].lstrip("0"), bodyFormat
                    )
                    if "Boot Camp" in row["Section Title"]:
                        worksheet.write(
                            excelRow, 5, row["Section Title"], laptopReadyFormat
                        )
                    else:
                        worksheet.write(excelRow, 5, row["Section Title"], bodyFormat)
                    if row["Instructor"] == "Instructor To Be Announced":
                        worksheet.write(excelRow, 6, "TBA", instructorFormat)
                    elif not pd.isnull(row["Instructor"]):
                        worksheet.write(
                            excelRow, 6, row["Instructor"], instructorFormat
                        )
                    worksheet.write(excelRow, 7, row["Section Number"], bodyFormat)
                    worksheet.write(excelRow, 8, row["Technology"], bodyFormat)
                    worksheet.write(excelRow, 9, row["Notes"], bodyFormat)
                    excelRow += 1

        workbook.close()
        return fileName, dateList[0].strftime("%Y-%m-%d")

    def SFCSchedule(self, schedule: pd.DataFrame, location: str) -> Tuple[str, str]:
        # Sort the schedule.
        sortedSchedule = schedule.sort_values(by=["Date", "Room", "Start Time"])
        sortedSchedule["Start Time"] = sortedSchedule["Start Time"].dt.strftime(
            "%I:%M %p"
        )
        sortedSchedule["End Time"] = sortedSchedule["End Time"].dt.strftime("%I:%M %p")
        sortedSchedule = sortedSchedule.fillna("")
        dateList = pd.to_datetime(sortedSchedule["Date"].unique())

        # Create the Excel file to write to.
        fileName = (
            f"{location} Schedule {dateList[0].strftime('%Y-%m-%d')} "
            f"{dateList[0].strftime('%A')}.xlsx"
        )
        writer = pd.ExcelWriter(
            f"{self.saveReportToPath}\\{fileName}", engine="xlsxwriter"
        )
        workbook = writer.book

        # Set the Excel workbook formatting.
        worksheet = workbook.add_worksheet(location)
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column("K:XFD", None, None, {"hidden": True})
        worksheet.freeze_panes(2, 0)
        worksheet.set_landscape()  # Page orientation as landscape.
        worksheet.hide_gridlines(0)  # Don’t hide gridlines.
        worksheet.fit_to_pages(1, 0)  # Fit to 1x1 pages.
        worksheet.center_horizontally()
        worksheet.center_vertically()
        worksheet.set_paper(1)  # Set paper size to 8.5" x 11"
        worksheet.set_margins(left=0.25, right=0.25, top=0.25, bottom=0.25)
        worksheet.set_header("", {"margin": 0})
        worksheet.set_footer("", {"margin": 0})

        worksheet.set_column("A:A", 29)  # Column A (Date) width
        worksheet.set_column("B:B", 7.29)  # Column B (Room) width
        worksheet.set_column("C:C", 3.86)  # Column C (Size) width
        worksheet.set_column("D:D", 9.29)  # Column D (Start Date) width
        worksheet.set_column("E:E", 8.43)  # Column E (End Date) width
        worksheet.set_column("F:F", 46.29)  # Column F (Section Title) width
        worksheet.set_column("G:G", 12.57)  # Column G (Instructor) width
        worksheet.set_column("H:H", 16)  # Column H (Section Number) width
        worksheet.set_column("I:I", 18.57)  # Column H Technology) width
        worksheet.set_column("J:J", 64)  # Column H (Notes) width

        locationFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#000000",
                "bg_color": "#FFC000",
            }
        )

        genFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#C00000",
                "bg_color": "#FDE9D9",
            }
        )

        headerFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                # 'valign': 'vcenter',
                "bottom": 2,
                "bottom_color": "#000000",
            }
        )

        bodyFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": False,
                "valign": "top",
                "text_wrap": False,
                "font_color": "#000000",
            }
        )

        daySeparator = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#000000",
                "bg_color": "#00B050",
            }
        )

        roomFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "align": "right",
                "font_color": "#C00000",
            }
        )

        roomFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#C00000",
            }
        )

        laptopReadyFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "italic": True,
                "text_wrap": False,
                "font_color": "#0070C0",
            }
        )

        instructorFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": False,
                "text_wrap": False,
                "font_color": "#C00000",
            }
        )

        # AM
        amFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": False,
                "text_wrap": False,
                "font_color": "#FF0000",
                "bg_color": "#DAEEF3",
            }
        )

        # Computer Lab
        labFormat = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "text_wrap": False,
                "font_color": "#FF0000",
                "bg_color": "#FFFF00",
            }
        )

        worksheet.write(0, 0, f"{sortedSchedule.iloc[0][6]}", locationFormat)
        worksheet.write(0, 1, "", locationFormat)
        worksheet.write(0, 2, "", locationFormat)
        worksheet.write(0, 3, "", locationFormat)
        worksheet.write(0, 4, "", locationFormat)

        worksheet.write(
            0,
            5,
            (
                f"Report generated as of {dateList[0].strftime('%A')}, "
                f"{dateList[0].strftime('%B %d, %Y').replace(' 0', ' ')}"
            ),
            genFormat,
        )
        for col_num, value in enumerate(
            [
                "Date",
                "Room",
                "Size",
                "Start Time",
                "End Time",
                "Section Title",
                "Instructor",
                "Section Number",
                "Technology",
                "Notes",
            ]
        ):
            worksheet.write(1, col_num, value, headerFormat)
        excelRow = 2
        # Loop through each day
        for i in range(0, len(dateList)):
            singleDaySched = sortedSchedule.loc[
                sortedSchedule["Date"] == dateList[i], :
            ]
            daytimeBlock = singleDaySched.loc[
                singleDaySched["Start Time"].astype("datetime64") < "17:00:00", :
            ]
            eveningBlock = singleDaySched.loc[
                singleDaySched["Start Time"].astype("datetime64") >= "17:00:00", :
            ]

            worksheet.write(
                excelRow,
                0,
                (
                    f"{dateList[i].strftime('%A')}, "
                    f"{dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}"
                ),
                daySeparator,
            )
            worksheet.write_number(excelRow, 3, len(singleDaySched.index), amFormat)
            excelRow += 1
            if not daytimeBlock.empty:
                for idx, row in daytimeBlock.iterrows():
                    worksheet.write(
                        excelRow,
                        0,
                        (
                            f"{dateList[i].strftime('%A')}, "
                            f"{dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}"
                        ),
                        bodyFormat,
                    )
                    if row["Room"].replace("Classroom ", "") in [
                        "502",
                        "510",
                        "514",
                        "515",
                    ]:
                        worksheet.write_number(
                            excelRow,
                            1,
                            int(row["Room"].replace("Classroom ", "")),
                            labFormat,
                        )
                    else:
                        if (
                            "Conference Room" in row["Room"]
                            or "Conference Room" in row["Room"]
                        ):
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Conference Room ", "CR"),
                                roomFormat,
                            )
                        elif row["Room"].replace("Classroom ", "").isdigit():
                            worksheet.write_number(
                                excelRow,
                                1,
                                int(row["Room"].replace("Classroom ", "")),
                                roomFormat,
                            )
                        else:
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Classroom ", "").lstrip("0"),
                                roomFormat,
                            )
                    worksheet.write(excelRow, 2, row["Section Size"], bodyFormat)
                    if "AM" in row["Start Time"]:
                        worksheet.write(
                            excelRow, 3, row["Start Time"].lstrip("0"), amFormat
                        )
                    else:
                        worksheet.write(
                            excelRow, 3, row["Start Time"].lstrip("0"), bodyFormat
                        )
                    if "AM" in row["End Time"]:
                        worksheet.write(
                            excelRow, 4, row["End Time"].lstrip("0"), amFormat
                        )
                    else:
                        worksheet.write(
                            excelRow, 4, row["End Time"].lstrip("0"), bodyFormat
                        )
                    if "Boot Camp" in row["Section Title"]:
                        worksheet.write(
                            excelRow, 5, row["Section Title"], laptopReadyFormat
                        )
                    else:
                        worksheet.write(excelRow, 5, row["Section Title"], bodyFormat)
                    if row["Instructor"] == "Instructor To Be Announced":
                        worksheet.write(excelRow, 6, "TBA", instructorFormat)
                    elif not pd.isnull(row["Instructor"]):
                        worksheet.write(
                            excelRow, 6, row["Instructor"], instructorFormat
                        )
                    worksheet.write(excelRow, 7, row["Section Number"], bodyFormat)
                    worksheet.write(excelRow, 8, row["Technology"], bodyFormat)
                    worksheet.write(excelRow, 9, row["Notes"], bodyFormat)
                    excelRow += 1

            if not eveningBlock.empty:
                for idx, row in eveningBlock.iterrows():
                    worksheet.write(
                        excelRow,
                        0,
                        (
                            f"{dateList[i].strftime('%A')}, "
                            f"{dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}"
                        ),
                        bodyFormat,
                    )
                    if row["Room"].replace("Classroom ", "") in [
                        "502",
                        "510",
                        "514",
                        "515",
                    ]:
                        worksheet.write_number(
                            excelRow,
                            1,
                            int(row["Room"].replace("Classroom ", "")),
                            labFormat,
                        )
                    else:
                        if (
                            "Conference Room" in row["Room"]
                            or "Conference Room" in row["Room"]
                        ):
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Conference Room ", "CR"),
                                roomFormat,
                            )
                        elif row["Room"].replace("Classroom ", "").isdigit():
                            worksheet.write_number(
                                excelRow,
                                1,
                                int(row["Room"].replace("Classroom ", "")),
                                roomFormat,
                            )
                        else:
                            worksheet.write(
                                excelRow,
                                1,
                                row["Room"].replace("Classroom ", "").lstrip("0"),
                                roomFormat,
                            )
                    worksheet.write(excelRow, 2, row["Section Size"], bodyFormat)
                    worksheet.write(
                        excelRow, 3, row["Start Time"].lstrip("0"), bodyFormat
                    )
                    worksheet.write(
                        excelRow, 4, row["End Time"].lstrip("0"), bodyFormat
                    )
                    if "Boot Camp" in row["Section Title"]:
                        worksheet.write(
                            excelRow, 5, row["Section Title"], laptopReadyFormat
                        )
                    else:
                        worksheet.write(excelRow, 5, row["Section Title"], bodyFormat)
                    if row["Instructor"] == "Instructor To Be Announced":
                        worksheet.write(excelRow, 6, "TBA", instructorFormat)
                    elif not pd.isnull(row["Instructor"]):
                        worksheet.write(
                            excelRow, 6, row["Instructor"], instructorFormat
                        )
                    worksheet.write(excelRow, 7, row["Section Number"], bodyFormat)
                    worksheet.write(excelRow, 8, row["Technology"], bodyFormat)
                    worksheet.write(excelRow, 9, row["Notes"], bodyFormat)
                    excelRow += 1

        workbook.close()
        return fileName, dateList[0].strftime("%Y-%m-%d")


if __name__ == "__main__":
    # os.environ["QT_AUTO_SCREEN_FACTOR"] = "1"
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = QtWidgets.QWidget()
    ui = Ui_mainWindow()
    ui.setupUi(mainWindow)
    mainWindow.show()
    sys.exit(app.exec_())
