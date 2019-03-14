import sys
import os
import threading
import time
import queue
import ipaddress
#from winreg import *
import re
import traceback
import ctypes
from shutil import copyfile
from collections import OrderedDict
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QAction, QLabel, QFormLayout,QBoxLayout, QVBoxLayout, QHBoxLayout, QLineEdit, QPlainTextEdit, QPushButton, QProgressBar, QTabWidget, QFileDialog, QMessageBox, QScrollArea, QStatusBar, QDialog, QTableWidget, QTableWidgetItem, QSplitter, QSizePolicy, QMenu, QCheckBox, QDateTimeEdit
from PyQt5.QtGui import QFont, QBrush, QColor,QCursor
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QObject, pyqtSlot, QDateTime
from ComputerInfoSharedResources.CIForms import ShortcutCheckboxForm, AuthenticationForm, FileForm, AppsForm
from ComputerInfoSharedResources.dynamic_forms.forms import DynamicForm
from ComputerInfoSharedResources.dynamic_forms.models import DynamicModel
from ComputerInfoSharedResources.CIExcel import CIWorkbook
from ComputerInfoSharedResources.CIProgram import ProgramChoices
from ComputerInfoSharedResources.CITime import format_time
from ComputerInfoSharedResources.CIStorage import ThreadSafeCounter, ThreadSafeBool
from ComputerInfoSharedResources.CIWMI import ComputerInfo, WMIThread
from ComputerInfoSharedResources.CICustomWidgets import CustomScrollBox
from smtplib import SMTP_SSL, SMTP
import getpass
from email.mime.text import MIMEText
from email.utils import formataddr
import argparse
from urllib import request
import logging
from functools import partial
import pythoncom
from pathlib import Path
from datetime import datetime

try:
    from win10toast import ToastNotifier
except: pass

def safe_divide(x,y):
    if not x is None and y:
        return x/y
    else:
        return 0

class TestThreadClass(QObject):
    started = pyqtSignal()
    send_update = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, callback, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs
        self.callback = callback
        super().__init__()
    
    @pyqtSlot()
    def run(self):
        self.callback(*self.args, **self.kwargs)

class GuiThreadClass(QObject):

    started = pyqtSignal()
    progress_update = pyqtSignal(int,int)
    configure_prog = pyqtSignal(int)
    complete_run = pyqtSignal()
    summary_dict = pyqtSignal(int,dict,dict,dict,dict,list,dict,dict,dict)

    def __init__(self,callback, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs
        self.callback = callback
        super().__init__()

    @pyqtSlot()
    def run(self):
        self.callback(*self.args,**self.kwargs)

class WorkThreadClass(QObject):
    started = pyqtSignal()
    finished = pyqtSignal()

    def __init__(self, q, cancel_bool, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs
        self.q = q
        self.cancel_bool = cancel_bool
        super().__init__()

    @pyqtSlot()
    def run(self):
        while not self.q.empty():
            if not self.cancel_bool.get():
                pythoncom.CoInitialize()
                self.q.get()(*self.args, **self.kwargs)
                pythoncom.CoUninitialize()
            else:
                self.q.get()

class App(QMainWindow):

    def __init__(self,parent=None,timeout=None,main_wind=None,logger=None, icon=None):
        super().__init__()

        self.main_wind = main_wind
        self.lock_toast = threading.Lock()
        #model defined by json in file. Creates fields for each dictionary entry
        self.settings = DynamicModel("multi_comp_settings.cfg",os.getenv("APPDATA") + '\\Computer Info\\multi_comp_settings.cfg')
        #takes json of programs and creates a check list
        self.other_applications = ProgramChoices(["other_applications.prg"],default_folder=os.getenv("APPDATA") + '\\Computer Info',default_filename='other_applications.prg')

        self.logger = logger
        self.timeout = timeout
        self.icon = icon
        
        self.started_threads = []
        self.custom_user = ""
        self.custom_passwd = ""
        self.running = False
        self.execution_time = None
        self.info_queue = queue.Queue()
        self.cancel_bool = ThreadSafeBool()

        #widgets
        self.innerframe = QTabWidget(parent=self)

        self.createMainWidgets()
        self.showEditOptionsSidebar()

        self.logger.debug("********Running in Debug Mode********")
        self.scan_thread = QThread()
        self.scan_thread.start()

        self.gui_work_thread = QThread()
        self.gui_work_thread.start()

        self.testthread = QThread()
        self.testthread.start()
        

    def createMainWidgets(self):
        self.setWindowTitle("Computer Info")
        self.window_menu = self.menuBar()

        self.file_menu = self.window_menu.addMenu('File')

        self.other_credentials_button = QAction('Other Credentials', self)
        self.other_credentials_button.triggered.connect(self.getCredentials)
        self.file_menu.addAction(self.other_credentials_button)
        self.schedule_button = QAction("Delay Runtime", self)
        self.schedule_button.triggered.connect(self.getDelay)
        self.file_menu.addAction(self.schedule_button)
        self.exit_button = QAction('Exit', self)
        self.exit_button.setShortcut('Ctrl+Q')
        self.exit_button.triggered.connect(self.close)
        self.file_menu.addAction(self.exit_button)

        self.options_menu = self.window_menu.addMenu('Edit')

        self.push_shortcut_btn = QAction('Push Shortcut', self)
        self.push_shortcut_btn.setCheckable(True)
        self.push_shortcut_btn.triggered.connect(self.showEditOptionsSidebar)
        self.options_menu.addAction(self.push_shortcut_btn)

        self.find_scanners_btn = QAction('Find Scanners', self)
        self.find_scanners_btn.setCheckable(True)
        self.find_scanners_btn.triggered.connect(self.showEditOptionsSidebar)
        self.options_menu.addAction(self.find_scanners_btn)

        self.find_monitors_btn = QAction('Find Monitors (unreliable)', self)
        self.find_monitors_btn.setCheckable(True)
        self.find_monitors_btn.triggered.connect(self.showEditOptionsSidebar)
        self.options_menu.addAction(self.find_monitors_btn)

        self.find_printers_btn = QAction("Find Printers", self)
        self.find_printers_btn.setCheckable(True)
        self.find_printers_btn.triggered.connect(self.showEditOptionsSidebar)
        self.options_menu.addAction(self.find_printers_btn)

        self.find_apps_btn = QAction('Find Apps', self)
        self.find_apps_btn.setCheckable(True)
        self.find_apps_btn.triggered.connect(self.showEditOptionsSidebar)
        self.options_menu.addAction(self.find_apps_btn)

        self.install_app_btn = QAction('Install App', self)
        self.install_app_btn.setCheckable(True)
        self.install_app_btn.triggered.connect(self.showEditOptionsSidebar)
        self.options_menu.addAction(self.install_app_btn)

        self.options_menu.addSeparator()

        adv_options = QAction('Advanced Options', self)
        adv_options.triggered.connect(self.showSettingsWindow)
        self.options_menu.addAction(adv_options)

        self.help_menu = self.window_menu.addMenu('Help')

        about = QAction('About', self)
        about.triggered.connect(
            partial(QMessageBox.information, self, "About", "Computer Info\nVersion 2.3"))
        self.help_menu.addAction(about)

        self.containerWidget = QWidget()
        self.containerlayout = QVBoxLayout()
        self.containerlayout.setContentsMargins(10, 0, 0, 0)
        self.containerWidget.setLayout(self.containerlayout)
        self.centralWidget = QWidget(self.containerWidget)

        self.counterbox = QStatusBar()
        self.counterbox.showMessage('- Computers Left')
        self.counterbox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.containerlayout.addWidget(self.centralWidget)
        self.containerlayout.addWidget(self.counterbox)

        self.innermainlayout = QHBoxLayout()
        self.centralWidget.setLayout(self.innermainlayout)
        self.setCentralWidget(self.containerWidget)

        self.col_one = QWidget(self)
        self.col_one.setMinimumSize(400, 500)
        self.col_one_layout = QVBoxLayout()
        self.col_one.setLayout(self.col_one_layout)
        self.title_label = QLabel(
            "Input a list of computers to get details", self.col_one)
        self.subtitle_label = QLabel("", self.col_one)
        font = QFont()
        font.setPointSize(16)
        self.title_label.setFont(font)
        font2 = QFont()
        font2.setPointSize(10)
        self.subtitle_label.setFont(font2)

        self.waitingbox = QLabel(self.col_one)
        self.waitingbox.setFont(font)

        self.inbox = QPlainTextEdit(self.col_one)

        self.table_layout = QVBoxLayout()
        self.table_btn_layout = QHBoxLayout()
        self.table_frame = QWidget(self.col_one)
        self.table_tabs = QTabWidget(self.table_frame)
        self.table1 = QTableWidget()
        self.table1.cellDoubleClicked.connect(self.singleCompAppHook)
        self.table_tabs.insertTab(0, self.table1, "Computers")
        self.shortcuts_table = QTableWidget()
        self.scanners_table = QTableWidget()
        self.monitors_table = QTableWidget()
        self.printers_table = QTableWidget()
        self.install_apps_table = QTableWidget()
        self.find_apps_table = QTableWidget()
        self.find_apps_installs_table = QTableWidget()
        self.install_apps_table.cellDoubleClicked.connect(
            self.manualInstallOutput)
        self.find_apps_installs_table.cellDoubleClicked.connect(
            self.checkboxAppsInstallOutput)

        self.table1.verticalHeader().setVisible(False)
        self.shortcuts_table.verticalHeader().setVisible(False)
        self.scanners_table.verticalHeader().setVisible(False)
        self.monitors_table.verticalHeader().setVisible(False)
        self.printers_table.verticalHeader().setVisible(False)
        self.install_apps_table.verticalHeader().setVisible(False)
        self.find_apps_table.verticalHeader().setVisible(False)
        self.find_apps_installs_table.verticalHeader().setVisible(False)
        self.table_hide_btn = QPushButton("Close Table", self.table_frame)
        self.table_save_btn = QPushButton("Save Excel", self.table_frame)
        self.table_save_btn.setEnabled(False)
        self.table_hide_btn.clicked.connect(self.restoreInputBox)
        self.table_save_btn.clicked.connect(self.saveExcel)

        self.table_frame.setLayout(self.table_layout)
        self.table_layout.addWidget(self.table_tabs)

        self.table_btn_layout.addWidget(self.table_hide_btn)
        self.table_btn_layout.addWidget(self.table_save_btn)
        self.table_layout.addLayout(self.table_btn_layout)

        self.table_frame.hide()

        self.run_button = QPushButton('Start', self.col_one)
        self.run_button.clicked.connect(self.startScanFacilitator)
        self.run_button.setSizePolicy(
            QSizePolicy.Expanding, QSizePolicy.Minimum)

        self.table_unavailable_btn = QCheckBox(
            "Hide 'Unavailable'", self.col_one)
        self.table_unavailable_btn.stateChanged.connect(self.setRowsHidden)
        self.table_unavailable_btn.setSizePolicy(
            QSizePolicy.Minimum, QSizePolicy.Minimum)

        self.start_layout = QHBoxLayout()
        self.start_layout.addWidget(self.run_button)
        self.start_layout.addWidget(self.table_unavailable_btn)

        self.running_frame = QWidget()
        self.running_frame.setContentsMargins(0, 0, 0, 0)
        self.running_frame_layout = QHBoxLayout()
        self.running_frame.setLayout(self.running_frame_layout)
        self.prog = QProgressBar()
        self.prog.setAlignment(Qt.AlignCenter)
        self.cancelbtn = QPushButton("Cancel", self.running_frame)
        self.cancelbtn.setEnabled(False)
        self.cancelbtn.clicked.connect(self.cancelBtnAction)
        self.running_frame_layout.addWidget(self.prog)
        self.running_frame_layout.addWidget(self.cancelbtn)

        self.col_one_layout.addWidget(self.title_label)
        self.col_one_layout.addWidget(self.subtitle_label)
        self.col_one_layout.addWidget(self.waitingbox)
        self.col_one_layout.addWidget(self.inbox)
        self.col_one_layout.addLayout(self.start_layout)
        self.col_one_layout.addWidget(self.running_frame)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.subtitle_label.setAlignment(Qt.AlignCenter)
        self.waitingbox.setAlignment(Qt.AlignCenter)

        self.split = QHBoxLayout()
        self.split.addWidget(self.col_one)
        self.split.addWidget(self.innerframe)
        self.innerframe.setSizePolicy(
            QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.innermainlayout.addLayout(self.split)

        self.push_shortcut_tab = QWidget()
        self.push_shortcut_tab_layout = QVBoxLayout()
        self.push_shortcut_tab_layout.setAlignment(Qt.AlignTop)
        self.push_shortcut_tab.setLayout(self.push_shortcut_tab_layout)

        self.find_apps_tab = CustomScrollBox()
        self.find_apps_tab.setWidgetResizable(True)
        self.find_apps_tab.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.find_apps_tab_layout = QVBoxLayout()
        self.find_apps_tab_layout.setAlignment(Qt.AlignTop)
        self.find_apps_tab.setLayout(self.find_apps_tab_layout)
        self.innerframe.insertTab(2, self.find_apps_tab, "Find Apps")
        self.innerframe.removeTab(2)

        self.install_app_tab = QWidget()
        self.install_app_tab_layout = QVBoxLayout()
        self.install_app_tab_layout.setAlignment(Qt.AlignTop)
        self.install_app_tab.setLayout(self.install_app_tab_layout)

        self.check_form = ShortcutCheckboxForm(
            self.push_shortcut_tab, title="Choose where to place Shortcuts")
        self.push_shortcut_tab_layout.addWidget(self.check_form)

        self.shortcut_file_form = FileForm(
            extensionsallowed="Shortcut Files (*.url;*.lnk;*.exe;*.rdp)", title="Choose Shortcut File")
        self.push_shortcut_tab_layout.addWidget(self.shortcut_file_form)

        self.apps_form = AppsForm(programs_obj=self.other_applications)
        self.find_apps_tab.setWidget(self.apps_form)
        

        self.app_file_form = FileForm(
            extensionsallowed="VBScript, Powershell Script, Python Script (*.vbs *.ps1 *.py)", title="Choose Script File")
        self.install_app_tab_layout.addWidget(self.app_file_form)
        self.app_file_form.form_change.connect(self.get_extra_parameters)


        self.show()
        self.innerframe.show()

    def cancelBtnAction(self):
        """
        Sets cancel_bool to safely end work
        """
        self.cancel_bool.setTrue()
        self.cancelbtn.setText("Cancelling...")
        self.cancelbtn.setEnabled(False)

    def restoreInputBox(self):
        """
        Close Table and replace with input box
        """
        self.col_one_layout.replaceWidget(self.table_frame, self.inbox)
        self.table_frame.hide()
        self.inbox.show()
        self.table_save_btn.setEnabled(False)

    def showEditOptionsSidebar(self):
        """
        Displays tabs based off of edit menu checkboxes
        """
        self.innerframe.clear()
        if self.push_shortcut_btn.isChecked() or self.find_apps_btn.isChecked() or self.install_app_btn.isChecked():
            self.innerframe.show()

            if self.push_shortcut_btn.isChecked():
                self.innerframe.addTab(self.push_shortcut_tab, "Shortcuts")

            if self.find_apps_btn.isChecked():
                self.innerframe.addTab(self.find_apps_tab, "Find Apps")

            if self.install_app_btn.isChecked():
                self.innerframe.addTab(self.install_app_tab, "Install Apps")
        else:
            self.innerframe.hide()

    def showSettingsWindow(self):
        """
        Creates window for editing settings.
        """
        top = QDialog(self)
        top.setWindowTitle("Settings")
        top.setSizeGripEnabled(True)
        top_layout = QVBoxLayout()
        top.setLayout(top_layout)

        settings_form = DynamicForm(top,title="Settings",submit_callback=top.destroy,submit_callback_kwargs={},dynamicmodel=self.settings)
        top_layout.addWidget(settings_form)
        top_layout.setAlignment(settings_form,Qt.AlignTop)
        top.show()
        top.activateWindow()

    def setCredentials(self,user,passwd,top):
        self.custom_user=user.text()
        if self.custom_user:
            self.custom_passwd=passwd.text()
            self.setWindowTitle("Computer Info (User: %s)" % self.custom_user)
        else:
            self.setWindowTitle("Computer Info")
        top.close()
        
    def getCredentials(self):
        top = QDialog(self)
        top.setWindowTitle("Enter Credentials")
        top.setSizeGripEnabled(True)
        top_layout = QVBoxLayout()
        top.setLayout(top_layout)
        settings_form = QFormLayout()
        usernamefield = QLineEdit(self.custom_user)
        passwordfield = QLineEdit(self.custom_passwd)
        passwordfield.setEchoMode(QLineEdit.Password)
        submitbtn = QPushButton("Submit")
        settings_form.addRow("Username", usernamefield)
        settings_form.addRow("Password", passwordfield)
        settings_form.addRow(submitbtn)
        submitbtn.clicked.connect(partial(self.setCredentials, usernamefield, passwordfield, top))

        top_layout.addLayout(settings_form)
        top_layout.setAlignment(settings_form, Qt.AlignTop)
        top.show()
        top.activateWindow()

    def getDelay(self):
        top = QDialog(self)
        top.setWindowTitle("Enter Credentials")
        top.setSizeGripEnabled(True)
        top_layout = QVBoxLayout()
        top.setLayout(top_layout)
        settings_form = QVBoxLayout()
        datetimefield = QDateTimeEdit()
        if self.execution_time:
            datetimefield.setDateTime(self.execution_time)
        datetimefield.setMinimumDateTime(QDateTime.currentDateTime())
        submitbtn = QPushButton("Submit")
        settings_form.addWidget(datetimefield)
        settings_form.addWidget(submitbtn)
        submitbtn.clicked.connect(partial(self.printDT,datetimefield,top))

        top_layout.addLayout(settings_form)
        top_layout.setAlignment(settings_form, Qt.AlignTop)
        top.show()
        top.activateWindow()

    def printDT(self, val, top):
        self.execution_time = val.dateTime()
        self.subtitle_label.setText("Delay set to %s" % self.execution_time.toString("MM/dd/yyyy h:mm ap"))
        top.close()

    def countdownTime(self):
        cur = QDateTime.currentDateTime()
        timedelta = cur.secsTo(self.execution_time)
        while timedelta > 0 and not self.cancel_bool.get():
            self.testworker.send_update.emit(timedelta)
            cur = QDateTime.currentDateTime()
            timedelta = cur.secsTo(self.execution_time)
            QThread.sleep(1)
        if timedelta <= 0 or self.cancel_bool.get():
            self.testworker.finished.emit()

    @pyqtSlot(int)
    def updateTimeLabel(self, secs):
        hours, remainder = divmod(secs, 3600)
        minutes, seconds = divmod(remainder, 60)
        time_string = ""
        if hours > 0:
            time_string = time_string + "%dh " % hours
        if minutes > 0:
            time_string = time_string + "%dm " % minutes
        else:
            if hours > 0:
                time_string = time_string + "%dm " % minutes
        time_string = time_string + "%ds" % seconds
        self.waitingbox.setText("%s until execution\n(%s)" % (time_string, self.execution_time.toString("MM/dd/yyyy h:mm ap")))
        self.counterbox.showMessage("%s until execution (%s)" % (time_string, self.execution_time.toString("MM/dd/yyyy h:mm ap")))

        

    def singleCompAppHook(self,row,col):
        """
        Hooks into companion program passing along computer name.
        Requires registry entry that associates 'singlecomputerinfo' URI with Single Computer Info
        """
        if col == 1 or col == 2:
            os.startfile("singlecomputerinfo:%s" % self.table1.item(row,col).text())

    def toggleRunningState(self):
        """
        Enables and disables buttons while running or stopped and resets variables for new run
        """
        if not self.running:
            self.comp_info_objs = []
            self.excel_row_printer_count = 0
            self.table_row_printer_count = 0
            self.comp_obj_complete = {}
            self.table_hide_btn.setEnabled(False)
            self.table_save_btn.setEnabled(False)
            self.cancelbtn.setEnabled(True)
            self.run_button.setEnabled(False)
            self.run_button.show()
            self.inbox.setEnabled(False)
            self.options_menu.setEnabled(False)
            self.table_unavailable_btn.setEnabled(False)
            self.table_unavailable_btn.show()
            self.check_form.form_disable()
            self.shortcut_file_form.form_disable()
            self.apps_form.form_disable()
            self.app_file_form.form_disable()
            self.col_one_layout.replaceWidget(self.inbox,self.table_frame)
            self.inbox.hide()

            self.table_tabs.removeTab(1)
            self.table_tabs.removeTab(2)
            self.table_tabs.removeTab(3)
            self.table_tabs.removeTab(4)
            self.table_tabs.removeTab(5)
            self.table_tabs.removeTab(6)

            self.table1.setRowCount(1)
            self.shortcuts_table.setRowCount(1)
            self.scanners_table.setRowCount(1)
            self.monitors_table.setRowCount(1)
            self.printers_table.setRowCount(1)
            self.install_apps_table.setRowCount(1)
            self.find_apps_table.setRowCount(1)
            self.find_apps_installs_table.setRowCount(1)

            self.table_frame.show()
            self.start_time = time.time()
            self.running = True
            del self.info_queue
            self.info_queue = queue.Queue()

            try:
                self.install_script_name = os.path.splitext(os.path.basename(self.app_file_form.filename))[0].title()[:30]
            except:
                self.install_script_name = "Manual Install"

            if isinstance(self.install_script_name, str):
                self.install_script_name = self.install_script_name.replace("_"," ")

        else:
            self.table_hide_btn.setEnabled(True)
            self.waitingbox.hide()
            self.subtitle_label.show()
            self.subtitle_label.setText("")
            self.end_time = time.time()
            self.cancelbtn.setEnabled(False)
            self.run_button.setEnabled(True)
            self.run_button.show()
            self.inbox.setEnabled(True)
            self.options_menu.setEnabled(True)
            self.table_unavailable_btn.setEnabled(True)
            self.table_unavailable_btn.show()
            self.check_form.form_enable()
            self.shortcut_file_form.form_enable()
            self.apps_form.form_enable()
            self.app_file_form.form_enable()
            self.running = False
            self.execution_time = None

    def setWaitingState(self):
        """
        Shows waiting information
        """
        self.waitingbox.show()
        self.inbox.hide()
        self.subtitle_label.hide()
        self.run_button.hide()
        self.table_unavailable_btn.hide()

    def getCheckboxApps(self):
        """
        Retrieve checkbox values for apps
        """
        self.chosen_executes = []
        self.chosen_apps = []
        exe_widget_list = []
        app_widget_list = []
        for val in self.apps_form.widget_list:
            if val.get():
                app_widget_list.append(val.text.lower().strip())
            if val.sub_get():
                exe_widget_list.append(val.text.lower().strip())
        for val in self.other_applications.dict_list:
            if val['title'].lower().strip() in app_widget_list:
                self.chosen_apps.append(val)
            if val['title'].lower().strip() in exe_widget_list:
                self.chosen_executes.append(val)
        return (self.chosen_apps,self.chosen_executes)

    def setTableColumns(self):
        """
        Sets columns and sheets for excel output
        """
        self.table_tabs.clear()
        self.table_tabs.addTab(self.table1,"Computers")
        self.main_columns = ["Status","Name","IP Address","Serial","Model","Username","OS","Resolution","Monitors","CPU","Memory","Error","Profile Time", "Time Completed"]
        self.table1.setColumnCount(len(self.main_columns))
        self.table1.setHorizontalHeaderLabels(self.main_columns)

        if self.push_shortcut_btn.isChecked():
            self.table_tabs.addTab(self.shortcuts_table,"Shortcuts")
            self.icon_columns = ["Status","Name","Public Desktop","Startup Folder"]
            for i in self.settings.settings_dict['desktop profiles']:
                if i:
                    self.icon_columns.append(i.title() + " Desktop")
            self.icon_columns.append("Error")
            self.shortcuts_table.setColumnCount(len(self.icon_columns))
            self.shortcuts_table.setHorizontalHeaderLabels(self.icon_columns)

        if self.find_scanners_btn.isChecked():
            self.table_tabs.addTab(self.scanners_table,"Scanners")
            self.scanner_columns = ["Status","Name","IP Address","Scanners"]
            self.scanners_table.setColumnCount(len(self.scanner_columns))
            self.scanners_table.setHorizontalHeaderLabels(self.scanner_columns)

        if self.find_monitors_btn.isChecked():
            self.table_tabs.addTab(self.monitors_table,"Monitors")
            self.monitor_columns = ["Status","Name","IP Address","Monitors"]
            self.monitors_table.setColumnCount(len(self.monitor_columns))
            self.monitors_table.setHorizontalHeaderLabels(self.monitor_columns)

        if self.find_printers_btn.isChecked():
            self.table_tabs.addTab(self.printers_table,"Printers")
            self.printer_columns = ["Source PC Name","Name","PortName"]
            self.printers_table.setColumnCount(len(self.printer_columns))
            self.printers_table.setHorizontalHeaderLabels(self.printer_columns)

        if self.install_app_btn.isChecked():
            try:
                splitstring = re.split(r"[\/]",self.app_file_form.filename)
                tabname = splitstring[-1]
                tabname = tabname.replace(".vbs ","")
            except: tabname = "Install Apps"
            self.table_tabs.addTab(self.install_apps_table,tabname)
            self.install_columns = ["Status","Name","IP Address","Result"]
            self.install_apps_table.setColumnCount(len(self.install_columns))
            self.install_apps_table.setHorizontalHeaderLabels(self.install_columns)

        self.chosen_apps = None
        self.chosen_executes = None
        if self.find_apps_btn.isChecked():

            self.chosen_apps, self.chosen_executes = self.getCheckboxApps()

            self.table_tabs.addTab(self.find_apps_table,"Find Apps")
            if self.chosen_executes:
                self.table_tabs.addTab(self.find_apps_installs_table,"Install Missing Apps")

            self.apps_columns = ["Status","Name","IP Address"]
            self.exes_columns = ["Status","Name","IP Address"]
            for p in self.chosen_apps:
                self.apps_columns.append(p['title'].title())
            for p in self.chosen_executes:
                self.exes_columns.append(p['title'].title())
            self.find_apps_table.setColumnCount(len(self.apps_columns))
            self.find_apps_table.setHorizontalHeaderLabels(self.apps_columns)
            self.find_apps_installs_table.setColumnCount(len(self.exes_columns))
            self.find_apps_installs_table.setHorizontalHeaderLabels(self.exes_columns)

    def queueThreads(self):
        """
        Starts new threads for each computer up to the 'thread clusters' number specified in the settings.
        """
        self.worker_threads = []
        max_threads = int(self.settings.settings_dict.get('thread clusters', 0))
        if not max_threads:
            max_threads = QThread.idealThreadCount()
            
        self.logger.debug("Starting up %s threads" % max_threads)

        if not self.started_threads:
            for _ in range(0,max_threads):
                self.started_threads.append(QThread())
                self.started_threads[-1].start()

        # if len(self.started_threads) < max_threads:
        #     for _ in range(0,max_threads-len(self.started_threads)):
        #         self.started_threads.append(QThread())
        #         self.started_threads[-1].start()

        # elif len(self.started_threads) > max_threads:
        #     for st in self.started_threads[max_threads:]:
        #         st.quit()
        #         del st
        #     del self.started_threads[max_threads:]
        #     print("gt Threads: %s" % len(self.started_threads))


        for st in self.started_threads:
            self.worker_threads.append(WorkThreadClass(self.threads, self.cancel_bool))
            self.worker_threads[-1].moveToThread(st)
            self.worker_threads[-1].started.connect(self.worker_threads[-1].run)
            self.worker_threads[-1].started.emit()
            
    @pyqtSlot(str)
    def get_extra_parameters(self, script_file):
        self.app_file_form.remove_fields()
        if str(script_file).endswith(".py"):
            line = ""
            with open(script_file,"r") as f:
                line = f.readline()
            if line.startswith("\""):
                all_split = re.split(r';', line.replace("\"","").strip())
                self.extra_parameters = all_split
                for param in self.extra_parameters:
                    self.app_file_form.add_field(param)
            else:
                self.extra_parameters = []
        else:
            self.extra_parameters = []


    @pyqtSlot(int,int)
    def updateProgressBar(self,val,max_val):
        """
        Updates progress bar and remaining computer count
        """
        self.prog.setValue(max_val-val)
        percent = self.prog.value()
        avg_time = safe_divide((time.time() - self.start_time),int(percent))

        remaining_time = format_time(int(avg_time*val))

        if remaining_time.strip() == "": remaining_time = "-"

        if percent != max_val:
            self.counterbox.showMessage("%s Computers Left (%s remaining)" % (str(val),remaining_time))
        else:
            self.counterbox.showMessage("Done.")

    @pyqtSlot(int)
    def initializeProgressUI(self,val):
        """
        Shows relevant tables and initializes percentage and counter
        """
        self.prog.setMaximum(self.count.get())

        self.table1.setRowCount(0)
        self.shortcuts_table.setRowCount(0)
        self.scanners_table.setRowCount(0)
        self.monitors_table.setRowCount(0)
        self.printers_table.setRowCount(0)
        self.install_apps_table.setRowCount(0)
        self.find_apps_table.setRowCount(0)
        self.find_apps_installs_table.setRowCount(0)

        self.table1.setRowCount(self.count.get())
        self.shortcuts_table.setRowCount(self.count.get())
        self.scanners_table.setRowCount(self.count.get())
        self.monitors_table.setRowCount(self.count.get())
        self.install_apps_table.setRowCount(self.count.get())
        self.find_apps_table.setRowCount(self.count.get())
        self.find_apps_installs_table.setRowCount(self.count.get())

        for c in self.master_pc_list:
            table1_temp_item = QTableWidgetItem("Queued")
            table1_temp_item.setBackground(QBrush(QColor("#FFFFD0")))
            table1_temp_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)

            self.table1.setItem(c['count'],next((i for i,x in enumerate(self.main_columns) if x.lower().strip() == "status"),None),table1_temp_item)

            table1_temp_name = QTableWidgetItem(c['name'])
            table1_temp_name.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            self.table1.setItem(c['count'],next((i for i,x in enumerate(self.main_columns) if x.lower().strip() == "name"),None),table1_temp_name)

            if self.push_shortcut_btn.isChecked():
                shortcuts_table_temp_item = QTableWidgetItem(c['name'])
                shortcuts_table_temp_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.shortcuts_table.setItem(c['count'],next((i for i,x in enumerate(self.icon_columns) if x.lower().strip() == "name"),None),shortcuts_table_temp_item)

            if self.find_scanners_btn.isChecked():
                scanners_table_temp_item = QTableWidgetItem(c['name'])
                scanners_table_temp_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.scanners_table.setItem(c['count'],next((i for i,x in enumerate(self.scanner_columns) if x.lower().strip() == "name"),None),scanners_table_temp_item)

            if self.find_monitors_btn.isChecked():
                monitors_table_temp_item = QTableWidgetItem(c['name'])
                monitors_table_temp_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.monitors_table.setItem(c['count'],next((i for i,x in enumerate(self.monitor_columns) if x.lower().strip() == "name"),None),monitors_table_temp_item)

            if self.install_app_btn.isChecked():
                install_apps_table_temp_item = QTableWidgetItem(c['name'])
                install_apps_table_temp_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.install_apps_table.setItem(c['count'],next((i for i,x in enumerate(self.install_columns) if x.lower().strip() == "name"),None),install_apps_table_temp_item)

            if self.find_apps_btn.isChecked():
                find_apps_table_temp_item = QTableWidgetItem(c['name'])
                find_apps_table_temp_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.find_apps_table.setItem(c['count'],next((i for i,x in enumerate(self.apps_columns) if x.lower().strip() == "name"),None),find_apps_table_temp_item)
                find_apps_install_table_temp_item = QTableWidgetItem(c['name'])
                find_apps_install_table_temp_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.find_apps_installs_table.setItem(c['count'],next((i for i,x in enumerate(self.exes_columns) if x.lower().strip() == "name"),None),find_apps_install_table_temp_item)
        self.counterbox.showMessage("%s Computers Left (- remaining)" % (str(val)))

    @pyqtSlot()
    def finalizeProgress(self):
        """
        Toggles running state and fixes any PCs that may have been skipped over due to error.
        Then posts a toast notification that the application has completed on all hosts.
        """
        self.toggleRunningState()
        
        self.fixBlanks()
        temp_t = threading.Thread(target=self.postToast,daemon=True)
        temp_t.start()
        self.table_save_btn.setEnabled(True)
        self.setSummary()
        if self.cancel_bool.get():
            self.cancel_bool.setFalse()
            self.cancelbtn.setText("Cancel")

    @pyqtSlot(int, dict, dict, dict, dict, list, dict, dict, dict)
    def updateCounts(self,row,temp_dict,temp_icon_dict,temp_scanner_dict,temp_monitor_dict,temp_printer_dict_list,temp_manual_app_dict,temp_checkbox_apps_dict,temp_checkbox_exes_dict):
        """
        Adds items to tables and formats them appropriately
        """
        def add_items(table,columns,input_dict,color_scheme=[]):
            temp_list = [t.lower() for t in columns]
            for k,v in input_dict.items():
                col = temp_list.index(k.lower().strip())
                table_item = QTableWidgetItem(str(v))
                if k.lower().strip() == "status" and v.lower().strip() == "online":
                    table_item.setBackground(QBrush(QColor("#98FB98")))
                    table.setItem(row,col,table_item)
                elif k.lower().strip() == "status" and v.lower().strip() == "unavailable":
                    table_item.setBackground(QBrush(QColor("#FFDDD0")))
                    table.setItem(row,col,table_item)
                else:
                    if color_scheme:
                        for c in color_scheme:
                            if str(v).lower().strip() == c[0].lower().strip():
                                table_item.setBackground(QBrush(QColor(c[1])))
                    table.setItem(row,col,table_item)
                table_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            table.resizeColumnsToContents()

        def add_multiple_items(table,columns,input_dict_list):
            temp_list = [t.lower() for t in columns]
            for input_dict in input_dict_list:
                for k,v, in input_dict.items():
                    col = temp_list.index(k.lower().strip())
                    table_item = QTableWidgetItem(str(v))
                    self.printers_table.setRowCount(self.table_row_printer_count+1)
                    table.setItem(self.table_row_printer_count,col,table_item)
                    table_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.table_row_printer_count += 1
            table.resizeColumnsToContents()


        add_items(self.table1,self.main_columns,temp_dict)
        try:
            add_items(self.shortcuts_table,self.icon_columns,temp_icon_dict,color_scheme=[('Done',"#98FB98")])
        except AttributeError:pass
        except Exception: self.logger.debug("Failed adding item to table", exc_info=True)
        try:
            add_items(self.scanners_table,self.scanner_columns,temp_scanner_dict)
        except AttributeError:pass
        except Exception: self.logger.debug("Failed adding item to table", exc_info=True)
        try:
            add_items(self.monitors_table,self.monitor_columns,temp_monitor_dict)
        except AttributeError:pass
        except Exception: self.logger.debug("Failed adding item to table", exc_info=True)
        try:
            add_multiple_items(self.printers_table,self.printer_columns,temp_printer_dict_list)
        except AttributeError:pass
        except Exception: self.logger.debug("Failed adding item to table", exc_info=True)
        try:
            add_items(self.install_apps_table,self.install_columns,temp_manual_app_dict,color_scheme=[('Success',"#98FB98"),('Already Installed',"#F0F8FF")])
        except AttributeError:pass
        except Exception: self.logger.debug("Failed adding item to table", exc_info=True)
        try:
            add_items(self.find_apps_table,self.apps_columns,temp_checkbox_apps_dict,color_scheme=[('Success',"#98FB98"),('Already Installed',"#F0F8FF")])
        except AttributeError:pass
        except Exception: self.logger.debug("Failed adding item to table", exc_info=True)
        try:
            add_items(self.find_apps_installs_table,self.exes_columns,temp_checkbox_exes_dict,color_scheme=[('Success',"#98FB98"),('Already Installed',"#F0F8FF")])
        except AttributeError:pass
        except Exception:
            self.logger.debug("Failed adding item to table")

        if self.table_unavailable_btn.isChecked() and self.table1.item(row,0).text().lower().strip() == "unavailable":
            self.table1.setRowHidden(row,True)

            if self.shortcuts_table.rowCount():
                self.shortcuts_table.setRowHidden(row,True)

            if self.scanners_table.rowCount():
                self.scanners_table.setRowHidden(row,True)

            if self.monitors_table.rowCount():
                self.monitors_table.setRowHidden(row,True)

            if self.install_apps_table.rowCount():
                self.install_apps_table.setRowHidden(row,True)

            if self.find_apps_table.rowCount():
                self.find_apps_table.setRowHidden(row,True)

            if self.find_apps_installs_table.rowCount():
                self.find_apps_installs_table.setRowHidden(row,True)

    def startScanFacilitator(self):        
        self.waitingbox.show()
        self.subtitle_label.hide()

        self.toggleRunningState()
        if self.cancel_bool.get():
            self.toggleRunningState()
            self.restoreInputBox()
            self.cancel_bool.setFalse()
            self.cancelbtn.setText("Cancel")
            return

        if self.push_shortcut_btn.isChecked() and not self.shortcut_file_form.filename:
            QMessageBox.critical(self, "Icon File Missing", "Please choose an icon file to push")
            self.toggleRunningState()
            self.restoreInputBox()
            return
        if self.install_app_btn.isChecked() and not self.app_file_form.filename:
            QMessageBox.critical(self, "Installer Missing", "Please choose a vbs file to apply to PCs")
            self.toggleRunningState()
            self.restoreInputBox()
            return

        if self.execution_time:
            self.setWaitingState()
            self.table_frame.hide()
            self.testworker = TestThreadClass(self.countdownTime)
            self.testworker.moveToThread(self.testthread)
            self.testworker.started.connect(self.testworker.run)
            self.testworker.send_update.connect(self.updateTimeLabel)
            self.testworker.finished.connect(self.startScan)
            self.testworker.started.emit()
        else:
            self.startScan()
            

    def startScan(self):
        """
        Activated by pressing start button. Creates a GUIThreadClass Object with all necessary options.
        Connects thread to signals to change UI after each PC completes.
        """
        self.table_frame.show()
        if self.cancel_bool.get():
            self.toggleRunningState()
            self.restoreInputBox()
            self.cancel_bool.setFalse()
            self.cancelbtn.setText("Cancel")
            return

        self.threads = queue.Queue()
        self.countdown = ThreadSafeCounter()
        self.count = ThreadSafeCounter()

        self.master_pc_list = []
        
        self.setTableColumns()

        self.wt = GuiThreadClass(self.getComputerNames,
                fullbox = self.inbox.toPlainText().splitlines(),
                icon = self.push_shortcut_btn.isChecked(),
                get_devices = self.find_scanners_btn.isChecked(),
                get_monitors = self.find_monitors_btn.isChecked(),
                get_printers = self.find_printers_btn.isChecked(),
                other_profiles = self.check_form.check1.isChecked(),
                public_check = self.check_form.check2.isChecked(),
                startup_check = self.check_form.check3.isChecked(),
                get_apps = self.find_apps_btn.isChecked(),
                install_app = self.install_app_btn.isChecked(),
        )
        self.wt.moveToThread(self.scan_thread)
        self.wt.started.connect(self.wt.run)
        self.wt.progress_update.connect(self.updateProgressBar)
        self.wt.configure_prog.connect(self.initializeProgressUI)
        self.wt.complete_run.connect(self.finalizeProgress)
        self.wt.summary_dict.connect(self.updateCounts)
        self.wt.started.emit()

        self.counterbox.showMessage('Queuing Threads (May take some time with many PCs)')

    def getComputerNames(self,**kwargs):
        """
        Main method for handling objects and starting each query. It then updates the spreadsheet.
        """
        #toggle running state then check if folders
        fullbox = kwargs.pop('fullbox',None)
        icon = kwargs.pop('icon',None)
        get_devices = kwargs.pop('get_devices',None)
        get_monitors = kwargs.pop('get_monitors',None)
        get_printers = kwargs.pop('get_printers',None)
        other_profiles = kwargs.pop('other_profiles',None)
        public_check = kwargs.pop('public_check',None)
        startup_check = kwargs.pop('startup_check',None)
        get_apps = kwargs.pop('get_apps',None)
        install_app = kwargs.pop('install_app',None)

        self.summary = OrderedDict()
        self.summary['totals'] = {'success':0,'total computers':0}
        self.summary['icons'] = {}
        self.summary['scanners'] = {}
        self.summary['apps found'] = {}
        self.summary['apps installed'] = {}

        self.workbook = None
        self.computers_key = None
        self.workbook = CIWorkbook()

        self.summary_key = self.workbook.new_summary("Summary")

        self.computers_key = self.workbook.new_sheet("Computers",columns=self.main_columns)
        self.summary['totals']['success'] = 0

        extra_parameters = []
        if install_app:
            single_app_install = self.app_file_form.filename
            extra_parameters = self.app_file_form.get_field_list()
        else:
            single_app_install = None

        #expand fullbox with subnets here
        for i,l in enumerate(fullbox):
            pattern = re.compile(r'^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}/[0-9]{1,2}$')
            pmatch = pattern.match(l.strip())
            if pmatch:
                self.logger.debug("Found match for %s" % l.strip())
                fullbox = fullbox[:i] + [str(t) for t in ipaddress.IPv4Network(l.strip()).hosts()] + fullbox[i+1:]

        for line in fullbox:
            if line.strip() != "":
                self.master_pc_list.append({'name':line.strip(),'count':self.threads.qsize()})
                self.summary['totals']['total computers'] += 1
                self.comp_info_objs.append(
                    ComputerInfo(
                        q=self.info_queue,
                        input_name=line.strip(),
                        count=self.threads.qsize(),
                        icon=icon,
                        get_devices=get_devices,
                        get_monitors=get_monitors,
                        other_profiles=other_profiles,
                        profile_list=self.settings.settings_dict.get('desktop profiles',[]),
                        public=public_check,
                        startup=startup_check,
                        shortcut_filename=self.shortcut_file_form.filename,
                        input_domain = self.settings.settings_dict.get('domain',''),
                        input_group = self.settings.settings_dict.get('group',''),
                        get_apps=get_apps,
                        other_applications = self.chosen_apps,
                        install_applications = self.chosen_executes,
                        single_app_install = single_app_install,
                        extra_parameters = extra_parameters,
                        get_printers = get_printers,
                        logger = self.logger,
                        profile = True,
                        manual_user = self.custom_user,
                        manual_pass = self.custom_passwd
                    )
                )
                #t = WMIThread(target=self.comp_info_objs[-1].get_info,daemon=True)
                self.threads.put(self.comp_info_objs[-1].get_info)

        self.countdown.set(self.threads.qsize())
        self.count.set(self.threads.qsize())

        self.wt.configure_prog.emit(self.countdown.get())
        self.queueThreads()

        #Set queue timeout for computers that hang WMI
        if self.timeout:
            timeout = self.timeout
        else:
            if get_apps:
                timeout = 500
            elif install_app:
                temp_timeout = self.settings.settings_dict.get('install timeout',None)
                try:
                    if temp_timeout and int(temp_timeout) > 1000:
                        timeout = int(temp_timeout)
                    else:
                        timeout = 1000
                except:
                    timeout = 1000
            else:
                timeout = 300

        while self.countdown.get() > 0 and not self.cancel_bool.get():
            item = None
            temp_dict = {}
            temp_icon_dict = {}
            temp_scanner_dict = {}
            temp_monitor_dict = {}
            temp_printer_dict_list = []
            temp_manual_app_dict = {}
            temp_checkbox_apps_dict = {}
            temp_checkbox_exes_dict = {}
            try:
                if self.cancel_bool.get():
                    timeout = 0

                item = self.info_queue.get(timeout=timeout)
                self.comp_obj_complete[item.count] = item

                self.workbook.set_working_sheet(self.computers_key)
                if not item.status:
                    temp_dict = {"status":"Online",'name':item.name,'serial':item.serial,'model':item.model,'username':item.user,'os':item.os,'cpu':item.cpu,'memory':item.memory,'error':item.status,'monitors':str(item.monitors)}

                    try: temp_dict['ip address'] = item.ip_addresses[0]
                    except: temp_dict['ip address'] = ""

                    if item.resolution:
                        temp_dict['resolution'] = item.resolution.replace("\n",", ")

                    if item.profile_time:
                        temp_dict['profile time'] = "%.2f seconds" % item.profile_time

                    temp_dict['time completed'] = datetime.now()

                    self.workbook.working_sheet.add_row(temp_dict,row=item.count+2)
                    self.summary['totals']['success'] += 1

                    if icon:
                        self.workbook.set_or_create_worksheet("Icon Push",columns=self.icon_columns)
                        temp_icon_dict = {'status':temp_dict.get('status',"Unknown"),'name':temp_dict.get('name',item.input_name)}
                        for k,v in item.paths.items():
                            temp_icon_dict[k] = v['result']
                        self.workbook.working_sheet.add_row(temp_icon_dict,row=item.count+2)
                        if self.summary['icons'] == {}:
                            self.summary['icons']['pushed'] = 1
                        else:
                            self.summary['icons']['pushed'] += 1

                    if get_devices:
                        self.workbook.set_or_create_worksheet("Scanners",columns=self.scanner_columns)
                        scanner_text = ""
                        for s in list(set(item.devices)):
                            if s and "fi-" in s:
                                scanner_text += s + "\n"
                        temp_scanner_dict = {'status':temp_dict.get('status',"Unknown"),'name':item.name,'ip address':temp_dict['ip address'],"scanners":scanner_text}
                        if not item.ip_addresses:
                            temp_scanner_dict['name'] = item.input_name
                        self.workbook.working_sheet.add_row(temp_scanner_dict,row=item.count+2)

                        if self.summary['scanners'] == {}:
                            if scanner_text:
                                self.summary['scanners']['count'] = 1
                            else:
                                self.summary['scanners']['count'] = 0
                        else:
                            if scanner_text:
                                self.summary['scanners']['count'] += 1

                    if get_monitors:
                        self.workbook.set_or_create_worksheet("Monitors",columns=self.monitor_columns,wrap=True)
                        monitor_text = ""
                        if item.monitors_detail:
                            for s in list(set(item.monitors_detail)):
                                monitor_text += s + "\n"
                        temp_monitor_dict = {'status':temp_dict.get('status',"Unknown"),'name':item.name,'ip address':temp_dict['ip address'],"monitors":monitor_text}
                        if not item.ip_addresses:
                            temp_monitor_dict['name'] = item.input_name
                        self.workbook.working_sheet.add_row(temp_monitor_dict,row=item.count+2)

                    if get_printers:
                        self.workbook.set_or_create_worksheet("Printers",columns=self.printer_columns,wrap=True)
                        if item.printers:
                            for p in item.printers:
                                temp_printer_dict = {'source pc name':item.input_name,'name':p.printer,'portname':p.port}
                                temp_printer_dict_list.append(temp_printer_dict)
                                self.workbook.working_sheet.add_row(temp_printer_dict,row=self.excel_row_printer_count + 2)
                                self.excel_row_printer_count += 1
                    if install_app:
                        if not self.install_script_name in self.summary['apps installed'] or not type(self.summary['apps installed'][self.install_script_name]) is int:
                            self.summary['apps installed'][self.install_script_name] = 0

                        if type(item.single_install_status) is str and "success" in item.single_install_status.lower():
                            self.summary['apps installed'][self.install_script_name] += 1

                        self.workbook.set_or_create_worksheet(self.install_script_name,columns=self.install_columns,wrap=True)
                        temp_manual_app_dict = {'status':temp_dict.get('status',"Unknown"),'name':item.name,'ip address':temp_dict['ip address'],'result':item.single_install_status}
                        if not item.ip_addresses:
                            temp_manual_app_dict['name'] = item.input_name
                        self.workbook.working_sheet.add_row(temp_manual_app_dict,row=item.count+2)

                    if get_apps:
                        self.workbook.set_or_create_worksheet("Apps",columns=self.apps_columns,wrap=True)
                        temp_checkbox_apps_dict = {'status':temp_dict.get('status',"Unknown"),'name':item.name,'ip address':temp_dict['ip address']}
                        if item.ip_addresses:
                            temp_checkbox_apps_dict['name'] = item.input_name
                        for k,v in item.found_apps.items():
                            if k.lower() in self.summary['apps found']:
                                if v and str(v).strip():
                                    self.summary['apps found'][k.lower()] += 1
                            else:
                                if v:
                                    self.summary['apps found'][k.lower()] = 1
                                else:
                                    self.summary['apps found'][k.lower()] = 0
                            if v:
                                if type(v) is list:
                                    temp_checkbox_apps_dict[k.lower()] = str(v[-1])
                                elif not type(v) is str:
                                    temp_checkbox_apps_dict[k.lower()] = str(v)
                                else:
                                    temp_checkbox_apps_dict[k.lower()] = v

                        self.workbook.working_sheet.add_row(temp_checkbox_apps_dict,row=item.count+2)

                        temp_checkbox_exes_dict = {'status':temp_dict.get('status',"Unknown"),'name':item.name,'ip address':temp_dict['ip address']}
                        if item.ip_addresses:
                            temp_checkbox_exes_dict['name'] = item.input_name

                        for k,v in item.install_status.items():
                            if v:
                                if not k.lower() in self.summary['apps installed']:
                                    self.summary['apps installed'][k.lower()] = None
                                if not self.summary['apps installed'][k.lower()] or not type(self.summary['apps installed'][k.lower()]) is int:
                                    self.summary['apps installed'][k.lower()] = 0
                                if type(v) is list:
                                    temp_checkbox_exes_dict[k.lower()] = v[-1]
                                elif not type(v) is str:
                                    temp_checkbox_exes_dict[k.lower()] = str(v)
                                else:
                                    temp_checkbox_exes_dict[k.lower()] = v
                            elif v == 0:
                                temp_checkbox_exes_dict[k.lower()] = "Success"
                                if k.lower() in self.summary['apps installed']:
                                    self.summary['apps installed'][k.lower()] += 1
                                else:
                                    self.summary['apps installed'][k.lower()] = 1
                            self.logger.debug("installer: %s" % k)
                        if self.chosen_executes:
                            self.workbook.set_or_create_worksheet("Installs",columns=self.exes_columns,wrap=True)
                            self.workbook.working_sheet.add_row(temp_checkbox_exes_dict,row=item.count+2)

                else:
                    if self.cancel_bool.get():
                        temp_dict = {'status':"Cancelled",'name':item.input_name,'error':"Cancelled"}
                        try:
                            self.workbook.set_working_sheet(self.computers_key)
                            self.workbook.working_sheet.add_row(temp_dict,row=item.count+2)

                            if icon:
                                self.workbook.set_or_create_worksheet("Icon Push",columns=self.icon_columns)
                                self.workbook.working_sheet.add_row({'status':"Cancelled",'name':item.input_name},row=item.count+2)
                            if get_devices:
                                self.workbook.set_or_create_worksheet("Scanners",columns=self.scanner_columns)
                                self.workbook.working_sheet.add_row({'status':"Cancelled",'name':item.input_name},row=item.count+2)
                            if get_monitors:
                                self.workbook.set_or_create_worksheet("Monitors",columns=self.monitor_columns)
                                self.workbook.working_sheet.add_row({'status':"Cancelled",'name':item.input_name},row=item.count+2)
                            if install_app:
                                self.workbook.set_or_create_worksheet(self.install_script_name,columns=self.install_columns)
                                self.workbook.working_sheet.add_row({'status':"Cancelled",'name':item.input_name},row=item.count+2)
                            if get_apps:
                                self.workbook.set_or_create_worksheet("Apps",columns=self.apps_columns)
                                self.workbook.working_sheet.add_row({'status':"Cancelled",'name':item.input_name},row=item.count+2)
                                if self.chosen_executes:
                                    self.workbook.set_or_create_worksheet("Installs",columns=self.exes_columns)
                                    self.workbook.working_sheet.add_row({'status':"Cancelled",'name':item.input_name},row=item.count+2)
                        except Exception as e:
                            self.logger.debug("Error writing unavailable pc to sheet:", exc_info=True)
                    else:
                        temp_dict = {'status':"Unavailable",'name':item.input_name,'error':item.status}
                        try:
                            self.workbook.set_working_sheet(self.computers_key)
                            self.workbook.working_sheet.add_row(temp_dict,row=item.count+2)

                            if icon:
                                self.workbook.set_or_create_worksheet("Icon Push",columns=self.icon_columns)
                                self.workbook.working_sheet.add_row({'status':"Unavailable",'name':item.input_name},row=item.count+2)
                            if get_devices:
                                self.workbook.set_or_create_worksheet("Scanners",columns=self.scanner_columns)
                                self.workbook.working_sheet.add_row({'status':"Unavailable",'name':item.input_name},row=item.count+2)
                            if get_monitors:
                                self.workbook.set_or_create_worksheet("Monitors",columns=self.monitor_columns)
                                self.workbook.working_sheet.add_row({'status':"Unavailable",'name':item.input_name},row=item.count+2)
                            if install_app:
                                self.workbook.set_or_create_worksheet(self.install_script_name,columns=self.install_columns)
                                self.workbook.working_sheet.add_row({'status':"Unavailable",'name':item.input_name},row=item.count+2)
                            if get_apps:
                                self.workbook.set_or_create_worksheet("Apps",columns=self.apps_columns)
                                self.workbook.working_sheet.add_row({'status':"Unavailable",'name':item.input_name},row=item.count+2)
                                if self.chosen_executes:
                                    self.workbook.set_or_create_worksheet("Installs",columns=self.exes_columns)
                                    self.workbook.working_sheet.add_row({'status':"Unavailable",'name':item.input_name},row=item.count+2)
                        except Exception:
                            self.logger.debug("Error writing unavailable pc to sheet:", exc_info=True)

            except queue.Empty:
                self.logger.debug("Queue Timeout")
            except Exception:
                self.logger.debug("Unable to get computer info from queue:", exc_info=True)

                try:
                    self.workbook.set_working_sheet(self.computers_key)
                    temp_dict = {'status':"Unavailable",'name':item.input_name,'error':item.status}
                    self.workbook.working_sheet.add_row(temp_dict,row=item.count+2)
                    if icon:
                        self.workbook.set_or_create_worksheet("Icon Push",columns=self.icon_columns)
                        self.workbook.working_sheet.add_row({'name':item.input_name,'error':e},row=item.count+2)
                except Exception:
                    self.logger.debug(
                        "Failed setting PC unavailable in worksheet", exc_info=True)
            finally:
                self.countdown.decrement()
                self.wt.progress_update.emit(self.countdown.get(),self.count.get())
                if item:

                    self.wt.summary_dict.emit(
                                    item.count,
                                    temp_dict,
                                    temp_icon_dict,
                                    temp_scanner_dict,
                                    temp_monitor_dict,
                                    temp_printer_dict_list,
                                    temp_manual_app_dict,
                                    temp_checkbox_apps_dict,
                                    temp_checkbox_exes_dict,
                                    )
        #     w.stop()
        self.wt.complete_run.emit()

    def postToast(self):
        """
        Creates a toast notification to let the user know the scan has completed
        """
        if self.lock_toast.acquire(timeout=15):
            try:
                toaster = ToastNotifier()
                if self.icon and self.icon.exists():
                    toaster.show_toast("Complete!","Scan/Install is complete.\n%s out of %s computers online" % (self.summary['totals']['success'],self.summary['totals']['total computers']),icon_path=self.icon,duration=10)
                else:
                    toaster.show_toast("Complete!","Scan/Install is complete.",duration=10)
            except Exception: self.logger.debug("Failed showing notification", exc_info=True)
            finally: self.lock_toast.release()
        else:
            self.logger.debug("Unable to get lock on toast. Giving up.")

    def fixBlanks(self):
        """
        Fixes cases where queue timesout and no information returns. Ensures that the name fields are filled
        """
        self.workbook.set_working_sheet(self.computers_key)
        for pc in self.master_pc_list:
            if not self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=1).value:
                if self.cancel_bool.get():
                    self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=1,value="Cancelled")
                else:
                    self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=1,value="Unavailable")
                self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=2,value=pc['name'])
                try:
                    if self.push_shortcut_btn.isChecked():
                        self.workbook.set_or_create_worksheet("Icon Push",columns=self.icon_columns)
                        self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=2,value=pc['name'])
                    if self.find_scanners_btn.isChecked():
                        self.workbook.set_or_create_worksheet("Scanners",columns=self.scanner_columns)
                        self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=2,value=pc['name'])
                    if self.find_monitors_btn.isChecked():
                        self.workbook.set_or_create_worksheet("Monitors",columns=self.monitor_columns)
                        self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=2,value=pc['name'])
                    if self.install_app_btn.isChecked():
                        self.workbook.set_or_create_worksheet(self.install_script_name,columns=self.install_columns)
                        self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=2,value=pc['name'])
                    if self.find_apps_btn.isChecked():
                        self.workbook.set_or_create_worksheet("Apps",columns=self.apps_columns)
                        self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=2,value=pc['name'])
                        if self.chosen_executes:
                            self.workbook.set_or_create_worksheet("Installs",columns=self.exes_columns)
                            self.workbook.working_sheet.sheet.cell(row=pc['count']+2,column=2,value=pc['name'])
                except Exception:
                    self.logger.debug("Error fixing timed out computer names:", exc_info=True)

        for r in range(1,self.table1.rowCount()+1):
            if self.table1.item(r,0) and self.table1.item(r,0).text() == "Queued":
                if self.cancel_bool:
                    self.table1.item(r,0).setText("Cancelled")
                else:
                    self.table1.item(r,0).setText("Queue Timeout")
        self.table1.resizeColumnsToContents()

    def setRowsHidden(self,stateval):
        """
        Used to hide rows where the computer was "Unavailable"
        """
        self.table_unavailable_btn.setEnabled(False)
        if stateval:
            hide_row = True
        else:
            hide_row = False
        for r in range(0,self.table1.rowCount()+1):
            if self.table1.item(r,0) and self.table1.item(r,0).text() == "Unavailable":
                self.table1.setRowHidden(r,hide_row)

        if self.shortcuts_table.rowCount():
            for r in range(0,self.shortcuts_table.rowCount()+1):
                if not self.shortcuts_table.item(r,0) or self.shortcuts_table.item(r,0) and self.shortcuts_table.item(r,0).text() != "Online":
                    self.shortcuts_table.setRowHidden(r,hide_row)

        if self.scanners_table.rowCount():
            for r in range(0,self.scanners_table.rowCount()+1):
                if not self.scanners_table.item(r,0) or self.scanners_table.item(r,0) and self.scanners_table.item(r,0).text() != "Online":
                    self.scanners_table.setRowHidden(r,hide_row)

        if self.monitors_table.rowCount():
            for r in range(0,self.monitors_table.rowCount()+1):
                if not self.monitors_table.item(r,0) or self.monitors_table.item(r,0) and self.monitors_table.item(r,0).text() != "Online":
                    self.monitors_table.setRowHidden(r,hide_row)

        if self.install_apps_table.rowCount():
            for r in range(0,self.install_apps_table.rowCount()+1):
                if not self.install_apps_table.item(r,0) or self.install_apps_table.item(r,0) and self.install_apps_table.item(r,0).text() != "Online":
                    self.install_apps_table.setRowHidden(r,hide_row)

        if self.find_apps_table.rowCount():
            for r in range(0,self.find_apps_table.rowCount()+1):
                if not self.find_apps_table.item(r,0) or self.find_apps_table.item(r,0) and self.find_apps_table.item(r,0).text() != "Online":
                    self.find_apps_table.setRowHidden(r,hide_row)

        if self.find_apps_installs_table.rowCount():
            for r in range(0,self.find_apps_installs_table.rowCount()+1):
                if not self.find_apps_installs_table.item(r,0) or self.find_apps_installs_table.item(r,0) and self.find_apps_installs_table.item(r,0).text() != "Online":
                    self.find_apps_installs_table.setRowHidden(r,hide_row)
        self.table_unavailable_btn.setEnabled(True)

    def manualInstallOutput(self, row, col):
        """
        Generates a window with the output of a vbs installer.
        VBS installers are vbs files that attempt to run commands on remote machines.
        The VBscript will run for each host specified and will send a parameter to the VBS with the given hostname.
        """
        try:
            top = QDialog(self)
            top.setWindowTitle("Install Output")
            top.setSizeGripEnabled(True)
            #top.setWindowModality(Qt.ApplicationModal)
            top_layout = QVBoxLayout()
            top.setLayout(top_layout)
            titlelabel = QLabel(self.comp_obj_complete[row].input_name.upper())
            font = QFont()
            font.setPointSize(16)
            titlelabel.setFont(font)
            top_layout.addWidget(titlelabel)
            top_layout.setAlignment(titlelabel, Qt.AlignCenter)

            installout = QLabel()
            installout.setStyleSheet("background-color:black;color:white")
            installout.setContentsMargins(10, 10, 10, 10)
            top_layout.addWidget(installout)
            top_layout.setAlignment(installout, Qt.AlignTop)
            try:
                installout.setText(
                    self.comp_obj_complete[row].out1.decode('utf-8'))
            except:
                installout.setText("                    ")

            try:
                if self.comp_obj_complete[row].out1_err.decode('utf-8'):
                    errout = QLabel(
                        self.comp_obj_complete[row].out1_err.decode('utf-8'))
                    errout.setStyleSheet(
                        "background-color:#8B0000;color:white")
                    errout.setContentsMargins(10, 10, 10, 10)
                    top_layout.addWidget(errout)
                    top_layout.setAlignment(errout, Qt.AlignTop)
            except:
                installout.setText("                    ")

            top.show()
            top.activateWindow()
        except:
            pass

    def checkboxAppsInstallOutput(self, row, col):
        """
        Similar to manualInstallOutput except shows output of installers for apps chosen from the "Find Apps" list.
        """
        try:
            top = QDialog(self)
            top.setWindowTitle("Install Output")
            top.setSizeGripEnabled(True)
            #top.setWindowModality(Qt.ApplicationModal)
            top_layout = QVBoxLayout()
            top.setLayout(top_layout)
            titlelabel = QLabel(self.comp_obj_complete[row].input_name.upper())
            font = QFont()
            font.setPointSize(16)
            titlelabel.setFont(font)
            top_layout.addWidget(titlelabel)
            top_layout.setAlignment(titlelabel, Qt.AlignCenter)

            try:
                installout = QLabel(
                    self.comp_obj_complete[row].out2.decode('utf-8'))
                installout.setStyleSheet("background-color:black;color:white")
                installout.setContentsMargins(10, 10, 10, 10)
                top_layout.addWidget(installout)
                top_layout.setAlignment(installout, Qt.AlignTop)
            except:
                pass

            try:
                if self.comp_obj_complete[row].out2_err.decode('utf-8'):
                    errout = QLabel(
                        self.comp_obj_complete[row].out2_err.decode('utf-8'))
                    errout.setStyleSheet(
                        "background-color:#8B0000;color:white")
                    errout.setContentsMargins(10, 10, 10, 10)
                    top_layout.addWidget(errout)
                    top_layout.setAlignment(errout, Qt.AlignTop)
            except:
                pass

            top.show()
            top.activateWindow()
        except:
            pass

    def setSummary(self):
        """
        Create table of success numbers
        """
        self.workbook.set_working_sheet(self.summary_key)

        success_percent = safe_divide(self.summary['totals']['success'],self.summary['totals']['total computers'])
        failure_percent = safe_divide((self.summary['totals']['total computers']-self.summary['totals']['success']),self.summary['totals']['total computers'])

        self.workbook.working_sheet.add_data("Successful Computers",self.summary['totals']['success'],"%.2f%%" % success_percent,format_value="0.00%")
        self.workbook.working_sheet.add_data("Failed Computers",self.summary['totals']['total computers']-self.summary['totals']['success'],"%.2f%%" % failure_percent,format_value="0.00%")
        self.workbook.working_sheet.add_data("Total Computers Attempted",self.summary['totals']['total computers'])

        self.workbook.working_sheet.blank_data()
        if not self.summary['icons'] == {}:
            icon_percent = safe_divide(self.summary['icons']['pushed'],self.summary['totals']['total computers'])
            self.workbook.working_sheet.add_data("Icons Pushed",self.summary['icons']['pushed'],"%.2f%%" % icon_percent,format_value="0.00%")
            self.workbook.working_sheet.blank_data()
        if not self.summary['scanners'] == {}:
            scanner_percent = safe_divide(self.summary['scanners']['count'],self.summary['totals']['total computers'])

            self.workbook.working_sheet.add_data("Scanners Found",self.summary['scanners']['count'],"%.2f%%" % scanner_percent,format_value="0.00%")

        apps_installed_args = [(k.title(),v,"%.2f%%" % (safe_divide(v,self.summary['totals']['total computers']))) for k,v in self.summary['apps found'].items()]
        install_status_args = [(k.title(),v) for k,v in self.summary['apps installed'].items()]
        if self.find_apps_btn.isChecked():
            self.workbook.working_sheet.add_grouping("Apps Found",*apps_installed_args)

        if self.install_app_btn.isChecked():
            self.workbook.working_sheet.add_grouping("Installs",*install_status_args)

        smtp_thread = threading.Thread(target=lambda:self.send_smtp(['mjanos@wphospital.org', 'wpinventory@wphospital.org'], "Scan Complete!<br>Success: %s (%.2f%%)<br>Failure: %s (%.2f%%)" % (self.summary['totals']['success'], success_percent*100, self.summary['totals']['total computers']-self.summary['totals']['success'], failure_percent*100)))
        smtp_thread.start()        

    def send_smtp(self, recipients, content):
        send_from = 'ComputerInfo@wphospital.org'
        try:
            S = SMTP('wphospital-org.mail.protection.outlook.com', 25)
            #S.set_debuglevel(1)
            S.connect('wphospital-org.mail.protection.outlook.com',25)
            S.ehlo()
            try:
                S.starttls()
            except:
                self.logger.debug("No support for starttls")
            S.ehlo()
            #S.login("mjanos@wphospital.org", getpass.getpass("Input password for mjanos: "))
            msg = MIMEText(content.encode('utf-8'), 'html', 'UTF-8')
            msg['Subject'] = "Run Complete"
            msg['From'] = formataddr(('ComputerInfo', send_from))
            msg['Content-Type'] = "text/html; charset=utf-8"
            msg['to'] = ";".join(recipients)
            S.sendmail(send_from,";".join(recipients),msg.as_string())
            S.quit()
        except Exception as e:
            raise

    def saveExcel(self):
        """
        Creates and saves data to an Excel spreadsheet
        """
        discard_sheet = None
        done_save = False
        while done_save == False:
            f, _ = QFileDialog.getSaveFileName(self,"Save as","","Excel Spreadsheet (*.xlsx)")
            if f == '':
                discard_sheet = QMessageBox.question(self,"Save Excel", "Discard Spreadsheet?")
                if discard_sheet == QMessageBox.Yes:
                     done_save = True
            else:
                try:
                    done_save = True
                    self.workbook.save(f)
                except PermissionError:
                    self.logger.debug("Permission Error: Cannot save", exc_info=True)
                    QMessageBox.critical(self,"Error","Cannot save.\nThis can be caused by one of the following:\n1. You do not have access to the folder.\n2. You are replacing the file but it is still open.")
                    done_save = False
                except Exception:
                    self.logger.debug("Unable to save file:", exc_info=True)
                    done_save = False
                finally:
                    if done_save:
                        final_time = self.end_time - self.start_time
                        final_hours = ""
                        final_mins = ""
                        final_secs = ""
                        if final_time// 3600: final_hours ="%s hours, " % int(final_time// 3600)
                        if (final_time % 3600)//60: final_mins = "%s minutes, " % int((final_time % 3600)//60)
                        if int((final_time % 3600)%60): final_secs = "%s seconds" % int((final_time % 3600)%60)
                        if not final_secs.strip() and not final_mins.strip() and not final_hours.strip(): final_secs = "less than a second"
                        open_excel = QMessageBox.question(self,"Open File", "Report completed in %s%s%s\nOpen the saved Excel Spreadsheet?" % (final_hours,final_mins,final_secs))
                        if open_excel == QMessageBox.Yes: os.startfile(f.replace("\\","\\\\").replace("/","\\"))

if __name__ == '__main__':
    logger = logging.getLogger()
    handler = logging.StreamHandler()
    logger.addHandler(handler)
    formatter = logging.Formatter('%(levelname)s: %(message)s')
    file_formatter = logging.Formatter('%(asctime)s %(levelname)s: %(message)s')
    handler.setFormatter(formatter)
    

    ico_path = ""
    try:
        ico_path = sys._MEIPASS
    except:
        ico_path = sys.argv[0]
    try:
        if not os.path.exists(os.getenv("APPDATA") + '\\Computer Info'):
            os.makedirs(os.getenv("APPDATA") + '\\Computer Info')
    except Exception:
        logger.exception("Issue creating Computer Info directory")

    try:
        if not os.path.exists(os.getenv("APPDATA") + '\\Computer Info\\multi_comp_settings.cfg'):
            copyfile(os.path.dirname(ico_path) + '\\multi_comp_settings.cfg',
                     os.getenv("APPDATA") + '\\Computer Info\\multi_comp_settings.cfg')
    except Exception:
        logger.exception("Issue copying multi_comp_settings.cfg")

    try:
        if not os.path.exists(os.getenv("APPDATA") + '\\Computer Info\\other_applications.prg'):
            copyfile(os.path.dirname(ico_path) + '\\other_applications.prg',
                     os.getenv("APPDATA") + '\\Computer Info\\other_applications.prg')
    except Exception:
        logger.exception("Issue copying other_applications.prg")

    file_handler = logging.FileHandler(os.getenv("APPDATA") + '\\Computer Info\\comp_info.log',mode='w')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)


    wind = QApplication(sys.argv)

    parser = argparse.ArgumentParser(description="Find computers, install apps, etc.")

    parser.add_argument('-timeout', type=int, help="Set timeout for each PC. If computer info not found or program not installed in specified time (seconds) computer is skipped.")
    parser.add_argument('-debug', default=False, action='store_true')
    parser.add_argument('-verbose', default=False, action='store_true')
    args = parser.parse_args()

    if not ctypes.windll.UxTheme.IsThemeActive():
        wind.setStyle('Fusion')

    file_handler.setLevel(logging.DEBUG)
    logger.setLevel(logging.DEBUG)

    if args.debug:
        handler.setLevel(logging.DEBUG)
    elif args.verbose:
        handler.setLevel(logging.INFO)
    else:
        handler.setLevel(logging.WARNING)

    if args.timeout:
        logger.info("Timeout set to %ss" % (args.timeout))

    app = App(timeout=args.timeout,main_wind=wind,logger=logger, icon=Path(ico_path).parent.joinpath("logo.ico"))
    sys.exit(wind.exec_())
