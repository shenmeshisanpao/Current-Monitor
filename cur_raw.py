#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Real-time Dual-Channel Current Monitor with PyQt5 and Auto-Save
# Author: ZhiCheng Zhang <zhangzhicheng@cnncmail.cn>
# Date: 2025-07-27

import fcntl  # Linux/Unix系统的文件锁
import tempfile
import atexit
import sys
import serial
import struct 
import time
import ctypes
import numpy as np
import os
import shutil
import decimal
import re
from decimal import Decimal, getcontext
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, 
                             QPushButton, QHBoxLayout, QGridLayout, QLineEdit, QFileDialog,
                             QMessageBox, QComboBox, QDialog, QTextBrowser, QInputDialog, QScrollArea)
from PyQt5.QtCore import (QTimer, QUrl)
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib import rcParams

# 设置高精度计算
getcontext().prec = 5  # 设置Decimal精度为5位小数

# 设置Matplotlib参数
rcParams['font.size'] = 8
rcParams['axes.grid'] = True
rcParams['grid.linestyle'] = 'dotted'
rcParams['grid.alpha'] = 0.7

# CRC calculation function
def calculate_crc(data):
    crc = 0xFFFF
    for byte in data:
        crc ^= byte
        for _ in range(8):
            if crc & 0x0001:
                crc >>= 1
                crc ^= 0xA001
            else:
                crc >>= 1
    return crc

# Build request function
def build_request(slave_address, function_code, start_address, quantity):
    request = struct.pack('>B B H H', slave_address, function_code, start_address, quantity)
    crc = calculate_crc(request)
    request += struct.pack('<H', crc)  # CRC 的低字节在前，高字节在后
    return request

# Parse response function
def parse_response(response):
    # 解析帧内容
    slave_address = response[0]
    function_code = response[1]
    byte_count = response[2]
    data = response[3:3 + byte_count]

    # 提取寄存器数据
    registers = []
    for i in range(0, len(data), 2):
        register_value = int.from_bytes(data[i:i+2], byteorder='big')
        registers.append(register_value)

    return slave_address, function_code, registers

# 格式转换
def hex2float(h):
    i = int(h, 16)
    cp = ctypes.pointer(ctypes.c_int(i))
    fp = ctypes.cast(cp, ctypes.POINTER(ctypes.c_float))
    return fp.contents.value

class SingleInstanceLock:
    """单实例锁管理器"""
    def __init__(self, lock_file_name="current_monitor.lock"):
        self.lock_file_name = lock_file_name
        self.lock_file_path = os.path.join(tempfile.gettempdir(), lock_file_name)
        self.lock_file = None
        
    def acquire_lock(self):
        """获取锁"""
        try:
            self.lock_file = open(self.lock_file_path, 'w')
            # 尝试获取独占锁
            fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            # 写入进程ID
            self.lock_file.write(str(os.getpid()))
            self.lock_file.flush()
            # 注册退出时释放锁
            atexit.register(self.release_lock)
            return True
        except (IOError, OSError):
            if self.lock_file:
                self.lock_file.close()
                self.lock_file = None
            return False
    
    def release_lock(self):
        """释放锁"""
        if self.lock_file:
            try:
                fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_UN)
                self.lock_file.close()
                os.remove(self.lock_file_path)
            except:
                pass
            finally:
                self.lock_file = None

class RealTimePlotApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Real-Time Current Monitoring System - Dual Channel")

        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
        if os.path.exists(icon_path):
            icon = QtGui.QIcon(icon_path)
            self.setWindowIcon(icon)
            # 同时设置应用程序图标
            QApplication.instance().setWindowIcon(icon)

        self.setGeometry(100, 100, 1200, 800)
        
        # 初始化两个串口
        self.serialport1 = serial.Serial()
        self.serialport1.port = '/dev/ttyUSB0'
        self.serialport1.baudrate = 9600
        self.serialport1.parity = 'N'
        self.serialport1.bytesize = 8
        self.serialport1.stopbits = 1
        self.serialport1.timeout = 0.1
        
        self.serialport2 = serial.Serial()
        self.serialport2.port = '/dev/ttyUSB1'
        self.serialport2.baudrate = 9600
        self.serialport2.parity = 'N'
        self.serialport2.bytesize = 8
        self.serialport2.stopbits = 1
        self.serialport2.timeout = 0.1
        
        # 初始化变量
        self.run_stat = False
        self.column_int1 = Decimal('0.0')  # 通道1电荷量
        self.column_int2 = Decimal('0.0')  # 通道2电荷量
        self.start_time = None
        self.last_time = None
        self.last_current1 = None  # 通道1上次电流值
        self.last_current2 = None  # 通道2上次电流值
        self.data_points = 100  # 显示的数据点数
        self.x_data = np.linspace(0, self.data_points-1, self.data_points)
        self.y_data1 = np.zeros(self.data_points)  # 通道1数据
        self.y_data2 = np.zeros(self.data_points)  # 通道2数据
        self.time_data = np.zeros(self.data_points)
        self.file_handle = None  # 文件句柄
        self.filename = "Run_0000.csv"  # 默认文件名
        self.file_mode = "append"  # 默认文件模式：追加
        self.update_interval = 100  # 默认更新间隔100ms

        # 脉冲提醒相关变量
        self.pulse_reminder_enabled = True  # 脉冲提醒开关，默认开启
        self.pulse_reminder_timer = QTimer()  # 脉冲提醒定时器
        self.pulse_reminder_timer.timeout.connect(self.show_pulse_reminder)
        self.pulse_reminder_timer.setSingleShot(True)  # 单次触发
        self.reminder_suppressed = False  # 本轮是否已抑制提醒

        # 鼠标悬停相关属性
        self.hover_annotation = None       

        # 然后创建UI和菜单栏
        self.init_ui()
        self.create_menu_bar()
        
        # 创建定时器
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_data)
        self.timer.start(self.update_interval)
        
    def create_menu_bar(self):
        """创建菜单栏"""
        menu_bar = self.menuBar()
        
        # 文件菜单
        file_menu = menu_bar.addMenu('File')
        
        # 添加创建快照菜单项
        self.snapshot_action = QtWidgets.QAction('Create Data Snapshot', self)
        self.snapshot_action.triggered.connect(self.create_snapshot)
        self.snapshot_action.setShortcut('Ctrl+S')
        self.snapshot_action.setToolTip("Create an Independent Copy of Current Data with Additional Metadata")
        file_menu.addAction(self.snapshot_action)
        
        # 打开数据文件夹菜单项
        open_folder_action = QtWidgets.QAction('Open Data Folder', self)
        open_folder_action.triggered.connect(self.open_data_folder)
        open_folder_action.setShortcut('Ctrl+O')
        file_menu.addAction(open_folder_action)
        
        # 退出菜单项
        exit_action = QtWidgets.QAction('Exit', self)
        exit_action.triggered.connect(self.close)
        exit_action.setShortcut('Ctrl+Q')
        file_menu.addAction(exit_action)
        
        # 运行菜单
        run_menu = menu_bar.addMenu('Run')
        
        # 开始监控菜单项
        self.start_action = QtWidgets.QAction('Start Monitoring', self)
        self.start_action.triggered.connect(self.start_monitoring)
        self.start_action.setShortcut('Ctrl+R')
        run_menu.addAction(self.start_action)
        
        # 停止监控菜单项
        self.stop_action = QtWidgets.QAction('Stop Monitoring', self)
        self.stop_action.triggered.connect(self.stop_monitoring)
        self.stop_action.setShortcut('Ctrl+T')
        self.stop_action.setEnabled(False)
        run_menu.addAction(self.stop_action)

        run_menu.addSeparator()

        # 脉冲提醒开关菜单项
        self.pulse_reminder_action = QtWidgets.QAction('Pulse Reminder', self)
        self.pulse_reminder_action.setCheckable(True)
        self.pulse_reminder_action.setChecked(True)  # 默认开启
        self.pulse_reminder_action.triggered.connect(self.toggle_pulse_reminder)
        self.pulse_reminder_action.setToolTip("Enable/Disable Pulse Reminder")
        run_menu.addAction(self.pulse_reminder_action)

        run_menu.addSeparator()

        # 更新间隔设置菜单项
        update_interval_action = QtWidgets.QAction('Set Update Interval', self)
        update_interval_action.triggered.connect(self.set_update_interval)
        update_interval_action.setShortcut('Ctrl+I')
        update_interval_action.setToolTip("Set Data Update Interval")
        run_menu.addAction(update_interval_action)

        # 帮助菜单
        help_menu = menu_bar.addMenu('Help')
        
        # 教程菜单项
        tutorial_action = QtWidgets.QAction('Tutorial', self)
        tutorial_action.triggered.connect(self.show_tutorial)
        tutorial_action.setShortcut('Ctrl+H')
        help_menu.addAction(tutorial_action)

        # 关于菜单项
        about_action = QtWidgets.QAction('About', self)
        about_action.triggered.connect(self.show_about)
        about_action.setShortcut('Ctrl+A')
        help_menu.addAction(about_action)

    def init_ui(self):
        # 主布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # 双串口设置区域
        serial_layout = QVBoxLayout()
        
        # 通道1串口设置
        serial1_layout = QHBoxLayout()
        self.port1_label = QLabel("Channel 1 Serial Port:")
        self.port1_input = QLineEdit(self.serialport1.port)
        self.port1_input.setMinimumWidth(150)
        self.port1_input.setToolTip("Enter Channel 1 Serial Port Device Name")
        
        self.test_serial1_button = QPushButton("Test Port 1")
        self.test_serial1_button.clicked.connect(lambda: self.test_serial_connection(1))
        self.test_serial1_button.setStyleSheet("background-color: #9C27B0; color: white; font-weight: bold;")
        
        serial1_layout.addWidget(self.port1_label)
        serial1_layout.addWidget(self.port1_input)
        serial1_layout.addWidget(self.test_serial1_button)
        
        # 通道2串口设置
        serial2_layout = QHBoxLayout()
        self.port2_label = QLabel("Channel 2 Serial Port:")
        self.port2_input = QLineEdit(self.serialport2.port)
        self.port2_input.setMinimumWidth(150)
        self.port2_input.setToolTip("Enter Channel 2 Serial Port Device Name")
        
        self.test_serial2_button = QPushButton("Test Port 2")
        self.test_serial2_button.clicked.connect(lambda: self.test_serial_connection(2))
        self.test_serial2_button.setStyleSheet("background-color: #9C27B0; color: white; font-weight: bold;")
        
        serial2_layout.addWidget(self.port2_label)
        serial2_layout.addWidget(self.port2_input)
        serial2_layout.addWidget(self.test_serial2_button)
        
        serial_layout.addLayout(serial1_layout)
        serial_layout.addLayout(serial2_layout)
        
        # 文件设置区域
        file_layout = QHBoxLayout()
        
        # 文件名输入框
        self.filename_label = QLabel("Save File Name:")
        self.filename_input = QLineEdit(self.filename)
        self.filename_input.setMinimumWidth(300)
        
        # 浏览按钮
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self.browse_file)
        self.browse_button.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold;")
        
        # 文件模式选择
        self.file_mode_label = QLabel("File Mode:")
        self.file_mode_combo = QComboBox()
        self.file_mode_combo.addItem("Append (If File Exists)")
        self.file_mode_combo.addItem("Overwrite (If File Exists)")
        self.file_mode_combo.setCurrentIndex(0)
        self.file_mode_combo.currentIndexChanged.connect(self.file_mode_changed)
        
        file_layout.addWidget(self.filename_label)
        file_layout.addWidget(self.filename_input)
        file_layout.addWidget(self.browse_button)
        file_layout.addWidget(self.file_mode_label)
        file_layout.addWidget(self.file_mode_combo)
        
        # 状态显示区
        status_layout = QGridLayout()
        
        # 创建标签 - 双通道显示
        self.current1_label = QLabel("Channel 1 Current: --- mA")
        self.current2_label = QLabel("Channel 2 Current: --- mA")
        self.runtime_label = QLabel("Run Time: ---")
        self.integral1_label = QLabel("Channel 1 Integral: --- mC")
        self.integral2_label = QLabel("Channel 2 Integral: --- mC")
        self.timestamp_label = QLabel("Last Update Time (Local): ---")
        self.utc_timestamp_label = QLabel("UTC Timestamp: ---")
        self.save_status_label = QLabel("Save Status: Not Saved")
        
        # 设置标签样式
        for label in [self.current1_label, self.current2_label, self.runtime_label, 
                      self.integral1_label, self.integral2_label, self.timestamp_label, 
                      self.utc_timestamp_label, self.save_status_label]:
            label.setStyleSheet("font-size: 14px; font-weight: bold;")
            label.setMinimumHeight(30)
        
        # 设置不同颜色区分通道
        self.current1_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #1f77b4;")
        self.current2_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #ff7f0e;")
        self.integral1_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #1f77b4;")
        self.integral2_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #ff7f0e;")
        
        # 添加到布局
        status_layout.addWidget(QLabel("Dual-Channel Monitoring Status:"), 0, 0, 1, 2)
        status_layout.addWidget(self.current1_label, 1, 0)
        status_layout.addWidget(self.current2_label, 1, 1)
        status_layout.addWidget(self.integral1_label, 2, 0)
        status_layout.addWidget(self.integral2_label, 2, 1)
        status_layout.addWidget(self.runtime_label, 3, 0)
        status_layout.addWidget(self.timestamp_label, 3, 1)
        status_layout.addWidget(self.utc_timestamp_label, 4, 1)
        status_layout.addWidget(self.save_status_label, 4, 0)
        
        # 控制按钮
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("Start Monitoring")
        self.start_button.clicked.connect(self.start_monitoring)
        self.start_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        
        self.stop_button = QPushButton("Stop Monitoring")
        self.stop_button.clicked.connect(self.stop_monitoring)
        self.stop_button.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")
        self.stop_button.setEnabled(False)
        
        # 创建快照按钮
        self.snapshot_button = QPushButton("Create Snapshot")
        self.snapshot_button.clicked.connect(self.create_snapshot)
        self.snapshot_button.setStyleSheet("background-color: #FFC107; color: black; font-weight: bold;")
        self.snapshot_button.setToolTip("Create an Independent Copy of Current Data")
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(self.snapshot_button)
        
        # 绘图区域 - 双线绘图
        self.figure = Figure(figsize=(12, 6), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_xlabel('time (s)')
        self.ax.set_ylabel('current (mA)')
        
        # 创建两条线
        self.line1, = self.ax.plot(self.x_data, self.y_data1, 'b-', label='Channel 1', linewidth=2)
        self.line2, = self.ax.plot(self.x_data, self.y_data2, 'r-', label='Channel 2', linewidth=2)
        
        self.ax.set_ylim(0, 10)
        self.ax.legend()  # 显示图例
        self.ax.grid(True, alpha=0.3)

        # 鼠标悬停功能
        self.setup_mouse_hover()
        
        # 添加到主布局
        main_layout.addLayout(serial_layout)
        main_layout.addLayout(file_layout)
        main_layout.addLayout(status_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.canvas)
    
    def test_serial_connection(self, channel):
        """测试串口连接"""
        if channel == 1:
            port = self.port1_input.text().strip()
            serialport = self.serialport1
            channel_name = "Channel 1"
        else:
            port = self.port2_input.text().strip()
            serialport = self.serialport2
            channel_name = "Channel 2"
            
        if not port:
            QMessageBox.warning(self, "Warning", f"Please Enter {channel_name} Serial Port Name")
            return
            
        try:
            # 尝试打开串口
            serialport.port = port
            serialport.open()
            
            if serialport.is_open:
                serialport.close()
                QMessageBox.information(self, "Test Successful", f"{channel_name} Serial Port {port} Test Successful!")
                self.save_status_label.setText(f"{channel_name} Serial Port Test: Successful")
                self.save_status_label.setStyleSheet("color: green;")
            else:
                QMessageBox.warning(self, "Test Failed", f"Cannot Open {channel_name} Serial Port {port}")
                self.save_status_label.setText(f"{channel_name} Serial Port Test: Failed")
                self.save_status_label.setStyleSheet("color: red;")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"{channel_name} Serial Port Test Failed: {str(e)}")
            self.save_status_label.setText(f"{channel_name} Serial Port Test: Failed - {str(e)}")
            self.save_status_label.setStyleSheet("color: red;")
    
    def file_mode_changed(self, index):
        """文件模式改变时的处理"""
        self.file_mode = "append" if index == 0 else "overwrite"
        print(f"File Mode Changed to: {self.file_mode}")
    
    def browse_file(self):
        """浏览并选择保存文件"""
        filename, _ = QFileDialog.getSaveFileName(
            self, "Where to Save Data File?", self.filename_input.text(), "CSV File (*.csv)"
        )
        if filename:
            if not filename.lower().endswith('.csv'):
                filename += '.csv'
            self.filename_input.setText(filename)
    
    def open_data_file(self):
        """打开数据文件并写入表头"""
        self.filename = self.filename_input.text()
        
        # 确保文件名以.csv结尾
        if not self.filename.lower().endswith('.csv'):
            self.filename += '.csv'
            self.filename_input.setText(self.filename)
        
        try:
            # 检查文件是否存在
            file_exists = os.path.exists(self.filename)
            
            # 如果文件存在且模式为覆盖，提示用户确认
            if file_exists and self.file_mode == "overwrite":
                reply = QMessageBox.question(self, "Confirm Overwrite", 
                                        f"File '{self.filename}' Exists. Are You Sure You Want to Overwrite?",
                                        QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.No:
                    self.save_status_label.setText("Save Status: User Canceled Overwrite Operation")
                    self.save_status_label.setStyleSheet("color: orange;")
                    return False
            
            # 在追加模式下，先检查文件末尾是否需要换行符
            needs_newline = False
            if file_exists and self.file_mode == "append":
                try:
                    with open(self.filename, 'rb') as f:
                        # 移到文件末尾
                        f.seek(0, 2)
                        file_size = f.tell()
                        if file_size > 0:
                            # 读取最后一个字节
                            f.seek(-1, 2)
                            last_byte = f.read(1)
                            # 检查是否为换行符
                            needs_newline = last_byte != b'\n'
                except Exception as e:
                    print(f"Warning: Cannot check file ending: {e}")
                    needs_newline = True
            
            # 打开文件
            mode = "w" if (self.file_mode == "overwrite" or not file_exists) else "a"
            self.file_handle = open(self.filename, mode)
            
            # 如果是新文件或覆盖模式，写入表头
            if mode == "w" or (mode == "a" and not file_exists):
                # 修改表头以包含双通道数据
                self.file_handle.write("UTC Timestamp, Run Time (Seconds), Channel 1 Current (mA), Channel 2 Current (mA), Channel 1 Integral (mC), Channel 2 Integral (mC)\n")
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                self.file_handle.write(f"# New dual-channel monitoring session started at {timestamp}\n")
            else:
                # 追加模式且文件已存在，确保另起一行并添加分隔注释
                if needs_newline:
                    self.file_handle.write('\n')
                
                # 添加新会话开始的标记
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                self.file_handle.write(f"# New dual-channel monitoring session started at {timestamp}\n")
            
            self.save_status_label.setText(f"Save Status: Saving to {self.filename} ({'Overwrite' if mode == 'w' else 'Append'})")
            self.save_status_label.setStyleSheet("color: green;")
            print(f"Data File Opened: {self.filename} (Mode: {mode})")
            return True
        except Exception as e:
            self.save_status_label.setText(f"Save Status: File Open Failed - {str(e)}")
            self.save_status_label.setStyleSheet("color: red;")
            print(f"Failed to Open Data File: {e}")
            return False
    
    def close_data_file(self):
        """关闭数据文件"""
        if self.file_handle:
            try:
                self.file_handle.close()
                self.save_status_label.setText(f"Save Status: Saved to {self.filename}")
                print(f"Data File Closed: {self.filename}")
            except Exception as e:
                self.save_status_label.setText(f"Save Status: File Close Failed - {str(e)}")
                self.save_status_label.setStyleSheet("color: red;")
                print(f"Failed to Close Data File: {e}")
            finally:
                self.file_handle = None
    
    def write_data_row(self, time_val, runtime, current1, current2, integral1, integral2):
        """写入一行数据到文件"""
        if not self.file_handle:
            return
        
        try:
            # 写入双通道数据
            self.file_handle.write(f"{time_val:.1f},{runtime:.3f},{current1:.3f},{current2:.3f},{integral1:.3f},{integral2:.3f}\n")
            self.file_handle.flush()  # 确保数据立即写入
        except Exception as e:
            self.save_status_label.setText(f"Save Status: Write Failed - {str(e)}")
            self.save_status_label.setStyleSheet("color: red;")
            print(f"Failed to Write Data: {e}")
    
    def create_snapshot(self):
        """创建当前数据的独立快照副本"""
        if not self.run_stat:
            QMessageBox.critical(self, "Error", "Monitoring Not Running, No Data to Create Snapshot!")
            return
        
        # 检查文件是否为空
        if os.path.exists(self.filename) and os.path.getsize(self.filename) == 0:
            QMessageBox.critical(self, "Error", "Data File is Empty, Cannot Create Snapshot!")
            return
    
        # 确保文件已刷新
        if self.file_handle:
            try:
                self.file_handle.flush()
                os.fsync(self.file_handle.fileno())
            except Exception as e:
                QMessageBox.warning(self, "Warning", f"Cannot Flush File Buffer: {str(e)}")
        
        # 获取当前时间戳
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        
        # 构造快照文件名
        base_name, ext = os.path.splitext(self.filename)
        if not ext:
            ext = ".csv"
        snapshot_file = f"{base_name}_snapshot_{timestamp}{ext}"
        
        try:
            # 复制主数据文件到快照文件
            shutil.copyfile(self.filename, snapshot_file)
            
            # 在快照文件中添加元数据
            with open(snapshot_file, 'a') as f:
                # 添加空行分隔符
                f.write("\n")
                # 添加元数据
                f.write(f"# Snapshot Creation Time (Local): {time.strftime('%Y-%m-%d %H:%M:%S %Z', time.localtime())}\n")
                f.write(f"# Snapshot Creation Time (UTC): {time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime())}\n")
                f.write(f"# Monitoring Run Time: {self.runtime_label.text().replace('Run Time: ', '')}\n")
                f.write(f"# Channel 1 Integral Value: {self.integral1_label.text().replace('Channel 1 Integral: ', '')}\n")
                f.write(f"# Channel 2 Integral Value: {self.integral2_label.text().replace('Channel 2 Integral: ', '')}\n")
                f.write(f"# Snapshot Source File: {os.path.basename(self.filename)}\n")
            
            # 显示成功消息
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Snapshot Created Successfully")
            msg.setText(f"Data Snapshot File Created:\n{snapshot_file}")
            msg.setDetailedText(f"File Location: {os.path.abspath(snapshot_file)}\n"
                            f"Snapshot Time: {time.ctime()}\n"
                            f"Channel 1 Integral: {self.integral1_label.text()}\n"
                            f"Channel 2 Integral: {self.integral2_label.text()}")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
            
            # 更新状态栏
            self.save_status_label.setText(f"Snapshot Status: Created {os.path.basename(snapshot_file)}")
            self.save_status_label.setStyleSheet("color: blue;")
            
            print(f"Data Snapshot Created: {snapshot_file}")
            return True
            
        except Exception as e:
            error_msg = f"Snapshot Creation Failed: {str(e)}"
            QMessageBox.critical(self, "Error", error_msg)
            self.save_status_label.setText(f"Snapshot Status: {error_msg}")
            self.save_status_label.setStyleSheet("color: red;")
            print(f"Failed to Create Data Snapshot: {e}")
            return False
        
    def extract_number_from_filename(self, filename):
        """从文件名中提取末尾的四位数字"""
        import re
        # 移除.csv扩展名
        base_name = filename.replace('.csv', '')
        # 匹配末尾的四位数字
        match = re.search(r'_(\d{4})$', base_name)
        if match:
            return int(match.group(1)), base_name[:-5]  # 返回数字和前缀部分
        return None, base_name

    def generate_next_filename(self, current_filename):
        """生成下一个文件名"""
        number, prefix = self.extract_number_from_filename(current_filename)
        if number is not None:
            # 如果找到四位数字，递增
            next_number = (number + 1) % 10000  # 确保不超过四位数
            next_filename = f"{prefix}_{next_number:04d}.csv"
            return next_filename
        else:
            # 如果没有找到四位数字格式，保持原文件名不变
            return current_filename

    
    def send_data(self, channel):
        """发送数据请求到指定通道"""
        slave_address = 1
        function_code = 3
        start_address = 42
        quantity = 2
        
        try:
            request = build_request(slave_address, function_code, start_address, quantity)
            if channel == 1:
                self.serialport1.write(request)
            else:
                self.serialport2.write(request)
        except Exception as e:
            print(f"Failed to Send Request to Channel {channel}: {e}")
    
    def recv_data(self, channel):
        """接收并解析指定通道的数据"""
        try:
            # 读取响应数据
            if channel == 1:
                response = self.serialport1.read(9)
            else:
                response = self.serialport2.read(9)
                
            if len(response) < 9:
                raise ValueError(f"Channel {channel} Response Data Incomplete")
            
            # 解析响应
            slave_address, function_code, registers = parse_response(response)
            if len(registers) != 2:
                raise ValueError(f"Channel {channel} Incorrect Number of Registers")
            
            # 将两个寄存器组合成32位整数
            register_value = (registers[0] << 16) | registers[1]
            hex_str = format(register_value, '08X')
            return hex2float(hex_str)
        except Exception as e:
            print(f"Channel {channel} Parsing Error: {e}")
            return None
    
    def get_time(self):
        """获取当前时间戳"""
        return time.time()
    
    def start_monitoring(self):
        """开始监控"""
        # 获取串口名
        port1 = self.port1_input.text().strip()
        port2 = self.port2_input.text().strip()
        
        if not port1 or not port2:
            QMessageBox.warning(self, "Warning", "Please Enter Both Serial Port Names!")
            return
            
        self.serialport1.port = port1
        self.serialport2.port = port2
        
        try:
            # 打开两个串口
            if self.serialport1.is_open:
                self.serialport1.close()
            if self.serialport2.is_open:
                self.serialport2.close()
                
            self.serialport1.open()
            self.serialport2.open()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to Open Serial Ports: {e}")
            return
        
        # 打开数据文件
        if not self.open_data_file():
            QMessageBox.critical(self, "Error", "Cannot Open Data File, Monitoring Cannot Start!")
            return
        
        self.run_stat = True
        # 更新菜单状态
        self.start_action.setEnabled(False)  
        self.stop_action.setEnabled(True)    

        self.start_time = self.get_time()
        self.last_time = self.start_time
        self.column_int1 = Decimal('0.0')  # 重置通道1电荷量
        self.column_int2 = Decimal('0.0')  # 重置通道2电荷量
        self.last_current1 = None  # 重置通道1上次电流值
        self.last_current2 = None  # 重置通道2上次电流值
        self.y_data1 = np.zeros(self.data_points)
        self.y_data2 = np.zeros(self.data_points)
        self.time_data = np.zeros(self.data_points)
        
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)

        # 重置提醒抑制状态
        self.reminder_suppressed = False

        # 如果脉冲提醒功能开启，启动10秒后的提醒定时器
        if self.pulse_reminder_enabled:
            self.pulse_reminder_timer.start(10 * 1000)  # 10秒后提醒
                        
        print("Monitoring Started")
    
    def stop_monitoring(self):
        """停止监控"""
        self.run_stat = False
        # 更新菜单状态
        self.start_action.setEnabled(True)   
        self.stop_action.setEnabled(False)   

        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        
        # 关闭数据文件
        self.close_data_file()

        # 停止脉冲提醒定时器
        if self.pulse_reminder_timer.isActive():
            self.pulse_reminder_timer.stop()
        
        # 关闭串口
        if self.serialport1.is_open:
            try:
                self.serialport1.close()
                print("Serial Port 1 Closed")
            except Exception as e:
                print(f"Failed to Close Serial Port 1: {e}")
                
        if self.serialport2.is_open:
            try:
                self.serialport2.close()
                print("Serial Port 2 Closed")
            except Exception as e:
                print(f"Failed to Close Serial Port 2: {e}")
        
        # 自动更新文件名为下一个序号
        current_filename = self.filename_input.text().strip()
        if not current_filename:
            current_filename = self.filename
        
        next_filename = self.generate_next_filename(current_filename)
        self.filename_input.setText(next_filename)
        self.filename = next_filename
        
        print(f"Monitoring Stopped")

    
    def open_data_folder(self):
        """打开数据文件所在的文件夹"""
        # 获取当前设置的文件路径
        filepath = self.filename_input.text().strip()
        
        # 如果文件路径为空，使用默认文件名
        if not filepath:
            filepath = self.filename
        
        # 获取文件所在的目录
        dir_path = os.path.dirname(os.path.abspath(filepath)) if filepath else os.getcwd()
        
        # 检查目录是否存在
        if not dir_path or not os.path.exists(dir_path):
            QMessageBox.warning(self, "Warning", "Directory Does Not Exist, Please Select a Valid Save Location First")
            return
        
        try:
            # 使用系统默认方式打开文件夹
            if sys.platform == 'win32':
                os.startfile(dir_path)
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{dir_path}"')
            else:  # Linux
                os.system(f'xdg-open "{dir_path}"')
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Cannot Open Folder: {str(e)}")

    def set_update_interval(self):
        """设置更新间隔对话框"""
        # 获取当前更新间隔
        current_interval = self.update_interval
        
        # 显示输入对话框
        interval, ok = QInputDialog.getInt(
            self, 
            'Set Update Interval', 
            'Update Interval (milliseconds):\n\nAvailable range: 10-5000ms\nCurrent value: {}ms'.format(current_interval),
            current_interval,  # 默认值
            10,               # 最小值
            5000,             # 最大值
            10                # 步长
        )
        
        if ok and interval != current_interval:
            # 更新间隔值
            self.update_interval = interval
            
            # 如果定时器正在运行，重新启动定时器
            if self.timer.isActive():
                self.timer.stop()
                self.timer.start(self.update_interval)
                
                # 显示确认消息
                QMessageBox.information(
                    self, 
                    'Update Interval Changed', 
                    f'Update interval has been changed to {interval}ms.\n\n'
                    f'Note: This change takes effect immediately for running monitoring.'
                )
            else:
                # 显示确认消息
                QMessageBox.information(
                    self, 
                    'Update Interval Set', 
                    f'Update interval has been set to {interval}ms.\n\n'
                    f'This will take effect when monitoring starts.'
                )
            
            print(f"Update interval changed to: {interval}ms")

    def toggle_pulse_reminder(self):
        """切换脉冲提醒开关"""
        self.pulse_reminder_enabled = self.pulse_reminder_action.isChecked()
        status = "enabled" if self.pulse_reminder_enabled else "disabled"
        print(f"Pulse reminder {status}")
        
        # 如果关闭提醒且定时器正在运行，停止定时器
        if not self.pulse_reminder_enabled and self.pulse_reminder_timer.isActive():
            self.pulse_reminder_timer.stop()

    def show_pulse_reminder(self):
        """显示脉冲提醒对话框"""
        if not self.pulse_reminder_enabled or self.reminder_suppressed:
            return
        
        # 创建自定义对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("Pulse Reminder")
        dialog.setFixedSize(300, 150)
        dialog.setModal(True)
        
        # 设置对话框图标
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
        if os.path.exists(icon_path):
            dialog.setWindowIcon(QtGui.QIcon(icon_path))
        
        layout = QVBoxLayout()
        
        # 提醒文本
        message_label = QLabel("Reminder: Please apply pulse")
        message_label.setAlignment(QtCore.Qt.AlignCenter)
        message_label.setStyleSheet("font-size: 14px; font-weight: bold; margin: 10px;")
        layout.addWidget(message_label)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        
        # "本轮内不再提醒"按钮
        no_more_button = QPushButton("No More in This Run")
        no_more_button.clicked.connect(lambda: self.handle_reminder_choice(dialog, "no_more"))
        no_more_button.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 8px;")
        
        # "5分钟后再次提醒"按钮
        remind_later_button = QPushButton("Remind in 5 Minutes")
        remind_later_button.clicked.connect(lambda: self.handle_reminder_choice(dialog, "remind_later"))
        remind_later_button.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 8px;")
        
        button_layout.addWidget(no_more_button)
        button_layout.addWidget(remind_later_button)
        
        layout.addLayout(button_layout)
        dialog.setLayout(layout)
        
        # 显示对话框
        dialog.exec_()

    def handle_reminder_choice(self, dialog, choice):
        """处理用户的提醒选择"""
        dialog.accept()  # 关闭对话框
        
        if choice == "no_more":
            # 本轮内不再提醒
            self.reminder_suppressed = True
            print("Pulse reminder suppressed for this monitoring session")
        elif choice == "remind_later":
            # 5分钟后再次提醒
            self.pulse_reminder_timer.start(5 * 60 * 1000)  # 5分钟 = 300000毫秒
            print("Pulse reminder will show again in 5 minutes")

    def setup_mouse_hover(self):
        """设置鼠标悬停功能"""
        # 创建注释文本框
        self.hover_annotation = self.ax.annotate(
            '', 
            xy=(0, 0), 
            xytext=(20, 20), 
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.5", fc="yellow", alpha=0.8),
            arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0"),
            fontsize=10,
            visible=False
        )
        
        # 连接鼠标移动事件
        self.canvas.mpl_connect('motion_notify_event', self.on_hover)

    def on_hover(self, event):
        """鼠标悬停事件处理"""
        if not self.run_stat or event.inaxes != self.ax:
            self.hover_annotation.set_visible(False)
            self.canvas.draw_idle()
            return
        
        # 获取鼠标位置
        x_mouse = event.xdata
        y_mouse = event.ydata
        
        if x_mouse is None or y_mouse is None:
            self.hover_annotation.set_visible(False)
            self.canvas.draw_idle()
            return
        
        # 找到最接近的数据点
        closest_point = self.find_closest_point(x_mouse, y_mouse)
        
        if closest_point:
            channel, index, x_val, y_val, time_val = closest_point
            
            # 计算运行时间
            if self.start_time and time_val > 0:
                runtime = time_val - self.start_time
                time_str = f"{runtime:.2f}s"
            else:
                time_str = "N/A"
            
            # 创建显示文本
            hover_text = f"Channel {channel}\nTime: {time_str}\nCurrent: {y_val:.3f} mA"
            
            # 更新注释
            self.hover_annotation.xy = (x_val, y_val)
            self.hover_annotation.set_text(hover_text)
            self.hover_annotation.set_visible(True)
            
            # 设置不同通道的颜色
            if channel == 1:
                self.hover_annotation.get_bbox_patch().set_facecolor('lightblue')
            else:
                self.hover_annotation.get_bbox_patch().set_facecolor('lightcoral')
        else:
            self.hover_annotation.set_visible(False)
        
        self.canvas.draw_idle()

    def find_closest_point(self, x_mouse, y_mouse):
        """找到最接近鼠标位置的数据点"""
        min_distance = float('inf')
        closest_point = None
        
        # 检查通道1的数据点
        for i in range(len(self.x_data)):
            x_data_point = self.x_data[i]
            y_data_point = self.y_data1[i]
            
            # 计算距离（考虑坐标轴的比例）
            x_range = self.ax.get_xlim()[1] - self.ax.get_xlim()[0]
            y_range = self.ax.get_ylim()[1] - self.ax.get_ylim()[0]
            
            # 归一化距离计算
            dx = (x_mouse - x_data_point) / x_range
            dy = (y_mouse - y_data_point) / y_range
            distance = (dx**2 + dy**2)**0.5
            
            if distance < min_distance and distance < 0.05:  # 设置敏感度阈值
                min_distance = distance
                time_val = self.time_data[i] if hasattr(self, 'time_data') and i < len(self.time_data) else 0
                closest_point = (1, i, x_data_point, y_data_point, time_val)
        
        # 检查通道2的数据点
        for i in range(len(self.x_data)):
            x_data_point = self.x_data[i]
            y_data_point = self.y_data2[i]
            
            # 计算距离
            x_range = self.ax.get_xlim()[1] - self.ax.get_xlim()[0]
            y_range = self.ax.get_ylim()[1] - self.ax.get_ylim()[0]
            
            dx = (x_mouse - x_data_point) / x_range
            dy = (y_mouse - y_data_point) / y_range
            distance = (dx**2 + dy**2)**0.5
            
            if distance < min_distance and distance < 0.05:
                min_distance = distance
                time_val = self.time_data[i] if hasattr(self, 'time_data') and i < len(self.time_data) else 0
                closest_point = (2, i, x_data_point, y_data_point, time_val)
        
        return closest_point

    def show_about(self):
        """显示关于对话框"""
        about_dialog = QDialog(self)
        about_dialog.setWindowTitle("About")
        about_dialog.setFixedSize(520, 450)
        
        # 设置对话框图标
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
        if os.path.exists(icon_path):
            about_dialog.setWindowIcon(QtGui.QIcon(icon_path))
        
        # 设置对话框样式
        about_dialog.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                background-color: white;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 5px;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QScrollArea {
                border: none;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a0a0a0;
            }
        """)
        
        # 创建主布局
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        
        # 创建内容标签
        content_label = QLabel()
        content_label.setWordWrap(True)
        content_label.setAlignment(QtCore.Qt.AlignTop)
        content_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        
        about_text = """
            <div style="font-family: Arial, sans-serif; line-height: 1.6;">
            <center>
            <h2 style="color: #2E86AB; margin-bottom: 10px;">Real-Time Current Monitoring System - Dual Channel</h2>
            <p><b>Version 2025.07.28</b></p>
            <p>Compatible with 創鴻高精度智能盤面表 DM4D-An-Rs</p>
            <p><b>Developer:</b> Zhicheng Zhang</p>
            <p><b>Email:</b> <a href="mailto:zhangzhicheng@cnncmail.cn">zhangzhicheng@cnncmail.cn</a></p>
            <p><b>Special Thanks:</b> Dr. Taoyu Jiao</p>
            <p style="color: #666; font-size: 12px;">© 2025 CIAE & IMP Nuclear Astrophysics Group. All Rights Reserved.</p>
            </center>
            
            <hr style="margin: 20px 0; border: 1px solid #ddd;">
            
            <h3 style="color: #2E86AB;">About This Software</h3>
            <p>This software is designed for real-time monitoring of dual-channel current data and charge integral calculation, supporting automatic data saving and graphical visualization.</p>
            <p>Tested and verified on Ubuntu 24.04 LTS with Anaconda 25.5.1</p>
          
            <hr style="margin: 20px 0; border: 1px solid #ddd;">
            
            <h3 style="color: #2E86AB;">Update History</h3>
            
            <div style="margin-left: 15px;">
            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-07-28 (Latest)</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>An intelligent pulse reminder system was added, and its toggle switch is located in the Run menu.</li>
                <li>Added interactive data point hover functionality to the real-time plot.</li>
            </ul>

            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-07-27</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>Enhanced file naming system with automatic numbering.</li>
                <li>Auto-increment filename when stopping monitoring.</li>
            </ul>

            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-07-16</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>Added dual-channel monitoring support.</li>
                <li>Enhanced plotting with two current curves.</li>
                <li>Updated data saving format for dual channels.</li>
                <li>Improved UI layout for better dual-channel display.</li>
            </ul>
            
            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-07-03</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>Fixed the issue where the storage directory could not be opened when using the default file.</li>
                <li>Fixed the issue where appending to storage did not start a new line.</li>
                <li>Set the maximum value of the current plotting area to 30mA.</li>
                <li>Optimized some details when saving files.</li>
            </ul>

            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-06-30</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>Change the software interface language to English.</li>
            </ul>
            
            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-06-28</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>Added a user guide.</li>
            </ul>
            
            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-06-25</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>Added file locking functionality.</li>
                <li>Added the ability to customize the communication interval for the ammeter.</li>
                <li>Rearranged the interface layout.</li>
            </ul>
            
            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-06-20</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>Added serial port detection functionality.</li>
                <li>Added the ability to save snapshots.</li>
            </ul>
            
            <h4 style="color: #4A90E2; margin-bottom: 5px;"> 2025-06-16</h4>
            <ul style="margin-top: 5px; margin-bottom: 15px;">
                <li>Initial Release.</li>
            </ul>
            </div>
            
            <hr style="margin: 20px 0; border: 1px solid #ddd;">
            
            <center>
            <p style="font-size: 12px; color: #888;">
            For technical support or bug reports, please contact the developer.<br>
            This software is provided "as is" without warranty of any kind.
            </p>
            </center>
            </div>
        """
        
        content_label.setText(about_text)
        
        # 将内容标签添加到滚动区域
        scroll_area.setWidget(content_label)
        
        # 将滚动区域添加到主布局
        layout.addWidget(scroll_area)
        
        # 添加按钮区域
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 5, 0, 0)
        
        ok_button = QPushButton("OK")
        ok_button.setFixedSize(80, 32)
        ok_button.clicked.connect(about_dialog.accept)
        ok_button.setDefault(True)
        
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        # 设置对话框布局
        about_dialog.setLayout(layout)
        
        # 显示对话框
        about_dialog.exec_()

    def show_tutorial(self):
        """显示教程对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Tutorial")
        dialog.setFixedSize(600, 500)
        
        layout = QVBoxLayout()
        text_browser = QTextBrowser()
        text_browser.setOpenExternalLinks(True)
        
        # 获取 HTML 文件的绝对路径
        current_dir = os.path.dirname(os.path.abspath(__file__))
        html_file = os.path.join(current_dir, 'tutorial.html')
        
        if os.path.exists(html_file):
            # 使用 file:// URL 协议
            file_url = QUrl.fromLocalFile(html_file)
            text_browser.setSource(file_url)
        else:
            text_browser.setHtml("<h1>Error</h1><p>tutorial.html File Not Found</p>")
        
        close_button = QPushButton("Close")
        close_button.clicked.connect(dialog.accept)
        
        layout.addWidget(text_browser)
        layout.addWidget(close_button)
        dialog.setLayout(layout)
        
        dialog.exec_()
            
    def closeEvent(self, event):
        """关闭事件，确保资源被正确释放"""
        self.stop_monitoring()
        event.accept()
    
    def update_data(self):
        """定时更新数据 - 双通道版本"""
        if not self.run_stat:
            return
        
        try:
            # 发送请求到两个通道
            self.send_data(1)
            self.send_data(2)
            
            # 接收两个通道的数据
            current1 = self.recv_data(1)
            current2 = self.recv_data(2)
            
            # 检查数据有效性
            if current1 is None or current2 is None:
                return
            
            # 检查电流值是否有效
            # if current1 < 0 or current1 > 1000 or current2 < 0 or current2 > 1000:
            if current1 > 1000 or current2 > 1000:
                print(f"Invalid Current Values: Ch1={current1} mA, Ch2={current2} mA, Skipping Calculation")
                return
            
            # 获取当前时间
            now = self.get_time()
            
            # 获取本地时间和UTC时间
            local_time = time.localtime(now)
            utc_time = time.gmtime(now)
            
            # 格式化时间字符串
            local_time_str = time.strftime("%Y-%m-%d %H:%M:%S (UTC%z)", local_time)
            utc_time_str = time.strftime("%Y-%m-%dT%H:%M:%SZ", utc_time)
            
            # 优化积分计算：使用梯形法则 - 双通道
            if self.last_time is not None:
                delta_t = now - self.last_time
                
                # 跳过无效的时间间隔
                if delta_t <= 0:
                    print(f"Invalid Time Interval: {delta_t}, Skipping Calculation")
                    return
                
                # 通道1积分计算
                if self.last_current1 is not None:
                    avg_current1 = Decimal(str((self.last_current1 + current1) / 2.0))
                    charge_segment1 = avg_current1 * Decimal(str(delta_t))
                    self.column_int1 += charge_segment1
                else:
                    charge_segment1 = Decimal(str(current1)) * Decimal(str(delta_t))
                    self.column_int1 += charge_segment1
                
                # 通道2积分计算
                if self.last_current2 is not None:
                    avg_current2 = Decimal(str((self.last_current2 + current2) / 2.0))
                    charge_segment2 = avg_current2 * Decimal(str(delta_t))
                    self.column_int2 += charge_segment2
                else:
                    charge_segment2 = Decimal(str(current2)) * Decimal(str(delta_t))
                    self.column_int2 += charge_segment2
            
            # 保存当前电流值用于下次计算
            self.last_current1 = current1
            self.last_current2 = current2
            self.last_time = now
            
            # 更新显示标签
            runtime = now - self.start_time
            h = int(runtime // 3600)
            m = int((runtime - h * 3600) // 60)
            s = runtime - h * 3600 - m * 60
            
            # 获取积分值的浮点数表示（用于显示）
            integral1_float = float(self.column_int1)
            integral2_float = float(self.column_int2)
            
            self.current1_label.setText(f"Channel 1 Current: {current1:.3f} mA")
            self.current2_label.setText(f"Channel 2 Current: {current2:.3f} mA")
            self.runtime_label.setText(f"Run Time: {h:02d} Hours {m:02d} Minutes {s:05.2f} Seconds")
            self.integral1_label.setText(f"Channel 1 Integral: {integral1_float:.3f} mC")
            self.integral2_label.setText(f"Channel 2 Integral: {integral2_float:.3f} mC")
            self.timestamp_label.setText(f"Last Update Time (Local): {local_time_str}")
            self.utc_timestamp_label.setText(f"UTC Timestamp: {utc_time_str}")
            
            # 更新绘图数据 - 双通道
            self.y_data1 = np.roll(self.y_data1, -1)
            self.y_data1[-1] = current1
            
            self.y_data2 = np.roll(self.y_data2, -1)
            self.y_data2[-1] = current2
            
            self.time_data = np.roll(self.time_data, -1)
            self.time_data[-1] = now
            
            # 更新曲线
            self.line1.set_ydata(self.y_data1)
            self.line2.set_ydata(self.y_data2)
            
            # 更新X轴标签为运行时间
            x_labels = []
            for i in range(len(self.x_data)):
                time_sec = (self.time_data[i] - self.start_time) if self.time_data[i] > 0 else 0
                x_labels.append(f"{time_sec:.1f}")
            
            # 设置X轴刻度为运行时间
            tick_indices = range(0, self.data_points, max(1, self.data_points//10))
            self.ax.set_xticks(tick_indices)
            self.ax.set_xticklabels([x_labels[i] for i in tick_indices])
            
            # 自动调整Y轴范围 - 考虑两个通道的数据
            all_data = np.concatenate([self.y_data1, self.y_data2])
            min_val = max(-1, min(all_data) - 0.1)
            max_val = min(10, max(all_data) + 0.1)  # 上限设为10mA
            self.ax.set_ylim(min_val, max_val)
            
            # 重绘图形
            self.canvas.draw()
            
            # 写入数据到文件 - 双通道
            self.write_data_row(now, runtime, current1, current2, integral1_float, integral2_float)
            
            # 打印到控制台 - 双通道
            print(f"Ch1: {current1:.3f} mA, Ch2: {current2:.3f} mA, Runtime: {runtime:.3f} s, "
                  f"Int1: {integral1_float:.3f} mC, Int2: {integral2_float:.3f} mC, UTC: {utc_time_str}")
            
        except Exception as e:
            print(f"Error Updating Data: {e}")

def main():
    app = QApplication(sys.argv)
    # 设置应用程序图标
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
    if os.path.exists(icon_path):
        icon = QtGui.QIcon(icon_path)
        app.setWindowIcon(icon)

    try:
        window = RealTimePlotApp()
        window.show()
        sys.exit(app.exec_())
    except SystemExit as e:
        # 如果是因为单实例检查退出，直接退出
        if e.code == 1:
            app.quit()
            sys.exit(1)
        raise

if __name__ == "__main__":
    main()

