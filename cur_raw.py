#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Real-time Current Monitor
# Author: ZhiCheng Zhang <zhangzhicheng@cnncmail.cn>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import sys
import os
import tempfile
import atexit
import serial
import struct
import time
import ctypes
import numpy as np
import shutil
import re
import socket
import subprocess
from datetime import datetime
from decimal import Decimal, getcontext
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, 
                             QPushButton, QHBoxLayout, QGridLayout, QLineEdit, QFileDialog,
                             QMessageBox, QComboBox, QDialog, QTextBrowser, QInputDialog, QScrollArea,
                             QCompleter, QDirModel)
from PyQt5.QtCore import (QTimer, QUrl)
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib import rcParams
from collections import deque

# 文件锁
if os.name == 'nt':  # Windows
    import msvcrt
else:  # Linux/Unix
    import fcntl

# 设置高精度计算
getcontext().prec = 15  # 设置Decimal精度为15位小数

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

# 获取资源的绝对路径，用于 PyInstaller 打包
def resource_path(relative_path):
    try:
        # PyInstaller 创建临时文件夹，将路径存储在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class SingleInstanceLock:
    """跨平台单实例锁管理器"""
    def __init__(self, lock_file_name="current_monitor.lock"):
        self.lock_file_name = lock_file_name
        self.lock_file_path = os.path.join(tempfile.gettempdir(), lock_file_name)
        self.lock_file = None
        
    def acquire_lock(self):
        """获取锁"""
        try:
            self.lock_file = open(self.lock_file_path, 'w')
            
            if os.name == 'nt': # Windows 系统
                # 锁定文件的前10个字节，LK_NBLCK 表示非阻塞锁
                msvcrt.locking(self.lock_file.fileno(), msvcrt.LK_NBLCK, 10)
            else: # Linux/Unix 系统
                fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                
            # 写入进程ID
            self.lock_file.write(str(os.getpid()))
            self.lock_file.flush()
            # 注册退出时释放锁
            atexit.register(self.release_lock)
            return True
        except (IOError, OSError):
            if self.lock_file:
                try:
                    self.lock_file.close()
                except:
                    pass
                self.lock_file = None
            return False
    
    def release_lock(self):
        """释放锁"""
        if self.lock_file:
            try:
                if os.name == 'nt': # Windows 解锁
                    self.lock_file.seek(0)
                    msvcrt.locking(self.lock_file.fileno(), msvcrt.LK_UNLCK, 10)
                else: # Linux 解锁
                    fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_UN)
                
                self.lock_file.close()
                try:
                    os.remove(self.lock_file_path)
                except:
                    pass
            except:
                pass
            finally:
                self.lock_file = None

class LogBus(QtCore.QObject):   # 类: 日志总线 (单例, 发布-订阅)
    """日志总线：单例，收集最近 100 条日志并通过信号广播更新。
    Log bus singleton: collects up to 100 most-recent log entries and
    broadcasts updates via a Qt signal (publish-subscribe)."""
    log_updated = QtCore.pyqtSignal()
    _instance = None

    # 颜色映射: INFO=蓝灰, WARNING=橙, ERROR=红
    # Color mapping: INFO=blue-gray, WARNING=orange, ERROR=red
    LEVEL_COLORS = {
        "INFO":    "#555555",
        "WARNING": "#FF8C00",
        "ERROR":   "#D32F2F",
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        # deque(maxlen=100) 满员后追加自动丢弃最左侧 (FIFO 删首)
        # deque(maxlen=100) auto-drops oldest entry when full (FIFO)
        self._entries = deque(maxlen=100)  # (timestamp_str, level, message)

    @classmethod
    def get_instance(cls):
        """获取单例 (Get the singleton instance)"""
        if cls._instance is None:
            cls._instance = LogBus()
        return cls._instance

    def log(self, level, message):
        """记录一条日志 (Record one log entry)
        level: 'INFO' / 'WARNING' / 'ERROR'
        message: 日志正文"""
        ts = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        self._entries.append((ts, level, message))
        self.log_updated.emit()

    def get_latest_text(self):
        """返回最近一条的格式化文本 (供底部日志栏显示)
        Return formatted text of the latest entry (for the bottom log bar)."""
        if not self._entries:
            return "Event Log: (no events yet)"
        ts, level, msg = self._entries[-1]
        return f"[{ts}] {msg}"

    def get_latest_level(self):
        """返回最近一条的级别 (供着色)
        Return level of the latest entry (for coloring)."""
        if not self._entries:
            return "INFO"
        return self._entries[-1][1]

    def get_all_html(self):
        """生成带颜色 span 的 HTML, 最新在底部 (供日志窗口显示)
        Generate colored HTML, newest at bottom (for the log window)."""
        import html as html_mod  # 局部导入, 避免与变量名冲突 (local import)
        lines = []
        for ts, level, msg in self._entries:
            color = self.LEVEL_COLORS.get(level, "#555555")
            safe_msg = html_mod.escape(msg)
            lines.append(
                f'<div style="color:{color};white-space:pre-wrap;">'
                f'<span style="color:#888;">[{ts}]</span> '
                f'<b>[{level}]</b> {safe_msg}'
                f'</div>'
            )
        return "".join(lines)


class LogWindow(QDialog):   # 类: 非模态日志窗口
    """非模态日志窗口：只读 QTextBrowser, 按颜色区分级别, 最新在底部。
    Non-modal log window: read-only text browser, color-coded, newest at bottom.
    复用策略: closeEvent 仅 hide 不销毁, 由主窗口持有引用。
    Reuse strategy: closeEvent hides instead of destroying; main window keeps the ref."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Event Log")
        self.resize(700, 400)
        self.log_bus = LogBus.get_instance()

        layout = QVBoxLayout(self)
        self.text_edit = QTextBrowser()
        self.text_edit.setReadOnly(True)
        self.text_edit.setOpenExternalLinks(False)
        layout.addWidget(self.text_edit)

        # 订阅更新信号 (Subscribe to update signal)
        self.log_bus.log_updated.connect(self.refresh)

    def refresh(self):
        """刷新日志内容并滚动到底部 (Refresh and scroll to bottom)"""
        self.text_edit.setHtml(self.log_bus.get_all_html())
        sb = self.text_edit.verticalScrollBar()
        sb.setValue(sb.maximum())

    def showEvent(self, event):
        """打开时全量刷新 (Full refresh on show)"""
        self.refresh()
        super().showEvent(event)

    def closeEvent(self, event):
        """隐藏而非销毁, 支持复用 (Hide instead of destroy, for reuse)"""
        event.ignore()
        self.hide()

class ClickableLabel(QLabel):   # 类: 可点击的标签 (带闪烁功能)
    clicked = QtCore.pyqtSignal()

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setAlignment(QtCore.Qt.AlignCenter)
        
        # 基础样式
        self.base_style = "border: 1px solid #555; border-radius: 4px; font-weight: bold; qproperty-alignment: AlignCenter;"
        # 默认（灭灯/背景）样式
        self.default_css = "background-color: #E0E0E0; color: #555;" 
        
        self.setStyleSheet(self.base_style + self.default_css)
        
        # 闪烁定时器
        self.blink_timer = QTimer(self)
        self.blink_timer.timeout.connect(self.toggle_color)
        self.blink_timer.setInterval(200) # 闪烁间隔 200ms
        
        self.blink_state = False # True=亮色, False=暗色
        self.current_css = ""    # 当前应该显示的警告颜色
        self.last_status = ""    # 记录上一次的状态，防止重复重置定时器

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)

    def set_status(self, status, text=None):
        """根据状态改变颜色和文字，处理闪烁逻辑"""
        if text is None:
            text = status
        self.setText(text)
        
        # 如果状态没有变，且处于闪烁模式中，就不要打断定时器，直接返回
        if status == self.last_status:
            return
        
        self.last_status = status
        
        # 颜色定义
        style_map = {
            "STOP": "background-color: #9E9E9E; color: #FFFFFF;", # 灰色
            "RUN":  "background-color: #4CAF50; color: #FFFFFF;", # 绿色
            "PEAK": "background-color: #F44336; color: #FFFFFF;", # 红色
            "DROP": "background-color: #FF9800; color: #FFFFFF;", # 橙色
            "ZERO": "background-color: #D32F2F; color: #FFFFFF; border: 2px solid red;", # 深红
            "INIT": "background-color: #B0BEC5; color: #FFFFFF;"  # 浅灰
        }
        
        # 定义哪些状态需要闪烁
        blinking_states = ["PEAK", "DROP", "ZERO"]
        
        target_css = style_map.get(status, style_map["STOP"])
        
        if status in blinking_states:
            # 启动闪烁
            self.current_css = target_css
            if not self.blink_timer.isActive():
                self.blink_state = True
                self.setStyleSheet(self.base_style + self.current_css) # 立即亮起
                self.blink_timer.start()
        else:
            # 停止闪烁，显示常亮颜色
            self.blink_timer.stop()
            self.setStyleSheet(self.base_style + target_css)

    def toggle_color(self):
        """定时器回调：切换颜色"""
        self.blink_state = not self.blink_state
        if self.blink_state:
            # 亮状态：显示警告色
            self.setStyleSheet(self.base_style + self.current_css)
        else:
            # 灭状态：显示灰色背景 (模拟灯灭)
            self.setStyleSheet(self.base_style + self.default_css)

class StatusMonitor:    # 类: 监控逻辑核心
    def __init__(self):
        self.enabled = True
        self.history = deque()
        self.window_seconds = 10.0  # 窗口时间长度
        self.sample_interval = 0.1  # 采样间隔(s)，会在运行时更新
        
        # 阈值设置
        self.zero_threshold = 0.0001 # 0.1 uA (假设基础单位是mA，这里需要根据实际调整)
        
        self.spike_mode = "percent" # "value" or "percent"
        self.spike_threshold = 0.5 # 0.5 mA
        self.spike_percent = 50.0  # 20%
        
        self.hold_time = 10.0      # 警告保持时间
        
        # 脉冲模式设置
        self.pulse_mode = False     # 脉冲模式开关，默认关闭(直流模式)
        self.zero_timeout = 3.0     # 脉冲模式下零值持续超时时间(秒)
        self.last_nonzero_time = 0  # 最后一次检测到非零电流的时间戳
        
        # 内部状态
        self.warning_state = "RUN" # RUN, PEAK, DROP, ZERO
        self.warning_end_time = 0
        self.is_running = False

    def reset(self):
        self.history.clear()
        self.warning_state = "RUN"
        self.warning_end_time = 0
        self.is_running = False
        self.last_nonzero_time = 0

    def start(self):
        self.is_running = True
        self.history.clear()
        self.warning_state = "RUN"
        self.last_nonzero_time = time.time()  # 初始化为当前时间，避免启动瞬间误报

    def stop(self):
        self.is_running = False
        self.warning_state = "STOP"

    def clear_warning(self):
        """手动清除警告"""
        self.warning_state = "RUN"
        self.warning_end_time = 0

    def update_params(self, interval_ms):
        """更新采样率相关的参数"""
        self.sample_interval = interval_ms / 1000.0
        # 重新计算队列最大长度
        new_maxlen = int(self.window_seconds / self.sample_interval)
        if new_maxlen < 1: new_maxlen = 1
        
        # 如果长度变化，调整deque
        if self.history.maxlen != new_maxlen:
            # 创建一个新的deque，保留旧数据
            new_deque = deque(self.history, maxlen=new_maxlen)
            self.history = new_deque

    def process(self, current_val):
        if not self.enabled:
            return "OFF"
        if not self.is_running:
            return "STOP"

        now = time.time()
        abs_val = abs(current_val)

        # 脉冲模式：仅检测信号丢失，跳过突变检测
        # 电流持续为零超过 zero_timeout 秒才报警 ZERO，避免脉冲 low 期误触发
        if self.pulse_mode:
            if abs_val >= self.zero_threshold:
                # 检测到电流，更新时间戳
                self.last_nonzero_time = now
                if self.warning_state == "ZERO":
                    self.warning_state = "RUN"
                return "RUN"
            else:
                # 电流为零，检查是否超时
                if now - self.last_nonzero_time > self.zero_timeout:
                    self.warning_state = "ZERO"
                    return "ZERO"
                else:
                    # 在 low 期但未超时，正常
                    return "RUN"

        # 以下为直流模式逻辑

        # 1. 检查警告是否过期 (如果当前是警告状态)
        if self.warning_state in ["PEAK", "DROP"]:
            if now > self.warning_end_time:
                self.warning_state = "RUN"

        # 2. 零值检测 (优先级最高)
        if abs_val < self.zero_threshold:
            self.warning_state = "ZERO"
            # ZERO 状态：电流恢复后自动变 RUN (由下面的恢复逻辑处理)
            return "ZERO"

        # 2.5 ZERO 自动恢复 (直流模式)
        # 电流已恢复到 >= zero_threshold，若之前是 ZERO 则改回 RUN
        if self.warning_state == "ZERO":
            self.warning_state = "RUN"

        # 3. 填充历史数据
        # 如果历史数据太少，处于初始化阶段，不报警
        if len(self.history) < 2:
            self.history.append(abs_val)
            return self.warning_state if self.warning_state != "RUN" else "INIT"

        # 计算平均值 (基准线)
        avg_val = sum(self.history) / len(self.history)
        diff = abs_val - avg_val
        
        triggered = False
        trigger_type = "RUN"

        # 4. 突变检测
        if self.spike_mode == "value":
            if diff > self.spike_threshold:
                trigger_type = "PEAK"
                triggered = True
            elif diff < -self.spike_threshold:
                trigger_type = "DROP"
                triggered = True
        else: # percent mode
            if avg_val > 1e-9: # 防止分母为0
                pct = (diff / avg_val) * 100.0
                if pct > self.spike_percent:
                    trigger_type = "PEAK"
                    triggered = True
                elif pct < -self.spike_percent:
                    trigger_type = "DROP"
                    triggered = True

        # 5. 状态更新逻辑
        if triggered:
            # 只有当 新警告优先级更高 或 当前没有警告 时才覆盖
            # 这里简单处理：只要触发就刷新警告状态和计时器
            self.warning_state = trigger_type
            self.warning_end_time = now + self.hold_time
        
        # 6. 更新历史 (放在判断之后，这样突变点本身在下一帧才会进入平均值，
        # 从而保证当前帧能检测到差异。如果先append再算平均，突变会被平均值稀释)
        self.history.append(abs_val)

        return self.warning_state

class MonitorSettingsDialog(QDialog):       #类: 设置对话框
    def __init__(self, monitor1, monitor2, unit1, unit2, alarm_enabled=True, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Status Monitor Settings")
        self.monitor1 = monitor1
        self.monitor2 = monitor2
        self.unit1 = unit1
        self.unit2 = unit2
        self.alarm_enabled = alarm_enabled
        self.resize(400, 500)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        # 报警音开关 (全局)
        self.alarm_cb = QtWidgets.QCheckBox("Alarm Sound (beep on ZERO)")
        self.alarm_cb.setChecked(self.alarm_enabled)
        self.alarm_cb.setToolTip("Play a system beep every 3 seconds when any channel is in ZERO state.")
        layout.addWidget(self.alarm_cb)

        # 创建两个标签页分别设置两个通道
        tab_widget = QtWidgets.QTabWidget()
        tab_widget.addTab(self.create_channel_tab(self.monitor1, self.unit1), "Channel 1")
        tab_widget.addTab(self.create_channel_tab(self.monitor2, self.unit2), "Channel 2")
        layout.addWidget(tab_widget)

        # 底部按钮
        btn_box = QHBoxLayout()
        ok_btn = QPushButton("Apply && Close")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(ok_btn)
        btn_box.addWidget(cancel_btn)
        layout.addLayout(btn_box)

        self.setLayout(layout)

    def create_channel_tab(self, monitor, unit):
        widget = QWidget()
        form = QtWidgets.QFormLayout()

        # 1. 开关
        enable_cb = QtWidgets.QCheckBox("Enable Status Monitor")
        enable_cb.setChecked(monitor.enabled)
        form.addRow(enable_cb)

        # 2. 历史窗口时间
        window_spin = QtWidgets.QDoubleSpinBox()
        window_spin.setRange(1.0, 3600.0)
        window_spin.setValue(monitor.window_seconds)
        window_spin.setSuffix(" s")
        form.addRow("History Window:", window_spin)

        # 3. 零值阈值
        zero_spin = QtWidgets.QDoubleSpinBox()
        zero_spin.setRange(0.0, 99999.0)
        zero_spin.setDecimals(4)
        zero_spin.setValue(monitor.zero_threshold)
        zero_spin.setSuffix(f" {unit}")
        form.addRow("Zero Threshold:", zero_spin)

        # 脉冲模式设置
        form.addRow(QtWidgets.QLabel("--- Pulse Mode ---"))

        pulse_cb = QtWidgets.QCheckBox("Pulse Mode (skip spike detection)")
        pulse_cb.setChecked(monitor.pulse_mode)
        pulse_cb.setToolTip("Enable for pulsed signals. Only detects signal loss (zero for extended period).")
        form.addRow(pulse_cb)

        zero_timeout_spin = QtWidgets.QDoubleSpinBox()
        zero_timeout_spin.setRange(0.1, 60.0)
        zero_timeout_spin.setDecimals(1)
        zero_timeout_spin.setValue(monitor.zero_timeout)
        zero_timeout_spin.setSuffix(" s")
        zero_timeout_spin.setToolTip("In pulse mode, alarm ZERO if current stays at zero for this duration.")
        form.addRow("Zero Timeout:", zero_timeout_spin)

        form.addRow(QtWidgets.QLabel("--- Warning Logic ---"))

        # 4. 突变阈值模式
        mode_group = QtWidgets.QButtonGroup(widget) # 需要parent防止垃圾回收
        radio_val = QtWidgets.QRadioButton(f"By Value ({unit})")
        radio_pct = QtWidgets.QRadioButton("By Percentage (%)")
        mode_group.addButton(radio_val)
        mode_group.addButton(radio_pct)
        
        if monitor.spike_mode == "value":
            radio_val.setChecked(True)
        else:
            radio_pct.setChecked(True)
            
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(radio_val)
        mode_layout.addWidget(radio_pct)
        form.addRow("Threshold Mode:", mode_layout)

        # 5. 阈值数值
        val_spin = QtWidgets.QDoubleSpinBox()
        val_spin.setRange(0.0, 99999.0)
        val_spin.setDecimals(4)
        val_spin.setValue(monitor.spike_threshold)
        val_spin.setSuffix(f" {unit}")
        form.addRow("Value Threshold:", val_spin)

        pct_spin = QtWidgets.QDoubleSpinBox()
        pct_spin.setRange(0.1, 1000.0)
        pct_spin.setValue(monitor.spike_percent)
        pct_spin.setSuffix(" %")
        form.addRow("Percent Threshold:", pct_spin)

        # 6. 警告保持时间
        hold_spin = QtWidgets.QDoubleSpinBox()
        hold_spin.setRange(1.0, 300.0)
        hold_spin.setValue(monitor.hold_time)
        hold_spin.setSuffix(" s")
        form.addRow("Warning Hold Time:", hold_spin)

        # 保存引用以便 accept 时读取
        widget.inputs = {
            "enable": enable_cb,
            "window": window_spin,
            "zero": zero_spin,
            "pulse_mode": pulse_cb,
            "zero_timeout": zero_timeout_spin,
            "mode_val": radio_val,
            "thresh_val": val_spin,
            "thresh_pct": pct_spin,
            "hold": hold_spin
        }
        
        widget.setLayout(form)
        return widget

    def accept(self):
        # 应用设置到 monitor 对象
        # layout: [0]=alarm_cb, [1]=tab_widget, [2]=btn_box
        tab_widget = self.layout().itemAt(1).widget()
        for i in range(2):
            tab = tab_widget.widget(i)
            monitor = self.monitor1 if i == 0 else self.monitor2
            inputs = tab.inputs
            
            monitor.enabled = inputs["enable"].isChecked()
            monitor.window_seconds = inputs["window"].value()
            monitor.zero_threshold = inputs["zero"].value()
            monitor.pulse_mode = inputs["pulse_mode"].isChecked()
            monitor.zero_timeout = inputs["zero_timeout"].value()
            monitor.spike_mode = "value" if inputs["mode_val"].isChecked() else "percent"
            monitor.spike_threshold = inputs["thresh_val"].value()
            monitor.spike_percent = inputs["thresh_pct"].value()
            monitor.hold_time = inputs["hold"].value()
            
        super().accept()

class GDDAQSettingsDialog(QDialog):      # 类: GDDAQ 设置对话框
    def __init__(self, search_dir, proc_name, target_run, parent=None):
        super().__init__(parent)
        self.setWindowTitle("GDDAQ Settings")
        self.resize(450, 180)
        self.init_ui(search_dir, proc_name, target_run)

    def init_ui(self, search_dir, proc_name, target_run):
        layout = QtWidgets.QFormLayout()

        # 1. 数据根目录 (带 Browse 按钮)
        self.search_dir_input = QLineEdit(search_dir)
        # 路径自动补全 (Path auto-completion via QCompleter + QDirModel)
        _dir_completer = QCompleter(self)
        _dir_model = QDirModel()
        _dir_model.setFilter(QtCore.QDir.Dirs | QtCore.QDir.NoDotAndDotDot)
        _dir_completer.setModel(_dir_model)
        _dir_completer.setCompletionMode(QCompleter.PopupCompletion)
        _dir_completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.search_dir_input.setCompleter(_dir_completer)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_dir)
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(self.search_dir_input)
        dir_layout.addWidget(browse_btn)
        dir_widget = QWidget()
        dir_widget.setLayout(dir_layout)
        layout.addRow("Data Directory:", dir_widget)

        # 2. 进程名 (必填)
        self.proc_name_input = QLineEdit(proc_name)
        self.proc_name_input.setPlaceholderText("e.g. gddaq")
        self.proc_name_input.setToolTip("DAQ process name for pgrep -x detection (e.g. gddaq)")
        layout.addRow("Process Name:", self.proc_name_input)

        # 3. 监控轮次 (可留空 = 自动监控最大编号的轮次)
        self.target_run_input = QLineEdit(target_run)
        self.target_run_input.setPlaceholderText("Leave empty = auto (highest run number)")
        self.target_run_input.setToolTip("Leave empty to auto-monitor the run with the highest number; fill a number to monitor a specific run")
        layout.addRow("Run Number:", self.target_run_input)

        # 说明标签
        note = QLabel("Note: Run number corresponds to numbered subdirectories under the data directory.\n"
                      "Leave empty to auto-monitor the run with the highest number.")
        note.setStyleSheet("color: gray; font-size: 10px;")
        note.setWordWrap(True)
        layout.addRow(note)

        # 校验状态提示标签 (红字显示错误原因)
        # Validation status label (shows error reason in red)
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: red; font-size: 10px;")
        self.status_label.setWordWrap(True)
        layout.addRow(self.status_label)

        # 按钮
        btn_box = QHBoxLayout()
        self.ok_btn = QPushButton("Apply && Close")
        self.ok_btn.clicked.connect(self.on_apply)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(self.ok_btn)
        btn_box.addWidget(cancel_btn)
        layout.addRow(btn_box)

        self.setLayout(layout)

        # 必填项校验: Data Directory / Process Name 为空时禁用 Apply 按钮 (只能 Cancel)
        # Required-field validation: disable Apply when Data Directory / Process Name is empty
        self.search_dir_input.textChanged.connect(self.validate_inputs)
        self.proc_name_input.textChanged.connect(self.validate_inputs)
        self.validate_inputs()

    def validate_inputs(self):
        """校验必填项与目录有效性，控制 Apply 按钮可用状态
        (Validate required fields and directory existence, toggle Apply button)"""
        search_dir = self.search_dir_input.text().strip()
        proc_name = self.proc_name_input.text().strip()

        if not search_dir or not proc_name:
            reason = "Data Directory and Process Name are required."
        elif not os.path.isdir(search_dir):
            reason = "Data Directory does not exist."
        else:
            reason = ""

        self.ok_btn.setEnabled(not reason)
        self.status_label.setText(reason)
        self.ok_btn.setToolTip(reason)

    def on_apply(self):
        """Apply 前最终校验，防止目录在输入后被外部删除
        (Final validation before applying, in case the directory was removed)"""
        if not self.ok_btn.isEnabled():
            self.validate_inputs()
            return
        self.accept()

    def browse_dir(self):
        """浏览选择数据根目录"""
        current = self.search_dir_input.text().strip()
        start_dir = current if current else os.path.expanduser("~")
        chosen = QFileDialog.getExistingDirectory(self, "Select GDDAQ Data Directory", start_dir)
        if chosen:
            self.search_dir_input.setText(chosen)

    def get_values(self):
        """获取设置值"""
        return (
            self.search_dir_input.text().strip(),
            self.proc_name_input.text().strip() or "gddaq",
            self.target_run_input.text().strip()
        )

class RealTimePlotApp(QMainWindow):     # 类: 主应用窗口
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Real-Time Current Monitoring System")

        # 默认模式: "serial" 或 "network"
        self.connection_mode = "serial"
        self.socket1 = None
        self.socket2 = None

        # 设置窗口图标
        icon_path = resource_path("logo.png") 
        if os.path.exists(icon_path):
            icon = QtGui.QIcon(icon_path)
            self.setWindowIcon(icon)
            # 同时设置应用程序图标
            QApplication.instance().setWindowIcon(icon)
        
        # 定义默认波特率变量
        self.baud_rate = 9600 
        
        self.setGeometry(100, 100, 1200, 800)
        
        # 初始化两个串口
        self.serialport1 = serial.Serial()
        if sys.platform.startswith('win'):
            self.serialport1.port = 'COM3'
        else:
            self.serialport1.port = '/dev/ttyUSB0'
            
        self.serialport1.baudrate = self.baud_rate 
        self.serialport1.parity = 'N'
        self.serialport1.bytesize = 8
        self.serialport1.stopbits = 1
        self.serialport1.timeout = 0.1
        
        self.serialport2 = serial.Serial()
        if sys.platform.startswith('win'):
            self.serialport2.port = 'COM4'
        else:
            self.serialport2.port = '/dev/ttyUSB1'
            
        self.serialport2.baudrate = self.baud_rate
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
        self.single_channel_mode = False # 默认为双通道模式
        self.unit_ch1 = "mA"    # 初始化通道单位，默认为 mA
        self.unit_ch2 = "mA"   
        self.unit_factors = {   # 定义单位到 mA 的转换系数
            "mA": 1.0,
            "μA": 0.001,
            "nA": 0.000001
        }
        # 初始化电流过滤阈值 (默认为 1000 mA)
        self.limit_ch1_ma = 1000.0 
        self.limit_ch2_ma = 1000.0

        # 脉冲提醒相关变量
        self.pulse_reminder_enabled = False  # 脉冲提醒开关，默认关闭
        self.pulse_reminder_timer = QTimer()  # 脉冲提醒定时器
        self.pulse_reminder_timer.timeout.connect(self.show_pulse_reminder)
        self.pulse_reminder_timer.setSingleShot(True)  # 单次触发
        self.reminder_suppressed = False  # 本轮是否已抑制提醒

        # ZERO 报警音相关变量
        self.alarm_enabled = True        # 报警音开关，默认开启
        self.alarm_active = False        # 报警音是否正在响（避免重复启停定时器）
        self.alarm_timer = QTimer()      # 报警周期定时器 (每3秒响一次)
        self.alarm_timer.setInterval(3000)  # 3000 ms
        self.alarm_timer.timeout.connect(self.on_alarm_tick)

        # 鼠标悬停相关属性
        self.hover_annotation = None       

        # DAQ 监控初始化
        self.daq_status_file = "/tmp/daq_status.txt"
        self.daq_last_mtime = 0
        self.daq_timer = QTimer()
        self.daq_timer.timeout.connect(self.check_daq_status)
        self.daq_timer.setInterval(200)  # 每200ms检查一次文件

        # GDDAQ 监控初始化
        self.gddaq_search_dir = ""          # 数据根目录 (如 /home/rnb/Data/202605ams/raw/)
        self.gddaq_proc_name = "gddaq"      # DAQ 进程名 (用于 pgrep -x)
        self.gddaq_target_run = ""          # 目标轮次 (空=自动监控最新轮次)
        self.gddaq_timer = QTimer()
        self.gddaq_timer.timeout.connect(self.check_gddaq_status)
        self.gddaq_timer.setInterval(1000)  # 每1秒检查一次
        self.gddaq_last_state = "IDLE"      # 上次状态: IDLE / RUNNING / FINISHED / CRASHED
        self.gddaq_last_log_mtime = 0       # run.log 上次 mtime
        self.gddaq_last_alive_time = None   # 进程最后存活时刻 (崩溃时冻结计时)
        self.gddaq_dir_invalid_logged = False  # 数据目录失效日志去重标志 (每 run 只记一次)

        # DAQ 模式互斥标记: None / "daq_master" / "gddaq"
        self.active_daq_mode = None

        # 状态监控初始化
        self.monitor1 = StatusMonitor()
        self.monitor2 = StatusMonitor()
        # 初始化参数 (根据默认 update_interval)
        self.monitor1.update_params(self.update_interval)
        self.monitor2.update_params(self.update_interval)

        # 日志系统初始化 (Log system initialization)
        self.log_bus = LogBus.get_instance()
        self.log_window = None  # 延迟创建, 首次打开时实例化 (lazy-init)
        # About / Tutorial 对话框单实例引用 (非模态, 可同时操作主窗口)
        self._about_dialog = None
        self._tutorial_dialog = None
        # recv_data 连续无响应去抖计数器 (debounce counters for recv timeouts)
        # 达阈值(~5秒无数据)记一条 WARNING, 恢复时记 INFO
        self.recv_none_count_1 = 0
        self.recv_none_count_2 = 0
        self.recv_warned_1 = False
        self.recv_warned_2 = False

        # 然后创建UI和菜单栏
        self.init_ui()
        self.create_menu_bar()
        # 订阅日志更新信号 (Subscribe to log update signal)
        self.log_bus.log_updated.connect(self.on_log_updated)

        # 创建定时器
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_data)
        self.timer.start(self.update_interval)

        # 记录启动日志 (Record startup log)
        self.log_bus.log("INFO", "Application started")
        
    def create_menu_bar(self):
        """创建菜单栏"""
        menu_bar = self.menuBar()
        
        ## 文件菜单
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
        

        ## 连接菜单
        conn_menu = menu_bar.addMenu('Connection')
        
        self.mode_serial_action = QtWidgets.QAction('Serial Port Mode', self)
        self.mode_serial_action.setCheckable(True)
        self.mode_serial_action.setChecked(True)
        self.mode_serial_action.triggered.connect(lambda: self.switch_connection_mode("serial"))
        conn_menu.addAction(self.mode_serial_action)
        
        self.mode_network_action = QtWidgets.QAction('TCP Network Mode', self)
        self.mode_network_action.setCheckable(True)
        self.mode_network_action.setChecked(False)
        self.mode_network_action.triggered.connect(lambda: self.switch_connection_mode("network"))
        conn_menu.addAction(self.mode_network_action)

        # 互斥组，确保只能选一个
        mode_group = QtWidgets.QActionGroup(self)
        mode_group.addAction(self.mode_serial_action)
        mode_group.addAction(self.mode_network_action)

        ## 运行菜单
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

        # 单通道模式开关菜单项
        self.single_mode_action = QtWidgets.QAction('Single Channel Mode (CH1 Only)', self)
        self.single_mode_action.setCheckable(True)
        self.single_mode_action.setChecked(False)
        self.single_mode_action.triggered.connect(self.toggle_single_mode)
        self.single_mode_action.setShortcut('Ctrl+Shift+S')
        self.single_mode_action.setToolTip("Enable to monitor only Channel 1. Cannot be changed while running.")
        run_menu.addAction(self.single_mode_action)

        # run_menu.addSeparator()

        # 脉冲提醒开关菜单项
        self.pulse_reminder_action = QtWidgets.QAction('Pulse Reminder', self)
        self.pulse_reminder_action.setCheckable(True)
        self.pulse_reminder_action.setChecked(False)  # 默认关闭
        self.pulse_reminder_action.triggered.connect(self.toggle_pulse_reminder)
        self.pulse_reminder_action.setShortcut('Ctrl+Shift+P')
        self.pulse_reminder_action.setToolTip("Enable/Disable Pulse Reminder")
        run_menu.addAction(self.pulse_reminder_action)

        # run_menu.addSeparator()

        # DAQ 连接二级菜单
        daq_submenu = run_menu.addMenu('Connect to DAQ')

        # DAQ_Master 选项 (原 Connect to DAQ (beta))
        self.daq_connect_action = QtWidgets.QAction('DAQ_Master', self)
        self.daq_connect_action.setCheckable(True)
        self.daq_connect_action.setChecked(False)
        self.daq_connect_action.triggered.connect(self.toggle_daq_connection)
        self.daq_connect_action.setToolTip("Sync Start/Stop and Filename with DAQ_Master System")
        daq_submenu.addAction(self.daq_connect_action)

        # GDDAQ 选项
        self.gddaq_connect_action = QtWidgets.QAction('GDDAQ', self)
        self.gddaq_connect_action.setCheckable(True)
        self.gddaq_connect_action.setChecked(False)
        self.gddaq_connect_action.triggered.connect(self.toggle_gddaq_connection)
        self.gddaq_connect_action.setToolTip("Sync Start/Stop with GDDAQ System via run.log monitoring")
        daq_submenu.addAction(self.gddaq_connect_action)

        # 注意：不使用 QActionGroup 的互斥功能 (exclusive 默认 True 会阻止取消勾选)。
        # 互斥逻辑由 toggle_daq_connection / toggle_gddaq_connection 内部处理：
        # 启用一方时主动取消另一方的勾选并调用其 toggle 方法。

        daq_submenu.addSeparator()

        # GDDAQ 设置入口
        self.gddaq_settings_action = QtWidgets.QAction('GDDAQ Settings...', self)
        self.gddaq_settings_action.triggered.connect(self.open_gddaq_settings)
        self.gddaq_settings_action.setToolTip("Configure GDDAQ data directory, process name and run number")
        daq_submenu.addAction(self.gddaq_settings_action)

        run_menu.addSeparator()

        # 设置通道单位菜单项
        self.set_units_action = QtWidgets.QAction('Set Channel Units', self)
        self.set_units_action.triggered.connect(self.set_channel_units)
        self.set_units_action.setToolTip("Configure measurement units (mA, μA, nA)")
        run_menu.addAction(self.set_units_action)

        # run_menu.addSeparator()

        # 更新间隔设置菜单项
        self.update_interval_action = QtWidgets.QAction('Set Update Interval', self)
        self.update_interval_action.triggered.connect(self.set_update_interval)
        self.update_interval_action.setToolTip("Set Data Update Interval")
        run_menu.addAction(self.update_interval_action)
        
        # 设置波特率的菜单项
        self.set_baudrate_action = QtWidgets.QAction('Set Baud Rate', self)
        self.set_baudrate_action.triggered.connect(self.set_baud_rate)
        self.set_baudrate_action.setToolTip("Set Serial Port Baud Rate (Default: 9600)")
        run_menu.addAction(self.set_baudrate_action)        

        # 状态监控设置菜单项
        self.monitor_settings_action = QtWidgets.QAction('Status Monitor Settings', self)
        self.monitor_settings_action.triggered.connect(self.open_monitor_settings)
        run_menu.addAction(self.monitor_settings_action)

        # 设置电流阈值菜单项
        self.set_limit_action = QtWidgets.QAction('Set Current Threshold', self)
        self.set_limit_action.triggered.connect(self.set_current_threshold)
        self.set_limit_action.setToolTip("Set the maximum current limit for filtering noise")
        run_menu.addAction(self.set_limit_action)

        ## 帮助菜单
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

    def toggle_single_mode(self):
        """切换单通道/双通道模式"""
        self.single_channel_mode = self.single_mode_action.isChecked()
        
        # 视觉反馈：禁用/启用通道2的输入框和测试按钮
        is_dual = not self.single_channel_mode
        self.port2_input.setEnabled(is_dual)
        self.test_serial2_button.setEnabled(is_dual)
        
        # 更新标签提示
        if self.single_channel_mode:
            self.current2_label.setText("Channel 2 Current: --- (Disabled)")
            self.current2_label.setStyleSheet("font-size: 14px; font-weight: bold; color: gray;")
        else:
            self.current2_label.setText("Channel 2 Current: --- mA")
            self.current2_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #ff7f0e;")
            
        print(f"Mode Switched: {'Single Channel' if self.single_channel_mode else 'Dual Channel'}")
        self.log_bus.log("INFO", f"Mode switched to {'Single' if self.single_channel_mode else 'Dual'} Channel")

    def toggle_daq_connection(self):    #  DAQ_Master 联动功能实现
        """切换 DAQ_Master 连接模式 (与 GDDAQ 互斥)"""
        is_connected = self.daq_connect_action.isChecked()
        
        if is_connected:
            # 互斥：如果 GDDAQ 正在运行，先关闭它
            if self.active_daq_mode == "gddaq":
                self.gddaq_connect_action.setChecked(False)
                self.toggle_gddaq_connection()

            self.active_daq_mode = "daq_master"

            # 开启模式：锁定文件名输入，初始化时间戳，启动定时器
            self.filename_input.setEnabled(False)
            self.browse_button.setEnabled(False)
            self.file_mode_combo.setEnabled(False)
            
            # 初始化 last_mtime，忽略开启前的旧状态（参考你的 monitor_daq 逻辑）
            if os.path.exists(self.daq_status_file):
                try:
                    self.daq_last_mtime = os.path.getmtime(self.daq_status_file)
                except OSError:
                    self.daq_last_mtime = 0
            
            self.daq_timer.start()
            print("DAQ_Master Connection Enabled: Monitoring started.")
            self.log_bus.log("INFO", "DAQ_Master auto-link enabled")
            self.save_status_label.setText("DAQ_Master Mode: Waiting for signal...")
            self.save_status_label.setStyleSheet("color: blue;")
        else:
            self.active_daq_mode = None

            # 关闭模式：停止定时器，恢复输入框
            self.daq_timer.stop()

            # 如果正在运行，停止监控
            if self.run_stat:
                self.stop_monitoring(source="daq_disabled")

            self.filename_input.setEnabled(True)
            self.browse_button.setEnabled(True)
            self.file_mode_combo.setEnabled(True)
            print("DAQ_Master Connection Disabled.")
            self.log_bus.log("INFO", "DAQ_Master auto-link disabled")
            self.save_status_label.setText("DAQ_Master Mode: Disabled")
            self.save_status_label.setStyleSheet("color: black;")

    def check_daq_status(self):
        """定时检查 DAQ 状态文件"""
        if not os.path.exists(self.daq_status_file):
            return

        try:
            current_mtime = os.path.getmtime(self.daq_status_file)
            
            # 只有文件被修改过才读取
            if current_mtime > self.daq_last_mtime:
                self.daq_last_mtime = current_mtime
                
                # 读取文件内容
                with open(self.daq_status_file, 'r') as f:
                    lines = [line.strip() for line in f.readlines()]
                
                if not lines:
                    return

                status_line = lines[0]

                if "STATUS: RUNNING" in status_line:
                    # 如果已经在运行，先不处理，或者可以选择重启监控
                    if self.run_stat:
                        print("DAQ Signal: START received, but already running. Ignoring.")
                        return

                    if len(lines) >= 4:
                        run_name = lines[1]
                        try:
                            run_num = int(lines[2])
                        except ValueError:
                            run_num = 0
                        data_path = lines[3]

                        # 1. 确定保存路径：data_path 的上一级 + current_data
                        parent_dir = os.path.dirname(data_path)
                        save_dir = os.path.join(parent_dir, "current_data")
                        
                        # 创建文件夹
                        if not os.path.exists(save_dir):
                            try:
                                os.makedirs(save_dir, exist_ok=True)
                            except Exception as e:
                                print(f"Error creating directory: {e}")
                                return

                        # 2. 确定文件名：运行名称_运行编号 (5位)
                        # 例如: run_00048.csv
                        filename = f"{run_name}_{run_num:05d}.csv"
                        full_path = os.path.join(save_dir, filename)

                        # 3. 更新 UI 并启动
                        self.filename_input.setText(full_path)
                        print(f"DAQ Signal: START. File: {full_path}")
                        self.log_bus.log("INFO", f"DAQ_Master signal: START (file={filename})")

                        # 确保使用覆盖模式或追加模式，这里默认追加即可，因为是新文件
                        self.start_monitoring(source="daq_master")
                        
                        # 更新状态提示
                        self.save_status_label.setText(f"DAQ Linked: Running {filename}")
                        self.save_status_label.setStyleSheet("color: green;")

                elif "STATUS: STOPPED" in status_line:
                    if self.run_stat:
                        print("DAQ Signal: STOP received.")
                        self.log_bus.log("INFO", "DAQ_Master signal: STOP")
                        self.stop_monitoring(source="daq_master")
                        self.save_status_label.setText("DAQ Linked: Stopped")
                        self.save_status_label.setStyleSheet("color: orange;")

        except Exception as e:
            print(f"Error checking DAQ status: {e}")
            self.log_bus.log("WARNING", f"DAQ status file read error: {e}")

    # ===================== GDDAQ 联动功能 =====================

    def open_gddaq_settings(self):
        """打开 GDDAQ 设置对话框"""
        dialog = GDDAQSettingsDialog(
            self.gddaq_search_dir,
            self.gddaq_proc_name,
            self.gddaq_target_run,
            self
        )
        if dialog.exec_() == QDialog.Accepted:
            self.gddaq_search_dir, self.gddaq_proc_name, self.gddaq_target_run = dialog.get_values()
            print(f"GDDAQ Settings Updated: dir={self.gddaq_search_dir}, proc={self.gddaq_proc_name}, run={self.gddaq_target_run or '(auto)'}")
            self.log_bus.log("INFO",
                f"GDDAQ settings applied: dir={self.gddaq_search_dir}, "
                f"proc={self.gddaq_proc_name}, run={self.gddaq_target_run or '(auto)'}")

    def toggle_gddaq_connection(self):
        """切换 GDDAQ 连接模式 (与 DAQ_Master 互斥)"""
        is_connected = self.gddaq_connect_action.isChecked()

        if is_connected:
            # Windows 不支持 GDDAQ (依赖 pgrep)
            if sys.platform.startswith('win'):
                QMessageBox.warning(self, "Warning",
                    "GDDAQ mode is only available on Linux (requires pgrep).\n"
                    "Please use DAQ_Master mode instead.")
                self.gddaq_connect_action.setChecked(False)
                return

            # 互斥：如果 DAQ_Master 正在运行，先关闭它
            if self.active_daq_mode == "daq_master":
                self.daq_connect_action.setChecked(False)
                self.toggle_daq_connection()

            # 检查必填配置: 数据目录和进程名均不能为空
            # Required config check: both data directory and process name must be set
            if not self.gddaq_search_dir or not self.gddaq_proc_name:
                QMessageBox.warning(self, "Warning",
                    "GDDAQ Data Directory and Process Name must be set!\n"
                    "Please configure them via 'Connect to DAQ' -> 'GDDAQ Settings...' first.")
                self.gddaq_connect_action.setChecked(False)
                return

            if not os.path.isdir(self.gddaq_search_dir):
                QMessageBox.warning(self, "Warning",
                    f"GDDAQ data directory does not exist:\n{self.gddaq_search_dir}")
                self.gddaq_connect_action.setChecked(False)
                return

            self.active_daq_mode = "gddaq"

            # 锁定文件名输入 (与 DAQ_Master 行为一致)
            self.filename_input.setEnabled(False)
            self.browse_button.setEnabled(False)
            self.file_mode_combo.setEnabled(False)

            # 重置状态
            self.gddaq_last_state = "IDLE"
            self.gddaq_last_log_mtime = 0
            self.gddaq_last_alive_time = None
            self.gddaq_dir_invalid_logged = False  # 重置目录失效去重标志

            self.gddaq_timer.start()
            print("GDDAQ Connection Enabled: Monitoring started.")
            self.log_bus.log("INFO", "GDDAQ auto-link enabled")
            self.save_status_label.setText("GDDAQ Mode: Waiting for run.log signal...")
            self.save_status_label.setStyleSheet("color: blue;")
        else:
            # 关闭模式
            self.gddaq_timer.stop()
            self.active_daq_mode = None

            # 如果正在运行，停止监控
            if self.run_stat:
                self.stop_monitoring(source="daq_disabled")

            # 恢复 UI
            self.filename_input.setEnabled(True)
            self.browse_button.setEnabled(True)
            self.file_mode_combo.setEnabled(True)
            print("GDDAQ Connection Disabled.")
            self.log_bus.log("INFO", "GDDAQ auto-link disabled")
            self.save_status_label.setText("GDDAQ Mode: Disabled")
            self.save_status_label.setStyleSheet("color: black;")

    def is_gddaq_running(self):
        """通过 pgrep -x 检测 DAQ 进程是否存活 (参考 run_timer.py)"""
        try:
            result = subprocess.run(
                ["pgrep", "-x", self.gddaq_proc_name],
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
            )
            return result.returncode == 0
        except Exception:
            # pgrep 不存在 (如 Windows) 或其他错误，视为未运行
            return False

    def get_gddaq_run_dir(self):
        """
        返回目标轮次的数据目录路径。
        - target_run 为空: 返回 search_dir 下最大编号的数字子目录
        - target_run 非空: 返回 search_dir/<target_run> 目录
        返回 (run_dir, run_id_str) 或 (None, error_msg)
        参考 run_timer.py 的 get_run_data() 逻辑。
        """
        try:
            # 获取所有数字命名的子目录
            dirs = [
                d for d in os.listdir(self.gddaq_search_dir)
                if os.path.isdir(os.path.join(self.gddaq_search_dir, d)) and d.isdigit()
            ]

            if not dirs:
                return None, "No numbered run folders found"

            if self.gddaq_target_run:
                # 指定轮次
                target = self.gddaq_target_run.lstrip('0') or '0'
                # 尝试匹配: 数字比较
                matched = None
                for d in dirs:
                    if int(d) == int(target):
                        matched = d
                        break
                if matched is None:
                    return None, f"Run {self.gddaq_target_run} not found"
                latest_dir = matched
            else:
                # 自动取最大编号
                latest_dir = max(dirs, key=int)

            return os.path.join(self.gddaq_search_dir, latest_dir), latest_dir

        except Exception as e:
            return None, str(e)

    def parse_gddaq_run_log(self, log_path):
        """
        解析 run.log 文件，返回 (start_time, stop_time)。
        参考 run_timer.py 的解析逻辑:
        - 一个 run.log 可能含多个 Start/Stop 块 (重跑场景)，只取最后一个 Start 及其后是否有 Stop
        - Start 行格式: "Start: 2025-12-10 17:21:49"
        - Stop 行格式: "Stop : 2025-12-10 17:21:47" (注意 Stop 后有空格，用 split(":", 1) 处理)
        - 新的 Start 会使其之前的 Stop 失效
        返回 (datetime, datetime or None)，解析失败返回 (None, None)
        """
        start_time = None
        stop_time = None

        try:
            with open(log_path, "r") as f:
                for line in f:
                    if line.startswith("Start:"):
                        t_str = line.split("Start:")[1].strip()
                        try:
                            start_time = datetime.strptime(t_str, "%Y-%m-%d %H:%M:%S")
                            # 新的 Start 使之前的 Stop 失效
                            stop_time = None
                        except ValueError:
                            pass
                    elif line.startswith("Stop"):
                        # 格式 "Stop : 2025-12-10 17:21:47" — 用 split(":", 1) 松散处理
                        parts = line.split(":", 1)
                        if len(parts) > 1:
                            t_str = parts[1].strip()
                            try:
                                stop_time = datetime.strptime(t_str, "%Y-%m-%d %H:%M:%S")
                            except ValueError:
                                pass
        except Exception:
            pass

        return start_time, stop_time

    def check_gddaq_status(self):
        """
        定时检查 GDDAQ 状态 (参考 run_timer.py 的状态机)。
        状态流转:
        - IDLE → 检测到 Start 且无 Stop → RUNNING (自动 start_monitoring)
        - RUNNING → 检测到 Stop → FINISHED (自动 stop_monitoring)
        - RUNNING → pgrep 失败且无 Stop → CRASHED (自动 stop_monitoring，状态栏提示)
        - 每次轮询记录上次状态，避免重复触发启停
        """
        try:
            # 0. 检查数据根目录有效性 (运行期间可能被删除/卸载, 如 NFS 掉线)
            # Check data root validity (may be deleted/unmounted during a run, e.g. NFS dropout)
            if not os.path.isdir(self.gddaq_search_dir):
                if not self.gddaq_dir_invalid_logged:
                    self.log_bus.log("WARNING",
                        f"GDDAQ: data directory became invalid: {self.gddaq_search_dir}")
                    self.gddaq_dir_invalid_logged = True
                self.save_status_label.setText("GDDAQ: data directory invalid!")
                self.save_status_label.setStyleSheet("color: red;")
                return
            else:
                # 目录恢复有效时重置去重标志, 下次失效可再次记录
                self.gddaq_dir_invalid_logged = False

            # 1. 获取目标轮次目录
            run_dir, run_id = self.get_gddaq_run_dir()
            if run_dir is None:
                # 目录不存在或无轮次文件夹，保持等待
                if self.gddaq_last_state != "IDLE":
                    print(f"GDDAQ: {run_id}")
                    self.gddaq_last_state = "IDLE"
                return

            log_path = os.path.join(run_dir, "run.log")

            if not os.path.exists(log_path):
                # run.log 尚未创建，保持等待
                return

            # 2. 检查 mtime 是否变化 (减少 I/O)
            try:
                current_log_mtime = os.path.getmtime(log_path)
            except OSError:
                return

            log_changed = current_log_mtime > self.gddaq_last_log_mtime
            if log_changed:
                self.gddaq_last_log_mtime = current_log_mtime

            # 3. 解析 run.log
            start_time, stop_time = self.parse_gddaq_run_log(log_path)

            # 4. 检测进程存活
            proc_alive = self.is_gddaq_running()

            # 5. 状态判定 (参考 run_timer.py)
            now = datetime.now()

            if start_time is None:
                # 还没有 Start 行，等待中
                if self.gddaq_last_state != "IDLE":
                    self.gddaq_last_state = "IDLE"
                    self.save_status_label.setText(f"GDDAQ: Run {run_id} waiting for Start...")
                    self.save_status_label.setStyleSheet("color: blue;")
                return

            if stop_time is not None:
                # 有 Stop → 已结束
                if self.gddaq_last_state != "FINISHED":
                    print(f"GDDAQ: Run {run_id} FINISHED (Stop detected in run.log)")
                    self.log_bus.log("INFO", f"GDDAQ: run {run_id} finished")
                    self.gddaq_last_state = "FINISHED"
                    if self.run_stat:
                        self.stop_monitoring(source="gddaq")
                    self.save_status_label.setText(f"GDDAQ: Run {run_id} Finished")
                    self.save_status_label.setStyleSheet("color: green;")
                return

            # 无 Stop 的情况
            if proc_alive:
                # 进程存活且无 Stop → 正在运行
                self.gddaq_last_alive_time = now

                if self.gddaq_last_state != "RUNNING":
                    print(f"GDDAQ: Run {run_id} RUNNING (Start detected, process alive)")
                    self.log_bus.log("INFO", f"GDDAQ: run {run_id} started")
                    self.gddaq_last_state = "RUNNING"

                    # 生成保存路径: <search_dir>/CurrentData/run_<run_id:05d>.csv
                    save_dir = os.path.join(self.gddaq_search_dir, "CurrentData")
                    if not os.path.exists(save_dir):
                        try:
                            os.makedirs(save_dir, exist_ok=True)
                        except Exception as e:
                            print(f"GDDAQ: Error creating save directory: {e}")
                            self.log_bus.log("ERROR", f"GDDAQ: cannot create save directory: {e}")
                            return

                    run_num = int(run_id)
                    filename = f"run_{run_num:05d}.csv"
                    full_path = os.path.join(save_dir, filename)

                    # 更新 UI 文件名并启动监控
                    if not self.run_stat:
                        self.filename_input.setText(full_path)
                        print(f"GDDAQ: Auto-starting monitoring. File: {full_path}")
                        self.start_monitoring(source="gddaq")

                    self.save_status_label.setText(f"GDDAQ: Running {filename}")
                    self.save_status_label.setStyleSheet("color: green;")
                else:
                    # 已在 RUNNING 状态，仅更新运行时间显示
                    if self.gddaq_last_alive_time is not None:
                        runtime = now - self.gddaq_last_alive_time
                        # 不频繁更新，避免刷屏
            else:
                # 进程不存活且无 Stop → 可能崩溃
                # 参考 run_timer.py: 冻结计时在最后存活时刻
                if self.gddaq_last_alive_time is None:
                    self.gddaq_last_alive_time = now

                if self.gddaq_last_state != "CRASHED":
                    print(f"GDDAQ: Run {run_id} CRASHED (process '{self.gddaq_proc_name}' not found, no Stop in log)")
                    self.log_bus.log("WARNING", f"GDDAQ: process died (run {run_id})")
                    self.gddaq_last_state = "CRASHED"

                    if self.run_stat:
                        self.stop_monitoring(source="gddaq")

                    self.save_status_label.setText(f"GDDAQ: Run {run_id} CRASHED! Process '{self.gddaq_proc_name}' not found.")
                    self.save_status_label.setStyleSheet("color: red;")

        except Exception as e:
            print(f"Error checking GDDAQ status: {e}")
            import traceback
            traceback.print_exc()

    # ===================== GDDAQ 联动功能结束 =====================

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
        # 路径自动补全 (含目录和文件, 方便选择已有 CSV) (Path auto-completion: dirs + files)
        _file_completer = QCompleter(self)
        _file_model = QDirModel()
        _file_model.setFilter(QtCore.QDir.Dirs | QtCore.QDir.Files | QtCore.QDir.NoDotAndDotDot)
        _file_completer.setModel(_file_model)
        _file_completer.setCompletionMode(QCompleter.PopupCompletion)
        _file_completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.filename_input.setCompleter(_file_completer)
        
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
        
        # 双通道显示标签
        self.current1_label = QLabel("Channel 1 Current: --- mA")
        self.current2_label = QLabel("Channel 2 Current: --- mA")
        self.runtime_label = QLabel("Run Time: ---")
        self.integral1_label = QLabel("Channel 1 Integral: --- mC")
        self.integral2_label = QLabel("Channel 2 Integral: --- mC")
        self.timestamp_label = QLabel("Last Update Time (Local): ---")
        self.utc_timestamp_label = QLabel("UTC Timestamp: ---")
        self.save_status_label = QLabel("Save Status: Not Saved")
        
        # 状态监控标签
        self.status_label1 = ClickableLabel("STOP")
        self.status_label1.setFixedSize(80, 50) # 设置合适的大小
        self.status_label1.clicked.connect(self.monitor1.clear_warning) # 点击清除警告
        
        self.status_label2 = ClickableLabel("STOP")
        self.status_label2.setFixedSize(80, 50)
        self.status_label2.clicked.connect(self.monitor2.clear_warning)

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
        status_layout.addWidget(QLabel("Dual-Channel Monitoring Status:"), 0, 0, 1, 3)
        # 第一行：Ch1 Current | Ch1 Status
        status_layout.addWidget(self.current1_label, 1, 0)
        status_layout.addWidget(self.current2_label, 1, 1)
        # 将状态栏放在第2列 (最右侧)
        status_layout.addWidget(self.status_label1, 1, 2) 
        
        # 第二行：Ch2 Current | Ch2 Status
        status_layout.addWidget(self.integral1_label, 2, 0)
        status_layout.addWidget(self.integral2_label, 2, 1)
        status_layout.addWidget(self.status_label2, 2, 2)
        
        # 其他行保持不变
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

        # 左侧 Y 轴 - 通道 1 (蓝色)
        self.ax.set_ylabel(f'Channel 1 Current ({self.unit_ch1})', color='#1f77b4', fontweight='bold')
        self.ax.tick_params(axis='y', labelcolor='#1f77b4')
        
        # 右侧 Y 轴 - 通道 2 (橙色) - 共享 X 轴
        self.ax2 = self.ax.twinx()
        self.ax2.set_ylabel(f'Channel 2 Current ({self.unit_ch2})', color='#ff7f0e', fontweight='bold')
        self.ax2.tick_params(axis='y', labelcolor='#ff7f0e')
        
        # 创建两条线，分别绑定到不同的轴
        self.line1, = self.ax.plot(self.x_data, self.y_data1, 'b-', color='#1f77b4', label='Channel 1', linewidth=2)
        self.line2, = self.ax2.plot(self.x_data, self.y_data2, 'r-', color='#ff7f0e', label='Channel 2', linewidth=2)
        
        # 合并图例 (因为有两个轴，需要手动收集图例句柄)
        lines = [self.line1, self.line2]
        labels = [l.get_label() for l in lines]
        self.ax.legend(lines, labels, loc='best')
        self.ax.grid(True, alpha=0.3)

        # 鼠标悬停功能
        self.setup_mouse_hover()
        
        # 添加到主布局
        main_layout.addLayout(serial_layout)
        main_layout.addLayout(file_layout)
        main_layout.addLayout(status_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.canvas)

        # 底部日志栏 (一行, 显示最近一条日志, 点击打开日志窗口)
        # Bottom log bar: one line, shows latest entry, click to open log window
        self.log_bar = ClickableLabel("Event Log: (no events yet)")
        self.log_bar.setFixedHeight(24)
        self.log_bar.setStyleSheet(
            "background: #f0f0f0; border-top: 1px solid #ccc; "
            "padding: 2px 6px; color: #555555; font-size: 12px; "
            "qproperty-alignment: AlignVCenter; border-radius: 0px;"
        )
        self.log_bar.setToolTip("Click to view full event log")
        self.log_bar.clicked.connect(self.open_log_window)
        main_layout.addWidget(self.log_bar)

    def on_log_updated(self):
        """日志总线更新回调: 刷新底部日志栏文本/颜色, 并在窗口可见时刷新窗口
        (Log bus update callback: refresh bottom bar text/color and window if visible)"""
        if not hasattr(self, 'log_bar'):
            return
        text = self.log_bus.get_latest_text()
        level = self.log_bus.get_latest_level()
        color = LogBus.LEVEL_COLORS.get(level, "#555555")
        self.log_bar.setText(text)
        self.log_bar.setStyleSheet(
            f"background: #f0f0f0; border-top: 1px solid #ccc; "
            f"padding: 2px 6px; color: {color}; font-size: 12px; "
            f"qproperty-alignment: AlignVCenter; border-radius: 0px;"
        )
        # 若日志窗口已打开, 同步刷新 (Sync the log window if it is visible)
        if self.log_window is not None and self.log_window.isVisible():
            self.log_window.refresh()

    def open_log_window(self):
        """打开日志窗口 (复用并提至前台, 不弹新窗口)
        Open the log window (reuse & raise to front, no new window)."""
        if self.log_window is None:
            self.log_window = LogWindow(self)
        self.log_window.show()
        self.log_window.raise_()
        self.log_window.activateWindow()

    def test_serial_connection(self, channel):
        """测试连接（串口和网络）"""
        input_text = self.port1_input.text().strip() if channel == 1 else self.port2_input.text().strip()   # 获取输入内容
        channel_name = f"Channel {channel}"
        
        if not input_text:
            QMessageBox.warning(self, "Warning", f"Please Enter {channel_name} Configuration")
            return

        if self.connection_mode == "serial":
            # 串口测试
            serialport = self.serialport1 if channel == 1 else self.serialport2
            try:
                serialport.port = input_text
                if serialport.is_open: serialport.close()
                serialport.open()
                serialport.close()
                
                QMessageBox.information(self, "Test Successful", f"{channel_name} Serial Port Test Successful!")
                self.save_status_label.setText(f"{channel_name} Serial Port Test: Successful")
                self.save_status_label.setStyleSheet("color: green;")
                self.log_bus.log("INFO", f"Channel {channel} serial port test successful")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"{channel_name} Serial Port Test Failed: {str(e)}")
                self.save_status_label.setText(f"{channel_name} Serial Port Test: Failed - {str(e)}")
                self.save_status_label.setStyleSheet("color: red;")
                self.log_bus.log("WARNING", f"Channel {channel} serial port test failed: {e}")
        
        else:
            # 网络测试
            try:
                if ":" not in input_text:
                    raise ValueError("Invalid Format. Use IP:Port (e.g., 192.168.1.253:1030)")
                
                ip, port_str = input_text.split(":")
                port = int(port_str)
                
                # 创建临时 socket 测试连接
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(2) # 2秒超时
                s.connect((ip, port))
                s.close()
                
                QMessageBox.information(self, "Connect Successful", f"{channel_name} Network Connect Successful!")
                self.save_status_label.setText(f"{channel_name} Network Test: Successful")
                self.save_status_label.setStyleSheet("color: green;")
                self.log_bus.log("INFO", f"Channel {channel} network test successful")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"{channel_name} Network Test Failed: {str(e)}")
                self.save_status_label.setText(f"{channel_name} Network Test: Failed - {str(e)}")
                self.save_status_label.setStyleSheet("color: red;")
                self.log_bus.log("WARNING", f"Channel {channel} network test failed: {e}")

    
    def file_mode_changed(self, index):
        """文件模式改变时的处理"""
        self.file_mode = "append" if index == 0 else "overwrite"
        print(f"File Mode Changed to: {self.file_mode}")
        self.log_bus.log("INFO", f"File mode set to {self.file_mode}")
    
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
            self.log_bus.log("ERROR", f"Data file open failed: {e}")
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
                self.log_bus.log("ERROR", f"Data file close failed: {e}")
            finally:
                self.file_handle = None
    
    def write_data_row(self, time_val, runtime, current1, current2, integral1, integral2):
        """写入一行数据到文件"""
        if not self.file_handle:
            return
        
        try:
            # 写入双通道数据
            self.file_handle.write(f"{time_val:.1f},{runtime:.4f},{current1:.8e},{current2:.8e},{integral1:.4e},{integral2:.4e}\n")
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
            self.log_bus.log("WARNING", "Snapshot failed: data file empty")
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
            self.log_bus.log("INFO", f"Snapshot created: {os.path.basename(snapshot_file)}")
            return True
            
        except Exception as e:
            error_msg = f"Snapshot Creation Failed: {str(e)}"
            QMessageBox.critical(self, "Error", error_msg)
            self.save_status_label.setText(f"Snapshot Status: {error_msg}")
            self.save_status_label.setStyleSheet("color: red;")
            print(f"Failed to Create Data Snapshot: {e}")
            self.log_bus.log("ERROR", f"Snapshot failed: {e}")
            return False
        
    def extract_number_from_filename(self, filename):
        """从文件名中提取末尾的四位数字"""
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
        
    def switch_connection_mode(self, mode):
        """切换连接模式：串口 or 网络"""
        if self.run_stat:
            QMessageBox.warning(self, "Warning", "Cannot switch mode while monitoring is running!")

            # 恢复勾选状态
            if self.connection_mode == "serial":
                self.mode_serial_action.setChecked(True)
            else:
                self.mode_network_action.setChecked(True)
            return

        self.connection_mode = mode
        
        if mode == "network":
            self.port1_label.setText("Channel 1 Address (IP:Port):")
            self.port1_input.setToolTip("Format: 192.168.1.253:1030")
            self.port1_input.setText("192.168.1.253:1030") # 默认值示例
            
            self.port2_label.setText("Channel 2 Address (IP:Port):")
            self.port2_input.setToolTip("Format: 192.168.1.253:1031")
            self.port2_input.setText("192.168.1.253:1031")  # 默认值示例
            
            self.test_serial1_button.setText("Test Network 1")
            self.test_serial2_button.setText("Test Network 2")
        else:
            self.port1_label.setText("Channel 1 Serial Port:")
            self.port1_input.setToolTip("Enter Channel 1 Serial Port Device Name")
            # 恢复默认串口名，根据系统判断
            default_port1 = 'COM3' if sys.platform.startswith('win') else '/dev/ttyUSB0'
            self.port1_input.setText(default_port1)

            self.port2_label.setText("Channel 2 Serial Port:")
            self.port2_input.setToolTip("Enter Channel 2 Serial Port Device Name")
            # 恢复默认串口名，根据系统判断
            default_port2 = 'COM4' if sys.platform.startswith('win') else '/dev/ttyUSB1'
            self.port2_input.setText(default_port2) 

            self.test_serial1_button.setText("Test Port 1")
            self.test_serial2_button.setText("Test Port 2")

        print(f"Switched to {mode} mode")
        self.log_bus.log("INFO", f"Connection mode switched to {mode}")

    def send_data(self, channel):
        """发送数据请求到指定通道"""
        slave_address = 1
        function_code = 3
        start_address = 42
        quantity = 2
        
        try:
            request = build_request(slave_address, function_code, start_address, quantity)

            if self.connection_mode == "serial":
                # 串口发送
                if channel == 1:
                    self.serialport1.write(request)
                else:
                    self.serialport2.write(request)
            else:
                # 网络发送
                sock = self.socket1 if channel == 1 else self.socket2
                if sock:
                    sock.sendall(request)                
        except Exception as e:
            print(f"Failed to Send Request to Channel {channel}: {e}")
            self.log_bus.log("WARNING", f"Channel {channel} send request failed: {e}")

    def recv_data(self, channel):
        """接收并解析指定通道的数据"""
        try:
            response = b''

            if self.connection_mode == "serial":
                # 串口接收
                port = self.serialport1 if channel == 1 else self.serialport2
                response = port.read(9) # 期望读取9个字节
            else:
                # 网络接收
                sock = self.socket1 if channel == 1 else self.socket2
                if sock:
                    # 网络接收可能需要一点缓冲，这里简单处理
                    # 因为是透传，仪表回传的也是Modbus RTU帧，长度固定为9字节
                    # (地址1 + 功能1 + 字节数1 + 数据4 + CRC2 = 9)
                    response = sock.recv(1024) 
            
            if len(response) < 9:
                # 可以在这里加个日志，但不抛出异常以免刷屏
                return None
            
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
    
    def start_monitoring(self, source="manual"):
        """开始监控
        source: 'manual' / 'daq_master' / 'gddaq' — 标识触发来源, 用于日志"""
        # GDDAQ 模式下启动前再次校验数据目录有效性 (勾选后目录可能被删除/卸载)
        # Re-validate GDDAQ data directory before starting (it may have become invalid)
        if self.active_daq_mode == "gddaq":
            if not self.gddaq_search_dir or not os.path.isdir(self.gddaq_search_dir):
                self.log_bus.log("ERROR", "Start failed: GDDAQ data directory invalid")
                QMessageBox.critical(self, "Error",
                    f"GDDAQ data directory is invalid:\n"
                    f"{self.gddaq_search_dir or '(not set)'}\n"
                    "Monitoring cannot start. Please check GDDAQ Settings.")
                return

        # 获取串口名/网络地址
        addr1 = self.port1_input.text().strip()
        addr2 = self.port2_input.text().strip()
        
        if not addr1:
            self.log_bus.log("WARNING", "Start failed: Channel 1 configuration empty")
            QMessageBox.warning(self, "Warning", "Please Enter Channel 1 Configuration!")
            return
        if not self.single_channel_mode and not addr2:
            self.log_bus.log("WARNING", "Start failed: Channel 2 configuration empty")
            QMessageBox.warning(self, "Warning", "Please Enter Channel 2 Configuration!")
            return
            
        try:
            if self.connection_mode == "serial":
                # 串口模式
                self.serialport1.port = addr1
                if self.serialport1.is_open: self.serialport1.close()
                self.serialport1.open()
                self.serialport1.reset_input_buffer()
                
                if not self.single_channel_mode:
                    self.serialport2.port = addr2
                    if self.serialport2.is_open: self.serialport2.close()
                    self.serialport2.open()
                    self.serialport2.reset_input_buffer()
            else:
                # 网络模式
                # 解析地址1
                ip1, p1 = addr1.split(':')
                self.socket1 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                self.socket1.settimeout(1.0) # 设置超时
                self.socket1.connect((ip1, int(p1)))
                
                if not self.single_channel_mode:
                    # 解析地址2
                    ip2, p2 = addr2.split(':')
                    self.socket2 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    self.socket2.settimeout(1.0)
                    self.socket2.connect((ip2, int(p2)))

        except Exception as e:
            self.log_bus.log("ERROR", f"Start failed: connection: {e}")
            QMessageBox.critical(self, "Error", f"Connection Failed: {e}")
            return

        # 打开数据文件
        if not self.open_data_file():
            self.log_bus.log("ERROR", "Start failed: cannot open data file")
            QMessageBox.critical(self, "Error", "Cannot Open Data File, Monitoring Cannot Start!")
            return
        
        self.run_stat = True
        # 更新菜单状态
        self.start_action.setEnabled(False)
        self.stop_action.setEnabled(True)
        
        # 运行时禁止切换模式
        self.single_mode_action.setEnabled(False)
        # 运行时禁止调整更新间隔
        self.update_interval_action.setEnabled(False)
        # 运行时禁止设置单位和阈值
        self.set_units_action.setEnabled(False)
        self.set_limit_action.setEnabled(False)
        # 禁用模式切换
        self.mode_serial_action.setEnabled(False)
        self.mode_network_action.setEnabled(False)
        # 运行时禁止修改波特率
        self.set_baudrate_action.setEnabled(False) 

        # 初始化计时和数据
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
        
        # 启动监控逻辑
        self.monitor1.start()
        self.monitor2.start()
        # 更新参数以确保队列长度正确
        self.monitor1.update_params(self.update_interval)
        self.monitor2.update_params(self.update_interval)

        # 重新启动数据更新定时器
        self.timer.start(self.update_interval)                        
        print("Monitoring Started")
        self.log_bus.log("INFO", f"Monitoring started ({source})")
    
    def stop_monitoring(self, source="manual"):
        """停止监控
        source: 'manual' / 'daq_master' / 'gddaq' / 'daq_disabled' — 标识触发来源"""

        self.monitor1.stop()
        self.monitor2.stop()
        self.status_label1.set_status("STOP")
        self.status_label2.set_status("STOP")

        self.run_stat = False
        self.timer.stop()  # 先停定时器，防止继续调用 send/recv

        # 重置 recv 去抖计数器 (Reset recv debounce counters)
        self.recv_none_count_1 = 0
        self.recv_none_count_2 = 0
        self.recv_warned_1 = False
        self.recv_warned_2 = False

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

        # 关闭网络连接
        if self.socket1:
            try:
                self.socket1.close()
                print("Socket 1 Closed")
            except: pass
            self.socket1 = None
            
        if self.socket2:
            try:
                self.socket2.close()
                print("Socket 2 Closed")
            except: pass
            self.socket2 = None        

        # 关闭数据文件
        self.close_data_file()

        # 更新菜单状态
        self.start_action.setEnabled(True)
        self.stop_action.setEnabled(False)   

        # 停止后允许切换模式
        self.single_mode_action.setEnabled(True)
        # 停止后允许调整更新间隔
        self.update_interval_action.setEnabled(True)
        # 停止后允许设置单位和阈值
        self.set_units_action.setEnabled(True)
        self.set_limit_action.setEnabled(True)
        # 
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        # 启用模式切换
        self.mode_serial_action.setEnabled(True)
        self.mode_network_action.setEnabled(True)
        # 停止后允许修改波特率
        self.set_baudrate_action.setEnabled(True)

        # 停止脉冲提醒定时器
        if self.pulse_reminder_timer.isActive():
            self.pulse_reminder_timer.stop()

        # 停止 ZERO 报警音
        if self.alarm_active:
            self.alarm_timer.stop()
            self.alarm_active = False
        
        # 自动更新文件名为下一个序号
        current_filename = self.filename_input.text().strip()
        if not current_filename:
            current_filename = self.filename
        
        next_filename = self.generate_next_filename(current_filename)
        self.filename_input.setText(next_filename)
        self.filename = next_filename
        
        print(f"Monitoring Stopped")
        self.log_bus.log("INFO", f"Monitoring stopped ({source})")

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
            self.log_bus.log("WARNING", "Cannot open folder: directory does not exist")
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
            self.log_bus.log("WARNING", f"Cannot open folder: {e}")

    def set_channel_units(self):
        """设置两个通道的单位"""
        if self.run_stat:
            QMessageBox.warning(self, "Warning", "Cannot change units while monitoring is running!")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Set Channel Units")
        dialog.setFixedSize(300, 150)
        
        layout = QGridLayout()
        
        # 通道 1 设置
        layout.addWidget(QLabel("Channel 1 Unit:"), 0, 0)
        combo1 = QComboBox()
        combo1.addItems(["mA", "μA", "nA"])
        combo1.setCurrentText(self.unit_ch1)
        layout.addWidget(combo1, 0, 1)
        
        # 通道 2 设置
        layout.addWidget(QLabel("Channel 2 Unit:"), 1, 0)
        combo2 = QComboBox()
        combo2.addItems(["mA", "μA", "nA"])
        combo2.setCurrentText(self.unit_ch2)
        layout.addWidget(combo2, 1, 1)
        
        # 按钮
        btn_box = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        btn_box.addWidget(ok_btn)
        btn_box.addWidget(cancel_btn)
        layout.addLayout(btn_box, 2, 0, 1, 2)
        
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            self.unit_ch1 = combo1.currentText()
            self.unit_ch2 = combo2.currentText()
            
            # 更新 UI 上的标签
            self.current1_label.setText(f"Channel 1 Current: --- {self.unit_ch1}")
            if not self.single_channel_mode:
                self.current2_label.setText(f"Channel 2 Current: --- {self.unit_ch2}")
            
            # 更新绘图轴标签
            self.ax.set_ylabel(f'Channel 1 Current ({self.unit_ch1})', color='#1f77b4', fontweight='bold')
            self.ax2.set_ylabel(f'Channel 2 Current ({self.unit_ch2})', color='#ff7f0e', fontweight='bold')
            self.canvas.draw()
            
            print(f"Units set to: Ch1={self.unit_ch1}, Ch2={self.unit_ch2}")
            self.log_bus.log("INFO", f"Channel units set: Ch1={self.unit_ch1}, Ch2={self.unit_ch2}")

    def set_update_interval(self):
        """设置更新间隔对话框"""
        # 获取当前更新间隔
        current_interval = self.update_interval
        
        # 显示输入对话框
        interval, ok = QInputDialog.getInt(
            self, 
            'Set Update Interval', 
            'Update Interval (milliseconds):\n\nAvailable range: 1-5000ms\nCurrent value: {}ms'.format(current_interval),
            current_interval,  # 默认值
            1,               # 最小值
            5000,             # 最大值
            1                # 步长
        )
        
        if ok and interval != current_interval:
            # 更新间隔值
            self.update_interval = interval
            self.monitor1.update_params(self.update_interval)
            self.monitor2.update_params(self.update_interval)

            # 如果定时器正在运行，重新启动定时器
            if self.timer.isActive():
                self.timer.stop()
                self.timer.start(self.update_interval)
                
                # 显示确认消息
                QMessageBox.information(
                    self, 
                    'Update Interval Changed', 
                    f'Update interval has been changed to {interval}ms.\n\n'
                    f'This will apply to the next Run.'
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
            self.log_bus.log("INFO", f"Update interval set to {interval}ms")

    # 设置波特率的逻辑函数
    def set_baud_rate(self):
        """设置波特率对话框"""
        # 常用波特率列表
        rates = ["4800", "9600", "19200", "38400", "57600", "115200"]
        
        # 获取当前索引
        current_str = str(self.baud_rate)
        try:
            index = rates.index(current_str)
        except ValueError:
            index = 1 # 默认 9600
            
        # 弹出选择框
        item, ok = QInputDialog.getItem(
            self, 
            "Set Baud Rate", 
            "Select Baud Rate:", 
            rates, 
            index, 
            False # False 表示不可编辑，只能从列表中选
        )
        
        if ok and item:
            self.baud_rate = int(item)
            # 更新串口对象的配置
            self.serialport1.baudrate = self.baud_rate
            self.serialport2.baudrate = self.baud_rate
            
            QMessageBox.information(
                self, 
                "Baud Rate Set", 
                f"Baud Rate has been set to {self.baud_rate}.\n"
                f"This will take effect when monitoring starts."
            )
            print(f"Baud Rate changed to: {self.baud_rate}")
            self.log_bus.log("INFO", f"Baud rate set to {self.baud_rate}")

    def open_monitor_settings(self):
        """打开状态监控设置对话框"""
        dialog = MonitorSettingsDialog(
            self.monitor1, self.monitor2, 
            self.unit_ch1, self.unit_ch2, 
            self.alarm_enabled,
            self
        )
        if dialog.exec_():
            # 读取报警音开关
            self.alarm_enabled = dialog.alarm_cb.isChecked()
            # 如果关闭了报警音，立即停止正在响的报警
            if not self.alarm_enabled and self.alarm_active:
                self.alarm_timer.stop()
                self.alarm_active = False
            # 设置应用后，立即更新 Monitor 的采样率参数 (以防窗口时间改变)
            self.monitor1.update_params(self.update_interval)
            self.monitor2.update_params(self.update_interval)
            print(f"Monitor settings updated. Alarm sound: {'on' if self.alarm_enabled else 'off'}.")
            self.log_bus.log("INFO", "Status monitor settings applied")

    def set_current_threshold(self):
        """设置电流过滤阈值对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Set Current Threshold (Filter)")
        dialog.setFixedSize(350, 180)
        
        layout = QGridLayout()
        
        # 通道 1 设置
        layout.addWidget(QLabel("Channel 1 Max Limit:"), 0, 0)
        
        # 数值输入框
        spin1 = QtWidgets.QDoubleSpinBox()
        spin1.setRange(0, 999999) # 设置范围
        spin1.setDecimals(4)      # 设置小数位
        spin1.setValue(self.limit_ch1_ma) # 默认显示当前的mA值
        layout.addWidget(spin1, 0, 1)
        
        # 单位选择框
        combo1 = QComboBox()
        combo1.addItems(["mA", "μA", "nA"])
        combo1.setCurrentText("mA") # 默认显示单位为mA，因为spinbox里填的是mA值
        layout.addWidget(combo1, 0, 2)
        
        # 通道 2 设置
        layout.addWidget(QLabel("Channel 2 Max Limit:"), 1, 0)
        
        # 数值输入框
        spin2 = QtWidgets.QDoubleSpinBox()
        spin2.setRange(0, 999999)
        spin2.setDecimals(4)
        spin2.setValue(self.limit_ch2_ma)
        layout.addWidget(spin2, 1, 1)
        
        # 单位选择框
        combo2 = QComboBox()
        combo2.addItems(["mA", "μA", "nA"])
        combo2.setCurrentText("mA")
        layout.addWidget(combo2, 1, 2)
        
        # 说明标签
        note_label = QLabel("Note: Signals exceeding this value will be ignored.")
        note_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addWidget(note_label, 2, 0, 1, 3)

        # 按钮区域
        btn_box = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        btn_box.addWidget(ok_btn)
        btn_box.addWidget(cancel_btn)
        layout.addLayout(btn_box, 3, 0, 1, 3)
        
        dialog.setLayout(layout)
        
        # 如果用户点击了OK
        if dialog.exec_() == QDialog.Accepted:
            # 计算并保存通道1阈值 (转换为 mA)
            val1 = spin1.value()
            unit1 = combo1.currentText()
            self.limit_ch1_ma = val1 * self.unit_factors[unit1]
            
            # 计算并保存通道2阈值 (转换为 mA)
            val2 = spin2.value()
            unit2 = combo2.currentText()
            self.limit_ch2_ma = val2 * self.unit_factors[unit2]
            
            print(f"Thresholds Updated: Ch1={self.limit_ch1_ma} mA, Ch2={self.limit_ch2_ma} mA")
            self.log_bus.log("INFO", f"Current filter thresholds set: Ch1={self.limit_ch1_ma} mA, Ch2={self.limit_ch2_ma} mA")

            # 状态栏反馈
            QMessageBox.information(self, "Updated",
                                  f"Filter Thresholds Set:\n"
                                  f"Ch1: {self.limit_ch1_ma} mA\n"
                                  f"Ch2: {self.limit_ch2_ma} mA")

    def on_alarm_tick(self):
        """ZERO 报警音定时器回调：每3秒响一次"""
        QApplication.beep()

    def update_alarm_state(self, state1, state2):
        """根据通道状态更新报警音：任一通道为 ZERO 时响"""
        if not self.alarm_enabled:
            if self.alarm_active:
                self.alarm_timer.stop()
                self.alarm_active = False
            return

        any_zero = (state1 == "ZERO") or (state2 == "ZERO")

        if any_zero and not self.alarm_active:
            # 开始报警
            self.alarm_active = True
            QApplication.beep()  # 立即响一声
            self.alarm_timer.start()
        elif not any_zero and self.alarm_active:
            # 停止报警
            self.alarm_timer.stop()
            self.alarm_active = False

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
        if getattr(sys, 'frozen', False):
            # 如果是打包后的环境
            base_path = sys._MEIPASS
        else:
            # 如果是正常运行环境
            base_path = os.path.dirname(os.path.abspath(__file__))

        icon_path = os.path.join(base_path, "logo.png")

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
        # 为通道 1 创建注释 (绑定到 ax，蓝色背景)
        self.hover_annotation1 = self.ax.annotate(
            '', 
            xy=(0, 0), 
            xytext=(20, 20), 
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.5", fc="lightblue", alpha=0.8),
            arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0"),
            fontsize=10,
            visible=False
        )
        
        # 为通道 2 创建注释 (绑定到 ax2，橙色背景)
        self.hover_annotation2 = self.ax2.annotate(
            '', 
            xy=(0, 0), 
            xytext=(20, 20), 
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.5", fc="#ffcc99", alpha=0.8),
            arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0"),
            fontsize=10,
            visible=False
        )
        
        # 连接鼠标移动事件
        self.canvas.mpl_connect('motion_notify_event', self.on_hover)

    def on_hover(self, event):
        """鼠标悬停事件处理"""
        # 检查鼠标是否在任意一个轴内 (ax 或 ax2)
        if event.inaxes not in [self.ax, self.ax2]:
            self.hover_annotation1.set_visible(False)
            self.hover_annotation2.set_visible(False)
            self.canvas.draw_idle()
            return
        
        # 检查是否有有效数据
        if not hasattr(self, 'time_data') or len(self.time_data) == 0:
            self.hover_annotation1.set_visible(False)
            self.hover_annotation2.set_visible(False)
            self.canvas.draw_idle()
            return
        
        # 找到最接近的数据点 (传入 event 对象以处理坐标转换)
        closest_point = self.find_closest_point(event)
        
        if closest_point:
            channel, index, x_val, y_val, time_val = closest_point
            
            # 计算运行时间
            if self.start_time and time_val > 0:
                runtime = time_val - self.start_time
                time_str = f"{runtime:.2f}s"
            else:
                time_str = "N/A"
            
            # 获取当前通道的单位
            unit = self.unit_ch1 if channel == 1 else self.unit_ch2
            
            # 创建显示文本
            hover_text = f"Channel {channel}\nTime: {time_str}\nCurrent: {y_val:.4f} {unit}"
            
            # 根据通道显示对应的注释框，并隐藏另一个
            if channel == 1:
                self.hover_annotation1.xy = (x_val, y_val)
                self.hover_annotation1.set_text(hover_text)
                self.hover_annotation1.set_visible(True)
                self.hover_annotation2.set_visible(False)
            else:
                self.hover_annotation2.xy = (x_val, y_val)
                self.hover_annotation2.set_text(hover_text)
                self.hover_annotation2.set_visible(True)
                self.hover_annotation1.set_visible(False)
        else:
            self.hover_annotation1.set_visible(False)
            self.hover_annotation2.set_visible(False)
        
        self.canvas.draw_idle()

    def find_closest_point(self, event):
        """找到最接近鼠标位置的数据点"""
        min_distance = float('inf')
        closest_point = None
        
        # 获取鼠标在两个轴坐标系下的数据坐标
        # 注意：event.x 和 event.y 是屏幕像素坐标，我们需要将其分别转换回两个轴的数据坐标
        try:
            x1, y1 = self.ax.transData.inverted().transform((event.x, event.y))
            x2, y2 = self.ax2.transData.inverted().transform((event.x, event.y))
        except Exception:
            return None
        
        # 检查通道 1 的数据点 (使用 ax 的坐标系)
        xlim1 = self.ax.get_xlim()
        ylim1 = self.ax.get_ylim()
        x_range1 = xlim1[1] - xlim1[0]
        y_range1 = ylim1[1] - ylim1[0]
        
        if x_range1 > 0 and y_range1 > 0:
            for i in range(len(self.x_data)):
                x_data_point = self.x_data[i]
                y_data_point = self.y_data1[i]
                
                # 归一化距离计算 (使用 x1, y1)
                dx = (x1 - x_data_point) / x_range1
                dy = (y1 - y_data_point) / y_range1
                distance = (dx**2 + dy**2)**0.5
                
                if distance < min_distance and distance < 0.05:  # 阈值
                    min_distance = distance
                    time_val = self.time_data[i] if hasattr(self, 'time_data') and i < len(self.time_data) else 0
                    closest_point = (1, i, x_data_point, y_data_point, time_val)
        
        # 检查通道 2 的数据点 (使用 ax2 的坐标系)
        if not self.single_channel_mode:
            xlim2 = self.ax2.get_xlim()
            ylim2 = self.ax2.get_ylim()
            x_range2 = xlim2[1] - xlim2[0]
            y_range2 = ylim2[1] - ylim2[0]
            
            if x_range2 > 0 and y_range2 > 0:
                for i in range(len(self.x_data)):
                    x_data_point = self.x_data[i]
                    y_data_point = self.y_data2[i]
                    
                    # 归一化距离计算 (使用 x2, y2)
                    dx = (x2 - x_data_point) / x_range2
                    dy = (y2 - y_data_point) / y_range2
                    distance = (dx**2 + dy**2)**0.5
                    
                    if distance < min_distance and distance < 0.05:
                        min_distance = distance
                        time_val = self.time_data[i] if hasattr(self, 'time_data') and i < len(self.time_data) else 0
                        closest_point = (2, i, x_data_point, y_data_point, time_val)
        
        return closest_point

    def show_about(self):
        """显示关于对话框 (非模态, 单实例复用)

        若 About 对话框已存在 (即此前打开过、未被销毁)，则前置显示已存在的窗口，
        不重复创建。否则新建对话框并存入 self._about_dialog，以 show() 非模态弹出，
        允许用户在保留对话框打开的同时操作主程序。
        """
        # 单实例复用: 已打开则前置显示 (single-instance reuse: raise existing window)
        if self._about_dialog is not None:
            self._about_dialog.show()
            self._about_dialog.raise_()
            self._about_dialog.activateWindow()
            return

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
        
        # 读取外部HTML文件
        about_file = resource_path('about.html')
        
        try:
            with open(about_file, 'r', encoding='utf-8') as f:
                about_text = f.read()
        except FileNotFoundError:
            about_text = """
            <div style="font-family: Arial, sans-serif; line-height: 1.6;">
            <center>
            <h2 style="color: #2E86AB;">Real-Time Current Monitoring System</h2>
            <p><b>Error:</b> about.html file not found</p>
            <p>Please ensure the about.html file is in the same directory as the main program.</p>
            </center>
            </div>
            """
        except Exception as e:
            about_text = f"""
            <div style="font-family: Arial, sans-serif; line-height: 1.6;">
            <center>
            <h2 style="color: #2E86AB;">Real-Time Current Monitoring System</h2>
            <p><b>Error:</b> Failed to load about.html</p>
            <p>Error details: {str(e)}</p>
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

        # 保存单实例引用, 非模态弹出 (store singleton ref, modeless show)
        # QDialog 默认 windowModality 为 Qt.NonModal, show() 立即返回不阻塞主事件循环
        self._about_dialog = about_dialog
        about_dialog.show()

    def show_tutorial(self):
        """显示教程对话框 (非模态, 单实例复用)

        若 Tutorial 对话框已存在 (即此前打开过、未被销毁)，则前置显示已存在的窗口，
        不重复创建。否则新建对话框并存入 self._tutorial_dialog，以 show() 非模态弹出，
        允许用户在保留对话框打开的同时操作主程序。
        """
        # 单实例复用: 已打开则前置显示 (single-instance reuse: raise existing window)
        if self._tutorial_dialog is not None:
            self._tutorial_dialog.show()
            self._tutorial_dialog.raise_()
            self._tutorial_dialog.activateWindow()
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Tutorial")
        dialog.setFixedSize(600, 500)
        
        layout = QVBoxLayout()
        text_browser = QTextBrowser()
        text_browser.setOpenExternalLinks(True)
        
        # 获取 HTML 文件的绝对路径
        html_file = resource_path('tutorial.html')
        
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

        # 保存单实例引用, 非模态弹出 (store singleton ref, modeless show)
        # QDialog 默认 windowModality 为 Qt.NonModal, show() 立即返回不阻塞主事件循环
        self._tutorial_dialog = dialog
        dialog.show()

    def closeEvent(self, event):
        """关闭事件，确保资源被正确释放"""
        self.stop_monitoring()
        self.log_bus.log("INFO", "Application closed")
        event.accept()

    def update_data(self):
        """定时更新数据 - 支持单/双通道、多单位转换及双轴绘图"""
        if not self.run_stat:
            return
        
        try:
            # 1. 发送请求
            self.send_data(1)
            if not self.single_channel_mode:
                self.send_data(2)
            
            # 2. 接收数据 (这里获取的是对应单位的原始数值)
            current1 = self.recv_data(1)
            
            if not self.single_channel_mode:
                current2 = self.recv_data(2)
            else:
                current2 = 0.0
            
            # 检查数据有效性 + recv 连续无响应去抖
            # (Debounce recv timeouts: ~5s no data -> WARNING; resumed -> INFO)
            recv_warn_threshold = max(50, int(5000 / self.update_interval))

            if current1 is None:
                self.recv_none_count_1 += 1
                if self.recv_none_count_1 >= recv_warn_threshold and not self.recv_warned_1:
                    self.log_bus.log("WARNING",
                        "Channel 1: no data for ~5s (connection issue?)")
                    self.recv_warned_1 = True
                return
            else:
                # 有效数据恢复 (Valid data resumed)
                if self.recv_warned_1:
                    self.log_bus.log("INFO", "Channel 1: data resumed")
                self.recv_none_count_1 = 0
                self.recv_warned_1 = False

            if not self.single_channel_mode and current2 is None:
                self.recv_none_count_2 += 1
                if self.recv_none_count_2 >= recv_warn_threshold and not self.recv_warned_2:
                    self.log_bus.log("WARNING",
                        "Channel 2: no data for ~5s (connection issue?)")
                    self.recv_warned_2 = True
                return
            elif not self.single_channel_mode:
                if self.recv_warned_2:
                    self.log_bus.log("INFO", "Channel 2: data resumed")
                self.recv_none_count_2 = 0
                self.recv_warned_2 = False
            
            # 3. 数据转换：将原始读数转换为 mA，用于积分计算和文件保存
            factor1 = self.unit_factors[self.unit_ch1]
            current1_ma = current1 * factor1
            
            if not self.single_channel_mode:
                factor2 = self.unit_factors[self.unit_ch2]
                current2_ma = current2 * factor2
            else:
                current2_ma = 0.0
            
            # 简单过滤：如果转换后的 mA 值大得离谱(>1000mA)，可能是干扰
            if current1_ma > self.limit_ch1_ma or current2_ma > self.limit_ch2_ma:
                print(f"Invalid Current (mA converted): Ch1={current1_ma}, Ch2={current2_ma}, Skipping")
                return
            
            # 获取当前时间
            now = self.get_time()
            local_time = time.localtime(now)
            utc_time = time.gmtime(now)
            local_time_str = time.strftime("%Y-%m-%d %H:%M:%S (UTC%z)", local_time)
            utc_time_str = time.strftime("%Y-%m-%dT%H:%M:%SZ", utc_time)
            
            # 4. 积分计算 (必须使用 mA 值，确保积分单位是 mC)
            if self.last_time is not None:
                delta_t = now - self.last_time
                
                if delta_t > 0:
                    # 通道1积分
                    if self.last_current1 is not None:
                        # self.last_current1 存储的是上一次的 mA 值
                        avg_current1 = Decimal(str((self.last_current1 + current1_ma) / 2.0))
                        self.column_int1 += avg_current1 * Decimal(str(delta_t))
                    else:
                        self.column_int1 += Decimal(str(current1_ma)) * Decimal(str(delta_t))
                    
                    # 通道2积分
                    if not self.single_channel_mode:
                        if self.last_current2 is not None:
                            avg_current2 = Decimal(str((self.last_current2 + current2_ma) / 2.0))
                            self.column_int2 += avg_current2 * Decimal(str(delta_t))
                        else:
                            self.column_int2 += Decimal(str(current2_ma)) * Decimal(str(delta_t))
            
            # 保存当前的 mA 值用于下次积分计算
            self.last_current1 = current1_ma
            self.last_current2 = current2_ma
            self.last_time = now
            
            # 5. 更新 UI 显示 (显示原始数值 + 当前单位)
            runtime = now - self.start_time
            h = int(runtime // 3600)
            m = int((runtime - h * 3600) // 60)
            s = runtime - h * 3600 - m * 60
            
            integral1_float = float(self.column_int1)
            integral2_float = float(self.column_int2)
            
            self.current1_label.setText(f"Channel 1 Current: {current1:.4f} {self.unit_ch1}")
            
            if self.single_channel_mode:
                self.current2_label.setText("Channel 2 Current: --- (Disabled)")
            else:
                self.current2_label.setText(f"Channel 2 Current: {current2:.4f} {self.unit_ch2}")

            self.runtime_label.setText(f"Run Time: {h:02d} Hours {m:02d} Minutes {s:05.2f} Seconds")
            self.integral1_label.setText(f"Channel 1 Integral: {integral1_float:.4e} mC")
            self.integral2_label.setText(f"Channel 2 Integral: {integral2_float:.4e} mC")
            self.timestamp_label.setText(f"Last Update Time (Local): {local_time_str}")
            self.utc_timestamp_label.setText(f"UTC Timestamp: {utc_time_str}")
            
            # 6. 更新绘图数据 (使用原始数值，因为是双纵轴，各自显示各自的单位数值)
            self.y_data1 = np.roll(self.y_data1, -1)
            self.y_data1[-1] = current1
            
            self.y_data2 = np.roll(self.y_data2, -1)
            self.y_data2[-1] = current2
            
            self.time_data = np.roll(self.time_data, -1)
            self.time_data[-1] = now
            
            # 更新曲线
            self.line1.set_ydata(self.y_data1)
            self.line2.set_ydata(self.y_data2)
            
            # 控制可见性
            self.line2.set_visible(not self.single_channel_mode)
            self.ax2.set_visible(not self.single_channel_mode)
            
            # 更新X轴标签
            x_labels = []
            for i in range(len(self.x_data)):
                time_sec = (self.time_data[i] - self.start_time) if self.time_data[i] > 0 else 0
                x_labels.append(f"{time_sec:.1f}")
            
            tick_indices = range(0, self.data_points, max(1, self.data_points//10))
            self.ax.set_xticks(tick_indices)
            self.ax.set_xticklabels([x_labels[i] for i in tick_indices])
            
            # 7. 自动调整 Y 轴范围 (双轴独立调整)
            # 调整左轴 (Channel 1)
            min_y1 = np.min(self.y_data1)
            max_y1 = np.max(self.y_data1)
            range_y1 = max_y1 - min_y1

            if range_y1 == 0:
                # 如果是直线（数值不变），上下各留 25% 的绝对值空间，或者默认 0.1
                margin1 = max(0.1, abs(max_y1) * 0.25)
            else:
                # 如果有波动，上下各留波动幅度的 25%
                margin1 = range_y1 * 0.25

            self.ax.set_ylim(min_y1 - margin1, max_y1 + margin1)
            
            # 调整右轴 (Channel 2)
            if not self.single_channel_mode:
                min_y2 = np.min(self.y_data2)
                max_y2 = np.max(self.y_data2)
                range_y2 = max_y2 - min_y2

                if range_y2 == 0:
                    margin2 = max(0.1, abs(max_y2) * 0.25)
                else:
                    margin2 = range_y2 * 0.25

                self.ax2.set_ylim(min_y2 - margin2, max_y2 + margin2)    

            # 重绘
            self.canvas.draw()
            
            # 8. 插入监控逻辑
            if self.run_stat:
                # 处理通道 1
                # 注意：这里传入的值应该是对应单位的值。
                # 如果 current1 是 mA，而设置里也是 mA，直接传。
                state1 = self.monitor1.process(current1)
                self.status_label1.set_status(state1)
                
                # 处理通道 2
                if not self.single_channel_mode:
                    state2 = self.monitor2.process(current2)
                    self.status_label2.set_status(state2)
                else:
                    state2 = "STOP"
                    self.status_label2.set_status("STOP", "OFF")

                # 更新 ZERO 报警音
                self.update_alarm_state(state1, state2)

            # 9. 写入文件 (传入转换后的 mA 值，write_data_row 内部需使用 .8e 格式)
            self.write_data_row(now, runtime, current1_ma, current2_ma, integral1_float, integral2_float)
            
            # 打印日志
            print(f"Ch1: {current1:.4f} {self.unit_ch1}, Ch2: {current2:.4f} {self.unit_ch2}, "
                  f"Int1: {integral1_float:.4f} mC, Int2: {integral2_float:.4f} mC")
            
        except Exception as e:
            print(f"Error Updating Data: {e}")
            import traceback
            traceback.print_exc()

def main():
    app = QApplication(sys.argv)

    # 启用文件锁
    instance_lock = SingleInstanceLock()    # 创建锁对象
        # 尝试获取锁，如果失败说明已有程序在运行
    if not instance_lock.acquire_lock():
        QMessageBox.warning(None, "Warning", "Program is already running!")
        sys.exit(1)

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
