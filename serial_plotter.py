"""
Serial Plotter

Author: Nickchung
Date: July 16, 2024

Description:
This application is a serial port data plotter with CSV logging capabilities. It allows users to visualize 
and record data received from a serial port in real-time.

Features:
1. Real-time plotting of serial data
2. Customizable plot lines with checkboxes to show/hide data series
3. Adjustable maximum number of data points displayed
4. CSV logging of received data
5. Selectable serial port and baud rate
6. Ability to pause/resume data collection
7. Option to clear the plot
8. Settings persistence across sessions

Usage:
1. Select the appropriate serial port and baud rate
2. Click 'Connect' to establish a connection
3. Use the 'Run/Stop' button to control data collection
4. Adjust plot settings and CSV logging options as needed
5. Use checkboxes to show/hide specific data series

Note: Ensure that the serial device is sending comma-separated numeric values for proper functioning.

Requirements:
- Python 3.x
- PyQt5
- pyqtgraph
- pyserial

How to Run:
1. Ensure all required libraries are installed:
   pip install PyQt5 pyqtgraph pyserial
2. Run the script using Python:
   python serial_plotter.py

Creating an Executable:
To create a standalone executable file:
1. Install PyInstaller:
   pip install pyinstaller
2. Use PyInstaller to create the executable:
   pyinstaller --onefile --windowed serial_plotter.py
3. The executable will be created in the 'dist' folder

Note: When distributing the executable, ensure that the target machine has the necessary drivers 
for serial communication.

For any issues or feature requests, please contact the author.
"""

import sys
import os
import serial
import serial.tools.list_ports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QComboBox, QPushButton, QCheckBox, QLineEdit, QGroupBox, QLabel, QFileDialog)
from PyQt5.QtCore import QTimer, Qt, QSettings
from PyQt5.QtGui import QIntValidator
import pyqtgraph as pg
import numpy as np
import time
import csv

from PyQt5.QtCore import pyqtSignal

class CustomComboBox(QComboBox):
    popupAboutToBeShown = pyqtSignal()

    def showPopup(self):
        self.popupAboutToBeShown.emit()
        super().showPopup()

class SerialPlotter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Serial Plotter")
        self.setGeometry(100, 100, 1000, 600)

        self.serial = None
        self.data = []
        self.lines = []
        self.checkboxes = []
        self.settings = QSettings("MyCompany", "SerialPlotter")
        self.is_running = False
        self.max_points = 100
        self.legend = None
        self.total_data_count = 0
        self.csv_file = None
        self.csv_writer = None

        self.init_ui()
        self.load_settings()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # 控制面板
        control_panel = QGroupBox("Controls")
        control_layout = QVBoxLayout()
        control_panel.setLayout(control_layout)
        control_panel.setFixedWidth(200)
        main_layout.addWidget(control_panel)

        self.port_combo = CustomComboBox()
        self.port_combo.setPlaceholderText("Select Port")
        self.port_combo.popupAboutToBeShown.connect(self.update_ports)
        control_layout.addWidget(self.port_combo)

        self.baud_combo = QComboBox()
        self.baud_combo.addItems(['9600', '19200', '38400', '57600', '115200'])
        control_layout.addWidget(self.baud_combo)

        self.connect_button = QPushButton("Connect")
        self.connect_button.clicked.connect(self.toggle_connection)
        control_layout.addWidget(self.connect_button)

        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_plot)
        control_layout.addWidget(self.clear_button)

        self.run_stop_button = QPushButton("Run")
        self.run_stop_button.clicked.connect(self.toggle_run_stop)
        self.run_stop_button.setEnabled(False)
        control_layout.addWidget(self.run_stop_button)

        # 數據點數量限制
        data_points_layout = QHBoxLayout()
        data_points_layout.addWidget(QLabel("Max data points:"))
        self.max_points_edit = QLineEdit(str(self.max_points))
        self.max_points_edit.setValidator(QIntValidator(1, 10000))
        self.max_points_edit.returnPressed.connect(self.update_max_points)
        data_points_layout.addWidget(self.max_points_edit)
        control_layout.addLayout(data_points_layout)

        # CSV 檔案設置
        csv_group = QGroupBox("CSV Settings")
        csv_layout = QVBoxLayout()
        csv_group.setLayout(csv_layout)

        self.csv_filename_edit = QLineEdit("test")
        csv_layout.addWidget(QLabel("File Name:"))
        csv_layout.addWidget(self.csv_filename_edit)

        self.csv_folder_button = QPushButton("Select Folder")
        self.csv_folder_button.clicked.connect(self.select_csv_folder)
        csv_layout.addWidget(self.csv_folder_button)

        self.csv_folder_label = QLabel("Selected Folder: Not selected")
        self.csv_folder_label.setWordWrap(True)
        csv_layout.addWidget(self.csv_folder_label)

        control_layout.addWidget(csv_group)

        # Checkbox面板
        self.checkbox_layout = QVBoxLayout()
        control_layout.addLayout(self.checkbox_layout)

        control_layout.addStretch(1)

        # 錯誤訊息標籤
        self.error_label = QLabel()
        self.error_label.setWordWrap(True)
        control_layout.addWidget(self.error_label)

        # Restore Default 按鈕
        self.restore_button = QPushButton("Restore Default")
        self.restore_button.clicked.connect(self.restore_default)
        control_layout.addWidget(self.restore_button)

        # 添加設計者和支持者信息
        designer_label = QLabel("Designed by Nickchung")
        designer_label.setAlignment(Qt.AlignCenter)
        control_layout.addWidget(designer_label)

        powered_label = QLabel("Powered by NCKU ISP")
        powered_label.setAlignment(Qt.AlignCenter)
        control_layout.addWidget(powered_label)

        # 圖表
        plot_layout = QVBoxLayout()
        main_layout.addLayout(plot_layout)

        self.plot_widget = pg.PlotWidget(background='w')
        self.plot_widget.getAxis('bottom').setPen('k')
        self.plot_widget.getAxis('left').setPen('k')
        plot_layout.addWidget(self.plot_widget)

        # 創建圖例
        self.legend = pg.LegendItem(size=(100,60), offset=(70,20))
        self.legend.setParentItem(self.plot_widget.graphicsItem())
        self.legend.anchor(itemPos=(1, 0), parentPos=(1, 0), offset=(-10, 10))

        # 定時器
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_plot)
        self.timer.start(50)  # 50ms

    def update_ports(self):
        current_port = self.port_combo.currentText()
        self.port_combo.clear()
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo.addItems(ports)
        if current_port in ports:
            self.port_combo.setCurrentText(current_port)

    def toggle_connection(self):
        if self.serial is None or not self.serial.is_open:
            self.connect_serial()
        else:
            self.disconnect_serial()

    def connect_serial(self):
        port = self.port_combo.currentText()
        baud = int(self.baud_combo.currentText())
        retry_count = 3
        for attempt in range(retry_count):
            try:
                self.serial = serial.Serial(port, baud, timeout=1)
                self.connect_button.setText("Disconnect")
                self.clear_plot()
                self.error_label.setText("")
                self.run_stop_button.setEnabled(True)
                self.is_running = True
                self.run_stop_button.setText("Stop")
                self.legend.clear()
                self.open_csv_file()
                return
            except serial.SerialException as e:
                if attempt < retry_count - 1:
                    self.error_label.setText(f"Connection attempt {attempt + 1} failed. Retrying...")
                    time.sleep(1)
                else:
                    self.error_label.setText(f"Error: {str(e)}\nTry closing other programs using this port or restart your computer.")
            except Exception as e:
                self.error_label.setText(f"Unexpected error: {str(e)}")
                break

    def disconnect_serial(self):
        if self.serial:
            try:
                self.serial.close()
            except Exception as e:
                self.error_label.setText(f"Error closing port: {str(e)}")
            finally:
                self.serial = None
                self.connect_button.setText("Connect")
                self.run_stop_button.setEnabled(False)
                self.is_running = False
                self.run_stop_button.setText("Run")
                self.close_csv_file()

    def clear_plot(self):
        self.data = []
        for line in self.lines:
            self.plot_widget.removeItem(line)
        self.lines = []
        self.legend.clear()
        self.total_data_count = 0

    def toggle_run_stop(self):
        if self.serial and self.serial.is_open:
            self.is_running = not self.is_running
            self.run_stop_button.setText("Stop" if self.is_running else "Run")

    def update_max_points(self):
        try:
            self.max_points = int(self.max_points_edit.text())
        except ValueError:
            pass

    def update_plot(self):
        if self.serial and self.serial.is_open and self.is_running:
            try:
                new_data = False
                values = []
                while self.serial.in_waiting:
                    line = self.serial.readline().decode().strip()
                    try:
                        values = [float(x) for x in line.split(',')]
                        self.data.append(values)
                        self.total_data_count += 1
                        new_data = True
                        if self.csv_writer:
                            self.csv_writer.writerow(values)
                    except ValueError:
                        # 忽略非數字的輸入
                        pass

                self.data = self.data[-self.max_points:]

                if new_data and values:
                    while len(self.lines) < len(values):
                        color = pg.intColor(len(self.lines), hues=len(values), values=1, maxValue=255)
                        new_line = self.plot_widget.plot(pen=pg.mkPen(color=color, width=3))
                        self.lines.append(new_line)
                        if len(self.checkboxes) < len(values):
                            self.add_checkbox(f"Data {len(self.lines)}", color)

                    self.update_plot_data()

            except Exception as e:
                self.error_label.setText(f"Error: {str(e)}")

    def update_plot_data(self):
        self.legend.clear()

        for i, line in enumerate(self.lines):
            if i < len(self.checkboxes):
                if self.checkboxes[i].isChecked():
                    y_data = [d[i] for d in self.data if i < len(d)]
                    start_x = max(0, self.total_data_count - len(y_data))
                    x_data = list(range(start_x, self.total_data_count))
                    line.setData(x=x_data, y=y_data)
                    self.legend.addItem(line, self.checkboxes[i].text())
                else:
                    line.clear()

        self.plot_widget.enableAutoRange(axis='y')
        self.plot_widget.setXRange(max(0, self.total_data_count - self.max_points), self.total_data_count)

    def add_checkbox(self, name, color):
        checkbox = QCheckBox(name)
        checkbox.setChecked(True)
        checkbox.stateChanged.connect(self.update_plot_data)
        
        line_edit = QLineEdit(name)
        line_edit.setVisible(False)
        line_edit.editingFinished.connect(lambda: self.rename_checkbox(checkbox, line_edit))

        checkbox.mouseDoubleClickEvent = lambda event: self.edit_checkbox_name(checkbox, line_edit)

        self.checkbox_layout.addWidget(checkbox)
        self.checkbox_layout.addWidget(line_edit)
        self.checkboxes.append(checkbox)

    def edit_checkbox_name(self, checkbox, line_edit):
        checkbox.setVisible(False)
        line_edit.setVisible(True)
        line_edit.setFocus()

    def rename_checkbox(self, checkbox, line_edit):
        new_name = line_edit.text()
        checkbox.setText(new_name)
        checkbox.setVisible(True)
        line_edit.setVisible(False)
        self.update_plot_data()

    def load_settings(self):
        port = self.settings.value("port", "")
        baud = self.settings.value("baud", "9600")
        
        if port:
            index = self.port_combo.findText(port)
            if index >= 0:
                self.port_combo.setCurrentIndex(index)
        
        index = self.baud_combo.findText(baud)
        if index >= 0:
            self.baud_combo.setCurrentIndex(index)

        checkbox_names = self.settings.value("checkbox_names", [])
        checkbox_states = self.settings.value("checkbox_states", [])
        
        for name, state in zip(checkbox_names, checkbox_states):
            color = pg.intColor(len(self.checkboxes), hues=len(checkbox_names), values=1, maxValue=255)
            self.add_checkbox(name, color)
            self.checkboxes[-1].setChecked(state == "true")

        self.max_points = int(self.settings.value("max_points", 100))
        self.max_points_edit.setText(str(self.max_points))

        self.csv_filename_edit.setText(self.settings.value("csv_filename", "test"))
        folder = self.settings.value("csv_folder", "Not selected")
        self.csv_folder_label.setText(f"Selected Folder: {folder}")

    def save_settings(self):
        self.settings.setValue("port", self.port_combo.currentText())
        self.settings.setValue("baud", self.baud_combo.currentText())
        
        checkbox_names = [cb.text() for cb in self.checkboxes]
        checkbox_states = ["true" if cb.isChecked() else "false" for cb in self.checkboxes]
        
        self.settings.setValue("checkbox_names", checkbox_names)
        self.settings.setValue("checkbox_states", checkbox_states)
        self.settings.setValue("max_points", self.max_points)
        self.settings.setValue("csv_filename", self.csv_filename_edit.text())
        self.settings.setValue("csv_folder", self.csv_folder_label.text().replace("Selected Folder: ", ""))

    def restore_default(self):
        self.settings.clear()
        self.load_settings()
        self.clear_plot()
        for cb in self.checkboxes:
            self.checkbox_layout.removeWidget(cb)
            cb.deleteLater()
        self.checkboxes.clear()
        self.error_label.setText("Default settings restored")

    def select_csv_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.csv_folder_label.setText(f"Selected Folder: {folder}")

    def open_csv_file(self):
        folder = self.csv_folder_label.text().replace("Selected Folder: ", "")
        if folder == "Not selected":
            folder = os.getcwd()
        filename = self.csv_filename_edit.text() + ".csv"
        filepath = os.path.join(folder, filename)
        self.csv_file = open(filepath, 'a', newline='')
        self.csv_writer = csv.writer(self.csv_file)

    def close_csv_file(self):
        if self.csv_file:
            self.csv_file.close()
            self.csv_file = None
            self.csv_writer = None

    def closeEvent(self, event):
        self.save_settings()
        self.disconnect_serial()
        super().closeEvent(event)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    plotter = SerialPlotter()
    plotter.show()
    sys.exit(app.exec_())