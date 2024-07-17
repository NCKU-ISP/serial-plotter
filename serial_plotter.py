import sys
import os
import serial
import serial.tools.list_ports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QComboBox, QPushButton, QCheckBox, QLineEdit,
                             QGroupBox, QLabel, QFileDialog, QScroller)
from PyQt5.QtCore import QTimer, Qt, QSettings, pyqtSignal, QEvent
from PyQt5.QtGui import QIntValidator, QFontMetrics
import pyqtgraph as pg
import numpy as np
import time
import csv

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
        self.max_points = 200
        self.legend = None
        self.total_data_count = 0
        self.csv_file = None
        self.csv_writer = None
        self.checkbox_widgets = []

        self.init_ui()
        self.load_settings()

        # Modify the legend creation
        self.legend = pg.LegendItem((-1, -1), offset=(70,20))
        self.legend.setParentItem(self.plot_widget.graphicsItem())
        self.legend.anchor(itemPos=(1, 0), parentPos=(1, 0), offset=(-10, 10))
        self.legend.setVisible(False)  # Hide legend by default

        # Add border and reduce spacing
        self.legend.layout.setSpacing(1)
        self.legend.layout.setContentsMargins(5, 5, 5, 5)
        self.legend.setBrush(pg.mkBrush(255, 255, 255, 200))
        self.legend.setPen(pg.mkPen(color=(0, 0, 0), width=1))

        self.installEventFilter(self)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # Control panel
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

        # Max data points limit
        data_points_layout = QHBoxLayout()
        data_points_layout.addWidget(QLabel("Max data points:"))
        self.max_points_edit = QLineEdit(str(self.max_points))
        self.max_points_edit.setValidator(QIntValidator(1, 10000))
        self.max_points_edit.returnPressed.connect(self.update_max_points)
        data_points_layout.addWidget(self.max_points_edit)
        control_layout.addLayout(data_points_layout)

        # CSV file settings
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

        # Checkbox panel
        self.checkbox_layout = QVBoxLayout()
        control_layout.addLayout(self.checkbox_layout)

        control_layout.addStretch(1)

        # Error message label
        self.error_label = QLabel()
        self.error_label.setWordWrap(True)
        control_layout.addWidget(self.error_label)

        # Restore Default button
        self.restore_button = QPushButton("Restore Default")
        self.restore_button.clicked.connect(self.restore_default)
        control_layout.addWidget(self.restore_button)

        # Add designer and supporter information
        designer_label = QLabel("Designed by Nickchung")
        designer_label.setAlignment(Qt.AlignCenter)
        control_layout.addWidget(designer_label)

        powered_label = QLabel("Powered by NCKU ISP")
        powered_label.setAlignment(Qt.AlignCenter)
        control_layout.addWidget(powered_label)

        # Plot
        plot_layout = QVBoxLayout()
        main_layout.addLayout(plot_layout)

        self.plot_widget = pg.PlotWidget(background='w')
        self.plot_widget.getAxis('bottom').setPen('k')
        self.plot_widget.getAxis('left').setPen('k')
        plot_layout.addWidget(self.plot_widget)

        # Create legend
        self.legend = pg.LegendItem(size=(100,60), offset=(70,20))
        self.legend.setParentItem(self.plot_widget.graphicsItem())
        self.legend.anchor(itemPos=(1, 0), parentPos=(1, 0), offset=(-10, 10))

        # Timer
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

        # Clear checkboxes but keep their names
        for container, checkbox, line_edit, delete_button in self.checkbox_widgets:
            self.checkbox_layout.removeWidget(container)
            container.deleteLater()
        self.checkboxes.clear()
        self.checkbox_widgets.clear()

        # Recreate checkboxes with saved names
        checkbox_names = self.settings.value("checkbox_names", [])
        for name in checkbox_names:
            color = pg.intColor(len(self.checkboxes), hues=len(checkbox_names), values=1, maxValue=255)
            self.add_checkbox(name, color)

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
                        # Ignore non-numeric input
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
        has_visible_data = False
        visible_items = 0
        max_text_width = 0
        font_metrics = QFontMetrics(self.font())

        for i, line in enumerate(self.lines):
            if i < len(self.checkboxes):
                if self.checkboxes[i].isChecked():
                    y_data = [d[i] for d in self.data if i < len(d)]
                    start_x = max(0, self.total_data_count - len(y_data))
                    x_data = list(range(start_x, self.total_data_count))
                    line.setData(x=x_data, y=y_data)
                    self.legend.addItem(line, self.checkboxes[i].text())
                    has_visible_data = True
                    visible_items += 1

                    text_width = font_metrics.width(self.checkboxes[i].text())
                    max_text_width = max(max_text_width, text_width)
                else:
                    line.clear()

        # Dynamically adjust legend size
        if has_visible_data:
            self.legend.setVisible(True)
            new_height = max(30, visible_items * 23)
            new_width = max_text_width + 50
            self.legend.setGeometry(0, 0, new_width, new_height)
        else:
            self.legend.setVisible(False)
            self.legend.setGeometry(0, 0, 0, 0)

        self.plot_widget.enableAutoRange(axis='y')
        self.plot_widget.setXRange(max(0, self.total_data_count - self.max_points), self.total_data_count)

    def add_checkbox(self, name, color):
        checkbox = QCheckBox(name)
        checkbox.setChecked(True)
        checkbox.stateChanged.connect(self.update_plot_data)
        
        line_edit = QLineEdit(name)
        line_edit.setVisible(False)
        line_edit.editingFinished.connect(lambda: self.rename_checkbox(checkbox, line_edit))

        delete_button = QPushButton("×")
        delete_button.setFixedSize(20, 20)
        delete_button.setStyleSheet("""
            QPushButton {
                border: none;
                background-color: transparent;
                color: gray;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                color: red;
            }
        """)
        delete_button.clicked.connect(lambda: self.delete_checkbox(checkbox, line_edit, delete_button))

        # Modify double-click event
        checkbox.mouseDoubleClickEvent = lambda event: self.edit_checkbox_name(checkbox, line_edit, event)

        container = QWidget()
        layout = QHBoxLayout(container)
        layout.addWidget(checkbox)
        layout.addWidget(line_edit)
        layout.addWidget(delete_button)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        self.checkbox_layout.addWidget(container)
        self.checkboxes.append(checkbox)
        self.checkbox_widgets.append((container, checkbox, line_edit, delete_button))

    def delete_checkbox(self, checkbox, line_edit, delete_button):
        for container, cb, le, db in self.checkbox_widgets:
            if cb == checkbox:
                self.checkbox_layout.removeWidget(container)
                container.deleteLater()
                self.checkboxes.remove(checkbox)
                self.checkbox_widgets.remove((container, cb, le, db))
                break
        self.update_plot_data()
        self.save_checkbox_names()  # Save names after deletion

    def edit_checkbox_name(self, checkbox, line_edit, event):
        event.accept()  # Prevent event propagation to checkbox click handler
        checkbox.setVisible(False)
        line_edit.setVisible(True)
        line_edit.setFocus()
        line_edit.selectAll()

    def rename_checkbox(self, checkbox, line_edit):
        new_name = line_edit.text()
        checkbox.setText(new_name)
        checkbox.setVisible(True)
        line_edit.setVisible(False)
        self.update_plot_data()
        self.save_checkbox_names()  # Save names immediately after renaming

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonPress:
            for _, checkbox, line_edit, _ in self.checkbox_widgets:
                if line_edit.isVisible():
                    self.rename_checkbox(checkbox, line_edit)
                    return True
        return super().eventFilter(obj, event)

    def save_checkbox_names(self):
        checkbox_names = [cb.text() for cb in self.checkboxes]
        self.settings.setValue("checkbox_names", checkbox_names)

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

        # If no saved names, add default checkboxes
        if not checkbox_names:
            for i in range(1, 6):  # Add 5 default checkboxes
                color = pg.intColor(i-1, hues=5, values=1, maxValue=255)
                self.add_checkbox(f"Data {i}", color)

        self.max_points = int(self.settings.value("max_points", 200))
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

        self.save_checkbox_names()  # Ensure checkbox names are saved

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