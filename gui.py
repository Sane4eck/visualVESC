#gui.py
import base64
from PyQt5.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QLineEdit, QLabel, QComboBox, QSizePolicy
)
from PyQt5.QtGui import QColor, QPalette, QPixmap, QIcon
from PyQt5.QtCore import Qt, QTimer
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as Canvas
from matplotlib.figure import Figure

from ico.icon_bese64 import icon_base64
from logic import VESCWorker


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Іконка
        icon_data = base64.b64decode(icon_base64)
        pixmap = QPixmap()
        pixmap.loadFromData(icon_data)
        self.setWindowIcon(QIcon(pixmap))
        self.setWindowTitle("VESC Cyclogram")

        self.updating = True

        self.controller = VESCWorker()
        self.controller.data_ready.connect(self.update_plot)
        self.controller.connection_status.connect(self.update_connection_status)
        self.controller.mode_status.connect(self.update_mode_status)
        self.controller.lamp_status.connect(self.update_lamp)

        self.port_timer = QTimer()
        self.port_timer.timeout.connect(self.refresh_ports)
        self.port_timer.start(1000)

        self.x_data = []
        self.y_data = []
        self.duty_data = []
        self.current_data = []

        # --------------------- Графік ---------------------
        self.canvas = Canvas(Figure(figsize=(6, 4)))
        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.canvas.updateGeometry()
        self.canvas.figure.tight_layout()
        self.ax = self.canvas.figure.add_subplot(111)
        self.line, = self.ax.plot([], [], label="RPM", color="blue")
        self.ax.set_xlabel("Час (с)")
        self.ax.set_ylabel("RPM")
        self.ax.legend(loc="upper left")
        self.ax.grid(True)

        self.ax2 = self.ax.twinx()
        self.line_duty, = self.ax2.plot([], [], label="Duty", color="orange", linestyle="--")
        self.ax2.set_ylabel("Duty")
        self.ax2.legend(loc="upper center")

        self.ax3 = self.ax.twinx()
        self.ax3.spines["right"].set_position(("outward", 55))  # зсув праворуч
        self.line_current, = self.ax3.plot([], [], label="Current", color="green", linestyle=":")
        self.ax3.legend(loc="upper right")
        self.ax3.set_ylabel("Current (A)")
        self.canvas.figure.tight_layout()

        # --------------------- Параметри ---------------------
        self.pole_pairs_input = QLineEdit("3")
        self.reset_btn = QPushButton("Оновити")
        self.reset_btn.setStyleSheet("background-color: gray; font-size: 14px;")
        self.reset_btn.clicked.connect(self.reset_session)
        self.save_btn = QPushButton("Зберегти CSV")
        self.save_btn.setStyleSheet("background-color: orange; font-size: 14px;")
        self.save_btn.clicked.connect(self.save_csv)

        param_layout = QHBoxLayout()
        param_layout.addWidget(self.reset_btn)
        param_layout.addWidget(QLabel("pole pairs:"))
        param_layout.addWidget(self.pole_pairs_input)
        param_layout.addWidget(self.save_btn)

        # --------------------- Підключення ---------------------
        self.port_combo = QComboBox()
        self.refresh_ports()
        self.connect_btn = QPushButton("Підключити")
        self.disconnect_btn = QPushButton("Відключити")
        self.connection_label = QLabel("Статус: ❌")
        self.lamp_label = QLabel("   ")
        self.lamp_label.setAutoFillBackground(True)
        self.update_lamp("red")

        port_layout = QHBoxLayout()
        port_layout.addWidget(QLabel("COM порт:"))
        port_layout.addWidget(self.port_combo)
        port_layout.addWidget(self.connect_btn)
        port_layout.addWidget(self.disconnect_btn)
        port_layout.addWidget(self.connection_label)
        port_layout.addWidget(self.lamp_label)

        self.mode_label = QLabel("Mode: idle")
        port_layout.addWidget(self.mode_label)

        self.connect_btn.clicked.connect(self.connect_port)
        self.disconnect_btn.clicked.connect(self.disconnect_port)

        # --------------------- Циклограма ---------------------
        self.file_line = QLineEdit()
        self.file_line.setReadOnly(True)
        self.load_btn = QPushButton("Обрати файл")
        self.start_btn = QPushButton("СТАРТ")
        self.start_btn.setStyleSheet("background-color: green; font-size: 16px;")
        self.stop_btn = QPushButton("СТОП")
        self.stop_btn.setStyleSheet("background-color: red; font-size: 16px;")

        # НОВЕ: вибір режиму для циклограми
        self.cycle_mode_label = QLabel("Cycle by:")
        self.cycle_mode_combo = QComboBox()
        self.cycle_mode_combo.addItems(["Duty", "RPM"])

        self.manual_input = QLineEdit("0.07")
        self.manual_input.setPlaceholderText("Duty (0.0 ... 1.0)")
        self.manual_input.returnPressed.connect(self.manual_duty)

        self.manual_btn = QPushButton("Ручне")
        self.manual_btn.setStyleSheet("background-color: Blue; font-size: 16px;")

        self.manual_rpm_input = QLineEdit("1000")
        self.manual_rpm_input.setPlaceholderText("RPM mech")
        self.manual_rpm_input.returnPressed.connect(self.manual_rpm)

        self.manual_rpm_btn = QPushButton("Ручне RPM")
        self.manual_rpm_btn.setStyleSheet("background-color: purple; font-size: 16px;")
        self.manual_rpm_btn.clicked.connect(self.manual_rpm)

        cycle_layout = QHBoxLayout()
        cycle_layout.addWidget(self.file_line)
        cycle_layout.addWidget(self.load_btn)
        cycle_layout.addWidget(self.cycle_mode_label)     # нове
        cycle_layout.addWidget(self.cycle_mode_combo)     # нове
        cycle_layout.addWidget(self.start_btn)
        cycle_layout.addWidget(self.stop_btn)
        cycle_layout.addWidget(self.manual_btn)
        cycle_layout.addWidget(self.manual_input)
        cycle_layout.addWidget(self.manual_rpm_btn)
        cycle_layout.addWidget(self.manual_rpm_input)

        self.load_btn.clicked.connect(self.load_cycle)
        self.start_btn.clicked.connect(self.start_cycle)
        self.stop_btn.clicked.connect(self.stop_cycle)
        self.manual_btn.clicked.connect(self.manual_duty)

        # --------------------- RPM ---------------------
        self.rpm_display = QLabel("RPM: 0")
        self.rpm_display.setAlignment(Qt.AlignCenter)
        self.rpm_display.setStyleSheet("color: red; font-size: 20px; font-weight: bold;")

        self.current_display = QLabel("Current: 0 A")
        self.current_display.setAlignment(Qt.AlignCenter)

        info_layout = QHBoxLayout()
        info_layout.addWidget(self.rpm_display)
        info_layout.addWidget(self.current_display)

        # --------------------- Layout ---------------------
        layout = QVBoxLayout()
        layout.addLayout(param_layout)
        layout.addWidget(self.canvas)
        layout.addLayout(port_layout)
        layout.addLayout(cycle_layout)
        layout.addLayout(info_layout)
        self.setLayout(layout)

    def refresh_ports(self):
        prev = self.port_combo.currentText()
        new_ports = self.controller.get_available_ports()
        old_ports = [self.port_combo.itemText(i) for i in range(self.port_combo.count())]
        if new_ports == old_ports:
            return
        self.port_combo.blockSignals(True)
        self.port_combo.clear()
        self.port_combo.addItems(new_ports)
        if prev in new_ports:
            self.port_combo.setCurrentText(prev)
        self.port_combo.blockSignals(False)

    def connect_port(self):
        port = self.port_combo.currentText()
        self.controller.connect(port)

    def disconnect_port(self):
        self.controller.disconnect()

    def load_cycle(self):
        path, _ = QFileDialog.getOpenFileName(self, "Виберіть Excel-файл", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.file_line.setText(path)
            self.controller.load_cycle(path)

    def start_cycle(self):
        self.controller.pole_pairs = self.get_pole_pairs()
        chosen = self.cycle_mode_combo.currentText().strip().lower()
        self.controller.cycle_mode = "rpm" if chosen == "rpm" else "duty"
        self.controller.start_cycle()

    def save_csv(self):
        from datetime import datetime
        now = datetime.now()
        default_name = f"VESC_rpm_{now.strftime('%Y%m%d_%H%M%S')}.csv"
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти CSV", default_name, "CSV Files (*.csv)")
        if path:
            self.controller.export_csv(path)

    def stop_cycle(self):
        self.controller.stop_cycle()

    def manual_duty(self):
        try:
            duty = float(self.manual_input.text())
            self.controller.pole_pairs = self.get_pole_pairs()
            self.controller.set_manual_duty(duty)
            self.update_lamp("blue")
        except ValueError:
            pass

    def manual_rpm(self):
        try:
            rpm = float(self.manual_rpm_input.text())
            self.controller.pole_pairs = self.get_pole_pairs()
            self.controller.set_manual_rpm(rpm)
            self.update_lamp("purple")
        except ValueError:
            pass

    def update_plot(self, t, rpm, duty, current):
        if not getattr(self, "updating", True):
            return
        self.x_data.append(t)
        self.y_data.append(rpm)
        self.duty_data.append(duty)
        self.current_data.append(current)

        while self.x_data and (t - self.x_data[0] > 100):
            self.x_data.pop(0)
            self.y_data.pop(0)
            self.duty_data.pop(0)
            self.current_data.pop(0)

        if len(self.x_data) == len(self.y_data) == len(self.duty_data) == len(self.current_data):
            self.line.set_xdata(self.x_data)
            self.line.set_ydata(self.y_data)
            self.line_duty.set_xdata(self.x_data)
            self.line_duty.set_ydata(self.duty_data)
            self.line_current.set_xdata(self.x_data)
            self.line_current.set_ydata(self.current_data)

        if t > 0.1:
            self.ax.set_xlim(max(0, t - 100), t)
        else:
            self.ax.set_xlim(0, 1)

        self.ax.relim()
        self.ax.autoscale_view(True, True, True)
        self.ax2.relim()
        self.ax2.autoscale_view(True, True, True)
        self.ax3.relim()
        self.ax3.autoscale_view(True, True, True)

        self.canvas.draw()
        self.rpm_display.setText(f"RPM: {int(rpm)}")
        self.current_display.setText(f"Current: {current:.2f} A")

    def update_connection_status(self, status):
        self.connection_label.setText("Статус: ✅" if status else "Статус: ❌")

    def update_mode_status(self, mode):
        self.mode_label.setText(f"Mode: {mode}")

    def update_lamp(self, color):
        palette = self.lamp_label.palette()
        if color == "green":
            palette.setColor(QPalette.Window, QColor(0, 255, 0))
        elif color == "blue":
            palette.setColor(QPalette.Window, QColor(0, 0, 255))
        elif color == "purple":
            palette.setColor(QPalette.Window, QColor(128, 0, 128))
        else:
            palette.setColor(QPalette.Window, QColor(255, 0, 0))
        self.lamp_label.setPalette(palette)

    def get_pole_pairs(self):
        try:
            return int(float(self.pole_pairs_input.text()))
        except ValueError:
            return 1

    def reset_session(self):
        selected_port = self.port_combo.currentText()
        self.x_data = []
        self.y_data = []
        self.duty_data = []
        self.current_data = []
        self.canvas.draw()
        self.rpm_display.setText("RPM: 0")
        self.controller.reset_session()
        self.refresh_ports()
        if selected_port in [self.port_combo.itemText(i) for i in range(self.port_combo.count())]:
            self.port_combo.setCurrentText(selected_port)

    def closeEvent(self, event):
        self.controller.disconnect()
        super().closeEvent(event)

    def refresh_graphs(self):
        self.updating = False
        self.x_data.clear()
        self.y_data.clear()
        self.duty_data.clear()
        self.current_data.clear()
        self.canvas.draw()
        self.rpm_display.setText("RPM: 0")
        self.updating = True
