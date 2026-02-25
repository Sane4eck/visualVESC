#logic.py
import time
import csv
import threading
import serial
import serial.tools.list_ports
from pyvesc import encode, encode_request, decode
from pyvesc.VESC.messages import SetDutyCycle, GetValues, SetRPM
from PyQt5.QtCore import QObject, pyqtSignal, QThread


class VESCWorker(QObject):
    data_ready = pyqtSignal(float, float, float, float)  # elapsed_time, rpm, duty, current
    connection_status = pyqtSignal(bool)
    mode_status = pyqtSignal(str)
    lamp_status = pyqtSignal(str)
    error = pyqtSignal(str)
    log = pyqtSignal(str)

    def __init__(self, baudrate=115200, csv_file="rpm_log.csv", parent=None):
        super().__init__(parent)
        self.ser = None
        self.baudrate = baudrate

        self.running = False
        self.manual_duty = None
        self.manual_rpm = None
        self.control_mode = "duty"
        self.cycle_data = []              # для сумісності (duty)
        self.cycle_active = False
        self.start_time = time.time()
        self.last_save_time = time.time()
        self.csv_file = csv_file
        self.lock = threading.Lock()
        self.csv_lock = threading.Lock()
        self.pole_pairs = 1
        self.cycle_index = 0
        self.cycle_start_time = time.time()
        self._next_save_time = time.time() + 0.1

        # НОВЕ: режим і окремі масиви для duty/rpm
        self.cycle_mode = "duty"          # 'duty' або 'rpm'
        self.cycle_data_duty = []         # [(duration, duty)]
        self.cycle_data_rpm = []          # [(duration, rpm_mech)]

        # QThread для циклу зчитування
        self._thread = QThread()
        self.moveToThread(self._thread)
        self._thread.started.connect(self._read_loop)
        self._thread.start()

        # CSV header
        with open(self.csv_file, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["elapsed_time_sec", "rpm", "duty", "current"])

    # ---------- Порти ----------
    def get_available_ports(self):
        ports = serial.tools.list_ports.comports()
        return [p.device for p in ports]

    def connect(self, port):
        try:
            if self.ser and self.ser.is_open:
                self.log.emit("Already connected")
                self.connection_status.emit(True)
                self.lamp_status.emit("green")
                return True

            self.ser = serial.Serial(port, self.baudrate, timeout=0.1)
            try:
                self.ser.reset_input_buffer()
                self.ser.reset_output_buffer()
            except Exception:
                pass
            time.sleep(0.1)

            self.running = True
            self.connection_status.emit(True)
            self.lamp_status.emit("green")
            self.log.emit(f"Connected to {port}")
            return True
        except Exception as e:
            self.ser = None
            self.running = False
            self.connection_status.emit(False)
            self.lamp_status.emit("red")
            self.error.emit(f"Connect error: {e}")
            return False

    def disconnect(self):
        self.running = False
        try:
            if self.ser and self.ser.is_open:
                try:
                    self._set_duty(0)
                except Exception:
                    pass
                self.ser.flush()
                self.ser.close()
        except Exception as e:
            self.log.emit(f"Error closing serial: {e}")
        self.ser = None
        self.connection_status.emit(False)
        self.lamp_status.emit("red")
        self.cycle_active = False
        self.manual_duty = None
        try:
            self._set_duty(0)
        except:
            pass

    # ---------- Циклограма ----------
    def load_cycle(self, filepath):
        import pandas as pd
        try:
            df = pd.read_excel(filepath)
            cols = {c.lower(): c for c in df.columns}
            if "duration" not in cols:
                raise ValueError("Excel file must contain 'duration' column")

            dur = df[cols["duration"]].astype(float)

            self.cycle_data_duty = []
            self.cycle_data_rpm = []
            # duty колонка опціональна
            if "duty" in cols:
                duty = df[cols["duty"]].astype(float)
                self.cycle_data_duty = list(zip(dur, duty))
            # rpm колонка опціональна
            if "rpm" in cols:
                rpm = df[cols["rpm"]].astype(float)
                self.cycle_data_rpm = list(zip(dur, rpm))

            # сумісність зі старою логікою
            self.cycle_data = list(self.cycle_data_duty)

        except Exception as e:
            self.error.emit(f"Помилка при завантаженні циклограми: {e}")
            self.cycle_data_duty = []
            self.cycle_data_rpm = []
            self.cycle_data = []

    def start_cycle(self):
        with self.lock:
            # вибираємо активні дані відповідно до режиму
            if self.cycle_mode == "rpm":
                active = self.cycle_data_rpm
            else:
                active = self.cycle_data_duty if self.cycle_data_duty else self.cycle_data

            if not active:
                self.cycle_active = False
                self.manual_duty = None
                self.manual_rpm = None
                self.cycle_index = 0
                self.mode_status.emit("idle")
                self.lamp_status.emit("red")
                self.error.emit("Немає даних для обраного режиму циклограми. Додайте колонку 'duty' або 'rpm'.")
                return

            self.cycle_active = True
            self.manual_duty = None
            self.manual_rpm = None
            self.cycle_index = 0
            self.cycle_start_time = time.time()
        self.mode_status.emit("cycle")
        self.lamp_status.emit("green")

    def stop_cycle(self):
        with self.lock:
            self.cycle_active = False
            self.manual_duty = None
            self.manual_rpm = None
            self.cycle_index = 0
        self.mode_status.emit("idle")
        try:
            self._set_duty(0)
        except:
            pass
        self.data_ready.emit(time.time() - self.start_time, 0, 0, 0.0)
        self.lamp_status.emit("red")

    def set_manual_duty(self, duty):
        with self.lock:
            self.manual_duty = max(0.0, min(1.0, float(duty)))
            self.manual_rpm = None
            self.cycle_active = False
            self.control_mode = "duty"
        self.mode_status.emit("manual")
        self.lamp_status.emit("blue")

    # ---------- Керування ----------
    def _set_duty(self, duty):
        try:
            if not self.ser or not getattr(self.ser, "is_open", False):
                return
            duty = max(0.0, min(1.0, float(duty)))
            self.ser.write(encode(SetDutyCycle(duty)))
        except Exception as e:
            self.error.emit(f"_set_duty error: {e}")

    def set_manual_rpm(self, rpm_mech):
        with self.lock:
            self.manual_rpm = float(rpm_mech)
            self.manual_duty = None
            self.cycle_active = False
            self.control_mode = "rpm"
        self.mode_status.emit("manual")
        self.lamp_status.emit("purple")

    def _set_rpm(self, rpm_mech):
        try:
            if not self.ser or not getattr(self.ser, "is_open", False):
                return
            erpm = int(float(rpm_mech) * float(self.pole_pairs))
            self.ser.write(encode(SetRPM(erpm)))
        except Exception as e:
            self.error.emit(f"_set_rpm error: {e}")

    # ---------- Основний цикл ----------
    def _read_loop(self):
        while True:
            if self.running and self.ser and self.ser.is_open:
                current_time = time.time() - self.start_time
                duty_target = None
                rpm_target = None
                try:
                    with self.lock:
                        # пріоритет: manual_rpm -> manual_duty -> cycle -> idle
                        if self.manual_rpm is not None:
                            rpm_target = self.manual_rpm
                            self.mode_status.emit("manual")
                            self.lamp_status.emit("purple")
                        elif self.manual_duty is not None:
                            duty_target = self.manual_duty
                            self.mode_status.emit("manual")
                            self.lamp_status.emit("blue")
                        elif self.cycle_active and (self.cycle_data_duty or self.cycle_data_rpm or self.cycle_data):
                            # обрати активний набір під режим
                            if self.cycle_mode == "rpm":
                                active = self.cycle_data_rpm
                            else:
                                active = self.cycle_data_duty if self.cycle_data_duty else self.cycle_data

                            if self.cycle_index < len(active):
                                duration, value = active[self.cycle_index]
                                if (time.time() - self.cycle_start_time) >= duration:
                                    self.cycle_index += 1
                                    self.cycle_start_time = time.time()
                            else:
                                self.cycle_active = False
                                self.manual_duty = None
                                self.mode_status.emit("idle")
                                self.lamp_status.emit("red")
                                self._set_duty(0)
                                self.data_ready.emit(current_time, 0, 0, 0.0)
                                continue

                            if self.cycle_index < len(active):
                                _, value = active[self.cycle_index]
                                if self.cycle_mode == "rpm":
                                    rpm_target = float(value)
                                else:
                                    duty_target = float(value)
                                self.mode_status.emit("cycle")
                                self.lamp_status.emit("green")
                        else:
                            duty_target = 0
                            self.mode_status.emit("idle")
                            self.lamp_status.emit("red")

                    # Команда на VESC
                    if rpm_target is not None:
                        self._set_rpm(rpm_target)
                        duty_for_emit = 0
                    elif duty_target is not None:
                        self._set_duty(duty_target)
                        duty_for_emit = duty_target
                    else:
                        duty_for_emit = 0

                    # Запит значень з VESC
                    self.ser.write(encode_request(GetValues))
                    time.sleep(0.001)
                    response = self.ser.read(256)
                    values, _ = decode(response)
                    if values and hasattr(values, "rpm"):
                        erpm = values.rpm
                        rpm = erpm / self.pole_pairs if self.pole_pairs else erpm
                        motor_current = getattr(values, "avg_motor_current", 0.0)
                        self.data_ready.emit(current_time, rpm, duty_for_emit, motor_current)

                        now = time.time()
                        if now >= self._next_save_time:
                            with open(self.csv_file, "a", newline="") as f:
                                writer = csv.writer(f)
                                writer.writerow([current_time, rpm, duty_for_emit, motor_current])
                            self._next_save_time += 0.1
                except (serial.SerialException, OSError):
                    self.disconnect()
                except Exception as e:
                    self.error.emit(f"_read_loop error: {e}")
            time.sleep(0.005)

    # ---------- Скидання сесії ----------
    def reset_session(self):
        with self.lock:
            self.start_time = time.time()
            self.last_save_time = time.time()
            self.cycle_index = 0
            self.cycle_start_time = time.time()
            self.manual_duty = None
            self.cycle_active = False
            with self.csv_lock:
                with open(self.csv_file, "w", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["elapsed_time_sec", "rpm", "duty", "current"])
            self._next_save_time = time.time() + 0.1
        self.mode_status.emit("idle")
        self.lamp_status.emit("red")
        try:
            self._set_duty(0)
        except:
            pass
        self.data_ready.emit(0.0, 0.0, 0.0, 0.0)

    # ---------- Експорт CSV ----------
    def export_csv(self, path):
        try:
            with open(path, "w", newline="") as f_out:
                writer = csv.writer(f_out)
                with self.csv_lock:
                    with open(self.csv_file, "r", newline="") as f_in:
                        for row in csv.reader(f_in):
                            writer.writerow(row)
        except Exception as e:
            self.error.emit(f"Помилка при збереженні CSV: {e}")
