from PyQt6 import QtCore, QtGui, QtWidgets
import pandas as pd
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import subprocess
import sys
from pathlib import Path

# Lấy đường dẫn đầu vào
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return str(Path(sys._MEIPASS) / relative_path)
    else:
        return str(Path(__file__).parent / relative_path)

class Worker(QtCore.QThread):
    progress = QtCore.pyqtSignal(int)
    log = QtCore.pyqtSignal(str)
    finished = QtCore.pyqtSignal(object)

    def __init__(self, mnv, output_path):
        super().__init__()
        self.mnv = mnv
        self.output_path = output_path

    def run(self):
        self.log.emit("\U0001f300 Đang khởi tạo trình duyệt...")
        self.progress.emit(10)
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        browser = webdriver.Chrome(options=options)
        template_path = resource_path("test.docx")

        try:
            if not self.mnv.isnumeric():
                self.log.emit("❌ Mã nhân viên không hợp lệ.")
                self.progress.emit(0)
                self.finished.emit(None)
                return

            self.log.emit("🌐 Đang truy cập website...")
            browser.get("http://scfp.vn/Productscan.aspx")
            browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtMANV").send_keys(self.mnv)
            browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdbLoai_2").click()
            browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtXem").click()

            self.progress.emit(40)
            self.log.emit("📦 Đang lấy dữ liệu...")
            time.sleep(1)
            tbody = browser.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_gvPd"]/tbody')
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            data = [[cell.text for cell in row.find_elements(By.TAG_NAME, 'td')] for row in rows]

            df = pd.DataFrame(data[:], columns=data[0])
            df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], format="%m/%d/%Y %I:%M:%S %p").dt.strftime('%d/%m/%Y')
            df.iloc[:, 3] = df.iloc[:, 4] + ' / ' + df.iloc[:, 3]
            df.iloc[:, 4] = df.iloc[:, 5]
            df.iloc[:, 5] = ""
            df = df.iloc[:, :-1]

            self.progress.emit(70)
            self.log.emit("📝 Đang tạo Word...")

            document = Document(template_path)
            table = document.tables[0]

            for i, row in enumerate(df.values, start=1):
                if i >= len(table.rows):
                    table.add_row()
                for j, value in enumerate(row):
                    cell = table.rows[i].cells[j]
                    cell.text = str(value)
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            for i in range(len(table.rows) - 1, 0, -1):
                if not any(cell.text.strip() for cell in table.rows[i].cells):
                    table._element.remove(table.rows[i]._element)

            document.save(self.output_path)
            time.sleep(1)
            subprocess.Popen(['start', self.output_path], shell=True)
            self.progress.emit(100)
            self.finished.emit(df)

        except PermissionError:
            self.log.emit(f"❌ Lỗi: File word kết quả đang mở hãy đóng file và thử lại")
            self.progress.emit(0)
            self.finished.emit(None)
        except Exception as e:
            self.log.emit(f"❌ Lỗi: {str(e)}")
            self.progress.emit(0)
            self.finished.emit(None)
        finally:
            browser.quit()

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(620, 720)
        MainWindow.setStyleSheet("""
            QWidget {
                background-color: #f1fafe;
                font-family: 'Segoe UI';
            }
            QPushButton {
                background-color: #007acc;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 8px;
                font-size: 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #005b99;
            }
            QLineEdit, QTextEdit {
                background-color: #ffffff;
                border: 2px solid #d0eafc;
                border-radius: 8px;
                padding: 6px;
                font-size: 13px;
            }
            QLineEdit:invalid {
                border-color: red;
            }
            QLabel {
                color: #003366;
                font-size: 13pt;
            }
            QProgressBar {
                height: 24px;
                border-radius: 8px;
                background: #d0eafc;
            }
            QProgressBar::chunk {
                background-color: #007acc;
                border-radius: 8px;
                transition: all 0.5s ease-in-out;
            }
        """)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        layout = QtWidgets.QVBoxLayout(self.centralwidget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        label_mnv_layout = QtWidgets.QHBoxLayout()
        self.label_mnv = QtWidgets.QLabel("Nhập mã nhân viên:")
        self.mnv = QtWidgets.QLineEdit()
        self.mnv.setFont(QtGui.QFont("Segoe UI", 11))
        self.mnv.textChanged.connect(self.validate_input)
        label_mnv_layout.addWidget(self.label_mnv)
        label_mnv_layout.addWidget(self.mnv)

        btn_layout = QtWidgets.QHBoxLayout()
        self.create_btn = QtWidgets.QPushButton("Tạo phiếu xuất hàng")
        self.create_btn.clicked.connect(self.generate_report)
        self.clear_btn = QtWidgets.QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_input)
        btn_layout.addWidget(self.create_btn)
        btn_layout.addWidget(self.clear_btn)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setValue(0)

        self.log_output = QtWidgets.QTextEdit()
        self.log_output.setReadOnly(True)

        self.spinner_label = QtWidgets.QLabel()
        self.spinner_movie = QtGui.QMovie(resource_path("spinner.gif"))
        self.spinner_label.setMovie(self.spinner_movie)
        self.spinner_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.spinner_label.hide()

        layout.addLayout(label_mnv_layout)
        layout.addLayout(btn_layout)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.log_output)
        layout.addWidget(self.spinner_label)

        MainWindow.setCentralWidget(self.centralwidget)

    def validate_input(self):
        text = self.mnv.text()
        if not text.isnumeric():
            self.mnv.setStyleSheet("border: 2px solid red;")
        else:
            self.mnv.setStyleSheet("border: 2px solid #d0eafc;")

    def on_report_finished(self, df):
        self.spinner_movie.stop()
        self.spinner_label.hide()
        if df is not None:
            self.log_output.append("🎉 Hoàn thành!")

    def generate_report(self):
        mnv = self.mnv.text()
        if not mnv:
            QtWidgets.QMessageBox.warning(None, "Lỗi", "Vui lòng nhập mã nhân viên")
            return

        self.log_output.clear()
        self.progress_bar.setValue(0)
        self.spinner_label.show()
        self.spinner_movie.start()

        output_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            None, "Lưu phiếu xuất hàng", "phieu_xuat.docx", "Word Documents (*.docx)")
        if not output_path:
            self.log_output.append("❌ Đã hủy lưu file.")
            self.spinner_movie.stop()
            self.spinner_label.hide()
            return

        self.thread = Worker(mnv, output_path)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.log.connect(self.log_output.append)
        self.thread.finished.connect(self.on_report_finished)
        self.thread.start()

    def clear_input(self):
        self.mnv.clear()
        self.progress_bar.setValue(0)
        self.log_output.clear()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
