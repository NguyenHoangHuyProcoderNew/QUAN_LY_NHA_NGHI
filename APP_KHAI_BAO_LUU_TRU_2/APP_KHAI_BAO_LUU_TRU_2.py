import os
import sys

# Tắt output để tránh lỗi khi build exe và tăng tốc khởi động
if hasattr(sys, 'frozen'):
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(sys._MEIPASS, 'platforms')
    os.environ['QT_ENABLE_HIGHDPI_SCALING'] = '1'
    # Tắt logging của OpenCV
    os.environ["OPENCV_LOG_LEVEL"] = "OFF"
    # Tắt console output
    if not sys.stdout:
        sys.stdout = open(os.devnull, 'w')
    if not sys.stderr:
        sys.stderr = open(os.devnull, 'w')

# Import các thư viện cần thiết
from PyQt5 import QtCore, QtGui, QtWidgets, QtMultimedia, QtMultimediaWidgets
from PyQt5.QtMultimedia import QCameraImageCapture
import cv2
cv2.setLogLevel(0)  # Tắt logging của OpenCV

# Import các module khác theo thứ tự cần thiết
from datetime import datetime, timezone, timedelta
from unidecode import unidecode
import json
from qreader import QReader
import winsound
import time
import win32com.client as win32
import shutil
import win32gui
import win32process
import win32con
import psutil

# Tối ưu PyQt
QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

# Cache các biến toàn cục để tránh import nhiều lần
EXCEL_APP = None
QR_READER = None

def get_excel_app():
    global EXCEL_APP
    if EXCEL_APP is None:
        EXCEL_APP = win32.gencache.EnsureDispatch("Excel.Application")
    return EXCEL_APP

def get_qr_reader():
    global QR_READER
    if QR_READER is None:
        QR_READER = QReader()
    return QR_READER

def chuan_hoa_ngay(dmy: str) -> str:
    return f"{dmy[:2]}/{dmy[2:4]}/{dmy[4:]}" if len(dmy) == 8 else dmy

def parse_qr(data: str) -> dict:
    parts = data.split('|')
    if len(parts) >= 7:
        return {
            "Số giấy tờ": parts[0],
            "Số CMND cũ (nếu có)": parts[1],
            "Họ và tên": parts[2],
            "Ngày sinh": chuan_hoa_ngay(parts[3]),
            "Giới tính": parts[4],
            "Nơi thường trú": parts[5],
            "Ngày cấp giấy tờ": chuan_hoa_ngay(parts[6])
        }
    return None

def save_camera_config(qr_index, cam_index, path="config_cam.json"):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"qr": qr_index, "cam": cam_index}, f)
    except:
        pass

def load_camera_config(path="config_cam.json"):
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("qr", 0), data.get("cam", 0)
    except:
        return 0, 0

class QRDecodeThread(QtCore.QThread):
    qrDecoded = QtCore.pyqtSignal(str)
    def __init__(self, qreader, parent=None):
        super().__init__(parent)
        self.qreader = qreader
        self._running = True
        self._image = None
        self._lock = QtCore.QMutex()
        self._decode_next = False

    def stop(self):
        self._running = False
        self.wait()

    def request_decode(self, image):
        with QtCore.QMutexLocker(self._lock):
            self._image = image.copy() if image is not None else None
            self._decode_next = True

    def run(self):
        while self._running:
            image_to_decode = None
            with QtCore.QMutexLocker(self._lock):
                if self._decode_next and self._image is not None:
                    image_to_decode = self._image.copy()
                    self._decode_next = False
            if image_to_decode is not None:
                small = cv2.resize(image_to_decode, (320, 240))
                decoded = self.qreader.detect_and_decode(image=small)
                if decoded and decoded[0]:
                    self.qrDecoded.emit(decoded[0])
            self.msleep(10)

class OptionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, current_qr_idx=-1, current_cam_idx=-1):
        super().__init__(parent)
        self.setWindowTitle("Cài đặt")
        self.setModal(True)
        self.resize(500, 300)
        self.available_cameras = QtMultimedia.QCameraInfo.availableCameras()
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(20)
        grp = QtWidgets.QGroupBox("Cài đặt webcam quét thông tin và chụp ảnh mặt trước/sau")
        vbox = QtWidgets.QVBoxLayout(grp)
        vbox.setContentsMargins(10, 10, 10, 10)
        vbox.setSpacing(15)

        # --- Chọn webcam đọc QR ---
        h1 = QtWidgets.QHBoxLayout()
        lbl1 = QtWidgets.QLabel("Webcam đọc QR:")
        lbl1.setFixedWidth(120)
        lbl1.setStyleSheet("color: black;")
        self.cmb_qr = QtWidgets.QComboBox()
        self.cmb_qr.setStyleSheet("""
            QComboBox {
                background-color: #F7F7F7;
                color: #222222;
                padding: 4px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
            QComboBox::drop-down { border: none; }
            QComboBox QAbstractItemView {
                background-color: #FFFFFF;
                color: #222222;
                selection-background-color: #99CCFF;
                selection-color: #222222;
            }
        """)
        for cam in self.available_cameras:
            self.cmb_qr.addItem(cam.description())
        h1.addWidget(lbl1)
        h1.addWidget(self.cmb_qr)
        vbox.addLayout(h1)

        # --- Chọn webcam chụp ảnh ---
        h2 = QtWidgets.QHBoxLayout()
        lbl2 = QtWidgets.QLabel("Webcam chụp ảnh:")
        lbl2.setFixedWidth(120)
        lbl2.setStyleSheet("color: black;")
        self.cmb_cam = QtWidgets.QComboBox()
        self.cmb_cam.setStyleSheet("""
            QComboBox {
                background-color: #F7F7F7;
                color: #222222;
                padding: 4px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
            QComboBox::drop-down { border: none; }
            QComboBox QAbstractItemView {
                background-color: #FFFFFF;
                color: #222222;
                selection-background-color: #99CCFF;
                selection-color: #222222;
            }
        """)
        for cam in self.available_cameras:
            self.cmb_cam.addItem(cam.description())
        h2.addWidget(lbl2)
        h2.addWidget(self.cmb_cam)
        vbox.addLayout(h2)

        if 0 <= current_qr_idx < len(self.available_cameras):
            self.cmb_qr.setCurrentIndex(current_qr_idx)
        if 0 <= current_cam_idx < len(self.available_cameras):
            self.cmb_cam.setCurrentIndex(current_cam_idx)

        layout.addWidget(grp)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_save = QtWidgets.QPushButton("Lưu")
        btn_save.setFixedSize(100, 35)
        btn_save.setStyleSheet("""
            QPushButton {
                background-color: #005BEA;  /* Xanh dương đậm */
                color: white;
                font-size: 18px;                   
                font-weight: bold;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #3366FF;  /* Màu sáng hơn khi hover */
            }
        """)

        btn_save.clicked.connect(self.accept)
        btn_cancel = QtWidgets.QPushButton("Trở lại")
        btn_cancel.setFixedSize(100, 35)
        btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #FF2222;  /* Đỏ sáng rõ */
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #FF5555;  /* Đậm hơn một chút khi hover */
            }
        """)

        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

    def get_selected_camera_indexes(self):
        return self.cmb_qr.currentIndex(), self.cmb_cam.currentIndex()
    
class FloatingCalendarDateEdit(QtWidgets.QDateEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCalendarPopup(True)
        calendar = QtWidgets.QCalendarWidget()
        calendar.setVerticalHeaderFormat(QtWidgets.QCalendarWidget.NoVerticalHeader)
        calendar.setFixedSize(360, 260)  # Kích thước lịch popup
        self.setCalendarWidget(calendar)
        self.setDisplayFormat("dd/MM/yyyy")
        self.setDate(QtCore.QDate.currentDate())  # Đặt ngày hiện tại

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Phần mềm khai báo thông tin lưu trú được lập trình bởi Nguyễn Hoàng Huy - My phone: 033.293.6390")

        # Khởi tạo thuộc tính can_decode
        self.can_decode = True

        # Xác định đường dẫn thư mục app
        if hasattr(sys, 'frozen'):
            self.app_dir = sys._MEIPASS
        else:
            self.app_dir = os.path.dirname(os.path.abspath(__file__))

        # Khởi tạo đường dẫn Excel và folder ảnh
        self.excel_path = os.path.join(self.app_dir, "DS_Cong_Dan.xlsx")
        self.image_folder = os.path.join(self.app_dir, "DS_Anh_Cong_Dan_da_khai_bao")

        # ✅ Gán icon cửa sổ
        icon_path = os.path.join(self.app_dir, "logo_app.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QtGui.QIcon(icon_path))
    
        # Thay vì resize cứng, chỉ đặt minimumSize để cho phép người dùng kéo to nhỏ
        self.setMinimumSize(1200, 850)
        self.resize(1200, 800)
        self.setFixedSize(self.width(), self.height()) # Khóa không cho thay đổi kích thước cửa sổ giao diện
        
        # Khởi tạo các biến với giá trị mặc định
        self.cap_qr = None
        self.timer_cv = None
        self.last_qr_text = ""
        self.qreader = get_qr_reader()  # Sử dụng singleton QReader
        self.camera_cam = None
        self.qr_thread = None
        self.image_capture_cam = None
        self.front_img_path = None
        self.front_img_label = None
        self.back_img_path = None
        self.back_img_label = None
        self.front_image_temp = None
        self.back_image_temp = None

        # Đọc config webcam, nếu có
        self.qr_cam_index, self.index_camera_cam = load_camera_config()

        # Tạo giao diện
        self._createMenus()
        self._createCentralWidget()
        
        # Cập nhật hiển thị đường dẫn sau khi tạo UI
        self.update_paths_display()

        # Khởi động webcam sau khi giao diện đã được tạo
        QtCore.QTimer.singleShot(100, self.setup_cameras)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.viewfinder_cam and self.lbl_qr_preview:
            self.viewfinder_cam.setFixedSize(self.lbl_qr_preview.size())

    def _createMenus(self):
        menubar = self.menuBar()
        menubar.setNativeMenuBar(False)
        # Cài đặt
        self.action_caidat = QtWidgets.QAction("Cài đặt", self)
        self.action_caidat.triggered.connect(self.on_open_settings)
        menubar.addAction(self.action_caidat)

        # DS Công Dân
        self.action_open_excel = QtWidgets.QAction("DS Công dân", self)
        self.action_open_excel.triggered.connect(self.open_excel_file)
        menubar.addAction(self.action_open_excel)

    def _createCentralWidget(self):
        central = QtWidgets.QWidget()
        central.setStyleSheet("background-color: #E6F5FF;")
        self.setCentralWidget(central)
        main_layout = QtWidgets.QHBoxLayout(central)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # ------------------- Cột bên trái -------------------
        left_widget = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)

        # Frame hiển thị preview QR
        self.frame_qr = QtWidgets.QFrame()
        self.frame_qr.setMinimumSize(320, 180)
        self.frame_qr.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.frame_qr.setStyleSheet("""
            QFrame {
                background-color: #F0F7FF;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.lbl_qr_preview = QtWidgets.QLabel(self.frame_qr)
        self.lbl_qr_preview.setMinimumSize(360, 180)
        self.lbl_qr_preview.setStyleSheet("background-color: #000000;")
        self.lbl_qr_preview.setAlignment(QtCore.Qt.AlignCenter)
        self.btn_select_file = QtWidgets.QToolButton()
        folder_icon = self.style().standardIcon(QtWidgets.QStyle.SP_DirOpenIcon)
        self.btn_select_file.setIcon(folder_icon)
        self.btn_select_file.setToolTip("Chọn ảnh CCCD từ máy tính")
        self.btn_select_file.setFixedSize(28, 28)
        self.btn_select_file.clicked.connect(self.select_cccd_image)

        h_qr_row = QtWidgets.QHBoxLayout()
        h_qr_row.addWidget(self.frame_qr, stretch=1)

        # Tạo layout phụ để canh nút xuống 1px
        btn_layout = QtWidgets.QVBoxLayout()
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(0)
        btn_layout.addSpacing(300)  # Dịch xuống 1px
        btn_layout.addWidget(self.btn_select_file, alignment=QtCore.Qt.AlignTop)
        h_qr_row.addLayout(btn_layout)

        # ✅ THÊM DÒNG NÀY ĐỂ HIỂN THỊ
        left_layout.addLayout(h_qr_row)

        lbl_qr_text = QtWidgets.QLabel("Webcam đọc QR")
        lbl_qr_text.setAlignment(QtCore.Qt.AlignCenter)
        lbl_qr_text.setStyleSheet("color: black; font-weight: bold; font-size: 16px;")
        left_layout.addWidget(lbl_qr_text)

        # Frame hiển thị viewfinder cho chụp ảnh
        self.frame_cam = QtWidgets.QFrame()
        self.frame_cam.setMinimumSize(320, 180)  # đặt kích thước tối thiểu hợp lý
        self.frame_cam.setMaximumWidth(348)
        self.frame_cam.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.frame_cam.setStyleSheet("""
            QFrame {
                background-color: #F0F7FF;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)

        self.viewfinder_cam = QtMultimediaWidgets.QVideoWidget(self.frame_cam)
        self.viewfinder_cam.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        vf_layout = QtWidgets.QVBoxLayout(self.frame_cam)
        vf_layout.setContentsMargins(0, 0, 0, 0)
        vf_layout.setSpacing(0)
        vf_layout.addWidget(self.viewfinder_cam)
        lbl_cam_text = QtWidgets.QLabel("Webcam chụp mặt trước/sau")
        lbl_cam_text.setAlignment(QtCore.Qt.AlignCenter)
        lbl_cam_text.setStyleSheet("color: black; font-weight: bold; font-size: 16px;")
        left_layout.addWidget(self.frame_cam)
        left_layout.addWidget(lbl_cam_text)
        left_layout.addSpacing(20)

        # Các widget chọn ảnh mặt trước, mặt sau
        self.img_front_widget = self._create_image_widget("Ảnh mặt trước", is_front=True)
        self.img_back_widget = self._create_image_widget("Ảnh mặt sau", is_back=True)
        left_layout.addWidget(self.img_front_widget)
        left_layout.addWidget(self.img_back_widget)
        left_layout.addStretch()
        main_layout.addWidget(left_widget, stretch=2)

        # ------------------- Cột giữa: FORM -------------------
        # Đưa form_container vào scroll area
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        form_container = QtWidgets.QWidget()
        form_layout_v = QtWidgets.QVBoxLayout(form_container)
        form_layout_v.setContentsMargins(10, 10, 10, 10)
        form_layout_v.setSpacing(10)

        # Thêm widget hiển thị đường dẫn
        paths_group = QtWidgets.QGroupBox("Đường dẫn lưu trữ")
        paths_group.setStyleSheet("""
            QGroupBox {
                font-size: 16px;
                font-weight: bold;
                color: #005BEA;
                border: 2px solid #B0C4DE;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        paths_layout = QtWidgets.QVBoxLayout(paths_group)
        
        # Widget cho đường dẫn Excel
        excel_path_widget = QtWidgets.QWidget()
        excel_layout = QtWidgets.QHBoxLayout(excel_path_widget)
        excel_layout.setContentsMargins(0, 0, 0, 0)
        excel_label = QtWidgets.QLabel("File Excel:")
        excel_label.setStyleSheet("font-weight: bold; color: black; font-size: 14px;")
        excel_label.setFixedWidth(80)
        self.excel_path_display = QtWidgets.QLineEdit()
        self.excel_path_display.setReadOnly(True)
        self.excel_path_display.setStyleSheet("""
            QLineEdit {
                background-color: #F7F7F7;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                padding: 4px;
                color: #333333;
            }
        """)
        excel_open_btn = QtWidgets.QPushButton("Mở")
        excel_open_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 4px;
                padding: 4px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        excel_open_btn.clicked.connect(self.open_excel_file)
        excel_layout.addWidget(excel_label)
        excel_layout.addWidget(self.excel_path_display)
        excel_layout.addWidget(excel_open_btn)
        
        # Widget cho đường dẫn folder ảnh
        image_path_widget = QtWidgets.QWidget()
        image_layout = QtWidgets.QHBoxLayout(image_path_widget)
        image_layout.setContentsMargins(0, 0, 0, 0)
        image_label = QtWidgets.QLabel("Folder ảnh:")
        image_label.setStyleSheet("font-weight: bold; color: black; font-size: 14px;")
        image_label.setFixedWidth(80)
        self.image_path_display = QtWidgets.QLineEdit()
        self.image_path_display.setReadOnly(True)
        self.image_path_display.setStyleSheet("""
            QLineEdit {
                background-color: #F7F7F7;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                padding: 4px;
                color: #333333;
            }
        """)
        image_open_btn = QtWidgets.QPushButton("Mở")
        image_open_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 4px;
                padding: 4px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        image_open_btn.clicked.connect(self.open_image_folder)
        image_layout.addWidget(image_label)
        image_layout.addWidget(self.image_path_display)
        image_layout.addWidget(image_open_btn)
        
        # Thêm các widget vào group
        paths_layout.addWidget(excel_path_widget)
        paths_layout.addWidget(image_path_widget)
        
        # Thêm group vào form
        form_layout_v.addWidget(paths_group)
        form_layout_v.addSpacing(10)

        self.fields = {}

        label_style = "color: black; font-size: 16px; font-family: 'Segoe UI', 'Arial', sans-serif; font-weight: 500;"
        input_style = """
            background-color: #FFFFFF;
            color: #222222;
            font-size: 16px;
            border-radius: 10px;
            border: 1.5px solid #B0C4DE;
            padding-left: 10px;
            padding-right: 10px;
        """
        radio_style = "font-size: 16px; color: black;"

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignLeft)
        form.setFormAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        form.setHorizontalSpacing(20)
        form.setVerticalSpacing(18)

        # Số giấy tờ
        lbl_so_giay_to = QtWidgets.QLabel("Số giấy tờ:")
        lbl_so_giay_to.setStyleSheet(label_style)
        self.edt_so_giay_to = QtWidgets.QLineEdit()
        self.edt_so_giay_to.setFixedHeight(32)
        self.edt_so_giay_to.setStyleSheet(input_style)
        form.addRow(lbl_so_giay_to, self.edt_so_giay_to)
        self.fields["Số giấy tờ"] = self.edt_so_giay_to

        # Số CMND cũ (nếu có)
        lbl_cmnd_cu = QtWidgets.QLabel("Số CMND cũ (nếu có):")
        lbl_cmnd_cu.setStyleSheet(label_style)
        self.edt_cmnd_cu = QtWidgets.QLineEdit()
        self.edt_cmnd_cu.setFixedHeight(32)
        self.edt_cmnd_cu.setStyleSheet(input_style)
        form.addRow(lbl_cmnd_cu, self.edt_cmnd_cu)
        self.fields["Số CMND cũ (nếu có)"] = self.edt_cmnd_cu

        # Họ và tên
        lbl_hoten = QtWidgets.QLabel("Họ và tên:")
        lbl_hoten.setStyleSheet(label_style)
        self.edt_hoten = QtWidgets.QLineEdit()
        self.edt_hoten.setFixedHeight(32)
        self.edt_hoten.setStyleSheet(input_style)
        form.addRow(lbl_hoten, self.edt_hoten)
        self.fields["Họ và tên"] = self.edt_hoten

        # Ngày sinh
        lbl_ns = QtWidgets.QLabel("Ngày sinh:")
        lbl_ns.setStyleSheet(label_style)
        self.date_ns = FloatingCalendarDateEdit()
        self.date_ns.setDate(QtCore.QDate.currentDate())
        self.date_ns.setFixedHeight(32)
        self.date_ns.setStyleSheet(input_style)
        form.addRow(lbl_ns, self.date_ns)
        self.fields["Ngày sinh"] = self.date_ns

        # Giới tính
        lbl_gt = QtWidgets.QLabel("Giới tính:")
        lbl_gt.setStyleSheet(label_style)
        gender_widget = QtWidgets.QWidget()
        h_gender = QtWidgets.QHBoxLayout(gender_widget)
        h_gender.setContentsMargins(0, 0, 0, 0)
        h_gender.setSpacing(15)
        self.gender_group = QtWidgets.QButtonGroup(self)
        self.rb_nam = QtWidgets.QRadioButton("Nam")
        self.rb_nam.setStyleSheet(radio_style)
        self.rb_nu = QtWidgets.QRadioButton("Nữ")
        self.rb_nu.setStyleSheet(radio_style)
        h_gender.addWidget(self.rb_nam)
        h_gender.addWidget(self.rb_nu)
        self.gender_group.addButton(self.rb_nam)
        self.gender_group.addButton(self.rb_nu)
        form.addRow(lbl_gt, gender_widget)

        # Nơi thường trú
        lbl_noi_thuong_tru = QtWidgets.QLabel("Nơi thường trú:")
        lbl_noi_thuong_tru.setStyleSheet(label_style)
        self.edt_noi_thuong_tru = QtWidgets.QLineEdit()
        self.edt_noi_thuong_tru.setFixedHeight(32)
        self.edt_noi_thuong_tru.setStyleSheet(input_style)
        form.addRow(lbl_noi_thuong_tru, self.edt_noi_thuong_tru)
        self.fields["Nơi thường trú"] = self.edt_noi_thuong_tru

        # Ngày cấp giấy tờ
        lbl_ncgt = QtWidgets.QLabel("Ngày cấp giấy tờ:")
        lbl_ncgt.setStyleSheet(label_style)
        self.date_cap = FloatingCalendarDateEdit()
        self.date_cap.setDate(QtCore.QDate.currentDate())
        self.date_cap.setFixedHeight(32)
        self.date_cap.setStyleSheet(input_style)
        form.addRow(lbl_ncgt, self.date_cap)
        self.fields["Ngày cấp giấy tờ"] = self.date_cap

        # Loại giấy tờ
        lbl_loai_gt = QtWidgets.QLabel("Loại giấy tờ:")
        lbl_loai_gt.setStyleSheet(label_style)
        self.edt_loai_giay_to = QtWidgets.QLineEdit()
        self.edt_loai_giay_to.setFixedHeight(32)
        self.edt_loai_giay_to.setStyleSheet(input_style)
        self.edt_loai_giay_to.setText("CCCD")
        form.addRow(lbl_loai_gt, self.edt_loai_giay_to)
        self.fields["Loại giấy tờ"] = self.edt_loai_giay_to

        # Tên phòng lưu trú
        lbl_phong = QtWidgets.QLabel("Tên phòng lưu trú:")
        lbl_phong.setStyleSheet(label_style)
        self.cmb_phong = QtWidgets.QComboBox()
        self.cmb_phong.setFixedHeight(32)
        self.cmb_phong.setStyleSheet(input_style)
        ds_phong = [
            "", "Phòng 3 nhà cũ", "Phòng 4 nhà cũ", "Phòng 5 nhà cũ",
            "Phòng 7 nhà cũ", "Phòng 8 nhà cũ", "Phòng 9 nhà cũ",
            "Phòng 1 nhà mới", "Phòng 2 nhà mới", "Phòng 3 nhà mới",
            "Phòng 4 nhà mới", "Phòng 5 nhà mới"
        ]
        self.cmb_phong.addItems(ds_phong)
        form.addRow(lbl_phong, self.cmb_phong)
        self.fields["Tên phòng lưu trú"] = self.cmb_phong

        # Label status
        self.status_label = QtWidgets.QLabel("")
        self.status_label.setStyleSheet("color: green; font-size: 13px;")
        form_layout_v.addWidget(self.status_label)
        form_layout_v.addLayout(form)
        form_layout_v.addStretch()

        # Nút Lưu và Xóa thông tin
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_save = QtWidgets.QPushButton("Lưu")
        btn_save.setFixedHeight(40)
        btn_save.setFixedWidth(120)
        btn_save.setStyleSheet("""
            QPushButton {
                background-color: #005BEA;
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 15px;
            }
            QPushButton:hover {
                background-color: #0047BA;
            }
        """)
        btn_save.clicked.connect(self.write_to_excel)

        btn_clear = QtWidgets.QPushButton("Xóa")
        btn_clear.setFixedHeight(40)
        btn_clear.setFixedWidth(120)
        btn_clear.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FF7A7A, stop:1 #FF1C1C);  /* Sáng phía trên, đậm phía dưới */
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 15px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FF9999, stop:1 #E60000);
            }
        """)
        btn_clear.clicked.connect(self.clear_all_fields)
        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_clear)
        btn_row.addStretch()
        form_layout_v.addLayout(btn_row)
        scroll.setWidget(form_container)
        main_layout.addWidget(scroll, stretch=3)

        # ------------------- Cột bên phải (chỉ là background blue) -------------------
        right_panel = QtWidgets.QWidget()
        right_panel.setMinimumWidth(100)
        right_panel.setStyleSheet("background-color: #3399FF;")
        main_layout.addWidget(right_panel, stretch=1)

    def _create_image_widget(self, label_text: str, is_front=False, is_back=False) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        v_main = QtWidgets.QVBoxLayout(widget)
        v_main.setContentsMargins(0, 0, 0, 0)
        v_main.setSpacing(5)
        top_row = QtWidgets.QHBoxLayout()
        top_row.setSpacing(5)
        lbl = QtWidgets.QLabel(label_text)
        lbl.setFixedWidth(120)
        lbl.setStyleSheet("color: black; font-size: 15px; font-weight: 500;")
        top_row.addWidget(lbl)
        edt_path = QtWidgets.QLineEdit()
        edt_path.setFixedHeight(28)
        edt_path.setStyleSheet("""
            background-color: #FFFFFF;
            color: #222222;
            font-size: 15px;
            border-radius: 10px;
            border: 1.2px solid #B0C4DE;
            padding-left: 8px;
            padding-right: 8px;
        """)
        top_row.addWidget(edt_path, stretch=1)
        btn_folder = QtWidgets.QToolButton()
        btn_folder.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DirOpenIcon))
        btn_folder.setToolTip(f"Chọn file {label_text}")
        btn_camera = QtWidgets.QToolButton()
        btn_camera.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DesktopIcon))
        btn_camera.setToolTip(f"Chụp ảnh {label_text}")
        top_row.addWidget(btn_folder)
        top_row.addWidget(btn_camera)
        v_main.addLayout(top_row)
        lbl_image = QtWidgets.QLabel()
        lbl_image.setFixedSize(250, 160)  # bạn có thể điều chỉnh 250 và 160 tùy mong muốn
        lbl_image.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        lbl_image.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        lbl_image.setStyleSheet("""
            QLabel {
                background-color: #FFFFFF;
                border: 1px solid #CCCCCC;
                border-radius: 0px;
            }
        """)
        v_main.addWidget(lbl_image, alignment=QtCore.Qt.AlignLeft)
        v_main.addSpacing(5)

        if is_front:
            self.front_img_path = edt_path
            self.front_img_label = lbl_image
            btn_folder.clicked.connect(self.select_front_image_from_file)
            btn_camera.clicked.connect(self.capture_front_image_from_camera)
        if is_back:
            self.back_img_path = edt_path
            self.back_img_label = lbl_image
            btn_folder.clicked.connect(self.select_back_image_from_file)
            btn_camera.clicked.connect(self.capture_back_image_from_camera)
        return widget

    def select_front_image_from_file(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn ảnh mặt trước", "", "Image Files (*.jpg *.jpeg *.png)")
        if not file_path:
            return
        self.front_img_path.setText(file_path)
        img = cv2.imread(file_path)
        if img is not None:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            h, w, ch = img.shape
            bytes_per_line = ch * w
            qt_img = QtGui.QImage(img.data, w, h, bytes_per_line, QtGui.QImage.Format.Format_RGB888)
            pixmap = QtGui.QPixmap.fromImage(qt_img).scaled(
                self.front_img_label.width(),
                self.front_img_label.height(),
                QtCore.Qt.IgnoreAspectRatio,
                QtCore.Qt.SmoothTransformation
            )
            self.front_img_label.setPixmap(pixmap)

    def select_back_image_from_file(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn ảnh mặt sau", "", "Image Files (*.jpg *.jpeg *.png)")
        if not file_path:
            return
        self.back_img_path.setText(file_path)
        img = cv2.imread(file_path)
        if img is not None:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            h, w, ch = img.shape
            bytes_per_line = ch * w
            qt_img = QtGui.QImage(img.data, w, h, bytes_per_line, QtGui.QImage.Format.Format_RGB888)
            pixmap = QtGui.QPixmap.fromImage(qt_img).scaled(
                self.back_img_label.width(),
                self.back_img_label.height(),
                QtCore.Qt.IgnoreAspectRatio,
                QtCore.Qt.SmoothTransformation
            )
            self.back_img_label.setPixmap(pixmap)

    def capture_front_image_from_camera(self):
        if not hasattr(self, "image_capture_cam") or self.image_capture_cam is None or self.camera_cam is None:
            QtWidgets.QMessageBox.warning(self, "Lỗi", "Chưa cấu hình webcam chụp mặt trước/sau.")
            return
        try:
            self.image_capture_cam.imageCaptured.disconnect()
        except Exception:
            pass
        self.image_capture_cam.imageCaptured.connect(self.on_front_image_captured)
        self.image_capture_cam.capture()

    def on_front_image_captured(self, id, image):
        pixmap = QtGui.QPixmap.fromImage(image).scaled(
            self.front_img_label.width(),
            self.front_img_label.height(),
            QtCore.Qt.IgnoreAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        self.front_img_label.setPixmap(pixmap)
        self.front_image_temp = image.copy()  # Lưu ảnh tạm trong bộ nhớ
        
        try:
            self.image_capture_cam.imageCaptured.disconnect()
        except Exception:
            pass

    def capture_back_image_from_camera(self):
        if not hasattr(self, "image_capture_cam") or self.image_capture_cam is None or self.camera_cam is None:
            QtWidgets.QMessageBox.warning(self, "Lỗi", "Chưa cấu hình webcam chụp mặt trước/sau.")
            return
        try:
            self.image_capture_cam.imageCaptured.disconnect()
        except Exception:
            pass
        self.image_capture_cam.imageCaptured.connect(self.on_back_image_captured)
        self.image_capture_cam.capture()

    def on_back_image_captured(self, id, image):
        pixmap = QtGui.QPixmap.fromImage(image).scaled(
            self.back_img_label.width(),
            self.back_img_label.height(),
            QtCore.Qt.IgnoreAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        self.back_img_label.setPixmap(pixmap)
        self.back_image_temp = image.copy()  # Lưu ảnh tạm trong bộ nhớ
        
        try:
            self.image_capture_cam.imageCaptured.disconnect()
        except Exception:
            pass

    def save_temp_image(self, qimage, is_front=True):
        try:
            if qimage:
                current_dir = self.app_dir
                folder_anh = os.path.join(current_dir, "DS_Anh_Cong_Dan_da_khai_bao")
                os.makedirs(folder_anh, exist_ok=True)
                
                # Lấy họ tên và ngày sinh
                ho_ten = self.fields["Họ và tên"].text().strip()
                ngay_sinh = self.fields["Ngày sinh"].date()
                
                # Chuẩn hóa họ tên: thay khoảng trắng bằng dấu gạch dưới và loại bỏ dấu
                ho_ten = unidecode(ho_ten).replace(" ", "_")
                
                # Tạo tên file
                side = "mat_truoc" if is_front else "mat_sau"
                if ho_ten and ngay_sinh:
                    filename = f"{side}_{ho_ten}_{ngay_sinh.day():02d}_{ngay_sinh.month():02d}_{ngay_sinh.year()}.jpg"
                else:
                    # Nếu không có họ tên hoặc ngày sinh thì dùng timestamp
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{side}_{timestamp}.jpg"
                
                filepath = os.path.join(folder_anh, filename)
                
                # Nếu file đã tồn tại, thêm số thứ tự vào tên file
                base_name = os.path.splitext(filename)[0]
                ext = os.path.splitext(filename)[1]
                counter = 1
                while os.path.exists(filepath):
                    filename = f"{base_name}_{counter}{ext}"
                    filepath = os.path.join(folder_anh, filename)
                    counter += 1
                
                # Lưu ảnh
                qimage.save(filepath, "JPG", quality=95)
                return filepath
        except Exception as e:
            print(f"Lỗi khi lưu ảnh: {e}")
        return ""

    def on_open_settings(self):
        dlg = OptionDialog(
            parent=self,
            current_qr_idx=self.qr_cam_index if hasattr(self, "qr_cam_index") else -1,
            current_cam_idx=self.index_camera_cam
        )
        result = dlg.exec_()
        if result == QtWidgets.QDialog.Accepted:
            idx_qr, idx_cam = dlg.get_selected_camera_indexes()
            self.qr_cam_index = idx_qr
            self.index_camera_cam = idx_cam
            save_camera_config(self.qr_cam_index, self.index_camera_cam)
            self.setup_cameras()

    def setup_cameras(self):
        # Nếu đã khởi tạo trước đó, dừng và giải phóng
        if hasattr(self, "cap_qr") and self.cap_qr is not None:
            self.timer_cv.stop()
            self.cap_qr.release()
            self.cap_qr = None
            self.lbl_qr_preview.clear()
            self.lbl_qr_preview.setStyleSheet("background-color: #000000;")
            self.frame_qr.setStyleSheet("""
                QFrame {
                    background-color: #F0F7FF;
                    border: 1px solid #CCCCCC;
                    border-radius: 4px;
                }
            """)

        if self.camera_cam is not None:
            self.camera_cam.stop()
            self.camera_cam.deleteLater()
            self.camera_cam = None
            self.viewfinder_cam.hide()
            self.frame_cam.setStyleSheet("""
                QFrame {
                    background-color: #F0F7FF;
                    border: 1px solid #CCCCCC;
                    border-radius: 4px;
                }
            """)

        if self.qr_thread is not None:
            self.qr_thread.stop()
            self.qr_thread = None

        cameras_info = QtMultimedia.QCameraInfo.availableCameras()
        # Thiết lập webcam đọc QR
        if hasattr(self, "qr_cam_index") and 0 <= self.qr_cam_index < len(cameras_info):
            self.cap_qr = cv2.VideoCapture(self.qr_cam_index, cv2.CAP_DSHOW)
            self.cap_qr.set(cv2.CAP_PROP_FRAME_WIDTH, 600)
            self.cap_qr.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            self.frame_count = 0
            self.timer_cv = QtCore.QTimer(self)
            self.timer_cv.timeout.connect(self.read_frame_for_qr)
            self.timer_cv.start(30)
            self.qr_thread = QRDecodeThread(self.qreader)
            self.qr_thread.qrDecoded.connect(self.on_qr_decoded)
            self.qr_thread.start()

        # Thiết lập webcam chụp ảnh (mặt trước/sau)
        if 0 <= self.index_camera_cam < len(cameras_info):
            cam_info_cam = cameras_info[self.index_camera_cam]
            self.camera_cam = QtMultimedia.QCamera(cam_info_cam)
            
            # Thêm cấu hình camera
            settings = QtMultimedia.QCameraViewfinderSettings()
            settings.setResolution(640, 480)
            settings.setMinimumFrameRate(30.0)
            settings.setMaximumFrameRate(30.0)
            self.camera_cam.setViewfinderSettings(settings)
            
            self.camera_cam.setViewfinder(self.viewfinder_cam)
            self.viewfinder_cam.setAspectRatioMode(QtCore.Qt.KeepAspectRatioByExpanding)
            self.viewfinder_cam.show()
            self.viewfinder_cam.updateGeometry()
            self.frame_cam.setStyleSheet("""
                QFrame {
                    border: 1px solid #CCCCCC;
                }
            """)
            self.camera_cam.start()
            self.image_capture_cam = QCameraImageCapture(self.camera_cam)

    def read_frame_for_qr(self):
        if self.cap_qr is None or not self.cap_qr.isOpened():
            return
        ret, frame = self.cap_qr.read()
        if not ret:
            return
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb_frame.shape
        bytes_per_line = ch * w
        qt_img = QtGui.QImage(rgb_frame.data, w, h, bytes_per_line, QtGui.QImage.Format.Format_RGB888)
        pixmap = QtGui.QPixmap.fromImage(qt_img).scaled(
            self.lbl_qr_preview.width(),
            self.lbl_qr_preview.height(),
            QtCore.Qt.IgnoreAspectRatio,
            QtCore.Qt.SmoothTransformation
        )

        self.lbl_qr_preview.setPixmap(pixmap)
        if self.qr_thread is not None and not self.qr_thread.isRunning():
            return
        if self.qr_thread is not None and self.frame_count % 5 == 0:
            self.qr_thread.request_decode(rgb_frame)
        self.frame_count = (self.frame_count + 1) % 30

    def _reset_can_decode(self):
        self.can_decode = True

    def on_qr_decoded(self, qr_text):
        if not self.can_decode:
            return  # Đang chờ, bỏ qua mọi quét

        info = parse_qr(qr_text)
        if not info:
            return

        # Xử lý bình thường
        self.fill_form_from_info(info)
        self._play_sound("done.wav")
        self.last_qr_text = qr_text

        # Khóa lại 3 giây
        self.can_decode = False
        QtCore.QTimer.singleShot(3000, self._reset_can_decode)

    def _reset_ignore_decode(self):
        self.ignore_decode = False

    def _reset_can_process_trung(self):
        self.can_process_trung = True

    def fill_form_from_info(self, info: dict):
        self.clear_status()
        self.fields["Số giấy tờ"].setText(info["Số giấy tờ"])
        self.fields["Số CMND cũ (nếu có)"].setText(info["Số CMND cũ (nếu có)"])
        self.fields["Họ và tên"].setText(info["Họ và tên"])
        try:
            d_ns = datetime.strptime(info["Ngày sinh"], "%d/%m/%Y")
            self.date_ns.setDate(QtCore.QDate(d_ns.year, d_ns.month, d_ns.day))
        except:
            self.date_ns.clear()
        gt = info["Giới tính"]
        if gt == "Nam":
            self.rb_nam.setChecked(True)
        elif gt == "Nữ":
            self.rb_nu.setChecked(True)
        else:
            self.gender_group.setExclusive(False)
            self.rb_nam.setChecked(False)
            self.rb_nu.setChecked(False)
            self.gender_group.setExclusive(True)
        self.fields["Nơi thường trú"].setText(info["Nơi thường trú"])
        try:
            d_cap = datetime.strptime(info["Ngày cấp giấy tờ"], "%d/%m/%Y")
            self.date_cap.setDate(QtCore.QDate(d_cap.year, d_cap.month, d_cap.day))
        except:
            self.date_cap.clear()
        self.fields["Loại giấy tờ"].setText("CCCD")
        self.status_label.setText("✅ Quét QR thành công, đã điền thông tin.")
        self.status_label.setStyleSheet("color: green;")

    def select_cccd_image(self):
        self.clear_status()
        self.status_label.setText("📥 Chọn ảnh CCCD có QR để quét...")
        QtWidgets.QApplication.processEvents()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn ảnh CCCD", "", "Image Files (*.jpg *.jpeg *.png)")
        if not file_path:
            self.clear_status()
            return
        image = cv2.imread(file_path)
        if image is None:
            self.status_label.setText("❌ Không đọc được ảnh. Hãy kiểm tra định dạng.")
            self.status_label.setStyleSheet("color: red;")
            return
        image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        decoded = self.qreader.detect_and_decode(image=image_rgb)
        if not decoded or not decoded[0]:
            self.status_label.setText("❌ Không quét được QR từ ảnh.")
            self.status_label.setStyleSheet("color: red;")
            self._play_sound("error.wav")
            return
        qr_text = decoded[0]
        info = parse_qr(qr_text)
        if not info:
            self.status_label.setText("❌ Dữ liệu QR không đúng định dạng.")
            self.status_label.setStyleSheet("color: red;")
            self._play_sound("error.wav")
            return
        self.fill_form_from_info(info)
        self._play_sound("done.wav")
        self.last_qr_text = ""  # reset lại để cho phép webcam quét lại mã QR tiếp theo

    def _play_sound(self, wav_filename: str):
        try:
            base_dir = self.app_dir
            wav_path = os.path.join(base_dir, wav_filename)
            if os.path.exists(wav_path):
                winsound.PlaySound(wav_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception as e:
            print(f"⚠️ Lỗi khi phát âm thanh '{wav_filename}': {e}")

    def clear_status(self):
        self.status_label.setText("")
        self.status_label.setStyleSheet("color: green;")

    def dong_file_excel_neu_dang_mo(self, file_path, excel_app=None):
        try:
            # Tìm tất cả các processes của Excel đang chạy
            for proc in psutil.process_iter(['name']):
                try:
                    if proc.name().lower() in ['excel.exe', 'EXCEL.EXE']:
                        # Tìm tất cả các cửa sổ Excel
                        def callback(hwnd, hwnds):
                            if win32gui.IsWindowVisible(hwnd):
                                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                                if pid == proc.pid:
                                    hwnds.append(hwnd)
                            return True
                        
                        hwnds = []
                        win32gui.EnumWindows(callback, hwnds)
                        
                        # Gửi thông điệp đóng đến từng cửa sổ
                        for hwnd in hwnds:
                            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                        
                        # Đợi process kết thúc
                        try:
                            proc.wait(timeout=2)
                        except psutil.TimeoutExpired:
                            proc.kill()  # Nếu quá thời gian, buộc đóng
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            # Khởi tạo Excel application mới
            excel_app = win32.gencache.EnsureDispatch("Excel.Application")
            return excel_app

        except Exception as e:
            print(f"Lỗi khi đóng Excel: {e}")
            return None

    def write_to_excel(self):
        try:
            current_dir = self.app_dir
            excel_path = os.path.join(current_dir, "DS_Cong_Dan.xlsx")
            folder_anh = os.path.join(current_dir, "DS_Anh_Cong_Dan_da_khai_bao")
            os.makedirs(folder_anh, exist_ok=True)

            # Lưu ảnh từ bộ nhớ tạm vào file nếu có
            front_path = ""
            back_path = ""
            if self.front_image_temp:
                front_path = self.save_temp_image(self.front_image_temp, is_front=True)
            if self.back_image_temp:
                back_path = self.save_temp_image(self.back_image_temp, is_front=False)

            # Khởi tạo Excel application mới
            excel = self.dong_file_excel_neu_dang_mo(excel_path)
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            if not os.path.exists(excel_path):
                wb = excel.Workbooks.Add()
                ws = wb.Sheets(1)
                headers = ["Số giấy tờ", "Số CMND cũ", "Họ và tên", "Giới tính", "Ngày sinh",
                        "Nơi thường trú", "Ngày cấp", "Loại giấy tờ", "Tên phòng", "Thời gian ghi",
                        "Ảnh mặt trước", "Ảnh mặt sau"]
                rng = ws.Range(ws.Cells(1, 1), ws.Cells(1, len(headers)))
                rng.Value = headers
                rng.Font.Name = "Times New Roman"
                rng.Font.Size = 16
                rng.Font.Bold = True
                rng.Interior.Color = 0x00C0FF
                rng.HorizontalAlignment = -4108

                # Định dạng tất cả các cột ngày tháng là text
                ws.Range("E:E").NumberFormat = "@"  # Ngày sinh
                ws.Range("G:G").NumberFormat = "@"  # Ngày cấp
                ws.Range("J:J").NumberFormat = "@"  # Thời gian ghi

                # Tự động điều chỉnh độ rộng cột
                ws.Range("A:L").EntireColumn.AutoFit()

                wb.SaveAs(excel_path)
                wb.Close()

            wb = excel.Workbooks.Open(excel_path)
            ws = wb.Sheets(1)
            
            # Tìm dòng cuối cùng một cách chính xác hơn
            last_row = 1  # Bắt đầu từ dòng 1
            while True:
                if not ws.Cells(last_row, 1).Value:  # Nếu ô A{last_row} trống
                    # Kiểm tra thêm các cột khác để đảm bảo dòng thực sự trống
                    is_empty = True
                    for col in range(1, 13):  # Kiểm tra từ cột 1 đến 12
                        if ws.Cells(last_row, col).Value:
                            is_empty = False
                            break
                    if is_empty:
                        break
                last_row += 1

            # Nếu là dòng đầu tiên (chỉ có header), ghi từ dòng 2
            if last_row == 1:
                last_row = 2

            # Lấy dữ liệu từ các trường, nếu trống thì để chuỗi rỗng
            # Lấy thời gian hiện tại từ hệ thống
            t = time.localtime()
            current_time = f"{t.tm_mday:02d}/{t.tm_mon:02d}/{t.tm_year} {t.tm_hour:02d}:{t.tm_min:02d}"
            
            # Lấy ngày sinh và ngày cấp như người dùng nhập
            ngay_sinh = self.fields["Ngày sinh"].date()
            ngay_cap = self.fields["Ngày cấp giấy tờ"].date()
            
            data = [
                self.fields["Số giấy tờ"].text().strip(),
                self.fields["Số CMND cũ (nếu có)"].text().strip(),
                self.fields["Họ và tên"].text().strip(),
                "Nam" if self.rb_nam.isChecked() else ("Nữ" if self.rb_nu.isChecked() else ""),
                f"{ngay_sinh.day():02d}/{ngay_sinh.month():02d}/{ngay_sinh.year()}" if ngay_sinh else "",
                self.fields["Nơi thường trú"].text().strip(),
                f"{ngay_cap.day():02d}/{ngay_cap.month():02d}/{ngay_cap.year()}" if ngay_cap else "",
                self.fields["Loại giấy tờ"].text().strip(),
                self.cmb_phong.currentText().strip(),
                current_time,  # Thời gian ghi đã được định dạng sẵn
                "Ảnh mặt trước" if front_path else "",  # Chỉ hiển thị text, hyperlink sẽ được thêm sau
                "Ảnh mặt sau" if back_path else ""      # Chỉ hiển thị text, hyperlink sẽ được thêm sau
            ]

            # Ghi dữ liệu một lần
            rng = ws.Range(ws.Cells(last_row, 1), ws.Cells(last_row, len(data)))
            rng.Value = [data]
            
            # Định dạng
            rng.Font.Size = 11
            rng.Font.Name = "Times New Roman"
            rng.HorizontalAlignment = -4108

            # Định dạng cột thời gian ghi (cột J) là text để giữ nguyên định dạng
            ws.Range(f"J{last_row}").NumberFormat = "@"
            
            # Định dạng cột ngày sinh và ngày cấp là text
            ws.Range(f"E{last_row}").NumberFormat = "@"
            ws.Range(f"G{last_row}").NumberFormat = "@"

            # Thêm hyperlink cho ảnh nếu có
            if front_path and os.path.exists(front_path):
                # Chuyển đổi đường dẫn thành đường dẫn tuyệt đối và chuẩn hóa
                abs_front_path = os.path.abspath(front_path)
                file_url = f"file:///{abs_front_path.replace(os.sep, '/')}"
                ws.Hyperlinks.Add(
                    Anchor=ws.Range(f"K{last_row}"),
                    Address=file_url,
                    TextToDisplay="Ảnh mặt trước"
                )

            if back_path and os.path.exists(back_path):
                # Chuyển đổi đường dẫn thành đường dẫn tuyệt đối và chuẩn hóa
                abs_back_path = os.path.abspath(back_path)
                file_url = f"file:///{abs_back_path.replace(os.sep, '/')}"
                ws.Hyperlinks.Add(
                    Anchor=ws.Range(f"L{last_row}"),
                    Address=file_url,
                    TextToDisplay="Ảnh mặt sau"
                )

            # AutoFit và border
            ws.Range("A:L").EntireColumn.AutoFit()
            rng = ws.Range(f"A1:L{last_row}")
            for i in range(7, 13):
                rng.Borders(i).LineStyle = 1
                rng.Borders(i).Weight = 2

            # Bật lại các tính năng Excel trước khi lưu
            excel.ScreenUpdating = True
            excel.EnableEvents = True
            
            wb.Save()
            wb.Close()

            self.front_image_temp = None
            self.back_image_temp = None
            self.clear_all_fields()
            self.show_success_message(f"✅ Đã lưu công dân tại dòng {last_row} vào DS_Cong_Dan.xlsx")

        except Exception as e:
            # Đảm bảo bật lại các tính năng Excel trong trường hợp lỗi
            try:
                if excel:
                    excel.ScreenUpdating = True
                    excel.EnableEvents = True
            except:
                pass
            QtWidgets.QMessageBox.critical(self, "Lỗi", f"Không thể ghi Excel: {e}")

    def open_excel_file(self):
        try:
            excel_path = os.path.join(self.app_dir, "DS_Cong_Dan.xlsx")
            if os.path.exists(excel_path):
                os.startfile(excel_path)
            else:
                QtWidgets.QMessageBox.warning(self, "Không tìm thấy file", "Chưa có dữ liệu khai báo hoặc file chưa được tạo.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", f"Không thể mở file Excel: {e}")

    def clear_all_fields(self):
        for field in self.fields.values():
            if isinstance(field, QtWidgets.QLineEdit):
                field.clear()
            elif isinstance(field, QtWidgets.QComboBox):
                field.setCurrentIndex(0)
            elif isinstance(field, QtWidgets.QDateEdit):
                field.setDate(QtCore.QDate.currentDate())
        self.rb_nam.setChecked(False)
        self.rb_nu.setChecked(False)
        if self.front_img_label:
            self.front_img_label.clear()
        if self.back_img_label:
            self.back_img_label.clear()
        self.front_image_temp = None
        self.back_image_temp = None
        self.clear_status()

    def show_success_message(self, message: str):
        self.status_label.setText(message)
        self.status_label.setStyleSheet("color: blue; font-size: 14px; font-weight: bold;")

    def update_paths_display(self):
        """Cập nhật hiển thị đường dẫn"""
        if hasattr(self, 'excel_path_display'):
            self.excel_path_display.setText(self.excel_path)
        if hasattr(self, 'image_path_display'):
            self.image_path_display.setText(self.image_folder)

    def open_image_folder(self):
        """Mở folder chứa ảnh bằng Windows Explorer"""
        try:
            if not os.path.exists(self.image_folder):
                os.makedirs(self.image_folder)
            os.startfile(self.image_folder)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Lỗi", f"Không thể mở thư mục ảnh: {e}")

if __name__ == "__main__":
    try:
        # Tắt thông báo lỗi của OpenCV
        os.environ["OPENCV_LOG_LEVEL"] = "OFF"
        cv2.setLogLevel(0)
        
        # Khởi tạo QApplication với các tùy chọn tối ưu
        app = QtWidgets.QApplication(sys.argv)
        app.setStyle('Fusion')  # Style nhẹ và nhanh
        
        # Tắt các hiệu ứng không cần thiết
        app.setEffectEnabled(QtCore.Qt.UI_AnimateCombo, False)
        app.setEffectEnabled(QtCore.Qt.UI_AnimateTooltip, False)
        
        # Khởi tạo cửa sổ chính
        window = MainWindow()
        window.show()
        
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Lỗi khởi động: {e}")
        sys.exit(1)
