import sys
import os
import cv2
import numpy as np
from datetime import datetime
from unidecode import unidecode
import json

from PyQt5 import QtCore, QtGui, QtWidgets, QtMultimedia, QtMultimediaWidgets
from PyQt5.QtMultimedia import QCameraImageCapture

from qreader import QReader
import winsound

import time

import win32com.client as win32
import shutil

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
        self.setWindowTitle("Option")
        self.setModal(True)
        self.resize(500, 300)
        self.available_cameras = QtMultimedia.QCameraInfo.availableCameras()
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(20)
        grp = QtWidgets.QGroupBox("Cài đặt chung")
        vbox = QtWidgets.QVBoxLayout(grp)
        vbox.setContentsMargins(10, 10, 10, 10)
        vbox.setSpacing(15)
        h1 = QtWidgets.QHBoxLayout()
        lbl1 = QtWidgets.QLabel("Webcam đọc QR")
        lbl1.setFixedWidth(120)
        lbl1.setStyleSheet("color: black;")
        self.cmb_qr = QtWidgets.QComboBox()
        for cam in self.available_cameras:
            self.cmb_qr.addItem(cam.description())
        h1.addWidget(lbl1)
        h1.addWidget(self.cmb_qr)
        vbox.addLayout(h1)
        h2 = QtWidgets.QHBoxLayout()
        lbl2 = QtWidgets.QLabel("Webcam chụp ảnh")
        lbl2.setFixedWidth(120)
        lbl2.setStyleSheet("color: black;")
        self.cmb_cam = QtWidgets.QComboBox()
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
        btn_save = QtWidgets.QPushButton("Save")
        btn_save.setFixedSize(100, 35)
        btn_save.clicked.connect(self.accept)
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setFixedSize(100, 35)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

    def get_selected_camera_indexes(self):
        return self.cmb_qr.currentIndex(), self.cmb_cam.currentIndex()

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Kiểm soát ra/vào bằng CCCD 0.3.3")
        self.resize(1280, 900)
        self.setMinimumSize(1024, 700)
        self.cap_qr = None
        self.timer_cv = None
        self.last_qr_text = ""
        self.qreader = QReader()
        self.camera_cam = None
        self.qr_thread = None
        self.image_capture_cam = None
        self.front_img_path = None
        self.front_img_label = None
        self.back_img_path = None
        self.back_img_label = None
        self.qr_cam_index, self.index_camera_cam = load_camera_config()
        if getattr(sys, 'frozen', False):
            self.app_dir = sys._MEIPASS
        else:
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
        self.can_decode = True

        self._createMenus()
        self._createCentralWidget()
        self.setup_cameras()

        self.front_image_temp = None
        self.back_image_temp = None

    def _createMenus(self):
        menubar = self.menuBar()
        menubar.setNativeMenuBar(False)
        self.menu_ds_congdan = menubar.addMenu("DS Công dân")
        self.menu_ds_luotvao = menubar.addMenu("DS Lượt vào")
        self.menu_phan_nhom = menubar.addMenu("Phân nhóm")
        self.menu_caidat = menubar.addMenu("Cài đặt")
        self.menu_gioithieu = menubar.addMenu("Giới thiệu")
        self.menu_test = menubar.addMenu("Test")
        option_action = QtWidgets.QAction("Option…", self)
        option_action.triggered.connect(self.on_open_settings)
        self.menu_caidat.addAction(option_action)

    def _createCentralWidget(self):
        central = QtWidgets.QWidget()
        central.setStyleSheet("background-color: #E6F5FF;")
        self.setCentralWidget(central)
        main_layout = QtWidgets.QHBoxLayout(central)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # --- LEFT PANEL ---
        left_widget = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)
        # QR Webcam Frame
        self.frame_qr = QtWidgets.QFrame()
        self.frame_qr.setMinimumSize(200, 120)
        self.frame_qr.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.frame_qr.setStyleSheet("""
            QFrame {
                background-color: #F0F7FF;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.lbl_qr_preview = QtWidgets.QLabel(self.frame_qr)
        self.lbl_qr_preview.setMinimumSize(200, 120)
        self.lbl_qr_preview.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.lbl_qr_preview.setStyleSheet("background-color: #000000;")
        self.lbl_qr_preview.setAlignment(QtCore.Qt.AlignCenter)
        qr_layout = QtWidgets.QVBoxLayout(self.frame_qr)
        qr_layout.setContentsMargins(0, 0, 0, 0)
        qr_layout.addWidget(self.lbl_qr_preview)
        self.btn_select_file = QtWidgets.QToolButton()
        folder_icon = self.style().standardIcon(QtWidgets.QStyle.SP_DirOpenIcon)
        self.btn_select_file.setIcon(folder_icon)
        self.btn_select_file.setToolTip("Chọn ảnh CCCD từ máy tính")
        self.btn_select_file.setFixedSize(28, 28)
        self.btn_select_file.clicked.connect(self.select_cccd_image)
        h_qr_row = QtWidgets.QHBoxLayout()
        h_qr_row.addWidget(self.frame_qr)
        h_qr_row.addWidget(self.btn_select_file, alignment=QtCore.Qt.AlignBottom)
        left_layout.addLayout(h_qr_row)
        lbl_qr_text = QtWidgets.QLabel("Webcam đọc QR")
        lbl_qr_text.setAlignment(QtCore.Qt.AlignCenter)
        lbl_qr_text.setStyleSheet("color: black; font-weight: bold; font-size: 16px;")
        left_layout.addWidget(lbl_qr_text)
        left_layout.addSpacing(10)
        # Cam chụp mặt trước/sau
        self.frame_cam = QtWidgets.QFrame()
        self.frame_cam.setMinimumSize(200, 110)
        self.frame_cam.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.frame_cam.setStyleSheet("""
            QFrame {
                background-color: #F0F7FF;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.viewfinder_cam = QtMultimediaWidgets.QCameraViewfinder(self.frame_cam)
        layout = QtWidgets.QVBoxLayout(self.frame_cam)
        layout.setContentsMargins(0,0,0,0)
        layout.setSpacing(0)
        layout.addWidget(self.viewfinder_cam)
        lbl_cam_text = QtWidgets.QLabel("Webcam chụp mặt trước/sau")
        lbl_cam_text.setAlignment(QtCore.Qt.AlignCenter)
        lbl_cam_text.setStyleSheet("color: black; font-weight: bold; font-size: 16px;")
        left_layout.addWidget(self.frame_cam)
        left_layout.addWidget(lbl_cam_text)
        left_layout.addSpacing(10)
        self.img_front_widget = self._create_image_widget("Ảnh mặt trước", is_front=True)
        self.img_back_widget = self._create_image_widget("Ảnh mặt sau", is_back=True)
        left_layout.addWidget(self.img_front_widget)
        left_layout.addWidget(self.img_back_widget)
        left_layout.addStretch()
        main_layout.addWidget(left_widget, stretch=2)

        # --- FORM PANEL ---
        form_container = QtWidgets.QWidget()
        form_layout_v = QtWidgets.QVBoxLayout(form_container)
        form_layout_v.setContentsMargins(0, 0, 0, 0)
        form_layout_v.setSpacing(10)
        self.fields = {}
        label_style = "color: black; font-size: 16px; font-weight: 500;"
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
        form.setVerticalSpacing(15)

        lbl_so_giay_to = QtWidgets.QLabel("Số giấy tờ:")
        lbl_so_giay_to.setStyleSheet(label_style)
        self.edt_so_giay_to = QtWidgets.QLineEdit()
        self.edt_so_giay_to.setMinimumHeight(30)
        self.edt_so_giay_to.setStyleSheet(input_style)
        self.edt_so_giay_to.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        form.addRow(lbl_so_giay_to, self.edt_so_giay_to)
        self.fields["Số giấy tờ"] = self.edt_so_giay_to

        lbl_cmnd_cu = QtWidgets.QLabel("Số CMND cũ (nếu có):")
        lbl_cmnd_cu.setStyleSheet(label_style)
        self.edt_cmnd_cu = QtWidgets.QLineEdit()
        self.edt_cmnd_cu.setMinimumHeight(30)
        self.edt_cmnd_cu.setStyleSheet(input_style)
        self.edt_cmnd_cu.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        form.addRow(lbl_cmnd_cu, self.edt_cmnd_cu)
        self.fields["Số CMND cũ (nếu có)"] = self.edt_cmnd_cu

        lbl_hoten = QtWidgets.QLabel("Họ và tên:")
        lbl_hoten.setStyleSheet(label_style)
        self.edt_hoten = QtWidgets.QLineEdit()
        self.edt_hoten.setMinimumHeight(30)
        self.edt_hoten.setStyleSheet(input_style)
        self.edt_hoten.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        form.addRow(lbl_hoten, self.edt_hoten)
        self.fields["Họ và tên"] = self.edt_hoten

        lbl_ns = QtWidgets.QLabel("Ngày sinh:")
        lbl_ns.setStyleSheet(label_style)
        self.date_ns = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.date_ns.setDisplayFormat("dd/MM/yyyy")
        self.date_ns.setCalendarPopup(True)
        self.date_ns.setMinimumHeight(30)
        self.date_ns.setStyleSheet(input_style)
        self.date_ns.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        form.addRow(lbl_ns, self.date_ns)
        self.fields["Ngày sinh"] = self.date_ns

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

        lbl_noi_thuong_tru = QtWidgets.QLabel("Nơi thường trú:")
        lbl_noi_thuong_tru.setStyleSheet(label_style)
        self.edt_noi_thuong_tru = QtWidgets.QLineEdit()
        self.edt_noi_thuong_tru.setMinimumHeight(30)
        self.edt_noi_thuong_tru.setStyleSheet(input_style)
        self.edt_noi_thuong_tru.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        form.addRow(lbl_noi_thuong_tru, self.edt_noi_thuong_tru)
        self.fields["Nơi thường trú"] = self.edt_noi_thuong_tru

        lbl_ncgt = QtWidgets.QLabel("Ngày cấp giấy tờ:")
        lbl_ncgt.setStyleSheet(label_style)
        self.date_cap = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.date_cap.setDisplayFormat("dd/MM/yyyy")
        self.date_cap.setCalendarPopup(True)
        self.date_cap.setMinimumHeight(30)
        self.date_cap.setStyleSheet(input_style)
        self.date_cap.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        form.addRow(lbl_ncgt, self.date_cap)
        self.fields["Ngày cấp giấy tờ"] = self.date_cap

        lbl_loai_gt = QtWidgets.QLabel("Loại giấy tờ:")
        lbl_loai_gt.setStyleSheet(label_style)
        self.edt_loai_giay_to = QtWidgets.QLineEdit()
        self.edt_loai_giay_to.setMinimumHeight(30)
        self.edt_loai_giay_to.setStyleSheet(input_style)
        self.edt_loai_giay_to.setText("CCCD")
        self.edt_loai_giay_to.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        form.addRow(lbl_loai_gt, self.edt_loai_giay_to)
        self.fields["Loại giấy tờ"] = self.edt_loai_giay_to

        lbl_phong = QtWidgets.QLabel("Tên phòng lưu trú:")
        lbl_phong.setStyleSheet(label_style)
        self.cmb_phong = QtWidgets.QComboBox()
        self.cmb_phong.setMinimumHeight(30)
        self.cmb_phong.setStyleSheet(input_style)
        ds_phong = [
            "", "Phòng 3 nhà cũ", "Phòng 4 nhà cũ", "Phòng 5 nhà cũ",
            "Phòng 7 nhà cũ", "Phòng 8 nhà cũ", "Phòng 9 nhà cũ",
            "Phòng 1 nhà mới", "Phòng 2 nhà mới", "Phòng 3 nhà mới",
            "Phòng 4 nhà mới", "Phòng 5 nhà mới"
        ]
        self.cmb_phong.addItems(ds_phong)
        self.cmb_phong.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        form.addRow(lbl_phong, self.cmb_phong)
        self.fields["Tên phòng lưu trú"] = self.cmb_phong

        # Status
        self.status_label = QtWidgets.QLabel("")
        self.status_label.setStyleSheet("color: green; font-size: 13px;")
        form_layout_v.addWidget(self.status_label)
        form_layout_v.addLayout(form)
        form_layout_v.addStretch()

        # Button row
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_save = QtWidgets.QPushButton("Save")
        btn_save.setMinimumHeight(36)
        btn_save.setMinimumWidth(110)
        btn_save.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #00C6FB, stop:1 #005BEA);
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 15px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #00E1FF, stop:1 #005BEA);
            }
        """)
        btn_save.clicked.connect(self.write_to_excel)
        btn_clear = QtWidgets.QPushButton("Clear")
        btn_clear.setMinimumHeight(36)
        btn_clear.setMinimumWidth(110)
        btn_clear.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #FFDEE9, stop:1 #FF5050);
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 15px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #FFB3B3, stop:1 #FF1744);
            }
        """)
        btn_clear.clicked.connect(self.clear_all_fields)
        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_clear)
        btn_row.addStretch()
        form_layout_v.addLayout(btn_row)

        main_layout.addWidget(form_container, stretch=3)

        # --- RIGHT PANEL ---
        right_panel = QtWidgets.QWidget()
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
        edt_path.setMinimumHeight(26)
        edt_path.setStyleSheet("""
            background-color: #FFFFFF;
            color: #222222;
            font-size: 15px;
            border-radius: 10px;
            border: 1.2px solid #B0C4DE;
            padding-left: 8px;
            padding-right: 8px;
        """)
        edt_path.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
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
        lbl_image.setMinimumSize(150, 90)
        lbl_image.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        lbl_image.setAlignment(QtCore.Qt.AlignCenter)
        lbl_image.setStyleSheet("""
            QLabel {
                background-color: #FFFFFF;
                border: 1px solid #CCCCCC;
                border-radius: 0px;
            }
        """)
        v_main.addWidget(lbl_image)
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

    # ========= Xử lý ảnh giữ đúng tỷ lệ =========
    def set_label_image(self, label, img):
        h, w = img.shape[:2]
        qt_img = QtGui.QImage(img.data, w, h, 3*w, QtGui.QImage.Format.Format_RGB888)
        pixmap = QtGui.QPixmap.fromImage(qt_img)
        scaled_pixmap = pixmap.scaled(
            label.width(), label.height(),
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        label.setPixmap(scaled_pixmap)

    def select_front_image_from_file(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn ảnh mặt trước", "", "Image Files (*.jpg *.jpeg *.png)")
        if not file_path:
            return
        self.front_img_path.setText(file_path)
        img = cv2.imread(file_path)
        if img is not None:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            self.set_label_image(self.front_img_label, img)

    def select_back_image_from_file(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn ảnh mặt sau", "", "Image Files (*.jpg *.jpeg *.png)")
        if not file_path:
            return
        self.back_img_path.setText(file_path)
        img = cv2.imread(file_path)
        if img is not None:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            self.set_label_image(self.back_img_label, img)

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
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        self.front_img_label.setPixmap(pixmap)
        self.front_img_path.setText("")
        self.front_image_temp = image.copy()
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
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        self.back_img_label.setPixmap(pixmap)
        self.back_img_path.setText("")
        self.back_image_temp = image.copy()
        try:
            self.image_capture_cam.imageCaptured.disconnect()
        except Exception:
            pass

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
        if hasattr(self, "qr_cam_index") and 0 <= self.qr_cam_index < len(cameras_info):
            self.cap_qr = cv2.VideoCapture(self.qr_cam_index, cv2.CAP_DSHOW)
            self.cap_qr.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.cap_qr.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            self.frame_count = 0
            self.timer_cv = QtCore.QTimer(self)
            self.timer_cv.timeout.connect(self.read_frame_for_qr)
            self.timer_cv.start(30)
            self.qr_thread = QRDecodeThread(self.qreader)
            self.qr_thread.qrDecoded.connect(self.on_qr_decoded)
            self.qr_thread.start()

        if 0 <= self.index_camera_cam < len(cameras_info):
            cam_info_cam = cameras_info[self.index_camera_cam]
            self.camera_cam = QtMultimedia.QCamera(cam_info_cam)
            self.camera_cam.setViewfinder(self.viewfinder_cam)
            self.viewfinder_cam.show()
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
        pixmap = QtGui.QPixmap.fromImage(qt_img)
        scaled_pixmap = pixmap.scaled(
            self.lbl_qr_preview.width(),
            self.lbl_qr_preview.height(),
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        self.lbl_qr_preview.setPixmap(scaled_pixmap)
        if self.qr_thread is not None and not self.qr_thread.isRunning():
            return
        if self.qr_thread is not None and self.frame_count % 5 == 0:
            self.qr_thread.request_decode(rgb_frame)
        self.frame_count = (self.frame_count + 1) % 30

    def _reset_can_decode(self):
        self.can_decode = True

    def on_qr_decoded(self, qr_text):
        if not self.can_decode:
            return
        info = parse_qr(qr_text)
        if not info:
            return
        self.fill_form_from_info(info)
        self._play_sound("done.wav")
        self.last_qr_text = qr_text
        self.can_decode = False
        QtCore.QTimer.singleShot(3000, self._reset_can_decode)

    def fill_form_from_info(self, info: dict):
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
        self.status_label.setText("✅ Quét QR từ webcam thành công, đã điền thông tin.")
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
        self.last_qr_text = ""

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

    def write_to_excel(self):
        import win32com.client as win32
        from datetime import datetime
        import shutil

        ten_phong = self.cmb_phong.currentText() if hasattr(self.cmb_phong, "currentText") else self.cmb_phong.currentText()
        if not ten_phong or ten_phong.strip() == "":
            QtWidgets.QMessageBox.warning(self, "Thiếu thông tin", "Bạn chưa chọn tên phòng!")
            return
        try:
            current_dir = self.app_dir
            excel_path = os.path.join(current_dir, "DU_LIEU_KHAI_BAO.xlsx")
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False

            if not os.path.exists(excel_path):
                wb = excel.Workbooks.Add()
                ws = wb.Sheets(1)
                headers = ["Số giấy tờ", "Số CMND cũ", "Họ và tên", "Giới tính", "Ngày sinh",
                            "Nơi thường trú", "Ngày cấp", "Loại giấy tờ", "Tên phòng", "Thời gian ghi",
                            "Ảnh mặt trước", "Ảnh mặt sau"]
                for col, header in enumerate(headers, start=1):
                    cell = ws.Cells(1, col)
                    cell.Value = header
                    cell.Font.Name = "Times New Roman"
                    cell.Font.Size = 16
                    cell.Font.Bold = True
                    cell.Interior.Color = 0x00C0FF
                    cell.HorizontalAlignment = -4108
                wb.SaveAs(excel_path)
                wb.Close()

            wb = excel.Workbooks.Open(excel_path)
            ws = wb.Sheets(1)

            row = 2
            while ws.Cells(row, 1).Value:
                row += 1

            def save_image_when_needed(image_temp, old_path, prefix):
                if old_path and os.path.exists(old_path):
                    folder = os.path.join(current_dir, "Anh_CCCD_da_khai_bao")
                    os.makedirs(folder, exist_ok=True)
                    filename = os.path.basename(old_path)
                    dest_path = os.path.join(folder, filename)
                    if old_path != dest_path:
                        shutil.copyfile(old_path, dest_path)
                    return dest_path
                elif image_temp is not None:
                    folder = os.path.join(current_dir, "Anh_CCCD_da_khai_bao")
                    os.makedirs(folder, exist_ok=True)
                    filename = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                    save_path = os.path.join(folder, filename)
                    image_temp.save(save_path)
                    return save_path
                return ""

            front_path_text = self.front_img_path.text() if self.front_img_path else ""
            back_path_text = self.back_img_path.text() if self.back_img_path else ""
            front_final = save_image_when_needed(self.front_image_temp, front_path_text, "mat_truoc")
            back_final = save_image_when_needed(self.back_image_temp, back_path_text, "mat_sau")

            data = [
                self.fields["Số giấy tờ"].text(),
                self.fields["Số CMND cũ (nếu có)"].text(),
                self.fields["Họ và tên"].text(),
                "Nam" if self.rb_nam.isChecked() else ("Nữ" if self.rb_nu.isChecked() else ""),
                self.fields["Ngày sinh"].date().toString("dd/MM/yyyy"),
                self.fields["Nơi thường trú"].text(),
                self.fields["Ngày cấp giấy tờ"].date().toString("dd/MM/yyyy"),
                self.fields["Loại giấy tờ"].text(),
                ten_phong,
                datetime.now().strftime("%H:%M:%S %d/%m/%Y"),
                front_final,
                back_final
            ]
            for col, value in enumerate(data, start=1):
                cell = ws.Cells(row, col)
                if col in [11, 12] and value and os.path.exists(value):
                    label = "Ảnh mặt trước" if col == 11 else "Ảnh mặt sau"
                    cell.Formula = f'=HYPERLINK("{value}", "{label}")'
                else:
                    cell.Value = value
                cell.Font.Size = 11
                cell.Font.Name = "Times New Roman"
                cell.HorizontalAlignment = -4108

            ws.Range("A:L").EntireColumn.AutoFit()
            rng = ws.Range(f"A1:L{row}")
            borders = rng.Borders
            for i in range(7, 13):
                borders(i).LineStyle = 1
                borders(i).Weight = 2

            wb.Save()
            wb.Close()
            self.front_image_temp = None
            self.back_image_temp = None
            QtWidgets.QMessageBox.information(self, "Lưu Excel", f"Lưu thông tin thành công.\nDòng thứ {row} trong file DU_LIEU_KHAI_BAO.xlsx")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi Excel", f"Lỗi khi ghi Excel: {e}")

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
        if self.front_img_path:
            self.front_img_path.setText("")
        if self.back_img_path:
            self.back_img_path.setText("")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
