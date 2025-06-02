import sys
import io

# Bảo vệ tqdm khỏi lỗi stderr None khi đóng gói bằng pyinstaller --noconsole
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

import os
import json
import cv2
import numpy as np
from unidecode import unidecode
from PyQt5 import QtCore, QtGui, QtWidgets, QtMultimedia, QtMultimediaWidgets
from PyQt5.QtMultimedia import QCameraImageCapture
from qreader import QReader

# ========== Tiện ích chung ==========
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

def image_to_qpixmap(image, target_label):
    img_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    h, w, ch = img_rgb.shape
    bytes_per_line = ch * w
    qt_img = QtGui.QImage(img_rgb.data, w, h, bytes_per_line, QtGui.QImage.Format_RGB888)
    return QtGui.QPixmap.fromImage(qt_img).scaled(
        target_label.width(), target_label.height(),
        QtCore.Qt.IgnoreAspectRatio, QtCore.Qt.SmoothTransformation
    )

# ========== Luồng giải mã QR ==========
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
            self.msleep(30)  # giảm CPU load hơn so với 10ms

# ========== Khởi động chương trình ==========
if __name__ == "__main__":
    from PyQt5.QtWidgets import QApplication
    from main_window import MainWindow  # Tách phần MainWindow ra file riêng nếu cần

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
