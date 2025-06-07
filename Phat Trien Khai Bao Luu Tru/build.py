import PyInstaller.__main__
import os
import sys

# Đường dẫn tới thư mục hiện tại
current_dir = os.path.dirname(os.path.abspath(__file__))

# Các file cần copy
datas = [
    ('logo_app.ico', '.'),
    ('loading.gif', '.'),
    ('done.wav', '.'),
    ('error.wav', '.'),
    ('database.py', '.')
]

# Các thư viện cần exclude để giảm kích thước
excludes = [
    'matplotlib', 'tkinter', 'scipy', 'PIL', 'pandas',
    'notebook', 'jedi', 'IPython', 'ipykernel', 'jupyter_client',
    'tornado', 'zmq', 'debugpy'
]

# Các options cho PyInstaller
options = [
    'APP_KHAI_BAO_LUU_TRU_2.py',  # Tên file chính
    '--name=Khai Bao Luu Tru',  # Tên file exe
    '--onefile',  # Đóng gói thành 1 file
    '--noconsole',  # Không hiển thị console
    '--icon=logo_app.ico',  # Icon cho file exe
    '--clean',  # Xóa cache cũ
    '--windowed',  # Ứng dụng GUI
    '--noupx',  # Không nén UPX để tránh lỗi
    '--noconfirm',  # Không hỏi khi ghi đè
]

# Thêm data files
for src, dst in datas:
    options.extend(['--add-data', f'{src};{dst}'])

# Thêm excludes
for pkg in excludes:
    options.extend(['--exclude-module', pkg])

# Thêm hidden imports cần thiết
hidden_imports = [
    'cv2', 'numpy', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets',
    'PyQt5.QtMultimedia', 'PyQt5.QtMultimediaWidgets', 'qreader',
    'win32com.client', 'psutil', 'unidecode'
]

for imp in hidden_imports:
    options.extend(['--hidden-import', imp])

# Thêm các options tối ưu hiệu năng
options.extend([
    '--disable-windowed-traceback',  # Tắt traceback trong GUI
    '--optimize=2',  # Tối ưu bytecode level 2
])

# Chạy PyInstaller
PyInstaller.__main__.run(options)