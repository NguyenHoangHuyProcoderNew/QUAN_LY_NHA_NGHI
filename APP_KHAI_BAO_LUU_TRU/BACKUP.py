import win32com.client as win32
from tkinter import *
from tkinter import filedialog, ttk
from pyzbar.pyzbar import decode
from PIL import Image
from datetime import datetime
from unidecode import unidecode
import cv2
import os
import builtins
open = builtins.open
import sys
sys.stdout = open(os.devnull, 'w')
sys.stderr = open(os.devnull, 'w')
from qreader import QReader
import winsound
from tkinter import filedialog
from pygrabber.dshow_graph import FilterGraph
import threading

def chuan_hoa_ngay(dmy: str):
    return f"{dmy[:2]}/{dmy[2:4]}/{dmy[4:]}" if len(dmy) == 8 else dmy

def parse_qr(data: str):
    parts = data.split('|')
    if len(parts) >= 7:
        return {
            "Số giấy tờ": parts[0],
            "CMND": parts[1],
            "Họ tên": parts[2],
            "Ngày sinh": chuan_hoa_ngay(parts[3]),
            "Giới tính": parts[4],
            "Địa chỉ": parts[5],
            "Ngày cấp": chuan_hoa_ngay(parts[6])
        }
    return None

class QRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phần mềm khai báo lưu trú BY Nguyễn Hoàng Huy - My phone: 033.293.6390")
        self.root.geometry("1000x700")
        self.root.configure(bg="#e6f2ff")

        self.fields = {}
        self.info = None
        self.front_img_path = None
        self.back_img_path = None
        self.front_path_var = StringVar()
        self.back_path_var = StringVar()

        labels = [
            "Số giấy tờ", "Số CMND cũ (nếu có)", "Họ và tên",
            "Ngày sinh", "Giới tính", "Nơi thường trú",
            "Ngày cấp giấy tờ", "Loại giấy tờ", "Tên phòng lưu trú"
        ]

        self.ds_phong = [
            "", "Phòng 3 nhà cũ", "Phòng 4 nhà cũ", "Phòng 5 nhà cũ",
            "Phòng 7 nhà cũ", "Phòng 8 nhà cũ", "Phòng 9 nhà cũ",
            "Phòng 1 nhà mới", "Phòng 2 nhà mới", "Phòng 3 nhà mới",
            "Phòng 4 nhà mới", "Phòng 5 nhà mới"
        ]

        font_label = ("Segoe UI", 11)
        font_entry = ("Segoe UI", 11)

        main_frame = Frame(root, bg="#e6f2ff")
        main_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)

        form_frame = Frame(main_frame, bg="#e6f2ff")
        form_frame.pack(side=LEFT, padx=10, pady=10, anchor=N)

        button_frame = Frame(main_frame, bg="#e6f2ff")
        button_frame.pack(side=RIGHT, padx=10, pady=10, anchor=N)

        for idx, label in enumerate(labels):
            Label(form_frame, text=label + ":", font=font_label, bg="#e6f2ff", anchor="w", width=25).grid(row=idx, column=0, sticky=W, padx=10, pady=8)

            if label == "Giới tính":
                self.gender_var = StringVar(value="none")
                gender_frame = Frame(form_frame, bg="#e6f2ff")
                gender_frame.grid(row=idx, column=1, sticky=W)
                Radiobutton(gender_frame, variable=self.gender_var, value="none").pack_forget()
                Radiobutton(gender_frame, text="Nam", variable=self.gender_var, value="Nam", bg="#e6f2ff", font=font_entry).pack(side=LEFT, padx=5)
                Radiobutton(gender_frame, text="Nữ", variable=self.gender_var, value="Nữ", bg="#e6f2ff", font=font_entry).pack(side=LEFT, padx=5)
            elif label == "Tên phòng lưu trú":
                cb = ttk.Combobox(form_frame, values=self.ds_phong, state="readonly", width=33, font=font_entry)
                cb.grid(row=idx, column=1, sticky=W, padx=5)
                cb.set("")
                self.fields[label] = cb
            else:
                entry = Entry(form_frame, width=40, font=font_entry)
                if label == "Loại giấy tờ":
                    entry.insert(0, "CCCD")
                entry.grid(row=idx, column=1, sticky=W, padx=5)
                self.fields[label] = entry

        # 🆕 Thêm ô hiển thị đường dẫn ảnh
        Label(form_frame, text="Đường dẫn ảnh mặt trước:", font=font_label, bg="#e6f2ff").grid(row=len(labels), column=0, sticky=W, padx=10, pady=8)
        Entry(form_frame, textvariable=self.front_path_var, font=font_entry, width=40, state='readonly').grid(row=len(labels), column=1, sticky=W, padx=5)

        Label(form_frame, text="Đường dẫn ảnh mặt sau:", font=font_label, bg="#e6f2ff").grid(row=len(labels)+1, column=0, sticky=W, padx=10, pady=8)
        Entry(form_frame, textvariable=self.back_path_var, font=font_entry, width=40, state='readonly').grid(row=len(labels)+1, column=1, sticky=W, padx=5)
        btn_style = {"font": ("Segoe UI", 11, "bold"), "padx": 10, "pady": 5}
        Button(button_frame, text="📷 Quét thông tin CCCD từ ảnh trên máy tính", command=self.chon_anh, bg="#007acc", fg="white", **btn_style).pack(fill=X, pady=6)
        Button(button_frame, text="📸 Quét thông tin CCCD bằng Webcam", command=self.quet_qr_webcam, bg="#ffa500", fg="white", **btn_style).pack(fill=X, pady=6)
        Button(button_frame, text="🖼️ Chụp ảnh mặt trước", command=lambda: self.chup_anh('front'), bg="#17a2b8", fg="white", **btn_style).pack(fill=X, pady=6)
        Button(button_frame, text="🖼️ Chụp ảnh mặt sau", command=lambda: self.chup_anh('back'), bg="#6f42c1", fg="white", **btn_style).pack(fill=X, pady=6)
        Button(button_frame, text="📁 Tải ảnh mặt trước từ máy tính", command=self.tai_anh_mat_truoc, bg="#20c997", fg="white", **btn_style).pack(fill=X, pady=6)
        Button(button_frame, text="📁 Tải ảnh mặt sau từ máy tính", command=self.tai_anh_mat_sau, bg="#6610f2", fg="white", **btn_style).pack(fill=X, pady=6)
        Button(button_frame, text="📝 Lưu thông tin", command=self.ghi_excel, bg="#28a745", fg="white", **btn_style).pack(fill=X, pady=6)
        Button(button_frame, text="📂 Xem thông tin đã lưu", command=self.mo_excel, bg="#343a40", fg="white", **btn_style).pack(fill=X, pady=6)
        Button(button_frame, text="❌ Xóa dữ liệu", command=self.clear_fields, bg="#dc3545", fg="white", **btn_style).pack(fill=X, pady=6)

        self.status_label = Label(self.root, text="", font=("Segoe UI", 11), bg="#e6f2ff")
        self.status_label.pack(pady=5)

        # 🆕 Thêm combobox chọn webcam (hiển thị theo tên thiết bị)
        Label(button_frame, text="🖥️ Chọn webcam:", font=("Segoe UI", 11, "bold"), bg="#e6f2ff").pack(pady=(0, 5))

        self.webcam_name_var = StringVar()
        webcams = self.liet_ke_webcam()  # Gán danh sách thiết bị và lưu vào self.ten_webcams
        self.webcam_combobox = ttk.Combobox(button_frame, textvariable=self.webcam_name_var, state="readonly", font=("Segoe UI", 10))
        self.webcam_combobox['values'] = webcams
        self.webcam_combobox.pack(fill=X, pady=5)

        if webcams:
            self.webcam_combobox.current(0)  # Chọn webcam đầu tiên làm mặc định
        
        # 🔧 Đảm bảo folder ảnh cùng cấp với file chương trình
        if getattr(sys, 'frozen', False):
            # Chạy từ file .exe
            self.app_dir = os.path.dirname(sys.executable)
        else:
            # Chạy từ file .py
            self.app_dir = os.path.dirname(os.path.abspath(__file__))

        self.folder_anh = os.path.join(self.app_dir, "Anh_CCCD_da_khai_bao")
        os.makedirs(self.folder_anh, exist_ok=True)


    def show_status(self, message, error=False):
        self.status_label.config(text=message, fg="red" if error else "green")

    def xoa_thong_bao(self):
        self.status_label.config(text="")

    def chon_anh(self):
        self.xoa_thong_bao()
        self.show_status("📥 Chọn ảnh có chứa mã QR gắn trên CCCD.")
        self.root.update_idletasks()

        image_path = filedialog.askopenfilename(
            title="Chọn ảnh CCCD",
            filetypes=[("Image files", "*.jpg *.jpeg *.png")]
        )

        if not image_path:
            return

        try:
            image = cv2.imread(image_path)

            if image is None:
                self.show_status("❌ Không đọc được ảnh. File có thể bị hỏng, sai định dạng, hoặc đang bị khóa.", error=True)
                return

            image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
            qreader = QReader()
            decoded_data = qreader.detect_and_decode(image=image_rgb)

            if not decoded_data or not decoded_data[0]:
                self.show_status("❌ Không quét được mã QR từ ảnh.", error=True)
                self.phat_am_thanh("error.wav")
                return

            qr_text = decoded_data[0]
            self.dien_thong_tin_tu_qr(qr_text)

            self.front_img_path = image_path
            self.front_path_var.set(image_path)

        except Exception as e:
            # ❌ Bỏ ghi file log, chỉ in ra console và báo lỗi cho người dùng
            print(f"Lỗi khi xử lý ảnh: {e}")
            self.show_status(f"❌ Lỗi khi xử lý ảnh: {e}", error=True)

    def quet_qr_webcam(self):
        def scan_qr():
            self.xoa_thong_bao()
            self.show_status(
                "📸 Đưa phần mã QR của CCCD lại gần webcam để quét\n"
                "Nhấn 'Q' trên bàn phím nếu muốn thoát"
            )

            webcam_index = self.ten_webcams.index(self.webcam_name_var.get())
            cap = cv2.VideoCapture(webcam_index)
            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

            qreader = QReader()
            found_data = None
            frame_count = 0

            while True:
                ret, frame = cap.read()
                if not ret:
                    break

                frame_count += 1
                frame_display = cv2.resize(frame, (640, 360))

                # Chỉ xử lý mỗi 6 frames (giảm tải)
                if frame_count % 6 == 0:
                    image_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    result = qreader.detect_and_decode(image=image_rgb)
                    if result and result[0]:
                        found_data = result[0]
                        break

                cv2.imshow("📷 Webcam quét QR CCCD", frame_display)
                key = cv2.waitKey(1)
                if key & 0xFF == ord('q'):
                    break

            cap.release()
            cv2.destroyAllWindows()
            self.xoa_thong_bao()

            if found_data:
                self.dien_thong_tin_tu_qr(found_data)
            else:
                self.show_status("❌ Không quét được mã QR từ webcam.", error=True)
                self.phat_am_thanh("error.wav")

        # Chạy trong luồng riêng
        threading.Thread(target=scan_qr, daemon=True).start()

    def chup_anh(self, loai):
        self.xoa_thong_bao()
        huong_dan = "📸 Hướng dẫn: Đặt CCCD mặt trước lên nền sáng, rõ nét. Nhấn 'S' để chụp, 'Q' để thoát." if loai == 'front' \
            else "📸 Hướng dẫn: Đặt CCCD mặt sau lên nền sáng, rõ nét. Nhấn 'S' để chụp, 'Q' để thoát."
        self.show_status(huong_dan)
        self.root.update_idletasks()

        webcam_index = self.ten_webcams.index(self.webcam_name_var.get())
        cap = cv2.VideoCapture(webcam_index)

        while True:
            ret, frame = cap.read()
            if not ret:
                break
            cv2.imshow("Webcam chup mat truoc/sau", frame)
            key = cv2.waitKey(1)
            if key & 0xFF == ord('s'):
                folder = self.folder_anh
                os.makedirs(folder, exist_ok=True)

                ho_ten_raw = self.fields["Họ và tên"].get().strip()
                ho_ten_khong_dau = unidecode(ho_ten_raw).replace(" ", "_") if ho_ten_raw else "Khong_ten"
                ngay = datetime.now().strftime("%d_%m_%Y")

                filename = f"{'mat_truoc' if loai == 'front' else 'mat_sau'}_{ho_ten_khong_dau}_{ngay}.jpg"
                img_path = os.path.join(folder, filename)
                cv2.imwrite(img_path, frame)

                if loai == 'front':
                    self.front_img_path = img_path
                    self.front_path_var.set(img_path)
                else:
                    self.back_img_path = img_path
                    self.back_path_var.set(img_path)

                self.show_status(f"✅ Đã lưu ảnh {loai}: {img_path}")
                break
            elif key & 0xFF == ord('q'):
                break
        cap.release()
        cv2.destroyAllWindows()
        self.xoa_thong_bao()

    def tai_anh_mat_truoc(self):
        self.xoa_thong_bao()
        image_path = filedialog.askopenfilename(
            title="Tải lên ảnh mặt trước của giấy tờ",
            filetypes=[("Image files", "*.jpg *.jpeg *.png")]
        )
        if not image_path:
            return

        # Gán đường dẫn ảnh vào biến và giao diện (không copy file)
        self.front_img_path = image_path
        self.front_path_var.set(image_path)
        self.show_status("Tải lên ảnh mặt trước thành công ✅")

    def tai_anh_mat_sau(self):
        self.xoa_thong_bao()
        image_path = filedialog.askopenfilename(
            title="Tải lên ảnh mặt sau của giấy tờ",
            filetypes=[("Image files", "*.jpg *.jpeg *.png")]
        )
        if not image_path:
            return

        # Gán đường dẫn ảnh vào biến và giao diện (không copy file)
        self.back_img_path = image_path
        self.back_path_var.set(image_path)
        self.show_status("Tải lên ảnh mặt sau thành công ✅")

    def phat_am_thanh(self, ten_file_wav):
        try:
            import winsound
            import sys
            import os
            path = os.path.join(
                getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))),
                ten_file_wav
            )
            if os.path.exists(path):
                winsound.PlaySound(path, winsound.SND_FILENAME)
        except Exception as e:
            print(f"Lỗi phát âm thanh {ten_file_wav}: {e}")

    def dien_thong_tin_tu_qr(self, data):
        import winsound
        import sys
        import os

        self.info = parse_qr(data)
        if not self.info:
            self.show_status("❌ Không đọc được dữ liệu từ mã QR.", error=True)

            # 🔊 Phát âm thanh lỗi trực tiếp
            try:
                path = os.path.join(
                    getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))),
                    "error.wav"
                )
                if os.path.exists(path):
                    winsound.PlaySound(path, winsound.SND_FILENAME)
            except Exception as e:
                print(f"⚠️ Không phát được âm thanh lỗi: {e}")
            return

        # ✅ Điền thông tin vào các trường
        self.fields["Số giấy tờ"].delete(0, END)
        self.fields["Số giấy tờ"].insert(0, self.info["Số giấy tờ"])

        self.fields["Số CMND cũ (nếu có)"].delete(0, END)
        self.fields["Số CMND cũ (nếu có)"].insert(0, self.info["CMND"])

        self.fields["Họ và tên"].delete(0, END)
        self.fields["Họ và tên"].insert(0, self.info["Họ tên"])

        self.fields["Ngày sinh"].delete(0, END)
        self.fields["Ngày sinh"].insert(0, self.info["Ngày sinh"])

        self.gender_var.set(self.info["Giới tính"] if self.info["Giới tính"] in ["Nam", "Nữ"] else "none")

        self.fields["Nơi thường trú"].delete(0, END)
        self.fields["Nơi thường trú"].insert(0, self.info["Địa chỉ"])

        self.fields["Ngày cấp giấy tờ"].delete(0, END)
        self.fields["Ngày cấp giấy tờ"].insert(0, self.info["Ngày cấp"])

        self.fields["Loại giấy tờ"].delete(0, END)
        self.fields["Loại giấy tờ"].insert(0, "CCCD")

        self.show_status("Quét thông tin từ QR trên CCCD thành công ✅")

        # 🔊 Phát âm thanh thành công trực tiếp
        try:
            path = os.path.join(
                getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))),
                "done.wav"
            )
            if os.path.exists(path):
                winsound.PlaySound(path, winsound.SND_FILENAME)
        except Exception as e:
            print(f"⚠️ Không phát được âm thanh thành công: {e}")


    def ghi_excel(self):
        self.xoa_thong_bao()
        ten_phong = self.fields["Tên phòng lưu trú"].get().strip()
        if not ten_phong:
            self.show_status("⚠️ Bạn chưa chọn tên phòng.", error=True)
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

            # Tìm dòng trống tiếp theo
            row = 2
            while ws.Cells(row, 1).Value:
                row += 1

            # Tạo bản sao ảnh vào thư mục lưu nếu cần
            def copy_if_needed(src_path):
                try:
                    if src_path and os.path.exists(src_path):
                        filename = os.path.basename(src_path)
                        dest_path = os.path.join(self.folder_anh, filename)

                        if src_path != dest_path:
                            from shutil import copyfile
                            copyfile(src_path, dest_path)

                        return dest_path
                except Exception as e:
                    print(f"⚠️ Lỗi khi copy ảnh: {e}")
                return ""


            # Xử lý ảnh
            front_final = copy_if_needed(self.front_img_path)
            back_final = copy_if_needed(self.back_img_path)

            # Dữ liệu cần ghi
            data = [
                self.fields["Số giấy tờ"].get(),
                self.fields["Số CMND cũ (nếu có)"].get(),
                self.fields["Họ và tên"].get(),
                self.gender_var.get(),
                self.fields["Ngày sinh"].get(),
                self.fields["Nơi thường trú"].get(),
                self.fields["Ngày cấp giấy tờ"].get(),
                self.fields["Loại giấy tờ"].get(),
                ten_phong,
                datetime.now().strftime("%H:%M:%S %d/%m/%Y"),
                front_final,
                back_final
            ]

            for col, value in enumerate(data, start=1):
                cell = ws.Cells(row, col)
                if col in [11, 12] and os.path.exists(value):
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
            self.clear_fields()
            self.show_status(f"Lưu thông tin thành công ✅\nVào dòng thứ {row} trong file dữ liệu khai báo")

        except Exception as e:
            self.show_status(f"❌ Lỗi khi ghi Excel: {e}", error=True)

    def mo_excel(self):
        try:
            excel_path = os.path.join(self.app_dir, "DU_LIEU_KHAI_BAO.xlsx")
            if os.path.exists(excel_path):
                os.startfile(excel_path)
                self.show_status(f"Mở file dữ liệu khai báo thành công ✅\nĐường dẫn file: {excel_path}")
            else:
                self.show_status("⚠️ Chưa có file Excel nào được tạo.", error=True)
        except Exception as e:
            self.show_status(f"❌ Không mở được file dữ liệu khai báo: {e}", error=True)

    def clear_fields(self):
        self.xoa_thong_bao()
        for field in self.fields.values():
            if isinstance(field, Entry):
                field.delete(0, END)
            elif isinstance(field, ttk.Combobox):
                field.set("")
        self.gender_var.set("none")
        self.front_img_path = None
        self.back_img_path = None
        self.front_path_var.set("")
        self.back_path_var.set("")

    def liet_ke_webcam(self):
        graph = FilterGraph()
        self.ten_webcams = graph.get_input_devices()
        return self.ten_webcams

if __name__ == "__main__":
    root = Tk()

    # Lấy đường dẫn đến icon đúng, hỗ trợ cả khi chạy từ .exe
    if getattr(sys, 'frozen', False):
        app_dir = sys._MEIPASS
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))

    icon_path = os.path.join(app_dir, "logo_app.ico")
    if os.path.exists(icon_path):
        root.iconbitmap(icon_path)

    app = QRApp(root)
    root.mainloop()