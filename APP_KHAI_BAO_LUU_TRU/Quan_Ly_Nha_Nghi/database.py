import sqlite3
import os
import shutil
from datetime import datetime
from typing import Optional, List, Dict, Any

class Database:
    def __init__(self, app_dir: str):
        # Tạo cấu trúc thư mục
        self.app_dir = app_dir
        self.data_dir = os.path.join(app_dir, "data")
        self.images_dir = os.path.join(self.data_dir, "images")
        self.db_path = os.path.join(self.data_dir, "cong_dan.db")
        
        # Đảm bảo các thư mục tồn tại
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.images_dir, exist_ok=True)
        
        self._create_tables()
    
    def _create_tables(self):
        """Tạo các bảng nếu chưa tồn tại"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Bảng thông tin công dân
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cong_dan (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    so_giay_to TEXT,
                    so_cmnd_cu TEXT,
                    ho_ten TEXT,
                    gioi_tinh TEXT,
                    ngay_sinh TEXT,
                    noi_thuong_tru TEXT,
                    ngay_cap TEXT,
                    loai_giay_to TEXT,
                    ten_phong TEXT,
                    thoi_gian_ghi TEXT,
                    anh_mat_truoc TEXT,
                    anh_mat_sau TEXT
                )
            ''')
            
            # Tạo index để tìm kiếm nhanh
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_ho_ten ON cong_dan(ho_ten)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_so_giay_to ON cong_dan(so_giay_to)')
            
            conn.commit()

    def them_cong_dan(self, data: Dict[str, Any]) -> bool:
        """Thêm một công dân mới vào database"""
        try:
            print("Đang thêm công dân với dữ liệu:")
            print(f"- Họ tên: {data.get('ho_ten', '')}")
            
            # Di chuyển ảnh vào thư mục images nếu cần
            anh_mat_truoc = data.get('anh_mat_truoc', '')
            anh_mat_sau = data.get('anh_mat_sau', '')
            
            if anh_mat_truoc:
                # Kiểm tra xem ảnh có nằm ngoài thư mục images không
                if not anh_mat_truoc.startswith(os.path.join("data", "images")):
                    # Di chuyển ảnh vào thư mục images
                    ten_file = os.path.basename(anh_mat_truoc)
                    duong_dan_moi = os.path.join("data", "images", ten_file)
                    duong_dan_day_du = os.path.join(self.app_dir, duong_dan_moi)
                    
                    # Copy file ảnh vào thư mục mới
                    duong_dan_cu = os.path.join(self.app_dir, anh_mat_truoc)
                    if os.path.exists(duong_dan_cu):
                        os.makedirs(os.path.dirname(duong_dan_day_du), exist_ok=True)
                        shutil.copy2(duong_dan_cu, duong_dan_day_du)
                        anh_mat_truoc = duong_dan_moi
            
            if anh_mat_sau:
                # Tương tự cho ảnh mặt sau
                if not anh_mat_sau.startswith(os.path.join("data", "images")):
                    ten_file = os.path.basename(anh_mat_sau)
                    duong_dan_moi = os.path.join("data", "images", ten_file)
                    duong_dan_day_du = os.path.join(self.app_dir, duong_dan_moi)
                    
                    duong_dan_cu = os.path.join(self.app_dir, anh_mat_sau)
                    if os.path.exists(duong_dan_cu):
                        os.makedirs(os.path.dirname(duong_dan_day_du), exist_ok=True)
                        shutil.copy2(duong_dan_cu, duong_dan_day_du)
                        anh_mat_sau = duong_dan_moi
            
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO cong_dan (
                        so_giay_to, so_cmnd_cu, ho_ten, gioi_tinh,
                        ngay_sinh, noi_thuong_tru, ngay_cap, loai_giay_to,
                        ten_phong, thoi_gian_ghi, anh_mat_truoc, anh_mat_sau
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    data.get('so_giay_to', ''),
                    data.get('so_cmnd_cu', ''),
                    data.get('ho_ten', ''),
                    data.get('gioi_tinh', ''),
                    data.get('ngay_sinh', ''),
                    data.get('noi_thuong_tru', ''),
                    data.get('ngay_cap', ''),
                    data.get('loai_giay_to', ''),
                    data.get('ten_phong', ''),
                    data.get('thoi_gian_ghi', ''),
                    anh_mat_truoc,
                    anh_mat_sau
                ))
                conn.commit()
                print(f"✅ Đã thêm công dân thành công với ID: {cursor.lastrowid}")
                return True
        except Exception as e:
            print(f"❌ Lỗi khi thêm công dân: {e}")
            return False

    def tim_kiem_theo_ten(self, ten: str) -> List[Dict[str, Any]]:
        """Tìm kiếm công dân theo tên"""
        try:
            print(f"Đang tìm kiếm công dân với tên: {ten}")
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT * FROM cong_dan 
                    WHERE LOWER(ho_ten) LIKE ?
                    ORDER BY thoi_gian_ghi DESC
                ''', (f'%{ten.lower()}%',))
                
                results = []
                for row in cursor.fetchall():
                    result = dict(row)
                    print(f"Tìm thấy công dân:")
                    print(f"- Họ tên: {result['ho_ten']}")
                    print(f"- Ảnh mặt trước: {result['anh_mat_truoc']}")
                    print(f"- Ảnh mặt sau: {result['anh_mat_sau']}")
                    results.append(result)
                print(f"✅ Tìm thấy {len(results)} kết quả")
                return results
        except Exception as e:
            print(f"❌ Lỗi khi tìm kiếm: {e}")
            return []

    def tim_kiem_theo_ten_va_ngay(self, ten: str, tu_ngay: str, den_ngay: str) -> List[Dict[str, Any]]:
        """Tìm kiếm công dân theo tên và khoảng thời gian"""
        try:
            print(f"Tìm kiếm với tham số: tên='{ten}', từ ngày='{tu_ngay}', đến ngày='{den_ngay}'")
            
            # Chuyển đổi định dạng ngày để so sánh
            tu_ngay_obj = datetime.strptime(tu_ngay, "%d/%m/%Y")
            den_ngay_obj = datetime.strptime(den_ngay, "%d/%m/%Y")
            
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                
                # Base query với điều kiện thời gian
                base_query = '''
                    SELECT * FROM cong_dan 
                    WHERE substr(thoi_gian_ghi, 1, 10) BETWEEN ? AND ?
                '''
                
                # Thêm điều kiện tìm kiếm theo tên nếu có
                if ten.strip():
                    query = base_query + " AND LOWER(ho_ten) LIKE ? ORDER BY id ASC"
                    cursor.execute(query, (
                        tu_ngay_obj.strftime("%d/%m/%Y"),
                        den_ngay_obj.strftime("%d/%m/%Y"),
                        f'%{ten.lower()}%'
                    ))
                else:
                    query = base_query + " ORDER BY id ASC"
                    cursor.execute(query, (
                        tu_ngay_obj.strftime("%d/%m/%Y"),
                        den_ngay_obj.strftime("%d/%m/%Y")
                    ))
                
                results = []
                for row in cursor.fetchall():
                    result = dict(row)
                    print(f"Tìm thấy: {result['ho_ten']} - {result['thoi_gian_ghi']}")
                    results.append(result)
                
                print(f"Tổng số kết quả: {len(results)}")
                return results
                
        except Exception as e:
            print(f"❌ Lỗi khi tìm kiếm: {e}")
            return []

    def lay_tat_ca_cong_dan(self) -> List[Dict[str, Any]]:
        """Lấy danh sách tất cả công dân"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('SELECT * FROM cong_dan ORDER BY id ASC')  # Sắp xếp theo ID tăng dần
                return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            print(f"Lỗi khi lấy danh sách công dân: {e}")
            return []

    def _convert_date_format(self, date_str: str) -> str:
        """Chuyển đổi chuỗi ngày tháng sang định dạng chuẩn"""
        if not date_str:
            return ""
        try:
            # Nếu là định dạng dd/mm/yyyy
            if '/' in date_str:
                parts = date_str.split('/')
                if len(parts) == 3:
                    return f"'{parts[0]}/{parts[1]}/{parts[2]}"
            return f"'{date_str}"
        except:
            return f"'{date_str}"

    def xuat_excel(self, excel_path: str) -> bool:
        """Xuất dữ liệu ra file Excel"""
        try:
            import win32com.client as win32
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            
            # Thêm header
            headers = ["Số giấy tờ", "Số CMND cũ", "Họ và tên", "Giới tính", 
                      "Ngày sinh", "Nơi thường trú", "Ngày cấp", "Loại giấy tờ", 
                      "Tên phòng", "Thời gian ghi", "Ảnh mặt trước", "Ảnh mặt sau"]
            
            for col, header in enumerate(headers, 1):
                ws.Cells(1, col).Value = header
            
            # Format header
            header_range = ws.Range(ws.Cells(1, 1), ws.Cells(1, len(headers)))
            header_range.Font.Bold = True
            header_range.Interior.Color = 0x00C0FF
            
            # Đặt định dạng cột là Text cho các cột ngày tháng
            ws.Range("E:E,G:G,J:J").NumberFormat = "@"
            
            # Thêm dữ liệu
            data = self.lay_tat_ca_cong_dan()
            for row, item in enumerate(data, 2):
                ws.Cells(row, 1).Value = item['so_giay_to']
                ws.Cells(row, 2).Value = item['so_cmnd_cu']
                ws.Cells(row, 3).Value = item['ho_ten']
                ws.Cells(row, 4).Value = item['gioi_tinh']
                ws.Cells(row, 5).Value = item['ngay_sinh']  # Ngày sinh
                ws.Cells(row, 6).Value = item['noi_thuong_tru']
                ws.Cells(row, 7).Value = item['ngay_cap']  # Ngày cấp
                ws.Cells(row, 8).Value = item['loai_giay_to']
                ws.Cells(row, 9).Value = item['ten_phong']
                ws.Cells(row, 10).Value = item['thoi_gian_ghi']  # Thời gian ghi
                
                # Thêm hyperlink cho ảnh
                if item['anh_mat_truoc']:
                    # Lấy đường dẫn tuyệt đối từ thư mục gốc của ứng dụng
                    app_dir = os.path.dirname(self.db_path)  # Thư mục data
                    app_dir = os.path.dirname(app_dir)  # Thư mục gốc ứng dụng
                    image_path = os.path.join(app_dir, item['anh_mat_truoc'])
                    
                    if os.path.exists(image_path):
                        try:
                            # Chuyển đổi đường dẫn sang dạng URL
                            url_path = f"file:///{image_path.replace(os.sep, '/')}"
                            ws.Hyperlinks.Add(
                                Anchor=ws.Cells(row, 11),
                                Address=url_path,
                                TextToDisplay="Ảnh mặt trước"
                            )
                        except Exception as e:
                            print(f"Lỗi tạo hyperlink ảnh mặt trước: {e}")
                            ws.Cells(row, 11).Value = "Lỗi đường dẫn"
                    else:
                        print(f"Không tìm thấy ảnh mặt trước: {image_path}")
                        ws.Cells(row, 11).Value = "Không tìm thấy ảnh"

                if item['anh_mat_sau']:
                    # Lấy đường dẫn tuyệt đối từ thư mục gốc của ứng dụng
                    app_dir = os.path.dirname(self.db_path)  # Thư mục data
                    app_dir = os.path.dirname(app_dir)  # Thư mục gốc ứng dụng
                    image_path = os.path.join(app_dir, item['anh_mat_sau'])
                    
                    if os.path.exists(image_path):
                        try:
                            # Chuyển đổi đường dẫn sang dạng URL
                            url_path = f"file:///{image_path.replace(os.sep, '/')}"
                            ws.Hyperlinks.Add(
                                Anchor=ws.Cells(row, 12),
                                Address=url_path,
                                TextToDisplay="Ảnh mặt sau"
                            )
                        except Exception as e:
                            print(f"Lỗi tạo hyperlink ảnh mặt sau: {e}")
                            ws.Cells(row, 12).Value = "Lỗi đường dẫn"
                    else:
                        print(f"Không tìm thấy ảnh mặt sau: {image_path}")
                        ws.Cells(row, 12).Value = "Không tìm thấy ảnh"
            
            # Format bảng
            ws.Range("A:L").EntireColumn.AutoFit()
            ws.Range(f"A1:L{len(data)+1}").Borders.Weight = 2
            
            # Lưu và đóng
            wb.SaveAs(excel_path)
            wb.Close()
            excel.Quit()
            return True
            
        except Exception as e:
            print(f"Lỗi khi xuất Excel: {e}")
            return False

    def xuat_excel_tu_ket_qua(self, excel_path: str, results: List[Dict[str, Any]]) -> bool:
        """Xuất dữ liệu từ kết quả tìm kiếm ra file Excel"""
        try:
            import win32com.client as win32
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            
            # Thêm header
            headers = [
                "STT", "Số giấy tờ", "Số CMND cũ", "Họ và tên", "Giới tính", 
                "Ngày sinh", "Nơi thường trú", "Ngày cấp", "Loại giấy tờ", 
                "Tên phòng", "Thời gian ghi", "Ảnh mặt trước", "Ảnh mặt sau"
            ]
            
            for col, header in enumerate(headers, 1):
                ws.Cells(1, col).Value = header
            
            # Format header
            header_range = ws.Range(ws.Cells(1, 1), ws.Cells(1, len(headers)))
            header_range.Font.Bold = True
            header_range.Interior.Color = 0x00C0FF
            
            # Đặt định dạng cột là Text trước khi nhập dữ liệu
            ws.Range("F:F,H:H,K:K").NumberFormat = "@"
            
            # Thêm dữ liệu
            for idx, item in enumerate(results, 1):
                row = idx + 1
                # STT
                ws.Cells(row, 1).Value = idx
                
                # Các trường thông thường
                ws.Cells(row, 2).Value = item['so_giay_to']
                ws.Cells(row, 3).Value = item['so_cmnd_cu']
                ws.Cells(row, 4).Value = item['ho_ten']
                ws.Cells(row, 5).Value = item['gioi_tinh']
                
                # Xử lý các trường ngày tháng
                ws.Cells(row, 6).Value = self._convert_date_format(item['ngay_sinh'])
                ws.Cells(row, 7).Value = item['noi_thuong_tru']
                ws.Cells(row, 8).Value = self._convert_date_format(item['ngay_cap'])
                ws.Cells(row, 9).Value = item['loai_giay_to']
                ws.Cells(row, 10).Value = item['ten_phong']
                ws.Cells(row, 11).Value = self._convert_date_format(item['thoi_gian_ghi'])
                
                # Thêm hyperlink cho ảnh nếu có
                if item['anh_mat_truoc']:
                    full_path = os.path.abspath(os.path.join(os.path.dirname(self.db_path), item['anh_mat_truoc']))
                    ws.Hyperlinks.Add(
                        Anchor=ws.Cells(row, 12),
                        Address=f"file:///{full_path.replace(os.sep, '/')}",
                        TextToDisplay="Ảnh mặt trước"
                    )
                if item['anh_mat_sau']:
                    full_path = os.path.abspath(os.path.join(os.path.dirname(self.db_path), item['anh_mat_sau']))
                    ws.Hyperlinks.Add(
                        Anchor=ws.Cells(row, 13),
                        Address=f"file:///{full_path.replace(os.sep, '/')}",
                        TextToDisplay="Ảnh mặt sau"
                    )
            
            # Format bảng
            ws.Range("A:M").EntireColumn.AutoFit()
            ws.Range(f"A1:M{len(results)+1}").Borders.Weight = 2
            
            # Lưu và đóng
            wb.SaveAs(excel_path)
            wb.Close()
            excel.Quit()
            return True
            
        except Exception as e:
            print(f"Lỗi khi xuất Excel: {e}")
            return False 