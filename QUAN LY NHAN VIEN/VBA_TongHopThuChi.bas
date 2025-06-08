' ===================================================================
' VBA CODE CHO EXCEL - TƯƠNG ĐƯƠNG VỚI GOOGLE APPS SCRIPT
' Chức năng: Tự động tạo báo cáo tổng hợp thu chi
' ===================================================================

' Đặt code này vào Sheet Module của sheet "TỔNG HỢP THU - CHI NHÀ NGHỈ"
' Hoặc vào ThisWorkbook nếu muốn áp dụng cho toàn bộ workbook

Option Explicit

' ===================================================================
' SỰ KIỆN WORKSHEET_CHANGE - Tương đương với onEdit trong Google Apps Script
' ===================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim allowedSheets As String
    Dim ws As Worksheet
    Dim sheetName As String
    Dim col As Long
    Dim row As Long
    Dim valueC As Variant
    Dim value As Variant
    
    ' Tắt tính toán tự động và events để tránh loop
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Kiểm tra sheet được phép
    allowedSheets = "TỔNG HỢP THU - CHI NHÀ NGHỈ"
    Set ws = Target.Worksheet
    sheetName = ws.Name
    
    If InStr(allowedSheets, sheetName) = 0 Then
        GoTo CleanUp
    End If
    
    ' Chỉ xử lý khi thay đổi 1 ô duy nhất
    If Target.Cells.Count > 1 Then
        GoTo CleanUp
    End If
    
    col = Target.Column
    row = Target.row
    
    ' === PHẦN 1: Ghi thời gian vào cột D khi nhập số tiền ở cột C ===
    If col = 3 Then ' Cột C
        valueC = ws.Cells(row, 3).Value
        If valueC <> "" And IsNumeric(valueC) Then
            ws.Cells(row, 4).Value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        Else
            ws.Cells(row, 4).Value = ""
        End If
    End If
    
    ' === PHẦN 2: Tạo bảng báo cáo khi nhập ngày vào cột A ===
    If col = 1 Then ' Cột A
        value = ws.Cells(row, 1).Value
        Call TaoBangBaoCao(ws, row, value)
    End If
    
CleanUp:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Lỗi: " & Err.Description, vbCritical
End Sub

' ===================================================================
' HÀM TẠO BẢNG BÁO CÁO
' ===================================================================
Private Sub TaoBangBaoCao(ws As Worksheet, row As Long, value As Variant)
    Dim startCol As Long
    Dim valueCol As Long
    Dim defaultEndRow As Long
    Dim labels As Variant
    Dim highlightColors As Variant
    Dim formulas As Variant
    Dim i As Long
    Dim titleCell As Range
    Dim formattedDate As String
    Dim labelCell As Range
    Dim valueCellRange As Range
    
    startCol = 8 ' Cột H
    valueCol = 9 ' Cột I
    defaultEndRow = 1000
    
    ' Mảng nhãn
    labels = Array( _
        "Số dư ban đầu tiền mặt:", _
        "Số dư ban đầu tài khoản:", _
        "Thu tiền mặt:", _
        "Chi tiền mặt:", _
        "Thu tài khoản:", _
        "Chi tài khoản:", _
        "Tổng thu:", _
        "Tổng chi:", _
        "Tiền mặt hiện có:", _
        "Tài khoản hiện có:", _
        "Tổng tiền hiện có:" _
    )
    
    ' Mảng màu nền (RGB)
    highlightColors = Array( _
        RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), _
        RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), _
        RGB(84, 129, 53), RGB(84, 129, 53), RGB(1, 176, 80) _
    )
    
    ' Nếu giá trị rỗng, xóa bảng báo cáo
    If IsEmpty(value) Or value = "" Then
        ws.Range(ws.Cells(row, startCol), ws.Cells(row + UBound(labels) + 1, valueCol)).Clear
        Exit Sub
    End If
    
    ' Tạo tiêu đề
    Set titleCell = ws.Range(ws.Cells(row, startCol), ws.Cells(row, valueCol))
    titleCell.Merge
    
    If IsDate(value) Then
        formattedDate = Format(value, "dd/mm/yyyy")
    Else
        formattedDate = CStr(value)
    End If
    
    With titleCell
        .Value = "BÁO CÁO TỔNG HỢP - NGÀY " & formattedDate
        .Font.Name = "Times New Roman"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(254, 242, 203)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Mảng công thức
    formulas = Array( _
        "", _
        "", _
        "=SUMIFS(C" & row & ":C" & defaultEndRow & ",F" & row & ":F" & defaultEndRow & ",""Tiền mặt"",E" & row & ":E" & defaultEndRow & ",""Thu"")", _
        "=SUMIFS(C" & row & ":C" & defaultEndRow & ",F" & row & ":F" & defaultEndRow & ",""Tiền mặt"",E" & row & ":E" & defaultEndRow & ",""Chi"")", _
        "=SUMIFS(C" & row & ":C" & defaultEndRow & ",F" & row & ":F" & defaultEndRow & ",""Chuyển khoản"",E" & row & ":E" & defaultEndRow & ",""Thu"")", _
        "=SUMIFS(C" & row & ":C" & defaultEndRow & ",F" & row & ":F" & defaultEndRow & ",""Chuyển khoản"",E" & row & ":E" & defaultEndRow & ",""Chi"")", _
        "=I" & (row + 3) & "+I" & (row + 5), _
        "=I" & (row + 4) & "+I" & (row + 6), _
        "=I" & (row + 1) & "+I" & (row + 3) & "-I" & (row + 4), _
        "=I" & (row + 2) & "+I" & (row + 5) & "-I" & (row + 6), _
        "=I" & (row + 9) & "+I" & (row + 10) _
    )
    
    ' Tạo các dòng dữ liệu
    For i = 0 To UBound(labels)
        Set labelCell = ws.Cells(row + 1 + i, startCol)
        Set valueCellRange = ws.Cells(row + 1 + i, valueCol)
        
        ' Đặt nhãn
        labelCell.Value = labels(i)
        
        ' Đặt công thức hoặc giá trị rỗng
        If formulas(i) <> "" Then
            valueCellRange.Formula = formulas(i)
        Else
            valueCellRange.Value = ""
        End If
        
        ' Định dạng nhãn
        With labelCell
            .Font.Name = "Times New Roman"
            .Font.Size = 15
            .VerticalAlignment = xlCenter
            .Interior.Color = highlightColors(i)
            .Borders.LineStyle = xlContinuous
            
            ' Định dạng font theo điều kiện
            If i <= 7 Then
                .Font.Bold = False
                .Font.Color = RGB(0, 0, 0) ' Đen
            ElseIf i = 8 Or i = 9 Then
                .Font.Bold = True
                .Font.Color = RGB(228, 193, 178) ' Màu nâu nhạt
            ElseIf i = 10 Then
                .Font.Bold = True
                .Font.Color = RGB(0, 0, 0) ' Đen
            End If
        End With
        
        ' Định dạng ô giá trị
        With valueCellRange
            .Font.Name = "Times New Roman"
            .Font.Size = 15
            .VerticalAlignment = xlCenter
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Đặt border cho cả 2 ô
        ws.Range(ws.Cells(row + 1 + i, startCol), ws.Cells(row + 1 + i, valueCol)).Borders.LineStyle = xlContinuous
    Next i
End Sub

' ===================================================================
' HÀM DI CHUYỂN ĐẾN DÒNG CUỐI CÙNG
' ===================================================================
Sub DenDongCuoiCung()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    ws.Cells(lastRow + 20, 1).Select
End Sub

' ===================================================================
' SỰ KIỆN WORKBOOK_OPEN - Tương đương với onOpen trong Google Apps Script
' ===================================================================
' Đặt code này vào ThisWorkbook module
Private Sub Workbook_Open()
    Call DenDongCuoiCung
End Sub
