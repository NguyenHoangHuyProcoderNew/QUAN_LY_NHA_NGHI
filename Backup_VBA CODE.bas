Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False

    If Target.Cells.CountLarge > 1 Then GoTo SafeExit
    
    ' Xử lý thời gian cho cột D khi cột C thay đổi
    If Not Intersect(Target, Columns(3)) Is Nothing Then
        ' Lấy dòng hiện tại
        Dim currentRow As Long
        currentRow = Target.Row
        
        ' Kiểm tra nếu cột C có dữ liệu
        If Not IsEmpty(Me.Cells(currentRow, 3).Value) Then
            ' Ghi thời gian vào cột D
            Me.Cells(currentRow, 4).Value = Format(Now, "hh:mm:ss dd/mm/yyyy")
            Me.Cells(currentRow, 4).NumberFormat = "hh:mm:ss dd/mm/yyyy"
        Else
            ' Xóa thời gian ở cột D nếu cột C trống
            Me.Cells(currentRow, 4).ClearContents
        End If
    End If

    Dim labels(0 To 10) As String
    Dim highlightColors As Variant
    Dim formulas(0 To 10) As String
    Dim r As Long
    Dim startCol As Long, valueCol As Long
    Dim i As Long
    Dim labelSheet As Worksheet
    Dim cashLabel As String, bankLabel As String
    Dim formattedDate As String
    Dim firstRow As Long, lastRow As Long
    Dim currentDate As Variant
    Dim maxLoop As Long: maxLoop = 1000
    
    ' Thêm biến để lưu giá trị số dư ban đầu
    Dim initialCashValue As Variant
    Dim initialBankValue As Variant
    Dim currentReportRow As Long

    If Intersect(Target, Columns(1)) Is Nothing And _
       Intersect(Target, Columns(3)) Is Nothing And _
       Intersect(Target, Columns(5)) Is Nothing And _
       Intersect(Target, Columns(6)) Is Nothing Then GoTo SafeExit

    r = Target.Row
    startCol = 8
    valueCol = 9
    Set labelSheet = ThisWorkbook.Sheets("SETTINGS VBA CODE")

    For i = 0 To 10
        labels(i) = labelSheet.Cells(i + 2, 1).Value
    Next i

    cashLabel = labelSheet.Range("A13").Value
    bankLabel = labelSheet.Range("A14").Value

    highlightColors = Array( _
        RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), _
        RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), _
        RGB(84, 129, 53), RGB(84, 129, 53), RGB(1, 176, 80))

    ' Tìm vị trí của báo cáo hiện tại và lưu giá trị số dư ban đầu
    currentReportRow = r
    Do While currentReportRow > 1
        If Me.Cells(currentReportRow, startCol).Value Like "*BÁO CÁO TỔNG HỢP*" Then
            ' Chỉ lưu giá trị số dư ban đầu nếu KHÔNG phải đang tạo bảng mới
            If Target.Column <> 1 Then
                initialCashValue = Me.Cells(currentReportRow + 1, valueCol).Value
                initialBankValue = Me.Cells(currentReportRow + 2, valueCol).Value
            End If
            Exit Do
        End If
        currentReportRow = currentReportRow - 1
        If currentReportRow <= 1 Then Exit Do
    Loop

    If Target.Column = 1 Then
        If Trim(Target.Value) = "" Then
            With Me.Range(Me.Cells(r, startCol), Me.Cells(r + UBound(labels) + 1, valueCol))
                .UnMerge
                .ClearContents
                .ClearFormats
                .Borders.LineStyle = xlNone
            End With
            GoTo SafeExit
        Else
            currentDate = Me.Cells(r, 1).Value
        End If
    Else
        Dim loopCount As Long: loopCount = 0
        Do While r > 1
            If Me.Cells(r, startCol).Value Like "*BÁO CÁO TỔNG HỢP*" Then Exit Do
            r = r - 1
            loopCount = loopCount + 1
            If loopCount > maxLoop Then GoTo SafeExit
        Loop
        If r = 1 And Me.Cells(r, startCol).Value <> "BÁO CÁO TỔNG HỢP" Then GoTo SafeExit
        If InStr(1, Me.Cells(r, startCol).Value, "NGÀY") > 0 Then
            currentDate = Split(Me.Cells(r, startCol).Value, "NGÀY")(1)
        Else
            currentDate = Me.Cells(r, startCol).Value
        End If
        currentDate = Trim(currentDate)
        If Not IsDate(currentDate) Then GoTo SafeExit
    End If

    If IsDate(currentDate) Then
        formattedDate = Format(CDate(currentDate), "dd/mm/yyyy")
    Else
        formattedDate = CStr(currentDate)
    End If

    firstRow = r
    lastRow = r
    loopCount = 0
    
    ' Sửa lại cách tìm lastRow
    Do While lastRow < Me.Rows.Count
        If loopCount > maxLoop Then Exit Do
        loopCount = loopCount + 1
        
        ' Kiểm tra dòng tiếp theo
        Dim nextRow As Long
        nextRow = lastRow + 1
        
        ' Kiểm tra nếu gặp tiêu đề báo cáo khác
        If Me.Cells(nextRow, startCol).Value Like "*BÁO CÁO TỔNG HỢP*" Then
            Exit Do
        End If
        
        ' Kiểm tra có dữ liệu trong các cột quan trọng không
        If Not IsEmpty(Me.Cells(nextRow, 1).Value) Or _
           Not IsEmpty(Me.Cells(nextRow, 2).Value) Or _
           Not IsEmpty(Me.Cells(nextRow, 3).Value) Or _
           Not IsEmpty(Me.Cells(nextRow, 5).Value) Or _
           Not IsEmpty(Me.Cells(nextRow, 6).Value) Then
            lastRow = nextRow
        Else
            ' Kiểm tra 5 dòng tiếp theo xem có dữ liệu không
            Dim hasDataInNextRows As Boolean
            hasDataInNextRows = False
            
            Dim k As Long
            For k = 2 To 5
                If nextRow + k > Me.Rows.Count Then Exit For
                
                If Not IsEmpty(Me.Cells(nextRow + k, 1).Value) Or _
                   Not IsEmpty(Me.Cells(nextRow + k, 2).Value) Or _
                   Not IsEmpty(Me.Cells(nextRow + k, 3).Value) Or _
                   Not IsEmpty(Me.Cells(nextRow + k, 5).Value) Or _
                   Not IsEmpty(Me.Cells(nextRow + k, 6).Value) Then
                    hasDataInNextRows = True
                    Exit For
                End If
            Next k
            
            If Not hasDataInNextRows Then
                Exit Do
            End If
            
            lastRow = nextRow
        End If
    Loop

    With Me.Range(Me.Cells(r, startCol), Me.Cells(r, valueCol))
        .Merge
        .Value = labelSheet.Range("A1").Value & " " & formattedDate
        .Font.Name = "Times New Roman"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(254, 242, 203)
        .Borders.LineStyle = xlContinuous
    End With

    formulas(0) = ""  ' Để trống cho số dư ban đầu tiền mặt
    formulas(1) = ""  ' Để trống cho số dư ban đầu tài khoản
    formulas(2) = "=SUMIFS(C" & firstRow & ":C" & lastRow & ",F" & firstRow & ":F" & lastRow & ",""" & cashLabel & """,E" & firstRow & ":E" & lastRow & ",""Thu"")"
    formulas(3) = "=SUMIFS(C" & firstRow & ":C" & lastRow & ",F" & firstRow & ":F" & lastRow & ",""" & cashLabel & """,E" & firstRow & ":E" & lastRow & ",""Chi"")"
    formulas(4) = "=SUMIFS(C" & firstRow & ":C" & lastRow & ",F" & firstRow & ":F" & lastRow & ",""" & bankLabel & """,E" & firstRow & ":E" & lastRow & ",""Thu"")"
    formulas(5) = "=SUMIFS(C" & firstRow & ":C" & lastRow & ",F" & firstRow & ":F" & lastRow & ",""" & bankLabel & """,E" & firstRow & ":E" & lastRow & ",""Chi"")"
    formulas(6) = "=I" & (r + 3) & "+I" & (r + 5)
    formulas(7) = "=I" & (r + 4) & "+I" & (r + 6)
    formulas(8) = "=I" & (r + 1) & "+I" & (r + 3) & "-I" & (r + 4)
    formulas(9) = "=I" & (r + 2) & "+I" & (r + 5) & "-I" & (r + 6)
    formulas(10) = "=I" & (r + 9) & "+I" & (r + 10)

    For i = 0 To 10
        With Me.Cells(r + 1 + i, startCol)
            .Value = labels(i)
            .Font.Name = "Times New Roman"
            .Font.Size = 15
            .Interior.Color = highlightColors(i)
            .Borders.LineStyle = xlContinuous

            Select Case i
                Case 0 To 7
                    .Font.Bold = False
                    .Font.Color = RGB(0, 0, 0)
                Case 8, 9
                    .Font.Bold = True
                    .Font.Color = RGB(228, 193, 178)
                Case 10
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0)
            End Select
        End With

        With Me.Cells(r + 1 + i, valueCol)
            ' Áp dụng định dạng số cho tất cả các ô trong cột giá trị
            .NumberFormat = "#,##0"
            
            Select Case i
                Case 0  ' Số dư ban đầu tiền mặt
                    If Not IsEmpty(initialCashValue) And Target.Column <> 1 Then
                        .Value = initialCashValue
                    Else
                        .Value = ""
                    End If
                Case 1  ' Số dư ban đầu tài khoản
                    If Not IsEmpty(initialBankValue) And Target.Column <> 1 Then
                        .Value = initialBankValue
                    Else
                        .Value = ""
                    End If
                Case Else
                    If formulas(i) <> "" Then
                        .Formula = formulas(i)
                    Else
                        .Value = ""
                    End If
            End Select
            
            .Font.Name = "Times New Roman"
            .Font.Size = 15
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlCenter
        End With
    Next i

    ' Định dạng số cho cột số tiền (cột C)
    Dim dataRange As Range
    Set dataRange = Me.Range(Me.Cells(firstRow, 3), Me.Cells(lastRow, 3))
    dataRange.NumberFormat = "#,##0"

SafeExit:
    Application.EnableEvents = True
End Sub