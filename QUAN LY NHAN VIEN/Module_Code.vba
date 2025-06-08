' ===================================================================
' VBA CODE CHO MODULE
' Tạo một Module mới và đặt code này vào
' ===================================================================

Option Explicit

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
