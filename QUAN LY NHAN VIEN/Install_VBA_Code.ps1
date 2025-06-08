# ===================================================================
# SCRIPT POWERSHELL TO INSTALL VBA CODE INTO EXCEL
# ===================================================================

param(
    [string]$ExcelFilePath = "QUAN LY NHA NGHI.xlsm"
)

# Check if Excel file exists
if (-not (Test-Path $ExcelFilePath)) {
    Write-Host "Error: Cannot find file $ExcelFilePath" -ForegroundColor Red
    exit 1
}

Write-Host "Installing VBA code into Excel file..." -ForegroundColor Green

try {
    # Create Excel Application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Open workbook
    $workbook = $excel.Workbooks.Open((Resolve-Path $ExcelFilePath).Path)

    # Find sheet "TONG HOP THU - CHI NHA NGHI"
    $targetSheet = $null
    foreach ($sheet in $workbook.Worksheets) {
        if ($sheet.Name -eq "TONG HOP THU - CHI NHA NGHI") {
            $targetSheet = $sheet
            break
        }
    }

    if ($targetSheet -eq $null) {
        Write-Host "Warning: Cannot find sheet 'TONG HOP THU - CHI NHA NGHI'" -ForegroundColor Yellow
        Write-Host "Will create code for first sheet..." -ForegroundColor Yellow
        $targetSheet = $workbook.Worksheets.Item(1)
    }

    # VBA Code for Worksheet_Change event
    $worksheetCode = @"
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim allowedSheets As String
    Dim ws As Worksheet
    Dim sheetName As String
    Dim col As Long
    Dim row As Long
    Dim valueC As Variant
    Dim value As Variant

    ' Turn off automatic calculation and events to avoid loop
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler

    ' Check allowed sheets
    allowedSheets = "TONG HOP THU - CHI NHA NGHI"
    Set ws = Target.Worksheet
    sheetName = ws.Name

    If InStr(allowedSheets, sheetName) = 0 Then
        GoTo CleanUp
    End If

    ' Only process when changing 1 cell
    If Target.Cells.Count > 1 Then
        GoTo CleanUp
    End If

    col = Target.Column
    row = Target.row

    ' === PART 1: Write timestamp to column D when entering amount in column C ===
    If col = 3 Then ' Column C
        valueC = ws.Cells(row, 3).Value
        If valueC <> "" And IsNumeric(valueC) Then
            ws.Cells(row, 4).Value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        Else
            ws.Cells(row, 4).Value = ""
        End If
    End If

    ' === PART 2: Create report table when entering date in column A ===
    If col = 1 Then ' Column A
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
    MsgBox "Error: " + Err.Description, vbCritical
End Sub

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

    startCol = 8 ' Column H
    valueCol = 9 ' Column I
    defaultEndRow = 1000

    ' Label array
    labels = Array( _
        "So du ban dau tien mat:", _
        "So du ban dau tai khoan:", _
        "Thu tien mat:", _
        "Chi tien mat:", _
        "Thu tai khoan:", _
        "Chi tai khoan:", _
        "Tong thu:", _
        "Tong chi:", _
        "Tien mat hien co:", _
        "Tai khoan hien co:", _
        "Tong tien hien co:" _
    )

    ' Background color array (RGB)
    highlightColors = Array( _
        RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), _
        RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), RGB(197, 224, 179), _
        RGB(84, 129, 53), RGB(84, 129, 53), RGB(1, 176, 80) _
    )

    ' If value is empty, clear report table
    If IsEmpty(value) Or value = "" Then
        ws.Range(ws.Cells(row, startCol), ws.Cells(row + UBound(labels) + 1, valueCol)).Clear
        Exit Sub
    End If

    ' Create title
    Set titleCell = ws.Range(ws.Cells(row, startCol), ws.Cells(row, valueCol))
    titleCell.Merge

    If IsDate(value) Then
        formattedDate = Format(value, "dd/mm/yyyy")
    Else
        formattedDate = CStr(value)
    End If

    With titleCell
        .Value = "BAO CAO TONG HOP - NGAY " + formattedDate
        .Font.Name = "Times New Roman"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(254, 242, 203)
        .Borders.LineStyle = xlContinuous
    End With

    ' Formula array
    formulas = Array( _
        "", _
        "", _
        "=SUMIFS(C" + CStr(row) + ":C" + CStr(defaultEndRow) + ",F" + CStr(row) + ":F" + CStr(defaultEndRow) + ",""Tien mat"",E" + CStr(row) + ":E" + CStr(defaultEndRow) + ",""Thu"")", _
        "=SUMIFS(C" + CStr(row) + ":C" + CStr(defaultEndRow) + ",F" + CStr(row) + ":F" + CStr(defaultEndRow) + ",""Tien mat"",E" + CStr(row) + ":E" + CStr(defaultEndRow) + ",""Chi"")", _
        "=SUMIFS(C" + CStr(row) + ":C" + CStr(defaultEndRow) + ",F" + CStr(row) + ":F" + CStr(defaultEndRow) + ",""Chuyen khoan"",E" + CStr(row) + ":E" + CStr(defaultEndRow) + ",""Thu"")", _
        "=SUMIFS(C" + CStr(row) + ":C" + CStr(defaultEndRow) + ",F" + CStr(row) + ":F" + CStr(defaultEndRow) + ",""Chuyen khoan"",E" + CStr(row) + ":E" + CStr(defaultEndRow) + ",""Chi"")", _
        "=I" + CStr(row + 3) + "+I" + CStr(row + 5), _
        "=I" + CStr(row + 4) + "+I" + CStr(row + 6), _
        "=I" + CStr(row + 1) + "+I" + CStr(row + 3) + "-I" + CStr(row + 4), _
        "=I" + CStr(row + 2) + "+I" + CStr(row + 5) + "-I" + CStr(row + 6), _
        "=I" + CStr(row + 9) + "+I" + CStr(row + 10) _
    )

    ' Create data rows
    For i = 0 To UBound(labels)
        Set labelCell = ws.Cells(row + 1 + i, startCol)
        Set valueCellRange = ws.Cells(row + 1 + i, valueCol)

        ' Set label
        labelCell.Value = labels(i)

        ' Set formula or empty value
        If formulas(i) <> "" Then
            valueCellRange.Formula = formulas(i)
        Else
            valueCellRange.Value = ""
        End If

        ' Format label
        With labelCell
            .Font.Name = "Times New Roman"
            .Font.Size = 15
            .VerticalAlignment = xlCenter
            .Interior.Color = highlightColors(i)
            .Borders.LineStyle = xlContinuous

            ' Format font based on condition
            If i <= 7 Then
                .Font.Bold = False
                .Font.Color = RGB(0, 0, 0) ' Black
            ElseIf i = 8 Or i = 9 Then
                .Font.Bold = True
                .Font.Color = RGB(228, 193, 178) ' Light brown
            ElseIf i = 10 Then
                .Font.Bold = True
                .Font.Color = RGB(0, 0, 0) ' Black
            End If
        End With

        ' Format value cell
        With valueCellRange
            .Font.Name = "Times New Roman"
            .Font.Size = 15
            .VerticalAlignment = xlCenter
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle = xlContinuous
        End With

        ' Set border for both cells
        ws.Range(ws.Cells(row + 1 + i, startCol), ws.Cells(row + 1 + i, valueCol)).Borders.LineStyle = xlContinuous
    Next i
End Sub
"@

    # VBA Code for ThisWorkbook
    $workbookCode = @"
Private Sub Workbook_Open()
    Call DenDongCuoiCung
End Sub
"@

    # VBA Code for Module
    $moduleCode = @"
Sub DenDongCuoiCung()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    ws.Cells(lastRow + 20, 1).Select
End Sub
"@

    # Install code into Worksheet
    Write-Host "Installing code for Worksheet..." -ForegroundColor Yellow
    $targetSheet.CodeName = $targetSheet.CodeName
    $vbaProject = $workbook.VBProject
    $worksheetModule = $vbaProject.VBComponents.Item($targetSheet.CodeName)
    $worksheetModule.CodeModule.DeleteLines(1, $worksheetModule.CodeModule.CountOfLines)
    $worksheetModule.CodeModule.AddFromString($worksheetCode)

    # Install code into ThisWorkbook
    Write-Host "Installing code for ThisWorkbook..." -ForegroundColor Yellow
    $thisWorkbook = $vbaProject.VBComponents.Item("ThisWorkbook")
    $thisWorkbook.CodeModule.DeleteLines(1, $thisWorkbook.CodeModule.CountOfLines)
    $thisWorkbook.CodeModule.AddFromString($workbookCode)

    # Create new Module and install code
    Write-Host "Creating new Module..." -ForegroundColor Yellow
    $newModule = $vbaProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
    $newModule.Name = "TienIchModule"
    $newModule.CodeModule.AddFromString($moduleCode)

    # Save file
    Write-Host "Saving file..." -ForegroundColor Yellow
    $workbook.Save()

    Write-Host "VBA code installation successful!" -ForegroundColor Green
    Write-Host "Functions added:" -ForegroundColor Cyan
    Write-Host "  1. Auto timestamp when entering amount (column C -> column D)" -ForegroundColor White
    Write-Host "  2. Auto create report table when entering date (column A)" -ForegroundColor White
    Write-Host "  3. Auto move to last row when opening file" -ForegroundColor White

} catch {
    Write-Host "Error installing VBA code: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Close Excel
    if ($workbook) {
        $workbook.Close()
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

Write-Host "Complete!" -ForegroundColor Green
Write-Host "Please reopen the Excel file to test the new functions." -ForegroundColor Yellow
