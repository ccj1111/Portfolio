Attribute VB_Name = "Module1"
' ==========================================================
' 已進單整理
'   - 將「已進單底稿」依欄位對應整理為「已進單整理」
'   - CA 欄新增 PROGRAM CATEGORY (來源：已進單底稿 E 欄)
'   - AG 欄 (MK外發) VLOOKUP 上週排單 -> 計算後轉為值
'   - BD 欄 (未核可) VLOOKUP 上週排單 -> 計算後轉為值
'   - 字型：Calibri (拉丁) + 微軟正黑體 (中文)
'   - 指定欄位欄寬縮為 0.1：
'       D, F, J, L, M, Q~W, AA~AB, AD~AE, AK~AO, AR~BA
'   - 自動套用篩選 (AutoFilter)
' ==========================================================

Sub ProcessDataConversion_Full()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wsLookup As Worksheet
    Dim wb As Workbook
    Dim lastRow As Long
    Dim targetLastRow As Long
    Dim i As Long
    Dim targetRow As Long
    Dim j As Long
    Dim factoryCode As String
    Dim cellE As Range, cellK As Range, cellAJ As Range
    Dim rng As Range
    Dim rngA As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws1 = wb.Sheets("已進單底稿")
    Set wsLookup = wb.Sheets("上週排單")
    On Error GoTo 0

    If ws1 Is Nothing Then
        MsgBox "錯誤：找不到來源工作表「已進單底稿」。", vbCritical
        GoTo CleanExit
    End If

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("已進單整理").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws2 = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws2.Name = "已進單整理"

    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "來源工作表「已進單底稿」沒有有效的資料。", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------
    ' 1. 複製資料並調整欄位順序
    ' ------------------------------------------------------
    For i = 2 To lastRow
        targetRow = i - 1

        ws2.Cells(targetRow, 1).Value = ws1.Cells(i, 1).Value
        ws2.Cells(targetRow, 2).Value = ws1.Cells(i, 2).Value
        ws2.Cells(targetRow, 3).Value = ws1.Cells(i, 3).Value
        ws2.Cells(targetRow, 4).Value = ws1.Cells(i, 6).Value
        ws2.Cells(targetRow, 5).Value = ws1.Cells(i, 9).Value
        ws2.Cells(targetRow, 6).Value = ws1.Cells(i, 10).Value
        ws2.Cells(targetRow, 7).Value = ws1.Cells(i, 11).Value
        ws2.Cells(targetRow, 8).Value = ws1.Cells(i, 13).Value
        ws2.Cells(targetRow, 9).Value = ws1.Cells(i, 14).Value
        ws2.Cells(targetRow, 10).Value = ws1.Cells(i, 17).Value
        ws2.Cells(targetRow, 11).Value = ws1.Cells(i, 15).Value
        ws2.Cells(targetRow, 12).Value = ws1.Cells(i, 18).Value
        ws2.Cells(targetRow, 13).Value = ws1.Cells(i, 19).Value
        ws2.Cells(targetRow, 14).Value = ws1.Cells(i, 20).Value
        ws2.Cells(targetRow, 15).Value = ws1.Cells(i, 21).Value
        ws2.Cells(targetRow, 16).Value = ws1.Cells(i, 22).Value
        ws2.Cells(targetRow, 17).Value = ws1.Cells(i, 24).Value
        ws2.Cells(targetRow, 18).Value = ws1.Cells(i, 26).Value
        ws2.Cells(targetRow, 19).Value = ws1.Cells(i, 27).Value
        ws2.Cells(targetRow, 20).Value = ws1.Cells(i, 28).Value
        ws2.Cells(targetRow, 21).Value = ws1.Cells(i, 31).Value
        ws2.Cells(targetRow, 22).Value = ws1.Cells(i, 29).Value
        ws2.Cells(targetRow, 23).Value = ws1.Cells(i, 33).Value
        ws2.Cells(targetRow, 24).Value = ws1.Cells(i, 35).Value
        ws2.Cells(targetRow, 25).Value = ws1.Cells(i, 36).Value
        ws2.Cells(targetRow, 26).Value = ws1.Cells(i, 34).Value
        ws2.Cells(targetRow, 27).Value = ws1.Cells(i, 38).Value
        ws2.Cells(targetRow, 28).Value = ws1.Cells(i, 39).Value
        ws2.Cells(targetRow, 29).Value = ws1.Cells(i, 40).Value
        ws2.Cells(targetRow, 30).Value = ws1.Cells(i, 42).Value
        ws2.Cells(targetRow, 31).Value = ws1.Cells(i, 43).Value
        ws2.Cells(targetRow, 32).Value = ws1.Cells(i, 44).Value
        ws2.Cells(targetRow, 38).Value = ws1.Cells(i, 45).Value
        ws2.Cells(targetRow, 39).Value = ws1.Cells(i, 46).Value
        ws2.Cells(targetRow, 40).Value = ws1.Cells(i, 75).Value
        ws2.Cells(targetRow, 41).Value = ws1.Cells(i, 47).Value
        ws2.Cells(targetRow, 42).Value = ws1.Cells(i, 48).Value
        ws2.Cells(targetRow, 43).Value = ws1.Cells(i, 49).Value
        ws2.Cells(targetRow, 44).Value = ws1.Cells(i, 55).Value
        ws2.Cells(targetRow, 45).Value = ws1.Cells(i, 56).Value
        ws2.Cells(targetRow, 46).Value = ws1.Cells(i, 58).Value
        ws2.Cells(targetRow, 47).Value = ws1.Cells(i, 65).Value
        ws2.Cells(targetRow, 48).Value = ws1.Cells(i, 66).Value
        ws2.Cells(targetRow, 49).Value = ws1.Cells(i, 67).Value
        ws2.Cells(targetRow, 50).Value = ws1.Cells(i, 51).Value
        ws2.Cells(targetRow, 51).Value = ws1.Cells(i, 68).Value
        ws2.Cells(targetRow, 52).Value = ws1.Cells(i, 59).Value
        ws2.Cells(targetRow, 53).Value = ws1.Cells(i, 61).Value
        ws2.Cells(targetRow, 54).Value = ws1.Cells(i, 4).Value
        ws2.Cells(targetRow, 55).Value = ws1.Cells(i, 7).Value
        ws2.Cells(targetRow, 57).Value = ws1.Cells(i, 73).Value
        ws2.Cells(targetRow, 58).Value = ws1.Cells(i, 71).Value
        ws2.Cells(targetRow, 59).Value = ws1.Cells(i, 72).Value
        ws2.Cells(targetRow, 60).Value = ws1.Cells(i, 67).Value
        ws2.Cells(targetRow, 61).Value = ws1.Cells(i, 66).Value
        ws2.Cells(targetRow, 62).Value = ws1.Cells(i, 74).Value
        ws2.Cells(targetRow, 63).Value = ws1.Cells(i, 65).Value
        ws2.Cells(targetRow, 64).Value = ws1.Cells(i, 76).Value
        ws2.Cells(targetRow, 65).Value = ws1.Cells(i, 77).Value
        ws2.Cells(targetRow, 66).Value = ws1.Cells(i, 53).Value
        ws2.Cells(targetRow, 67).Value = ws1.Cells(i, 54).Value
        ws2.Cells(targetRow, 68).Value = ws1.Cells(i, 63).Value
        ws2.Cells(targetRow, 69).Value = ws1.Cells(i, 27).Value
        ws2.Cells(targetRow, 70).Value = ws1.Cells(i, 28).Value
        ws2.Cells(targetRow, 71).Value = ws1.Cells(i, 8).Value
        ws2.Cells(targetRow, 72).Value = ws1.Cells(i, 29).Value
        ws2.Cells(targetRow, 73).Value = ws1.Cells(i, 81).Value
        ws2.Cells(targetRow, 74).Value = ws1.Cells(i, 82).Value
        ws2.Cells(targetRow, 75).Value = ws1.Cells(i, 83).Value
        ws2.Cells(targetRow, 76).Value = ws1.Cells(i, 84).Value
        ws2.Cells(targetRow, 77).Value = ws1.Cells(i, 89).Value
        ws2.Cells(targetRow, 78).Value = ws1.Cells(i, 90).Value
        ws2.Cells(targetRow, 79).Value = ws1.Cells(i, 5).Value

        ws2.Cells(targetRow, 5).Interior.Color = ws1.Cells(i, 9).Interior.Color
        ws2.Cells(targetRow, 7).Interior.Color = ws1.Cells(i, 11).Interior.Color
        ws2.Cells(targetRow, 11).Interior.Color = ws1.Cells(i, 15).Interior.Color
        ws2.Cells(targetRow, 14).Interior.Color = ws1.Cells(i, 20).Interior.Color
        ws2.Cells(targetRow, 26).Interior.Color = ws1.Cells(i, 34).Interior.Color
        ws2.Cells(targetRow, 29).Interior.Color = ws1.Cells(i, 40).Interior.Color
    Next i

    targetLastRow = targetRow

    ' ------------------------------------------------------
    ' 2. 設定特殊欄位標題
    ' ------------------------------------------------------
    With ws2
        .Cells(1, 33).Value = "MK外發"
        .Cells(1, 34).Value = "原始工廠"
        .Cells(1, 35).Value = "Factory"
        .Cells(1, 36).Value = "LC_NO"
        .Cells(1, 37).Value = "RFP2 YM"
        .Cells(1, 56).Value = "未核可"
        .Cells(1, 79).Value = "PROGRAM CATEGORY"
    End With

    ' ------------------------------------------------------
    ' 3. 插入公式 / 邏輯判斷
    ' ------------------------------------------------------
    For j = 2 To targetLastRow
        ws2.Cells(j, 37).Formula = _
            "=YEAR(Z" & j & ")&IF(MONTH(Z" & j & ")<10,""0""&MONTH(Z" & j & "),MONTH(Z" & j & "))"

        If Not wsLookup Is Nothing Then
            ws2.Cells(j, 33).Formula = _
                "=IFERROR(VLOOKUP(E" & j & "&AF" & j & "&BC" & j & _
                ",上週排單!$A:$AH,34,FALSE),"""")"
            ws2.Cells(j, 56).Formula = _
                "=IFERROR(VLOOKUP(E" & j & "&AF" & j & "&BC" & j & _
                ",上週排單!$A:$BE,57,FALSE),"""")"
        End If

        ws2.Cells(j, 34).Formula = "=LEFT(AG" & j & ",3)"

        factoryCode = Trim(UCase(ws2.Cells(j, 32).Value))
        Select Case factoryCode
            Case "MK1", "MK2", "MK5", "MH1", "MH2", "MH3"
                ws2.Cells(j, 35).Value = "自製"
            Case Else
                ws2.Cells(j, 35).Value = "外發"
        End Select

        Set cellE = ws2.Cells(j, 5)
        Set cellK = ws2.Cells(j, 11)
        Set cellAJ = ws2.Cells(j, 36)

        If cellE.Interior.Color = vbYellow Then
            cellAJ.Value = "New Order"
        End If

        If cellK.Interior.Color = vbYellow Then
            cellAJ.Value = "New CPO"
        ElseIf cellK.Interior.Color = vbRed Then
            cellAJ.Value = "CPO Change"
        End If
    Next j

    ' ------------------------------------------------------
    ' 4. 強制計算後將 VLOOKUP 結果及衍生公式轉為值
    '    AG (33) VLOOKUP / AH (34) LEFT / AK (37) YEAR / BD (56) VLOOKUP
    ' ------------------------------------------------------
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    With ws2.Range(ws2.Cells(2, 33), ws2.Cells(targetLastRow, 33))
        .Value = .Value
    End With
    With ws2.Range(ws2.Cells(2, 34), ws2.Cells(targetLastRow, 34))
        .Value = .Value
    End With
    With ws2.Range(ws2.Cells(2, 37), ws2.Cells(targetLastRow, 37))
        .Value = .Value
    End With
    With ws2.Range(ws2.Cells(2, 56), ws2.Cells(targetLastRow, 56))
        .Value = .Value
    End With

    Application.Calculation = xlCalculationManual

    ' ------------------------------------------------------
    ' 5. 格式設定
    ' ------------------------------------------------------
    Set rng = ws2.UsedRange

    With ws2.Rows(1).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Color = vbBlack
        .Weight = xlThick
    End With

    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlThin
    rng.Borders.Color = vbBlack

    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter

    ws2.Columns(5).VerticalAlignment = xlTop
    ws2.Columns(11).VerticalAlignment = xlTop

    ws2.Range("AG1, AH1, BD1, CA1").Interior.Color = vbYellow

    ' 字型：Calibri (拉丁) + 微軟正黑體 (中文)
    ' 僅套用到資料範圍，避免對整個 Cells 操作太大
    Dim fontRng As Range
    Set fontRng = ws2.Range(ws2.Cells(1, 1), ws2.Cells(targetLastRow, 79))

    With fontRng.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 12
    End With

    ' NameFarEast 在部分 Excel 版本不支援 -> 用錯誤處理避免中斷
    On Error Resume Next
    fontRng.Font.NameFarEast = "微軟正黑體"
    On Error GoTo 0

    ws2.Columns("G:J").NumberFormat = "#,##0"

    ' 預設欄寬 14、列高 15
    ws2.Columns.ColumnWidth = 14
    ws2.Rows.RowHeight = 15

    ' 指定欄位欄寬縮為 0.1
    ws2.Columns("D").ColumnWidth = 0.1
    ws2.Columns("F").ColumnWidth = 0.1
    ws2.Columns("J").ColumnWidth = 0.1
    ws2.Columns("L").ColumnWidth = 0.1
    ws2.Columns("M").ColumnWidth = 0.1
    ws2.Range("Q:W").ColumnWidth = 0.1
    ws2.Range("AA:AB").ColumnWidth = 0.1
    ws2.Range("AD:AE").ColumnWidth = 0.1
    ws2.Range("AK:AO").ColumnWidth = 0.1
    ws2.Range("AR:BA").ColumnWidth = 0.1

    Set rngA = ws2.Range("A2:A" & targetLastRow)
    rngA.TextToColumns Destination:=ws2.Range("A2"), _
        DataType:=xlFixedWidth, _
        FieldInfo:=Array(1, xlTextFormat)

    ' ------------------------------------------------------
    ' 6. 開啟篩選
    ' ------------------------------------------------------
    If ws2.AutoFilterMode Then ws2.AutoFilterMode = False
    ws2.Range(ws2.Cells(1, 1), ws2.Cells(targetLastRow, 79)).AutoFilter

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
