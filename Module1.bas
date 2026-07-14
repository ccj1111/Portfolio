Attribute VB_Name = "Module1"
' ==========================================================
' 已進單整理
'   - 將「已進單底稿」依欄位對應整理為「已進單整理」
'   - CA 欄新增 PROGRAM CATEGORY (來源：已進單底稿 E 欄)
'   - AG (MK外發) 用 F_Style + MK 兩鍵在 上週排單 查 AH 值
'        無對應或錯誤時填入 "NA"
'   - BD (未核可) 用 F_Style + MK 兩鍵在 上週排單 查 BE 值
'        無對應時留空
'   - AH (原始工廠) 維持 LEFT 公式 (不轉值)
'   - 字型：Calibri (拉丁) + 微軟正黑體 (中文)
'   - 指定欄位欄寬縮為 0.1
'   - 自動套用篩選 (AutoFilter)
' ==========================================================

Sub ProcessDataConversion_Full()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsLookup As Worksheet
    Dim wb As Workbook
    Dim lastRow As Long, targetLastRow As Long
    Dim i As Long, targetRow As Long, j As Long, k As Long
    Dim factoryCode As String
    Dim cellE As Range, cellK As Range, cellAJ As Range
    Dim rng As Range, rngA As Range
    Dim fontRng As Range
    Dim mkDict As Object, unappDict As Object
    Dim lookupLastRow As Long
    Dim fVal As String, agVal As String, keyStr As String
    Dim lookupKey As String
    Dim sourceVal As Variant

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
    ' 3. 建立 上週排單 查找字典 (key = F_Style & MK)
    '    上週排單 F = F_Style (col 6)、AG = MK (col 33)
    '    回傳 AH = MK外發 (col 34)、BE = 未核可 (col 57)
    ' ------------------------------------------------------
    Set mkDict = CreateObject("Scripting.Dictionary")
    Set unappDict = CreateObject("Scripting.Dictionary")

    If Not wsLookup Is Nothing Then
        lookupLastRow = wsLookup.Cells(wsLookup.Rows.Count, 6).End(xlUp).Row
        For k = 2 To lookupLastRow
            fVal = CStr(wsLookup.Cells(k, 6).Value)
            agVal = CStr(wsLookup.Cells(k, 33).Value)
            keyStr = fVal & Chr(7) & agVal
            If Not mkDict.Exists(keyStr) Then
                mkDict.Add keyStr, wsLookup.Cells(k, 34).Value
                unappDict.Add keyStr, wsLookup.Cells(k, 57).Value
            End If
        Next k
    End If

    ' ------------------------------------------------------
    ' 4. 插入公式 / 邏輯判斷
    ' ------------------------------------------------------
    For j = 2 To targetLastRow

        ' AK: RFP2 YM 公式 (稍後會轉為值)
        ws2.Cells(j, 37).Formula = _
            "=YEAR(Z" & j & ")&IF(MONTH(Z" & j & ")<10,""0""&MONTH(Z" & j & "),MONTH(Z" & j & "))"

        ' AG / BD 查找 (key = E F_Style & AF MK)
        lookupKey = CStr(ws2.Cells(j, 5).Value) & Chr(7) & CStr(ws2.Cells(j, 32).Value)

        ' AG (MK外發) - 失敗或錯誤填 "NA"
        If mkDict.Exists(lookupKey) Then
            sourceVal = mkDict(lookupKey)
            If IsError(sourceVal) Then
                ws2.Cells(j, 33).Value = "NA"
            Else
                ws2.Cells(j, 33).Value = sourceVal
            End If
        Else
            ws2.Cells(j, 33).Value = "NA"
        End If

        ' BD (未核可) - 失敗填 ""
        If unappDict.Exists(lookupKey) Then
            sourceVal = unappDict(lookupKey)
            If IsError(sourceVal) Then
                ws2.Cells(j, 56).Value = ""
            Else
                ws2.Cells(j, 56).Value = sourceVal
            End If
        Else
            ws2.Cells(j, 56).Value = ""
        End If

        ' AH (原始工廠) - 維持 LEFT 公式 (不轉值)
        ws2.Cells(j, 34).Formula = "=LEFT(AG" & j & ",3)"

        ' AI (Factory) - 維持公式：依 AF 欄 (MK) 判斷自製/外發
        ws2.Cells(j, 35).Formula = _
            "=IF(OR(TRIM(AF" & j & ")=""MK1"",TRIM(AF" & j & ")=""MK2""," & _
            "TRIM(AF" & j & ")=""MK5"",TRIM(AF" & j & ")=""MH1""," & _
            "TRIM(AF" & j & ")=""MH2"",TRIM(AF" & j & ")=""MH3"")," & _
            """自製"",""外發"")"

        ' AJ LC_NO 顏色判斷
        Set cellE = ws2.Cells(j, 5)
        Set cellK = ws2.Cells(j, 11)
        Set cellAJ = ws2.Cells(j, 36)
        If cellE.Interior.Color = vbYellow Then cellAJ.Value = "New Order"
        If cellK.Interior.Color = vbYellow Then
            cellAJ.Value = "New CPO"
        ElseIf cellK.Interior.Color = vbRed Then
            cellAJ.Value = "CPO Change"
        End If
    Next j

    ' ------------------------------------------------------
    ' 5. 強制計算後將 AK 轉為值 (AG/BD 已是值，AH 維持公式)
    ' ------------------------------------------------------
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    With ws2.Range(ws2.Cells(2, 37), ws2.Cells(targetLastRow, 37))
        .Value = .Value
    End With

    Application.Calculation = xlCalculationManual

    ' ------------------------------------------------------
    ' 6. 格式設定
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
    Set fontRng = ws2.Range(ws2.Cells(1, 1), ws2.Cells(targetLastRow, 79))

    With fontRng.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 12
    End With

    On Error Resume Next
    fontRng.Font.NameFarEast = "微軟正黑體"
    On Error GoTo 0

    ws2.Columns("G:J").NumberFormat = "#,##0"

    ws2.Columns.ColumnWidth = 14
    ws2.Rows.RowHeight = 15

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
    ' 7. 開啟篩選
    ' ------------------------------------------------------
    If ws2.AutoFilterMode Then ws2.AutoFilterMode = False
    ws2.Range(ws2.Cells(1, 1), ws2.Cells(targetLastRow, 79)).AutoFilter

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
