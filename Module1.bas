Attribute VB_Name = "Module1"
' ==========================================================
' 已進單整理
'   - 將「已進單底稿」依欄位對應整理為「已進單整理」
'   - CA 欄新增 PROGRAM CATEGORY (來源：已進單底稿 E 欄)
'   - AG 欄 (MK外發) VLOOKUP 上週排單
'   - BD 欄 (未核可) VLOOKUP 上週排單
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

        ws2.Cells(targetRow, 1).Value = ws1.Cells(i, 1).Value    ' A  <- A  GROUP
        ws2.Cells(targetRow, 2).Value = ws1.Cells(i, 2).Value    ' B  <- B  Cust
        ws2.Cells(targetRow, 3).Value = ws1.Cells(i, 3).Value    ' C  <- C  Sub_Group
        ws2.Cells(targetRow, 4).Value = ws1.Cells(i, 6).Value    ' D  <- F  Index
        ws2.Cells(targetRow, 5).Value = ws1.Cells(i, 9).Value    ' E  <- I  F_Style
        ws2.Cells(targetRow, 6).Value = ws1.Cells(i, 10).Value   ' F  <- J  C_Style
        ws2.Cells(targetRow, 7).Value = ws1.Cells(i, 11).Value   ' G  <- K  PCS
        ws2.Cells(targetRow, 8).Value = ws1.Cells(i, 13).Value   ' H  <- M  CPO_QTY
        ws2.Cells(targetRow, 9).Value = ws1.Cells(i, 14).Value   ' I  <- N  CPO_QTYIE
        ws2.Cells(targetRow, 10).Value = ws1.Cells(i, 17).Value  ' J  <- Q  FOB
        ws2.Cells(targetRow, 11).Value = ws1.Cells(i, 15).Value  ' K  <- O  CPO
        ws2.Cells(targetRow, 12).Value = ws1.Cells(i, 18).Value  ' L  <- R
        ws2.Cells(targetRow, 13).Value = ws1.Cells(i, 19).Value  ' M  <- S
        ws2.Cells(targetRow, 14).Value = ws1.Cells(i, 20).Value  ' N  <- T  EXP
        ws2.Cells(targetRow, 15).Value = ws1.Cells(i, 21).Value  ' O  <- U
        ws2.Cells(targetRow, 16).Value = ws1.Cells(i, 22).Value  ' P  <- V
        ws2.Cells(targetRow, 17).Value = ws1.Cells(i, 24).Value  ' Q  <- X
        ws2.Cells(targetRow, 18).Value = ws1.Cells(i, 26).Value  ' R  <- Z
        ws2.Cells(targetRow, 19).Value = ws1.Cells(i, 27).Value  ' S  <- AA
        ws2.Cells(targetRow, 20).Value = ws1.Cells(i, 28).Value  ' T  <- AB
        ws2.Cells(targetRow, 21).Value = ws1.Cells(i, 31).Value  ' U  <- AE
        ws2.Cells(targetRow, 22).Value = ws1.Cells(i, 29).Value  ' V  <- AC
        ws2.Cells(targetRow, 23).Value = ws1.Cells(i, 33).Value  ' W  <- AG  PPM2
        ws2.Cells(targetRow, 24).Value = ws1.Cells(i, 35).Value  ' X  <- AI
        ws2.Cells(targetRow, 25).Value = ws1.Cells(i, 36).Value  ' Y  <- AJ
        ws2.Cells(targetRow, 26).Value = ws1.Cells(i, 34).Value  ' Z  <- AH  CRFP2
        ws2.Cells(targetRow, 27).Value = ws1.Cells(i, 38).Value  ' AA <- AL
        ws2.Cells(targetRow, 28).Value = ws1.Cells(i, 39).Value  ' AB <- AM
        ws2.Cells(targetRow, 29).Value = ws1.Cells(i, 40).Value  ' AC <- AN  IE
        ws2.Cells(targetRow, 30).Value = ws1.Cells(i, 42).Value  ' AD <- AP
        ws2.Cells(targetRow, 31).Value = ws1.Cells(i, 43).Value  ' AE <- AQ
        ws2.Cells(targetRow, 32).Value = ws1.Cells(i, 44).Value  ' AF <- AR  MK
        ' AG (33)、AH (34)、AI (35)、AJ (36)、AK (37) 由公式填入
        ws2.Cells(targetRow, 38).Value = ws1.Cells(i, 45).Value  ' AL <- AS
        ws2.Cells(targetRow, 39).Value = ws1.Cells(i, 46).Value  ' AM <- AT
        ws2.Cells(targetRow, 40).Value = ws1.Cells(i, 75).Value  ' AN <- BW
        ws2.Cells(targetRow, 41).Value = ws1.Cells(i, 47).Value  ' AO <- AU
        ws2.Cells(targetRow, 42).Value = ws1.Cells(i, 48).Value  ' AP <- AV
        ws2.Cells(targetRow, 43).Value = ws1.Cells(i, 49).Value  ' AQ <- AW
        ws2.Cells(targetRow, 44).Value = ws1.Cells(i, 55).Value  ' AR <- BC
        ws2.Cells(targetRow, 45).Value = ws1.Cells(i, 56).Value  ' AS <- BD
        ws2.Cells(targetRow, 46).Value = ws1.Cells(i, 58).Value  ' AT <- BF
        ws2.Cells(targetRow, 47).Value = ws1.Cells(i, 65).Value  ' AU <- BM
        ws2.Cells(targetRow, 48).Value = ws1.Cells(i, 66).Value  ' AV <- BN
        ws2.Cells(targetRow, 49).Value = ws1.Cells(i, 67).Value  ' AW <- BO
        ws2.Cells(targetRow, 50).Value = ws1.Cells(i, 51).Value  ' AX <- AY
        ws2.Cells(targetRow, 51).Value = ws1.Cells(i, 68).Value  ' AY <- BP
        ws2.Cells(targetRow, 52).Value = ws1.Cells(i, 59).Value  ' AZ <- BG
        ws2.Cells(targetRow, 53).Value = ws1.Cells(i, 61).Value  ' BA <- BI
        ws2.Cells(targetRow, 54).Value = ws1.Cells(i, 4).Value   ' BB <- D  PROGRAM
        ws2.Cells(targetRow, 55).Value = ws1.Cells(i, 7).Value   ' BC <- G  Status
        ' BD (56) 未核可 由 VLOOKUP 公式填入
        ws2.Cells(targetRow, 57).Value = ws1.Cells(i, 73).Value  ' BE <- BU
        ws2.Cells(targetRow, 58).Value = ws1.Cells(i, 71).Value  ' BF <- BS
        ws2.Cells(targetRow, 59).Value = ws1.Cells(i, 72).Value  ' BG <- BT
        ws2.Cells(targetRow, 60).Value = ws1.Cells(i, 67).Value  ' BH <- BO
        ws2.Cells(targetRow, 61).Value = ws1.Cells(i, 66).Value  ' BI <- BN
        ws2.Cells(targetRow, 62).Value = ws1.Cells(i, 74).Value  ' BJ <- BV
        ws2.Cells(targetRow, 63).Value = ws1.Cells(i, 65).Value  ' BK <- BM
        ws2.Cells(targetRow, 64).Value = ws1.Cells(i, 76).Value  ' BL <- BX
        ws2.Cells(targetRow, 65).Value = ws1.Cells(i, 77).Value  ' BM <- BY
        ws2.Cells(targetRow, 66).Value = ws1.Cells(i, 53).Value  ' BN <- BA
        ws2.Cells(targetRow, 67).Value = ws1.Cells(i, 54).Value  ' BO <- BB
        ws2.Cells(targetRow, 68).Value = ws1.Cells(i, 63).Value  ' BP <- BK
        ws2.Cells(targetRow, 69).Value = ws1.Cells(i, 27).Value  ' BQ <- AA
        ws2.Cells(targetRow, 70).Value = ws1.Cells(i, 28).Value  ' BR <- AB
        ws2.Cells(targetRow, 71).Value = ws1.Cells(i, 8).Value   ' BS <- H  Status2
        ws2.Cells(targetRow, 72).Value = ws1.Cells(i, 29).Value  ' BT <- AC
        ws2.Cells(targetRow, 73).Value = ws1.Cells(i, 81).Value  ' BU <- CC Season
        ws2.Cells(targetRow, 74).Value = ws1.Cells(i, 82).Value  ' BV <- CD
        ws2.Cells(targetRow, 75).Value = ws1.Cells(i, 83).Value  ' BW <- CE
        ws2.Cells(targetRow, 76).Value = ws1.Cells(i, 84).Value  ' BX <- CF
        ws2.Cells(targetRow, 77).Value = ws1.Cells(i, 89).Value  ' BY <- CK
        ws2.Cells(targetRow, 78).Value = ws1.Cells(i, 90).Value  ' BZ <- CL
        ' === 新增欄位 ===
        ws2.Cells(targetRow, 79).Value = ws1.Cells(i, 5).Value   ' CA <- E  PROGRAM CATEGORY

        ws2.Cells(targetRow, 5).Interior.Color = ws1.Cells(i, 9).Interior.Color
        ws2.Cells(targetRow, 7).Interior.Color = ws1.Cells(i, 11).Interior.Color
        ws2.Cells(targetRow, 11).Interior.Color = ws1.Cells(i, 15).Interior.Color
        ws2.Cells(targetRow, 14).Interior.Color = ws1.Cells(i, 20).Interior.Color
        ws2.Cells(targetRow, 26).Interior.Color = ws1.Cells(i, 34).Interior.Color
        ws2.Cells(targetRow, 29).Interior.Color = ws1.Cells(i, 40).Interior.Color
    Next i

    targetLastRow = targetRow

    ' ------------------------------------------------------
    ' 2. 設定第 1 列特殊欄位的標題
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
    ' 3. 插入公式 / 邏輯判斷 (第 2 列 ~ 最末列)
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

        Dim factoryCode As String
        factoryCode = Trim(UCase(ws2.Cells(j, 32).Value))
        Select Case factoryCode
            Case "MK1", "MK2", "MK5", "MH1", "MH2", "MH3"
                ws2.Cells(j, 35).Value = "自製"
            Case Else
                ws2.Cells(j, 35).Value = "外發"
        End Select

        Dim cellE As Range, cellK As Range, cellAJ As Range
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
    ' 4. 格式設定
    ' ------------------------------------------------------
    Dim rng As Range
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

    With ws2.Cells.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 12
    End With

    ws2.Columns("G:J").NumberFormat = "#,##0"

    ws2.Columns.ColumnWidth = 14
    ws2.Rows.RowHeight = 15

    Dim rngA As Range
    Set rngA = ws2.Range("A2:A" & targetLastRow)
    rngA.TextToColumns Destination:=ws2.Range("A2"), _
        DataType:=xlFixedWidth, _
        FieldInfo:=Array(1, xlTextFormat)

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
