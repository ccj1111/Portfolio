Attribute VB_Name = "Module2"
' ==========================================================
' PJT 附加到 已進單整理
'   AF (MK) 邏輯完全複製 BIS轉換 公式，改用 VBA Dictionary 實作
'   對應：BIS轉換    ? 已進單整理
'         A (GROUP)  ? A (col 1)
'         B (Cust)   ? B (col 2)
'         G (Status) ? BC (col 55)
'         H (Status2)? BS (col 71)
'         I (Style)  ? E (col 5)
'         N (CPO_IE) ? I (col 9, CPO_QTYIE)
'         AU (BIS Maker) ? BIS.K (col 11) 直接讀取
'
'   查找優先序 (不再限定 Status2 = 待確認)：
'     1) 上週移出!A:AY 查 Status+GROUP+Style+CPO_IE  ? 若 = "移出" 就填 "移出"
'     2) B = "G.U"   → 上周排單!K:AZ 用 Style 查      → 抓 AZ (col 52)
'     3) Status2 = "有產詢有Maker" → 上周排單!A:AZ 用 Status+GROUP+Style 查 → 抓 AZ
'     4) 其他         → 上周排單!B:AZ 用 Status+GROUP+Style+CPO_IE 查 → 抓 AZ
'     每層失敗 → fallback: BIS.K (BIS Maker)
'   執行前提：先執行 Module1 的 ProcessDataConversion_Full
' ==========================================================

Sub AppendPJT_To_FinishedOrders()
    Dim wb As Workbook
    Dim wsTarget As Worksheet
    Dim wsPJT As Worksheet
    Dim wsLookup As Worksheet
    Dim wsMoveOut As Worksheet
    Dim wsBIS As Worksheet
    Dim targetLastRow As Long
    Dim pjtLastRow As Long
    Dim startRow As Long
    Dim copyRows As Long
    Dim copyCols As Long
    Dim i As Long, k As Long
    Dim aVal As Variant
    Dim appendedRng As Range
    Dim pjtRowIdx As Long

    Dim dictMoveOut As Object, dictGU As Object
    Dim dictProdMaker As Object, dictDefault As Object
    Dim lookupLastRow As Long
    Dim tmpVal As Variant, azVal As Variant
    Dim keyStr As String

    ' 已進單整理 該列的值
    Dim grpVal As String, custVal As String, styleVal As String
    Dim ieVal As String, statusVal As String, status2Val As String
    Dim bisMk As Variant, bisMkStr As String
    Dim assigned As Boolean

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate                    ' 先讓所有 =BIS!... 公式算好
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsTarget = wb.Sheets("已進單整理")
    Set wsPJT = wb.Sheets("PJT")
    Set wsBIS = wb.Sheets("BIS")
    ' 兩種 週/周 用字都試
    Set wsLookup = wb.Sheets("上週排單")
    If wsLookup Is Nothing Then Set wsLookup = wb.Sheets("上周排單")
    Set wsMoveOut = wb.Sheets("上週移出")
    If wsMoveOut Is Nothing Then Set wsMoveOut = wb.Sheets("上周移出")
    On Error GoTo 0

    If wsTarget Is Nothing Then
        MsgBox "找不到「已進單整理」，請先執行 Module1 的 ProcessDataConversion_Full。", vbCritical
        GoTo CleanExit
    End If
    If wsPJT Is Nothing Then
        MsgBox "找不到「PJT」工作表。", vbCritical
        GoTo CleanExit
    End If

    targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row

    ' PJT 真正最後一列：往回掃過公式回傳空字串的列
    pjtLastRow = wsPJT.Cells(wsPJT.Rows.Count, 1).End(xlUp).Row
    Do While pjtLastRow >= 2
        aVal = wsPJT.Cells(pjtLastRow, 1).Value
        If Not IsEmpty(aVal) Then
            If Trim(CStr(aVal)) <> "" Then Exit Do
        End If
        pjtLastRow = pjtLastRow - 1
    Loop

    If pjtLastRow < 2 Then
        MsgBox "PJT 沒有可附加的資料。", vbExclamation
        GoTo CleanExit
    End If

    copyRows = pjtLastRow - 1
    copyCols = 79
    startRow = targetLastRow + 1

    ' ==========================================
    ' 建立 4 個 Dictionary 對應公式的四層查找
    ' ==========================================
    Set dictMoveOut = CreateObject("Scripting.Dictionary")
    Set dictGU = CreateObject("Scripting.Dictionary")
    Set dictProdMaker = CreateObject("Scripting.Dictionary")
    Set dictDefault = CreateObject("Scripting.Dictionary")

    ' (1) 上週移出：key = col A，value = col AY (51)
    If Not wsMoveOut Is Nothing Then
        lookupLastRow = wsMoveOut.Cells(wsMoveOut.Rows.Count, 1).End(xlUp).Row
        For k = 2 To lookupLastRow
            keyStr = CStr(wsMoveOut.Cells(k, 1).Value)
            If keyStr <> "" And Not dictMoveOut.Exists(keyStr) Then
                dictMoveOut.Add keyStr, wsMoveOut.Cells(k, 51).Value
            End If
        Next k
    End If

    ' (2)(3)(4) 上週排單：三個 dict 都指向 col AZ (52) 當回傳
    '   dictGU        key = col K (11)
    '   dictProdMaker key = col A (1)
    '   dictDefault   key = col B (2)
    If Not wsLookup Is Nothing Then
        lookupLastRow = wsLookup.Cells(wsLookup.Rows.Count, 1).End(xlUp).Row
        For k = 2 To lookupLastRow
            azVal = wsLookup.Cells(k, 52).Value

            keyStr = CStr(wsLookup.Cells(k, 11).Value)
            If keyStr <> "" And Not dictGU.Exists(keyStr) Then dictGU.Add keyStr, azVal

            keyStr = CStr(wsLookup.Cells(k, 1).Value)
            If keyStr <> "" And Not dictProdMaker.Exists(keyStr) Then dictProdMaker.Add keyStr, azVal

            keyStr = CStr(wsLookup.Cells(k, 2).Value)
            If keyStr <> "" And Not dictDefault.Exists(keyStr) Then dictDefault.Add keyStr, azVal
        Next k
    End If

    ' ==========================================
    ' 貼上 PJT 資料
    ' ==========================================
    wsTarget.Range(wsTarget.Cells(startRow, 1), _
                   wsTarget.Cells(startRow + copyRows - 1, copyCols)).Value = _
        wsPJT.Range(wsPJT.Cells(2, 1), _
                    wsPJT.Cells(pjtLastRow, copyCols)).Value

    ' A 欄強制文字並前綴 '
    wsTarget.Range(wsTarget.Cells(startRow, 1), _
                   wsTarget.Cells(startRow + copyRows - 1, 1)).NumberFormat = "@"

    For i = startRow To startRow + copyRows - 1

        ' A 欄前加 '
        aVal = wsTarget.Cells(i, 1).Value
        If Not IsEmpty(aVal) Then
            If CStr(aVal) <> "" Then
                wsTarget.Cells(i, 1).Value = "'" & CStr(aVal)
            End If
        End If

        pjtRowIdx = i - startRow + 2

        ' ---- 讀該列各欄 ----
        grpVal = CStr(wsTarget.Cells(i, 1).Value)
        ' 去掉 A 欄前加的 ' 以便 concat 對得起來 (VBA 讀出的 .Value 本來就不含 ')
        If Left(grpVal, 1) = "'" Then grpVal = Mid(grpVal, 2)
        custVal = CStr(wsTarget.Cells(i, 2).Value)
        styleVal = CStr(wsTarget.Cells(i, 5).Value)
        ieVal = CStr(wsTarget.Cells(i, 9).Value)         ' CPO_QTYIE
        statusVal = CStr(wsTarget.Cells(i, 55).Value)    ' BC = Status
        status2Val = CStr(wsTarget.Cells(i, 71).Value)   ' BS = Status2

        ' ---- 讀 BIS Maker (fallback 用) ----
        bisMkStr = ""
        If Not wsBIS Is Nothing Then
            bisMk = wsBIS.Cells(pjtRowIdx, 11).Value    ' BIS!K
            If Not IsError(bisMk) Then
                If Not IsEmpty(bisMk) Then
                    If Trim(CStr(bisMk)) <> "" Then bisMkStr = CStr(bisMk)
                End If
            End If
        End If

        assigned = False

        ' ==== Step 1: 上週移出 ====
        keyStr = statusVal & grpVal & styleVal & ieVal
        If dictMoveOut.Exists(keyStr) Then
            tmpVal = dictMoveOut(keyStr)
            If Not IsError(tmpVal) Then
                If CStr(tmpVal) = "移出" Then
                    wsTarget.Cells(i, 32).Value = "移出"
                    assigned = True
                End If
            End If
        End If

        If Not assigned Then
            ' ==== Step 2: B = "G.U" ====
            If custVal = "G.U" Then
                If dictGU.Exists(styleVal) Then
                    tmpVal = dictGU(styleVal)
                    If Not IsError(tmpVal) And Not IsEmpty(tmpVal) Then
                        If Trim(CStr(tmpVal)) <> "" Then
                            wsTarget.Cells(i, 32).Value = tmpVal
                            assigned = True
                        End If
                    End If
                End If
                If Not assigned And bisMkStr <> "" Then
                    wsTarget.Cells(i, 32).Value = bisMkStr
                    assigned = True
                End If

            ' ==== Step 3: Status2 = "有產詢有Maker" ====
            ElseIf status2Val = "有產詢有Maker" Then
                keyStr = statusVal & grpVal & styleVal
                If dictProdMaker.Exists(keyStr) Then
                    tmpVal = dictProdMaker(keyStr)
                    If Not IsError(tmpVal) And Not IsEmpty(tmpVal) Then
                        If Trim(CStr(tmpVal)) <> "" Then
                            wsTarget.Cells(i, 32).Value = tmpVal
                            assigned = True
                        End If
                    End If
                End If
                If Not assigned And bisMkStr <> "" Then
                    wsTarget.Cells(i, 32).Value = bisMkStr
                    assigned = True
                End If

            ' ==== Step 4: 其他 (預設) ====
            Else
                keyStr = statusVal & grpVal & styleVal & ieVal
                If dictDefault.Exists(keyStr) Then
                    tmpVal = dictDefault(keyStr)
                    If Not IsError(tmpVal) And Not IsEmpty(tmpVal) Then
                        If Trim(CStr(tmpVal)) <> "" Then
                            wsTarget.Cells(i, 32).Value = tmpVal
                            assigned = True
                        End If
                    End If
                End If
                If Not assigned And bisMkStr <> "" Then
                    wsTarget.Cells(i, 32).Value = bisMkStr
                    assigned = True
                End If
            End If
        End If

    Next i

    ' IE (AC, col 29) 顯示 2 位小數
    wsTarget.Range(wsTarget.Cells(startRow, 29), _
                   wsTarget.Cells(startRow + copyRows - 1, 29)).NumberFormat = "0.00"

    ' 套用 Module1 同款格式到附加列
    Set appendedRng = wsTarget.Range(wsTarget.Cells(startRow, 1), _
                                     wsTarget.Cells(startRow + copyRows - 1, copyCols))
    With appendedRng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = vbBlack
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 15
    End With
    With appendedRng.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 12
    End With
    wsTarget.Range(wsTarget.Cells(startRow, 5), _
                   wsTarget.Cells(startRow + copyRows - 1, 5)).VerticalAlignment = xlTop
    wsTarget.Range(wsTarget.Cells(startRow, 11), _
                   wsTarget.Cells(startRow + copyRows - 1, 11)).VerticalAlignment = xlTop

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
