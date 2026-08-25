Attribute VB_Name = "Module2"
' ==========================================================
' PJT 附加到 已進單整理
'   1. 從 PJT 第 2 列 ~ A 欄有值的最後一列以「值」貼上
'   2. A 欄每格前加 ' (強制文字)
'   3. AF (col 32) 填入 MK，兩段式查找：
'        (a) 上週排單 同款號 (F_Style) + 客戶 (Cust) + 數量 (CPO_QTY)
'            回傳 上週排單 AG (MK)
'        (b) 找不到才用 PJT CB (col 80，值來自 BIS!K) 當 fallback
'        只在 Status2 (BS/col 71) = "待確認" 時執行
'   4. AC (IE, col 29) 顯示 2 位小數
'   5. 套用 Module1 相同格式
'   執行前提：已先執行 Module1 的 ProcessDataConversion_Full
' ==========================================================

Sub AppendPJT_To_FinishedOrders()
    Dim wb As Workbook
    Dim wsTarget As Worksheet
    Dim wsPJT As Worksheet
    Dim wsLookup As Worksheet
    Dim targetLastRow As Long
    Dim pjtLastRow As Long
    Dim lookupLastRow As Long
    Dim startRow As Long
    Dim copyRows As Long
    Dim copyCols As Long
    Dim i As Long, k As Long
    Dim aVal As Variant
    Dim cbVal As Variant
    Dim mkVal As Variant
    Dim s2Val As String
    Dim pjtRowIdx As Long
    Dim appendedRng As Range
    Dim mkDict As Object
    Dim lkey As String
    Dim fs As String, cs As String
    Dim tFS As String, tCust As String
    Dim qt As Variant, tQty As Variant
    Dim mkAssigned As Boolean

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsTarget = wb.Sheets("已進單整理")
    Set wsPJT = wb.Sheets("PJT")
    Set wsLookup = wb.Sheets("上週排單")
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

    ' PJT 真正最後一列：End(xlUp) 起步後往回掃過公式空字串列
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
    copyCols = 79                ' A~CA (含 PROGRAM CATEGORY)
    startRow = targetLastRow + 1

    ' ==========================================
    ' 建立 上週排單 查找字典
    '   key = F_Style (F/col 6) + Cust (C/col 3) + CPO_QTY (I/col 9)
    '   value = MK (AG/col 33)
    '   (上週排單 A 欄是 Vlookup 鍵，所以欄位序比 已進單整理 多 1)
    ' ==========================================
    Set mkDict = CreateObject("Scripting.Dictionary")
    If Not wsLookup Is Nothing Then
        lookupLastRow = wsLookup.Cells(wsLookup.Rows.Count, 6).End(xlUp).Row
        For k = 2 To lookupLastRow
            fs = CStr(wsLookup.Cells(k, 6).Value)
            cs = CStr(wsLookup.Cells(k, 3).Value)
            qt = wsLookup.Cells(k, 9).Value
            If fs <> "" And cs <> "" And IsNumeric(qt) Then
                lkey = fs & Chr(7) & cs & Chr(7) & CStr(qt)
                If Not mkDict.Exists(lkey) Then
                    mkDict.Add lkey, wsLookup.Cells(k, 33).Value
                End If
            End If
        Next k
    End If

    ' 以「值」貼上 (避免 =BIS!... 公式位移失效)
    wsTarget.Range(wsTarget.Cells(startRow, 1), _
                   wsTarget.Cells(startRow + copyRows - 1, copyCols)).Value = _
        wsPJT.Range(wsPJT.Cells(2, 1), _
                    wsPJT.Cells(pjtLastRow, copyCols)).Value

    ' A 欄強制文字格式，並在每格值前加 ' (與手動輸入 '202608 同效)
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

        ' AF (col 32) 兩段式查找 - 只在 Status2 = "待確認"
        pjtRowIdx = i - startRow + 2
        s2Val = Trim(CStr(wsTarget.Cells(i, 71).Value))

        If s2Val = "待確認" Then
            mkAssigned = False

            ' 優先: 上週排單 (F_Style + Cust + CPO_QTY)
            tFS = CStr(wsTarget.Cells(i, 5).Value)   ' E = F_Style
            tCust = CStr(wsTarget.Cells(i, 2).Value) ' B = Cust
            tQty = wsTarget.Cells(i, 8).Value        ' H = CPO_QTY
            If tFS <> "" And tCust <> "" And IsNumeric(tQty) Then
                lkey = tFS & Chr(7) & tCust & Chr(7) & CStr(tQty)
                If mkDict.Exists(lkey) Then
                    mkVal = mkDict(lkey)
                    If Not IsError(mkVal) And Not IsEmpty(mkVal) Then
                        If CStr(mkVal) <> "" Then
                            wsTarget.Cells(i, 32).Value = mkVal
                            mkAssigned = True
                        End If
                    End If
                End If
            End If

            ' Fallback: PJT CB (BIS Maker)
            If Not mkAssigned Then
                cbVal = wsPJT.Cells(pjtRowIdx, 80).Value
                If Not IsError(cbVal) And Not IsEmpty(cbVal) Then
                    wsTarget.Cells(i, 32).Value = cbVal
                End If
            End If
        End If
    Next i

    ' IE (AC, col 29) 顯示 2 位小數
    wsTarget.Range(wsTarget.Cells(startRow, 29), _
                   wsTarget.Cells(startRow + copyRows - 1, 29)).NumberFormat = "0.00"

    ' ==========================================
    ' 套用 Module1 同款格式到附加列
    ' ==========================================
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

    ' E (5) 和 K (11) 欄靠上 (與 Module1 一致)
    wsTarget.Range(wsTarget.Cells(startRow, 5), _
                   wsTarget.Cells(startRow + copyRows - 1, 5)).VerticalAlignment = xlTop
    wsTarget.Range(wsTarget.Cells(startRow, 11), _
                   wsTarget.Cells(startRow + copyRows - 1, 11)).VerticalAlignment = xlTop

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
