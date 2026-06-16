Attribute VB_Name = "Module2"
' ==========================================================
' PJT 附加到 已進單整理 (重寫版)
'   1. 從 PJT 第 2 列開始，向後掃到「實際 A 欄有值」的最後一列
'      (跳過 =BIS!... 公式回傳空字串造成的尾端假資料列)
'   2. 以「值」貼到 已進單整理 末尾
'   3. 依 B 欄 (Cust) 對應填入 AF 欄 (MK)：
'        B = "G.U"  → AF = "MK2"
'        B = "JOE"  → AF = "MK1"
'        B = "DKS"  → AF = "MH3"
'   4. 將 Module1 的格式 (字型/框線/置中/列高/E,K 靠上) 套到附加列
'   執行前提：已先執行 Module1 的 ProcessDataConversion_Full
' ==========================================================

Sub AppendPJT_To_FinishedOrders()
    Dim wb As Workbook
    Dim wsTarget As Worksheet
    Dim wsPJT As Worksheet
    Dim targetLastRow As Long
    Dim pjtLastRow As Long
    Dim startRow As Long
    Dim copyRows As Long
    Dim copyCols As Long
    Dim i As Long
    Dim custVal As String
    Dim aVal As Variant
    Dim appendedRng As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsTarget = wb.Sheets("已進單整理")
    Set wsPJT = wb.Sheets("PJT")
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

    ' 用 End(xlUp) 起步，再往回掃跳過公式回傳空字串的列
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
    copyCols = 73                ' A~BU，對齊 已進單整理 結構
    startRow = targetLastRow + 1

    ' 以「值」貼上 (避免 =BIS!... 位移後抓錯)
    wsTarget.Range(wsTarget.Cells(startRow, 1), _
                   wsTarget.Cells(startRow + copyRows - 1, copyCols)).Value = _
        wsPJT.Range(wsPJT.Cells(2, 1), _
                    wsPJT.Cells(pjtLastRow, copyCols)).Value

    ' AF (col 32) 依 B (col 2 Cust) 對應廠別
    For i = startRow To startRow + copyRows - 1
        custVal = Trim(CStr(wsTarget.Cells(i, 2).Value))
        Select Case UCase(custVal)
            Case "G.U"
                wsTarget.Cells(i, 32).Value = "MK2"
            Case "JOE"
                wsTarget.Cells(i, 32).Value = "MK1"
            Case "DKS"
                wsTarget.Cells(i, 32).Value = "MH3"
        End Select
    Next i

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
