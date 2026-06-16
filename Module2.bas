Attribute VB_Name = "Module2"
' ==========================================================
' PJT 附加到 已進單整理
'   1. 從 PJT 複製內容（第 2 列 ~ A 欄有值的最後一列）以「值」貼到
'      已進單整理 末尾
'   2. 依 B 欄 (Cust) 對應填入 AF 欄 (MK)：
'        B = "G.U"  → AF = "MK2"
'        B = "JOE"  → AF = "MK1"
'        B = "DKS"  → AF = "MH3"
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

    ' 已進單整理 最後一列 (依 A 欄)
    targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row

    ' PJT A 欄有值的最後一列
    pjtLastRow = wsPJT.Cells(wsPJT.Rows.Count, 1).End(xlUp).Row

    If pjtLastRow < 2 Then
        MsgBox "PJT 沒有可附加的資料。", vbExclamation
        GoTo CleanExit
    End If

    copyRows = pjtLastRow - 1                   ' PJT 第 2 列 ~ 最末列
    copyCols = wsPJT.UsedRange.Columns.Count    ' PJT 實際使用欄位數
    startRow = targetLastRow + 1

    ' 以「值」複製 (避免 PJT 中的 =BIS!... 公式失效)
    wsTarget.Range(wsTarget.Cells(startRow, 1), _
                   wsTarget.Cells(startRow + copyRows - 1, copyCols)).Value = _
        wsPJT.Range(wsPJT.Cells(2, 1), _
                    wsPJT.Cells(pjtLastRow, copyCols)).Value

    ' AF 欄依 B 欄 (Cust) 對應廠別代碼
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

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
