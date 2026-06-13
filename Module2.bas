Attribute VB_Name = "Module2"
' ==========================================================
' PJT 附加
'   - 將 PJT (BIS 轉換結果) 以「值」附加到「已進單整理」末尾
'   - IE = CPO_QTYIE (I) / CPO_QTY (H) → 寫入 AC 欄 (col 29)
'   - 同步補上 AG / BD VLOOKUP、AH / AK 公式與 AI 廠別判斷
'   - PJT 實際有效列數依據 BIS 的最後一列判斷
' 執行前提：必須先執行 Module1 的 ProcessDataConversion_Full
' ==========================================================

Sub AppendPJT_To_FinishedOrders()
    Dim wb As Workbook
    Dim wsTarget As Worksheet
    Dim wsPJT As Worksheet
    Dim wsBIS As Worksheet
    Dim wsLookup As Worksheet
    Dim targetLastRow As Long
    Dim pjtLastRow As Long
    Dim bisLastRow As Long
    Dim startRow As Long
    Dim copyRows As Long
    Dim copyCols As Long
    Dim i As Long
    Dim qty As Variant, qtyIE As Variant
    Dim factoryCode As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsTarget = wb.Sheets("已進單整理")
    Set wsPJT = wb.Sheets("PJT")
    Set wsBIS = wb.Sheets("BIS")
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

    ' 已進單整理 最後一列 (依 A 欄)
    targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row

    ' PJT 公式參照 BIS，因此 PJT 的有效資料 = BIS 實際資料列數
    If Not wsBIS Is Nothing Then
        bisLastRow = wsBIS.Cells(wsBIS.Rows.Count, 1).End(xlUp).Row
        pjtLastRow = bisLastRow
    Else
        pjtLastRow = wsPJT.Cells(wsPJT.Rows.Count, 1).End(xlUp).Row
    End If

    If pjtLastRow < 2 Then
        MsgBox "PJT 沒有可附加的資料。", vbExclamation
        GoTo CleanExit
    End If

    copyRows = pjtLastRow - 1   ' PJT 第 2 列開始為資料
    copyCols = 73               ' A~BU，對齊已進單整理欄位範圍
    startRow = targetLastRow + 1

    ' 將 PJT 公式結果以「值」貼到 已進單整理
    wsTarget.Range(wsTarget.Cells(startRow, 1), _
                   wsTarget.Cells(startRow + copyRows - 1, copyCols)).Value = _
        wsPJT.Range(wsPJT.Cells(2, 1), _
                    wsPJT.Cells(pjtLastRow, copyCols)).Value

    ' 為新附加列填入計算 / 公式
    For i = startRow To startRow + copyRows - 1

        ' BC 若為空則補上 'PJT'
        If Trim(CStr(wsTarget.Cells(i, 55).Value)) = "" Then
            wsTarget.Cells(i, 55).Value = "PJT"
        End If

        ' IE = CPO_QTYIE (I) / CPO_QTY (H)，寫入 AC (col 29)
        qty = wsTarget.Cells(i, 8).Value
        qtyIE = wsTarget.Cells(i, 9).Value
        If IsNumeric(qty) And IsNumeric(qtyIE) Then
            If qty <> 0 Then
                wsTarget.Cells(i, 29).Value = qtyIE / qty
            End If
        End If

        ' AG / BD VLOOKUP，與既有列保持一致
        If Not wsLookup Is Nothing Then
            wsTarget.Cells(i, 33).Formula = _
                "=IFERROR(VLOOKUP(E" & i & "&AF" & i & "&BC" & i & _
                ",上週排單!$A:$AH,34,FALSE),"""")"
            wsTarget.Cells(i, 56).Formula = _
                "=IFERROR(VLOOKUP(E" & i & "&AF" & i & "&BC" & i & _
                ",上週排單!$A:$BE,57,FALSE),"""")"
        End If

        wsTarget.Cells(i, 34).Formula = "=LEFT(AG" & i & ",3)"
        wsTarget.Cells(i, 37).Formula = _
            "=YEAR(Z" & i & ")&IF(MONTH(Z" & i & ")<10,""0""&MONTH(Z" & i & "),MONTH(Z" & i & "))"

        ' AI 廠別判斷
        factoryCode = Trim(UCase(wsTarget.Cells(i, 32).Value))
        Select Case factoryCode
            Case "MK1", "MK2", "MK5", "MH1", "MH2", "MH3"
                wsTarget.Cells(i, 35).Value = "自製"
            Case Else
                wsTarget.Cells(i, 35).Value = "外發"
        End Select
    Next i

    ' 格式套用
    Dim appendedRng As Range
    Set appendedRng = wsTarget.Range(wsTarget.Cells(startRow, 1), _
                                     wsTarget.Cells(startRow + copyRows - 1, copyCols))

    appendedRng.Borders.LineStyle = xlContinuous
    appendedRng.Borders.Weight = xlThin
    appendedRng.Borders.Color = vbBlack
    appendedRng.HorizontalAlignment = xlCenter
    appendedRng.VerticalAlignment = xlCenter

    With appendedRng.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 12
    End With

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
