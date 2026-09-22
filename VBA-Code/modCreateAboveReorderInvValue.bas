Attribute VB_Name = "modCreateAboveReorderInvValue"
Option Explicit

Public Sub CreateAboveReorderInventoryValue()

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim qtyCol As Long
    Dim priceCol As Long
    Dim resultCol As Long

    Dim lastRow As Long
    Dim r As Long

    Dim qty As Double
    Dim price As Double
    Dim oldValue As Variant
    Dim v As Double

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("CleanedData")

    qtyCol = FindColumn(ws, _
        "Above Reorder Quantity", _
        "AboveReorderQuantity")

    priceCol = FindColumn(ws, _
        "Unit Price", _
        "Price")

    If qtyCol = 0 Or priceCol = 0 Then Exit Sub

    resultCol = FindColumn(ws, _
        "Above Reorder Inventory Value", _
        "AboveReorderInventoryValue")

    If resultCol = 0 Then

        resultCol = ws.Cells(1, _
            ws.Columns.count).End(xlToLeft).Column + 1

        ws.Cells(1, resultCol).value = _
            "Above Reorder Inventory Value"

    End If

    lastRow = ws.Cells(ws.Rows.count, qtyCol).End(xlUp).Row

    For r = 2 To lastRow

        oldValue = ws.Cells(r, resultCol).value

        If IsNumeric(ws.Cells(r, qtyCol).value) Then
            qty = CDbl(ws.Cells(r, qtyCol).value)
        Else
            qty = 0
        End If

        If IsNumeric(ws.Cells(r, priceCol).value) Then
            price = CDbl(ws.Cells(r, priceCol).value)
        Else
            price = 0
        End If

        v = qty * price

        If v < 0 Then v = 0

        ws.Cells(r, resultCol).value = v

        If CStr(oldValue) <> CStr(v) Then

            RecordChange _
                ws.Cells(r, resultCol), _
                oldValue, _
                v

        End If

    Next r

    ws.Range( _
        ws.Cells(2, resultCol), _
        ws.Cells(lastRow, resultCol) _
    ).NumberFormat = "#,##0.00"

End Sub

