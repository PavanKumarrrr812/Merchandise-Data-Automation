Attribute VB_Name = "modCreateInventoryValue"
Option Explicit

Public Sub CreateInventoryValue()

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim qtyCol As Long
    Dim priceCol As Long
    Dim invCol As Long

    Dim lastRow As Long
    Dim r As Long

    Dim qty As Double
    Dim price As Double
    Dim v As Double

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("CleanedData")

    invCol = FindColumn(ws, _
        "Inventory Value", _
        "InventoryValue")

    If invCol = 0 Then

        invCol = ws.Cells(1, _
            ws.Columns.count).End(xlToLeft).Column + 1

        ws.Cells(1, invCol).value = _
            "Inventory Value"

    End If

    qtyCol = QuantityColumn(ws)

    priceCol = FindColumn(ws, _
        "Unit Price", _
        "Price")

    If qtyCol = 0 Or priceCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, qtyCol).End(xlUp).Row

    For r = 2 To lastRow

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

        ws.Cells(r, invCol).value = v

    Next r

    ws.Range( _
        ws.Cells(2, invCol), _
        ws.Cells(lastRow, invCol) _
    ).NumberFormat = "#,##0.00"

End Sub

