Attribute VB_Name = "modCleanDateColumn"
Option Explicit

Public Sub CleanDateColumn()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim col As Long
    Dim lastRow As Long
    Dim r As Long

    Dim oldValue As Variant
    Dim d As Variant

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("CleanedData")

    col = FindColumn(ws, _
        "Date", _
        "Order Date", _
        "Sales Date", _
        "Purchase Date", _
        "Transaction Date")

    If col = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

    For r = 2 To lastRow

        oldValue = ws.Cells(r, col).value

        d = ParseDateValue(oldValue)

        If IsEmpty(d) Then
            d = "Unknown"
        End If

        If CStr(oldValue) <> CStr(d) Then

            ws.Cells(r, col).value = d

            RecordChange _
                ws.Cells(r, col), _
                oldValue, _
                d

        End If

    Next r

    ws.Columns(col).NumberFormat = "dd-mm-yyyy"

    SortDatesWithUnknownLast ws, col

End Sub

