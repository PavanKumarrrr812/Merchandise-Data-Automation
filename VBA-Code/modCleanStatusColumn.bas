Attribute VB_Name = "modCleanStatusColumn"
Option Explicit

Public Sub CleanStatusColumn()

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim col As Long
    Dim lastRow As Long
    Dim r As Long

    Dim oldValue As Variant
    Dim newValue As String

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("CleanedData")

    col = StatusColumn(ws)

    If col = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

    For r = 2 To lastRow

        oldValue = ws.Cells(r, col).value

        newValue = NormalizeStockStatus(oldValue)

        If CStr(oldValue) <> newValue Then

            ws.Cells(r, col).value = newValue

            RecordChange _
                ws.Cells(r, col), _
                oldValue, _
                newValue

        End If

    Next r

End Sub

