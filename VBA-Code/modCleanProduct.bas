Attribute VB_Name = "modCleanProduct"
Option Explicit

Public Sub CleanProductColumn()

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim col As Long
    Dim lastRow As Long
    Dim r As Long

    Dim oldValue As Variant
    Dim v As String

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("CleanedData")

    col = ProductColumn(ws)

    If col = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

    For r = 2 To lastRow

        oldValue = ws.Cells(r, col).value

        v = CStr(oldValue)

        v = Replace(v, Chr(160), "")
        v = Application.WorksheetFunction.Clean(v)
        v = Trim$(v)

        If v = "" Then
            v = "Unknown"
        End If

        If IsNumeric(v) Then
            v = "Unknown"
        End If

        If CStr(oldValue) <> v Then

            ws.Cells(r, col).value = v

            RecordChange _
                ws.Cells(r, col), _
                oldValue, _
                v

        End If

        If Len(v) > 0 Then

            If IsNumeric(Left$(v, 1)) Then
                ws.Cells(r, col).Interior.Color = RGB(255, 255, 0)
            End If

        End If

        If Len(v) > 25 Then
            ws.Cells(r, col).Interior.Color = RGB(255, 255, 0)
        End If

    Next r

End Sub

