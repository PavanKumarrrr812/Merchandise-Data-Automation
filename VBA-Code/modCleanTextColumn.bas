Attribute VB_Name = "modCleanTextColumn"
Option Explicit

Public Sub CleanTextColumn(ByVal col As Long)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Dim oldValue As Variant
    Dim v As String

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("CleanedData")

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

        'Category, Supplier and Warehouse
        'are not allowed to contain digits.

        If col = CategoryColumn(ws) _
           Or col = SupplierColumn(ws) _
           Or col = WarehouseColumn(ws) Then

            If ContainsDigit(v) Then
                v = "Unknown"
            End If

        End If

        If CStr(oldValue) <> v Then

            ws.Cells(r, col).NumberFormat = "@"

            ws.Cells(r, col).value = v

            RecordChange _
                ws.Cells(r, col), _
                oldValue, _
                v

        End If

        If Len(v) > 25 Then

            ws.Cells(r, col).Interior.Color = _
                RGB(255, 255, 0)

        End If

    Next r

End Sub


Private Function ContainsDigit( _
    ByVal textValue As String) As Boolean

    Dim i As Long
    Dim ch As String

    ContainsDigit = False

    For i = 1 To Len(textValue)

        ch = Mid$(textValue, i, 1)

        If ch >= "0" And ch <= "9" Then

            ContainsDigit = True

            Exit Function

        End If

    Next i

End Function

