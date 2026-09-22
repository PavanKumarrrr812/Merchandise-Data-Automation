Attribute VB_Name = "modCleanPriceValue"
Option Explicit

Public Sub CleanPriceValueColumn(ByVal col As Long)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Dim oldValue As Variant
    Dim v As Variant
    Dim numericValue As Double

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("CleanedData")

    If col = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

    For r = 2 To lastRow

        oldValue = ws.Cells(r, col).value

        If IsNumeric(Trim$(CStr(oldValue))) Then

            numericValue = CDbl(Trim$(CStr(oldValue)))

            If numericValue < 0 Then

                ws.Cells(r, col).Interior.Color = _
                    RGB(255, 255, 0)

                v = 0

            Else

                v = numericValue

            End If

        ElseIf Trim$(CStr(oldValue)) = "" Then

            v = 0

        Else

            v = WordToNumber(CStr(oldValue))

            If IsError(v) Then
                v = 0
            End If

        End If

        If CStr(oldValue) <> CStr(v) Then

            ws.Cells(r, col).value = v

            RecordChange _
                ws.Cells(r, col), _
                oldValue, _
                v

        End If

    Next r

End Sub

