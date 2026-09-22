Attribute VB_Name = "modQuantityOutliers"
Option Explicit

Public Sub HighlightQuantityOutliers( _
    ByVal ws As Worksheet, _
    ByVal col As Long)

    Dim lastRow As Long
    Dim r As Long

    Dim avgQty As Double
    Dim count As Long
    Dim v As Double

    If col = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

    For r = 2 To lastRow

        If IsNumeric(ws.Cells(r, col).value) Then

            avgQty = avgQty + _
                CDbl(ws.Cells(r, col).value)

            count = count + 1

        End If

    Next r

    If count = 0 Then Exit Sub

    avgQty = avgQty / count

    If avgQty >= 100 And avgQty <= 1000 Then

        For r = 2 To lastRow

            If IsNumeric(ws.Cells(r, col).value) Then

                v = CDbl(ws.Cells(r, col).value)

                If v < 100 Or v > 1000 Then

                    ws.Cells(r, col).Interior.Color = _
                        RGB(255, 255, 0)

                End If

            End If

        Next r

    End If

End Sub

