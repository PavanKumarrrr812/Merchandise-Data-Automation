Attribute VB_Name = "modRemoveEmptyColumns"
Option Explicit

Public Sub RemoveEmptyColumns(ByVal ws As Worksheet)

    Dim lastCol As Long
    Dim c As Long

    If ws Is Nothing Then Exit Sub

    lastCol = ws.Cells(1, _
        ws.Columns.count).End(xlToLeft).Column

    For c = lastCol To 1 Step -1

        If Application.WorksheetFunction.CountA( _
            ws.Columns(c)) = 0 Then

            ws.Columns(c).Delete

        End If

    Next c

End Sub

