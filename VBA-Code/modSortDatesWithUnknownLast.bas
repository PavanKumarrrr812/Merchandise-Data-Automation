Attribute VB_Name = "modSortDatesWithUnknownLast"
Option Explicit

Public Sub SortDatesWithUnknownLast( _
    ByVal ws As Worksheet, _
    ByVal dateCol As Long)

    Dim lastRow As Long
    Dim lastCol As Long

    If ws Is Nothing Then Exit Sub
    If dateCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, dateCol).End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ws.Sort.SortFields.Clear

    ws.Sort.SortFields.Add _
        key:=ws.Range( _
            ws.Cells(2, dateCol), _
            ws.Cells(lastRow, dateCol)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    With ws.Sort

        .SetRange ws.Range( _
            ws.Cells(1, 1), _
            ws.Cells(lastRow, lastCol))

        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom

        .Apply

    End With

End Sub

