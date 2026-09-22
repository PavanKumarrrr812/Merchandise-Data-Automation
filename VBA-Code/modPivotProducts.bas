Attribute VB_Name = "modPivotProducts"
Option Explicit

Public Sub CreateTop10Products( _
    ByVal wsPivot As Worksheet, _
    ByVal sourceRange As Range)

    Dim pt As PivotTable
    Dim pf As PivotField
    Dim df As PivotField

    If gPivotCache Is Nothing Then Exit Sub

    Set pt = gPivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("E8"), _
        TableName:="ptTop10Products")

    With pt

        .ManualUpdate = True

        Set pf = .PivotFields("Product")

        pf.Orientation = xlRowField
        pf.Position = 1

        Set df = .AddDataField( _
            .PivotFields("Quantity"), _
            "Total Quantity", _
            xlSum)

        pf.AutoSort xlDescending, df.Name

        On Error Resume Next

        pf.PivotFilters.Add _
            Type:=xlTopCount, _
            dataField:=df, _
            Value1:=10

        On Error GoTo 0

        .RowAxisLayout xlTabularRow

        .ManualUpdate = False

    End With

    FormatRealPivot pt

End Sub

