Attribute VB_Name = "modPivotWarehouse"
Option Explicit

Public Sub CreateQuantityByWarehouse( _
    ByVal wsPivot As Worksheet, _
    ByVal sourceRange As Range)

    Dim pt As PivotTable
    Dim df As PivotField

    If gPivotCache Is Nothing Then Exit Sub

    Set pt = gPivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("B8"), _
        TableName:="ptWarehouse")

    With pt

        .ManualUpdate = True

        With .PivotFields("Warehouse")
            .Orientation = xlRowField
            .Position = 1
        End With

        Set df = .AddDataField( _
            .PivotFields("Quantity"), _
            "Total Quantity", _
            xlSum)

        .RowAxisLayout xlTabularRow

        .ManualUpdate = False

    End With

    FormatRealPivot pt

End Sub

