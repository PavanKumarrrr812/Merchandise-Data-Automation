Attribute VB_Name = "modPivotStockStatus"
Option Explicit

Public Sub CreateStockStatusPivot( _
    ByVal wsPivot As Worksheet, _
    ByVal sourceRange As Range)

    Dim pt As PivotTable
    Dim df As PivotField

    If gPivotCache Is Nothing Then Exit Sub

    Set pt = gPivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("K8"), _
        TableName:="ptStockStatus")

    With pt

        .ManualUpdate = True

        With .PivotFields("Stock Status")
            .Orientation = xlRowField
            .Position = 1
        End With

        Set df = .AddDataField( _
            .PivotFields("Product"), _
            "Product Count", _
            xlCount)

        .RowAxisLayout xlTabularRow

        .ManualUpdate = False

    End With

    FormatRealPivot pt

End Sub

