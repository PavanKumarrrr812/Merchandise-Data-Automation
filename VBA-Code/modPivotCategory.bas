Attribute VB_Name = "modPivotCategory"
Option Explicit

Public Sub CreateInventoryByCategory( _
    ByVal wsPivot As Worksheet, _
    ByVal sourceRange As Range)

    Dim pt As PivotTable
    Dim dataField As PivotField

    If gPivotCache Is Nothing Then

        MsgBox "Pivot Cache was not created.", _
               vbCritical, "Pivot Error"

        Exit Sub

    End If

    On Error GoTo ErrorHandler

    'Create the real PivotTable
    Set pt = gPivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("H8"), _
        TableName:="ptInventoryCategory")

    With pt

        .ManualUpdate = True

        'CATEGORY ? ROWS
        With .PivotFields("Category")

            .Orientation = xlRowField
            .Position = 1

        End With

        'INVENTORY VALUE ? VALUES
        Set dataField = .AddDataField( _
            .PivotFields("Inventory Value"), _
            "Total Inventory Value", _
            xlSum)

        .RowAxisLayout xlTabularRow

        .ManualUpdate = False

    End With

    'Format the PivotTable
    FormatRealPivot pt

    'Format the numeric values
    On Error Resume Next

    If Not pt.DataBodyRange Is Nothing Then

        pt.DataBodyRange.NumberFormat = "#,##0.00"

    End If

    On Error GoTo 0

    Exit Sub


ErrorHandler:

    On Error Resume Next

    'If the PivotTable was partially created, remove it
    If Not pt Is Nothing Then
        pt.TableRange2.Clear
    End If

    On Error GoTo 0

    MsgBox _
        "Inventory Value by Category failed." & _
        vbCrLf & vbCrLf & _
        "Error: " & Err.Description & _
        vbCrLf & _
        "Error Number: " & Err.Number, _
        vbCritical, _
        "Pivot Error"

End Sub

