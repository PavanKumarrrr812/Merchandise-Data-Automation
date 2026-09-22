Attribute VB_Name = "modPivotSalesCategory"
Option Explicit

Public Sub CreateSalesByCategory( _
    ByVal wsPivot As Worksheet, _
    ByVal sourceRange As Range)

    Dim pt As PivotTable
    Dim dataField As PivotField

    Dim tbl As ListObject
    Dim col As ListColumn

    Dim categoryHeader As String
    Dim salesHeader As String

    If gPivotCache Is Nothing Then
        MsgBox "Pivot Cache was not created.", _
               vbCritical, "Sales Category Error"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Set tbl = sourceRange.Worksheet.ListObjects(1)

    'Get exact headers from the Excel Table
    For Each col In tbl.ListColumns

        If LCase(Trim(CStr(col.Name))) = "category" Then
            categoryHeader = col.Name
        End If

        If InStr( _
            1, _
            LCase(Trim(CStr(col.Name))), _
            "sales value", _
            vbTextCompare) > 0 Then

            salesHeader = col.Name

        End If

    Next col

    If categoryHeader = "" Then

        MsgBox "Category column was not found in CleanedData.", _
               vbCritical, "Sales Category Error"

        Exit Sub

    End If

    If salesHeader = "" Then

        MsgBox "Sales Value column was not found in CleanedData.", _
               vbCritical, "Sales Category Error"

        Exit Sub

    End If

    'Create real PivotTable
    Set pt = gPivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("H45"), _
        TableName:="ptSalesCategory")

    With pt

        .ManualUpdate = True

        'CATEGORY ? ROWS
        With .PivotFields(categoryHeader)

            .Orientation = xlRowField
            .Position = 1

        End With

        'SALES VALUE ? VALUES
        Set dataField = .AddDataField( _
            .PivotFields(salesHeader), _
            "Total Sales Value", _
            xlSum)

        .RowAxisLayout xlTabularRow

        .ManualUpdate = False

    End With

    'Format PivotTable
    FormatRealPivot pt

    On Error Resume Next

    If Not pt.DataBodyRange Is Nothing Then
        pt.DataBodyRange.NumberFormat = "#,##0.00"
    End If

    On Error GoTo 0

    Exit Sub


ErrorHandler:

    MsgBox _
        "Sales by Category failed." & _
        vbCrLf & vbCrLf & _
        "Category header: " & categoryHeader & _
        vbCrLf & _
        "Sales header: " & salesHeader & _
        vbCrLf & vbCrLf & _
        "Error: " & Err.Description & _
        vbCrLf & _
        "Error Number: " & Err.Number, _
        vbCritical, _
        "Sales Category Error"

End Sub

