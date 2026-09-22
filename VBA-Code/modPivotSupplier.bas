Attribute VB_Name = "modPivotSupplier"
Option Explicit

Public Sub CreatePurchaseBySupplier( _
    ByVal wsPivot As Worksheet, _
    ByVal sourceRange As Range)

    Dim pt As PivotTable
    Dim dataField As PivotField

    Dim tbl As ListObject
    Dim col As ListColumn

    Dim supplierHeader As String
    Dim purchaseHeader As String

    If gPivotCache Is Nothing Then
        MsgBox "Pivot Cache was not created.", _
               vbCritical, "Purchase Supplier Error"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Set tbl = sourceRange.Worksheet.ListObjects(1)

    'Get exact headers from the Excel Table
    For Each col In tbl.ListColumns

        If LCase(Trim(CStr(col.Name))) = "supplier" Then
            supplierHeader = col.Name
        End If

        If InStr( _
            1, _
            LCase(Trim(CStr(col.Name))), _
            "purchase value", _
            vbTextCompare) > 0 Then

            purchaseHeader = col.Name

        End If

    Next col

    If supplierHeader = "" Then

        MsgBox "Supplier column was not found in CleanedData.", _
               vbCritical, "Purchase Supplier Error"

        Exit Sub

    End If

    If purchaseHeader = "" Then

        MsgBox "Purchase Value column was not found in CleanedData.", _
               vbCritical, "Purchase Supplier Error"

        Exit Sub

    End If

    'Create real PivotTable
    Set pt = gPivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("E22"), _
        TableName:="ptPurchaseSupplier")

    With pt

        .ManualUpdate = True

        'SUPPLIER ? ROWS
        With .PivotFields(supplierHeader)

            .Orientation = xlRowField
            .Position = 1

        End With

        'PURCHASE VALUE ? VALUES
        Set dataField = .AddDataField( _
            .PivotFields(purchaseHeader), _
            "Total Purchase Value", _
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
        "Purchase Value by Supplier failed." & _
        vbCrLf & vbCrLf & _
        "Supplier header: " & supplierHeader & _
        vbCrLf & _
        "Purchase header: " & purchaseHeader & _
        vbCrLf & vbCrLf & _
        "Error: " & Err.Description & _
        vbCrLf & _
        "Error Number: " & Err.Number, _
        vbCritical, _
        "Purchase Supplier Error"

End Sub

