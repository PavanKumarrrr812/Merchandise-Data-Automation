Attribute VB_Name = "modPivotMonthly"
Option Explicit

Public Sub CreateMonthlyTrend( _
    ByVal wsPivot As Worksheet, _
    ByVal sourceRange As Range)

    Dim pt As PivotTable
    Dim dataField As PivotField

    Dim tbl As ListObject
    Dim col As ListColumn

    Dim monthHeader As String
    Dim salesHeader As String

    If gPivotCache Is Nothing Then
        MsgBox "Pivot Cache was not created.", _
               vbCritical, "Monthly Sales Error"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Set tbl = sourceRange.Worksheet.ListObjects(1)

    'Get the EXACT headers from the Excel Table
    For Each col In tbl.ListColumns

        If LCase(Trim(CStr(col.Name))) = "month" Then
            monthHeader = col.Name
        End If

        If InStr(1, _
                 LCase(Trim(CStr(col.Name))), _
                 "sales value", _
                 vbTextCompare) > 0 Then

            salesHeader = col.Name

        End If

    Next col

    If monthHeader = "" Then

        MsgBox "Month column was not found in CleanedData.", _
               vbCritical, "Monthly Sales Error"

        Exit Sub

    End If

    If salesHeader = "" Then

        MsgBox "Sales Value column was not found in CleanedData.", _
               vbCritical, "Monthly Sales Error"

        Exit Sub

    End If

    'Create real PivotTable
    Set pt = gPivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("H22"), _
        TableName:="ptMonthlySales")

    With pt

        .ManualUpdate = True

        'MONTH ? ROWS
        With .PivotFields(monthHeader)

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

    'Format values
    On Error Resume Next

    If Not pt.DataBodyRange Is Nothing Then
        pt.DataBodyRange.NumberFormat = "#,##0.00"
    End If

    On Error GoTo 0

    Exit Sub


ErrorHandler:

    MsgBox _
        "Monthly Sales Trend failed." & _
        vbCrLf & vbCrLf & _
        "Sales header detected: " & salesHeader & _
        vbCrLf & _
        "Month header detected: " & monthHeader & _
        vbCrLf & vbCrLf & _
        "Error: " & Err.Description & _
        vbCrLf & _
        "Error Number: " & Err.Number, _
        vbCritical, _
        "Monthly Sales Error"

End Sub

