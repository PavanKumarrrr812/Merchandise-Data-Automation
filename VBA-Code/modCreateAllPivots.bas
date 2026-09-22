Attribute VB_Name = "modCreateAllPivots"
Option Explicit

Public gPivotCache As PivotCache

Public Sub CreateAllPivots()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim tbl As ListObject

    Set wb = ThisWorkbook

    If Not PivotSheetExists("CleanedData") Then
        MsgBox "CleanedData sheet was not found.", _
               vbExclamation, "Missing Data"
        Exit Sub
    End If

    Set wsData = wb.Worksheets("CleanedData")

    If wsData.ListObjects.count = 0 Then
        MsgBox "No Excel Table was found in CleanedData.", _
               vbExclamation, "Missing Table"
        Exit Sub
    End If

    Set tbl = wsData.ListObjects(1)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Delete previous PivotTables sheet
    If PivotSheetExists("PivotTables") Then
        wb.Worksheets("PivotTables").Delete
    End If

    Application.DisplayAlerts = True

    'Create PivotTables sheet
    Set wsPivot = wb.Worksheets.Add( _
                    After:=wb.Worksheets(wb.Worksheets.count))

    wsPivot.Name = "PivotTables"

    'Create ONE shared PivotCache
    Set gPivotCache = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=tbl.Range.Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlR1C1, _
            External:=True))

    'Report title
    With wsPivot.Range("B1:L1")

        .Merge
        .value = "MERCHANDISE ANALYTICS - PIVOT REPORT"

        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)

        .Interior.Color = RGB(25, 35, 55)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    wsPivot.Rows(1).RowHeight = 30

    'KPI section
    CreateKPISection wsPivot, wsData.UsedRange

    'REAL PIVOTTABLES
    CreateQuantityByWarehouse wsPivot, wsData.UsedRange
    CreateTop10Products wsPivot, wsData.UsedRange
    CreateInventoryByCategory wsPivot, wsData.UsedRange
    CreateStockStatusPivot wsPivot, wsData.UsedRange

    CreateMonthlyTrend wsPivot, wsData.UsedRange
    CreatePurchaseBySupplier wsPivot, wsData.UsedRange
    CreateSalesByCategory wsPivot, wsData.UsedRange

    'Formatting
    wsPivot.Columns("A:L").ColumnWidth = 18

    wsPivot.Columns("B:L").VerticalAlignment = xlCenter

    Application.ScreenUpdating = True

    wsPivot.Activate
    wsPivot.Range("B2").Select

    Set gPivotCache = Nothing

    MsgBox "All PivotTables created successfully.", _
           vbInformation, "Pivot Report"

End Sub


Public Function PivotSheetExists( _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    PivotSheetExists = False

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If Not ws Is Nothing Then
        PivotSheetExists = True
    End If

End Function

