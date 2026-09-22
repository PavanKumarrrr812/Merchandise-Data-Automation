Attribute VB_Name = "modCreateDashboardReport"
Option Explicit

'=========================================================
' DASHBOARD REPORT
'=========================================================

Public Sub CreateDashboardReport()

    Dim wb As Workbook
    Dim wsDash As Worksheet
    Dim wsPivot As Worksheet
    Dim wsData As Worksheet

    Set wb = ThisWorkbook

    If Not DashboardSheetExists("CleanedData") Then
        MsgBox "CleanedData sheet was not found.", vbExclamation
        Exit Sub
    End If

    If Not DashboardSheetExists("PivotTables") Then
        MsgBox "PivotTables sheet was not found.", vbExclamation
        Exit Sub
    End If

    Set wsData = wb.Worksheets("CleanedData")
    Set wsPivot = wb.Worksheets("PivotTables")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    '=====================================================
    ' DELETE OLD DASHBOARD
    '=====================================================

    DeleteDashboardSheet "DashboardReport"
    DeleteDashboardSheet "Dashboard_Report"

    '=====================================================
    ' CREATE NEW DASHBOARD
    '=====================================================

    Set wsDash = wb.Worksheets.Add( _
                    After:=wb.Worksheets(wb.Worksheets.count))

    wsDash.Name = "Dashboard_Report"

    '=====================================================
    ' BASIC SHEET SETTINGS
    '=====================================================

    With wsDash

        .Cells.Clear

        .Cells.Interior.Color = RGB(76, 39, 105)

        .Columns("A:T").ColumnWidth = 10

        .Rows("1:60").RowHeight = 18

    End With

    ActiveWindow.DisplayGridlines = False

    '=====================================================
    ' BACKGROUND
    '=====================================================

    CreateDashboardBackground wsDash

    '=====================================================
    ' KPI HEADER
    '=====================================================

    CreateKPIHeader wsDash, wsData

    '=====================================================
    ' FILTER PANEL
    '=====================================================

    CreateSlicerPanel wsDash

    '=====================================================
    ' CHART PANELS
    '=====================================================

    CreateChartPanel _
        wsDash, _
        210, 125, _
        270, 195, _
        "MONTHLY SALES TREND"

    CreateChartPanel _
        wsDash, _
        490, 125, _
        270, 195, _
        "STOCK STATUS"

    CreateChartPanel _
        wsDash, _
        770, 125, _
        270, 195, _
        "SALES BY CATEGORY"

    CreateChartPanel _
        wsDash, _
        210, 335, _
        380, 215, _
        "WAREHOUSE STOCK"

    CreateChartPanel _
        wsDash, _
        610, 335, _
        430, 215, _
        "TOP 10 PRODUCTS"

    '=====================================================
    ' CREATE CHARTS
    '=====================================================

    CreateMonthlyChart wsDash, wsPivot

    CreateStockStatusChart wsDash, wsPivot

    CreateSalesCategoryChart wsDash, wsPivot

    CreateWarehouseChart wsDash, wsPivot

    CreateTopProductsChart wsDash, wsPivot

    '=====================================================
    ' CREATE SLICERS
    '=====================================================

    CreateDashboardSlicers _
        wb, _
        wsDash, _
        wsPivot

    '=====================================================
    ' FINAL SETTINGS
    '=====================================================

    wsDash.Activate

    ActiveWindow.DisplayGridlines = False

    wsDash.Range("A1").Select

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Dashboard_Report created successfully.", _
           vbInformation, _
           "Dashboard Complete"

    Exit Sub

ErrorHandler:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Dashboard creation stopped." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, _
           "Dashboard Error"

End Sub


'=========================================================
' DASHBOARD BACKGROUND
'=========================================================

Private Sub CreateDashboardBackground( _
    ByVal ws As Worksheet)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                35, 20, _
                1050, 550)

    With shp

        .Name = "DashboardMainBackground"

        .Fill.Visible = msoTrue

        .Fill.ForeColor.RGB = _
            RGB(43, 18, 67)

        .Fill.BackColor.RGB = _
            RGB(115, 55, 145)

        .Fill.TwoColorGradient _
            Style:=msoGradientHorizontal, _
            Variant:=1

        .Line.ForeColor.RGB = _
            RGB(35, 12, 55)

        .Line.Weight = 2

        .Adjustments.Item(1) = 0.08

        .Shadow.Visible = msoFalse

    End With

End Sub


'=========================================================
' KPI HEADER
'=========================================================

Private Sub CreateKPIHeader( _
    ByVal ws As Worksheet, _
    ByVal wsData As Worksheet)

    Dim header As Shape

    Dim totalQuantity As Double
    Dim avgUnitPrice As Double
    Dim salesValue As Double
    Dim inventoryValue As Double
    Dim purchaseValue As Double

    Dim lowStock As Long

    Dim colQuantity As Long
    Dim colUnitPrice As Long
    Dim colSalesValue As Long
    Dim colInventoryValue As Long
    Dim colPurchaseValue As Long
    Dim colStockStatus As Long

    Dim lastRow As Long
    Dim i As Long

    Dim v As Variant

    Dim priceValue As Double
    Dim priceTotal As Double
    Dim priceCount As Long

    '=====================================================
    ' HEADER BACKGROUND
    '=====================================================

    Set header = ws.Shapes.AddShape( _
                    msoShapeRoundedRectangle, _
                    210, 35, _
                    845, 70)

    With header

        .Name = "KPI_Header"

        .Fill.ForeColor.RGB = _
            RGB(45, 91, 145)

        .Line.ForeColor.RGB = _
            RGB(30, 55, 90)

        .Line.Weight = 1.5

        .Adjustments.Item(1) = 0.15

    End With

    '=====================================================
    ' LAST ROW
    '=====================================================

    lastRow = _
        wsData.Cells( _
            wsData.Rows.count, _
            1).End(xlUp).Row

    '=====================================================
    ' FIND COLUMNS
    '=====================================================

    colQuantity = DashboardFindColumnContains( _
                    wsData, _
                    "quantity", _
                    "qty")

    colUnitPrice = _
        DashboardFindUnitPriceColumn(wsData)

    colSalesValue = _
        DashboardFindValueColumn( _
            wsData, _
            "sales value")

    colInventoryValue = _
        DashboardFindValueColumn( _
            wsData, _
            "inventory value")

    colPurchaseValue = _
        DashboardFindValueColumn( _
            wsData, _
            "purchase value")

    colStockStatus = _
        DashboardFindColumnContains( _
            wsData, _
            "stock status", _
            "status")

    '=====================================================
    ' TOTAL QUANTITY
    '=====================================================

    If colQuantity > 0 Then

        totalQuantity = _
            DashboardSumColumn( _
                wsData, _
                colQuantity, _
                lastRow)

    End If

    '=====================================================
    ' TOTAL SALES VALUE
    '=====================================================

    If colSalesValue > 0 Then

        salesValue = _
            DashboardSumColumn( _
                wsData, _
                colSalesValue, _
                lastRow)

    End If

    '=====================================================
    ' TOTAL INVENTORY VALUE
    '=====================================================

    If colInventoryValue > 0 Then

        inventoryValue = _
            DashboardSumColumn( _
                wsData, _
                colInventoryValue, _
                lastRow)

    End If

    '=====================================================
    ' TOTAL PURCHASE VALUE
    '=====================================================

    If colPurchaseValue > 0 Then

        purchaseValue = _
            DashboardSumColumn( _
                wsData, _
                colPurchaseValue, _
                lastRow)

    End If

    '=====================================================
    ' AVERAGE UNIT PRICE
    '=====================================================

    priceTotal = 0

    priceCount = 0

    If colUnitPrice > 0 Then

        For i = 2 To lastRow

            v = wsData.Cells(i, colUnitPrice).value

            If DashboardTryNumber(v, priceValue) Then

                If priceValue >= 0 Then

                    priceTotal = _
                        priceTotal + priceValue

                    priceCount = _
                        priceCount + 1

                End If

            End If

        Next i

    End If

    If priceCount > 0 Then

        avgUnitPrice = _
            priceTotal / priceCount

    Else

        avgUnitPrice = 0

    End If

    '=====================================================
    ' LOW STOCK
    '=====================================================

    lowStock = 0

    If colStockStatus > 0 Then

        For i = 2 To lastRow

            v = _
                wsData.Cells( _
                    i, _
                    colStockStatus).value

            If Not IsError(v) Then

                If InStr( _
                    1, _
                    LCase(CStr(v)), _
                    "low", _
                    vbTextCompare) > 0 Then

                    lowStock = _
                        lowStock + 1

                End If

            End If

        Next i

    End If

    '=====================================================
    ' KPI CARDS
    '=====================================================

    CreateKPICard _
        ws, _
        225, 47, _
        "AVG UNIT PRICE", _
        DashboardFormatMoney(avgUnitPrice)

    CreateKPICard _
        ws, _
        350, 47, _
        "LOW STOCK ITEMS", _
        Format(lowStock, "#,##0")

    CreateKPICard _
        ws, _
        475, 47, _
        "TOTAL SALES VALUE", _
        DashboardFormatCompactMoney(salesValue)

    CreateKPICard _
        ws, _
        600, 47, _
        "TOTAL STOCK QTY", _
        DashboardFormatCompactNumber(totalQuantity)

    CreateKPICard _
        ws, _
        725, 47, _
        "TOTAL INV VALUE", _
        DashboardFormatCompactMoney(inventoryValue)

    CreateKPICard _
        ws, _
        850, 47, _
        "TOTAL PURCHASE VALUE", _
        DashboardFormatCompactMoney(purchaseValue)

End Sub


'=========================================================
' KPI CARD
'=========================================================

Private Sub CreateKPICard( _
    ByVal ws As Worksheet, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal titleText As String, _
    ByVal valueText As String)

    Dim card As Shape
    Dim txt As Shape

    Set card = ws.Shapes.AddShape( _
                    msoShapeRectangle, _
                    leftPos, _
                    topPos, _
                    110, _
                    40)

    With card

        .Fill.ForeColor.RGB = _
            RGB(255, 255, 255)

        .Line.ForeColor.RGB = _
            RGB(190, 200, 215)

        .Line.Weight = 1

    End With

    Set txt = ws.Shapes.AddTextbox( _
                    msoTextOrientationHorizontal, _
                    leftPos + 3, _
                    topPos + 2, _
                    104, _
                    36)

    With txt

        .Line.Visible = msoFalse

        .Fill.Visible = msoFalse

        .TextFrame2.TextRange.Text = _
            titleText & vbCrLf & _
            valueText

        .TextFrame2.TextRange.Font.Name = _
            "Calibri"

        .TextFrame2.TextRange.Font.Size = 8

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(25, 35, 55)

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle

    End With

End Sub


'=========================================================
' CHART PANEL
'=========================================================

Private Sub CreateChartPanel( _
    ByVal ws As Worksheet, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal panelWidth As Double, _
    ByVal panelHeight As Double, _
    ByVal titleText As String)

    Dim panel As Shape
    Dim titleBox As Shape

    Set panel = ws.Shapes.AddShape( _
                    msoShapeRoundedRectangle, _
                    leftPos, _
                    topPos, _
                    panelWidth, _
                    panelHeight)

    With panel

        .Fill.ForeColor.RGB = _
            RGB(45, 91, 145)

        .Line.ForeColor.RGB = _
            RGB(30, 55, 90)

        .Line.Weight = 1.5

        .Adjustments.Item(1) = 0.12

    End With

    Set titleBox = ws.Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        leftPos + 10, _
                        topPos + 5, _
                        panelWidth - 20, _
                        25)

    With titleBox

        .Line.Visible = msoFalse

        .Fill.Visible = msoFalse

        .TextFrame2.TextRange.Text = _
            titleText

        .TextFrame2.TextRange.Font.Name = _
            "Calibri"

        .TextFrame2.TextRange.Font.Size = 11

        .TextFrame2.TextRange.Font.Bold = _
            msoTrue

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 153, 0)

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

    End With

End Sub


'=========================================================
' MONTHLY CHART
'=========================================================

Private Sub CreateMonthlyChart( _
    ByVal ws As Worksheet, _
    ByVal wsPivot As Worksheet)

    Dim pt As PivotTable

    Set pt = GetDashboardPivot( _
                wsPivot, _
                "ptMonthlySales", _
                "ptMonthly")

    If pt Is Nothing Then Exit Sub

    CreatePivotChart _
        ws, _
        pt, _
        "chtMonthlySales", _
        220, _
        150, _
        250, _
        155, _
        xlLineMarkers

End Sub


'=========================================================
' STOCK STATUS CHART
'=========================================================

Private Sub CreateStockStatusChart( _
    ByVal ws As Worksheet, _
    ByVal wsPivot As Worksheet)

    Dim pt As PivotTable

    Set pt = GetDashboardPivot( _
                wsPivot, _
                "ptStockStatus")

    If pt Is Nothing Then Exit Sub

    CreatePivotChart _
        ws, _
        pt, _
        "chtStockStatus", _
        500, _
        150, _
        250, _
        155, _
        xlDoughnut

End Sub


'=========================================================
' SALES CATEGORY CHART
'=========================================================

Private Sub CreateSalesCategoryChart( _
    ByVal ws As Worksheet, _
    ByVal wsPivot As Worksheet)

    Dim pt As PivotTable

    Set pt = GetDashboardPivot( _
                wsPivot, _
                "ptSalesCategory")

    If pt Is Nothing Then Exit Sub

    CreatePivotChart _
        ws, _
        pt, _
        "chtSalesCategory", _
        780, _
        150, _
        250, _
        155, _
        xlColumnClustered

End Sub


'=========================================================
' WAREHOUSE CHART
'=========================================================

Private Sub CreateWarehouseChart( _
    ByVal ws As Worksheet, _
    ByVal wsPivot As Worksheet)

    Dim pt As PivotTable

    Set pt = GetDashboardPivot( _
                wsPivot, _
                "ptWarehouse")

    If pt Is Nothing Then Exit Sub

    CreatePivotChart _
        ws, _
        pt, _
        "chtWarehouse", _
        220, _
        360, _
        360, _
        165, _
        xlColumnClustered

End Sub


'=========================================================
' TOP PRODUCTS CHART
'=========================================================

Private Sub CreateTopProductsChart( _
    ByVal ws As Worksheet, _
    ByVal wsPivot As Worksheet)

    Dim pt As PivotTable

    Set pt = GetDashboardPivot( _
                wsPivot, _
                "ptTop10Products")

    If pt Is Nothing Then Exit Sub

    CreatePivotChart _
        ws, _
        pt, _
        "chtTopProducts", _
        620, _
        360, _
        410, _
        165, _
        xlBarClustered

End Sub


'=========================================================
' PIVOT CHART
'=========================================================

Private Sub CreatePivotChart( _
    ByVal ws As Worksheet, _
    ByVal pt As PivotTable, _
    ByVal chartName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal chartWidth As Double, _
    ByVal chartHeight As Double, _
    ByVal chartType As XlChartType)

    Dim co As ChartObject
    Dim ch As Chart
    Dim ser As Series

    Dim i As Long
    Dim pointCount As Long

    Set co = ws.ChartObjects.Add( _
                leftPos, _
                topPos, _
                chartWidth, _
                chartHeight)

    co.Name = chartName

    Set ch = co.Chart

    With ch

        .SetSourceData Source:=pt.TableRange1

        .chartType = chartType

        .HasTitle = False

        .HasLegend = False

        '=================================================
        ' CHART BACKGROUND
        '=================================================

        .ChartArea.Format.Fill.Visible = msoTrue

        .ChartArea.Format.Fill.ForeColor.RGB = _
            RGB(45, 91, 145)

        .ChartArea.Format.Line.ForeColor.RGB = _
            RGB(45, 91, 145)

        .PlotArea.Format.Fill.Visible = msoTrue

        .PlotArea.Format.Fill.ForeColor.RGB = _
            RGB(45, 91, 145)

        .PlotArea.Format.Line.Visible = msoFalse

        '=================================================
        ' AXES
        '=================================================

        On Error Resume Next

        .Axes(xlCategory).TickLabels.Font.Color = _
            RGB(255, 255, 255)

        .Axes(xlValue).TickLabels.Font.Color = _
            RGB(255, 255, 255)

        .Axes(xlCategory).TickLabels.Font.Size = 8

        .Axes(xlValue).TickLabels.Font.Size = 8

        .Axes(xlCategory).MajorGridlines.Format.Line.ForeColor.RGB = _
            RGB(190, 210, 230)

        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = _
            RGB(190, 210, 230)

        On Error GoTo 0

    End With

    '=====================================================
    ' SERIES COLORS
    '=====================================================

    If ch.SeriesCollection.count > 0 Then

        For Each ser In ch.SeriesCollection

            On Error Resume Next

            pointCount = ser.Points.count

            '=================================================
            ' COLUMN / BAR / DOUGHNUT
            '=================================================

            If chartType = xlColumnClustered _
               Or chartType = xlBarClustered _
               Or chartType = xlDoughnut Then

                For i = 1 To pointCount

                    With ser.Points(i)

                        .Format.Fill.Visible = msoTrue

                        .Format.Fill.ForeColor.RGB = _
                            DashboardChartColor(i)

                        .Format.Line.Visible = msoTrue

                        .Format.Line.ForeColor.RGB = _
                            RGB(255, 255, 255)

                        .Format.Line.Weight = 0.75

                    End With

                Next i

            Else

                '=================================================
                ' MONTHLY LINE
                '=================================================

                With ser.Format.Line

                    .Visible = msoTrue

                    .ForeColor.RGB = _
                        RGB(46, 204, 113)

                    .Weight = 2.5

                End With

                With ser

                    .MarkerStyle = _
                        xlMarkerStyleCircle

                    .MarkerSize = 6

                    .MarkerBackgroundColor = _
                        RGB(46, 204, 113)

                    .MarkerForegroundColor = _
                        RGB(255, 255, 255)

                End With

            End If

            '=================================================
            ' DATA LABELS
            '=================================================

            ser.ApplyDataLabels

            With ser.DataLabels

                .Font.Color = _
                    RGB(255, 255, 255)

                .Font.Size = 8

                .Font.Bold = False

            End With

            On Error GoTo 0

        Next ser

    End If

End Sub


'=========================================================
' CHART COLOR PALETTE
'=========================================================

Private Function DashboardChartColor( _
    ByVal index As Long) As Long

    Select Case ((index - 1) Mod 12) + 1

        Case 1
            DashboardChartColor = RGB(52, 152, 219)

        Case 2
            DashboardChartColor = RGB(46, 204, 113)

        Case 3
            DashboardChartColor = RGB(241, 196, 15)

        Case 4
            DashboardChartColor = RGB(230, 126, 34)

        Case 5
            DashboardChartColor = RGB(231, 76, 60)

        Case 6
            DashboardChartColor = RGB(26, 188, 156)

        Case 7
            DashboardChartColor = RGB(155, 89, 182)

        Case 8
            DashboardChartColor = RGB(52, 73, 94)

        Case 9
            DashboardChartColor = RGB(22, 160, 133)

        Case 10
            DashboardChartColor = RGB(243, 156, 18)

        Case 11
            DashboardChartColor = RGB(41, 128, 185)

        Case 12
            DashboardChartColor = RGB(192, 57, 43)

    End Select

End Function


'=========================================================
' SLICER PANEL
'=========================================================

Private Sub CreateSlicerPanel( _
    ByVal ws As Worksheet)

    Dim panel As Shape
    Dim titleBox As Shape

    Set panel = ws.Shapes.AddShape( _
                    msoShapeRoundedRectangle, _
                    50, _
                    125, _
                    145, _
                    425)

    With panel

        .Name = "SlicerPanel"

        .Fill.ForeColor.RGB = _
            RGB(45, 91, 145)

        .Line.ForeColor.RGB = _
            RGB(30, 55, 90)

        .Line.Weight = 1.5

        .Adjustments.Item(1) = 0.1

    End With

    Set titleBox = ws.Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        60, _
                        130, _
                        125, _
                        25)

    With titleBox

        .Line.Visible = msoFalse

        .Fill.Visible = msoFalse

        .TextFrame2.TextRange.Text = _
            "FILTERS"

        .TextFrame2.TextRange.Font.Name = _
            "Calibri"

        .TextFrame2.TextRange.Font.Size = 12

        .TextFrame2.TextRange.Font.Bold = _
            msoTrue

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 153, 0)

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

    End With

End Sub


'=========================================================
' CREATE SLICERS
'=========================================================

Private Sub CreateDashboardSlicers( _
    ByVal wb As Workbook, _
    ByVal wsDash As Worksheet, _
    ByVal wsPivot As Worksheet)

    Dim pt As PivotTable

    Dim scMonth As SlicerCache
    Dim scCategory As SlicerCache
    Dim scWarehouse As SlicerCache

    Dim sl As Slicer

    '=====================================================
    ' GET A PIVOT
    '=====================================================

    Set pt = GetDashboardPivot( _
                wsPivot, _
                "ptMonthlySales")

    If pt Is Nothing Then

        Set pt = GetDashboardPivot( _
                    wsPivot, _
                    "ptMonthly")

    End If

    If pt Is Nothing Then

        Set pt = GetDashboardPivot( _
                    wsPivot, _
                    "ptWarehouse")

    End If

    If pt Is Nothing Then Exit Sub

    '=====================================================
    ' DELETE OLD SLICER CACHES
    '=====================================================

    DeleteSlicerCache _
        wb, _
        "DashboardMonthCache"

    DeleteSlicerCache _
        wb, _
        "DashboardCategoryCache"

    DeleteSlicerCache _
        wb, _
        "DashboardWarehouseCache"

    '=====================================================
    ' MONTH SLICER
    '=====================================================

    On Error Resume Next

    Set scMonth = wb.SlicerCaches.Add2( _
                        pt, _
                        "Month")

    On Error GoTo 0

    If Not scMonth Is Nothing Then

        Set sl = scMonth.Slicers.Add(wsDash)

        With sl

            .Name = "DashboardMonthSlicer"

            .Caption = "MONTH"

            .Left = 60

            .Top = 160

            .Width = 125

            .Height = 105

        End With

        ConnectSlicerToAllPivots _
            scMonth, _
            wsPivot

    End If

    '=====================================================
    ' CATEGORY SLICER
    '=====================================================

    Set scCategory = Nothing

    On Error Resume Next

    Set scCategory = wb.SlicerCaches.Add2( _
                        pt, _
                        "Category")

    On Error GoTo 0

    If Not scCategory Is Nothing Then

        Set sl = scCategory.Slicers.Add(wsDash)

        With sl

            .Name = "DashboardCategorySlicer"

            .Caption = "CATEGORY"

            .Left = 60

            .Top = 275

            .Width = 125

            .Height = 105

        End With

        ConnectSlicerToAllPivots _
            scCategory, _
            wsPivot

    End If

    '=====================================================
    ' WAREHOUSE SLICER
    '=====================================================

    Set scWarehouse = Nothing

    On Error Resume Next

    Set scWarehouse = wb.SlicerCaches.Add2( _
                        pt, _
                        "Warehouse")

    On Error GoTo 0

    If Not scWarehouse Is Nothing Then

        Set sl = scWarehouse.Slicers.Add(wsDash)

        With sl

            .Name = "DashboardWarehouseSlicer"

            .Caption = "WAREHOUSE"

            .Left = 60

            .Top = 390

            .Width = 125

            .Height = 105

        End With

        ConnectSlicerToAllPivots _
            scWarehouse, _
            wsPivot

    End If

End Sub


'=========================================================
' CONNECT SLICER TO ALL PIVOTS
'=========================================================

Private Sub ConnectSlicerToAllPivots( _
    ByVal sc As SlicerCache, _
    ByVal wsPivot As Worksheet)

    Dim pt As PivotTable

    If sc Is Nothing Then Exit Sub

    On Error Resume Next

    For Each pt In wsPivot.PivotTables

        sc.PivotTables.AddPivotTable pt

    Next pt

    On Error GoTo 0

End Sub


'=========================================================
' GET PIVOT TABLE
'=========================================================

Private Function GetDashboardPivot( _
    ByVal wsPivot As Worksheet, _
    ParamArray PivotNames() As Variant) As PivotTable

    Dim i As Long

    Dim pt As PivotTable

    For i = LBound(PivotNames) To UBound(PivotNames)

        Set pt = Nothing

        On Error Resume Next

        Set pt = wsPivot.PivotTables( _
                    CStr(PivotNames(i)))

        On Error GoTo 0

        If Not pt Is Nothing Then

            Set GetDashboardPivot = pt

            Exit Function

        End If

    Next i

    Set GetDashboardPivot = Nothing

End Function


'=========================================================
' DELETE SLICER CACHE
'=========================================================

Private Sub DeleteSlicerCache( _
    ByVal wb As Workbook, _
    ByVal cacheName As String)

    Dim sc As SlicerCache

    On Error Resume Next

    Set sc = wb.SlicerCaches(cacheName)

    If Not sc Is Nothing Then

        sc.Delete

    End If

    On Error GoTo 0

End Sub


'=========================================================
' FIND COLUMN CONTAINING TEXT
'=========================================================

Private Function DashboardFindColumnContains( _
    ByVal ws As Worksheet, _
    ParamArray HeaderNames() As Variant) As Long

    Dim lastCol As Long

    Dim c As Long
    Dim i As Long

    Dim actualHeader As String
    Dim wantedHeader As String

    lastCol = _
        ws.Cells( _
            1, _
            ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol

        actualHeader = _
            DashboardNormalizeHeader( _
                ws.Cells(1, c).value)

        For i = LBound(HeaderNames) To UBound(HeaderNames)

            wantedHeader = _
                DashboardNormalizeHeader( _
                    HeaderNames(i))

            If InStr( _
                1, _
                actualHeader, _
                wantedHeader, _
                vbTextCompare) > 0 Then

                DashboardFindColumnContains = c

                Exit Function

            End If

        Next i

    Next c

    DashboardFindColumnContains = 0

End Function


'=========================================================
' FIND VALUE COLUMN
'=========================================================

Private Function DashboardFindValueColumn( _
    ByVal ws As Worksheet, _
    ByVal searchText As String) As Long

    Dim lastCol As Long

    Dim c As Long

    Dim h As String
    Dim target As String

    target = _
        DashboardNormalizeHeader(searchText)

    lastCol = _
        ws.Cells( _
            1, _
            ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol

        h = _
            DashboardNormalizeHeader( _
                ws.Cells(1, c).value)

        If h = target Then

            DashboardFindValueColumn = c

            Exit Function

        End If

        If InStr( _
            1, _
            h, _
            target, _
            vbTextCompare) > 0 Then

            DashboardFindValueColumn = c

            Exit Function

        End If

    Next c

    DashboardFindValueColumn = 0

End Function


'=========================================================
' FIND UNIT PRICE COLUMN
'=========================================================

Private Function DashboardFindUnitPriceColumn( _
    ByVal ws As Worksheet) As Long

    Dim lastCol As Long

    Dim c As Long

    Dim h As String

    lastCol = _
        ws.Cells( _
            1, _
            ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol

        h = _
            DashboardNormalizeHeader( _
                ws.Cells(1, c).value)

        If InStr( _
            1, _
            h, _
            "unit", _
            vbTextCompare) > 0 _
           And _
           InStr( _
            1, _
            h, _
            "price", _
            vbTextCompare) > 0 Then

            DashboardFindUnitPriceColumn = c

            Exit Function

        End If

    Next c

    DashboardFindUnitPriceColumn = 0

End Function


'=========================================================
' SUM COLUMN
'=========================================================

Private Function DashboardSumColumn( _
    ByVal ws As Worksheet, _
    ByVal colNumber As Long, _
    ByVal lastRow As Long) As Double

    Dim i As Long

    Dim v As Variant

    Dim num As Double

    DashboardSumColumn = 0

    For i = 2 To lastRow

        v = _
            ws.Cells( _
                i, _
                colNumber).value

        If DashboardTryNumber(v, num) Then

            DashboardSumColumn = _
                DashboardSumColumn + num

        End If

    Next i

End Function


'=========================================================
' NORMALIZE HEADER
'=========================================================

Private Function DashboardNormalizeHeader( _
    ByVal v As Variant) As String

    Dim s As String

    If IsError(v) Then

        DashboardNormalizeHeader = ""

        Exit Function

    End If

    s = CStr(v)

    s = Replace( _
            s, _
            Chr(160), _
            " ")

    s = Replace( _
            s, _
            vbCr, _
            " ")

    s = Replace( _
            s, _
            vbLf, _
            " ")

    s = LCase(s)

    s = Trim(s)

    Do While InStr(s, "  ") > 0

        s = Replace( _
                s, _
                "  ", _
                " ")

    Loop

    DashboardNormalizeHeader = s

End Function


'=========================================================
' CONVERT VALUE TO NUMBER
'=========================================================

Private Function DashboardTryNumber( _
    ByVal v As Variant, _
    ByRef result As Double) As Boolean

    Dim s As String

    DashboardTryNumber = False

    result = 0

    If IsError(v) Then Exit Function

    '=====================================================
    ' NORMAL NUMERIC CELL
    '=====================================================

    If IsNumeric(v) Then

        result = CDbl(v)

        DashboardTryNumber = True

        Exit Function

    End If

    '=====================================================
    ' TEXT NUMBER
    '=====================================================

    s = Trim(CStr(v))

    If s = "" Then Exit Function

    s = Replace(s, "?", "")

    s = Replace(s, "?", "")

    s = Replace(s, "$", "")

    s = Replace(s, "€", "")

    s = Replace(s, "£", "")

    s = Replace(s, ",", "")

    s = Replace(s, " ", "")

    If IsNumeric(s) Then

        result = CDbl(s)

        DashboardTryNumber = True

    End If

End Function


'=========================================================
' MONEY FORMAT
'=========================================================
' IMPORTANT:
' No ? symbol is used here.
' This prevents the ? character from appearing.
'=========================================================

Private Function DashboardFormatMoney( _
    ByVal value As Double) As String

    DashboardFormatMoney = _
        Format(value, "#,##0.00")

End Function


'=========================================================
' COMPACT MONEY FORMAT
'=========================================================
' IMPORTANT:
' No ? symbol is used here.
'=========================================================

Private Function DashboardFormatCompactMoney( _
    ByVal value As Double) As String

    If Abs(value) >= 10000000 Then

        DashboardFormatCompactMoney = _
            Format( _
                value / 10000000, _
                "0.0") & " Cr"

    ElseIf Abs(value) >= 100000 Then

        DashboardFormatCompactMoney = _
            Format( _
                value / 100000, _
                "0.0") & " L"

    ElseIf Abs(value) >= 1000 Then

        DashboardFormatCompactMoney = _
            Format( _
                value / 1000, _
                "0.0") & " K"

    Else

        DashboardFormatCompactMoney = _
            Format( _
                value, _
                "#,##0")

    End If

End Function


'=========================================================
' COMPACT NUMBER
'=========================================================

Private Function DashboardFormatCompactNumber( _
    ByVal value As Double) As String

    If Abs(value) >= 1000000 Then

        DashboardFormatCompactNumber = _
            Format( _
                value / 1000000, _
                "0.0") & " M"

    ElseIf Abs(value) >= 1000 Then

        DashboardFormatCompactNumber = _
            Format( _
                value / 1000, _
                "0.0") & " K"

    Else

        DashboardFormatCompactNumber = _
            Format( _
                value, _
                "#,##0")

    End If

End Function


'=========================================================
' CHECK SHEET EXISTS
'=========================================================

Private Function DashboardSheetExists( _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    DashboardSheetExists = False

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If Not ws Is Nothing Then

        DashboardSheetExists = True

    End If

End Function


'=========================================================
' DELETE DASHBOARD SHEET
'=========================================================

Private Sub DeleteDashboardSheet( _
    ByVal sheetName As String)

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If Not ws Is Nothing Then

        Application.DisplayAlerts = False

        ws.Delete

        Application.DisplayAlerts = True

    End If

End Sub

