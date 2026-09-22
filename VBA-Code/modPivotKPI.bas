Attribute VB_Name = "modPivotKPI"
Option Explicit

'=========================================================
' KPI SECTION
'=========================================================

Public Sub CreateKPISection( _
    ByVal wsPivot As Worksheet, _
    ByVal sourceRange As Range)

    Dim wsData As Worksheet
    Dim lastRow As Long

    Dim colQuantity As Long
    Dim colInventoryValue As Long
    Dim colUnitPrice As Long
    Dim colReorderQty As Long
    Dim colAboveReorderQty As Long
    Dim colProduct As Long

    Dim totalQuantity As Double
    Dim totalInventoryValue As Double
    Dim avgUnitPrice As Double
    Dim totalReorderQty As Double
    Dim aboveReorderQty As Double
    Dim productCount As Long

    Dim priceTotal As Double
    Dim priceCount As Long

    Dim i As Long
    Dim v As Variant
    Dim numericPrice As Double

    Set wsData = ThisWorkbook.Worksheets("CleanedData")

    '=====================================================
    ' FIND LAST ROW
    '=====================================================

    lastRow = wsData.Cells( _
                    wsData.Rows.count, 1).End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    '=====================================================
    ' FIND REQUIRED COLUMNS
    '=====================================================

    colQuantity = FindKPIColumn( _
                    wsData, _
                    "Quantity", _
                    "Qty")

    colInventoryValue = FindKPIColumn( _
                    wsData, _
                    "Inventory Value")

    'Find ONLY Unit Price
    colUnitPrice = FindUnitPriceColumn(wsData)

    colReorderQty = FindKPIColumn( _
                    wsData, _
                    "Reorder QTY", _
                    "Reorder Quantity")

    colAboveReorderQty = FindKPIColumn( _
                    wsData, _
                    "Above Reorder QTY", _
                    "Above Reorder Quantity")

    colProduct = FindKPIColumn( _
                    wsData, _
                    "Product", _
                    "Product Name", _
                    "Item", _
                    "Item Name")

    '=====================================================
    ' UNIT PRICE CHECK
    '=====================================================

    If colUnitPrice = 0 Then

        MsgBox _
            "Unit Price column was not found." & vbCrLf & vbCrLf & _
            "VBA checked row 1 of CleanedData for a header " & _
            "containing both 'Unit' and 'Price'.", _
            vbCritical, _
            "Unit Price Error"

        Exit Sub

    End If

    '=====================================================
    ' TOTAL QUANTITY
    '=====================================================

    If colQuantity > 0 Then

        totalQuantity = Application.Sum( _
            wsData.Range( _
                wsData.Cells(2, colQuantity), _
                wsData.Cells(lastRow, colQuantity)))

    End If

    '=====================================================
    ' TOTAL INVENTORY VALUE
    '=====================================================

    If colInventoryValue > 0 Then

        totalInventoryValue = Application.Sum( _
            wsData.Range( _
                wsData.Cells(2, colInventoryValue), _
                wsData.Cells(lastRow, colInventoryValue)))

    End If

    '=====================================================
    ' AVERAGE UNIT PRICE
    '=====================================================

    priceTotal = 0
    priceCount = 0

    For i = 2 To lastRow

        v = wsData.Cells(i, colUnitPrice).value

        If TryGetNumber(v, numericPrice) Then

            If numericPrice >= 0 Then

                priceTotal = priceTotal + numericPrice
                priceCount = priceCount + 1

            End If

        End If

    Next i

    If priceCount > 0 Then

        avgUnitPrice = priceTotal / priceCount

    Else

        avgUnitPrice = 0

    End If

    '=====================================================
    ' TOTAL REORDER QUANTITY
    '=====================================================

    If colReorderQty > 0 Then

        totalReorderQty = Application.Sum( _
            wsData.Range( _
                wsData.Cells(2, colReorderQty), _
                wsData.Cells(lastRow, colReorderQty)))

    End If

    '=====================================================
    ' ABOVE REORDER QUANTITY
    '=====================================================

    If colAboveReorderQty > 0 Then

        aboveReorderQty = Application.Sum( _
            wsData.Range( _
                wsData.Cells(2, colAboveReorderQty), _
                wsData.Cells(lastRow, colAboveReorderQty)))

    End If

    '=====================================================
    ' PRODUCT COUNT
    '=====================================================

    productCount = 0

    If colProduct > 0 Then

        For i = 2 To lastRow

            v = wsData.Cells(i, colProduct).value

            If Not IsError(v) Then

                If Trim(CStr(v)) <> "" Then

                    If LCase(Trim(CStr(v))) <> "unknown" Then

                        productCount = productCount + 1

                    End If

                End If

            End If

        Next i

    End If

    '=====================================================
    ' CLEAR OLD KPI AREA
    '=====================================================

    With wsPivot.Range("B4:G5")

        .UnMerge
        .Clear
        .Borders.LineStyle = xlNone

    End With

    '=====================================================
    ' KPI HEADINGS
    '=====================================================

    wsPivot.Range("B4").value = "TOTAL QUANTITY"
    wsPivot.Range("C4").value = "INVENTORY VALUE"
    wsPivot.Range("D4").value = "AVG UNIT PRICE"
    wsPivot.Range("E4").value = "TOTAL REORDER QTY"
    wsPivot.Range("F4").value = "ABOVE REORDER QTY"
    wsPivot.Range("G4").value = "PRODUCT COUNT"

    '=====================================================
    ' KPI VALUES
    '=====================================================

    wsPivot.Range("B5").value = totalQuantity
    wsPivot.Range("C5").value = totalInventoryValue
    wsPivot.Range("D5").value = avgUnitPrice
    wsPivot.Range("E5").value = totalReorderQty
    wsPivot.Range("F5").value = aboveReorderQty
    wsPivot.Range("G5").value = productCount

    '=====================================================
    ' NUMBER FORMATTING
    '=====================================================

    wsPivot.Range("B5").NumberFormat = "#,##0"
    wsPivot.Range("C5").NumberFormat = "#,##,##0"
    wsPivot.Range("D5").NumberFormat = "#,##0.00"
    wsPivot.Range("E5").NumberFormat = "#,##0"
    wsPivot.Range("F5").NumberFormat = "#,##0"
    wsPivot.Range("G5").NumberFormat = "#,##0"

    '=====================================================
    ' KPI HEADER FORMATTING
    '=====================================================

    With wsPivot.Range("B4:G4")

        .Interior.Color = RGB(25, 35, 55)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 185, 190)

    End With

    '=====================================================
    ' KPI VALUE FORMATTING
    '=====================================================

    With wsPivot.Range("B5:G5")

        .Interior.Color = RGB(245, 247, 250)

        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(25, 35, 55)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 185, 190)

    End With

    '=====================================================
    ' KPI COLUMN WIDTHS
    '=====================================================

    wsPivot.Columns("B:G").ColumnWidth = 18

    '=====================================================
    ' ROW HEIGHT
    '=====================================================

    wsPivot.Rows("4").RowHeight = 32
    wsPivot.Rows("5").RowHeight = 28

    '=====================================================
    ' ALIGNMENT
    '=====================================================

    wsPivot.Range("B4:G5").HorizontalAlignment = xlCenter
    wsPivot.Range("B4:G5").VerticalAlignment = xlCenter

End Sub


'=========================================================
' FIND UNIT PRICE COLUMN
'=========================================================

Private Function FindUnitPriceColumn( _
    ByVal ws As Worksheet) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim headerText As String

    FindUnitPriceColumn = 0

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol

        headerText = CStr(ws.Cells(1, c).value)

        'Remove invisible characters
        headerText = Replace(headerText, Chr(160), " ")
        headerText = Replace(headerText, vbCr, " ")
        headerText = Replace(headerText, vbLf, " ")

        'Convert to lowercase
        headerText = LCase(Trim(headerText))

        'Remove multiple spaces
        Do While InStr(headerText, "  ") > 0
            headerText = Replace(headerText, "  ", " ")
        Loop

        '-------------------------------------------------
        ' ONLY LOOK FOR A HEADER CONTAINING:
        '
        ' UNIT
        ' PRICE
        '
        ' This will detect:
        '
        ' Unit Price
        ' Unit Price (?)
        ' Unit Price (?)
        ' UNIT PRICE
        '
        ' But will NOT detect:
        '
        ' Purchase Value
        ' Sales Value
        ' Inventory Value
        '-------------------------------------------------

        If InStr(1, headerText, "unit", vbTextCompare) > 0 _
           And InStr(1, headerText, "price", vbTextCompare) > 0 Then

            FindUnitPriceColumn = c
            Exit Function

        End If

    Next c

End Function


'=========================================================
' CONVERT VALUE TO NUMBER
'=========================================================

Private Function TryGetNumber( _
    ByVal InputValue As Variant, _
    ByRef result As Double) As Boolean

    Dim s As String

    TryGetNumber = False
    result = 0

    If IsError(InputValue) Then Exit Function

    '---------------------------------------------
    ' Already a numeric Excel value
    '---------------------------------------------

    If IsNumeric(InputValue) Then

        result = CDbl(InputValue)
        TryGetNumber = True

        Exit Function

    End If

    '---------------------------------------------
    ' Treat as text
    '---------------------------------------------

    s = Trim(CStr(InputValue))

    If s = "" Then Exit Function

    'Remove currency symbol
    s = Replace(s, "?", "")

    'Handle encoding replacement character
    s = Replace(s, "?", "")

    'Remove commas
    s = Replace(s, ",", "")

    'Remove spaces
    s = Replace(s, " ", "")

    'Try conversion again
    If IsNumeric(s) Then

        result = CDbl(s)
        TryGetNumber = True

    End If

End Function


'=========================================================
' FIND NORMAL KPI COLUMN
'=========================================================

Private Function FindKPIColumn( _
    ByVal ws As Worksheet, _
    ParamArray HeaderNames() As Variant) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim i As Long

    Dim actualHeader As String
    Dim wantedHeader As String

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol

        actualHeader = NormalizeKPIHeader( _
                        ws.Cells(1, c).value)

        If actualHeader <> "" Then

            For i = LBound(HeaderNames) To UBound(HeaderNames)

                wantedHeader = NormalizeKPIHeader( _
                                HeaderNames(i))

                If actualHeader = wantedHeader Then

                    FindKPIColumn = c
                    Exit Function

                End If

            Next i

        End If

    Next c

    FindKPIColumn = 0

End Function


'=========================================================
' NORMALIZE HEADER
'=========================================================

Private Function NormalizeKPIHeader( _
    ByVal HeaderValue As Variant) As String

    Dim s As String

    If IsError(HeaderValue) Then

        NormalizeKPIHeader = ""
        Exit Function

    End If

    s = CStr(HeaderValue)

    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    s = Trim(s)

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeKPIHeader = LCase(s)

End Function

