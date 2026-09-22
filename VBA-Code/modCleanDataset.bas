Attribute VB_Name = "modCleanDataset"
Option Explicit

Public Sub CleanDataset()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim col As Long
    Dim lastCol As Long

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("CleanedData")

    '1. Remove completely empty columns
    RemoveEmptyColumns ws

    '2. Clean Product
    CleanProductColumn

    '3. Clean Category
    CleanTextColumn CategoryColumn(ws)

    '4. Clean Supplier
    CleanTextColumn SupplierColumn(ws)

    '5. Clean Warehouse
    CleanTextColumn WarehouseColumn(ws)

    '6. Clean Quantity
    CleanQuantity

    '7. Clean Reorder Level
    CleanReorderLevel

    '8. Clean Stock Status
    CleanStatusColumn

    '9. Clean existing Price / Value columns
    lastCol = ws.Cells(1, _
        ws.Columns.count).End(xlToLeft).Column

    For col = 1 To lastCol

        If IsValueOrPriceHeader( _
            CStr(ws.Cells(1, col).value)) Then

            CleanPriceValueColumn col

        End If

    Next col

    '10. Create derived columns
    CreateInventoryValue
    CreateReorderQuantity
    CreateAboveReorderQuantity
    CreateAboveReorderInventoryValue

    '11. Clean Inventory Value
    col = FindColumn(ws, _
        "Inventory Value", _
        "InventoryValue")

    If col > 0 Then
        CleanPriceValueColumn col
    End If

    '12. Clean Above Reorder Inventory Value
    col = FindColumn(ws, _
        "Above Reorder Inventory Value", _
        "AboveReorderInventoryValue")

    If col > 0 Then
        CleanPriceValueColumn col
    End If

    '13. Clean Date
    CleanDateColumn

    '14. Refresh highlighted copy
    RefreshHighlightCopy

End Sub

