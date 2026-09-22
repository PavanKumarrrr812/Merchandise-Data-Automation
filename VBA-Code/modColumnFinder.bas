Attribute VB_Name = "modColumnFinder"
Option Explicit

Public Function ProductColumn(ByVal ws As Worksheet) As Long

    ProductColumn = FindColumn(ws, _
        "Product", _
        "Product Name", _
        "Item", _
        "Item Name")

End Function


Public Function CategoryColumn(ByVal ws As Worksheet) As Long

    CategoryColumn = FindColumn(ws, _
        "Category", _
        "Product Category")

End Function


Public Function SupplierColumn(ByVal ws As Worksheet) As Long

    SupplierColumn = FindColumn(ws, _
        "Supplier", _
        "Vendor")

End Function


Public Function WarehouseColumn(ByVal ws As Worksheet) As Long

    WarehouseColumn = FindColumn(ws, _
        "Warehouse", _
        "Warehouse Name", _
        "Location", _
        "Hub")

End Function


Public Function QuantityColumn(ByVal ws As Worksheet) As Long

    QuantityColumn = FindColumn(ws, _
        "Quantity", _
        "Qty")

End Function


Public Function ReorderLevelColumn(ByVal ws As Worksheet) As Long

    ReorderLevelColumn = FindColumn(ws, _
        "Reorder Level", _
        "ReorderLevel")

End Function


Public Function StatusColumn(ByVal ws As Worksheet) As Long

    StatusColumn = FindColumn(ws, _
        "Stock Status", _
        "Status")

End Function

