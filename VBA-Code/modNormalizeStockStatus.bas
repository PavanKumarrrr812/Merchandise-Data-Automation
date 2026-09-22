Attribute VB_Name = "modNormalizeStockStatus"
Option Explicit

Public Function NormalizeStockStatus( _
    ByVal value As Variant) As String

    Dim v As String

    NormalizeStockStatus = "Unknown"

    v = LCase$(Trim$(CStr(value)))

    Select Case v

        Case "true"
            NormalizeStockStatus = "TRUE"

        Case "false"
            NormalizeStockStatus = "FALSE"

        Case "0"
            NormalizeStockStatus = "0"

        Case "1"
            NormalizeStockStatus = "1"

        Case "in stock"
            NormalizeStockStatus = "In Stock"

        Case "out of stock"
            NormalizeStockStatus = "Out Of Stock"

        Case "low stock"
            NormalizeStockStatus = "Low Stock"

        Case "normal stock"
            NormalizeStockStatus = "Normal Stock"

        Case "high stock"
            NormalizeStockStatus = "High Stock"

        Case ""
            NormalizeStockStatus = "Unknown"

        Case Else
            NormalizeStockStatus = "Unknown"

    End Select

End Function

