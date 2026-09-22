Attribute VB_Name = "modRefreshHighlightCopy"
Option Explicit

Public Sub RefreshHighlightCopy()

    Dim wb As Workbook
    Dim wsClean As Worksheet
    Dim wsHighlight As Worksheet

    Set wb = ThisWorkbook

    Set wsClean = wb.Worksheets("CleanedData")
    Set wsHighlight = _
        wb.Worksheets("Changes_Highlighted")

    wsHighlight.Cells.Clear

    wsClean.UsedRange.Copy _
        Destination:=wsHighlight.Range("A1")

End Sub

