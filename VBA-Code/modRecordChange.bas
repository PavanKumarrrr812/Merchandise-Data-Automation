Attribute VB_Name = "modRecordChange"
Option Explicit

Public Sub RecordChange( _
    ByVal cell As Range, _
    ByVal oldValue As Variant, _
    ByVal newValue As Variant)

    Dim wb As Workbook
    Dim wsLog As Worksheet
    Dim wsHighlight As Worksheet
    Dim r As Long
    Dim columnName As String

    Set wb = ThisWorkbook
    Set wsLog = wb.Worksheets("Change_Log")

    columnName = cell.Worksheet.Cells(1, cell.Column).value

    r = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).Row + 1

    wsLog.Cells(r, 1).value = cell.Worksheet.Name
    wsLog.Cells(r, 2).value = cell.Address(False, False)
    wsLog.Cells(r, 3).value = columnName
    wsLog.Cells(r, 4).value = oldValue
    wsLog.Cells(r, 5).value = newValue

    cell.Interior.Color = RGB(255, 255, 0)

    If cell.Worksheet.Name = "CleanedData" Then

        Set wsHighlight = wb.Worksheets("Changes_Highlighted")

        wsHighlight.Cells( _
            cell.Row, _
            cell.Column).Interior.Color = RGB(255, 255, 0)

    End If

End Sub

