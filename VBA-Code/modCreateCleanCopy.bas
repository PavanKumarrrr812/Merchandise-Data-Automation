Attribute VB_Name = "modCreateCleanCopy"
Option Explicit

Public Sub CreateCleanCopy()

    Dim wb As Workbook
    Dim wsOriginal As Worksheet
    Dim wsCleaned As Worksheet
    Dim wsHighlighted As Worksheet

    Set wb = ThisWorkbook

    If Not CleanCopySheetExists("OriginalData") Then

        MsgBox "Please upload a dataset first.", _
               vbExclamation, _
               "Dataset Required"

        Exit Sub

    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    DeleteCleanCopySheet "CleanedData"
    DeleteCleanCopySheet "Changes_Highlighted"
    DeleteCleanCopySheet "Change_Log"

    Set wsOriginal = wb.Worksheets("OriginalData")

    wsOriginal.Copy After:=wsOriginal

    Set wsCleaned = wb.Worksheets(wsOriginal.index + 1)

    wsCleaned.Name = "CleanedData"

    wsCleaned.Copy After:=wsCleaned

    Set wsHighlighted = wb.Worksheets(wsCleaned.index + 1)

    wsHighlighted.Name = "Changes_Highlighted"

    SetChangeLog

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


Public Sub SetChangeLog()

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ThisWorkbook

    DeleteCleanCopySheet "Change_Log"

    Set ws = wb.Worksheets.Add( _
        After:=wb.Worksheets("Changes_Highlighted"))

    ws.Name = "Change_Log"

    ws.Range("A1:E1").value = Array( _
        "Sheet", _
        "Cell", _
        "Column", _
        "Original Value", _
        "New Value")

    With ws.Range("A1:E1")

        .Font.Bold = True
        .Interior.Color = RGB(25, 35, 55)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter

    End With

    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 25
    ws.Columns("E").ColumnWidth = 25

End Sub


Private Function CleanCopySheetExists( _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    CleanCopySheetExists = False

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If Not ws Is Nothing Then
        CleanCopySheetExists = True
    End If

End Function


Private Sub DeleteCleanCopySheet( _
    ByVal sheetName As String)

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If Not ws Is Nothing Then

        Application.DisplayAlerts = False

        ws.Delete

        Application.DisplayAlerts = True

    End If

End Sub

