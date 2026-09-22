Attribute VB_Name = "modUploadFolder"
Option Explicit

Public Sub UploadFromFolder()

    Dim fd As FileDialog
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As String

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd

        .title = "Select Dataset Folder"

        If .Show <> -1 Then Exit Sub

        folderPath = .SelectedItems(1)

    End With

    fileName = Dir(folderPath & "\*.xlsx")

    If fileName = "" Then
        fileName = Dir(folderPath & "\*.xlsm")
    End If

    If fileName = "" Then
        fileName = Dir(folderPath & "\*.xls")
    End If

    If fileName = "" Then
        fileName = Dir(folderPath & "\*.csv")
    End If

    If fileName = "" Then

        MsgBox "No Excel or CSV dataset found.", _
               vbExclamation, "Dataset Not Found"

        Exit Sub

    End If

    filePath = folderPath & "\" & fileName

    ImportExternalFile filePath

End Sub


Public Sub ImportExternalFile(ByVal filePath As String)

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsUploaded As Worksheet

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    DeleteSheetIfExists "OriginalData"

    Set wbSource = Workbooks.Open(filePath)

    Set wsSource = wbSource.Worksheets(1)

    wsSource.Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

    Set wsUploaded = ThisWorkbook.Sheets( _
        ThisWorkbook.Sheets.count)

    wsUploaded.Name = "OriginalData"

    wbSource.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Dataset imported successfully.", _
           vbInformation, "Import Complete"

    Exit Sub

ErrorHandler:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not wbSource Is Nothing Then
        On Error Resume Next
        wbSource.Close SaveChanges:=False
        On Error GoTo 0
    End If

    MsgBox "Unable to import dataset." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "Import Error"

End Sub

