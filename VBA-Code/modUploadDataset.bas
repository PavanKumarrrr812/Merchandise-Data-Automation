Attribute VB_Name = "modUploadDataset"
Option Explicit

'==========================================================
' UPLOAD DATASET - MAIN MENU
'==========================================================

Public Sub UploadDataset()

    Dim choice As String

    choice = InputBox( _
        "Select upload method:" & vbCrLf & vbCrLf & _
        "1 = File" & vbCrLf & _
        "2 = Folder" & vbCrLf & _
        "3 = URL", _
        "Upload Dataset")

    Select Case Trim(choice)

        Case "1"
            UploadFromFile

        Case "2"
            UploadFromFolder

        Case "3"
            UploadFromURL

        Case ""
            Exit Sub

        Case Else
            MsgBox "Invalid selection." & vbCrLf & _
                   "Please enter 1, 2, or 3.", _
                   vbExclamation, _
                   "Upload Dataset"

    End Select

End Sub


'==========================================================
' UPLOAD FROM FILE
'==========================================================

Public Sub UploadFromFile()

    Dim fd As Object
    Dim filePath As String

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsUploaded As Worksheet

    On Error GoTo ErrorHandler

    Set fd = Application.FileDialog(3)

    With fd

        .title = "Select Dataset"

        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls"
        .Filters.Add "CSV Files", "*.csv"

        .AllowMultiSelect = False

        If .Show <> -1 Then
            Exit Sub
        End If

        filePath = .SelectedItems(1)

    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbSource = Workbooks.Open(filePath)

    Set wsSource = wbSource.Worksheets(1)

    DeleteSheetIfExists "OriginalData"

    wsSource.Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

    Set wsUploaded = ThisWorkbook.Sheets( _
        ThisWorkbook.Sheets.count)

    wsUploaded.Name = "OriginalData"

    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Dataset uploaded successfully.", _
           vbInformation, _
           "Upload Complete"

    Exit Sub


ErrorHandler:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not wbSource Is Nothing Then

        On Error Resume Next
        wbSource.Close SaveChanges:=False
        On Error GoTo 0

    End If

    MsgBox "Unable to upload dataset." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Upload Error"

End Sub


'==========================================================
' UPLOAD FROM FOLDER
'==========================================================

Public Sub UploadFromFolder()

    Dim fd As Object
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsUploaded As Worksheet

    On Error GoTo ErrorHandler

    Set fd = Application.FileDialog(4)

    With fd

        .title = "Select Dataset Folder"

        .AllowMultiSelect = False

        If .Show <> -1 Then
            Exit Sub
        End If

        folderPath = .SelectedItems(1)

    End With

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If


    '------------------------------------------------------
    ' Find first Excel or CSV file
    '------------------------------------------------------

    fileName = Dir(folderPath & "*.xlsx")

    If fileName = "" Then
        fileName = Dir(folderPath & "*.xlsm")
    End If

    If fileName = "" Then
        fileName = Dir(folderPath & "*.xls")
    End If

    If fileName = "" Then
        fileName = Dir(folderPath & "*.csv")
    End If


    If fileName = "" Then

        MsgBox "No Excel or CSV dataset was found in the selected folder.", _
               vbExclamation, _
               "No Dataset Found"

        Exit Sub

    End If


    fullPath = folderPath & fileName

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbSource = Workbooks.Open(fullPath)

    Set wsSource = wbSource.Worksheets(1)

    DeleteSheetIfExists "OriginalData"

    wsSource.Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

    Set wsUploaded = ThisWorkbook.Sheets( _
        ThisWorkbook.Sheets.count)

    wsUploaded.Name = "OriginalData"

    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Dataset uploaded successfully." & vbCrLf & vbCrLf & _
           "File: " & fileName, _
           vbInformation, _
           "Upload Complete"

    Exit Sub


ErrorHandler:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not wbSource Is Nothing Then

        On Error Resume Next
        wbSource.Close SaveChanges:=False
        On Error GoTo 0

    End If

    MsgBox "Unable to upload dataset from folder." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Upload Error"

End Sub


'==========================================================
' UPLOAD FROM URL
'==========================================================

Public Sub UploadFromURL()

    Dim url As String
    Dim tempPath As String

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsUploaded As Worksheet

    On Error GoTo ErrorHandler

    url = InputBox( _
        "Enter the direct download URL of the dataset:" & vbCrLf & vbCrLf & _
        "Example:" & vbCrLf & _
        "https://example.com/data.xlsx", _
        "Upload Dataset From URL")

    If Trim(url) = "" Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    tempPath = Environ$("TEMP") & "\UploadedDataset.xlsx"

    DownloadFileFromURL url, tempPath

    Set wbSource = Workbooks.Open(tempPath)

    Set wsSource = wbSource.Worksheets(1)

    DeleteSheetIfExists "OriginalData"

    wsSource.Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

    Set wsUploaded = ThisWorkbook.Sheets( _
        ThisWorkbook.Sheets.count)

    wsUploaded.Name = "OriginalData"

    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    MsgBox "Dataset uploaded successfully from URL.", _
           vbInformation, _
           "Upload Complete"

    Exit Sub


ErrorHandler:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not wbSource Is Nothing Then

        On Error Resume Next
        wbSource.Close SaveChanges:=False
        On Error GoTo 0

    End If

    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    MsgBox "Unable to upload dataset from URL." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Upload Error"

End Sub


'==========================================================
' DOWNLOAD FILE FROM URL
'==========================================================

Private Sub DownloadFileFromURL( _
    ByVal url As String, _
    ByVal savePath As String)

    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False

    http.send

    If http.status <> 200 Then

        Err.Raise _
            vbObjectError + 1000, _
            "DownloadFileFromURL", _
            "Unable to download the file." & vbCrLf & _
            "HTTP Status: " & http.status

    End If


    Set stream = CreateObject("ADODB.Stream")

    With stream

        .Type = 1
        .Open

        .Write http.responseBody

        .SaveToFile savePath, 2

        .Close

    End With

    Set stream = Nothing
    Set http = Nothing

End Sub


'==========================================================
' DELETE SHEET IF IT EXISTS
'==========================================================

Public Sub DeleteSheetIfExists(ByVal sheetName As String)

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

