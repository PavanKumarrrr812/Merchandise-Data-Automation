Attribute VB_Name = "modUploadURL"
Option Explicit

Public Sub UploadFromURL()

    Dim url As String
    Dim tempPath As String

    url = InputBox( _
        "Enter direct dataset URL:", _
        "Dataset URL")

    If Trim(url) = "" Then Exit Sub

    tempPath = Environ$("TEMP") & "\DownloadedDataset.xlsx"

    If DownloadFile(url, tempPath) Then

        ImportExternalFile tempPath

        On Error Resume Next
        Kill tempPath
        On Error GoTo 0

    Else

        MsgBox "Unable to download dataset.", _
               vbCritical, "Download Error"

    End If

End Sub


Public Function DownloadFile( _
    ByVal url As String, _
    ByVal savePath As String) As Boolean

    Dim http As Object
    Dim stream As Object

    On Error GoTo Failed

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False
    http.send

    If http.status <> 200 Then GoTo Failed

    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile savePath, 2
    stream.Close

    Set stream = Nothing
    Set http = Nothing

    DownloadFile = True

    Exit Function

Failed:

    On Error Resume Next

    If Not stream Is Nothing Then
        stream.Close
    End If

    Set stream = Nothing
    Set http = Nothing

    DownloadFile = False

End Function

