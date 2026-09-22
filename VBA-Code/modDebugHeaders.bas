Attribute VB_Name = "modDebugHeaders"
Sub CheckTableHeaders()

    Dim tbl As ListObject
    Dim col As ListColumn

    Set tbl = ThisWorkbook.Worksheets("CleanedData").ListObjects("Table172")

    For Each col In tbl.ListColumns
        Debug.Print col.index & " | [" & col.Name & "]"
    Next col

End Sub
