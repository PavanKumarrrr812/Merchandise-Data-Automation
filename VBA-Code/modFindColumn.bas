Attribute VB_Name = "modFindColumn"
Option Explicit

Public Function FindColumn( _
    ByVal ws As Worksheet, _
    ParamArray names() As Variant) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim i As Long
    Dim HeaderName As String
    Dim SearchName As String

    FindColumn = 0

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol

        HeaderName = Trim$(CStr(ws.Cells(1, c).value))

        For i = LBound(names) To UBound(names)

            SearchName = Trim$(CStr(names(i)))

            If StrComp(HeaderName, SearchName, vbTextCompare) = 0 Then
                FindColumn = c
                Exit Function
            End If

        Next i

    Next c

End Function

