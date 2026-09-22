Attribute VB_Name = "modHeaderChecker"
Option Explicit

Public Function IsValueOrPriceHeader( _
    ByVal header As String) As Boolean

    Dim h As String

    h = LCase$(Trim$(header))

    IsValueOrPriceHeader = False

    If InStr(1, h, "price", vbTextCompare) > 0 Then
        IsValueOrPriceHeader = True
        Exit Function
    End If

    If InStr(1, h, "value", vbTextCompare) > 0 Then
        IsValueOrPriceHeader = True
        Exit Function
    End If

End Function

