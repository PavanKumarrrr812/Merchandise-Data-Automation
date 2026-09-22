Attribute VB_Name = "modMonthFromFirstThreeLetters"
Option Explicit

Public Function MonthFromFirstThreeLetters( _
    ByVal monthText As String) As Long

    Select Case LCase$(Left$(Trim$(monthText), 3))

        Case "jan": MonthFromFirstThreeLetters = 1
        Case "feb": MonthFromFirstThreeLetters = 2
        Case "mar": MonthFromFirstThreeLetters = 3
        Case "apr": MonthFromFirstThreeLetters = 4
        Case "may": MonthFromFirstThreeLetters = 5
        Case "jun": MonthFromFirstThreeLetters = 6
        Case "jul": MonthFromFirstThreeLetters = 7
        Case "aug": MonthFromFirstThreeLetters = 8
        Case "sep": MonthFromFirstThreeLetters = 9
        Case "oct": MonthFromFirstThreeLetters = 10
        Case "nov": MonthFromFirstThreeLetters = 11
        Case "dec": MonthFromFirstThreeLetters = 12

        Case Else
            MonthFromFirstThreeLetters = 0

    End Select

End Function

