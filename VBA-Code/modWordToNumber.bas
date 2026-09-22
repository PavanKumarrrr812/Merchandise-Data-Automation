Attribute VB_Name = "modWordToNumber"
Option Explicit

Public Function WordToNumber( _
    ByVal v As String) As Variant

    Select Case LCase$(Trim$(v))

        Case "zero"
            WordToNumber = 0

        Case "one"
            WordToNumber = 1

        Case "two"
            WordToNumber = 2

        Case "three"
            WordToNumber = 3

        Case "four"
            WordToNumber = 4

        Case "five"
            WordToNumber = 5

        Case "six"
            WordToNumber = 6

        Case "seven"
            WordToNumber = 7

        Case "eight"
            WordToNumber = 8

        Case "nine"
            WordToNumber = 9

        Case "ten"
            WordToNumber = 10

        Case Else
            WordToNumber = CVErr(xlErrNA)

    End Select

End Function

