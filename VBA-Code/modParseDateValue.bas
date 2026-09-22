Attribute VB_Name = "modParseDateValue"
Option Explicit

Public Function ParseDateValue( _
    ByVal value As Variant) As Variant

    Dim s As String
    Dim parts() As String
    Dim d As Long
    Dim m As Long
    Dim y As Long
    Dim testDate As Date

    If IsEmpty(value) Then
        ParseDateValue = Empty
        Exit Function
    End If

    s = Trim(CStr(value))

    If s = "" Then
        ParseDateValue = Empty
        Exit Function
    End If

    'Already a real Excel date
    If IsDate(value) And IsNumeric(value) Then
        ParseDateValue = CDate(value)
        Exit Function
    End If

    'Remove ordinal endings
    s = LCase$(s)
    s = Replace(s, "st", "")
    s = Replace(s, "nd", "")
    s = Replace(s, "rd", "")
    s = Replace(s, "th", "")

    'Replace common separators
    s = Replace(s, "/", "-")
    s = Replace(s, ".", "-")

    parts = Split(s, "-")

    If UBound(parts) = 2 Then

        'Day
        If IsNumeric(parts(0)) Then
            d = CLng(parts(0))
        Else
            ParseDateValue = Empty
            Exit Function
        End If

        'Month - first 3 letters
        m = MonthFromFirstThreeLetters(parts(1))

        If m = 0 Then

            If IsNumeric(parts(1)) Then
                m = CLng(parts(1))
            Else
                ParseDateValue = Empty
                Exit Function
            End If

        End If

        'Year
        If IsNumeric(parts(2)) Then

            y = CLng(parts(2))

            If y < 100 Then
                y = y + 2000
            End If

        Else

            ParseDateValue = Empty
            Exit Function

        End If

        'Validate actual calendar date
        On Error GoTo InvalidDate

        testDate = DateSerial(y, m, d)

        If Day(testDate) <> d _
           Or Month(testDate) <> m _
           Or Year(testDate) <> y Then

            GoTo InvalidDate

        End If

        ParseDateValue = testDate
        Exit Function

    End If

InvalidDate:

    ParseDateValue = Empty

End Function

