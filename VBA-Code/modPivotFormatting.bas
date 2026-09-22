Attribute VB_Name = "modPivotFormatting"
Option Explicit

Public Sub FormatRealPivot(ByVal pt As PivotTable)

    Dim ws As Worksheet

    Set ws = pt.Parent

    On Error Resume Next

    'Turn off unnecessary buttons
    pt.ShowDrillIndicators = False
    pt.DisplayFieldCaptions = True

    'General formatting
    With pt.TableRange1

        .Font.Name = "Calibri"
        .Font.Size = 11

        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin

        .VerticalAlignment = xlCenter

    End With

    'Header
    With pt.TableRange1.Rows(1)

        .Font.Bold = True
        .Interior.Color = RGB(25, 35, 55)
        .Font.Color = RGB(255, 255, 255)

    End With

    'Numbers
    If pt.DataFields.count > 0 Then

        pt.DataFields(1).NumberFormat = "#,##0.00"

    End If

    pt.TableRange1.Columns.AutoFit

    On Error GoTo 0

End Sub
