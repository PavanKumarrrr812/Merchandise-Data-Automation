Attribute VB_Name = "Dashboard"
Option Explicit

'========================================================
' BUILD DASHBOARD
'========================================================

Public Sub BuildDashboard()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape

    Set wb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '----------------------------------------------------
    ' Check Sheet1
    '----------------------------------------------------
    On Error Resume Next
    Set ws = wb.Worksheets("Sheet1")
    On Error GoTo 0

    If ws Is Nothing Then

        Application.DisplayAlerts = True
        Application.ScreenUpdating = True

        MsgBox "Sheet1 was not found.", _
               vbExclamation, _
               "Sheet1 Not Found"

        Exit Sub

    End If


    '----------------------------------------------------
    ' Delete existing Dashboard
    '----------------------------------------------------
    On Error Resume Next
    wb.Worksheets("Dashboard").Delete
    On Error GoTo 0


    '----------------------------------------------------
    ' Create Dashboard
    '----------------------------------------------------
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "Dashboard"


    '====================================================
    ' BACKGROUND
    '====================================================

    ws.Cells.Interior.Color = RGB(245, 247, 250)


    '====================================================
    ' COLUMN WIDTHS
    '====================================================

    ws.Columns("A").ColumnWidth = 3

    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 14

    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 14

    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 14

    ws.Columns("K").ColumnWidth = 3


    '====================================================
    ' ROW HEIGHTS
    '====================================================

    ws.Rows("1:30").RowHeight = 22


    '====================================================
    ' HEADER
    '====================================================

    With ws.Range("B1:J2")

        .Merge

        .Interior.Color = RGB(25, 35, 55)

        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 20

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .value = "MERCHANDISE DATA AUTOMATION"

    End With


    '====================================================
    ' SUBTITLE
    '====================================================

    With ws.Range("B3:J3")

        .Merge

        .Font.Size = 10
        .Font.Color = RGB(90, 90, 90)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .value = "Automate data cleaning, pivot analysis and dashboard creation"

    End With


    '====================================================
    ' DATASET CARD
    '====================================================

    With ws.Range("B6:J8")

        .Interior.Color = RGB(255, 255, 255)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(220, 225, 230)

    End With


    'Dataset title
    With ws.Range("B6:D6")

        .Merge

        .value = "DATASET"

        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(25, 35, 55)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    'Dataset description
    With ws.Range("B7:D8")

        .Merge

        .value = "Upload your Excel dataset to begin automation."

        .Font.Size = 10
        .Font.Color = RGB(100, 100, 100)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    '====================================================
    ' UPLOAD BUTTON
    '====================================================

    Set shp = ws.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                ws.Range("E6").Left, _
                ws.Range("E6").Top + 5, _
                200, _
                40)

    With shp

        .Name = "btnUploadDataset"

        .Fill.ForeColor.RGB = RGB(52, 152, 219)

        .Line.Visible = msoFalse

        .TextFrame2.TextRange.Text = "UPLOAD DATASET"

        .TextFrame2.TextRange.Font.Size = 10

        .TextFrame2.TextRange.Font.Bold = msoTrue

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        'CALL UPLOAD MACRO
        .OnAction = "UploadDataset"

    End With


    '====================================================
    ' STATUS
    '====================================================

    With ws.Range("H6:J8")

        .Merge

        .Interior.Color = RGB(235, 248, 240)

        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(39, 120, 70)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .value = "READY"

    End With


    '====================================================
    ' AUTOMATION OPTIONS TITLE
    '====================================================

    With ws.Range("B10:J10")

        .Merge

        .value = "AUTOMATION OPTIONS"

        .Font.Bold = True
        .Font.Size = 15
        .Font.Color = RGB(25, 35, 55)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    '====================================================
    ' CLEAN DATA CARD
    '====================================================

    With ws.Range("B12:D16")

        .Interior.Color = RGB(230, 242, 255)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 210, 240)

    End With


    AddAutomationCheckbox _
        ws, _
        ws.Range("B13").Left + 8, _
        ws.Range("B13").Top + 5, _
        "CLEAN"


    With ws.Range("C13:D14")

        .Merge

        .value = "CLEAN DATA"

        .Font.Bold = True
        .Font.Size = 15
        .Font.Color = RGB(30, 100, 170)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    With ws.Range("B15:D15")

        .Merge

        .value = "Clean and standardize dataset"

        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    '====================================================
    ' PIVOT TABLES CARD
    '====================================================

    With ws.Range("E12:G16")

        .Interior.Color = RGB(232, 247, 238)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 220, 190)

    End With


    AddAutomationCheckbox _
        ws, _
        ws.Range("E13").Left + 8, _
        ws.Range("E13").Top + 5, _
        "PIVOT"


    With ws.Range("F13:G14")

        .Merge

        .value = "PIVOT TABLES"

        .Font.Bold = True
        .Font.Size = 15
        .Font.Color = RGB(35, 125, 70)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    With ws.Range("E15:G15")

        .Merge

        .value = "Create analytical pivot tables"

        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    '====================================================
    ' DASHBOARD CARD
    '====================================================

    With ws.Range("H12:J16")

        .Interior.Color = RGB(255, 243, 225)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(240, 205, 160)

    End With


    AddAutomationCheckbox _
        ws, _
        ws.Range("H13").Left + 8, _
        ws.Range("H13").Top + 5, _
        "DASHBOARD"


    With ws.Range("I13:J14")

        .Merge

        .value = "DASHBOARD"

        .Font.Bold = True
        .Font.Size = 15
        .Font.Color = RGB(210, 120, 30)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    With ws.Range("H15:J15")

        .Merge

        .value = "Create visual business dashboard"

        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    '====================================================
    ' RUN AUTOMATION TITLE
    '====================================================

    With ws.Range("B18:J18")

        .Merge

        .value = "RUN AUTOMATION"

        .Font.Bold = True
        .Font.Size = 15
        .Font.Color = RGB(25, 35, 55)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With


    '====================================================
    ' RUN AUTOMATION BUTTON
    '====================================================

    Set shp = ws.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                ws.Range("D21").Left, _
                ws.Range("D21").Top, _
                400, _
                55)

    With shp

        .Name = "btnRunAutomation"

        .Fill.ForeColor.RGB = RGB(39, 174, 96)

        .Line.Visible = msoFalse

        .TextFrame2.TextRange.Text = "RUN AUTOMATION"

        .TextFrame2.TextRange.Font.Size = 15

        .TextFrame2.TextRange.Font.Bold = msoTrue

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        .OnAction = "RunAutomation"

    End With


    '====================================================
    ' FOOTER
    '====================================================

    With ws.Range("B25:J25")

        .Merge

        .value = "Merchandise Data Automation | Excel VBA"

        .Font.Size = 9
        .Font.Color = RGB(130, 130, 130)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With


    '====================================================
    ' SETTINGS
    '====================================================

    ws.Activate

    ActiveWindow.DisplayGridlines = False

    ActiveWindow.Zoom = 90


    Application.DisplayAlerts = True
    Application.ScreenUpdating = True


    MsgBox "Dashboard created successfully.", _
           vbInformation, _
           "Dashboard Ready"

End Sub


'========================================================
' ADD CHECKBOX
'========================================================

Private Sub AddAutomationCheckbox( _
    ByVal ws As Worksheet, _
    ByVal LeftPosition As Double, _
    ByVal TopPosition As Double, _
    ByVal OptionName As String)

    Dim chk As Shape

    Set chk = ws.Shapes.AddFormControl( _
                xlCheckBox, _
                LeftPosition, _
                TopPosition, _
                20, _
                20)

    With chk

        .AlternativeText = "AUTOMATION_" & OptionName

        .TextFrame.Characters.Text = ""

        .ControlFormat.value = xlOff

    End With

End Sub


'========================================================
' SHEET EXISTS
'========================================================

Public Function SheetExists( _
    ByVal sheetName As String, _
    ByVal wb As Workbook) As Boolean

    Dim ws As Worksheet

    SheetExists = False

    For Each ws In wb.Worksheets

        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then

            SheetExists = True

            Exit Function

        End If

    Next ws

End Function

