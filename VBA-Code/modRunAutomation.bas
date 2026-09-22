Attribute VB_Name = "modRunAutomation"
Option Explicit

Public Sub RunAutomation()

    Dim doClean As Boolean
    Dim doPivot As Boolean
    Dim doDashboard As Boolean

    If Not RunSheetExists("OriginalData") Then

        MsgBox "Please upload a dataset first.", _
               vbExclamation, "Dataset Required"

        Exit Sub

    End If

    'Read dashboard checkboxes
    doClean = GetAutomationCheckbox("CLEAN")
    doPivot = GetAutomationCheckbox("PIVOT")
    doDashboard = GetAutomationCheckbox("DASHBOARD")

    'At least one option must be selected
    If Not doClean And Not doPivot And Not doDashboard Then

        MsgBox "Please select at least one automation option.", _
               vbExclamation, "No Option Selected"

        Exit Sub

    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    '==================================================
    ' CLEAN DATA
    '==================================================

    If doClean Then

        CreateCleanCopy
        CleanDataset

    End If


    '==================================================
    ' PIVOT TABLES
    '==================================================

    If doPivot Then

        'Pivot tables require cleaned data
        If Not RunSheetExists("CleanedData") Then

            CreateCleanCopy
            CleanDataset

        End If

        CreateAllPivots

    End If


    '==================================================
    ' DASHBOARD
    '==================================================

    If doDashboard Then

        'Dashboard requires cleaned data
        If Not RunSheetExists("CleanedData") Then

            CreateCleanCopy
            CleanDataset

        End If

        CreateDashboardReport

    End If


    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox "Automation completed successfully.", _
           vbInformation, "Automation Complete"

    Exit Sub


'==================================================
' ERROR HANDLER
'==================================================

ErrorHandler:

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Automation stopped." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "Automation Error"

End Sub


'==================================================
' GET CHECKBOX VALUE
'==================================================

Public Function GetAutomationCheckbox( _
    ByVal OptionName As String) As Boolean

    Dim ws As Worksheet
    Dim shp As Shape

    GetAutomationCheckbox = False

    Set ws = ThisWorkbook.Worksheets("Dashboard")

    For Each shp In ws.Shapes

        If StrComp( _
            shp.AlternativeText, _
            "AUTOMATION_" & OptionName, _
            vbTextCompare) = 0 Then

            If shp.Type = msoFormControl Then

                If shp.FormControlType = xlCheckBox Then

                    GetAutomationCheckbox = _
                        (shp.ControlFormat.value = xlOn)

                    Exit Function

                End If

            End If

        End If

    Next shp

End Function


'==================================================
' CHECK IF SHEET EXISTS
'==================================================

Private Function RunSheetExists( _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    RunSheetExists = False

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If Not ws Is Nothing Then

        RunSheetExists = True

    End If

End Function

