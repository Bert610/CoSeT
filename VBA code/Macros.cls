VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' callbacks from the macro buttons

Private Sub Export_Click()
    Export_Sub
End Sub

Public Sub Export_Sub()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    If ExportCompetitionWorkbook = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    FreeArrays
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub BuildMasterScoresheet_Click()
    BuildMasterScoresheet_Sub
End Sub
Public Sub BuildMasterScoresheet_Sub()
    InitMessages
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If LoadScoresAndComments = False Then
        ReportMessages
    End If
    ReportMessages
    FreeArrays
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "All done."
End Sub



Private Sub GenerateExpertiseSheets_Click()
    GenerateExpertiseSheets_sub
End Sub
Public Sub GenerateExpertiseSheets_sub()
    
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    
    MakeProjectExpertiseSheets
    
    ReportMessages
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub CreateKeywordExpertiseSheets_Click()
    CreateKeywordExpertiseSheets_Sub
End Sub
Public Sub CreateKeywordExpertiseSheets_Sub()
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    
    MakeKeywordExpertiseSheets
    
    ReportMessages
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub


Private Sub LoadExpertiseByKeyword_Click()
    LoadExpertiseByKeyword_Sub
End Sub
Public Sub LoadExpertiseByKeyword_Sub()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If LoadMarkerKeywordExpertiseIntoPXM = False Then
        ReportMessages
        Exit Sub
    End If
    If CreatePXMFromProjectRelevanceAndMarkerExpertise = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    ThisWorkbook.Activate
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub LoadPerProjectExpertise_Click()
    LoadExpertiseByProject_Sub
End Sub
Public Sub LoadExpertiseByProject_Sub()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If LoadMarkerProjectExpertiseIntoPXM = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    ThisWorkbook.Activate
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub MakeAssignments_Click()
    MakeAssignments_Sub
End Sub
Public Sub MakeAssignments_Sub()
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    InitMessages
    If AssignMarkers = False Then
        ReportMessages
        Exit Sub
    End If
    If PopulateResultsSheet = False Then
        ReportMessages
        Exit Sub
    End If
    If PopulateAnalysisSheet(False) = False Then
        ReportMessages
        Exit Sub
    End If
    If MakeAverageAndSpanChart = False Then
        ReportMessages
        Exit Sub
    End If
    ShowMasterAssignmentsSheet
    ReportMessages
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Public Sub PopulateResultsAndAnalysis_Sub()
    ' allows for rudimentary editing of the Assignments Master
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    InitMessages
    If PopulateResultsSheet = False Then
        ReportMessages
        Exit Sub
    End If
    If PopulateAnalysisSheet(False) = False Then
        ReportMessages
        Exit Sub
    End If
    If MakeAverageAndSpanChart = False Then
        ReportMessages
        Exit Sub
    End If
    ShowMasterAssignmentsSheet
    ReportMessages
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Public Sub PopulateAnalysisSheet_Sub()
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If ActivateCompetitionWorkbook = False Then
        Exit Sub
    End If
    making_competition_workbook = False
    If DefineGlobals = False Then
        ReportMessages
        Exit Sub
    End If
    InitMessages
    If PopulateAnalysisSheet(True) = False Then
        Exit Sub
    End If
    ' make sure the chart exists
    If MakeAverageAndSpanChart = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub CreateMarkerScoresheets1_Click()
    CreateMarkerScoresheets_Sub
End Sub
Public Sub CreateMarkerScoresheets_Sub()
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If CreateAllMarkingSheets = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub CalculateFOM_Click()
    CalculateProjectByMarkerTable_Sub
End Sub
Public Sub CalculateProjectByMarkerTable_Sub()
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    InitMessages
    If (CreateFigureOfMeritTable = False) Then
        ReportMessages
        Exit Sub
    Else
        AddMessage "Finished combining Project and Expertise keyword ratings for " & num_projects & " projects, " & _
                    num_markers & " markers."
    End If
    ThisWorkbook.Activate
    Sheets(start_sheet).Activate
    ReportMessages
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub Expertise2Scoresheets_Click()
    Expertise2Scoresheets_Sub
End Sub
Public Sub Expertise2Scoresheets_Sub()
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Expertise2MarkingSheets
    ThisWorkbook.Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub


Private Sub FOM2Scoresheets_Click()
    InitMessages
    If KeywordTablesToScoresheets = False Then
    End If
    ReportMessages
    FreeArrays
End Sub
Private Sub HideUnusedColumnsRows_Click()
    HideUnusedColumnsAndRows_Sub
End Sub
Public Sub HideUnusedColumnsAndRows_Sub()
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    If HideUnusedColumnsRowsF = False Then
        ReportMessages
        Exit Sub
    End If
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    ReportMessages
    FreeArrays
    Application.StatusBar = "All done."
End Sub
Private Sub ShowAllRowsAndColumns_Click()
    ShowAllRowsAndColumns_Sub
End Sub

Public Sub ShowAllRowsAndColumns_Sub()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If UnhideUnusedColumnsRowsF = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub

