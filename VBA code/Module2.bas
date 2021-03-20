Attribute VB_Name = "Module2"
    Option Explicit
    
    Const SYSTEM_PARAMETERS_SHEET As String = "System Parameters"
    Const SP_MAX_NUM_OF_CRITERIA_RANGE As String = "C2"
    Const SP_MAX_PROJECTS_CELL As String = "C3"
    Const SP_MAX_NUMBER_OF_MARKERS_CELL As String = "C4"
    Const SP_MAX_NUMBER_OF_ASSIGNMENTS_PER_MARKER As String = "C5"
    Const SP_MAX_NUMBER_OF_MARKERS_PER_PROJ As String = "C6" 'where to max # markers per project supported
    Const SP_MAX_KEYWORDS_CELL As String = "C7"
    Const SP_LOCKED_SHEET_PWD_CELL As String = "C8"         ' where to find the sheet password on the parameters sheet.
    Const SP_PROJECT_EXPERTISE_FILE_PATTERN = "C9"          ' used for loading expertise info about the markers
    Const SP_KEYWORD_EXPERTISE_FILE_PATTERN = "C10"          ' used for loading expertise info about the markers
    Const SP_MARKS_ONLY_FILE_PATTERN As String = "C11"   ' for loading the scores from markers
    Const SP_MARKS_AND_CMTS_FILE_PATTERN As String = "C12"
    Const SP_SAME_ORGANIZATION_TEXT_CELL As String = "C13"
    Const SP_SIMULATE_MARKER_RESPONSES_CELL As String = "C14"
    
    Const COMPETITION_PARAMETERS_SHEET As String = "Competition Parameters"
    Const CP_TARGET_MARKERS_PER_PROJ As String = "C3"   'how many markers per project are desired
    Const CP_TARGET_ASSIGNMENTS_PER_MARKER = "C4"       'how many many projects assigned to a marker are desired
    Const CP_NUM_KEYWORDS_CELL As String = "C5"
    Const CP_COMPETITION_ROOT_FOLDER As String = "C10"
    Const CP_EXPERTISE_BY_PROJECT_REQUESTED_FOLDER_CELL As String = "C11"
    Const CP_EXPERTISE_BY_PROJECT_RECEIVED_FOLDER_CELL As String = "C12"
    Const CP_EXPERTISE_BY_KEYWORD_REQUESTED_FOLDER_CELL As String = "C13"
    Const CP_EXPERTISE_BY_KEYWORD_RECEIVED_FOLDER_CELL As String = "C14"
    Const CP_SCORES_REQUESTED_FOLDER_CELL As String = "C15"
    Const CP_SCORES_RECEIVED_CELL As String = "C16"
    Const CP_COMMENTS_FOLDER_CELL As String = "C17"
    Const CP_USE_ORG_DISAMBIGUATION_CELL As String = "C18"
    Const CP_USE_EMAIL_DISAMBIGUATION_CELL As String = "C19"
    Const CP_USE_NORMALIZED_SCORING_CELL = "C20"
    Const CP_GATHER_COMMENTS_CELL = "C21"
    Const CP_COMMENT_OUTPUT_FORMAT As String = "C22"
    Const CP_BLANK_EXPERTISE_TREATMENT As String = "C23"
    Const CP_MAX_FIRST_READER_ASSIGNMENTS_CELL As String = "K15"
    
    Const CRITERIA_SHEET As String = "Criteria"
    Const C_NUM_CRITERIA_CELL As String = "H1"
    Const C_FIRST_DATA_ROW As Long = 3
    Const C_FIRST_CRITERIA_MINVALUE_RN As Long = 3
    Const C_FIRST_CRITERIA_MINVALUE_CN As Long = 3
    Const C_TRANSPOSE_COPY_CELL As String = "G9"
    
    Const PROJECTS_SHEET As String = "Projects"
    Const P_NUM_PROJECTS_CELL As String = "L1"
    Const P_PROJECT_NAME_COLUMN As Long = 2
    Const P_CONTACT_NAME_COLUMN As Long = 3
    Const P_ORG_COLUMN As Long = 4                 ' organization of the submitters
    Const P_CONTACT_EMAIL_COLUMN As Long = 5
    Const P_MENTOR_ID_COLUMN As Long = 7
    Const P_FIRST_DATA_ROW As Long = 3
    
    Const MARKERS_SHEET As String = "Markers"      ' sheet containg the marker names and their associated marker number
    Const M_NUM_MARKERS_CELL As String = "G1"      ' cell containing a count of all the people who registered as markers
    Const M_NUMBER_COL As Long = 1
    Const M_NAME_COL As Long = 2
    Const M_ORG_COL As Long = 3
    Const M_EMAIL_COL As Long = 4
    Const M_NUM_TEAMS_MENTORED_COL As Long = 5
    Const M_FIRST_DATA_ROW As Long = 2
    
    Const KEYWORDS_SHEET As String = "Keywords"
    Const KEYWORD_STRING As String = "Keyword"
    Const KW_NUM_KEYWORDS_CELL As String = "G2"
    Const KW_KEYWORDS_COL As Long = 3
    Const KW_WEIGHTS_COL As Long = 4
    Const KW_FIRST_DATA_ROW As Long = 3
    Const KW_WEIGHTS_ROW As Long = 3
    Const KW_TRANSPOSE_COPY_CELL As String = "F4"
    
    Const PROJECT_KEYWORDS_SHEET As String = "Project Keywords"
    Const PK_FIRST_DATA_COL As Long = 3
    Const PK_FIRST_DATA_ROW As Long = 4
    Const PK_COLUMNS_BETWEEN_DATA_TABLES As Long = 3
    
    Const MARKER_EXPERTISE_SHEET As String = "Marker Expertise"
    Const ME_FIRST_DATA_COL As Long = 3
    Const ME_FIRST_DATA_ROW As Long = 4
    Const ME_COLUMNS_BETWEEN_DATA_TABLES As Long = 3
    Const ME_NUM_COLUMNS2SECOND_TABLE As Long = 5

    Const PROJECT_X_MARKER_SHEET As String = "Project X Marker Table"
    Const PXM_FIRST_DATA_ROW As Long = 4
    Const PXM_FIRST_DATA_COL As Long = 4
    Const PXM_MARKER_NUM_ROW As Long = 1
    
    Const EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET As String = "Expertise by Projects - Instr."
    
    Const EXPERTISE_CROSSWALK_SHEET As String = "Expertise Crosswalk"
    Const EC_FIRST_DATA_ROW As Long = 7
    Const EC_XLMH_CONF_PER_PROJECT_COL As Long = 5          ' 4 col. with COIs, low, medium, high about a project
    Const EC_ASSMT_CONF_FIRST_COL As Long = 10              ' col. with project confidence for marker assigned
    Const EC_XLMH_MARKER_TABLE_FIRST_ROW As Long = 2        ' 4 rows with COIs, low, medium, high about a marker
        
    Const MASTER_ASSIGNMENTS_SHEET As String = "Assignments Master"
    Const MAS_FIRST_ASSMT_COL As Long = 6               ' column "F" columns that span the marking assignments
    Const MAS_MENTOR_NUM_COL As Long = 4
    Const MAS_FIRST_ASSMT_ROW As Long = 3
    Const MAS_NUM_ASSIGNMENTS_COL As Long = 17          '# of assignments for each reader
    
    Const MSS_PROJECT_COL As Long = 1
    Const MSS_FIRST_PROJECT_ROW As Long = 6
    Const MSS_FIRST_SCORE_COL As Long = 9
    Const MSS_TOTAL_SCORES_COL As Long = 6       ' column with the total scores for a project
    Const MSS_LAST_COL As Long = 127                 ' last column of the master storing sheet
    Const MSS_MARKER_NUMBER_ROW As Long = 2         ' for "Marker #N (Normalized/Raw) Criteria Scores"
    
    Const RESULTS_SHEET As String = "Results"
    Const R_MARKER_NUM_COLUMN As Long = 1
    Const R_READER_NUM_COLUMN As Long = 3
    Const R_EXPERTISE_LETTERS_COLUMN As Long = 4
    Const R_PROJECT_NUM_COLUMN As Long = 5
    Const R_FIRST_RAW_COLUMN As Long = 8
    Const R_FIRST_DATA_ROW As Long = 5
    ' define the Start column #s and the Number of columns for each of the four tables on this sheet
    Const R_T1S As Long = 1
    Const R_T1N As Long = 13
    Const R_T2S As Long = 15
    Const R_T2N As Long = 6
    Const R_T3S As Long = 22
    Const R_T3N As Long = 9
    Const R_T4S As Long = 32
    Const R_T4N As Long = 7

    Const ANALYSIS_SHEET As String = "Analysis"
    Const A_PROJ_NUM_COL As Long = 2
    Const A_FIRST_RAW_READER_COLUMN As Long = 3
    Const A_FIRST_DATA_ROW As Long = 4
    ' define the Start column #s and the Number of columns for each of the tables on this sheet
    Const A_T1S As Long = 1
    Const A_T1N As Long = 11
    Const A_T2S As Long = 12
    Const A_T2N As Long = 10
    Const ANALYSIS_CHART_NAME As String = "Analysis Chart"
    
    Const EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET As String = "Expertise by Keywords - Instr."
    Const EBKI_FOR_MORE_INFO_ROW As Long = 3
    
    Const MARKER_PROJECT_EXPERTISE_TEMPLATE As String = "Marker Project - template"
    Const MPET_FIRST_DATA_ROW As Long = 2
    Const MPET_COI_COLUMN As Long = 6
    Const MPET_EXPERTISE_COLUMN As Long = 7
    Const MPET_MARKER_INFO_COLUMN As Long = 11
    
    Const EBPI_FOR_MORE_INFO_ROW As Long = 6    'expertise by project - instructions ...
    
    Const MARKER_KEYWORD_EXPERTISE_TEMPLATE As String = "Marker Keyword - template"
    Const MKET_FIRST_DATA_ROW As Long = 2
    Const MKET_EXPERTISE_COLUMN As String = "C"
    Const MKET_MARKER_NAME_CELL As String = "F1"
    Const MKET_MARKER_NUM_CELL As String = "F2"
    Const MKET_COI_SHEET_NAME As String = "Conflicts of Interest"
    
    Const MARKER_SCORING_TEMPLATE As String = "Marker Scoresheet - template"
    Const MST_MARKER_NUMBER_CELL As String = "B1"
    Const MST_MARKER_NAME_CELL As String = "D1"
    Const MST_FIRST_SCORING_COL As Long = 4
    Const MST_FIRST_SCORING_ROW As Long = 9             ' row of first project score in marker' sheet
    Const MST_READER_NUM_COL As Long = 3
    Const MST_PROJECT_NUM_COL As Long = 1
    
    Const SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE As String = "Instructions to Markers"
    Const SCI_INSTRUCTION_SHEET_NAME As String = "Instructions"
    Const SCI_COMPETITION_NAME_CELL As String = "A1"
    Const SCI_FOR_MORE_INFO_CELL As String = "A4"
    Const SCI_FIRST_DATA_ROW As Long = 8
    Const SCI_PROJECT_COUNT_ROW As Long = 6
    
    Const SCORES_AND_COMMENTS_TEMPLATE_SHEET As String = "Scores and Comments - template"
    Const SCT_COMPETITION_NAME_CELL = "B1"
    Const SCT_MARKER_NAME_CELL As String = "B2"
    Const SCT_MARKER_NUM_CELL As String = "B3"
    Const SCT_PROJECT_NUM_CELL As String = "B4"
    Const SCT_PROJECT_NAME_CELL As String = "B5"
    Const SCT_READER_NUM_CELL As String = "E4"
    Const SCT_CRITERIA_ONE_NAME_CELL As String = "D12"
    Const SCT_CRITERIA_ONE_ROW As Long = 12
    Const SCT_CRITERIA_ONE_MIN_CELL As String = "B13"
    Const SCT_CRITERIA_ONE_MAX_CELL As String = "D13"
    Const SCT_FIRST_CRITERIA_SCORE As String = "F13"
    Const SCT_SCORE_CHECK_CELL As String = "G2"
    Const SCT_COI_RESPONSE_CELL As String = "K6"
    Const SCT_CONFIDENCE_LOW_CELL As String = "B8"
    Const SCT_CONFIDENCE_MEDIUM_CELL As String = "E8"
    Const SCT_CONFIDENCE_HIGH_CELL As String = "H8"
    Const SCT_SCORE_COLUMN As Long = 6
    Const SCT_GENERAL_COMMENT_CELL As String = "A10"
    Const SCT_ROWS_PER_CRITERIA As Long = 5
'    Const SCT_FULL_SHEET_RANGE As String = "A1:K61"
    
    Const PROJECT_COMMENTS_SHEET As String = "Project Comments - template"
    Const PC_PROJECT_NUM_CELL As String = "B2"
    Const PC_PROJECT_NAME_CELL As String = "B3"
    Const PC_GENERAL_COMMENTS_CELL As String = "A5"
    Const PC_FIRST_CRITERIA_COMMENTS_ROW As Long = 9
    Const PC_ROWS_PER_CRITERIA As Long = 4
    
    Const MACROS_SHEET As String = "Macros"
    
    Const COMPETITION_WORKBOOK_DEFAULT_NAME As String = "CompetitionWorkbook.xlsx"
    
    Const ONE_THIRD As Double = 1 / 3
    Const TWO_THIRDS As Double = 2 / 3
    Const MAX_EMAIL_LENGTH As Long = 30
    Const COSET As String = "CoSeT"
    
' GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS
    
    Global main_workbook As Variant
    Global cwb As String                        'name of the competition workbook
    Global num_criteria As Long                 ' number from the Criteria sheet
    Global num_projects As Long                 ' number from the Projects sheet
    Global num_projects_scored As Long          ' accounts for did-not-finish
    Global num_markers As Long                  ' number of people from the Markers sheet (calculated)
    Global max_markers_per_proj As Long         ' max # of markers per project supported by this set of tools
    Global target_markers_per_proj As Long      ' # of markers as specified in the competition parameters sheet
    Global max_ass_per_marker As Long
    Global target_ass_per_marker As Long
    Global max_first_reader_assignments As Long
    Global n_per_assignment_col() As Long       ' number of readers assigned at this level
    Global retain_worksheets As Boolean
    Global num_keywords As Long
    Global max_keywords As Long
    Global normalize_scoring As Boolean
    Global project_expertise_file_pattern As String
    Global keyword_expertise_file_pattern As String
    Global expertise_by_project_ending As String
    Global expertise_by_keyword_ending As String
    Global marks_only_file_pattern As String            'e.g., * marks.xlsx
    Global marks_and_cmts_file_pattern As String
    Global same_organization_text As String
    Global use_email_disambiguation As Boolean      ' for disambiguating personalized files
    Global use_org_disambiguation As Boolean
    Global gather_comments As Boolean
    Global output_comments_format As String
    Global blank_expertise_means_exclusion  As Boolean ' should blanks be treated as exclusions (True) or low expertise (False)
    Global mst_first_normalized_score_column As Long
    Global ec_assignments_first_column As Long     ' col. with marker numbers assigned to a project
    Global ec_data_first_marker_column As Long
'    Global max_projects As Long
'    Global max_markers As Long
'    Global max_criteria As Long
   
    Global current_expertise_sheet As String          'this next bunch is for the expertise requests
    Global COI_sheet As String
    Global lock_sheet_pwd As String
    Global expertise_template_workbook As String
    Global expertise_col As Long
    
    Global root_folder As String
    Global expertise_by_project_requested_folder As String
    Global expertise_by_project_received_folder As String
    Global expertise_by_keyword_requested_folder As String
    Global expertise_by_keyword_received_folder As String
    Global scores_received_folder As String
    Global scores_requested_folder As String
    Global comments_folder As String
    
    Global scores_only_ending As String          ' for sheets with only scores
    Global scores_and_cmts_ending As String        ' for sheets with comments and scores
    
    ' Arrays for assigning markers
    Global pxm_table() As Variant           'table of project/marker confidences, eXclusions and Assignments
    Global pxm_at_start() As Variant        'as read in from the expertise crosswalk
    Global pn_array() As Variant            ' pn for project number (project numbers can be in random order)
    Global mn_array() As Variant            ' mn for marker number (marker numbers can be in random order)
    Global coa_array() As Variant           ' confidence of the assigned marker array
        ' dimension 1 = rows, one project's assigned markers' COA
        ' dimension 2 has N items, one for each possible marker's confidence
    Global mentor_column() As Variant       ' column of COI text for expertise sheet
    Global competition_COIs() As Variant    ' array of COIs for all markers on all projects
    Global ss_marker_col() As Variant       ' the marker# column in the raw table
    Global ss_project_col() As Variant      ' the project# column in the raw table
    
    Global xlmh_per_marker() As Variant     'number of eXcluded, Low, Medium & High selections for a marker
    Global xlmh_per_project() As Variant    'number of eXcluded, Low, Medium & High markers for a project
    Global assignments() As Variant         ' markers/readers assigned to the projects
        ' dimension 1 = rows, one per project
        ' dimension 2 has N items, with the number of each marker assigned
    Global n_assigned2project() As Long        ' count on the number of markers assigned to a project
    Global n_assigned2marker() As Long      ' count on the number of projects assigned to a marker
    Global markers_table() As Variant       ' Table read from the markers sheet
    Global projects_table() As Variant       ' Table read from the markers sheet
    Global comments() As Variant            ' array of comments from each marker for each criteria on each project
                                            ' Projects x Criteria
    Global general_comments() As Variant    ' array of the general comments provided by markers (by projects)
    
    Global assignment_failed_for_this_proj() As Boolean
    
    ' these 3 arrays are used to store scores, moved to global to see if it avoids crashes.
    ' this is for one marker
    Global pn_1m() As Variant, rn_1m() As Variant, mn_1m() As Variant
    Global marker_scores() As Variant
    ' this is the full column for the full competition
    Global proj_num_col() As Variant, reader_num_col() As Variant, marker_num_col() As Variant
    Global competition_scores() As Variant, num_competition_scores As Long
    Global messages() As String             ' for buffering messages
    Global num_messages As Long
    Global buffer_messages As Boolean
    
    Global simulate_marker_responses As Boolean
    
    Global making_competition_workbook As Boolean   ' flag that the macro to build the CWB is active
    Global globals_defined As Boolean
    Global fps As String                        ' folder path separator (mac vs windows)
    

' bundle the middle steps together
Sub Expertise2MarkingSheets()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    InitMessages
    
    If LoadMarkerExpertiseIntoCrosswalk = False Then
        ReportMessages
        Exit Sub
    End If
    
    
    If AssignMarkers = False Then
        ReportMessages
        Exit Sub
    End If
    
    If CreateAllMarkingSheets = False Then
        ReportMessages
        Exit Sub
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ReportMessages
    
End Sub

Public Function KeywordTablesToScoresheets() As Boolean

    
    globals_defined = False
    making_competition_workbook = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    If (CreatePXMFromProjectRelevanceAndMarkerExpertise = False) Then
        Exit Function
    End If
    
    If AssignMarkers = False Then
        ReportMessages
        Exit Function
    End If
    
    If CreateAllMarkingSheets = False Then
        ReportMessages
        Exit Function
    End If
    
End Function

Sub UnsortMasterScoresheet()
'
' UnsortMasterScoresheet Macro
'
    MsgBox "needs updating for V6 unlimits on projects and markers", vbCritical
    Exit Sub
    
    making_competition_workbook = False
    DefineGlobals
    Dim start_address As String
    start_address = ActiveCell.Address

    Const sort_key_col As String = "A"
    Dim sort_range As String
    sort_range = sort_key_col & MSS_FIRST_PROJECT_ROW & ":" & _
                  c2l(MSS_LAST_COL + 2) & (MSS_FIRST_PROJECT_ROW + max_projects - 1)
    Range(FirstCell(sort_range)).Select
    Range(FirstCell(sort_range)).Activate
    Dim first_row As Long, sort_key As String
    first_row = ActiveCell.row
    sort_key = sort_key_col & first_row & ":" & sort_key_col & (first_row + max_projects - 1)
    ActiveWorkbook.Worksheets(RESULTS_BY_READER_SHEET).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(RESULTS_BY_READER_SHEET).Sort.SortFields.Add2 Key:=Range(sort_key), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(RESULTS_BY_READER_SHEET).Sort
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' go back to where we started
    Range(start_address).Select
End Sub

Public Sub MakeProjectExpertiseSheets()
    InitMessages
    WriteAllExpertiseSheets expertise_by_project_requested_folder, expertise_by_project_ending, ""
    ReportMessages
End Sub

Public Sub MakeKeywordExpertiseSheets()
    InitMessages
    WriteAllExpertiseSheets expertise_by_keyword_requested_folder, expertise_by_keyword_ending, KEYWORD_STRING
    ReportMessages
End Sub

Public Function WriteAllExpertiseSheets(expertise_type As String, expertise_type_ending As String, _
                                        keyword_or_project As String) As Boolean
    'Turn off events and screen flickering.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim start_range As String
    start_range = ActiveCell.Address
    Dim starting_sheet_name As String, n_expertise_books As Long
    starting_sheet_name = ActiveSheet.Name
    If ActivateCompetitionWorkbook = False Then
        WriteAllExpertiseSheets = False
        Exit Function
    End If
        
    making_competition_workbook = False
    globals_defined = False
    making_competition_workbook = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    LoadMarkerAndProjectTables      ' need to fill in the COI column
    Sheets(MARKERS_SHEET).Activate
    
    ' as long as we have a marker number and marker name create a marker expertise sheet
    Dim marker_name As String, marker_number As Long
    Dim i As Long
    
    Dim folder_for_expertises As String
'    ChDir root_folder
    folder_for_expertises = SelectFolder("Select folder to store the blank" & expertise_type_ending & " sheets", _
                root_folder & expertise_type)
    If Len(folder_for_expertises) = 0 Then
        Exit Function
    End If

    ' read in the arrays of markers
    Dim markers_range As String
    markers_range = c2l(M_NUMBER_COL) & M_FIRST_DATA_ROW & ":" & _
                    c2l(M_EMAIL_COL) & M_FIRST_DATA_ROW + num_markers - 1
    markers_table = Range(markers_range)
    ' look through the rows on the markers sheet and create an expertise workbook for each person
    expertise_template_workbook = ""
    For i = 1 To num_markers
        marker_number = markers_table(i, M_NUMBER_COL)
        marker_name = markers_table(i, M_NAME_COL)
        If (marker_number <> i) Or (Len(marker_name) = 0) Then
           AddMessage "[writeAllExpertiseSheets] expected marker number and name, got [" & marker_number _
                    & "] and {" & marker_name & "} - exiting."
            Exit Function
        End If
        If keyword_or_project = KEYWORD_STRING Then
            WriteOneKeywordExpertiseBook marker_number, marker_name, folder_for_expertises
        Else
            WriteOneProjectExpertiseBook marker_number, marker_name, folder_for_expertises
        End If
        n_expertise_books = n_expertise_books + 1
    Next i
    
    ' close the template workbook that was used to write out the files
    If keyword_or_project = KEYWORD_STRING Then
        WriteOneKeywordExpertiseBook -1, marker_name, folder_for_expertises
    Else
        WriteOneProjectExpertiseBook -1, marker_name, folder_for_expertises
    End If
    
    ThisWorkbook.Activate
    Sheets(starting_sheet_name).Activate
    Range(start_range).Select
    Range(start_range).Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    AddMessage "Created " & n_expertise_books & " sheets for markers to provide their confidence about marking projects."

End Function

Public Function WriteOneProjectExpertiseBook(marker_num As Long, marker_name As String, folder As String) As String

    Dim form_str As String
    form_str = "=IF(LEN($K$1)>0,IF(AND(LEN(D2)>0,VLOOKUP($K$1,Markers!A:D,3,FALSE)=" & _
            "'Marker Project - template'!D2),""SAME ORGANIZATION"",IF(ISNA(VLOOKUP(A2,Projects!A:G,5,FALSE))" & _
            ","""",IF(VLOOKUP(A2,Projects!A:G,7,FALSE)='Marker Project - template'!$K$1,""MENTOR"",""""))),"""")"

    ' create the workbook that has the Project expertise sheet and associated instructions for each expert
    If marker_num < 0 Then
        If Len(expertise_template_workbook) = 0 Then
            PopMessage "Expecting to close project expertise book, but not available", vbCritical
            Exit Function
        Else
            Workbooks(expertise_template_workbook).Activate
            ActiveWindow.Close
            expertise_template_workbook = ""
            Workbooks(cwb).Activate
            Sheets(Array(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, MARKER_PROJECT_EXPERTISE_TEMPLATE)).Visible = False
            Exit Function
        End If
    End If
    
    Dim new_sheet As String
    If Len(expertise_template_workbook) = 0 Then
        ' create the workbook containing the expertise sheet and associated instructions
        
        ' make sure the templates we want are visible (they are hidden by default)
        Sheets(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET).Visible = True
        Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Visible = True
       
        ' move the new sheet out into its own workbook
        Sheets(Array(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, MARKER_PROJECT_EXPERTISE_TEMPLATE)).Select
        Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Activate
        Sheets(Array(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, MARKER_PROJECT_EXPERTISE_TEMPLATE)).Copy
        
        ' remove the external link in the instruction sheet
        Sheets(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET).Activate
        ConvertRangeToText "A1"
        ConvertRangeToText "A" & EBPI_FOR_MORE_INFO_ROW
        
        Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Select
        ' remove external links in the expertise template
        ConvertRangeToText "A1:E1"          ' column headers
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "A", num_projects
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "B", num_projects
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "C", num_projects
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "D", num_projects
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "E", num_projects
        
        ' rename the expertise sheet to the marker's number and name
        ActiveSheet.Name = GoodTabName(marker_num & " " & marker_name)
        current_expertise_sheet = ActiveSheet.Name
    Else
        ' for 2nd and onward marker
        Unlock2Sheets EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, ActiveSheet.Name
        Sheets(current_expertise_sheet).Select
        PutFormulaAndDragDown c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW, form_str, num_projects - 1
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, c2l(MPET_COI_COLUMN), num_projects
    End If
    
    ' name the expertise sheet to the marker's number and name
    Sheets(current_expertise_sheet).Select
    ActiveSheet.Name = GoodTabName(marker_num & " " & marker_name)
    current_expertise_sheet = ActiveSheet.Name

    ' store the marker number and name on this sheet
    Range(c2l(MPET_MARKER_INFO_COLUMN) & 1).Value = marker_num
    Range(c2l(MPET_MARKER_INFO_COLUMN) & 2).Value = marker_name
    ' pre-seed the COI column with any obvious conflicts (same org, or mentor roles)
    FillCOIcolumn MPET_COI_COLUMN, MPET_FIRST_DATA_ROW, marker_num

    If simulate_marker_responses Then
        ' read in the existing column (to capture an data already specified)
        Dim COI_range As String, COI_column() As Variant
        COI_range = c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                    c2l(MPET_COI_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1)
        COI_column = Range(COI_range)
        ' fill the expertise column of the sheet with made up high-medium-low
        ' so they can be used for assignment testing (but respect the COI info)
        Range(c2l(MPET_EXPERTISE_COLUMN) & MPET_FIRST_DATA_ROW).Select
        WriteRandomExpertiseRatings num_projects, True, COI_column
    End If
    
    ' lock the sheet
    LockUserProjectExpertiseSheet ActiveSheet.Name, MPET_EXPERTISE_COLUMN
    
    ' if requested, insert the marker's organization and email into the filename
    Dim file_stub As String
    file_stub = DisambiguateFilename(GoodTabName(marker_num & " " & marker_name), marker_num)
    ' write this workbook to the current directory and close it.
    WriteOneProjectExpertiseBook = folder & fps & file_stub & expertise_by_project_ending & ".xlsx"
    ActiveWorkbook.SaveAs Filename:=WriteOneProjectExpertiseBook, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    expertise_template_workbook = ActiveWorkbook.Name
    
End Function

Public Function WriteOneKeywordExpertiseBook(marker_num As Long, marker_name As String, folder As String) As String
    
    Dim form_str As String
    form_str = "=IF(LEN($I$1)>0,IF(AND(LEN(D2)>0,VLOOKUP($I$1,Markers!A:D,3,FALSE)=" & _
            "'Marker Project - template'!D2),""SAME ORGANIZATION"",IF(ISNA(VLOOKUP(A2,Projects!A:G,5,FALSE))" & _
            ","""",IF(VLOOKUP(A2,Projects!A:G,7,FALSE)='Marker Project - template'!$I$1,""MENTOR"",""""))),"""")"
    
    ' special case - marker_num negative means we are done, so delete the template workbook
    If marker_num < 0 Then
        If Len(expertise_template_workbook) = 0 Then
            PopMessage "Expecting to close Keyword Expertise book, but not available", vbCritical
            Exit Function
        Else
            Workbooks(expertise_template_workbook).Activate
            ActiveWindow.Close
            expertise_template_workbook = ""
            Workbooks(cwb).Activate
            Sheets(Array(EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, MARKER_KEYWORD_EXPERTISE_TEMPLATE, _
            MARKER_PROJECT_EXPERTISE_TEMPLATE)).Visible = False
            Exit Function
        End If
    End If
    
    If Len(expertise_template_workbook) = 0 Then
        ' create the workbook containing the expertise sheet and associated instructions
        
        ' make sure the templates we want are visible (they are hidden by default)
        Sheets(EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET).Visible = True
        Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Visible = True
        Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Visible = True
                    
        ' move the required sheets out into their own workbook
        Sheets(Array(EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, MARKER_KEYWORD_EXPERTISE_TEMPLATE, _
                    MARKER_PROJECT_EXPERTISE_TEMPLATE)).Select
        Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Activate
        Sheets(Array(EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, MARKER_KEYWORD_EXPERTISE_TEMPLATE, _
                    MARKER_PROJECT_EXPERTISE_TEMPLATE)).Copy
        expertise_template_workbook = ActiveWorkbook.Name
        
        ' remove the external link in the instruction sheet
        Sheets(EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET).Activate
        ConvertRangeToText "A1"
        ConvertRangeToText "A" & EBKI_FOR_MORE_INFO_ROW
        
        ' remove external links from the keyword rating sheet
        Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Select
        Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Activate
        ConvertCellsDownFromFormula2Text MKET_FIRST_DATA_ROW, "A", num_keywords
        ConvertCellsDownFromFormula2Text MKET_FIRST_DATA_ROW, "B", num_keywords
        current_expertise_sheet = MARKER_KEYWORD_EXPERTISE_TEMPLATE
    
        '   update what has become the COI sheet
        Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Select
        Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Activate
        ConvertRangeToText "A1:E1"
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "A", num_projects
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "B", num_projects
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "C", num_projects
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "D", num_projects
        ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "E", num_projects
    
        ' remove the expertise column, and the error checking column from the project expertise sheet
        ' with that, the sheet becomes a COI request sheet.
        expertise_col = FindHeaderColumn(1, "Expertise:", False)
        If expertise_col = 0 Then
            PopMessage "WriteOneKeywordExpertiseBook: unable to find Expertise: column", vbCritical
            Exit Function
        End If
        Columns(c2l(expertise_col) & ":" & c2l(expertise_col + 1)).Select
        Selection.Delete Shift:=xlToLeft
        ActiveSheet.Name = "Conflicts of Interest"
        COI_sheet = ActiveSheet.Name
    Else
        ' for 2nd and onward marker
        Unlock2Sheets current_expertise_sheet, COI_sheet
    End If
    
    ' name the expertise sheet to the marker's number and name
    Sheets(current_expertise_sheet).Select
    ActiveSheet.Name = GoodTabName(marker_num & " " & marker_name)
    current_expertise_sheet = ActiveSheet.Name
    
    ' store the marker number and name on the keyword expertise sheet
    Sheets(current_expertise_sheet).Range(MKET_MARKER_NUM_CELL).Value = marker_num
    Sheets(current_expertise_sheet).Range(MKET_MARKER_NAME_CELL).Value = marker_name
    ' do the same for the COI sheet, and also pre-seed any obvious COI's
    Sheets(COI_sheet).Select
    Range(c2l(MPET_MARKER_INFO_COLUMN - 2) & 1).Value = marker_num
    Range(c2l(MPET_MARKER_INFO_COLUMN - 2) & 2).Value = marker_name
    FillCOIcolumn MPET_COI_COLUMN, MPET_FIRST_DATA_ROW, marker_num
    
    If simulate_marker_responses Then
        ' fill the expertise column of the sheet with made up high-medium-low
        ' so they can be used for assignment testing (but respect the COI info)
        Sheets(current_expertise_sheet).Activate
        Range(MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW).Select
        Dim COI_column() As Variant
        WriteRandomExpertiseRatings num_keywords, False, COI_column
    End If
            
    ' Lock the appropriate sheets, and make the instructions sheet visible
    LockUserKeywordExpertAndCOISheets current_expertise_sheet, COI_sheet, expertise_col - 2
    Sheets(1).Select    ' make the instructions sheet visible
    Sheets(1).Activate
    
    ' if requested, insert the marker's organization and email into the filename
    Dim file_stub As String
    file_stub = DisambiguateFilename(GoodTabName(marker_num & " " & marker_name), marker_num)
    ' now save the file to disk
    WriteOneKeywordExpertiseBook = folder & fps & file_stub & expertise_by_keyword_ending & ".xlsx"
    ActiveWorkbook.SaveAs Filename:=WriteOneKeywordExpertiseBook, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    expertise_template_workbook = ActiveWorkbook.Name
End Function

Public Function LoadMarkerAndProjectTables() As Boolean
    Dim table_range As String
    Sheets(MARKERS_SHEET).Select
    Sheets(MARKERS_SHEET).Activate
    table_range = c2l(1) & M_FIRST_DATA_ROW & ":" & c2l(M_ORG_COL) & (M_FIRST_DATA_ROW + num_markers - 1)
    markers_table = Range(table_range)
    Sheets(PROJECTS_SHEET).Select
    Sheets(PROJECTS_SHEET).Activate
    table_range = c2l(1) & P_FIRST_DATA_ROW & ":" & c2l(P_MENTOR_ID_COLUMN) & (P_FIRST_DATA_ROW + num_projects - 1)
    projects_table = Range(table_range)
End Function

Public Function FillCOIcolumn(col_num As Long, start_row As Long, marker_num As Long) As Boolean
    ' check whether this marker is from the same organization or a reader on any of the projects.
    ' If so, put text in the column to signal this.
    
    'load the markers table
    Dim col_entries() As Variant
    ReDim col_entries(1 To num_projects, 1 To 1)
    
    Dim i As Long
    For i = 1 To num_projects
        ' check if they are the same organization
        If (Len(projects_table(i, P_ORG_COLUMN)) > 0) And _
            projects_table(i, P_ORG_COLUMN) = markers_table(marker_num, M_ORG_COL) Then
            col_entries(i, 1) = "SAME ORG"
        Else
            If projects_table(i, P_MENTOR_ID_COLUMN) = marker_num Then
                col_entries(i, 1) = "MENTOR"
            End If
        End If
    Next i
        
    ' write out the column of entries
    Dim Destination As Range
    Set Destination = Range(c2l(col_num) & start_row)
    Destination.Resize(UBound(col_entries, 1), UBound(col_entries, 2)).Value = col_entries
    Erase col_entries
    
End Function

Public Function DisambiguateFilename(filestub_in As String, marker_num As Long) As String
    ' add a version of the marker's organization and email if indicated by the competition parameters
    Dim cell_range As String, filestub As String
    filestub = filestub_in
    If use_org_disambiguation Then
        Dim org As String, marker_org As String
        cell_range = c2l(M_ORG_COL) & (M_FIRST_DATA_ROW + marker_num - 1)
        marker_org = ThisWorkbook.Sheets(MARKERS_SHEET).Range(cell_range).Value
        'read the org name for this expert
        org = Email2Text(marker_org, MAX_EMAIL_LENGTH)
        ' if the org is non-null tack it on
        If Len(org) > 0 Then
            filestub = filestub & " " & org
        End If
    End If
    If use_email_disambiguation Then
        Dim email As String, marker_email As String
        cell_range = c2l(M_EMAIL_COL) & (M_FIRST_DATA_ROW + marker_num - 1)
        marker_email = ThisWorkbook.Sheets(MARKERS_SHEET).Range(cell_range).Value
        'read the email address for this expert
        email = Email2Text(marker_email, MAX_EMAIL_LENGTH)
        ' if the email is non-null tack it on
        If Len(email) > 0 Then
            filestub = filestub & " " & email
        End If
    End If
    DisambiguateFilename = filestub
    
End Function

Public Function LockUserProjectExpertiseSheet(project_expertise_sheet As String, lc As Long) As Boolean
    'lock most of the sheet for recording COIs
    Sheets(project_expertise_sheet).Select
    Dim lock_range As String, free_range As String
    lock_range = "A:" & c2l(lc - 2) & "," & c2l(lc + 1) & ":" & c2l(lc + 1) & "," & _
                c2l(lc + 3) & "1:" & c2l(lc + 4) & "2," & c2l(lc) & "1"
    Range(lock_range).Select
    Range("A1").Activate
    Selection.Locked = True
    Selection.FormulaHidden = False
    
    ' make sure the areas for data entry are not locked.
    free_range = c2l(lc - 1) & MPET_FIRST_DATA_ROW & ":" & c2l(lc) & (MPET_FIRST_DATA_ROW - 1 + num_projects)
    Range(free_range).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' lock the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=lock_sheet_pwd

    LockUserProjectExpertiseSheet = True
End Function

Public Function Unlock2Sheets(Sheet1 As String, sheet2 As String)
    Sheets(Sheet1).Unprotect lock_sheet_pwd
    Sheets(sheet2).Unprotect lock_sheet_pwd
    
End Function
Public Function LockUserKeywordExpertAndCOISheets(keyword_sheet As String, C_O_I_sheet As String, _
                     lc As Long) As Boolean
'
    'lock most of the sheet for recording COIs
    Sheets(C_O_I_sheet).Select
    Dim lock_range As String, free_range As String
    lock_range = "A:" & c2l(lc) & "," & c2l(lc + 1) & "1," & c2l(lc + 3) & "1:" & c2l(lc + 4) & "2"
    Range(lock_range).Select
    Range("A1").Activate
    Selection.Locked = True
    Selection.FormulaHidden = False
    
    ' make sure the areas for data entry are not locked.
    free_range = c2l(lc + 1) & MKET_FIRST_DATA_ROW & ":" & c2l(lc + 1) & (MKET_FIRST_DATA_ROW - 1 + num_projects)
    Range(free_range).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' lock the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=lock_sheet_pwd
    
    'now protect most of the Keyword sheet
    Sheets(keyword_sheet).Select
    lock_range = "A:B,E1:F2"
    Range(lock_range).Select
    Range("C2").Activate
    Selection.Locked = True
    Selection.FormulaHidden = False
    ' make sure the areas for data entry are not locked.
    free_range = MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW & ":" & _
          MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW - 1 + max_keywords
    Range(free_range).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' lock the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=lock_sheet_pwd

End Function

Public Function random_expertise_rating_LMH() As String
    ' create a random rating letter grade: L, M or H, with the bias indicated by the case statement below.
    Const upperbound As Long = 100
    Const lowerbound As Long = 1
    Dim rating_letters(1 To 3) As String, rating As Long
    rating_letters(1) = "L"
    rating_letters(2) = "M"
    rating_letters(3) = "H"
    rating = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    Select Case rating  '40 null, 40 L, 17 M, 3 H
    Case 1 To 40
        If blank_expertise_means_exclusion Then
            random_expertise_rating_LMH = ""
        Else
            random_expertise_rating_LMH = rating_letters(1)
        End If
    Case 41 To 80
        random_expertise_rating_LMH = rating_letters(1)
    Case 81 To 97
        random_expertise_rating_LMH = rating_letters(2)
    Case 98 To 100
        random_expertise_rating_LMH = rating_letters(3)
    Case Else
        PopMessage "Error in case statement", vbCritical
    End Select
    
End Function

Function WriteRandomExpertiseRatings(num_ratings As Long, check_COI As Boolean, COI_column() As Variant) As Boolean
    ' write a column of random expertise ratings starting at the current location
    Dim i As Long, rating As Long
    Dim ratings_column() As String
    ReDim ratings_column(1 To num_ratings)
    ' create the random numbers
    For i = 1 To num_ratings
        If check_COI Then
            If Len(COI_column(i, 1)) = 0 Then
                ' bias the confidence ratings to low, then medium, then high
                ratings_column(i) = random_expertise_rating_LMH()
            Else
                ' this marker is in conflict for this project, flag it.
                ratings_column(i) = "X"
            End If
        Else
            ratings_column(i) = random_expertise_rating_LMH()
        End If
    Next i
    
    ' now write the array to the expertise column of the expertise sheet
    Dim Destination As Range
    Set Destination = Range(ActiveCell.Address)
    Set Destination = Destination.Resize(UBound(ratings_column), 1)
    Destination.Value = Application.Transpose(ratings_column)
    Erase ratings_column
    WriteRandomExpertiseRatings = True
End Function

Sub test_createallmarkingsheets()
    CreateAllMarkingSheets
End Sub

Function CreateScoringBook(num_score_sheets As Long) As Boolean
    Dim i As Long, sheet_names() As Variant
    sheet_names = Array(SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE, SCORES_AND_COMMENTS_TEMPLATE_SHEET)
    HideOrShowSheets sheet_names, True
    Sheets(sheet_names).Select
    Sheets(sheet_names).Copy
    Erase sheet_names
    
    ' remove external references from the instructions sheet
    Sheets(SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE).Name = SCI_INSTRUCTION_SHEET_NAME
    Sheets(SCI_INSTRUCTION_SHEET_NAME).Select
    ConvertRangeToText SCI_COMPETITION_NAME_CELL
    ConvertRangeToText SCI_FOR_MORE_INFO_CELL
    InsertAndExpandDown 1, SCI_FIRST_DATA_ROW, 3, num_score_sheets - 2
    
    ' build enough sheets in this workbook to hand the maximum # of assignments
    Sheets(SCORES_AND_COMMENTS_TEMPLATE_SHEET).Select
    Dim after_pos As Long
    For i = 2 To num_score_sheets
        after_pos = Sheets.Count
        Sheets(SCORES_AND_COMMENTS_TEMPLATE_SHEET).Copy after:=Sheets(after_pos)
    Next i
    CreateScoringBook = True
End Function

Public Function CreateAllMarkingSheets() As Boolean
    ' create the sheets for markers to enter their scoring of projects

' -----------------------------------------------------------------------------------------
' There was a bug with renaming sheets that occurs when the scoring book is re-used
' The bug (error 400 & exit) appears when trying to rename a sheet using the name of another sheet.
' If the bug re-occurs can be avoided by creating a new scoring book each time (slower, but functional)
' by setting this flag to FALSE
' only applies to the case where books for scores and comments are needed
Const reuse_scores_and_comments_book As Boolean = True
' -----------------------------------------------------------------------------------------

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'make sure we are starting in the right sheet
    Dim starting_sheet As String, starting_book As String
    starting_sheet = ActiveSheet.Name
    starting_book = ActiveWorkbook.Name
    If ActivateCompetitionWorkbook = False Then
        CreateAllMarkingSheets = False
        Exit Function
    End If

    globals_defined = False
    making_competition_workbook = False
    If DefineGlobals = False Then
        Exit Function
    End If

    ' check the number of markers
    If num_markers < 1 Then
        PopMessage "[CreateAllMarkingSheets] Need positive number of markers, found " & num_markers, vbOKOnly
        Exit Function
    End If
    
    ' load the markers and projects
    LoadMarkerAndProjectTables
    
   'load the assignments table
    Dim table_range As String, ass_per_marker_count() As Variant
    Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
    table_range = c2l(MAS_FIRST_ASSMT_COL) & MAS_FIRST_ASSMT_ROW & ":" & _
                c2l(MAS_FIRST_ASSMT_COL + target_markers_per_proj - 1) & _
                                                     (MAS_FIRST_ASSMT_ROW + num_projects - 1)
    assignments = Range(table_range)
    ' the assignments per marker
    table_range = c2l(MAS_NUM_ASSIGNMENTS_COL) & MAS_FIRST_ASSMT_ROW & ":" & _
                  c2l(MAS_NUM_ASSIGNMENTS_COL) & (MAS_FIRST_ASSMT_ROW + num_markers - 1)
    ass_per_marker_count = Range(table_range)
    Dim i As Long, max_assignments As Long
    For i = 1 To num_markers
        If ass_per_marker_count(i, 1) > max_assignments Then
            max_assignments = ass_per_marker_count(i, 1)
        End If
    Next i
    If max_assignments = 0 Then
        PopMessage "Cannot make score sheets - No markers have marking assignments yet!", vbCritical
        Exit Function
    End If
    
    ' Create the workbook that will be personalized for each marker
    Dim sheet_names() As Variant, tab_name As String
    If gather_comments Then
        If reuse_scores_and_comments_book Then
            CreateScoringBook max_assignments
        End If
    Else
        sheet_names = Array(MARKER_SCORING_TEMPLATE)
        HideOrShowSheets sheet_names, True
        Erase sheet_names
        Sheets(MARKER_SCORING_TEMPLATE).Select
        Sheets(MARKER_SCORING_TEMPLATE).Copy
        ' convert the target normalization fraction to text
        ' do it here in case it is edited by the user before the competition sheets are set
        ConvertRangeToText c2l(MST_FIRST_SCORING_COL + num_criteria) & (MST_FIRST_SCORING_ROW + 4)
    End If

    ' Ask the user to specify the folder where to store the scoring files
    scores_requested_folder = SelectFolder("Specify where to save the blank scoresheets", _
                            root_folder & scores_requested_folder)
    If Len(scores_requested_folder) = 0 Then
        Exit Function
    End If
    
    ' for each marker save a scoresheet or workbook for scores and comments containing their assigned projects
    Dim mn As Long, marker_name As String
    Dim marker_assmts() As Long, marker_rn() As Long, num_assigned As Long, insert_array() As Variant
    Dim Destination As Range, marking_sheet As String, filestub As String
    Dim num_scoring_files_created As Long, link_row(1 To 3) As Variant
    Dim rows_in_table As Long
    rows_in_table = 2       ' the scores-only sheet starts with two scoring rows
    num_scoring_files_created = 0
    For mn = 1 To num_markers
        marker_name = markers_table(mn, 2)
        
        ' get the assignments for this marker
        ReDim marker_assmts(1 To ass_per_marker_count(mn, 1))
        ReDim marker_rn(1 To ass_per_marker_count(mn, 1))
        num_assigned = GetAssignedProjects(mn, marker_assmts, marker_rn)
        
        ' now put the information in the sheet(s) for these assignments
        If num_assigned > 0 Then
            If gather_comments Then
                If reuse_scores_and_comments_book = False Then
                    CreateScoringBook num_assigned
                Else
                    ' make sure the table on the Instructions sheet is clear
                    Dim rng As String
                    rng = "A" & SCI_FIRST_DATA_ROW & ":" & "C" & SCI_FIRST_DATA_ROW + target_ass_per_marker - 1
                    Sheets(1).Select
                    Range(rng).Select
                    Selection.Clear
                    
                    ' make sure the sheet names in the book don't conflict with the sheetnames for the last marker
                    For i = 2 To Sheets.Count
                        With Sheets(i)
                            .Visible = True
                            .Select
                            .Unprotect lock_sheet_pwd
                            .Name = "sheet_" & (i)
                        End With
                    Next i

                End If
                
                ' now process each assignment
                For i = 1 To num_assigned
                    ' the book starts with an instructions sheet & enough scores and comments sheets
                    ' for the max # of assignements
                    tab_name = marker_assmts(i) & " " & projects_table(marker_assmts(i), 2)
                    marking_sheet = GoodTabName(tab_name)
                    With Sheets(i + 1)
                        .Select
                        'now put the assignment information in the scores/comment sheet
                        .Range(SCT_MARKER_NUM_CELL).Value = mn
                        .Range(SCT_MARKER_NAME_CELL).Value = markers_table(mn, 2)
                        .Range(SCT_PROJECT_NUM_CELL).Value = marker_assmts(i)
                        .Range(SCT_PROJECT_NAME_CELL).Value = projects_table(marker_assmts(i), 2)
                        .Range(SCT_READER_NUM_CELL).Value = marker_rn(i)
On Error GoTo naming_failed
                        .Name = marking_sheet
                    If False Then
naming_failed:
                        If Sheets(i + 1).Name = SCI_INSTRUCTION_SHEET_NAME Then
                            MsgBox "Trying to rename instructions sheet", vbCritical
                        Else
                            MsgBox "Error trying to rename sheet " & Sheets(i + 1).Name, vbCritical
                        End If
                    End If
On Error GoTo 0
                    End With
                    ' make sure the sheet does not have external links
                    If mn = 1 Or (reuse_scores_and_comments_book = False) Then
                        If ConvertMarkerCommentSheetFormulaToText = False Then
                            Exit Function
                        End If
                    End If
                    
                    If simulate_marker_responses Then
                        ' enter random scores and text in the scoresheet
                        If SimulateScoresAndComments = False Then
                            Exit Function
                        End If
                    End If
                     
                    'lock most of the sheet
                    If mn = 1 Or (reuse_scores_and_comments_book = False) Then
                        If LockScoresAndCommentsSheet = False Then
                            Exit Function
                        End If
                    Else
                        Sheets(i + 1).Protect lock_sheet_pwd
                    End If
                    
                    ' add a row to the instructions sheet for this assignment
                    Sheets(SCI_INSTRUCTION_SHEET_NAME).Select
                    link_row(1) = marker_assmts(i)
                    link_row(2) = projects_table(marker_assmts(i), 2)
                    link_row(3) = marking_sheet
                    Set Destination = Range("A" & SCI_PROJECT_COUNT_ROW + 1 + i)
                    Set Destination = Destination.Resize(1, UBound(link_row))
                    Destination.Value = link_row
                    Range("C" & SCI_PROJECT_COUNT_ROW + 1 + i).Select
                    ' hyperlink to the tab name
                    Sheets(SCI_INSTRUCTION_SHEET_NAME).Hyperlinks.Add Anchor:=Selection, Address:="", _
                        SubAddress:="'" & marking_sheet & "'!" & SCT_COI_RESPONSE_CELL, _
                        TextToDisplay:=marking_sheet
                Next i
                If reuse_scores_and_comments_book Then ' since the current implementation makes a new book for each marker
                    'hide any unnecessary sheets (leaving in the blank template)
                    For i = num_assigned + 2 To ActiveWorkbook.Sheets.Count
                        Sheets(i).Visible = False
                    Next i
                End If
                
                ' save the scores and comments workbook
                filestub = DisambiguateFilename(mn & " " & marker_name, mn)
                ActiveWorkbook.SaveAs Filename:= _
                    scores_requested_folder & fps & filestub & scores_and_cmts_ending & ".xlsx", _
                    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                num_scoring_files_created = num_scoring_files_created + 1
                If reuse_scores_and_comments_book = False Then
                    ' makes a new book for each marker
                    ActiveWindow.Close
                End If
            Else
                ' we have one sheet with rows for the scores on each project
                Sheets(1).Unprotect lock_sheet_pwd
                Range(MST_MARKER_NUMBER_CELL).Value = mn
                Range(MST_MARKER_NAME_CELL).Value = marker_name
'                Range(MST_EXPECTED_NUMBER_OF_SCORES).Value = num_criteria * num_assignments
                'ensure the sheet is sized for this marker's number of scoring rows
                Dim num2add_or_delete As Long
                num2add_or_delete = num_assigned - rows_in_table
                InsertAndExpandDown 1, MST_FIRST_SCORING_ROW, 5 + 2 * num_criteria, num2add_or_delete
                
                Dim rn As Long
                rn = MST_FIRST_SCORING_ROW - 1
                ReDim insert_array(1 To num_assigned, 1 To 3)
                For i = 1 To num_assigned
                    insert_array(i, 1) = marker_assmts(i)                      ' project #
                    insert_array(i, 2) = projects_table(marker_assmts(i), 2)   ' project name
                    insert_array(i, 3) = marker_rn(i)                          ' reader #
                Next i
                ' write out the assignments
                Set Destination = Range(c2l(MST_PROJECT_NUM_COL) & MST_FIRST_SCORING_ROW)
                Destination.Resize(UBound(insert_array, 1), UBound(insert_array, 2)).Value = insert_array
                
                ' put the focus on the first project to mark
                Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Select
                Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Activate
                
                ' fill the scoresheets with made up numbers so they can be loaded into the master sheet
                If simulate_marker_responses Then
                    MakeRandomScores num_assigned
                End If
                
                ' lock most of the score sheet so the formulas and layout don't get messed up
                If LockMarkerScoresheet = False Then
                    Exit Function
                End If
                
                'name the sheet after the marker and save the scoring sheet to its own file
                ActiveSheet.Name = GoodTabName(mn & " " & marker_name)
                filestub = DisambiguateFilename(ActiveSheet.Name, mn)
                ActiveWorkbook.SaveAs Filename:= _
                    scores_requested_folder & fps & filestub & scores_only_ending & ".xlsx", _
                    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                num_scoring_files_created = num_scoring_files_created + 1
                ' all done for this one prepare the sheet for the next marker
                rows_in_table = num_assigned
            End If
        End If
    Next mn
    
    If (Not gather_comments) Or reuse_scores_and_comments_book Then
        ' close this working workbook
        ActiveWindow.Close
    End If

    ' hide the sheet(s) in the competition workbook used to as templates
    If gather_comments Then
        sheet_names = Array(SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE, SCORES_AND_COMMENTS_TEMPLATE_SHEET)
        HideOrShowSheets sheet_names, False
    Else
        sheet_names = Array(MARKER_SCORING_TEMPLATE)
        HideOrShowSheets sheet_names, False
    End If
    'all done. put the focus back where it was before the macro ran
    Workbooks(starting_book).Activate
    Sheets(starting_sheet).Activate
    
    If num_markers = num_scoring_files_created Then
        If gather_comments Then
            AddMessage "Created " & num_markers & " files for markers to use for scoring and comments."
        Else
            AddMessage "Created " & num_markers & " files for markers to use for scoring."
        End If
    Else
        AddMessage num_markers & " possible markers, only found " & num_scoring_files_created & " scoring files."
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Erase sheet_names, marker_rn, marker_assmts, insert_array
    
    CreateAllMarkingSheets = True
End Function

Function GetAssignedProjects(mn As Long, ByRef marker_assmts() As Long, ByRef marker_rn() As Long) As Long
    ' go through the assignments array and pull out the assignments for mn, and the reader # (column #)
    Dim i As Long, j As Long, k As Long
    For i = 1 To num_projects
        For j = 1 To target_markers_per_proj
            If assignments(i, j) = mn Then
              k = k + 1
              marker_assmts(k) = i
              marker_rn(k) = j
            End If
        Next j
    Next i
    GetAssignedProjects = k
End Function

Function ConvertMarkerCommentSheetFormulaToText() As Boolean
    
    ConvertRangeToText SCT_COMPETITION_NAME_CELL
    ConvertRangeToText SCT_PROJECT_NAME_CELL
'    ConvertRangeToText SCT_SCORE_CHECK_CELL
    Dim i As Long, name_range As String, min_score_range As String, max_score_range As String
    For i = 1 To num_criteria
        ' criteria name
        name_range = c2l(Range(SCT_CRITERIA_ONE_NAME_CELL).Column) & _
                    (Range(SCT_CRITERIA_ONE_NAME_CELL).row + SCT_ROWS_PER_CRITERIA * (i - 1))
        ConvertRangeToText name_range
        min_score_range = c2l(Range(SCT_CRITERIA_ONE_MIN_CELL).Column) & _
                    (Range(SCT_CRITERIA_ONE_MIN_CELL).row + SCT_ROWS_PER_CRITERIA * (i - 1))
        ConvertRangeToText min_score_range
        max_score_range = c2l(Range(SCT_CRITERIA_ONE_MAX_CELL).Column) & _
                    (Range(SCT_CRITERIA_ONE_MAX_CELL).row + SCT_ROWS_PER_CRITERIA * (i - 1))
        ConvertRangeToText max_score_range
    Next i
    
    ConvertMarkerCommentSheetFormulaToText = True

End Function

Function SimulateScoresAndComments() As Boolean
    
    ' fill the array with random numbers
    Dim i As Long, first_score_col As Long, first_score_row As Long, score As Double
    Dim lb_cell As String, ub_cell As String, lowerbound As Double, upperbound As Double
    Dim score_range As String, comment_range As String
    first_score_col = Range(SCT_FIRST_CRITERIA_SCORE).Column
    first_score_row = Range(SCT_FIRST_CRITERIA_SCORE).row
    Range(SCT_GENERAL_COMMENT_CELL).Value = "Simply dummy text ... Lorem Ipsum."
    For i = 1 To num_criteria
        lb_cell = c2l(C_FIRST_CRITERIA_MINVALUE_CN) & C_FIRST_CRITERIA_MINVALUE_RN + i - 1
        ub_cell = c2l(C_FIRST_CRITERIA_MINVALUE_CN + 1) & C_FIRST_CRITERIA_MINVALUE_RN + i - 1
        lowerbound = Workbooks(cwb).Sheets(CRITERIA_SHEET).Range(lb_cell).Value
        upperbound = Workbooks(cwb).Sheets(CRITERIA_SHEET).Range(ub_cell).Value
        score = (upperbound - lowerbound) * Rnd + lowerbound
        score_range = c2l(first_score_col) & (first_score_row + SCT_ROWS_PER_CRITERIA * (i - 1))
        Range(score_range).Value = score
        comment_range = "A" & (first_score_row + 2 + SCT_ROWS_PER_CRITERIA * (i - 1))
        Range(comment_range).Value = "Simply dummy text ... Lorem Ipsum."
    Next i

    SimulateScoresAndComments = True
End Function

Public Function LockScoresAndCommentsSheet() As Boolean
    ' make score sheet so it can only be modified where we want data
    
    Dim sct_full_sheet_range As String
    Dim i As Long, first_score_col As Long, first_score_row As Long
    first_score_col = Range(SCT_FIRST_CRITERIA_SCORE).Column
    first_score_row = Range(SCT_FIRST_CRITERIA_SCORE).row
    sct_full_sheet_range = "A1:" & c2l(Range(SCT_COI_RESPONSE_CELL).Column) & _
        first_score_row + num_criteria * SCT_ROWS_PER_CRITERIA - 2

    'first lock the whole work area, then unlock fields that should be modified.
    Range(sct_full_sheet_range).Select
    Range(FirstCell(sct_full_sheet_range)).Activate
    Selection.Locked = True
    Selection.FormulaHidden = False
    
    ' now unlock the fields at the top of the sheet
    Range(SCT_COI_RESPONSE_CELL & "," & SCT_CONFIDENCE_LOW_CELL & "," & _
          SCT_CONFIDENCE_MEDIUM_CELL & "," & SCT_CONFIDENCE_HIGH_CELL).Select
    Range(SCT_COI_RESPONSE_CELL).Activate
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' unlock the scoring field and the comments field for each criteria
    Dim unlock_range As String
    unlock_range = SCT_GENERAL_COMMENT_CELL
    For i = 1 To num_criteria
        unlock_range = unlock_range & ","
        unlock_range = unlock_range & _
            c2l(first_score_col) & (first_score_row + SCT_ROWS_PER_CRITERIA * (i - 1))
            unlock_range = unlock_range & ","
        unlock_range = unlock_range & _
                        "A" & ((first_score_row + 2) + SCT_ROWS_PER_CRITERIA * (i - 1)) & _
                        ":K" & ((first_score_row + 2) + SCT_ROWS_PER_CRITERIA * (i - 1))
    Next i
    Range(unlock_range).Select
    Range(FirstCell(unlock_range)).Activate
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' apply this access control by locking the sheet
    Range(SCT_COI_RESPONSE_CELL).Select
    Range(SCT_COI_RESPONSE_CELL).Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=lock_sheet_pwd
    
    LockScoresAndCommentsSheet = True
End Function

Public Function GetMarkerName(marker_number As Long) As String
' get the name associated with a marker number, looking it up from the table of marker numbers and names
    
    
    ' Look up the name of a marker from table on the global assignment sheet
    Dim col_num As Long, row_num As Long, marker_row As Long
    
    ' make sure we are on the right sheet, and get the active cell
    Dim current_sheet As String, cell_ref As String
    current_sheet = ActiveSheet.Name
    
    marker_row = FindMarkerRow(marker_number)
    If marker_row <= 0 Then
        GetMarkerName = ""
        Return
    End If
    
    ' make sure we have found the marker number in the table
    col_num = ActiveCell.Column
    row_num = ActiveCell.row
    If (ActiveCell.Value) <> marker_number Then
        PopMessage "[GetMarkerName] Unable to find marker number " & marker_number & " on assignment sheet " & _
                    MASTER_ASSIGNMENTS_SHEET, vbOKOnly
        GetMarkerName = ""
        Return
    End If
    ' move to the right to get the marker's name
    ChangeActiveCell 0, 1
    GetMarkerName = ActiveCell.Value
    
    ' move back to the starting sheet
    Sheets(current_sheet).Activate
    
End Function

Public Function FindMarkerRow(marker_number As Long) As Long 'returns the row of the active marker
    
    Sheets(MARKERS_SHEET).Activate
    ' search in the column of marker numbers for this marker number
    Columns(M_NUMBER_AND_NAME_COLUMNS).Select
    Range("A1").Activate
    If Selection.Find(What:=marker_number, after:=ActiveCell, LookIn:=xlFormulas2, LookAt _
                        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                        False, SearchFormat:=False) Is Nothing Then
        PopMessage "unable to find row for marker " & marker_number, vbCritical
        FindMarkerRow = -1
        Exit Function
    End If
    Selection.Find(What:=marker_number, after:=ActiveCell, LookIn:=xlFormulas2, LookAt _
                        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                        False, SearchFormat:=False).Activate
    FindMarkerRow = ActiveCell.row
    
End Function

Public Function LockMarkerScoresheet() As Boolean
    
    ' lock the cells that should not be edited in the scoresheet that the marker will use to score with
    ' these are generally grey-filled
    Const MST_LOCKING_RANGE As String = "A2:C28,D1:O8,I9:O29,D29:H29,G30:I32,J31:M31"
    Range(MST_LOCKING_RANGE).Select
    ' select the first cell named in the range
    Range(FirstCell(MST_LOCKING_RANGE)).Activate
    Selection.FormulaHidden = False
    Selection.Locked = True
    
    ' make sure the scoring entries are not locked.
    Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW & ":" & _
          c2l(MST_FIRST_SCORING_COL + num_criteria - 1) & _
          (MST_FIRST_SCORING_ROW + max_ass_per_marker - 1)).Select
    Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Activate
    Selection.Locked = False
    
    'now lock the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=lock_sheet_pwd
    
    LockMarkerScoresheet = True
    
End Function

Public Function MakeRandomScores(num_assignments As Long) As Boolean
' fill in the marker scoresheet with random numbers
    
    ' create the random numbers
    Dim upperbound As Long, lowerbound As Long
    Dim i As Long, j As Long, lb_cell As String, ub_cell As String
    Dim random_scores() As Double
    ReDim random_scores(1 To num_assignments, 1 To num_criteria)
    
    ' fill the array with random numbers
    For j = 1 To num_criteria
        lb_cell = c2l(C_FIRST_CRITERIA_MINVALUE_CN) & C_FIRST_CRITERIA_MINVALUE_RN + j - 1
        ub_cell = c2l(C_FIRST_CRITERIA_MINVALUE_CN + 1) & C_FIRST_CRITERIA_MINVALUE_RN + j - 1
        lowerbound = Workbooks(main_workbook).Sheets(CRITERIA_SHEET).Range(lb_cell).Value
        upperbound = Workbooks(main_workbook).Sheets(CRITERIA_SHEET).Range(ub_cell).Value
        For i = 1 To num_assignments
            random_scores(i, j) = (upperbound - lowerbound) * Rnd + lowerbound
        Next i
    Next j
    
    ' now write the array to the expertise column of the expertise sheet
    Dim Destination As Range
    Set Destination = Range(c2l(ActiveCell.Column) & ActiveCell.row)
    Destination.Resize(num_assignments, num_criteria) = random_scores
'    Destination.Value = Application.Transpose(random_scores)
        
    Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Select
    Erase random_scores
    MakeRandomScores = False
    
End Function

Public Function LoadScoresAndComments() As Boolean
'   start with a master scoresheet template and populate it from the scoresheets found in a folder

    Dim starting_sheet As String, starting_workbook As String, read_range As String
    starting_sheet = ActiveSheet.Name
    starting_workbook = ActiveWorkbook.Name

    If ActivateCompetitionWorkbook = False Then
        LoadScoresAndComments = False
        Exit Function
    End If
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    If False Then   ' we don't need to read this in as it will come from the score sheets/books
        ' read in the marker and project columns from the shared scoresheet (so we know where to store scores)
        Sheets(RESULTS_SHEET).Select
            read_range = c2l(R_PROJECT_NUM_COLUMN) & R_FIRST_DATA_ROW & ":" & _
                     c2l(R_PROJECT_NUM_COLUMN) & R_FIRST_DATA_ROW + num_projects * target_markers_per_proj - 1
        ss_project_col = Range(read_range)
        read_range = c2l(R_MARKER_NUM_COLUMN) & R_FIRST_DATA_ROW & ":" & _
                     c2l(R_MARKER_NUM_COLUMN) & R_FIRST_DATA_ROW + num_projects * target_markers_per_proj - 1
        ss_marker_col = Range(read_range)
    Else
        ' allocate space for the data to be compiled
        ReDim proj_num_col(1 To num_projects * target_markers_per_proj, 1 To 1)
        ReDim reader_num_col(1 To num_projects * target_markers_per_proj, 1 To 1)
        ReDim marker_num_col(1 To num_projects * target_markers_per_proj, 1 To 1)
        ReDim competition_scores(1 To num_projects * target_markers_per_proj, 1 To num_criteria)
    End If
    num_competition_scores = 0
    
    ' Ask the user to specify the folder containing the marker score files
    Dim folder_with_scores As String
    If Len(root_folder) > 0 Then
        ChDir root_folder
    End If
    folder_with_scores = SelectFolder("Select folder containing the scoresheets to compile", _
                        root_folder & scores_received_folder)
    If Len(folder_with_scores) = 0 Then
        LoadScoresAndComments = False
        Exit Function
    End If

    'load and process the xlsx files that have the right filename pattern to contain scores
    Dim looking As Boolean, marking_sheet As String, marker_num As Long, marker_name As String
    Dim file_name As String, num_ms As Long, num_scores As Long, project_num As Long
    Dim i As Long, j As Long, output_row_num As Long
    Dim num_scores_missing As Long
    Dim file_path As String
    Dim assignment_col As Long
    Dim file_pattern As String, score_read As String
    ' make space for the comments
    If gather_comments Then
        ReDim comments(1 To num_projects, 1 To num_criteria)
        ReDim general_comments(1 To num_projects)
    End If
    
    ' process the files
    num_ms = 0          ' counter of the number of marking sheets
    looking = True
    output_row_num = 0
    While looking
        If num_ms = 0 Then
            If gather_comments Then
                file_name = Dir(folder_with_scores & fps & marks_and_cmts_file_pattern)
            Else
                file_name = Dir(folder_with_scores & fps & marks_only_file_pattern)
            End If
        Else
            file_name = Dir()
        End If
        If Len(file_name) = 0 Then
            looking = False     'no more files to process
        Else
            file_path = folder_with_scores & fps & file_name
            If gather_comments Then
                marker_num = ReadCommentsAndScores(file_path, output_row_num)
            Else
                marker_num = ReadScoresFromSingleSheet(file_path)
            End If
            
            If marker_num <= 0 Then
                AddMessage "LoadScoresAndComments: Error reading scores from " & file_name
                LoadScoresAndComments = False
                Exit Function
            Else
                num_ms = num_ms + 1
            End If
        End If
    Wend
    
    ' check that scores were loaded
    If num_ms = 0 Then
        PopMessage "No scores were loaded - check the folder for marking files", vbCritical
        LoadScoresAndComments = False
        Exit Function
    End If
    
    ' load the table of scores etc into the Results sheet
    If AddScoresToResultsSheet(file_path, marker_num) = False Then
        LoadScoresAndComments = False
        Exit Function
    End If
    If PopulateAnalysisSheet(True) = False Then
        LoadScoresAndComments = False
        Exit Function
    End If
            
    ' output the evaluators comments if they were loaded
    Dim with_without_comments As String
    If gather_comments Then
        If num_ms > 0 Then
            If OutputComments = False Then
                Exit Function
            End If
        End If
        with_without_comments = " with comments"
    Else
        with_without_comments = "without comments"
    End If
    
    Dim raw_or_normalized As String
    If normalize_scoring Then
        raw_or_normalized = "normalized"
    Else
        raw_or_normalized = "raw"
    End If
    AddMessage "Compiled " & raw_or_normalized & " scores " & with_without_comments & " from " & num_ms & " markers."
    
    ' put the focus back where it was when the macro started
    Workbooks(starting_workbook).Activate
    Sheets(starting_sheet).Activate
    
    LoadScoresAndComments = True
End Function

Function OutputComments() As Boolean
    'save a file for each projects comments
    Dim i As Long, j As Long, rn As Long
    Dim file_name As String
    Dim sheet_name As String
    
    ' load the projects table
    Workbooks(cwb).Activate
    Sheets(PROJECTS_SHEET).Select
    Dim table_range As String, sheet_names() As Variant
    table_range = c2l(1) & P_FIRST_DATA_ROW & ":" & c2l(P_MENTOR_ID_COLUMN) & (P_FIRST_DATA_ROW + num_projects - 1)
    projects_table = Range(table_range)

    ' copy the comments template sheet to a new document
    sheet_names = Array(PROJECT_COMMENTS_SHEET)
    HideOrShowSheets sheet_names, True
    Sheets(sheet_names).Select
    Sheets(sheet_names).Copy
    
    ' Ask the user to specify the folder to save the comments files
    Dim comments_destination_folder As String
    If Len(root_folder) > 0 Then
        ChDir root_folder
    End If
    comments_destination_folder = SelectFolder("Select the destination folder for comments", _
                        root_folder & comments_folder)
    If Len(comments_destination_folder) = 0 Then
        OutputComments = False
        Exit Function
    End If
    
    ' we will repopulate it for each comment sheet
    For i = 1 To num_projects
        'populate it with the available comments
        With ActiveSheet
            .Range(PC_PROJECT_NUM_CELL).Value = projects_table(i, 1)
            .Range(PC_PROJECT_NAME_CELL).Value = projects_table(i, 2)
            .Range(PC_GENERAL_COMMENTS_CELL).Value = general_comments(i)
            rn = PC_FIRST_CRITERIA_COMMENTS_ROW
            For j = 1 To num_criteria
                .Range("A" & rn).Value = comments(i, j)
                rn = rn + PC_ROWS_PER_CRITERIA
            Next j
            ' make sure the comments are visible
            
            ' save it to a format accessible by a word-processor
            Dim filestub As String, ending As String, saveas_type As Long
            
            Select Case output_comments_format
            Case "PRINTER"
                saveas_type = xlTextPrinter
                ending = ".prn"
            Case "HTML"
                saveas_type = xlHtml
                ending = ".htm"
            Case "XLSX"
                saveas_type = xlOpenXMLWorkbook
                ending = ".xlsx"
            Case "TEXT"
                saveas_type = xlTextWindows
                ending = ".txt"
            Case Else
                PopMessage "unknown file output format " & output_comments_format, vbCritical
                Exit Function
                
            End Select
            filestub = GoodTabName(i & " " & ActiveSheet.Range(PC_PROJECT_NAME_CELL).Value)
            file_name = comments_destination_folder & fps & filestub & ending
            ActiveWorkbook.SaveAs Filename:=file_name _
             , FileFormat:=saveas_type, ReadOnlyRecommended:=False, CreateBackup:=False
        End With
    Next i
    
    ' delete the template workbook
    ActiveWorkbook.Close
    ' hide the template
    Workbooks(cwb).Activate
    HideOrShowSheets sheet_names, False
    Sheets(ANALYSIS_CHART_NAME).Select
    Sheets(ANALYSIS_CHART_NAME).Activate

    OutputComments = True
End Function
    
Function SortFinalScoresheets(mss_workbook As String, master_sheet As String) As Boolean
    'sort both versions of the final score sheets by decreasing scores
    
    Workbooks(mss_workbook).Activate
    Sheets(RESULTS_SHEET).Activate
    Dim sort_range As String, sort_key As String
    sort_key = c2l(R_FINAL_TOTAL_SCORES_COLUMN) & R_FIRST_DATA_ROW & ":" & _
               c2l(R_FINAL_TOTAL_SCORES_COLUMN) & (R_FIRST_DATA_ROW + num_projects - 1)
    sort_range = c2l(R_FINAL_PROJ_COLUMN) & R_FIRST_DATA_ROW & ":" & _
               c2l(R_FINAL_TOTAL_SCORES_COLUMN) & (R_FIRST_DATA_ROW + num_projects - 1)
    Range(sort_range).Select
    Range(FirstCell(sort_key)).Activate
    ActiveWorkbook.Worksheets(RESULTS_SHEET).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(RESULTS_SHEET).Sort.SortFields.Add2 Key:= _
        Range(sort_key), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(RESULTS_SHEET).Sort
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets(master_sheet).Activate
    sort_key = c2l(MSS_TOTAL_SCORES_COL) & MSS_FIRST_PROJECT_ROW & ":" & _
               c2l(MSS_TOTAL_SCORES_COL) & (MSS_FIRST_PROJECT_ROW + num_projects - 1)
    sort_range = "A" & MSS_FIRST_PROJECT_ROW & ":" & _
               c2l(MSS_LAST_COL + 2) & (MSS_FIRST_PROJECT_ROW + num_projects - 1)
    Range(sort_range).Select
    Range(FirstCell(sort_key)).Activate
    ActiveWorkbook.Worksheets(master_sheet).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(master_sheet).Sort.SortFields.Add2 Key:= _
        Range(sort_key), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(master_sheet).Sort
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    SortFinalScoresheets = True
End Function
    
Function LabelMasterScoresheetHeaders() As Boolean
    ' put in the correct header (normalized or not)
    msbox "LabelMasterScoresheetHeaders needs updating", vbCritical
    LabelMasterScoresheetHeaders = False
    Exit Function
    
    Dim stub As String, i As Long
    Dim title_range As String
    If normalize_scoring Then
        stub = "normalized"
    Else
        stub = "raw"
    End If
    stub = stub & ") criteria scores"
    For i = 1 To max_markers_per_proj
        title_range = c2l(MSS_FIRST_SCORE_COL + 1 + (i - 1) * (max_criteria + 2)) & MSS_MARKER_NUMBER_ROW
        Range(title_range).Value = "Marker #" & i & " (" & stub
    Next i
    LabelMasterScoresheetHeaders = True
End Function

Function AddScoresToResultsSheet(file_path As String, marker_num As Long) As Boolean
    'this function assumes the shared scoresheet is active
    
    ' load the markers and projects
    LoadMarkerAndProjectTables
    ' get the information for the column of Expertise confidences
    Dim coa_array_range As String
    'target_markers_per_proj' columns wide by num_projects rows
    coa_array_range = c2l(EC_ASSMT_CONF_FIRST_COL) & EC_FIRST_DATA_ROW & ":" & _
                c2l(EC_ASSMT_CONF_FIRST_COL + target_markers_per_proj - 1) & (EC_FIRST_DATA_ROW + num_projects - 1)
    Sheets(EXPERTISE_CROSSWALK_SHEET).Select
'    Sheets(EXPERTISE_CROSSWALK_SHEET).Activate
    coa_array = Range(coa_array_range)
    
    ' update columns for reviewer confidence levels, marker names and project names
    Dim i As Long
    Dim marker_names() As Variant, project_names() As Variant, confidence_levels() As Variant
    ReDim marker_names(1 To UBound(marker_num_col, 1))
    ReDim project_names(1 To UBound(marker_num_col, 1))
    ReDim confidence_levels(1 To UBound(marker_num_col, 1))
    For i = 1 To UBound(marker_num_col, 1)
        If IsEmpty(marker_num_col(i, 1)) = False Then
            ' there was data in the Results sheet for this row, so update it.
            marker_names(i) = markers_table(marker_num_col(i, 1), 2)
            project_names(i) = projects_table(proj_num_col(i, 1), 2)
            confidence_levels(i) = num2LMH(coa_array(proj_num_col(i, 1), reader_num_col(i, 1)))
        End If
    Next i

    ' now we are read to write out the columns
    Dim write_range As String, dest As Range
    ' first the marker numbers column
    Sheets(RESULTS_SHEET).Select
    write_range = c2l(R_MARKER_NUM_COLUMN) & R_FIRST_DATA_ROW
    Set dest = Range(write_range)
    dest.Resize(UBound(marker_num_col, 1), UBound(marker_num_col, 2)) = marker_num_col
    ' next to it the corresponding marker names
    write_range = c2l(R_MARKER_NUM_COLUMN + 1) & R_FIRST_DATA_ROW
    Set dest = Range(write_range)
    dest.Resize(UBound(marker_names, 1)) = Application.Transpose(marker_names)
    
    ' the reader #
    write_range = c2l(R_READER_NUM_COLUMN) & R_FIRST_DATA_ROW
    Set dest = Range(write_range)
    dest.Resize(UBound(reader_num_col, 1), UBound(reader_num_col, 2)) = reader_num_col
   
    ' the reviewer's confidence on this project
    write_range = c2l(R_READER_NUM_COLUMN + 1) & R_FIRST_DATA_ROW
    Set dest = Range(write_range)
    dest.Resize(UBound(confidence_levels, 1)) = Application.Transpose(confidence_levels)
    
    ' The Project numbers
    write_range = c2l(R_PROJECT_NUM_COLUMN) & R_FIRST_DATA_ROW
    Set dest = Range(write_range)
    dest.Resize(UBound(proj_num_col, 1), UBound(proj_num_col, 2)) = proj_num_col
    ' next to it the corresponding marker names
    write_range = c2l(R_PROJECT_NUM_COLUMN + 1) & R_FIRST_DATA_ROW
    Set dest = Range(write_range)
    dest.Resize(UBound(project_names, 1)) = Application.Transpose(project_names)
    
    ' The scores
    write_range = c2l(R_FIRST_RAW_COLUMN) & R_FIRST_DATA_ROW
    Set dest = Range(write_range)
    dest.Resize(UBound(competition_scores, 1), UBound(competition_scores, 2)) = competition_scores
    
    SortResultRawScoresTable UBound(competition_scores, 1)
    SortResultsFinalProjectTable
    
    Erase marker_num_col, reader_num_col, proj_num_col, competition_scores
    AddScoresToResultsSheet = True
    
End Function

Function SortResultsFinalProjectTable() As Boolean
    Dim sort_range As String, key1_range As String, key2_range As String
  ' sort the final table by project score
    key1_range = c2l(R_T4S + R_T4N - 2 + 3 * (num_criteria - 2)) & R_FIRST_DATA_ROW & ":" & _
                 c2l(R_T4S + R_T4N - 2 + 3 * (num_criteria - 2)) & (R_FIRST_DATA_ROW + num_projects - 1) 'first sort by marker
    sort_range = c2l(R_T4S + 2 * (num_criteria - 2)) & R_FIRST_DATA_ROW & ":" & _
                 c2l(R_T4S + R_T4N - 2 + 3 * (num_criteria - 2)) & (R_FIRST_DATA_ROW + num_projects - 1)
    With ActiveWorkbook.Worksheets(RESULTS_SHEET).Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range(key1_range), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    SortResultsFinalProjectTable = True
End Function
Function SortResultRawScoresTable(row_num As Long) As Boolean
    ' sort the raw scores by increasing marker, and then ascending project numbers

    Dim sort_range As String, key1_range As String, key2_range As String
    key1_range = c2l(R_MARKER_NUM_COLUMN) & R_FIRST_DATA_ROW & ":" & c2l(R_MARKER_NUM_COLUMN) & _
                                        (R_FIRST_DATA_ROW + row_num - 1) 'first sort by marker
    key2_range = c2l(R_PROJECT_NUM_COLUMN) & R_FIRST_DATA_ROW & ":" & c2l(R_PROJECT_NUM_COLUMN) & _
                                        (R_FIRST_DATA_ROW + row_num - 1) ' then sort by project
    sort_range = c2l(R_T1S) & R_FIRST_DATA_ROW & ":" & _
              c2l(R_T1S + R_T1N - 1) & (R_FIRST_DATA_ROW + row_num - 1)
    With ActiveWorkbook.Worksheets(RESULTS_SHEET).Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range(key1_range), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range(key2_range), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    SortResultRawScoresTable = True
End Function

Private Function ReadCommentsAndScores(file_path As String, output_row_num As Long) As Long
    'open the file with a scoresheet
    Dim marker_num As Long, num_scores As Long, i As Long, j As Long, num_assigned_as_long
    Dim num_assigned As Long, read_row As Long, project_nums() As Variant
    Dim tab_names() As String, comment As Variant
    Workbooks.Open Filename:=file_path
    
    ' load in the structure information for the book from the table on the instructions sheet
    num_assigned = Sheets(SCI_INSTRUCTION_SHEET_NAME).Range("B" & SCI_PROJECT_COUNT_ROW).Value
    ReDim tab_names(1 To num_assigned)
    ReDim project_nums(1 To num_assigned)
    For i = 1 To num_assigned
        project_nums(i) = Sheets(SCI_INSTRUCTION_SHEET_NAME).Range("A" & (SCI_PROJECT_COUNT_ROW + 1 + i)).Value
        tab_names(i) = Sheets(SCI_INSTRUCTION_SHEET_NAME).Range("C" & (SCI_PROJECT_COUNT_ROW + 1 + i)).Value
    Next i

    ' go through each of the marks/comments sheets and extract the scores and comments info
    Dim reader_num As Long, COI_response As String, row_num As Long
    For i = 1 To num_assigned
        output_row_num = output_row_num + 1
        If output_row_num > num_projects * target_markers_per_proj Then
            PopMessage "ReadCommentsAndScores Error: # of rows exceeds space for scores", vbCritical
            ReadCommentsAndScores = -1
            Exit Function
        End If
        With Sheets(tab_names(i))
            COI_response = .Range(SCT_COI_RESPONSE_CELL).Value
            If InStr(COI_response, "N") > 0 Then
                AddMessage "ReadCommentsAndScores: marker " & marker_num_col(output_row_num, 1) & _
                            " flags a conflict of interest for project " & project_nums(i) & _
                            ". Their input will be ignored for this project."
                output_row_num = output_row_num - 1
            Else
                read_row = 2
                proj_num_col(output_row_num, 1) = project_nums(i)
                marker_num = .Range(SCT_MARKER_NUM_CELL).Value
                marker_num_col(output_row_num, 1) = marker_num
                reader_num_col(output_row_num, 1) = .Range(SCT_READER_NUM_CELL).Value
                AppendComment CLng(reader_num_col(output_row_num, 1)), .Range(SCT_GENERAL_COMMENT_CELL).Value, _
                                    general_comments(project_nums(i))
                read_row = Range(SCT_FIRST_CRITERIA_SCORE).row
                For j = 1 To num_criteria
                    competition_scores(output_row_num, j) = .Range(c2l(SCT_SCORE_COLUMN) & read_row).Value
                    ' get the comment (if provided and append it to the comments for that criteria
                    AppendComment CLng(reader_num_col(output_row_num, 1)), .Range("A" & (read_row + 2)).Value, _
                                    comments(project_nums(i), j)
                    read_row = read_row + SCT_ROWS_PER_CRITERIA
                Next j
            End If
        End With
    Next i
    Erase tab_names, project_nums
    ActiveWorkbook.Close
    ReadCommentsAndScores = marker_num
    
 End Function

Private Function AppendComment(reader_num As Long, comment_read As String, comments As Variant) As Boolean
    If Len(CStr(comments)) > 1 Then
        comments = comments & vbCrLf
    End If
    comments = comments & "Comment from reader #" & reader_num & ":" & vbCrLf & comment_read
    AppendComment = True
End Function

Private Function ReadScoresFromSingleSheet(file_path As String) As Long
    ' returns marker number who did the scoring
    
    'open the file with a scoresheet
    Dim marker_num As Long, num_projects As Long
    Workbooks.Open Filename:=file_path
    ' get name of scoring sheet (should be the only sheet)
    If Sheets.Count > 1 Then
        PopMessage "[ReadScoresFromSingleSheet] Expected only one sheet in book, found " & Sheets.Count, vbCritical
        Exit Function
    End If
    marker_num = ActiveSheet.Range(MST_MARKER_NUMBER_CELL).Value
    
    ' Extract the scores and associated data from the sheet
    Dim num_projects_range As String
    num_projects_range = c2l(MST_FIRST_SCORING_COL + 2 * num_criteria + 1) & 1
    num_projects = Range(num_projects_range).Value
    Dim first_score_column As Long
    Dim pn_1m_range As String, rn_1m_range As String, scores_range As String
    
    'load the project numbers assigned to this marker
    pn_1m_range = c2l(MST_PROJECT_NUM_COL) & MST_FIRST_SCORING_ROW & ":" & _
                  c2l(MST_PROJECT_NUM_COL) & MST_FIRST_SCORING_ROW + (num_projects - 1)
    pn_1m = Range(pn_1m_range)
    'load the reader numbers for each assignment
    rn_1m_range = c2l(MST_READER_NUM_COL) & MST_FIRST_SCORING_ROW & ":" & _
                  c2l(MST_READER_NUM_COL) & MST_FIRST_SCORING_ROW + (num_projects - 1)
    rn_1m = Range(rn_1m_range)
    If False Then   'change in approach - always read the raw scores, and leave to the competition to
                    ' decide whether to normalize (was if normalize_scoring then)
        first_score_column = MST_FIRST_SCORING_COL + num_criteria + 1
    Else
        first_score_column = MST_FIRST_SCORING_COL
    End If
    scores_range = c2l(first_score_column) & MST_FIRST_SCORING_ROW & ":" & _
                    c2l(first_score_column + num_criteria - 1) & (MST_FIRST_SCORING_ROW + num_projects - 1)
    marker_scores = Range(scores_range)
    
    AddMarkerScores marker_num, pn_1m, rn_1m, marker_scores
    Erase pn_1m, rn_1m, marker_scores
    
    ActiveWorkbook.Close
    
    ReadScoresFromSingleSheet = marker_num
    
End Function

Function AddMarkerScores(marker_num As Long, ByRef pn_1m(), ByRef rn_1m(), ByRef scores_read() As Variant) As Boolean
    ' add the data read from a scoresheet/scorebook to the data columns to be written to the Results sheet
    Dim old_num_competition_scores As Long, i As Long, j As Long, k As Long
    old_num_competition_scores = num_competition_scores
    num_competition_scores = num_competition_scores + UBound(pn_1m, 1)
    j = 0
    For i = old_num_competition_scores + 1 To num_competition_scores
        j = j + 1
        proj_num_col(i, 1) = pn_1m(j, 1)
        reader_num_col(i, 1) = rn_1m(j, 1)
        marker_num_col(i, 1) = marker_num
        For k = 1 To num_criteria
            competition_scores(i, k) = marker_scores(j, k)
        Next k
    Next i
End Function

Function AddScoresToMasterSheet(file_path As String, marker_num As Long) As Boolean
    ' this function assumes the master scoresheet is selected
    
    ' for loading the scores into the master sheet
    Dim num_scores As Long, num_scores_missing, i As Long, j As Long
    num_scores = UBound(scores, 1)
    Dim scores_row() As Double
    ReDim scores_row(1 To num_criteria)
    For i = 1 To num_scores
        num_scores_missing = 0
        For j = 1 To num_criteria
            If Len(scores(i, j)) > 0 Then
                scores_row(j) = CDbl(scores(i, j))
            Else
                num_scores_missing = num_scores_missing + 1
            End If
        Next j
        If num_scores_missing > 0 Then
            AddMessage "Project " & pn_1m(i, 1) & " in file <" & file_path & "> is missing " & _
                        num_scores_missing & " score(s)."
        Else
            'find the project that corresponds to this row of scores
            Range(c2l(MSS_PROJECT_COL) & MSS_FIRST_PROJECT_ROW & ":" & _
                    c2l(MSS_PROJECT_COL) & MSS_FIRST_PROJECT_ROW + num_projects - 1).Select
            ' was xlformulas2
            Selection.Find(What:=pn_1m(i, 1), after:=ActiveCell, LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False).Activate
            'insert the scores
            If (InsertScores(CLng(rn_1m(i, 1)), marker_num, scores_row) = False) Then
                Exit Function
            End If
        End If
    Next i
    Erase scores_row
    AddScoresToMasterSheet = True
    
End Function

Public Function ActivateCompetitionWorkbook() As Boolean
    Dim book_name As String, i As Long
    If Len(cwb) > 0 Then
        For i = 1 To Workbooks.Count
            If Workbooks(i).Name = cwb Then
                Workbooks(cwb).Activate
                ActivateCompetitionWorkbook = True
                Exit Function
            End If
        Next i
    End If
    book_name = ActivateWorkbookBySheetname(COMPETITION_PARAMETERS_SHEET)
    cwb = book_name
    If Len(book_name) > 0 Then
        ActivateCompetitionWorkbook = True
    End If
    
End Function
Function ActivateWorkbookBySheetname(sheet_name As String) As String ' returns the name of the workbook
    ' activate or open the workbook containing a particular sheet

    'loop through the open workbooks looking for an XLSX file containing a sheet
    ' whose name is given
    Dim i As Long, j As Long
    Dim found As Boolean
    For i = 1 To Workbooks.Count
        With Workbooks(i)
            If Right(.FullName, 4) = "xlsx" Then
                For j = 1 To .Sheets.Count
                    If .Sheets(j).Name = sheet_name Then
                        found = True
                        .Activate
                        .Sheets(j).Activate
                        ActivateWorkbookBySheetname = ActiveWorkbook.Name
                        Exit Function
                    End If
                Next j
            End If
        End With
    Next i
    
    ' the one we need is not among the active workbooks, ask the user to select a file

    Dim looking As Boolean
    looking = True
    While looking
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Dim file_name As String
        file_name = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsx*), *.xlsx", _
                                title:="Choose the Excel file to update its master scoresheet", MultiSelect:=False)
        If (Len(file_name) = 0) Or (file_name = "False") Then
            Exit Function
        End If
        
        ' open the selected file
        Workbooks.Open Filename:=file_name
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        ' check that it contains the desired sheet
        With ActiveWorkbook
            For j = 1 To .Sheets.Count
                If .Sheets(j).Name = sheet_name Then
                    looking = False
                    ActivateWorkbookBySheetname = ActiveWorkbook.Name
                    Exit Function
                End If
            Next j
        End With
        ' sheet not found, let the user know
        Dim msg As String
        msg = " workbook " & file_name & " does not contain a (required) sheet named " & sheet_name & _
                ". Please select a file that does"
        If PopMessage(msg, vbOKCancel) = vbCancel Then
            ActivateWorkbookBySheetname = ""
            Exit Function
        End If
    Wend
    
End Function

Function InsertScores(assignment_col As Long, marker_num As Long, scores_row() As Double) As Boolean
    ' insert the three (normalized) scores in the master sheet, and the marker who made them,
    ' assuming we are on the correct row.
    If marker_num < 1 Or marker_num > num_markers Then
        MsgBox "[InsertScores] unexpected marker_num: " & marker_num, vbCritical
        InsertScores = False
        Exit Function
    End If
    Dim start_col As Long, i As Long
    start_col = MSS_FIRST_SCORE_COL + (max_criteria + 2) * (assignment_col - 1)
    Range(c2l(start_col - 1) & ActiveCell.row).Value = marker_num
    For i = 1 To UBound(scores_row)
        Range(c2l(start_col + i - 1) & ActiveCell.row).Value = scores_row(i)
    Next i
    InsertScores = True
    
End Function


Public Function Foo()
'
'   replace the vlookups for the marker # columns, and the marker name lookups
'   with their text equivalents, so that the MSS does not have links to other workbooks
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "B", num_projects
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "C", num_projects
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "J", num_projects
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "Q", num_projects
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "X", num_projects
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "AE", num_projects
    
    ' convert the vlookup for the criteria and scoring ranges to text
    ConvertRangeToText "D4:H4"

End Function

Public Function getMarkingSheetName(file_name As String, name_suffix As String, _
                                    marker_num As Long, marker_name As String) _
                                    As Long
' file_name = "18 Sigurer.xlsx"
' sheet_name = "18 Sigurer"
' marker_num = 18
' Marker name = "Sigurer"  (not the exact maker's name as accents and punctuation were removed)

    Dim i As Long, sheet_name As String
    Dim number_from_sheet_name As String, number_from_file_name As String
    number_from_file_name = Left(file_name, InStr(file_name, " ") - 1)      ' file_name leads with a number
    For i = 1 To ActiveWorkbook.Sheets.Count        ' make sure the book name starts with a number
        sheet_name = Sheets(i).Name
        If InStr(Sheets(i).Name, " ") > 0 Then      ' and has to have a space in the sheet name
            number_from_sheet_name = Left(Sheets(i).Name, InStr(Sheets(i).Name, " ") - 1)
            If number_from_file_name = number_from_sheet_name Then
                marker_num = Val(number_from_file_name)
                sheet_name = Left(file_name, Len(file_name) - 5) ' strip off the extent
                marker_name = Right(file_name, Len(file_name) - InStr(file_name, " "))   ' get string without the number
                marker_name = Left(marker_name, InStr(marker_name, name_suffix) - 1)    ' remove the suffix to get the name
                getMarkingSheetName = i
                Exit Function
            End If
        End If
    Next i
    MsgBox "ERROR getting marker number and name from file " & file_name, vbCritical
    
End Function

Public Function LoadMarkerProjectExpertiseIntoPXM() As Boolean
    ' load the data from all the project expertise sheets in a folder into the PXM table
    
    If ActivateCompetitionWorkbook = False Then
        LoadMarkerProjectExpertiseIntoPXM = False
        Exit Function
    End If
    making_competition_workbook = False
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' initialize the PXM sheet
    Dim PXM_workbook As String
    PXM_workbook = ActiveWorkbook.Name
    
    ' get the folder name with the expertise files to load
    Dim folder_with_expertises As String
'    ChDir root_folder
    folder_with_expertises = SelectFolder("Select folder containing the expertise about projects of potential markers", _
                root_folder & expertise_by_project_received_folder)
    If Len(folder_with_expertises) = 0 Then
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'load and process the xlsx files that have the right filename pattern.
    Dim expertise_sheet As String, marker_num As Long, marker_name As String
    Dim file_name As String, num_ms As Long, project_num As Long
    Dim rn As Long, i As Long, j As Long, num_es As Long
    num_es = 0                      ' counter of the number of expertise sheets
    Dim expertise() As Variant      ' arrays read from the expertise sheet
    Dim COIs() As Variant           ' the contents will be strings, but use variants to get from the S/S
    Dim finding As Boolean
    Dim expertise_out() As Variant
    ReDim expertise_out(1 To num_projects)
    Dim expertise_workbook As String
    ReDim pxm_table(1 To num_projects, 1 To num_markers)
    
    Sheets(PROJECT_X_MARKER_SHEET).Select
    
    ' ignore any information currently stored in the PxM table
    ClearPXMSheet
        
    finding = True
    Dim num_blank As Long, num_excluded As Long
    num_blank = 0
    num_excluded = 0
    While finding
        If num_es = 0 Then
            file_name = Dir(folder_with_expertises & fps & project_expertise_file_pattern)
        Else
            file_name = Dir()
        End If
        If Len(file_name) = 0 Then
            finding = False     'no more files to process
        Else        'open the file with a expertise sheet
            ' get name of expertise sheet from filename
            Workbooks.Open Filename:=folder_with_expertises & fps & file_name
            expertise_workbook = ActiveWorkbook.Name
            Dim sht_num As Long
            sht_num = getMarkingSheetName(file_name, expertise_by_project_ending, marker_num, marker_name)
            expertise_sheet = Sheets(sht_num).Name
  'PROJECT VERSION
            ' load the keyword expertise column from that sheet
            COIs = Range(c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                         c2l(MPET_COI_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1))
            
            ' load the project COI column from the COI sheet
            expertise = Range(c2l(MPET_EXPERTISE_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                                c2l(MPET_EXPERTISE_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1))
  
            ' close the expertise workbook
            Workbooks(expertise_workbook).Close
            'combine the expertise info provided with the COIs signalled
            For i = 1 To num_projects
                Select Case COIs(i, 1)
                Case "X", "Y", "Mentor", "MENTOR", same_organization_text
                    pxm_table(i, marker_num) = "X" '
                    num_excluded = num_excluded + 1
                Case "N", ""
                    If (expertise(i, 1) = "") And blank_expertise_means_exclusion Then
                            pxm_table(i, marker_num) = ""
                            num_blank = num_blank + 1
                        Else
                            If expertise(i, 1) = "" Then
                                pxm_table(i, marker_num) = LMH2Percent("L")
                            Else
                                pxm_table(i, marker_num) = LMH2Percent(CStr(expertise(i, 1)))
                            End If
                        End If
                Case Else
                    MsgBox "[LoadMarkerProjectExpertiseIntoPXM] Marker " & marker_name & _
                            "'s file has unexpected COI for project " & i & ", value: " & COIs(i, 1) & ", ignored.", vbOKOnly
                    Exit Function
                End Select
            Next i
            num_es = num_es + 1
        End If
    Wend
    
    'write the pxm_table to the pxm table sheet
    Sheets(PROJECT_X_MARKER_SHEET).Select
    Dim Destination As Range, pxm_range As String
    pxm_range = c2l(PXM_FIRST_DATA_COL) & PXM_FIRST_DATA_ROW
    Set Destination = Range(pxm_range)
    Destination.Resize(UBound(pxm_table, 1), UBound(pxm_table, 2)).Value = pxm_table
    
    ' formatting for the table
    Selection.NumberFormat = "0.0"
    Range(c2l(ec_data_first_marker_column) & 1).Activate
    Columns(c2l(PXM_FIRST_DATA_COL) & ":" & _
            c2l(PXM_FIRST_DATA_COL - 1 + num_markers)).EntireColumn.AutoFit

    Range(c2l(PXM_FIRST_DATA_COL) & (PXM_FIRST_DATA_ROW)).Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Dim msg As String
    If blank_expertise_means_exclusion Then
        AddMessage num_blank & " of " & num_markers * num_projects & " Project X Marker table entries are blank (excluded)"
    End If
    If num_excluded > 0 Then
        AddMessage num_excluded & " of " & num_markers * num_projects & " Project X Marker table entries are conflicts (excluded)"
    End If
    msg = "Loaded " & num_es & " expertise and conflict profiles into PXM table. "
    AddMessage msg
    
    Erase expertise, COIs, expertise_out
    LoadMarkerProjectExpertiseIntoPXM = True
    
End Function

Public Function LoadMarkerKeywordExpertiseIntoPXM() As Boolean
    ' load the data from all the project expertise sheets in a folder into the PXM table
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If ActivateCompetitionWorkbook = False Then
        LoadMarkerKeywordExpertiseIntoPXM = False
        Exit Function
    End If
    making_competition_workbook = False
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    
    ' initialize the PXM sheet
    Dim PXM_workbook As String
    PXM_workbook = ActiveWorkbook.Name
    Sheets(PROJECT_X_MARKER_SHEET).Select
    ClearPXMSheet
    
    ' get the folder name with the expertise files to load
    Dim folder_with_expertises As String
'    ChDir root_folder
    folder_with_expertises = SelectFolder("Select folder containing the confidence of markers by Keywords", _
                root_folder & expertise_by_keyword_received_folder)
    If Len(folder_with_expertises) = 0 Then
        Exit Function
    End If

    'load and process the xlsx files that have the right filename pattern.
    Dim looking As Boolean, expertise_sheet As String, marker_num As Long, marker_name As String
    Dim file_name As String, num_ms As Long, project_num As Long
    Dim rn As Long, i As Long, j As Long, num_es As Long
    num_es = 0                      ' counter of the number of expertise sheets
    Dim expertise() As Variant      ' arrays read from the expertise sheet
    Dim COIs() As Variant           ' the contents will be strings, but use variants to get from the S/S
    Dim finding As Boolean
    Dim keyword_row_range As String
    Dim expertise_out() As Variant
    ReDim expertise_out(1 To num_markers, 1 To num_keywords)
    Dim expertise_workbook As String
    ReDim competition_COIs(1 To num_projects, 1 To num_markers)
    
    finding = True
    While finding
        If num_es = 0 Then
            file_name = Dir(folder_with_expertises & fps & keyword_expertise_file_pattern)
        Else
            file_name = Dir()
        End If
        If Len(file_name) = 0 Then
            finding = False     'no more files to process
        Else        'open the file with a expertise sheet
            ' get name of expertise sheet from filename
            Workbooks.Open Filename:=folder_with_expertises & fps & file_name
            expertise_workbook = ActiveWorkbook.Name
            Dim sht_num As Long
            sht_num = getMarkingSheetName(file_name, expertise_by_keyword_ending, marker_num, marker_name)
            expertise_sheet = Sheets(sht_num).Name

'KEYWORD version
            ' load the keyword expertise column from that sheet, and store it in the output array
            Sheets(expertise_sheet).Select
            expertise = Range(MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW & ":" & _
                              MKET_EXPERTISE_COLUMN & (MKET_FIRST_DATA_ROW + num_keywords - 1))
            For j = 1 To num_keywords
                expertise_out(marker_num, j) = expertise(j, 1)
            Next j
            ' load the project COI column from the COI sheet
            Sheets(MKET_COI_SHEET_NAME).Select
            COIs = Range(c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                         c2l(MPET_COI_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1))
            'load this marker's set of project COIs into the project x marker array
            For i = 1 To num_projects
                If Len(COIs(i, 1)) > 0 Then
                    competition_COIs(i, marker_num) = "X"
                Else
                    ' empty cells mean no COI
                End If
            Next i
            ' close the expertise workbook
            Workbooks(expertise_workbook).Close
            num_es = num_es + 1
        End If
    Wend
    
    ' move the focus to the competition workbook where the results need to be stored
    If ActivateCompetitionWorkbook = False Then
        LoadMarkerKeywordExpertiseIntoPXM = False
        Exit Function
    End If
    
    ' Write the keyword expertise table to the marker expertise sheet
    Sheets(MARKER_EXPERTISE_SHEET).Select
    'find the column for this marker
    keyword_row_range = c2l(ME_FIRST_DATA_COL) & ME_FIRST_DATA_ROW
    ' insert the data about the marker's expertise in the first table
    Dim Destination As Range
    Set Destination = Range(keyword_row_range)
    Destination.Resize(UBound(expertise_out, 1), UBound(expertise_out, 2)).Value = expertise_out
    
    ' write the array of competition COIs to the PXM sheet
    Sheets(PROJECT_X_MARKER_SHEET).Select
    Dim cf_range As String
    cf_range = c2l(PXM_FIRST_DATA_COL) & PXM_FIRST_DATA_ROW & ":" & _
                c2l(PXM_FIRST_DATA_COL + num_markers - 1) & (PXM_FIRST_DATA_ROW + num_projects - 1)
    Set Destination = Range(cf_range)
    Destination.Resize(UBound(competition_COIs, 1), UBound(competition_COIs, 2)) = competition_COIs

    ' save the workbook containing the PXM sheet
    Range(c2l(PXM_FIRST_DATA_COL) & (PXM_FIRST_DATA_ROW)).Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Dim msg As String
    msg = "Loaded " & num_es & " expertise and conflict profiles into " & cwb
    AddMessage msg
    Erase expertise_out
    LoadMarkerKeywordExpertiseIntoPXM = True
    
End Function

Public Function LoadMarkerExpertiseIntoCrosswalk() As Boolean
    'make a duplicate of the template sheet
    'load the list of projects to mark
    'for each marker load the expertises they signaled, paying attention to any COIs signaled
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' initialize the crosswalk sheet
    Dim crosswalk_sheet As String
    ' this approach adds the data to the sheet in the current workbook
    crosswalk_sheet = EXPERTISE_CROSSWALK_SHEET
    
    clear_crosswalk_sheet (crosswalk_sheet)
    
    Dim crosswalk_workbook As String
    crosswalk_workbook = ActiveWorkbook.Name
    
    ' get the folder name with the expertise files to load
    Dim folder_with_expertises As String
'    ChDir root_folder
    folder_with_expertises = SelectFolder("Select folder containing the expertise specified by potential markers", _
                root_folder & expertise_by_project_received_folder)
    If Len(folder_with_expertises) = 0 Then
        Exit Function
    End If

    'load and process the xlsx files that have the right filename pattern.
    Dim looking As Boolean, expertise_sheet As String, marker_num As Long, marker_name As String
    Dim file_name As String, num_ms As Long, project_num As Long
    Dim rn As Long, i As Long, j As Long, num_es As Long, insert_pos As Long
    num_es = 0                      ' counter of the number of expertise sheets
    Dim expertise() As Variant      ' arrays read from the expertise sheet
    Dim COIs() As Variant           ' the contents will be strings, but use variants to get from the S/S
    Dim finding As Boolean
    Dim marker_number_row As String
    Dim expertise_out() As Variant
    ReDim expertise_out(1 To num_projects)
    Dim expertise_workbook As String
    
    finding = True
    While finding
        If num_es = 0 Then
            file_name = Dir(folder_with_expertises & fps & expertise_file_pattern)
        Else
            file_name = Dir()
        End If
        If Len(file_name) = 0 Then
            finding = False     'no more files to process
        Else        'open the file with a expertise sheet
            Workbooks.Open Filename:=folder_with_expertises & fps & file_name
            expertise_workbook = ActiveWorkbook.Name
            
            ' get name of expertise sheet from filename
            Dim sht_num As Long
            sht_num = getMarkingSheetName(file_name, EXPERTISE_ENDING, marker_num, marker_name)
            expertise_sheet = Sheets(sht_num).Name
            
            'move it into the workbook
            Sheets(expertise_sheet).Select
            num_es = num_es + 1
            insert_pos = Workbooks(crosswalk_workbook).Sheets.Count
            Sheets(expertise_sheet).Move after:=Workbooks(crosswalk_workbook).Sheets(insert_pos)
            Sheets(expertise_sheet).Activate
            Workbooks(expertise_workbook).Close
            
            'copy the column of expertise and the column of COI information from the expertise sheet
            expertise = Range(c2l(MPET_EXPERTISE_COLUMN) & "2:" & c2l(MPET_EXPERTISE_COLUMN) & (2 + num_projects - 1))
            COIs = Range(c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                         c2l(MPET_COI_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1))
            'update the expertise to exclude rows with COI signalled
            For i = 1 To num_projects
                Select Case COIs(i, 1)
                Case "X", "Y", "Mentor", "MENTOR", same_organization_text
                    expertise_out(i) = "X"
                Case "N", ""
                    expertise_out(i) = LMH2Percent(CStr(expertise(i, 1)))
                Case Else
                    MsgBox "[load_marker_expertise_in_crosswalk] unexpected COI[" & i & "] value: " & COIs(i), vbCritical
                    Exit Function
                End Select
            Next i
            'find the column for this marker
            Sheets(crosswalk_sheet).Select
            marker_number_row = c2l(ec_data_first_marker_column) & _
                                    (EC_FIRST_DATA_ROW - 1) & ":" & _
                                c2l(ec_data_first_marker_column - 1 + num_markers) & _
                                    (EC_FIRST_DATA_ROW - 1)
            Range(marker_number_row).Select 'select the row containing the marker #'s
            If Selection.Find(What:=marker_num, LookIn:=xlFormulas2, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False) Is Nothing Then
                'nothing found, nothing for this marker so exit the loop
                MsgBox _
                  "[LoadMarkerExpertiseIntoCrosswalk] error finding the marker numbers - check it is the right sheet", _
                  vbCritical
                finding = False
                Exit Function
            End If
            Selection.Find(What:=marker_num, after:=ActiveCell, LookIn:=xlFormulas2, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
            
            ' insert the column of data about the marker's expertise and availability to the crosswalk sheet
            ChangeActiveCell 1, 0  ' move down to the first row of data
            Dim Destination As Range
            Set Destination = ActiveCell
            Set Destination = Destination.Resize(UBound(expertise_out), 1)
            Destination.Value = Application.Transpose(expertise_out)
            Sheets(expertise_sheet).Delete
        End If
    Wend
    
    ' save the workbook containing the expertise crosswalk sheet and the loaded expertise inputs
    Range(c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW).Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Dim msg As String
    msg = "Loaded " & num_es & " expertise and conflict profiles into crosswalk table. "
    AddMessage msg
    Erase expertise_out
    LoadMarkerExpertiseIntoCrosswalk = True
    
End Function

Public Function num2LMH(num As Variant) As String
    Select Case num
    Case 1
        num2LMH = "L"
    Case 2
        num2LMH = "M"
    Case 3
        num2LMH = "H"
    Case ""
        ' ignore blanks
    Case Else
        MsgBox "Unexpected letter to LMH2Percent <" & num & ">", vbCritical
    End Select
End Function

Public Function LMH2Percent(letter As String) As Double
    Select Case letter
    Case "H"
        LMH2Percent = 1
    Case "M"
        LMH2Percent = TWO_THIRDS
    Case "L"
        LMH2Percent = ONE_THIRD
    Case ""
        If blank_expertise_means_exclusion Then
        Else    ' if they are not exclusions, then treat blanks as low expertise
            LMH2Percent = ONE_THIRD
        End If
    Case "X"
        ' ignore conflicts
    Case Else
        MsgBox "Unexpected letter to LMH2Percent <" & letter & ">", vbCritical
        LMH2Percent = -1
    End Select
    
End Function

Public Function AssignMarkers() As Boolean
        
    Dim starting_book As String, starting_sheet As String
    starting_book = ActiveWorkbook.Name
    starting_sheet = ActiveSheet.Name
    If ActivateCompetitionWorkbook = False Then
        AssignMarkers = False
        Exit Function
    End If

    making_competition_workbook = False
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' a little error checking
    If max_first_reader_assignments * num_markers < num_projects Then
        MsgBox "! Need to increase the limit on the number of first reader assignment - " & _
                "see competition parameters sheet", vbCritical
        Exit Function
    End If
    If num_markers * target_ass_per_marker < num_projects * target_markers_per_proj Then
        MsgBox "! Not enough slots per marker to cover the number of readers per project desired - " & _
                    " see competition parameters sheet (and recreate competition workbook).", vbCritical
    End If
    
    ' read the PXM data from the PXM sheet into the marker confidence array
    Sheets(PROJECT_X_MARKER_SHEET).Select
    Dim pxm_range As String, mc_range As String
    pxm_range = c2l(PXM_FIRST_DATA_COL) & PXM_FIRST_DATA_ROW & ":" & _
                c2l(PXM_FIRST_DATA_COL - 1 + num_markers) & (PXM_FIRST_DATA_ROW + num_projects - 1)
    pxm_table = Range(pxm_range)
    ' make a copy of the marker confidence array to help with filling holes in the assignments (swapping)
    ReDim pxm_at_start(1 To num_projects, 1 To num_markers)
    pxm_at_start = pxm_table
        
    ' Read in the current assignment information.
    ' This allows users to define some assignments, and then let the software complete the assignments
    ' initialize the marker assignment, and confidence of assigned marker arrays (N per project)
    Dim assignments_range As String
    ' target_markers_per_proj wide x num_projects deep
    Sheets(EXPERTISE_CROSSWALK_SHEET).Select
    assignments_range = c2l(ec_assignments_first_column) & EC_FIRST_DATA_ROW & ":" & _
                c2l(ec_assignments_first_column + target_markers_per_proj - 1) & _
                        (EC_FIRST_DATA_ROW + num_projects - 1)
    assignments = Range(assignments_range)
    ' make a copy of the assignments array before this run starts making more assignments.
    Dim assmts_as_started() As Variant
    assmts_as_started = assignments
    
    ' apply the assignments to the pxm table, and,
    ' make sure the mentors show up as COIs in the PXM table (for the direct route, or edits to PXM_table)
    Sheets(PROJECTS_SHEET).Select
    Dim mentors_range As String, i As Long, j As Long
    Dim mentors_array() As Variant, num_x_added As Long
    mentors_range = c2l(P_MENTOR_ID_COLUMN) & P_FIRST_DATA_ROW & ":" & _
                    c2l(P_MENTOR_ID_COLUMN) & (P_FIRST_DATA_ROW + num_projects - 1)
    mentors_array = Range(mentors_range)
    For i = 1 To num_projects
        For j = 1 To target_markers_per_proj
            If IsEmpty(assignments(i, j)) = False Then
                If IsEmpty(pxm_table(i, assignments(i, j))) = False Then
                    pxm_table(i, assignments(i, j)) = "A"
                End If
            End If
        Next j
        If mentors_array(i, 1) > 0 Then
            If (mentors_array(i, 1) > num_markers) Then
                AddMessage "mentor # for project " & i & " greater than number of markers."
            Else
                If pxm_table(i, mentors_array(i, 1)) <> "X" Then
                    num_x_added = num_x_added + 1
                End If
                If pxm_table(i, mentors_array(i, 1)) = "A" Then
                    AddMessage "PXM table: marker " & mentors_array(i, 1) & " was assigned to project " & j & _
                        ") but is in conflict."
                    assignments(i, j) = ""
                End If
                pxm_table(i, mentors_array(i, 1)) = "X"
            End If
        End If
    Next i
    If num_x_added > 0 Then
        AddMessage "Added " & num_x_added & " exclusions based on marker roles."
    End If

    ' copy the PXM array with updated exclusion and assignment info into the the expertise crosswalk sheet
    ' this updates the xlmh arrays (marker and project)
    Sheets(EXPERTISE_CROSSWALK_SHEET).Activate
    Dim Destination As Range
    mc_range = c2l(ec_data_first_marker_column) & _
                    (EC_FIRST_DATA_ROW) & ":" & _
               c2l(ec_data_first_marker_column - 1 + num_markers) & _
                    (EC_FIRST_DATA_ROW + num_projects - 1)
    Set Destination = Range(mc_range)
    Destination.Resize(UBound(pxm_table, 1), UBound(pxm_table, 2)).Value = pxm_table
        
    ' load the marker number array
    Dim mn_range As String
    mn_range = c2l(ec_data_first_marker_column) & _
                    (EC_FIRST_DATA_ROW - 1) & ":" & _
               c2l(ec_data_first_marker_column - 1 + num_markers) & _
                    (EC_FIRST_DATA_ROW - 1)
    mn_array = Range(mn_range)
    
    ' load the counts of ratings (H, M, L, X) per project
    Dim xlmh_projects_range As String
    ' 4 columns wide X num_projects deep
    xlmh_projects_range = c2l(EC_XLMH_CONF_PER_PROJECT_COL) & EC_FIRST_DATA_ROW & ":" & _
                c2l(EC_XLMH_CONF_PER_PROJECT_COL + 3) & (EC_FIRST_DATA_ROW + num_projects - 1)
    xlmh_per_project = Range(xlmh_projects_range)
    Dim num_projects_without_experts As Long, initial_experts_count() As Long
    ReDim initial_experts_count(1 To num_projects)
    ' check on projects that have no expertise signaled in the P X M table
    For i = 1 To num_projects
        initial_experts_count(i) = xlmh_per_project(i, 4) + xlmh_per_project(i, 3) + xlmh_per_project(i, 2)
        If ProjectHasExpertsAvailable(i) = False Then
            num_projects_without_experts = num_projects_without_experts + 1
        End If
    Next i
    If num_projects_without_experts > 0 Then
        AddMessage num_projects_without_experts & " projects without any expertise signaled, will be ignored"
    End If
    ' load the array of marker total H,M and X's per marker
    Dim xlmh_markers_range As String
    ' num_markers wide X 4 rows deep
    xlmh_markers_range = c2l(ec_data_first_marker_column) & EC_XLMH_MARKER_TABLE_FIRST_ROW & ":" & _
                c2l(ec_data_first_marker_column + num_markers - 1) & (EC_XLMH_MARKER_TABLE_FIRST_ROW + 3)
    xlmh_per_marker = Range(xlmh_markers_range)
    
    
'    ' array with the confidence of the assignments made (will be updated)
'    Dim coa_array_range As String
'    ' 'target_markers_per_proj' columns wide by num_projects rows
'    coa_array_range = c2l(EC_ASSMT_CONF_FIRST_COL) & EC_FIRST_DATA_ROW & ":" & _
''                c2l(EC_ASSMT_CONF_FIRST_COL + target_markers_per_proj - 1) & (EC_FIRST_DATA_ROW + num_projects - 1)
'    coa_array = Range(coa_array_range)
      
    Dim next_proj As Long        ' # of next project that should be assigned.
    
    ' load the arrays counting numbers assigned from the marker assignments read from the worksheet
    ReDim n_assigned2project(1 To num_projects)
    ReDim n_assigned2marker(1 To num_markers)
    ReDim n_per_assignment_col(1 To target_markers_per_proj)

    ' make sure we have data in the Project x Marker array
    Dim pxm_table_blank As Long, pxm_table_excluded As Long, pxm_table_assigned As Long
    For i = 1 To num_projects
        For j = 1 To num_markers
            Select Case pxm_table(i, j)
                Case ""
                    pxm_table_blank = pxm_table_blank + 1
                Case "X"
                    pxm_table_excluded = pxm_table_excluded + 1
                Case "A"
                    pxm_table_assigned = pxm_table_assigned + 1
            End Select
        Next j
    Next i
    ' a little more data checking (do we have enough marking slots after exclusions
    If num_projects * num_markers - pxm_table_excluded < _
        num_projects * target_markers_per_proj Then
        MsgBox "! Not enough available marker slots (" & pxm_table_excluded & " exclusions, " & _
                num_projects * target_markers_per_proj & " marking slots required, " & _
                num_projects * num_markers - pxm_table_excluded & " slots available).", vbCritical
        Exit Function
    End If
    ' make sure there is some data in the PXM table
    Const EMPTY_SHEET_THRESHOLD_PERCENT As Double = 90
    If pxm_table_blank > EMPTY_SHEET_THRESHOLD_PERCENT / 100 * CDbl(num_projects * num_markers) Then
        MsgBox CInt(pxm_table_blank / CDbl(num_projects * num_markers) * 100) & _
                " percent of the Project X Marker array in sheet " & _
                PROJECT_X_MARKER_SHEET & " is blank.  Expect few assignments.", vbOKOnly
    End If

    ' build some arrays useful for tracking progress, and the array with the confidence levels for assignments
    Dim num_assigned As Long, num_assigned_at_start As Long
    ReDim coa_array(1 To num_projects, 1 To target_markers_per_proj)
    num_assigned_at_start = 0
    For i = 1 To num_projects
        For j = 1 To target_markers_per_proj
            If assignments(i, j) > 0 Then
                coa_array(i, j) = GetConfidenceCode(CDbl(pxm_at_start(i, assignments(i, j))))
                n_assigned2project(i) = n_assigned2project(i) + 1
                n_assigned2marker(assignments(i, j)) = n_assigned2marker(assignments(i, j)) + 1
                n_per_assignment_col(j) = n_per_assignment_col(j) + 1
                num_assigned_at_start = num_assigned_at_start + 1
            End If
        Next j
    Next i
    num_assigned = num_assigned_at_start
        
    Dim assignment_col As Long          ' assignment column (1 = first reader, 2 = second reader ...
    assignment_col = 1                  ' start with the projects that need the most markers
    Dim best_marker As Long             ' number of marker proposed to review this project
    Dim looking As Boolean
    looking = True
    Dim mentor_num As Long
    Dim num_conflicts As Long
    Dim num_first_reader_assignments() As Long
    ReDim num_first_reader_assignments(1 To num_markers)
    ReDim assignment_failed_for_this_proj(1 To num_projects)
    Dim num_assignments_failed As Long
    Dim last_marker As Long              ' DEBUG DEBUG DEBUG
    
    ' loop through the projects assigning first those with the least expertise
    ' ready to start assigning
    While looking
        'find the project with the lowest available expertise
        next_proj = FindNextProject(assignment_col, next_proj)
        If (next_proj > 0) Then
            mentor_num = Sheets(PROJECTS_SHEET).Range(c2l(P_MENTOR_ID_COLUMN) & (P_FIRST_DATA_ROW + next_proj - 1)).Value
            best_marker = 0
            ' for this project, find the marker with the highest available confidence rating
            ' and (if there are multiple possibilities) find the lowest number this confidence ratings available
            For j = 1 To num_markers
                If (pxm_table(next_proj, j) = "X") Or _
                   (pxm_table(next_proj, j) = "A") Or _
                   (pxm_table(next_proj, j) = "") _
                     Then     ' no data (i.e. no expertise profile)
                    ' don't consider this marker if there is a COI or
                    ' if they have already been assigned to this project or
                    ' they did not provide an expertise sheet
                Else
                    If (n_assigned2marker(j) >= target_ass_per_marker) Or _
                        ((assignment_col = 1) And _
                        (max_first_reader_assignments > 0) And _
                        (num_first_reader_assignments(j) >= max_first_reader_assignments)) Then
                        ' we reached this marker's limit on the number of assignments
                        ' or the limit on the number of first reader assignments
                    Else
                        If j <> mentor_num Then
                            ' the candidate marker is not the mentor
                            If best_marker = 0 Then
                                best_marker = j
                            Else
                                If CompareConfidence(next_proj, j, best_marker) Then
                                    best_marker = j
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
            If best_marker > 0 Then
                ' first a couple of checks
                If best_marker = mentor_num Then
                    MsgBox "[AssignMarkers] Proposed marker is also the mentor ?????", vbCritical
                    Exit Function
                End If
                ' make the assignment
                assignments(next_proj, assignment_col) = best_marker
                ' also store the confidence for this project of the marker assigned
                coa_array(next_proj, assignment_col) = GetConfidenceCode(CDbl(pxm_table(next_proj, best_marker)))
                'update the other information arrays since we have removed a project & assigned one to a marker
                UpdateArrays next_proj, best_marker
                
                n_assigned2project(next_proj) = n_assigned2project(next_proj) + 1
                n_assigned2marker(best_marker) = n_assigned2marker(best_marker) + 1
                If assignment_col = 1 Then
                    num_first_reader_assignments(best_marker) = num_first_reader_assignments(best_marker) + 1
                End If
                
                num_assigned = num_assigned + 1
            Else
                assignment_failed_for_this_proj(next_proj) = True
                num_assignments_failed = num_assignments_failed + 1
'                MsgBox "unable to find available marker for project " & next_proj & ", terminating search", vbCritical
'                looking = False
            End If
        Else
            looking = False ' no more projects need assigning
        End If
    Wend
    
    'see of we can fill in the assignment holes
    If num_assignments_failed > 0 Then
        If FillAssignmentHoles(num_assigned, assmts_as_started) = False Then
            Exit Function
        End If
    End If
    
    ' check for the same marker showing up more than once on a project (BUG) or unexpected use of macros/sheets
    Dim k As Long
    For i = 1 To num_projects
        For j = 1 To target_markers_per_proj
            If (IsEmpty(assignments(i, j)) = False) Then
                For k = j + 1 To target_markers_per_proj
                    If IsEmpty(assignments(i, k)) = False Then
                        If assignments(i, j) = assignments(i, k) Then
                            PopMessage "For project " & i & ", Marker " & _
                                assignments(i, j) & " has a repeated assignment as reader " & j & ", and " & k, vbCritical
                        End If
                    End If
                Next k
            End If
        Next j
    Next i
    
    ' update the assignment number (the calculations correspond to marker numbers that go from 1 to num_markers
    ' however since the data is taken from a spreadsheet, the marker columns may have been reordered.
    ' update the assignment numbers to reflect the marker numbers specified in the crosswalk table
    Dim n_empty As Long
    For i = 1 To num_projects
        If initial_experts_count(i) > 0 Then
            ' only chieck on the projects that expertise available initially
            For j = 1 To target_markers_per_proj
                If IsEmpty(assignments(i, j)) Then
                    n_empty = n_empty + 1
                Else
                    assignments(i, j) = mn_array(1, assignments(i, j))
                End If
            Next j
        End If
    Next i
    If num_assigned = 0 Then
        MsgBox "Nothing assigned! - check that inputs have been provided.", vbCritical
        AssignMarkers = False
        Exit Function
    Else
        If n_empty > 0 Then
            AddMessage "NOTE: not all assignments made, " & n_empty & " assignments by hand required!"
        End If
    End If
    
    Sheets(EXPERTISE_CROSSWALK_SHEET).Activate
    ' write the marker assignment array
    Set Destination = Range(assignments_range)
    Destination.Resize(UBound(assignments, 1), UBound(assignments, 2)).Value = assignments
    
    ' write out the confidence letter for the selected marker on this project
    Dim coa_array_range As String
    coa_array_range = c2l(EC_ASSMT_CONF_FIRST_COL) & EC_FIRST_DATA_ROW & ":" & _
                      c2l(EC_ASSMT_CONF_FIRST_COL + target_markers_per_proj - 1) & _
                      (EC_FIRST_DATA_ROW + num_projects - 1)
    Set Destination = Range(coa_array_range)
    Destination.Resize(UBound(coa_array, 1), UBound(coa_array, 2)).Value = coa_array
    ' fit the width of the assignment and confidence columns to their contents
    Columns(c2l(EC_ASSMT_CONF_FIRST_COL + 1) & ":" & _
            c2l(EC_ASSMT_CONF_FIRST_COL + target_markers_per_proj - 1)).EntireColumn.AutoFit
    Columns(c2l(EC_ASSMT_CONF_FIRST_COL + target_markers_per_proj + 1) & ":" & _
            c2l(EC_ASSMT_CONF_FIRST_COL + 2 * target_markers_per_proj - 1)).EntireColumn.AutoFit
    ' treat the first and last columns differently since they have text headers
    Columns(c2l(EC_ASSMT_CONF_FIRST_COL)).ColumnWidth = _
        Range(c2l(EC_ASSMT_CONF_FIRST_COL) & 1).ColumnWidth
    Columns(c2l(EC_ASSMT_CONF_FIRST_COL + 2 * target_markers_per_proj - 1)).ColumnWidth = _
         Range(c2l(EC_ASSMT_CONF_FIRST_COL) & 1).ColumnWidth
    
    ' write out the array of marker confidences (updated for assignments)
    mc_range = c2l(ec_data_first_marker_column) & _
                (EC_FIRST_DATA_ROW) & ":" & _
               c2l(ec_data_first_marker_column - 1 + num_markers) & _
                (EC_FIRST_DATA_ROW + num_projects - 1)
    Set Destination = Range(mc_range)
    Destination.Resize(UBound(pxm_table, 1), UBound(pxm_table, 2)).Value = pxm_table
    ' number formatting and column fit to contents
    Range(mc_range).Select
    Selection.NumberFormat = "0.0"
    Range(c2l(ec_data_first_marker_column) & 1).Activate
    Columns(c2l(ec_data_first_marker_column) & ":" & _
            c2l(ec_data_first_marker_column - 1 + num_markers)).EntireColumn.AutoFit
    
    ' copy the assignments into the assignment master sheet
    ' REPLACED BY FORMULA LINK TO EXPERTISE CROSSWALK
    Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
    Dim ass_sht_ass_range As String
    ass_sht_ass_range = c2l(MAS_FIRST_ASSMT_COL) & MAS_FIRST_ASSMT_ROW & ":" & _
        c2l(MAS_FIRST_ASSMT_COL + target_markers_per_proj - 1) & (MAS_FIRST_ASSMT_ROW + num_projects - 1)
    Set Destination = Range(ass_sht_ass_range)
    Destination.Resize(UBound(assignments, 1), UBound(assignments, 2)).Value = assignments
    ' a bit of formatting
    Dim cn As Long
    ' the width for the marker assignment numbers
    cn = MAS_FIRST_ASSMT_COL
    Columns(c2l(cn) & ":" & c2l(cn + target_markers_per_proj - 1)).Select
    Selection.ColumnWidth = 5
    ' the mentor # column
    Columns(c2l(4)).EntireColumn.AutoFit
    '  the width of the names for those assigned
    cn = MAS_FIRST_ASSMT_COL + target_markers_per_proj
    Columns(c2l(cn) & ":" & c2l(cn + target_markers_per_proj - 1)).Select
    Range(c2l(cn) & 1).Activate
    Selection.ColumnWidth = 14

    cn = MAS_FIRST_ASSMT_COL + 2 * target_markers_per_proj + 1
    Columns(c2l(cn)).EntireColumn.AutoFit  ' marker #
    cn = cn + 2
    Columns(c2l(cn)).EntireColumn.AutoFit   ' # projects assigned
    cn = cn + 1
    Columns(c2l(cn)).EntireColumn.AutoFit   ' Issues re # assigned
    cn = cn + 1
    Columns(c2l(cn)).EntireColumn.AutoFit   ' # as first reviewer assigned
    
    AddMessage "After Assign Markers: " & num_assigned & " marking assignments for " _
            & num_projects & " projects and " & num_markers & " markers."
    
    Workbooks(starting_book).Activate
    Sheets(starting_sheet).Activate
    Erase num_first_reader_assignments, assignment_failed_for_this_proj
    AssignMarkers = True
    
End Function

Function MarkerOnThisProjectAlready(marker_num As Long, project_num As Long) As Boolean
    Dim i As Long
    For i = 1 To n_assigned2project(project_num)
        If assignments(project_num, i) = marker_num Then
            MarkerOnThisProjectAlready = True
            Exit Function
        End If
    Next i
End Function

Public Function FillAssignmentHoles(num_assignments As Long, ByRef assmts_as_started() As Variant) As Boolean

    ' see if we can make all the remaining assignments by swapping pairs of markers
    ' (one of which does not yet have a full suite of assignments)
    
    ' basic approach - if:
    '   a marker has room for more assignments
    '   and is not in conflict for that project
    '   and the assigned marker is not in conflict for the project needing markers
    '   and marker to insert has confidence on that propsal equal to the existing marker
    ' then:
    '   move the assigned marker to the empty slot
    '   put the marker with room for more assignments into the newly vacated slot

    ' search order:
    '   down the projects in each column of assignments looking for empty slots
    '   down all assigned markers in an assignment slot looking of candidates to pop
    '   through all the markers looking for one that has good enough confidence to replace the marker popped out

    Dim empty_proj As Long, j As Long, k As Long, m As Long
    Dim marker2move As Long, conf2move As Long, conf2insert As Long

    For j = 1 To target_markers_per_proj        ' for a given assignment column
        For empty_proj = 1 To num_projects               ' for each of the projects (that still need to be assigned readers)
            If assignments(empty_proj, j) = 0 Then
                ' empty assignment slot is for project 'i' in column 'j'
                ' go through all the project assignments in this assignment_col
                ' looking for one that can be 'popped-out' and used to fill the empty slot
                If num_assignments < num_markers * target_ass_per_marker Then
                    ' there are still markers that could be given an additional marking assignment
                    For k = 1 To num_projects
                        'look through the markers assigned to other projects in this column
                        ' see if the markers can be popped out
                        ' k is the project that might have its marker (marker2move) put in the empty slot
                        If (IsEmpty(assmts_as_started(k, j)) = False) And (assmts_as_started(k, j) = 0) Then
                            ' this marker was specified in the current run, OK to move it
                            marker2move = assignments(k, j)
                        Else
                            ' this marker was loaded at the start of this run, don't move it
                            marker2move = 0
                        End If
                        If (k <> empty_proj) And (marker2move > 0) Then
                            If (pxm_at_start(k, marker2move) <> "A") And _
                                (pxm_at_start(empty_proj, marker2move) <> "X") Then
                                ' project 'k' is not the empty slot, and it currently has an assignment
                                ' its assignment was not read in (thus fixed)
                                ' the marker is eligible for the empty slot
                                If (MarkerOnThisProjectAlready(marker2move, empty_proj) = False) Then
                                    ' the marker to fill the empty slot for this project is not already a reviewer on it.
                                    ' marker2move looks like a candidate to be popped, and replaced by
                                    For m = 1 To num_markers
                                        ' Look for a marker to put on project K
                                        If (n_assigned2marker(m) < target_ass_per_marker) And _
                                            (marker2move <> m) And _
                                            (MarkerOnThisProjectAlready(m, k) = False) And _
                                            (pxm_at_start(k, m) <> "X") Then
                                            ' The marker to backfill marker2move on project k must:
                                            '   have room for more assignments
                                            '   is not already on this project
                                            '   and is not in conflict for this project
                                            If (blank_expertise_means_exclusion = False) Or _
                                                ((IsEmpty(pxm_at_start(empty_proj, marker2move)) = False) And _
                                                 (IsEmpty(pxm_at_start(k, m)) = False)) Then
                                                ' if required, make sure marker has expertise
                                                conf2move = GetConfidenceCode(CDbl(pxm_at_start(k, marker2move)))
                                                conf2insert = GetConfidenceCode(CDbl(pxm_at_start(k, m)))
                                                If conf2insert >= conf2move Then
                                                    ' the new marker must have at least the same confidence code for a
                                                    ' project as the marker assigned
                                                    ' move  marker2move into the empty slot
                                                    assignments(empty_proj, j) = marker2move
                                                    coa_array(empty_proj, j) = GetConfidenceCode(CDbl(pxm_at_start(empty_proj, marker2move)))
                                                    n_assigned2project(empty_proj) = n_assigned2project(empty_proj) + 1
                                                    pxm_table(empty_proj, marker2move) = "A"
                                                    num_assignments = num_assignments + 1
                                                    
                                                    ' replace marker2move with m
                                                    assignments(k, j) = m
                                                    coa_array(k, j) = GetConfidenceCode(CDbl(pxm_at_start(k, m)))
                                                    n_assigned2marker(m) = n_assigned2marker(m) + 1
                                                    pxm_table(k, m) = "A"
                                                    
                                                    ' exit the loop and go look for another empty slot
                                                    k = num_projects
                                                    m = num_markers
                                                End If
                                            End If
                                        End If
                                    Next m
                                End If
                            End If
                        End If
                    Next k
                Else
                    ' no more markers should get assignments so we are done
                    Erase assmts_as_started
                    FillAssignmentHoles = True
                    Exit Function
                End If
            End If
        Next empty_proj
    Next j
    
    Erase assmts_as_started
    FillAssignmentHoles = True
End Function

Public Function FindNextProject(assignment_col As Long, last_proj_assigned As Long) As Long
' from among the projects that need  more reader/markers at the current 'assignment_col'
' How? find the project with the lowest available confidence,
' particularly the fewest high-confidence reviewers.
' if no more projects to assign at this level, move to the next level

    ' assignment_col is the column of the assignment array currently being filled with marker #'s)
    Dim next_project2assign As Long, this_one_is_better As Boolean
    Dim num_highs As Long, num_meds As Long, num_lows As Long
    ' initialize
    num_highs = 2 * num_projects
    num_meds = 2 * num_projects
    num_lows = 2 * num_projects
    Dim looking As Boolean
    looking = True
    Dim i As Long
    i = 0    'go through the full project list, as the next project to assign could now be before the last one assigned
    While looking
        i = i + 1
        If (i <= num_projects) Then
            ' run through all the projects
            If ProjectHasExpertsAvailable(i) And (assignment_failed_for_this_proj(i) = False) And _
                (assignments(i, assignment_col) = 0) Then
                ' still worth trying, and this project# needs a marker in this assignment_col
                If xlmh_per_project(i, 4) < num_highs Then
                    ' this project has fewer people rating this as @High confidence (compared to current choice)
                    this_one_is_better = True
                Else
                    If xlmh_per_project(i, 4) = num_highs Then
                        If xlmh_per_project(i, 3) < num_meds Then
                            ' they are the same @High, but this one has fewer @Medium's
                            this_one_is_better = True
                        Else
                            If xlmh_per_project(i, 3) = num_meds Then
                                ' They are the same @High and @Medium but this one has fewer @Lows
                                If xlmh_per_project(i, 2) < num_lows Then
                                    this_one_is_better = True
                                Else
                                    ' the current 'next_project2assign is still a better choice
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If this_one_is_better Then
            num_highs = xlmh_per_project(i, 4)
            num_meds = xlmh_per_project(i, 3)
            num_lows = xlmh_per_project(i, 2)
            next_project2assign = i
            this_one_is_better = False
        End If
        If i = num_projects Then
            ' we have gone through all the projects looking for assignment candidates in this column
            If (num_highs > num_projects) And (num_meds > num_projects) Then
                ' no projects found needing markers at this level
'                If n_per_assignment_col(assignment_col) <> num_projects Then
'                    MsgBox n_per_assignment_col(assignment_col) & " of " & _
'                    num_projects & " had full assignments, manual assigning required!", vbOKOnly
'                End If
                If assignment_col = target_markers_per_proj Then
                    ' no more levels, we are done looking for projects to assign
                    FindNextProject = 0   ' flag that there are no projects that need a reader assigned
                    Exit Function
                Else
                    ' move to the next level
                    assignment_col = assignment_col + 1
                    i = 0
                End If
            Else
                ' a candidate was found, return it
                FindNextProject = next_project2assign
                looking = False
            End If
        End If
    Wend

End Function

Function ProjectHasExpertsAvailable(project_num As Long) As Boolean
    If xlmh_per_project(project_num, 4) > 0 Or _
       xlmh_per_project(project_num, 3) > 0 Or _
       xlmh_per_project(project_num, 2) > 0 Then
        ProjectHasExpertsAvailable = True
    Else
         ProjectHasExpertsAvailable = False
    End If

End Function

Function UpdateArrays(project_num As Long, marker_num As Long) As Boolean
    ' since a marker has been assigned to a project, the number of markers available for the project
    ' and the number of available confidence specifications for a marker has been reduced
    Dim conf As String
    conf = GetConfidenceCode(CDbl(pxm_table(project_num, marker_num)))
    Select Case conf
    Case 3
        xlmh_per_project(project_num, 4) = xlmh_per_project(project_num, 4) - 1
        xlmh_per_marker(4, marker_num) = xlmh_per_marker(4, marker_num) - 1
    Case 2
        xlmh_per_project(project_num, 3) = xlmh_per_project(project_num, 3) - 1
        xlmh_per_marker(3, marker_num) = xlmh_per_marker(3, marker_num) - 1
    Case 1
        xlmh_per_project(project_num, 2) = xlmh_per_project(project_num, 2) - 1
        xlmh_per_marker(2, marker_num) = xlmh_per_marker(2, marker_num) - 1
    Case Else
        MsgBox "! Error: marker selected for a project they are in conflict for", vbCritical
        UpdateArrays = False
        Exit Function
    End Select
    pxm_table(project_num, marker_num) = "A"   'flag the marker has been assigned to this project
    UpdateArrays = True
    
End Function

Public Function CompareConfidence(this_proj As Long, this_marker As Long, best_marker As Long) As Boolean
    ' select between two candidate markers. Choose the one with the higher confidence.
    ' in case of a tie choose the one with more confidence rankings at this level
    
    Dim best_ranking As Long, this_ranking As Long
    Dim this_marker_num_available As Long, best_marker_num_available As Long
    
    Dim best_letter As String, this_letter As String
    If pxm_table(this_proj, this_marker) > pxm_table(this_proj, best_marker) Then
        best_marker = this_marker
        CompareConfidence = True
    Else
        best_ranking = GetConfidenceCode(CDbl(pxm_table(this_proj, best_marker)))
        this_ranking = GetConfidenceCode(CDbl(pxm_table(this_proj, this_marker)))
        If this_ranking = best_ranking Then
            this_marker_num_available = xlmh_per_marker(this_ranking + 1, this_marker)
            best_marker_num_available = xlmh_per_marker(this_ranking + 1, best_marker)
            If this_marker_num_available > best_marker_num_available Then
                best_marker = this_marker
                CompareConfidence = True
            Else
                CompareConfidence = False
            End If
        End If
    End If
End Function

Public Function GetConfidenceCode(confidence_level As Double) As Long
    If confidence_level < 0 Then
        MsgBox "! Confidence level = " & confidence_level & " should be 0 to 1", vbCritical
        GetConfidenceCode = -1
    Else
        If confidence_level <= ONE_THIRD Then
            GetConfidenceCode = 1
        Else
            If confidence_level <= TWO_THIRDS Then
                GetConfidenceCode = 2
            Else
                If confidence_level <= 1 Then
                    GetConfidenceCode = 3
                Else
                    MsgBox "confidence level = " & confidence_level & " should be 0 to 1", vbCritical
                    GetConfidenceCode = -1
                End If
            End If
        End If
    End If
End Function

Public Function CreateFigureOfMeritTable() As Boolean
    
    If (CreatePXMFromProjectRelevanceAndMarkerExpertise = False) Then
        Exit Function
    End If
    CreateFigureOfMeritTable = True
End Function

Public Function CreatePXMFromProjectRelevanceAndMarkerExpertise() As Boolean
' calculate the Figure of Merits (FOM) which estimate the confidence of a marker to review a project
' given:
' a table of marker confidence with regards to a number of keywords or themes
' a table of the relevance of projects to the same keywords or themes
' weights on each of the keywords
' then the FOM for marker i's confidence to review project j is estimated as the sum of the products of
'   (the marker's confidence for a keyword times
'   the project's relevance to that keyword
'   the weight of that keyword)
'   divided by the sum of the weights

    If ActivateCompetitionWorkbook = False Then
        CreatePXMFromProjectRelevanceAndMarkerExpertise = False
        Exit Function
    End If
    making_competition_workbook = False
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    Dim i As Long, j As Long, k As Long
    Dim kw_range As String, me_range As String, pk_range As String, pxm_range As String
    Dim kw_weights() As Variant, me_array() As Variant, pk_array() As Variant
    ReDim pxm_table(1 To num_projects, 1 To num_markers)
    
    ' read in the keyword weights
    Sheets(KEYWORDS_SHEET).Activate
    kw_range = c2l(KW_WEIGHTS_COL) & KW_WEIGHTS_ROW & ":" & _
               c2l(KW_WEIGHTS_COL) & (KW_WEIGHTS_ROW - 1 + num_keywords)
    kw_weights = Range(kw_range)
    If is_array_empty(kw_weights) = True Then
        AddMessage "Keyword weights array is empty, check table in " & KEYWORDS_SHEET & " sheet."
        CreatePXMFromProjectRelevanceAndMarkerExpertise = False
        Exit Function
    End If
    
    ' read in the marker relevances to keywords
    Sheets(MARKER_EXPERTISE_SHEET).Activate
    me_range = c2l(ME_FIRST_DATA_COL) & ME_FIRST_DATA_ROW & ":" & _
               c2l(ME_FIRST_DATA_COL - 1 + num_keywords) & (ME_FIRST_DATA_ROW + num_markers - 1)
    me_array = Range(me_range)
    If is_array_empty(me_array) = True Then
        AddMessage "Array of marker confidence on keywords is empty, check table in " & _
                    MARKER_EXPERTISE_SHEET & " sheet."
        CreatePXMFromProjectRelevanceAndMarkerExpertise = False
        Exit Function
    End If

    ' read in the project keyword confidences
    Sheets(PROJECT_KEYWORDS_SHEET).Activate
    pk_range = c2l(PK_FIRST_DATA_COL) & PK_FIRST_DATA_ROW & ":" & _
               c2l(PK_FIRST_DATA_COL - 1 + num_keywords) & (PK_FIRST_DATA_ROW + num_projects - 1)
    pk_array = Range(pk_range)
    If is_array_empty(pk_array) = True Then
        AddMessage "Array of project ratings by keyword is empty, check table in " & _
                    PROJECT_KEYWORDS_SHEET & " sheet."
        CreatePXMFromProjectRelevanceAndMarkerExpertise = False
        Exit Function
    End If

    ' read in the column of mentor numbers for projects
    Sheets(PROJECTS_SHEET).Activate
    Dim mentor_range As String
    mentor_range = (c2l(P_MENTOR_ID_COLUMN) & P_FIRST_DATA_ROW) & ":" & _
                   (c2l(P_MENTOR_ID_COLUMN) & P_FIRST_DATA_ROW + num_projects - 1)
    mentor_column = Range(mentor_range)
    
    ' read in the existing PXM table to get the conflicts of interest already specified
    Sheets(PROJECT_X_MARKER_SHEET).Activate
    pxm_range = c2l(PXM_FIRST_DATA_COL) & PXM_FIRST_DATA_ROW & ":" & _
                c2l(PXM_FIRST_DATA_COL + num_markers - 1) & (PXM_FIRST_DATA_ROW + num_projects - 1)
    pxm_table = Range(pxm_range)
    
    ' calculate the PXM ratings
    Dim max_conf_calcd As Double, min_conf_calcd As Double, pk_percent As Double, me_percent As Double
    max_conf_calcd = 0
    min_conf_calcd = 99999999 ' should go zero to one (max)
    For i = 1 To num_projects
        For j = 1 To num_markers
            If IsEmpty(mentor_column(i, 1)) = False Then
                If mentor_column(i, 1) = j Then
                    ' marker is a mentor - exclude them
                    pxm_table(i, j) = "X"
                End If
            End If
            If i = 36 And j = 50 Then
                j = j
            End If
            If pxm_table(i, j) <> "X" Then
                ' this marker IS eligible to mark this project
                pxm_table(i, j) = 0
                For k = 1 To num_keywords
                    pk_percent = LMH2Percent(CStr(pk_array(i, k)))
                    me_percent = LMH2Percent(CStr(me_array(j, k)))
                    If pk_percent >= 0 And me_percent >= 0 Then
                        pxm_table(i, j) = pxm_table(i, j) + _
                        kw_weights(k, 1) * pk_percent * me_percent
                    Else
                        ' Somethings wrong with the data, exit
                        AddMessage "Looks like something is wrong with the Project/Marker expertise data"
                        CreatePXMFromProjectRelevanceAndMarkerExpertise = False
                        Exit Function
                    End If
                Next k
                If max_conf_calcd < pxm_table(i, j) Then
                    max_conf_calcd = pxm_table(i, j)
                Else
                    If min_conf_calcd > pxm_table(i, j) Then
                        min_conf_calcd = pxm_table(i, j)
                    End If
                End If
            Else
            End If
        Next j
    Next i
    
    ' scale the PXM to go from .1 to 1 (if blank treated as low confidence) or 0 to 1 (if blank treated as exclusion)
    Dim num_conflicts As Long, num_blanks As Long
    num_conflicts = 0
    num_blanks = 0
    Dim target_min_conf As Double, target_max_conf As Double
    If blank_expertise_means_exclusion Then
        target_min_conf = 0
        target_max_conf = 1
    Else    ' if they are not exclusions, then treat blanks as low expertise
        target_min_conf = 0.1
        target_max_conf = 1
    End If
    
    For i = 1 To num_projects
        For j = 1 To num_markers
            If pxm_table(i, j) = "X" Then
                num_conflicts = num_conflicts + 1
            Else
                If pxm_table(i, j) = 0 And blank_expertise_means_exclusion Then
                    pxm_table(i, j) = ""
                    num_blanks = num_blanks + 1
                Else
                    If mentor_column(i, 1) = j Then    ' double check, (in case of manual editting)
                        pxm_table(i, j) = "X" 'flag the mentor is in conflict for this project
                        num_conflicts = num_conflicts + 1
                    Else
                        pxm_table(i, j) = target_min_conf + (target_max_conf - target_min_conf) * (pxm_table(i, j) - min_conf_calcd) / _
                                                        (max_conf_calcd - min_conf_calcd)
                    End If
                End If
            End If
        Next j
    Next i
    Dim total_cells As Long, num_required As Long, num_available As Long
    total_cells = num_projects * num_markers
    num_required = num_projects * target_markers_per_proj
    If blank_expertise_means_exclusion And (num_blanks > 0) Then
        AddMessage num_blanks & " of " & total_cells & " Project X Marker table entries are blank (excluded)"
    End If
    If num_conflicts > 0 Then
        AddMessage num_conflicts & " of " & total_cells & " Project X Marker table entries are conflicts (excluded)"
    End If
    If blank_expertise_means_exclusion Then
        num_available = total_cells - num_blanks - num_conflicts
    Else
        num_available = total_cells - num_conflicts
        num_blanks = 0
    End If
    If num_available < num_required Then
        AddMessage "The " & (num_conflicts + num_blanks) & " exclusions only leave " & _
                num_available & " entries for assignments, BUT need " & _
                num_required & " entries. Expect incomplete assignments."
    Else
            AddMessage num_available & " slots available for " & num_required & " marking assignments."
    End If
    
    ' write the PXM array
    Dim Destination As Range
    Set Destination = Range(pxm_range)
    Destination.Resize(UBound(pxm_table, 1), UBound(pxm_table, 2)).Value = pxm_table
    Range(FirstCell(pxm_range)).Activate
    
    'make the columns readable
    Dim table_cols As String
    table_cols = c2l(PXM_FIRST_DATA_COL) & ":" & c2l(PXM_FIRST_DATA_COL - 1 + num_markers)
    Columns(table_cols).Select
    Columns(table_cols).EntireColumn.AutoFit

    CreatePXMFromProjectRelevanceAndMarkerExpertise = True
    
End Function

Public Function ClearPXMSheet() As Boolean
    
    Sheets(PROJECT_X_MARKER_SHEET).Activate
    Dim pxm_range As String
    pxm_range = c2l(PXM_FIRST_DATA_COL) & PXM_FIRST_DATA_ROW & ":" & _
                c2l(PXM_FIRST_DATA_COL - 1 + num_markers) & (PXM_FIRST_DATA_ROW + num_projects - 1)
    Range(pxm_range).Select
    Range(FirstCell(pxm_range)).Activate
    Selection.ClearContents
    Range(FirstCell(pxm_range)).Select
    Range(FirstCell(pxm_range)).Activate
    ActiveWindow.Zoom = 100

    ClearPXMSheet = True
    
End Function

Public Function FindHeaderColumn(row_num As Long, search_text As String, exact As Boolean) As Long
    ' find and move to a cell in the specified row that contains the search text
    Dim found As Boolean
    Dim search_flag As Long
    If exact = True Then
        search_flag = xlWhole
    Else
        search_flag = xlPart
    End If
    Rows(row_num).Select
    
    If Selection.Find(What:=search_text, after:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=search_flag, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False) Is Nothing Then
        'nothing found, nothing for this marker so exit the loop
        found = False
    Else
        Selection.Find(What:=search_text, after:=ActiveCell, LookIn:=xlFormulas2, _
            LookAt:=search_flag, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        found = True
    End If
    If found Then
        FindHeaderColumn = ActiveCell.Column
    Else
        FindHeaderColumn = 0
    End If
End Function

Public Function PopulateAnalysisSheet(loading_scores As Boolean) As Boolean
    Dim start_sheet As String, start_book As String
    start_book = ActiveWorkbook.Name
    start_sheet = ActiveSheet.Name
    If loading_scores Then
        ' don't need this initialization stuff, for this use
    Else
        If ActivateCompetitionWorkbook = False Then
            PopulateAnalysisSheet = False
            Exit Function
        End If
        making_competition_workbook = False
        globals_defined = False
        If DefineGlobals = False Then
            Exit Function
        End If '
    End If
    
    ' load the first table from the results sheet
    Sheets(RESULTS_SHEET).Select
    Dim results_table_range As String, table_out_range As String
    results_table_range = "A" & R_FIRST_DATA_ROW & ":" & _
                        c2l(R_T1N + num_criteria - 2) & _
                            (num_projects * target_markers_per_proj + R_FIRST_DATA_ROW - 1)
    Dim results_table() As Variant
    results_table = Range(results_table_range)

    ' get the maximum number of readers assigned to any project (for sizing the table)
    Dim i As Long, j As Long, num_readers() As Long, max_readers As Long, pn As Long
    Dim scores_provided As Long
    ReDim num_readers(1 To num_projects)
    For i = 1 To num_projects * target_markers_per_proj
        If IsEmpty(results_table(i, 5)) = False Then
            pn = results_table(i, 5)
            scores_provided = 0
            For j = 1 To num_criteria
                If IsEmpty(results_table(i, 7 + j)) = False Then
                    scores_provided = scores_provided + 1
                End If
            Next j
            ' make sure we have scores in all criteria from the marker
            If scores_provided = num_criteria Then
                num_readers(pn) = num_readers(pn) + 1
            Else
                If IsEmpty(results_table(i, 1)) = False Then
                    AddMessage "Warning: Marker " & results_table(i, 2) & " provided " & scores_provided & _
                            " of " & num_criteria & " required scores for project " & pn & " - ignored."
                End If
            End If
        End If
    Next i
    For i = 1 To num_projects
        If num_readers(i) > max_readers Then
            max_readers = num_readers(i)
        End If
    Next i
    
    ' set up the analysis table to recieve its data
    Sheets(ANALYSIS_SHEET).Visible = True
    Sheets(ANALYSIS_SHEET).Select
    If loading_scores Then
        ' don't need to expand this sheet if we are about to load in scores from markers
    Else
        ExpandAnalysisSheet max_readers
    End If
    
    ' build a table of raw project scores sorted by increasing project and reader #'s
    Dim table_out() As Variant, rdr_num As Long
    ReDim table_out(1 To num_projects, 1 To 2 + max_readers)
    For i = 1 To num_projects * target_markers_per_proj
        If IsEmpty(results_table(i, 5)) = False And IsEmpty(results_table(i, 3)) = False Then
            pn = results_table(i, 5)
            rdr_num = results_table(i, 3)
            table_out(pn, 1) = pn
            If (IsEmpty(results_table(i, R_T1N + num_criteria - 2)) = False) And _
                (Len(results_table(i, R_T1N + num_criteria - 2)) > 0) And _
                (results_table(i, R_T1N + num_criteria - 3) = num_criteria) Then
                ' only use out non-blank, full-criteria project scores
                table_out(pn, 1 + rdr_num) = results_table(i, R_T1N + num_criteria - 2)
            End If
            table_out(pn, 2 + max_readers) = num_readers(pn)
        End If
    Next i
    'write out the raw readers' score table to the analysis sheet
    Dim Destination As Range
    Set Destination = Range(c2l(A_PROJ_NUM_COL) & A_FIRST_DATA_ROW)
    Destination.Resize(UBound(table_out, 1), UBound(table_out, 2)).Value = table_out
    Erase results_table
    
    If loading_scores Then
        ' don't need this initialization stuff, its done already
    Else
        ' get rid of the formulae in the rank column in case things get sorted ...
        ConvertRangeToText "A1:" & "A" & (A_FIRST_DATA_ROW - 1 + num_projects)
    End If
    
    ' now repeat for the second table (normalized scores)
    ' load the normalized scores by criteria table from the third table on the results sheet
    Sheets(RESULTS_SHEET).Select
    results_table_range = c2l(R_T3S + num_criteria - 2) & R_FIRST_DATA_ROW & ":" & _
                          c2l(R_T3S + R_T3N + 2 * (num_criteria - 2) - 1) & _
                            (num_projects * target_markers_per_proj + R_FIRST_DATA_ROW - 1)
    results_table = Range(results_table_range)
    
    ' build the table of normalized reader scores
    ReDim table_out(1 To num_projects, 1 To max_readers)
    For i = 1 To num_projects * target_markers_per_proj
        If results_table(i, 1) > 0 And results_table(i, 4) > 0 Then
            pn = results_table(i, 1)
            rdr_num = results_table(i, 4)
            If (IsEmpty(results_table(i, R_T3N + num_criteria - 2)) = False) And _
                (Len(results_table(i, R_T3N + num_criteria - 2)) > 0) And _
                (CountScores(results_table, i) = num_criteria) Then
                'only use non-blank, full criteria project scores
                table_out(pn, rdr_num) = results_table(i, R_T3N + num_criteria - 2)
            End If
        End If
    Next i
    
    'write out the normalized readers score table to the analysis sheet
    Sheets(ANALYSIS_SHEET).Select
    Set Destination = Range(c2l(A_T2S + max_readers - 2) & A_FIRST_DATA_ROW)
    Destination.Resize(UBound(table_out, 1), UBound(table_out, 2)).Value = table_out
    Erase results_table
    
    ' for now, don't use the ranking from the results table
    ' this makes the analysis more a function of the scores, and less any manual reranking.
    If False Then
        ' load the 4th table to get the project rank
        Sheets(RESULTS_SHEET).Select
        results_table_range = c2l(R_T4S + 2 * (num_criteria - 2)) & R_FIRST_DATA_ROW & ":" & _
                            c2l(R_T4S + R_T4N + 3 * (num_criteria - 2) - 1) & _
                                (num_projects * target_markers_per_proj + R_FIRST_DATA_ROW - 1)
        results_table = Range(results_table_range)
        ' build a single column table for the project rank
        ReDim table_out(1 To num_projects, 1 To 1)
        For i = 1 To num_projects
            If IsEmpty(results_table(i, R_T4N + num_criteria - 2)) = False Then
                pn = results_table(i, 1)
                table_out(pn, 1) = results_table(i, R_T4N + num_criteria - 2)
            End If
        Next i
        'write the project rank column
        Sheets(ANALYSIS_SHEET).Select
        Set Destination = Range(c2l(A_T2S + A_T2N - 1 + 2 * (max_readers - 2)) & A_FIRST_DATA_ROW)
        Destination.Resize(UBound(table_out, 1), UBound(table_out, 2)).Value = table_out
    End If
    
    ' also keep the formulae on this sheet (should be OK since there are not off-sheet vlookups).
    If False Then
        ' get rid of the formulae in case things get sorted ... and other formatting
        i = A_T2S + A_T2N + 2 * (num_criteria - 2) - 1
        ConvertRangeToText "A1:" & c2l(i) & (A_FIRST_DATA_ROW + num_projects - 1)
    End If
    
    If loading_scores Then
        ' should already be formatted
    Else
        FormatAnalysisSheet max_readers
    End If
    
    ' sort the table by the normalized scores (in ascending order)
    Dim sort_range As String, key1_range As String, key2_range As String
    key1_range = c2l(A_T2S + 2 * (target_markers_per_proj - 2) + 3) & A_FIRST_DATA_ROW & ":" & _
                c2l(A_T2S + 2 * (target_markers_per_proj - 2) + 3) & _
                    (A_FIRST_DATA_ROW + num_projects - 1) 'first sort by marker
    sort_range = c2l(A_T1S + 1) & A_FIRST_DATA_ROW & ":" & _
              c2l(A_T2S + A_T2N + 2 * (target_markers_per_proj - 2) - 2) & (A_FIRST_DATA_ROW + num_projects - 1)
    With ActiveWorkbook.Worksheets(ANALYSIS_SHEET).Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range(key1_range), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A1").Select
    Workbooks(start_book).Activate
    Sheets(start_sheet).Select
    
    Erase table_out, num_readers, results_table

    PopulateAnalysisSheet = True
End Function

Function CountScores(ByRef results_table() As Variant, rownum As Long) As Long
    ' this only works for the third table (normalized marker score rows)
    Dim i As Long
    CountScores = 0
    For i = UBound(results_table, 2) - num_criteria To UBound(results_table, 2) - 1
        If IsEmpty(results_table(rownum, i)) = False And Len(results_table(rownum, i)) > 0 Then
            CountScores = CountScores + 1
        End If
    Next i
End Function

Public Function MakeAverageAndSpanChart() As Boolean
'
    Dim i As Long
    If ActivateCompetitionWorkbook = False Then
        MakeAverageAndSpanChart = False
        Exit Function
    End If
    ' see if the chart exists already
    For i = 1 To Sheets.Count
        If Sheets(i).Name = ANALYSIS_CHART_NAME Then
'            PopMessage "A Chart named " & ANALYSIS_CHART_NAME & _
'                        " exists already. Please rename or remove it to create the new chart.", vbCritical
            MakeAverageAndSpanChart = True
            Exit Function
        End If
    Next i
    
    Const RAW_AVG_COLOFF As Long = 4
    Const RAW_MIN_COLOFF As Long = 5
    Const RAW_MAX_COLOFF As Long = 6
    Const RAW_SPAN_COLOFF As Long = 7
    Const NORM_AVG_COLOFF As Long = 2
    Const NORM_MIN_COLOFF As Long = 3
    Const NORM_MAX_COLOFF As Long = 4
    Const NORM_SPAN_COLOFF As Long = 5
    
    Dim CHART_TITLE As String, nr As Long
    CHART_TITLE = "Competition Results:" & Chr(13) & "Ranked by Increasing (normalized) Scores"
    nr = target_markers_per_proj
    Dim col_range As String, an_sht As String
    an_sht = ANALYSIS_SHEET
    
    col_range = c2l(A_T1N + nr - 2 + nr + NORM_AVG_COLOFF) & A_FIRST_DATA_ROW - 1 & ":" & _
                c2l(A_T1N + nr - 2 + nr + NORM_AVG_COLOFF) & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
                c2l(A_T1N + nr - 2 + nr + NORM_MIN_COLOFF) & A_FIRST_DATA_ROW - 1 & ":" & _
                c2l(A_T1N + nr - 2 + nr + NORM_MIN_COLOFF) & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
                c2l(A_T1N + nr - 2 + nr + NORM_MAX_COLOFF) & A_FIRST_DATA_ROW - 1 & ":" & _
                c2l(A_T1N + nr - 2 + nr + NORM_MAX_COLOFF) & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
                c2l(A_T1N + nr - 2 + nr + NORM_SPAN_COLOFF) & A_FIRST_DATA_ROW - 1 & ":" & _
                c2l(A_T1N + nr - 2 + nr + NORM_SPAN_COLOFF) & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
                c2l(RAW_SPAN_COLOFF + nr) & A_FIRST_DATA_ROW - 1 & ":" & _
                c2l(RAW_SPAN_COLOFF + nr) & A_FIRST_DATA_ROW + num_projects_scored - 1
                
'                c2l(RAW_MIN_COLOFF + num_criteria) & A_FIRST_DATA_ROW - 1 & ":" & _
'                c2l(RAW_MIN_COLOFF + num_criteria) & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
'                c2l(RAW_MAX_COLOFF + num_criteria) & A_FIRST_DATA_ROW - 1 & ":" & _
'                c2l(RAW_MAX_COLOFF + num_criteria) & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _


    Range(col_range).Select
    Range(FirstCell(col_range)).Activate
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    col_range = an_sht & "!$" & c2l(A_T1N + nr - 2 + nr + NORM_AVG_COLOFF) & "$" & A_FIRST_DATA_ROW - 1 & ":" & _
                "$" & c2l(A_T1N + nr - 2 + nr + NORM_AVG_COLOFF) & "$" & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
                an_sht & "!$" & c2l(A_T1N + nr - 2 + nr + NORM_MIN_COLOFF) & "$" & A_FIRST_DATA_ROW - 1 & ":" & _
                "$" & c2l(A_T1N + nr - 2 + nr + NORM_MIN_COLOFF) & "$" & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
                an_sht & "!$" & c2l(A_T1N + nr - 2 + nr + NORM_MAX_COLOFF) & "$" & A_FIRST_DATA_ROW - 1 & ":" & _
                "$" & c2l(A_T1N + nr - 2 + nr + NORM_MAX_COLOFF) & "$" & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
                an_sht & "!$" & c2l(A_T1N + nr - 2 + nr + NORM_SPAN_COLOFF) & "$" & A_FIRST_DATA_ROW - 1 & ":" & _
                "$" & c2l(A_T1N + nr - 2 + nr + NORM_SPAN_COLOFF) & "$" & A_FIRST_DATA_ROW + num_projects_scored - 1 & "," & _
                an_sht & "!$" & c2l(RAW_SPAN_COLOFF + nr) & "$" & A_FIRST_DATA_ROW - 1 & ":" & _
                "$" & c2l(RAW_SPAN_COLOFF + nr) & "$" & A_FIRST_DATA_ROW + num_projects_scored - 1
    ActiveChart.SetSourceData Source:=Range(col_range)
'        "Analysis!$S$3:$S$10,Analysis!$T$3:$T$10,Analysis!$U$3:$U$10,Analysis!$V$3:$V$10,Analysis!$K$3:$K$10")
    
    ActiveChart.ChartTitle.Text = CHART_TITLE
    
    ' convert the spans to columns
'    ActiveChart.FullSeriesCollection(5).Select
'    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
'    ActiveChart.FullSeriesCollection(5).ChartType = xlColumnClustered
    
    ' convert the min and max plots to dashed
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .DashStyle = msoLineDash
    End With
    ActiveChart.FullSeriesCollection(4).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .DashStyle = msoLineDash
    End With
    ActiveChart.ChartTitle.Select
    Dim chart_name As String
    chart_name = ActiveSheet.ChartObjects(1).Name
    ActiveSheet.ChartObjects(chart_name).Activate

    ' not sure what this does
    If False Then
        Selection.Format.TextFrame2.TextRange.Characters.Text = CHART_TITLE
        With Selection.Format.TextFrame2.TextRange.Characters(1, 21).ParagraphFormat
            .TextDirection = msoTextDirectionLeftToRight
            .Alignment = msoAlignCenter
        End With
        With Selection.Format.TextFrame2.TextRange.Characters(1, 19).Font
            .BaselineOffset = 0
            .Bold = msoFalse
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(89, 89, 89)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 14
            .Italic = msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = msoNoUnderline
            .Spacing = 0
            .Strike = msoNoStrike
        End With
        With Selection.Format.TextFrame2.TextRange.Characters(20, 2).Font
            .BaselineOffset = 0
            .Bold = msoFalse
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(89, 89, 89)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 14
            .Italic = msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = msoNoUnderline
            .Spacing = 0
            .Strike = msoNoStrike
        End With
        With Selection.Format.TextFrame2.TextRange.Characters(22, 47).ParagraphFormat
            .TextDirection = msoTextDirectionLeftToRight
            .Alignment = msoAlignCenter
        End With
        With Selection.Format.TextFrame2.TextRange.Characters(22, 47).Font
            .BaselineOffset = 0
            .Bold = msoFalse
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(89, 89, 89)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 14
            .Italic = msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = msoNoUnderline
            .Spacing = 0
            .Strike = msoNoStrike
        End With
    End If
    
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=ANALYSIS_CHART_NAME
    Sheets(ANALYSIS_CHART_NAME).Select
    Sheets(ANALYSIS_CHART_NAME).Move after:=Sheets(ANALYSIS_SHEET)
    
    MakeAverageAndSpanChart = True
    
End Function

Public Function FormatAnalysisSheet(max_readers As Long) As Boolean
    Dim raw_readers_range As String, norm_readers_range As String, frr As Long
    frr = A_FIRST_RAW_READER_COLUMN
    raw_readers_range = c2l(frr) & "1:" & c2l(frr + max_readers - 1) & 1
    Range(raw_readers_range).Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    norm_readers_range = c2l(A_T2S + frr - 1) & "1:" & _
          c2l(A_T2S + frr + max_readers - 2) & 1
    Range(norm_readers_range).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Dim this_range As String
    this_range = c2l(frr) & ":" & c2l(frr + max_readers - 1) & "," & _
                    c2l(frr + max_readers + 1) & ":" & c2l(A_T1N + max_readers - 2) & "," & _
                    c2l(A_T2S + max_readers - 2) & ":" & c2l(A_T2S + frr - 1 + max_readers - 1) & "," & _
                    c2l(A_T2S + 2 * (max_readers - 2) + 3) & ":" & c2l(A_T2S + A_T2N + 2 * (max_readers - 2) - 2)
    Range(this_range).Select
    Range(c2l(frr) & A_FIRST_DATA_ROW).Activate
    Selection.NumberFormat = "0.0"
    this_range = "A:" & c2l(A_T2S + A_T2N + 2 * (max_readers - 2) - 1)
    Columns(this_range).Select
    Range("A1").Activate
    Columns(this_range).EntireColumn.AutoFit
    this_range = c2l(frr) & "2:" & c2l(frr + max_readers - 1) & "2," & _
          c2l(A_T2S + frr - 1) & "2:" & c2l(A_T2S + frr + max_readers - 2) & 2
    Range(this_range).Select
    Range(FirstCell(this_range)).Activate
    Selection.NumberFormat = "0"
    Range("A1").Select
    
    FormatAnalysisSheet = True
End Function

Public Function PopulateResultsSheet() As Boolean

    Dim start_sheet As String, start_book As String
    start_book = ActiveWorkbook.Name
    start_sheet = ActiveSheet.Name
    If ActivateCompetitionWorkbook = False Then
        PopulateResultsSheet = False
        Exit Function
    End If
    Sheets(RESULTS_SHEET).Visible = True
    Sheets(RESULTS_SHEET).Select
    
    making_competition_workbook = False
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' load the assignment information
    Sheets(MASTER_ASSIGNMENTS_SHEET).Select
    Dim assignments_range As String, coa_range As String
    assignments_range = c2l(MAS_FIRST_ASSMT_COL) & MAS_FIRST_ASSMT_ROW & ":" & _
                        c2l(MAS_FIRST_ASSMT_COL + target_markers_per_proj - 1) & _
                            MAS_FIRST_ASSMT_ROW + num_projects - 1
    assignments = Range(assignments_range)
    Sheets(EXPERTISE_CROSSWALK_SHEET).Select
    coa_range = c2l(EC_ASSMT_CONF_FIRST_COL) & EC_FIRST_DATA_ROW & ":" & _
                c2l(EC_ASSMT_CONF_FIRST_COL + target_markers_per_proj - 1) & _
                        (EC_ASSMT_CONF_FIRST_COL + num_projects - 1)
    coa_array = Range(coa_range)
    
    'populate the left hand table with:
    '   for each marker, one row for each of their marking assignments,
    '   including the corresponding project #
    Dim expertise_rows() As Variant
    ReDim mn_1m(1 To num_projects * target_markers_per_proj, 1 To 1)    ' marker number column
    ReDim pn_1m(1 To num_projects * target_markers_per_proj, 1 To 1)    ' project number column
    ReDim rn_1m(1 To num_projects * target_markers_per_proj, 1 To 1)    ' reader number column
    ReDim expertise_rows(1 To num_projects * target_markers_per_proj, 1 To 1)
    Dim row_num As Long, j As Long, num_assigned_this_proj As Long
    row_num = 0
    num_projects_scored = 0
    Dim i As Long
    For i = 1 To num_projects
        num_assigned_this_proj = 0
        For j = 1 To target_markers_per_proj
            If Len(assignments(i, j)) > 0 Then  'ignore unassigned marking slots
                row_num = row_num + 1
                pn_1m(row_num, 1) = i
                mn_1m(row_num, 1) = assignments(i, j)
                rn_1m(row_num, 1) = j
                expertise_rows(row_num, 1) = num2LMH(coa_array(i, j))
                num_assigned_this_proj = num_assigned_this_proj + 1
            End If
        Next j
        If num_assigned_this_proj > 0 Then
            num_projects_scored = num_projects_scored + 1
        End If
    Next i
    
    If (row_num < 2) Then
        MsgBox "[PopulateResultsSheet] Not enough assignments to have a competition", vbCritical
        Exit Function
    End If
    
    'put these arrays in the results sheet
    Sheets(RESULTS_SHEET).Select
    Dim pn_1m_range As String, mn_1m_range As String, coa_column_range As String
    Dim rn_1m_range As String
    ' column of project numbers
    pn_1m_range = c2l(R_PROJECT_NUM_COLUMN) & R_FIRST_DATA_ROW & ":" & _
        c2l(R_PROJECT_NUM_COLUMN) & (R_FIRST_DATA_ROW + UBound(pn_1m, 1) - 1)
    Dim Destination As Range
    Set Destination = Range(pn_1m_range)
    Destination.Resize(UBound(pn_1m, 1)).Value = pn_1m
    ' column of marker numbers
    mn_1m_range = c2l(R_MARKER_NUM_COLUMN) & R_FIRST_DATA_ROW & ":" & _
        c2l(R_MARKER_NUM_COLUMN) & (R_FIRST_DATA_ROW + UBound(mn_1m, 1) - 1)
    Set Destination = Range(mn_1m_range)
    Destination.Resize(UBound(mn_1m, 1)).Value = mn_1m
    'column of expertise for that assignment
    coa_column_range = c2l(R_EXPERTISE_LETTERS_COLUMN) & R_FIRST_DATA_ROW & ":" & _
        c2l(R_EXPERTISE_LETTERS_COLUMN) & (R_FIRST_DATA_ROW + UBound(expertise_rows, 1) - 1)
    Set Destination = Range(coa_column_range)
    Destination.Resize(UBound(expertise_rows, 1)).Value = expertise_rows
    ' reader nums
    rn_1m_range = c2l(R_READER_NUM_COLUMN) & R_FIRST_DATA_ROW & ":" & _
                        c2l(R_READER_NUM_COLUMN) & (R_FIRST_DATA_ROW + UBound(rn_1m, 1) - 1)
    Set Destination = Range(rn_1m_range)
    Destination.Resize(UBound(rn_1m, 1)).Value = rn_1m
    
    If simulate_marker_responses Then
        ' read the min and max for each of the scoring criteria
        Sheets(CRITERIA_SHEET).Select
        Dim scores_min() As Variant, scores_max() As Variant, read_range As String
        read_range = c2l(C_FIRST_CRITERIA_MINVALUE_CN) & C_FIRST_CRITERIA_MINVALUE_RN & ":" & _
                        c2l(C_FIRST_CRITERIA_MINVALUE_CN) & C_FIRST_CRITERIA_MINVALUE_RN + num_criteria - 1
        scores_min = Range(read_range)
        read_range = c2l(C_FIRST_CRITERIA_MINVALUE_CN + 1) & C_FIRST_CRITERIA_MINVALUE_RN & ":" & _
                        c2l(C_FIRST_CRITERIA_MINVALUE_CN + 1) & C_FIRST_CRITERIA_MINVALUE_RN + num_criteria - 1
        scores_max = Range(read_range)
        
        ' generate a table of random scores
        Dim random_scores() As Variant
        Dim rownum As Long, k As Long
        ReDim random_scores(1 To num_projects * target_markers_per_proj, 1 To num_criteria)
        rownum = 0
        For j = 1 To num_projects
            For k = 1 To target_markers_per_proj
                rownum = rownum + 1
                If IsEmpty(pn_1m(rownum, 1)) Or IsEmpty(rn_1m(rownum, 1)) Then
                    ' don't put any data if the project number or reader number is blank
                Else
                    If assignments(pn_1m(rownum, 1), rn_1m(rownum, 1)) > 0 Then
                        ' only create scores if there is an assignment
                        For i = 1 To num_criteria
                            random_scores(rownum, i) = scores_min(i, 1) + (scores_max(i, 1) - scores_min(i, 1)) * Rnd
                        Next i
                    End If
                End If
            Next k
        Next j
        ' write these random scores to the results table
        Sheets(RESULTS_SHEET).Select
        Dim write_range As String
        write_range = c2l(R_FIRST_RAW_COLUMN) & R_FIRST_DATA_ROW
        Set Destination = Range(write_range)
        Set Destination = Destination.Resize(UBound(random_scores, 1), UBound(random_scores, 2))
        Destination.Value = random_scores
        ' not sure why this is here since we just wrote raw numbers to the sheet ...
'        ConvertRangeToText write_range & ":" & c2l(R_FIRST_RAW_COLUMN + num_criteria - 1) & _
'                            R_FIRST_DATA_ROW + num_projects * target_markers_per_proj - 1
        AddMessage "Added random scores to " & RESULTS_SHEET
        Erase random_scores, scores_min, scores_max
    End If
    
    SortResultRawScoresTable row_num
            
    'convert the various lookup formulas in some columns to their text result:
    ConvertCellsDownFromFormula2Text R_FIRST_DATA_ROW, "B", num_projects * target_markers_per_proj    'marker name
    ConvertCellsDownFromFormula2Text R_FIRST_DATA_ROW, "F", num_projects * target_markers_per_proj    'project name
    ConvertCellsDownFromFormula2Text R_FIRST_DATA_ROW, c2l(R_T2S + num_criteria - 2 + 1), _
                                    num_projects * target_markers_per_proj 'Marker name
    ConvertCellsDownFromFormula2Text R_FIRST_DATA_ROW, c2l(R_T4S + 2 * (num_criteria - 2) + 1), _
                                    num_projects * target_markers_per_proj 'project name
    
    ' the criteria names and scoring ranges also need to be converted to text
    ConvertRangeToText "H2:" & c2l(8 + num_criteria - 1) & "4"
        
    ' format the sheet
    MergeHorizontal 1, R_T1S, R_T1S + 5
    MergeHorizontal 1, R_T2S + num_criteria - 2, R_T2S + num_criteria
    MergeHorizontal 1, R_T3S + num_criteria - 2, R_T3S + num_criteria - 2 + 7
    MergeHorizontal 1, R_T4S + 2 * (num_criteria - 2), R_T4S + 3 * (num_criteria - 2) + 4
    Cells.Select
    Range("A1").Activate
    Cells.EntireColumn.AutoFit
    Dim rng As String   '"4:4"
    rng = (R_FIRST_DATA_ROW - 1) & ":" & (R_FIRST_DATA_ROW - 1)
    Rows(rng).Select
    Rows(rng).EntireRow.AutoFit
    
    Range("A1").Select
    Range("A1").Activate

    SortResultsFinalProjectTable

    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    
    Erase mn_1m, rn_1m, expertise_rows
    
    PopulateResultsSheet = True
    
End Function

Public Function Email2Text(str_in As String, max_char As Long) As String
    If max_char < 1 Then
        MsgBox "[Email2Text] maximum character length input as " & max_char & " ???", vbCritical
        Exit Function
    End If
    Dim i As Long, one_char As String
    For i = 1 To Len(str_in)
        one_char = Mid(str_in, i, 1)
        Select Case one_char
        Case "@"
            Email2Text = Email2Text & "_at_"
        Case Else
            If ((one_char >= "a") And (one_char <= "z")) Or _
                ((one_char >= "A") And (one_char <= "Z")) Or _
                ((one_char >= "0") And (one_char <= "9")) Then
                ' only keep the alphanumeric
                Email2Text = Email2Text & one_char
            End If
        End Select
    Next i
    If (Len(Email2Text) > max_char) Then
        Email2Text = Left(Email2Text, max_char)
    End If
End Function

Public Function FreeArrays() As Boolean

    Erase pxm_table, pxm_at_start, pn_array, mn_array, coa_array, mentor_column
    Erase competition_COIs, ss_marker_col, ss_project_col, xlmh_per_marker, xlmh_per_project
    Erase assignments, n_assigned2project, n_assigned2marker, markers_table, projects_table
    Erase messages, comments, general_comments
    Erase assignment_failed_for_this_proj, marker_scores, pn_1m, rn_1m, mn_1m
    Erase competition_scores
    
    If num_messages > 0 Then    'report any messages before freeing the messages buffer
        ReportMessages
    End If
    Erase messages
    
    num_competition_scores = 0
    
    FreeArrays = True

End Function

Public Function ExpandKeywordsSheet() As Boolean
    Sheets(KEYWORDS_SHEET).Select
    Sheets(KEYWORDS_SHEET).Activate
    CopyAndTransposeCellsGrey 1, 2, KW_WEIGHTS_COL, num_keywords + 1, KW_TRANSPOSE_COPY_CELL
    ExpandKeywordsSheet = True
End Function

Public Function ExpandCriteriaSheet() As Boolean
    Sheets(CRITERIA_SHEET).Select
    Sheets(CRITERIA_SHEET).Activate
    CopyAndTransposeCellsGrey 1, 2, 5, num_criteria + 1, C_TRANSPOSE_COPY_CELL
    ExpandCriteriaSheet = True
End Function

Public Function ExpandMarkerExpertiseSheetTables()
    ' this sheet has a formula that depends on the Project Keywords sheet, so this function
    ' should be called after expanding the Project Keywords sheet.
    Sheets(MARKER_EXPERTISE_SHEET).Select
    Sheets(MARKER_EXPERTISE_SHEET).Activate
    Dim num_cols2add As Long, num_rows2add As Long
    num_cols2add = num_keywords - 2
    num_rows2add = num_markers - 2
    InsertAndExpandRight ME_FIRST_DATA_COL, 1, _
                        ME_FIRST_DATA_ROW + 1, num_cols2add  ' first the left hand table
    InsertAndExpandRight ME_FIRST_DATA_COL + ME_NUM_COLUMNS2SECOND_TABLE + num_cols2add, 1, _
                            ME_FIRST_DATA_ROW + 1, num_cols2add 'now do the table to the right as it depends on the left hand table
    AddErrorLMHCheckingFormula ME_FIRST_DATA_COL, ME_FIRST_DATA_ROW
    InsertAndExpandDown 1, ME_FIRST_DATA_ROW, _
            ME_FIRST_DATA_COL + num_cols2add + ME_NUM_COLUMNS2SECOND_TABLE + num_cols2add + 2, num_rows2add
' now some formatting
    Dim rng As String
    'size the columns in the two tables
'    Columns("B:B").ColumnWidth = 50         ' the project names column
    Columns("B:B").EntireColumn.AutoFit         ' the project names
    rng = c2l(ME_FIRST_DATA_COL) & ":" & c2l(ME_FIRST_DATA_COL + num_keywords - 1) & "," & _
          c2l(ME_FIRST_DATA_COL + num_keywords + ME_COLUMNS_BETWEEN_DATA_TABLES) & ":" & _
          c2l(ME_FIRST_DATA_COL + 2 * num_keywords + ME_COLUMNS_BETWEEN_DATA_TABLES - 1)
    Range(rng).Select
    Selection.ColumnWidth = 3.5
    ' size the error checking column, first oversized, and then autofit
    rng = c2l(ME_FIRST_DATA_COL + num_keywords)
    Range(rng & ":" & rng).Select
    Selection.ColumnWidth = 30
    Columns(rng & ":" & rng).EntireColumn.AutoFit
    ' size down the divider column
    rng = c2l(ME_FIRST_DATA_COL + num_keywords + 1)
    Range(rng & ":" & rng).Select
    Selection.ColumnWidth = 1#
    ' now autofit the row heights
    Cells.Select
    Cells.EntireRow.AutoFit
    Range(c2l(ME_FIRST_DATA_COL) & ME_FIRST_DATA_ROW).Select

End Function

Public Function AddErrorLMHCheckingFormula(col As Long, rn As Long) As Boolean
    'add the error checking formula for the rows of the first column
    Dim i As Long, colnum As Long, cr As String, formula_str As String
    
    formula_str = "=IF(OR("
    For i = 1 To num_keywords
        If i <> 1 Then
            formula_str = formula_str & ","
        End If
        cr = c2l(col + i - 1) & rn
        formula_str = formula_str & "AND(" & cr & "<>""L""," & cr & "<>""M""," & cr & "<>""H"")"
    Next i
    formula_str = formula_str & "),""Enter L, M or H in each cell"","""")"
    Range(c2l(col + num_keywords) & rn).Value = formula_str
    
    AddErrorLMHCheckingFormula = True
End Function

Sub MakeCompetitionWorkbook_sub()
    MakeCompetitionWorkbook
End Sub
Public Function MakeCompetitionWorkbook()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    making_competition_workbook = True
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' make sure the sheets are visible
    Dim sheet_names() As Variant
    sheet_names = Array(COMPETITION_PARAMETERS_SHEET, CRITERIA_SHEET, PROJECTS_SHEET, MARKERS_SHEET, _
        KEYWORDS_SHEET, PROJECT_KEYWORDS_SHEET, MARKER_EXPERTISE_SHEET, PROJECT_X_MARKER_SHEET, _
        EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, EXPERTISE_CROSSWALK_SHEET, MASTER_ASSIGNMENTS_SHEET, _
        RESULTS_SHEET, ANALYSIS_SHEET, EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, _
        MARKER_PROJECT_EXPERTISE_TEMPLATE, MARKER_KEYWORD_EXPERTISE_TEMPLATE, MARKER_SCORING_TEMPLATE, _
        SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE, SCORES_AND_COMMENTS_TEMPLATE_SHEET, PROJECT_COMMENTS_SHEET)
    HideOrShowSheets sheet_names, True
    
    ' now select and export them
    Sheets(Array(COMPETITION_PARAMETERS_SHEET, CRITERIA_SHEET, PROJECTS_SHEET, MARKERS_SHEET, _
        KEYWORDS_SHEET, PROJECT_KEYWORDS_SHEET, MARKER_EXPERTISE_SHEET, PROJECT_X_MARKER_SHEET, _
        EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, EXPERTISE_CROSSWALK_SHEET, MASTER_ASSIGNMENTS_SHEET, _
        RESULTS_SHEET, ANALYSIS_SHEET, EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, _
        MARKER_PROJECT_EXPERTISE_TEMPLATE, MARKER_KEYWORD_EXPERTISE_TEMPLATE, MARKER_SCORING_TEMPLATE, _
        SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE, SCORES_AND_COMMENTS_TEMPLATE_SHEET, PROJECT_COMMENTS_SHEET)).Select
    Sheets(PROJECT_KEYWORDS_SHEET).Activate
    
    Sheets(Array(COMPETITION_PARAMETERS_SHEET, CRITERIA_SHEET, PROJECTS_SHEET, MARKERS_SHEET, _
        KEYWORDS_SHEET, PROJECT_KEYWORDS_SHEET, MARKER_EXPERTISE_SHEET, PROJECT_X_MARKER_SHEET, _
        EXPERTISE_CROSSWALK_SHEET, MASTER_ASSIGNMENTS_SHEET, RESULTS_SHEET, ANALYSIS_SHEET, _
        EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, _
        MARKER_PROJECT_EXPERTISE_TEMPLATE, MARKER_KEYWORD_EXPERTISE_TEMPLATE, MARKER_SCORING_TEMPLATE, _
        SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE, SCORES_AND_COMMENTS_TEMPLATE_SHEET, PROJECT_COMMENTS_SHEET)).Copy
    
    cwb = ActiveWorkbook.Name
    
    ' expand the sheets to the suitable size
    ExpandKeywordsSheet
    ExpandCriteriaSheet
    ExpandProjectsKeywordsSheetTables
    ExpandMarkerExpertiseSheetTables
    ExpandProjectsXMarkersTable
    ExpandExpertiseCrosswalk
    ExpandAssignmentsMaster
    ExpandResultsSheet
    ExpandExpertiseTemplates
    ExpandMarkerScoresheetTemplate
    ExpandScoresWithCommentsSheet
    ExpandProjectCommentsSheet
    
    'HIDE the sheets that are templates (i.e., the user should not interact with them)
    Sheets(Array(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, _
        EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, MARKER_PROJECT_EXPERTISE_TEMPLATE, MARKER_KEYWORD_EXPERTISE_TEMPLATE, _
        MARKER_SCORING_TEMPLATE, SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE, SCORES_AND_COMMENTS_TEMPLATE_SHEET, _
        PROJECT_COMMENTS_SHEET)).Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets(PROJECTS_SHEET).Activate
    
    ' shrink the tab slider on this workbook to show more tabs
    ActiveWindow.TabRatio = 0.8
    Sheets(COMPETITION_PARAMETERS_SHEET).Activate
'    ActiveWindow.ScrollWorkbookTabs Sheets:=-2

    ' ask the user for the name to give competition book and save the file
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim comp_book As String, initial_name As String
    initial_name = COMPETITION_WORKBOOK_DEFAULT_NAME
    comp_book = GetFileSaveasName("Specify the name for this competition book", initial_name)
    If Len(comp_book) > 0 Then
        ActiveWorkbook.SaveAs Filename:=comp_book, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        cwb = ActiveWorkbook.Name
    End If
    
    ' go back to the macro book
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(MACROS_SHEET).Activate
    
    'hide the sheets that are not input sheets
    Sheets(MACROS_SHEET).Select
    Sheets(MACROS_SHEET).Activate
    sheet_names = Array(COMPETITION_PARAMETERS_SHEET, CRITERIA_SHEET, PROJECTS_SHEET, _
        MARKERS_SHEET, KEYWORDS_SHEET, PROJECT_KEYWORDS_SHEET, MARKER_EXPERTISE_SHEET, PROJECT_X_MARKER_SHEET, _
        EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, EXPERTISE_CROSSWALK_SHEET, MASTER_ASSIGNMENTS_SHEET, _
        RESULTS_SHEET, ANALYSIS_SHEET, EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, _
        MARKER_PROJECT_EXPERTISE_TEMPLATE, MARKER_KEYWORD_EXPERTISE_TEMPLATE, MARKER_SCORING_TEMPLATE, _
        SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE, SCORES_AND_COMMENTS_TEMPLATE_SHEET, PROJECT_COMMENTS_SHEET)
    '  hide them
    HideOrShowSheets sheet_names, False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MakeCompetitionWorkbook = True
End Function

Public Function ExpandProjectsKeywordsSheetTables()
    Sheets(PROJECT_KEYWORDS_SHEET).Select
    Sheets(PROJECT_KEYWORDS_SHEET).Activate
    Dim num_cols2add As Long, num_rows2add As Long
    num_cols2add = num_keywords - 2
    num_rows2add = num_projects - 2
    ' first the left hand table
    ' includes the formula below the table (that count the number of expertise by keyword)
    InsertAndExpandRight PK_FIRST_DATA_COL, 1, PK_FIRST_DATA_ROW + 2, num_cols2add
    AddErrorLMHCheckingFormula PK_FIRST_DATA_COL, PK_FIRST_DATA_ROW
    'now do the table to the right as it depends on the left hand table
    InsertAndExpandRight 8 + num_cols2add, 1, PK_FIRST_DATA_ROW + 1, num_cols2add
    ' fill out the table rows
    InsertAndExpandDown 1, PK_FIRST_DATA_ROW, 10 + 2 * num_cols2add, num_rows2add
    ' now some formatting
    Dim rng As String
    'size the columns in the two tables
    Columns("B:B").ColumnWidth = 50         ' the project names
    rng = c2l(PK_FIRST_DATA_COL) & ":" & c2l(PK_FIRST_DATA_COL + num_keywords - 1) & "," & _
          c2l(PK_FIRST_DATA_COL + num_keywords + PK_COLUMNS_BETWEEN_DATA_TABLES) & ":" & _
          c2l(PK_FIRST_DATA_COL + 2 * num_keywords + PK_COLUMNS_BETWEEN_DATA_TABLES - 1)
    Range(rng).Select
    Selection.ColumnWidth = 3.5
    ' size the error checking column, first oversized, and then autofit
    rng = c2l(PK_FIRST_DATA_COL + num_keywords)
    Range(rng & ":" & rng).Select
    Selection.ColumnWidth = 30
    Columns(rng & ":" & rng).EntireColumn.AutoFit
    ' size down the divider column
    rng = c2l(PK_FIRST_DATA_COL + num_keywords + 1)
    Range(rng & ":" & rng).Select
    Selection.ColumnWidth = 1#
    ' now autofit the row heights
    Cells.Select
    Cells.EntireRow.AutoFit
    Range(c2l(PK_FIRST_DATA_COL) & PK_FIRST_DATA_ROW).Select
                        
End Function
'
Public Function ExpandProjectsXMarkersTable()
    Sheets(PROJECT_X_MARKER_SHEET).Select
    Sheets(PROJECT_X_MARKER_SHEET).Activate
    Dim num_cols2add As Long, num_rows2add As Long
    num_cols2add = num_markers - 2
    num_rows2add = num_projects - 2
    InsertAndExpandRight PXM_FIRST_DATA_COL, 1, PXM_FIRST_DATA_ROW + 1, num_cols2add ' first the left hand table
    InsertAndExpandDown 1, PXM_FIRST_DATA_ROW, PXM_FIRST_DATA_COL + 1 + num_cols2add, num_rows2add
  
End Function

Public Function ExpandAssignmentsMaster()
    If (num_markers < 5) And (num_projects < 5) Then
        ' need at least 5 to avoid smearing the right hand table of statistics
        MsgBox "Too few markers and/or projects - need at least 5 to grow " & MASTER_ASSIGNMENTS_SHEET & _
            " with current software", vbCritical
        Exit Function
    End If

    Sheets(MASTER_ASSIGNMENTS_SHEET).Select
    Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
    Dim num_cols2add As Long, num_rows2add As Long
    num_cols2add = target_markers_per_proj - 2
    num_rows2add = num_projects - 2
    
    ' first the table of which markers are assigned to which project
    InsertAndExpandDown 1, MAS_FIRST_ASSMT_ROW, 9, num_rows2add
    ' first the marker numbers
    InsertAndExpandRight MAS_FIRST_ASSMT_COL, 1, MAS_FIRST_ASSMT_ROW + 1 + num_rows2add, num_cols2add
    ' update the formula for checking if the assignments conflict with the marker
    Dim form_str As String, vlookup_range As String, form_cell As String, i As Long
    form_str = "=IF(AND(" & c2l(MAS_MENTOR_NUM_COL) & MAS_FIRST_ASSMT_ROW & ">0,OR(" & _
                    c2l(MAS_FIRST_ASSMT_COL - 2) & MAS_FIRST_ASSMT_ROW & "=" & _
                    c2l(MAS_FIRST_ASSMT_COL) & MAS_FIRST_ASSMT_ROW
    For i = 2 To target_markers_per_proj
        form_str = form_str & "," & _
                    c2l(MAS_FIRST_ASSMT_COL - 2) & MAS_FIRST_ASSMT_ROW & "=" & _
                    c2l(MAS_FIRST_ASSMT_COL + i - 1) & MAS_FIRST_ASSMT_ROW
    Next i
    form_str = form_str & ")),""XX"","""")"
    form_cell = c2l(MAS_FIRST_ASSMT_COL - 1) & MAS_FIRST_ASSMT_ROW
    PutFormulaAndDragDown form_cell, form_str, num_projects - 1
    
    'then the marker names
    ' we need to update the formula that looks up the names of the markers
    vlookup_range = "$" & c2l(MAS_FIRST_ASSMT_COL + 2 * target_markers_per_proj + 1) & ":" & _
                    "$" & c2l(MAS_FIRST_ASSMT_COL + 2 * target_markers_per_proj + 2)
    form_str = "=IF(ISNA(VLOOKUP(F3," & vlookup_range & ",2,FALSE)),"""",VLOOKUP(F3," & vlookup_range & ",2,FALSE))"
    ' put it in the top left cell
    form_cell = c2l(MAS_FIRST_ASSMT_COL + target_markers_per_proj) & MAS_FIRST_ASSMT_ROW
    PutFormulaAndDragDown form_cell, form_str, num_projects - 1
    
    ' now grow the table of marker names for all the assignments per project
    InsertAndExpandRight MAS_FIRST_ASSMT_COL + target_markers_per_proj, 1, _
                         MAS_FIRST_ASSMT_ROW + 1 + num_rows2add, num_cols2add
    ' clear the duplicated text in the first row
    Dim clear_range As String
    clear_range = c2l(MAS_FIRST_ASSMT_COL + 1) & 1 & ":" & _
                  c2l(MAS_FIRST_ASSMT_COL + target_markers_per_proj - 1) & 1
    Range(clear_range).Select
    Range(FirstCell(clear_range)).Activate
    Selection.ClearContents
    clear_range = c2l(MAS_FIRST_ASSMT_COL + target_markers_per_proj + 1) & 1 & ":" & _
                  c2l(MAS_FIRST_ASSMT_COL + 2 * target_markers_per_proj - 1) & 1
    Range(clear_range).Select
    Range(FirstCell(clear_range)).Activate
    Selection.ClearContents
    
    AutofitOneColumn (3)    ' Organization Name
    ' now the table of statistics by marker
    num_rows2add = num_markers - 2
    InsertAndExpandDown 6 + 2 * target_markers_per_proj + 1, MAS_FIRST_ASSMT_ROW, 6, num_rows2add
    AutofitOneColumn (6 + 2 * target_markers_per_proj + 2) ' marker names
    AutofitOneColumn (18 + 2 * num_cols2add) 'labels for final table
    
    ' fix up a formula in the last (small) table
    Dim cn As Long
    cn = MAS_FIRST_ASSMT_COL + 2 * target_markers_per_proj + 1 + 6 + 1 + 1
    form_str = "=COUNTIF(" & c2l(MAS_FIRST_ASSMT_COL) & ":" & _
                             c2l(MAS_FIRST_ASSMT_COL + target_markers_per_proj - 1) & ",""<>""&"""")-" & _
                             c2l(cn) & "4-1"
    Range(c2l(cn) & 7).Value = form_str
        
    ' final formatting
    ResizeToNarrowColumn (6 + 2 * target_markers_per_proj)
    ResizeToNarrowColumn (17 + 2 * num_cols2add)

End Function

Public Function ExpandExpertiseCrosswalk()
    Sheets(EXPERTISE_CROSSWALK_SHEET).Select
    Sheets(EXPERTISE_CROSSWALK_SHEET).Activate
    
    Dim clear_range As String
    Dim project_rows2add As Long, marker_cols2add As Long, assmt_cols2add As Long
    marker_cols2add = num_markers - 2
    assmt_cols2add = target_markers_per_proj - 2
    project_rows2add = num_projects - 2
    InsertAndExpandRight EC_ASSMT_CONF_FIRST_COL, 1, 8, assmt_cols2add ' first the assignment columns
    ' clean up the repeated entries for the column titles created by expanding the table
    clear_range = c2l(EC_ASSMT_CONF_FIRST_COL + 1) & 4 & ":" & c2l(EC_ASSMT_CONF_FIRST_COL + assmt_cols2add + 1) & 4
    Range(clear_range).Select
    Range(FirstCell(clear_range)).Activate
    Selection.ClearContents
    InsertAndExpandRight EC_ASSMT_CONF_FIRST_COL + 2 + assmt_cols2add, 1, 8, assmt_cols2add ' now the marker names
    clear_range = c2l(EC_ASSMT_CONF_FIRST_COL + 1 + target_markers_per_proj) & 5 & ":" & c2l(EC_ASSMT_CONF_FIRST_COL + 2 * target_markers_per_proj - 1) & 5
    Range(clear_range).Select
    Range(FirstCell(clear_range)).Activate
    Selection.ClearContents
    marker_cols2add = num_markers - 2
    InsertAndExpandRight 10 + 2 + 3 + 2 * assmt_cols2add, 1, 8, marker_cols2add ' expertise by markers table
    InsertAndExpandDown 1, 7, 10 + 2 * target_markers_per_proj + num_markers, project_rows2add
    
    ' consistent widths for the assignment and confidence levels
    Dim rng As String
    rng = c2l(EC_ASSMT_CONF_FIRST_COL) & ":" & _
            c2l(EC_ASSMT_CONF_FIRST_COL + 2 * target_markers_per_proj - 1)
    Columns(rng).Select
    Range(c2l(EC_ASSMT_CONF_FIRST_COL) & EC_FIRST_DATA_ROW).Activate
    Selection.ColumnWidth = 5
    rng = c2l(EC_XLMH_CONF_PER_PROJECT_COL) & ":" & c2l(EC_XLMH_CONF_PER_PROJECT_COL + 3)
    Columns(rng).Select
    Range(c2l(EC_XLMH_CONF_PER_PROJECT_COL) & EC_FIRST_DATA_ROW).Activate
    Selection.ColumnWidth = 6

    ' merge the cells for expertise confidence label
    MergeVertical 9, 2, 6
 
End Function

Public Function InsertAndDragRight(dest_cell As String, form_str As String, num_cells As Long) As Boolean
    ' insert a formula/text  into a cell, and drag it so that it fills num_cells
    Range(dest_cell).Value = form_str
    Range(dest_cell).Select
    Dim fill_range As String
    fill_range = dest_cell & ":" & c2l(Range(dest_cell).Column + num_cells - 1) & Range(dest_cell).row
    Selection.AutoFill Destination:=Range(fill_range), Type:=xlFillDefault
    InsertAndDragRight = True
End Function

Public Function ExpandResultsSheet()
    Sheets(RESULTS_SHEET).Select
    Sheets(RESULTS_SHEET).Activate
    Dim criteria_cols2add As Long, pxm_rows2add As Long, marker_rows2add As Long, project_rows2add As Long
    criteria_cols2add = num_criteria - 2
    project_rows2add = num_projects - 2
    marker_rows2add = num_markers - 2
    pxm_rows2add = num_projects * target_markers_per_proj - 2
    
    ' fill out the table with all the columns for the criteria
    ' first the table for the raw pxm criteria scores
    InsertAndExpandRight R_FIRST_RAW_COLUMN, 1, R_FIRST_DATA_ROW + 1, criteria_cols2add
    ' update the formulas in the marker normalization factor table
    Dim form_str As String, range1 As String, range2 As String, range3 As String, range4 As String
    Dim destn As String
    ' # of scores used =SUMIF(A:A,O5,L:L)
    range1 = c2l(R_T2S + criteria_cols2add) & R_FIRST_DATA_ROW
    range2 = c2l(R_T1N + criteria_cols2add - 1) & ":" & c2l(R_T1N + criteria_cols2add - 1)
    form_str = "=SUMIF(A:A," & range1 & "," & range2 & ")"
    destn = c2l(R_T2S + criteria_cols2add + 2) & R_FIRST_DATA_ROW
    Range(destn).Value = form_str
    'next the count on the total of all the scores from that marker
    ' Total of Scores =IF(Q5>0,SUMIF($A:$A,O5,M:M),"")
    range1 = c2l(R_T2S + criteria_cols2add) & R_FIRST_DATA_ROW
    range2 = c2l(R_T2S + criteria_cols2add + 2) & R_FIRST_DATA_ROW
    range3 = c2l(R_T1N + criteria_cols2add) & ":" & c2l(R_T1N + criteria_cols2add)
    form_str = "=IF(" & range1 & ">0,SUMIF($A:$A," & range1 & "," & range3 & "),"""")"
    destn = c2l(R_T2S + criteria_cols2add + 3) & R_FIRST_DATA_ROW
    Range(destn).Value = form_str
    
    ' now the normalized pxm criteria scores
    InsertAndExpandRight R_T3S + 6 + criteria_cols2add, 1, R_FIRST_DATA_ROW + 1, criteria_cols2add
    ' Criteria 1 score (normalized) =IF(AND(LEN(H5)>0,$J5=1),H5*VLOOKUP($W5,$O:$T,6,FALSE),"")
    range1 = "$" & c2l(R_T3S + criteria_cols2add + 1) & R_FIRST_DATA_ROW
    range2 = "$" & c2l(R_T2S + criteria_cols2add) & ":$" & c2l(R_T2S + R_T2N - 1 + criteria_cols2add)
    form_str = "=IF(AND(LEN(H" & R_FIRST_DATA_ROW & ")>0,$" & c2l(R_T1N + criteria_cols2add - 3) & _
               R_FIRST_DATA_ROW & "=1),H" & R_FIRST_DATA_ROW & "*VLOOKUP(" & range1 & "," & range2 & "," & _
                R_T2N & ",FALSE),"""")"
    destn = c2l(R_T3S + criteria_cols2add + 6) & R_FIRST_DATA_ROW
    InsertAndDragRight destn, form_str, num_criteria
    
    ' now the normalized project scores by criteria
    InsertAndExpandRight R_T4S + 2 * criteria_cols2add + 2, 1, R_FIRST_DATA_ROW + 1, criteria_cols2add
    'Project Name: =IF(AH5>0,VLOOKUP(AH5,E:F,2,FALSE),"")
    range1 = c2l(R_T4S + 2 * criteria_cols2add) & R_FIRST_DATA_ROW
    form_str = "=IF(" & range1 & ">0,VLOOKUP(" & range1 & ",E:F,2,FALSE),"""")"
    destn = c2l(R_T4S + 2 * criteria_cols2add + 1) & R_FIRST_DATA_ROW
    Range(destn).Value = form_str
    'for the table 4 project criteria =IF(SUMIF($T:$T,$AH5,Z:Z)=0,"",SUMIF($T:$T,$AH5,Z:Z)/$AN5)
    range1 = "$" & c2l(R_T3S + criteria_cols2add) & ":$" & c2l(R_T3S + criteria_cols2add)
    range2 = "$" & c2l(R_T4S + 2 * criteria_cols2add) & R_FIRST_DATA_ROW
    range3 = c2l(R_T3S + criteria_cols2add + 6) & ":" & c2l(R_T3S + criteria_cols2add + 6)
    range4 = c2l(R_T4S + 3 * criteria_cols2add + 4) & R_FIRST_DATA_ROW
    form_str = "=IF(SUMIF(" & range1 & "," & range2 & "," & range3 & ")=0,""""," & _
                   "SUMIF(" & range1 & "," & range2 & "," & range3 & ")/$" & range4 & ")"
    destn = c2l(R_T4S + 2 * criteria_cols2add + 2) & R_FIRST_DATA_ROW
    InsertAndDragRight destn, form_str, num_criteria
    
    ' now fill out the rows of the different tables
    InsertAndExpandDown 1, R_FIRST_DATA_ROW, R_T1N + criteria_cols2add, pxm_rows2add
    InsertAndExpandDown R_T2S + criteria_cols2add, R_FIRST_DATA_ROW, R_T2N, marker_rows2add
    InsertAndExpandDown R_T3S + criteria_cols2add, R_FIRST_DATA_ROW, R_T3N + criteria_cols2add + 1, pxm_rows2add
    InsertAndExpandDown R_T4S + 2 * criteria_cols2add, R_FIRST_DATA_ROW, R_T4N + criteria_cols2add, project_rows2add
    
    ' some prettying up
    Dim cn As Long
    cn = 1
    MergeVertical cn + 0, 2, 4
    MergeVertical cn + 1, 2, 4
    MergeVertical cn + 2, 2, 4
    MergeVertical cn + 3, 2, 4
    MergeVertical cn + 4, 2, 4
    cn = R_T1N + criteria_cols2add
    MergeVertical cn - 3, 2, 4
    MergeVertical cn - 2, 2, 4
    MergeVertical cn - 1, 2, 4
    MergeVertical cn - 0, 2, 4
    cn = R_T2S + criteria_cols2add
    MergeVertical cn + 0, 2, 4
    MergeVertical cn + 1, 2, 4
    MergeVertical cn + 2, 2, 4
    MergeVertical cn + 3, 2, 4
    MergeVertical cn + 5, 2, 4
    cn = R_T2S + 4 + criteria_cols2add
    Columns(c2l(cn) & ":" & c2l(cn + 1)).ColumnWidth = 8
    cn = R_T3S + criteria_cols2add
    MergeVertical cn + 0, 2, 4
    MergeVertical cn + 1, 2, 4
    MergeVertical cn + 2, 2, 4
    MergeVertical cn + 3, 2, 4
    MergeVertical cn + 4, 2, 4
    cn = R_T4S + 3 * criteria_cols2add
    MergeVertical cn + 4, 2, 4
    MergeVertical cn + 5, 2, 4
    AutofitOneColumn (R_T4S + 2 * criteria_cols2add + 1) ' project name in final score table
    ' make the gap columns between the tables narrow
    ResizeToNarrowColumn (R_T1N + num_criteria - 2 + 1)
    ResizeToNarrowColumn (R_T2S + num_criteria - 2 + R_T2N)
    ResizeToNarrowColumn (R_T3S + 2 * (num_criteria - 2) + R_T3N)

End Function

Public Function ExpandAnalysisSheet(max_readers As Long) As Boolean
    If DoesSheetHaveCosetComment(ANALYSIS_SHEET, "A1") Then
        Exit Function ' make sure we don't do this twice
    End If
    Sheets(ANALYSIS_SHEET).Select
    Dim reader_cols2add As Long, project_rows2add As Long
    project_rows2add = num_projects - 2
    
    ' fill out each table with the suitable number of reader columns and update relevant formulae
    InsertAndExpandRight A_FIRST_RAW_READER_COLUMN, 1, A_FIRST_DATA_ROW + 1, max_readers - 2
    InsertAndExpandRight A_T2S + max_readers - 2, 1, A_FIRST_DATA_ROW + 1, max_readers - 2
    
    ' now fill out the rows for this multi-table
    InsertAndExpandDown 1, A_FIRST_DATA_ROW, A_T1N + A_T2N + 2 * (max_readers - 2), project_rows2add
    ' make the rank columns that frame the tables go from the number of projects down to 1
    Dim form_str As String
    form_str = "=" & "A" & A_FIRST_DATA_ROW & "-1"
    PutFormulaAndDragDown "A" & A_FIRST_DATA_ROW + 1, form_str, num_projects - 2
    ConvertCellsDownFromFormula2Text A_FIRST_DATA_ROW + 1, "A", num_projects
    PutFormulaAndDragDown c2l(A_T2S + A_T2N + 2 * (max_readers - 2) - 1) & A_FIRST_DATA_ROW + 1, form_str, num_projects - 2
    ConvertCellsDownFromFormula2Text A_FIRST_DATA_ROW, c2l(A_T2S + A_T2N + 2 * (max_readers - 2) - 1), num_projects
    
    FlagSheetExpanded
    ExpandAnalysisSheet = True
End Function

Public Function ExpandExpertiseTemplates()

    Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Select
    Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Activate

    Dim project_rows2add As Long, keyword_rows2add As Long
    
    project_rows2add = num_projects - 2
    InsertAndExpandDown 1, MPET_FIRST_DATA_ROW, 7, project_rows2add
    
    Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Select
    Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Activate
    keyword_rows2add = num_keywords - 2
    InsertAndExpandDown 1, MKET_FIRST_DATA_ROW, 3, keyword_rows2add
    
End Function

Public Function ExpandMarkerScoresheetTemplate()
    ' NOTE: this function only expand the columns of the sheet for the template
    '       adding the rows for each assignment is done on a per marker basis later
    If DefineGlobals = False Then
        Exit Function
    End If
    
    Sheets(MARKER_SCORING_TEMPLATE).Select
    Sheets(MARKER_SCORING_TEMPLATE).Activate
    Dim criteria_cols2add As Long
    criteria_cols2add = num_criteria - 2
    
    ' first the table for the raw scores
    InsertAndExpandRight MST_FIRST_SCORING_COL, 1, 15, criteria_cols2add
    ' now the table for the normalized scores
    InsertAndExpandRight MST_FIRST_SCORING_COL + num_criteria + 1, 1, 11, criteria_cols2add
    
    ' clean up the repeated entries for the column titles created by expanding the table
    Dim clear_range As String
    clear_range = c2l(MST_FIRST_SCORING_COL + 1) & 3 & ":" & c2l(MST_FIRST_SCORING_COL + criteria_cols2add + 1) & 3
    Range(clear_range).Select
    Range(FirstCell(clear_range)).Activate
    Selection.ClearContents
    
    clear_range = c2l(MST_FIRST_SCORING_COL + 1) & 8 & ":" & c2l(MST_FIRST_SCORING_COL + criteria_cols2add + 1) & 8
    Range(clear_range).Select
    Range(FirstCell(clear_range)).Activate
    Selection.ClearContents
    
    clear_range = c2l(MST_FIRST_SCORING_COL + num_criteria + 2) & 3 & ":" & _
                  c2l(MST_FIRST_SCORING_COL + num_criteria + 2 + criteria_cols2add) & 3
    Range(clear_range).Select
    Range(FirstCell(clear_range)).Activate
    Selection.ClearContents

    clear_range = c2l(MST_FIRST_SCORING_COL + num_criteria + 2) & 8 & ":" & _
                  c2l(MST_FIRST_SCORING_COL + num_criteria + 2 + criteria_cols2add) & 8
    Range(clear_range).Select
    Range(FirstCell(clear_range)).Activate
    Selection.ClearContents
End Function

Public Function ExpandScoresAndCommentsInstructionsSheet(n_proj As Long) As Boolean
    Sheets(SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE).Select
    Sheets(SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE).Activate

    InsertAndExpandDown 1, SCI_PROJECT_COUNT_ROW + 2, 3, n_proj
    ExpandScoresAndCommentsSheet = True
End Function
    
Public Function ExpandProjectCommentsSheet()
    Sheets(PROJECT_COMMENTS_SHEET).Select
    Sheets(PROJECT_COMMENTS_SHEET).Activate
    
    Const THANK_YOU_TEXT  As String = "This concludes the feedback on your project."
    Dim i As Long, next_row As Long
    next_row = PC_FIRST_CRITERIA_COMMENTS_ROW - 2
    For i = 2 To num_criteria
        ' copy the preceding comment section and paste below
        Rows(next_row & ":" & (next_row + PC_ROWS_PER_CRITERIA - 1)).Select
        Selection.Copy
        next_row = next_row + PC_ROWS_PER_CRITERIA
        Range("A" & next_row).Select
        ActiveSheet.Paste
         ' put in a formula for the criteria number (increment from the one above)
        Cells(next_row, 2).Value = "=" & "B" & next_row - PC_ROWS_PER_CRITERIA & "+1"
        ' make the comment cell suitibly sized
        Rows((next_row + 2) & ":" & (next_row + 2)).Select
        Selection.RowHeight = 95
    Next i
    ' put in the thankyou
    Cells(next_row + 3, 1).Value = THANK_YOU_TEXT
    Rows((next_row + 3) & ":" & (next_row + 4)).Select
    Selection.RowHeight = 20
End Function

Public Function ExpandScoresWithCommentsSheet()
    Sheets(SCORES_AND_COMMENTS_TEMPLATE_SHEET).Select
    Sheets(SCORES_AND_COMMENTS_TEMPLATE_SHEET).Activate
    
    Const THANK_YOU_TEXT  As String = "Thanks you for your contributions to this competition."
    Dim i As Long, next_row As Long
    next_row = Range(SCT_CRITERIA_ONE_NAME_CELL).row
    For i = 2 To num_criteria
        ' copy the preceding comment section and paste below
        Rows(next_row & ":" & (next_row + SCT_ROWS_PER_CRITERIA - 1)).Select
        Selection.Copy
        next_row = next_row + SCT_ROWS_PER_CRITERIA
        Range("A" & next_row).Select
        ActiveSheet.Paste
         ' put in a formula for the criteria number (increment from the one above)
        Cells(next_row, 2).Value = "=" & "B" & next_row - SCT_ROWS_PER_CRITERIA & "+1"
        ' make the comment cell suitibly sized
        Rows((next_row + 3) & ":" & (next_row + 3)).Select
        Selection.RowHeight = 95
    Next i
    ' put in the thankyou
    Cells(next_row + 4, 1).Value = THANK_YOU_TEXT
    Rows((next_row + 4) & ":" & (next_row + 4)).Select
    Selection.RowHeight = 20
    
    ExpandScoresWithCommentsSheet = True
End Function

Public Function CopyAndTransposeCellsGrey(start_col As Long, start_row As Long, _
                num_cols As Long, num_rows As Long, destination_cell As String) As Boolean
                
    ' copy a rectangle of cells to another location switching rows/columns
    
    ' replace the destination cells contents with formulae pointing back to the original cells
    Dim start_range As String
    Dim cell_formulae() As Variant, i As Long, j As Long, formula_str As String
    ReDim cell_formulae(1 To num_cols, 1 To num_rows)
    For i = 1 To num_cols           ' in the destination this is the rows
        For j = 1 To num_rows       ' in the destination this is the columns
            formula_str = "=" & c2l(start_col + i - 1) & (start_row + j - 1)
            cell_formulae(i, j) = formula_str
        Next j
    Next i
    
    ' write out the formulae
    Dim Destination As Range
    Set Destination = Range(destination_cell)
    Destination.Resize(UBound(cell_formulae, 1), UBound(cell_formulae, 2)).Value = cell_formulae
    
    ' copy the formatting of the cells (borders, font)
    start_range = c2l(start_col) & start_row & ":" & _
                 c2l(start_col + num_cols - 1) & (start_row + num_rows - 1)
    Range(start_range).Select
    Selection.Copy
    Range(destination_cell).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=True
        
    ' make the destination array grey (signal that they should not be touched)
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.15
        .PatternTintAndShade = 0
    End With
    
    Range(c2l(start_col + 1) & start_row + 1).Select
    CopyAndTransposeCellsGrey = True
End Function

Function DefineGlobals() As Boolean
    ' get some of the parameters, and make sure some global variables are initialized
    
    If (globals_defined) Then
        DefineGlobals = True
        Exit Function
    End If
    
    main_workbook = ThisWorkbook.Name
    
    With ThisWorkbook
        With .Sheets(SYSTEM_PARAMETERS_SHEET)
            max_markers_per_proj = .Range(SP_MAX_NUMBER_OF_MARKERS_PER_PROJ).Value
            max_ass_per_marker = .Range(SP_MAX_NUMBER_OF_ASSIGNMENTS_PER_MARKER).Value
            max_keywords = .Range(SP_MAX_KEYWORDS_CELL).Value
            
            'get the strings for selecting appropriate excel files from a folder
            project_expertise_file_pattern = .Range(SP_PROJECT_EXPERTISE_FILE_PATTERN).Value
            keyword_expertise_file_pattern = .Range(SP_KEYWORD_EXPERTISE_FILE_PATTERN).Value
            marks_only_file_pattern = .Range(SP_MARKS_ONLY_FILE_PATTERN).Value
            marks_and_cmts_file_pattern = .Range(SP_MARKS_AND_CMTS_FILE_PATTERN).Value
            
            ' extract the file ending
            expertise_by_project_ending = Left(project_expertise_file_pattern, _
                                                InStr(project_expertise_file_pattern, ".xlsx") - 1)
            expertise_by_project_ending = Right(expertise_by_project_ending, Len(expertise_by_project_ending) - 1)
            expertise_by_keyword_ending = Left(keyword_expertise_file_pattern, _
                                                InStr(keyword_expertise_file_pattern, ".xlsx") - 1)
            expertise_by_keyword_ending = Right(expertise_by_keyword_ending, Len(expertise_by_keyword_ending) - 1)
            scores_only_ending = Left(marks_only_file_pattern, InStr(marks_only_file_pattern, ".xlsx") - 1)
            scores_only_ending = Right(scores_only_ending, Len(scores_only_ending) - 1)
            scores_and_cmts_ending = Left(marks_and_cmts_file_pattern, InStr(marks_and_cmts_file_pattern, ".xlsx") - 1)
            scores_and_cmts_ending = Right(scores_and_cmts_ending, Len(scores_and_cmts_ending) - 1)
            
            same_organization_text = .Range(SP_SAME_ORGANIZATION_TEXT_CELL).Value
            ' get the filename ending that flags a user scoresheet
            simulate_marker_responses = .Range(SP_SIMULATE_MARKER_RESPONSES_CELL).Value
            lock_sheet_pwd = .Range(SP_LOCKED_SHEET_PWD_CELL).Value
        End With
    End With
    
    ' make sure we are pulling the competition parameters from the right book (depending on what we are doing)
    Dim wb_name As String
    If making_competition_workbook = False Then
        If Len(cwb) = 0 Then
            If ActivateCompetitionWorkbook = False Then
                DefineGlobals = False
                Exit Function
            End If
        End If
        wb_name = cwb               ' pull the parameters out of the competition workbook
    Else
        wb_name = ThisWorkbook.Name 'pull the parameters out of the macro book
    End If
    With Workbooks(wb_name)
        With .Sheets(COMPETITION_PARAMETERS_SHEET)
            target_ass_per_marker = .Range(CP_TARGET_ASSIGNMENTS_PER_MARKER).Value
            target_markers_per_proj = .Range(CP_TARGET_MARKERS_PER_PROJ).Value
            If target_markers_per_proj > max_markers_per_proj Then
                PopMessage "[DefineGlobals] desired markers per project (" & target_markers_per_proj & _
                        ") exceeds maximum currently possible (" & max_markers_per_proj & ")", vbCritical
                Exit Function
            End If
            root_folder = .Range(CP_COMPETITION_ROOT_FOLDER).Value
            max_first_reader_assignments = .Range(CP_MAX_FIRST_READER_ASSIGNMENTS_CELL).Value
            num_keywords = .Range(CP_NUM_KEYWORDS_CELL).Value
            normalize_scoring = .Range(CP_USE_NORMALIZED_SCORING_CELL).Value
            
            ' expected folder structure under the root folder
            expertise_by_project_requested_folder = .Range(CP_EXPERTISE_BY_PROJECT_REQUESTED_FOLDER_CELL).Value
            expertise_by_project_received_folder = .Range(CP_EXPERTISE_BY_PROJECT_RECEIVED_FOLDER_CELL).Value
            expertise_by_keyword_requested_folder = .Range(CP_EXPERTISE_BY_KEYWORD_REQUESTED_FOLDER_CELL).Value
            expertise_by_keyword_received_folder = .Range(CP_EXPERTISE_BY_KEYWORD_RECEIVED_FOLDER_CELL).Value
            scores_requested_folder = .Range(CP_SCORES_REQUESTED_FOLDER_CELL).Value
            scores_received_folder = .Range(CP_SCORES_RECEIVED_CELL).Value
            comments_folder = .Range(CP_COMMENTS_FOLDER_CELL).Value
            use_org_disambiguation = .Range(CP_USE_ORG_DISAMBIGUATION_CELL).Value
            use_email_disambiguation = .Range(CP_USE_EMAIL_DISAMBIGUATION_CELL).Value
            gather_comments = .Range(CP_GATHER_COMMENTS_CELL).Value
            output_comments_format = .Range(CP_COMMENT_OUTPUT_FORMAT).Value
            blank_expertise_means_exclusion = .Range(CP_BLANK_EXPERTISE_TREATMENT).Value
        End With
    
    ' various other variables
        num_criteria = .Sheets(CRITERIA_SHEET).Range(C_NUM_CRITERIA_CELL).Value
        num_markers = .Sheets(MARKERS_SHEET).Range(M_NUM_MARKERS_CELL).Value    '# of people marking
        num_projects = .Sheets(PROJECTS_SHEET).Range(P_NUM_PROJECTS_CELL).Value '# of projects to mark
        num_keywords = .Sheets(KEYWORDS_SHEET).Range(KW_NUM_KEYWORDS_CELL).Value
    End With
    
    If CheckSomeDataIsEntered = False Then
        DefineGlobals = False
        Exit Function
    End If
    
    ec_assignments_first_column = EC_ASSMT_CONF_FIRST_COL + target_markers_per_proj
    ec_data_first_marker_column = EC_ASSMT_CONF_FIRST_COL + 2 * target_markers_per_proj + 1
                        
    GetOSType   ' figure out whether this is running on Windows or a Mac (IOS)
    
    globals_defined = True
    DefineGlobals = True
    
End Function

Sub StartNewCompetition_Click()
    Const MAX_ROWS_ON_INPUT_SHEETS As Long = 100
    ' show the sheets that need to be filled before creating a competition workbook
    InitMessages
'    making_competition_workbook = True
'    If DefineGlobals = False Then
'        Exit Sub
'    End If
    With ThisWorkbook
        num_criteria = .Sheets(CRITERIA_SHEET).Range(C_NUM_CRITERIA_CELL).Value
        num_markers = .Sheets(MARKERS_SHEET).Range(M_NUM_MARKERS_CELL).Value    '# of people marking
        num_projects = .Sheets(PROJECTS_SHEET).Range(P_NUM_PROJECTS_CELL).Value '# of projects to mark
        num_keywords = .Sheets(KEYWORDS_SHEET).Range(KW_NUM_KEYWORDS_CELL).Value
    End With
    
    Dim sheet_names() As Variant
    sheet_names = Array(COMPETITION_PARAMETERS_SHEET, CRITERIA_SHEET, PROJECTS_SHEET, MARKERS_SHEET, _
                         KEYWORDS_SHEET)
    HideOrShowSheets sheet_names, True
    
    'clear the sheets of program data
    Dim clear_range As String
    Sheets(CRITERIA_SHEET).Select
    If num_criteria > 0 Then
        clear_range = "B" & C_FIRST_DATA_ROW & ":" & "D" & _
                        (C_FIRST_DATA_ROW + Imax(MAX_ROWS_ON_INPUT_SHEETS, num_criteria))
        Range(clear_range).Select
        Range(clear_range).Clear
    Else
        clear_range = "B" & C_FIRST_DATA_ROW
    End If
    Range(FirstCell(clear_range)).Select

    Sheets(PROJECTS_SHEET).Select
    If num_projects > 0 Then
        clear_range = c2l(P_PROJECT_NAME_COLUMN) & P_FIRST_DATA_ROW & ":" & _
                    c2l(P_MENTOR_ID_COLUMN + 3) & (P_FIRST_DATA_ROW + Imax(MAX_ROWS_ON_INPUT_SHEETS, num_projects))
        Range(clear_range).Select
        Range(clear_range).Clear
    Else
        clear_range = c2l(P_PROJECT_NAME_COLUMN) & C_FIRST_DATA_ROW
    End If
    Range(FirstCell(clear_range)).Select
    
    Sheets(MARKERS_SHEET).Select
    If num_markers > 0 Then
        clear_range = c2l(M_NAME_COL) & M_FIRST_DATA_ROW & ":" & _
                    c2l(M_EMAIL_COL) & (M_FIRST_DATA_ROW + Imax(MAX_ROWS_ON_INPUT_SHEETS, num_markers))
        Range(clear_range).Select
        Range(clear_range).Clear
    Else
        clear_range = c2l(M_NAME_COL) & M_FIRST_DATA_ROW
    End If
    Range(FirstCell(clear_range)).Select
    
    Sheets(KEYWORDS_SHEET).Select
    If num_keywords > 0 Then
        clear_range = c2l(2) & KW_FIRST_DATA_ROW & ":" & _
                    c2l(KW_WEIGHTS_COL) & (KW_FIRST_DATA_ROW + Imax(MAX_ROWS_ON_INPUT_SHEETS, max_keywords))
        Range(clear_range).Select
        Range(clear_range).Clear
    Else
        clear_range = c2l(2) & KW_FIRST_DATA_ROW
    End If
    Range(FirstCell(clear_range)).Select
    
    ' all start with the project sheet and a few messages.
    Sheets(PROJECTS_SHEET).Select
    Sheets(PROJECTS_SHEET).Activate
    
    AddMessage "The " & PROJECTS_SHEET & ", " & MARKERS_SHEET & ", " & CRITERIA_SHEET & ", " & _
                    KEYWORDS_SHEET & " sheets have been prepared for data entry."
    AddMessage "Enter data on these sheets as indicated by the column headers."
    AddMessage "When the data has been entered, verify (and adjust) the parameters on the " & _
                COMPETITION_PARAMETERS_SHEET & " sheet."
    AddMessage " "
    AddMessage "When ready, run the 'Build Competition Workbook' macro to create a workbook with the sheets needed for your competition."
    ReportMessages
End Sub

Public Function Imax(num1 As Long, num2 As Long) As Long
    If num1 > num2 Then
        Imax = num1
    Else
        Imax = num2
    End If
End Function
Public Function CheckSomeDataIsEntered() As Boolean
    Dim num_errors As Long
    If num_criteria <= 0 Then
        AddMessage "ERROR: " & num_criteria & " criteria found in the " & CRITERIA_SHEET & " sheet."
        num_errors = num_errors + 1
    End If
    If num_markers <= 0 Then
        AddMessage "ERROR: " & num_markers & " markers found in the " & MARKERS_SHEET & " sheet."
        num_errors = num_errors + 1
    End If
    If num_projects <= 1 Then
        AddMessage "ERROR: " & num_projects & " project(s) found in the " & PROJECTS_SHEET & " sheet."
        num_errors = num_errors + 1
    End If
    If num_keywords <= 0 Then
        AddMessage "ERROR: " & num_keywords & " keywords found in the " & KEYWORDS_SHEET & " sheet."
        num_errors = num_errors + 1
    End If
    If num_errors > 0 Then
        AddMessage "Fix the error(s) by entering data in the appropriate sheets, and then run the 'Build Competition Workbook' macro."
        CheckSomeDataIsEntered = False
        Exit Function
    End If
    CheckSomeDataIsEntered = True
End Function

Function FlagSheetExpanded(Optional comment_range As String = "A1")
    Range(comment_range).Value = "."
' Put a comment in the sheet to indicate it has been expanded (and should not be expanded again)
'    Range(comment_range).Select
'    Range(comment_range).AddCommentThreaded ( _
'        "CoSeT: Sheet has been expanded" & Chr(10) & "(Please do not delete this comment)")
End Function

Function DoesSheetHaveCosetComment(Optional sheet_name_in As String = "", _
                                  Optional comment_range As String = "A1") As Boolean

    Dim varComment As String, sheet_name As String, c As comment
    If Len(sheet_name_in) = 0 Then
        sheet_name = ActiveSheet.Name
    Else
        sheet_name = sheet_name_in
    End If
    If Sheets(sheet_name).Range(comment_range).Value = "." Then
        DoesSheetHaveCosetComment = True
    Else
    End If
    Exit Function
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' the comment approach is not working yet.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With Sheets(sheet_name)
        With Range(comment_range)
            On Error Resume Next
            Set c = .comment
            varComment = .comment
            On Error GoTo 0
            If c Is Nothing Then
                DoesSheetHaveCosetComment = False
            Else
                If InStr(c.Text, COSET) Then
                    DoesSheetHaveCosetComment = True
                Else
                    DoesSheetHaveCosetComment = False
                End If
           End If
        End With
    End With
End Function

Function ShowMasterAssignmentsSheet()
    Sheets(MASTER_ASSIGNMENTS_SHEET).Select
    Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
End Function
