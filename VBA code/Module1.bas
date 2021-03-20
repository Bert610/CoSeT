Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generic functions in this module (i.e. generally ones that don't require globals to work
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Global os_type As Variant                   ' for mac or windows
' this function (of course) is the exception, it needs the os_type global
Sub GetOSType()
    os_type = Application.OperatingSystem
    If InStr(os_type, "Windows") Then
        os_type = "Windows"
        fps = "\"           'this is the folder path delimiter
    Else
        os_type = "Mac"
        fps = "/"           'this is the folder path delimiter
    End If
End Sub

Public Function GetFileSaveasName(dialog_title As String, initial_name As String) As String
    Dim varResult As Variant
    'displays the save file dialog
    If os_type = "Windows" Then
        varResult = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", _
           title:=dialog_title, InitialFileName:=initial_name)
        If varResult = False Then
            Exit Function
        Else
            GetFileSaveasName = varResult
        End If
    Else
            ' mac: only option that works is InitialFilename
            GetFileSaveasName = Application.GetSaveAsFilename(InitialFileName:=initial_name)
    End If
End Function

Function MacGetSaveAsFilenameExcel(MyInitialFilename As String, FileExtension As String)
'Ron de Bruin, 03-April-2015
'Custom function for the Mac to save the activeworkbook in the format you want.
'If FileExtension = "" you can save in the following formats : xls, xlsx, xlsm, xlsb
'You can also set FileExtension to the extension you want like "xlsx" for example
    Dim FName As Variant
    Dim FileFormatValue As Long
    Dim TestIfOpen As Workbook
    Dim FileExtGetSaveAsFilename As String

Again: FName = False
    
    'Call VBA GetSaveAsFilename
    'Note: InitialFilename is the only parameter that works on a Mac
    FName = Application.GetSaveAsFilename(InitialFileName:=MyInitialFilename)

    If FName <> False Then
        'Get the file extension
        FileExtGetSaveAsFilename = LCase(Right(FName, Len(FName) - InStrRev(FName, ".", , 1)))

        If FileExtension <> "" Then
            If FileExtension <> FileExtGetSaveAsFilename Then
                MsgBox "Sorry you must save the file in this format : " & FileExtension
                GoTo Again
            End If
            If ActiveWorkbook.HasVBProject = True And LCase(FileExtension) = "xlsx" Then
                MsgBox "Your workbook have VBA code, please not save in xlsx format"
                Exit Function
            End If
        Else
            If ActiveWorkbook.HasVBProject = True And LCase(FileExtGetSaveAsFilename) = "xlsx" Then
                MsgBox "Your workbook have VBA code, please not save in xlsx format"
                GoTo Again
            End If
        End If

        'Find the correct FileFormat that match the choice in the "Save as type" list
        'and set the FileFormatValue, Extension and FileFormatValue must match.
        'Note : You can add or delete items to/from the list below if you want.
        Select Case FileExtGetSaveAsFilename
        Case "xls": FileFormatValue = 57
        Case "xlsx": FileFormatValue = 52
        Case "xlsm": FileFormatValue = 53
        Case "xlsb": FileFormatValue = 51
        Case Else: FileFormatValue = 0
        End Select
        If FileFormatValue = 0 Then
            MsgBox "Sorry, FileFormat not allowed"
            GoTo Again
        Else
            'Error check if there is a file open with that name
            Set TestIfOpen = Nothing
            On Error Resume Next
            Set TestIfOpen = Workbooks(LCase(Right(FName, Len(FName) - InStrRev(FName, _
                Application.PathSeparator, , 1))))
            On Error GoTo 0

            If Not TestIfOpen Is Nothing Then
                MsgBox "You are not allowed to overwrite a file that is open with the same name, " & _
                "use a different name or close the file with the same name first."
                GoTo Again
            End If
        End If

        'Now we have the information to Save the file
        Application.DisplayAlerts = False
        On Error Resume Next
        ActiveWorkbook.SaveAs FName, FileFormat:=FileFormatValue
        On Error GoTo 0
        Application.DisplayAlerts = True
    End If

End Function

 Sub test_SelectFolder()
    Dim title As String, start_folder As String
    title = "Testing Select Folder"
    start_folder = "C:\Users\bvmm\Dropbox\Berts Files\Work in DROPBOX\CoSeT\MacroTesting\P15M10R4_FROM_SHEETS\Expertise by Project Received\"
    MsgBox SelectFolder(title, start_folder), vbOKOnly
 End Sub

Public Function SelectFolder(title As String, start_folder As String) As String
    ' ask the user to select a folder
    Dim diaFolder As FileDialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.title = title
    Dim path As String
    path = Dir(start_folder, vbDirectory)   ' check if the folder exists (return will be non-null)
    If Len(path) = 0 Then
        ' PopMessage "WARNING: File path in [SelectFolder] <" & start_folder & "> does not exist", vbOK
    Else
        diaFolder.InitialFileName = path
    End If
    
    Dim asu As Boolean, ada As Boolean
    asu = Application.ScreenUpdating
    ada = Application.DisplayAlerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim selected As Boolean
    selected = diaFolder.Show

    Application.ScreenUpdating = asu
    Application.DisplayAlerts = ada
    
    If selected Then
       SelectFolder = diaFolder.SelectedItems(1)
    End If

'    Set diaFolder = Nothing
End Function

Public Function PopMessage(msg As String, status As Long) As Long
    ' wrapper around msgbox to make sure application status is enabled
    Dim SU_status As Boolean, DA_status As Boolean, ret_status As Long
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ret_status = MsgBox(msg, status)
    Application.ScreenUpdating = SU_status
    Application.DisplayAlerts = DA_status
    PopMessage = ret_status
End Function

Public Function AutofitOneColumn(column_num As Long) As Boolean
    ' resize the specified column to its contents
    Dim col_name As String, range_name As String
    col_name = c2l(column_num)
    range_name = col_name & ":" & col_name
    Columns(range_name).Select
    Range(col_name & 1).Activate
    Columns(range_name).EntireColumn.AutoFit
End Function

Public Function ResizeToNarrowColumn(column_num As Long) As Boolean
    ' make the specified column narrow
    Dim col_name As String, range_name As String
    col_name = c2l(column_num)
    range_name = col_name & ":" & col_name
    Columns(range_name).Select
    Range(col_name & 1).Activate
    Columns(range_name).ColumnWidth = 1.8
End Function

Public Function PutFormulaAndDragDown(form_cell As String, form_str As String, num_cells As Long) As Boolean
    Range(form_cell).Value = form_str
    Dim fill_range As String
    fill_range = form_cell & ":" & c2l(Range(form_cell).Column) & (Range(form_cell).row + num_cells)
    Range(form_cell).Select
    Selection.AutoFill Destination:=Range(fill_range), Type:=xlFillValues
    PutFormulaAndDragDown = True
End Function

Public Function InsertAndExpandRight(start_col As Long, start_row As Long, _
                        row_span As Long, num_cols2add As Long) As Boolean
' the template sheets have 2 columns of data (to ensure spanning formulas grow)
' this function grows the data columns and fills with the formulas & formatting in the first column

' insert columns to the rigth of start_col and then auto fill them, including the
' column to the right of the start_col.
' Since the template has two rows/columns of data (to allow insertions to keep spanning
' formulas) the function will delete the second row if only one row/colum is desired

    Dim start_range As String, insert_range As String, fill_range As String
    
    If num_cols2add < 0 Then
        ' only 1 column of data is wanted so delete the second one
        Dim delete_range As String
        delete_range = c2l(start_col + 1) & (start_row) & ":" & _
                c2l(start_col + num_cols2add) & (start_row + row_span - 1)
        Range(delete_range).Select
        Selection.Delete Shift:=xlToLeft
        Range(c2l(start_col) & start_row + 1).Select
        Exit Function
    End If
    If num_cols2add > 0 Then
        insert_range = c2l(start_col + 1) & start_row & ":" & _
                    c2l(start_col + num_cols2add) & (start_row + row_span - 1)
        Range(insert_range).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    start_range = c2l(start_col) & start_row & ":" & _
                c2l(start_col) & start_row + row_span - 1
    Range(start_range).Select
    fill_range = c2l(start_col) & start_row & ":" & _
                c2l(start_col + num_cols2add + 1) & (start_row + row_span - 1)
    Selection.AutoFill Destination:=Range(fill_range), Type:=xlFillDefault
    ' set the column width to that of start_col
    Dim col_width As Double
    col_width = Columns(start_col).Width
    Columns(c2l(start_col) & ":" & c2l(start_col + num_cols2add + 1)).ColumnWidth = col_width / (num_cols2add + 2)
    
    InsertAndExpandRight = True
End Function

Public Function InsertAndExpandDown(start_col As Long, start_row As Long, _
                        column_span As Long, num_rows2add As Long) As Boolean
' the template sheets have 2 rows of data (to ensure spanning formulas grow)
' this function grows the data rows and fills with the formulas & formatting in the first row

' insert rows below start row and then auto fill them, including the row
' below the insertion.
' Since the template has two rows/columns of data (to allow insertions to keep spanning
' formulas) the function will delete the second row if only one row/colum is desired

    
    Select Case num_rows2add
    Case Is > 0
        Dim start_range As String, insert_range As String, fill_range As String
        insert_range = c2l(start_col) & start_row + 1 & ":" & _
                    c2l(start_col + column_span - 1) & start_row + num_rows2add
        Range(insert_range).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        start_range = c2l(start_col) & start_row & ":" & _
                    c2l(start_col + column_span - 1) & start_row
        Range(start_range).Select
        fill_range = c2l(start_col) & start_row & ":" & _
                    c2l(start_col + column_span - 1) & start_row + num_rows2add + 1
        Selection.AutoFill Destination:=Range(fill_range), Type:=xlFillDefault
    Case 0
        ' do nothing
    Case Is < 0
        ' delete rows below the start row
        Dim delete_range As String
        delete_range = c2l(start_col) & start_row + 1 & ":" & _
                c2l(start_col + column_span - 1) & start_row + Abs(num_rows2add)
        Range(delete_range).Select
        Selection.Delete Shift:=xlUp
        Range(c2l(start_col) & start_row).Select
    End Select
    
    InsertAndExpandDown = True
End Function

Public Function MergeVertical(cn As Long, start_row As Long, end_row As Long) As Boolean
    Range(c2l(cn) & start_row & ":" & c2l(cn) & end_row).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 90
        .MergeCells = True
    End With

End Function

Public Function MergeHorizontal(rownum As Long, start_col As Long, end_col As Long) As Boolean
    Range(c2l(start_col) & rownum & ":" & c2l(end_col) & rownum).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .MergeCells = True
    End With

End Function


Public Function ConvertCellsDownFromFormula2Text(start_row As Long, column_name As String, num_cells)
    Dim last_row As Long
    last_row = start_row + num_cells - 1

    Range(column_name & start_row & ":" & column_name & last_row).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Function

Function HideOrShowSheets(sheets_array() As Variant, visible_flag As Boolean) As Boolean

    Dim i As Long
    For i = LBound(sheets_array) To UBound(sheets_array)
        If visible_flag Then
            ' show the sheets
            Sheets(sheets_array(i)).Visible = True
        Else
            ' hide the sheets
            Sheets(sheets_array(i)).Select
'            Sheets(sheets_array(i)).Activate
            ActiveWindow.SelectedSheets.Visible = False
        End If
    Next i
    HideOrShowSheets = True
End Function

Public Function is_array_empty(array2test() As Variant) As Boolean
    Dim i As Long, j As Long
    Dim upper1 As Long, upper2 As Long
    upper1 = UBound(array2test, 1)
    upper2 = UBound(array2test, 2)
    For i = 1 To upper1
        For j = 1 To upper2
            If Len(array2test(i, j)) > 0 Then
                is_array_empty = False
                Exit Function
            End If
        Next j
    Next i
    
    is_array_empty = True
    
End Function

Function AddMessage(msg2add As String) As Long
    If buffer_messages Then
        num_messages = num_messages + 1
        ReDim Preserve messages(1 To num_messages)
        messages(num_messages) = msg2add
        AddMessage = num_messages
    Else
        MsgBox msg2add, vbOKOnly
    End If
End Function

Function InitMessages()
    Erase messages
    buffer_messages = True
    num_messages = 0
End Function

Function ReportMessages()
    Dim msg_out As String
    If num_messages > 0 Then
        Dim i As Long
        For i = 1 To num_messages
            msg_out = msg_out & messages(i)
            If i <> num_messages Then
                msg_out = msg_out & vbCrLf
            End If
        Next i
        MsgBox msg_out, vbOKOnly
    End If
    num_messages = 0

End Function

Public Function formulas2text(sheet_name As String, range2change As String) As Boolean
    
    Dim current_cell As String, current_sheet As String, editsheet_celladdress As String
    current_sheet = ActiveSheet.Name
    current_cell = ActiveCell.Address
    Sheets(sheet_name).Select
    editsheet_celladdress = ActiveCell.Address
    Range(range2change).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range(editsheet_celladdress).Activate
    Sheets(current_sheet).Select
    Range(current_cell).Activate
    
    formulas2text = True
End Function

Public Function c2l(col_num As Long) As String

    ' whole different approach - from the stack overflow
    c2l = Split((Columns(col_num).Address(, 0)), ":")(0)
    Exit Function

End Function

Public Function ChangeActiveCell(num_row2move As Long, num_col2move As Long) As String 'returns the newly active cell name
    'positive arguments move the active cell to the right and/or down
    Dim curr_row As Long, curr_col As Long
    curr_row = ActiveCell.row
    curr_col = ActiveCell.Column
    Dim curr_cell_name As String
    curr_row = curr_row + num_row2move
    If (curr_row < 1) Or (curr_col + num_col2move < 1) Then
        PopMessage "[ChangeActiveCell] attempt to move before first row or column", vbCritical
        ChangeActiveCell = "A1"
        Range(ChangeActiveCell).Select
    Else
        curr_cell_name = c2l(curr_col + num_col2move) & curr_row
        Range(curr_cell_name).Select
        ChangeActiveCell = curr_cell_name
    End If
End Function

Public Function FirstCell(range_string As String) As String
    If InStr(range_string, ":") = 0 Then
        If InStr(range_string, ",") Then
            FirstCell = Left(range_string, InStr(range_string, ",") - 1)
        Else
            FirstCell = range_string
        End If
    Else
        FirstCell = Left(range_string, InStr(range_string, ":") - 1)
    End If
End Function

Public Function ConvertRangeToText(range_in As String) As Boolean

    Range(range_in).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    ConvertRangeToText = True

End Function


Public Function GoodTabName(name_in As String) As String
    Dim i As Long, j As Long, name_out As String, one_char As String
    Const MAX_TAB_NAME_LENGTH As Long = 30
    
    For i = 1 To Len(name_in)
        one_char = Mid(name_in, i, 1)
        If ((one_char >= "a") And (one_char <= "z")) Or _
        ((one_char >= "A") And (one_char <= "Z")) Or _
        ((one_char >= "0") And (one_char <= "9")) Or _
        (one_char = "_") Or (one_char = "-") Or (one_char = " ") Then
            j = j + 1
            name_out = name_out & one_char
        End If
    Next i
    If j = 0 Then name_out = "BAD TAB NAME"
    If Len(name_out) > MAX_TAB_NAME_LENGTH Then
        name_out = Left(name_out, MAX_TAB_NAME_LENGTH)
    End If
    GoodTabName = name_out
End Function

Public Function DuplicateTemplateSheet(sheet_name As String) As String ' returns the name of the new sheet
' make a duplicate of a given sheet and make it the active sheet
' new sheet's name (as created by Excel) is returned
    On Error GoTo duplicating_sheet_error
    
    Sheets(sheet_name).Copy Before:=Sheets(Sheets.Count)
    DuplicateTemplateSheet = ActiveSheet.Name
    Exit Function
duplicating_sheet_error:
    PopMessage "[DuplicateTemplateSheet] Error duplicating sheet {" & sheet_name & "}, check sheet exists", vbCritical
    DuplicateTemplateSheet = ""
    On Error GoTo 0
    Exit Function
End Function


