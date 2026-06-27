Attribute VB_Name = "B4_Process_Folder"
Option Explicit

' ============================================================
' B4_Process_Folder
' ------------------------------------------------------------
' Presents a multi-select file picker to the user, saves the
' folder of the first selected file back to SubmissionFolderPath,
' then processes each selected file through the full pipeline:
'
'   Open file
'     → B5: ValidateFile  (mandatory sheets, Support fields, spot checks)
'     → Read org name and submission descriptor via XLookup on Support sheet
'     → Close file
'     → Org/submission matching decision tree
'     → B1: FileImporter (file path, submission ID, org ID, org name,
'                         sub name, ByRef Lng_FirstRow, ByRef Lng_LastRow)
'
' Matching decision tree:
'   Case 1 - No org match:       message + skip, no user choice
'   Case 2 - One submission:     Yes/No confirmation before process
'   Case 3 - Multiple subs:      numbered InputBox + confirmation before process
'
' Org ID is read from column 1 of the Orgs sheet at the matched row
' and passed through to B1 for writing to the Home sheet.
'
' After all files processed, calls the following in sequence
' if any rows were imported, passing the full row range of the run:
'   B6a_DT_Converter      -- parses and formats DT question cells
'   B6_Response_Validator -- validates all response cells
'   B7_Duplicate_Detector -- detects duplicate records
'
' Clear prompt at start of run gives user the option to start
' fresh or append to existing Home sheet data. On clear:
'   1. Formats and contents wiped (.Clear)
'   2. Thin grid borders reapplied to FullDataArea (four sides
'      only -- diagonals explicitly excluded)
' ============================================================

Sub PickAndProcess()

    Dim Wsh_Home As Worksheet:          Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Wsh_Orgs As Worksheet:          Set Wsh_Orgs = ThisWorkbook.Worksheets("Orgs")
    Dim Rng_FolderPath As Range:        Set Rng_FolderPath = ThisWorkbook.Names("SubmissionFolderPath").RefersToRange
    Dim Rng_FullDataArea As Range:      Set Rng_FullDataArea = ThisWorkbook.Names("FullDataArea").RefersToRange

    '--Prompt user: clear existing Home sheet data or append?
    Dim Int_ClearChoice As Integer
    Int_ClearChoice = MsgBox("Do you want to clear existing data from the Home sheet before processing?" & vbCrLf & vbCrLf & _
                             "Click Yes to start with a clean slate." & vbCrLf & _
                             "Click No to append imported rows to existing data.", _
                             vbQuestion + vbYesNo, "Clear Existing Data?")

    If Int_ClearChoice = vbYes Then
        '--Clear all contents and formatting (removes orange validation colouring from previous runs)
        Rng_FullDataArea.Clear
        '--Reapply thin grid borders -- four sides only, diagonals excluded
        With Rng_FullDataArea
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
        End With
    End If

    '--Build initial folder for dialog
    Dim Str_StartFolder As String
    Str_StartFolder = Rng_FolderPath.Value

    If Str_StartFolder <> "" Then
        If InStr(Str_StartFolder, "\") > 0 Then
            If Dir(Str_StartFolder, vbDirectory) = "" Then
                Str_StartFolder = Left(Str_StartFolder, InStrRev(Str_StartFolder, "\"))
            End If
        End If
    End If

    '--Show multi-select file picker
    Dim Dlg As FileDialog
    Set Dlg = Application.FileDialog(msoFileDialogFilePicker)

    With Dlg
        .Title = "Select submission file(s) to process"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm; *.xlsb"
        .AllowMultiSelect = True
        If Str_StartFolder <> "" Then .InitialFileName = Str_StartFolder
    End With

    If Dlg.Show = False Then Exit Sub     '--User cancelled

    '--Save folder of first selected file back to SubmissionFolderPath
    Dim Str_FirstFile As String:        Str_FirstFile = Dlg.SelectedItems(1)
    Rng_FolderPath.Value = Left(Str_FirstFile, InStrRev(Str_FirstFile, "\"))

    '--Counters for end-of-run summary
    Dim Int_Processed As Integer:       Int_Processed = 0
    Dim Int_Skipped As Integer:         Int_Skipped = 0
    Dim Int_Total As Integer:           Int_Total = Dlg.SelectedItems.Count

    '--Row tracking -- spans the full run across all files
    Dim Lng_RunFirstRow As Long:        Lng_RunFirstRow = 0
    Dim Lng_RunLastRow As Long:         Lng_RunLastRow = 0
    Dim Lng_FileFirstRow As Long
    Dim Lng_FileLastRow As Long

    '--Loop over selected files
    Dim i As Integer
    For i = 1 To Int_Total

        Dim Str_FilePath As String:     Str_FilePath = Dlg.SelectedItems(i)
        Dim Str_FileName As String:     Str_FileName = Mid(Str_FilePath, InStrRev(Str_FilePath, "\") + 1)

        '--Open file
        Dim Wbk_Source As Workbook
        Set Wbk_Source = Nothing

        On Error Resume Next
        Set Wbk_Source = Workbooks.Open(Str_FilePath, ReadOnly:=True)
        On Error GoTo 0

        If Wbk_Source Is Nothing Then
            MsgBox "Could not open file:" & vbCrLf & vbCrLf & _
                   Str_FileName & vbCrLf & vbCrLf & _
                   "File will be skipped.", _
                   vbExclamation, "File Error"
            Int_Skipped = Int_Skipped + 1
            GoTo NextFile
        End If

        '--B5: Validate file structure and content
        If Not B5_File_Validator.ValidateFile(Wbk_Source, Str_FileName) Then
            Wbk_Source.Close SaveChanges:=False
            Set Wbk_Source = Nothing
            Int_Skipped = Int_Skipped + 1
            GoTo NextFile
        End If

        '--Read org name and submission descriptor via XLookup on Support sheet
        '--File has passed validation so Support sheet and fields are guaranteed present
        Dim Wsh_Support As Worksheet
        Set Wsh_Support = Wbk_Source.Worksheets("Support")

        Dim Str_OrgName As String
        Dim Str_SubDescriptor As String

        On Error Resume Next
        Str_OrgName = Trim(CStr(WorksheetFunction.XLookup( _
                        "Organisation Name", _
                        Wsh_Support.Columns(1), _
                        Wsh_Support.Columns(2), _
                        "")))
        Str_SubDescriptor = Trim(CStr(WorksheetFunction.XLookup( _
                        "Submission Name", _
                        Wsh_Support.Columns(1), _
                        Wsh_Support.Columns(2), _
                        "")))
        '--Fallback for Virtual Ward templates which use "Virtual Ward Name" instead of "Submission Name"
        If Str_SubDescriptor = "" Or Str_SubDescriptor = "0" Then
            Str_SubDescriptor = Trim(CStr(WorksheetFunction.XLookup( _
                            "Virtual Ward Name", _
                            Wsh_Support.Columns(1), _
                            Wsh_Support.Columns(2), _
                            "")))
        End If
        On Error GoTo 0

        Wbk_Source.Close SaveChanges:=False
        Set Wbk_Source = Nothing
        Set Wsh_Support = Nothing

        '--Find matching rows in Orgs sheet (column B = Org Name)
        Dim Lng_LastOrgRow As Long
        Lng_LastOrgRow = Wsh_Orgs.Cells(Wsh_Orgs.Rows.Count, 2).End(xlUp).Row

        '--Collect matching row indices
        Dim Arr_MatchRows() As Long
        Dim Int_MatchCount As Integer:  Int_MatchCount = 0
        ReDim Arr_MatchRows(1 To Lng_LastOrgRow)

        Dim r As Long
        For r = 2 To Lng_LastOrgRow
            If Trim(CStr(Wsh_Orgs.Cells(r, 2).Value)) = Str_OrgName Then
                Int_MatchCount = Int_MatchCount + 1
                Arr_MatchRows(Int_MatchCount) = r
            End If
        Next r

        '--Decision tree
        If Int_MatchCount = 0 Then

            '--Case 1: No match
            MsgBox "No matching organisation was found for this file:" & vbCrLf & vbCrLf & _
                   "File:         " & Str_FileName & vbCrLf & _
                   "Organisation: " & Str_OrgName & vbCrLf & vbCrLf & _
                   "Please check that the organisation name in the template matches exactly " & _
                   "with the organisations loaded in the tool. File will be skipped.", _
                   vbExclamation, "No Organisation Match"
            Int_Skipped = Int_Skipped + 1

        ElseIf Int_MatchCount = 1 Then

            '--Case 2: Exactly one submission
            Dim Lng_OrgID2 As Long
            Lng_OrgID2 = CLng(Wsh_Orgs.Cells(Arr_MatchRows(1), 1).Value)
            Dim Str_SubName2 As String
            Str_SubName2 = Trim(CStr(Wsh_Orgs.Cells(Arr_MatchRows(1), 3).Value))
            Dim Lng_SubID2 As Long
            Lng_SubID2 = CLng(Wsh_Orgs.Cells(Arr_MatchRows(1), 4).Value)

            Dim Str_Msg2 As String
            Str_Msg2 = "One submission found for this file." & vbCrLf & vbCrLf & _
                       "File:             " & Str_FileName & vbCrLf & _
                       "Organisation:     " & Str_OrgName & vbCrLf & _
                       "Submission:       " & Str_SubName2 & " (ID: " & Lng_SubID2 & ")"

            If Str_SubDescriptor <> "" And Str_SubDescriptor <> "0" Then
                Str_Msg2 = Str_Msg2 & vbCrLf & _
                           "Template note:    " & Str_SubDescriptor
            End If

            Str_Msg2 = Str_Msg2 & vbCrLf & vbCrLf & "Proceed with import?"

            If MsgBox(Str_Msg2, vbQuestion + vbYesNo, "Confirm Submission") = vbYes Then
                Lng_FileFirstRow = 0
                Lng_FileLastRow = 0
                Call B1_Importer.FileImporter(Str_FilePath, Lng_SubID2, Lng_OrgID2, Str_OrgName, Str_SubName2, _
                                              Lng_FileFirstRow, Lng_FileLastRow)
                If Lng_FileFirstRow > 0 Then
                    If Lng_RunFirstRow = 0 Then Lng_RunFirstRow = Lng_FileFirstRow
                    Lng_RunLastRow = Lng_FileLastRow
                End If
                Int_Processed = Int_Processed + 1
            Else
                Int_Skipped = Int_Skipped + 1
            End If

        Else

            '--Case 3: Multiple submissions
            Dim Str_Msg3 As String
            Str_Msg3 = "Multiple submissions found for this organisation." & vbCrLf & vbCrLf & _
                       "File:          " & Str_FileName & vbCrLf & _
                       "Organisation:  " & Str_OrgName & vbCrLf

            If Str_SubDescriptor <> "" And Str_SubDescriptor <> "0" Then
                Str_Msg3 = Str_Msg3 & "Template note: " & Str_SubDescriptor & vbCrLf
            End If

            Str_Msg3 = Str_Msg3 & vbCrLf & "Please enter the number of the correct submission:" & vbCrLf

            Dim j As Integer
            For j = 1 To Int_MatchCount
                Str_Msg3 = Str_Msg3 & vbCrLf & "  " & j & ".  " & _
                           Trim(CStr(Wsh_Orgs.Cells(Arr_MatchRows(j), 3).Value)) & _
                           " (ID: " & Trim(CStr(Wsh_Orgs.Cells(Arr_MatchRows(j), 4).Value)) & ")"
            Next j

            Dim Str_Input As String
            Str_Input = InputBox(Str_Msg3, "Select Submission")

            Dim Int_Choice As Integer:  Int_Choice = 0
            If IsNumeric(Str_Input) Then Int_Choice = CInt(Str_Input)

            If Int_Choice >= 1 And Int_Choice <= Int_MatchCount Then

                '--Valid choice - confirm before processing
                Dim Lng_OrgID3 As Long
                Lng_OrgID3 = CLng(Wsh_Orgs.Cells(Arr_MatchRows(Int_Choice), 1).Value)
                Dim Lng_SubID3 As Long
                Lng_SubID3 = CLng(Wsh_Orgs.Cells(Arr_MatchRows(Int_Choice), 4).Value)
                Dim Str_SubName3 As String
                Str_SubName3 = Trim(CStr(Wsh_Orgs.Cells(Arr_MatchRows(Int_Choice), 3).Value))

                MsgBox "Selection confirmed." & vbCrLf & vbCrLf & _
                       "File:        " & Str_FileName & vbCrLf & _
                       "Matched to:  " & Str_SubName3 & " (ID: " & Lng_SubID3 & ")" & vbCrLf & vbCrLf & _
                       "Processing will now begin.", _
                       vbInformation, "Submission Matched"

                Lng_FileFirstRow = 0
                Lng_FileLastRow = 0
                Call B1_Importer.FileImporter(Str_FilePath, Lng_SubID3, Lng_OrgID3, Str_OrgName, Str_SubName3, _
                                              Lng_FileFirstRow, Lng_FileLastRow)
                If Lng_FileFirstRow > 0 Then
                    If Lng_RunFirstRow = 0 Then Lng_RunFirstRow = Lng_FileFirstRow
                    Lng_RunLastRow = Lng_FileLastRow
                End If
                Int_Processed = Int_Processed + 1

            Else

                '--Invalid or blank entry
                MsgBox "No valid selection was made." & vbCrLf & vbCrLf & _
                       "File:  " & Str_FileName & vbCrLf & vbCrLf & _
                       "File will be skipped.", _
                       vbInformation, "No Selection Made"
                Int_Skipped = Int_Skipped + 1

            End If

        End If

NextFile:
        Set Wbk_Source = Nothing
        Set Wsh_Support = Nothing

    Next i

    '--Run post-import processing across all rows imported this run
    If Lng_RunFirstRow > 0 Then
        Call B6a_DT_Converter.ConvertDTColumns(Lng_RunFirstRow, Lng_RunLastRow)
        Call B6_Response_Validator.ValidateResponses(Lng_RunFirstRow, Lng_RunLastRow)
        Call B7_Duplicate_Detector.DetectDuplicates(Lng_RunFirstRow, Lng_RunLastRow)
    End If

    '--End of run summary
    MsgBox "Processing complete." & vbCrLf & vbCrLf & _
           "Files selected:   " & Int_Total & vbCrLf & _
           "Files processed:  " & Int_Processed & vbCrLf & _
           "Files skipped:    " & Int_Skipped, _
           vbInformation, "Run Complete"

End Sub
