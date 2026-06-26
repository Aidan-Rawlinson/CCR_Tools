Attribute VB_Name = "Process_Folder"
Option Explicit

' ============================================================
' Process_Folder
' ------------------------------------------------------------
' Presents a multi-select file picker to the user, saves the
' folder of the first selected file back to SubmissionFolderPath,
' then processes each selected file through the org/submission
' matching decision tree.
'
' Matching decision tree:
'   Case 1 - No org match:       message + skip, no user choice
'   Case 2 - One submission:     Yes/No confirmation before process
'   Case 3 - Multiple subs:      numbered InputBox + confirmation before process
'
' Once a file passes matching it calls ProcessValidFile (stub).
' ============================================================

Sub PickAndProcess()

    Dim Wsh_Orgs As Worksheet:          Set Wsh_Orgs = ThisWorkbook.Worksheets("Orgs")
    Dim Rng_FolderPath As Range:        Set Rng_FolderPath = ThisWorkbook.Names("SubmissionFolderPath").RefersToRange

    '--Build initial folder for dialog
    Dim Str_StartFolder As String
    Str_StartFolder = Rng_FolderPath.Value

    If Str_StartFolder <> "" Then
        '--If it's a file path, strip back to folder
        If InStr(Str_StartFolder, "\") > 0 Then
            If Dir(Str_StartFolder, vbDirectory) = "" Then
                '--Looks like a file path; take the folder portion
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

    '--Loop over selected files
    Dim i As Integer
    For i = 1 To Int_Total

        Dim Str_FilePath As String:     Str_FilePath = Dlg.SelectedItems(i)
        Dim Str_FileName As String:     Str_FileName = Mid(Str_FilePath, InStrRev(Str_FilePath, "\") + 1)

        '--Read org name and submission descriptor from Support sheet
        Dim Wbk_Source As Workbook
        Dim Str_OrgName As String
        Dim Str_SubDescriptor As String

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

        Dim Wsh_Support As Worksheet
        On Error Resume Next
        Set Wsh_Support = Wbk_Source.Worksheets("Support")
        On Error GoTo 0

        If Wsh_Support Is Nothing Then
            MsgBox "The file does not contain a 'Support' sheet and cannot be read:" & vbCrLf & vbCrLf & _
                   Str_FileName & vbCrLf & vbCrLf & _
                   "File will be skipped.", _
                   vbExclamation, "Invalid File"
            Wbk_Source.Close SaveChanges:=False
            Set Wbk_Source = Nothing
            Int_Skipped = Int_Skipped + 1
            GoTo NextFile
        End If

        Str_OrgName = Trim(CStr(Wsh_Support.Range("B5").Value))
        Str_SubDescriptor = Trim(CStr(Wsh_Support.Range("B6").Value))

        Wbk_Source.Close SaveChanges:=False
        Set Wbk_Source = Nothing

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
                Call ProcessValidFile(Str_FilePath, Lng_SubID2)
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
                Dim Lng_SubID3 As Long
                Lng_SubID3 = CLng(Wsh_Orgs.Cells(Arr_MatchRows(Int_Choice), 4).Value)
                Dim Str_SubName3 As String
                Str_SubName3 = Trim(CStr(Wsh_Orgs.Cells(Arr_MatchRows(Int_Choice), 3).Value))

                MsgBox "Selection confirmed." & vbCrLf & vbCrLf & _
                       "File:        " & Str_FileName & vbCrLf & _
                       "Matched to:  " & Str_SubName3 & " (ID: " & Lng_SubID3 & ")" & vbCrLf & vbCrLf & _
                       "Processing will now begin.", _
                       vbInformation, "Submission Matched"

                Call ProcessValidFile(Str_FilePath, Lng_SubID3)
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

    '--End of run summary
    MsgBox "Processing complete." & vbCrLf & vbCrLf & _
           "Files selected:   " & Int_Total & vbCrLf & _
           "Files processed:  " & Int_Processed & vbCrLf & _
           "Files skipped:    " & Int_Skipped, _
           vbInformation, "Run Complete"

End Sub


' ============================================================
' ProcessValidFile  (stub)
' ------------------------------------------------------------
' Called once a file has been matched to a submission.
' Replace this stub with the real import call in Session B.
' ============================================================

Private Sub ProcessValidFile(ByVal Str_FilePath As String, ByVal Lng_SubmissionID As Long)

    Dim Str_FileName As String
    Str_FileName = Mid(Str_FilePath, InStrRev(Str_FilePath, "\") + 1)

    MsgBox "Valid file." & vbCrLf & vbCrLf & _
           "File:          " & Str_FileName & vbCrLf & _
           "Submission ID: " & Lng_SubmissionID, _
           vbInformation, "Valid File"

End Sub
