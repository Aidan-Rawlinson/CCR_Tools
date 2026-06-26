Attribute VB_Name = "B1_Importer"
Option Explicit

' ============================================================
' B1_Importer
' ------------------------------------------------------------
' Pure data transfer. Called by B4_Process_Folder once a file
' has passed validation (B5) and been matched to a submission.
'
' Accepts file path, submission ID, org name, and submission
' name as parameters. No validation logic -- B5 owns that.
'
' Finds the next empty row on the Home sheet and appends;
' does not clear existing data (clearing is handled by B4
' at the start of the process run).
'
' Skips any patient position where all question cells are blank
' (threshold: at least 1 non-blank response required to import).
' Works for both Columns and Rows orientations.
'
' Hard fails if the entire file contains no patient data
' (i.e. every position across the full DataMax range is empty).
' ============================================================

Sub FileImporter(ByVal Str_FilePath As String, ByVal Lng_SubmissionID As Long, _
                 ByVal Str_OrgName As String, ByVal Str_SubName As String)

    Dim Wsh_Home As Worksheet:              Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Wbk_Source As Workbook
    Dim Wsh_Source As Worksheet
    Dim Str_SheetName As String:            Str_SheetName = ThisWorkbook.Names("DataSheetName").RefersToRange.Value
    Dim Str_Orientation As String:          Str_Orientation = ThisWorkbook.Names("Orientation").RefersToRange.Value
    Dim Str_DataStart As String:            Str_DataStart = ThisWorkbook.Names("DataStart").RefersToRange.Value
    Dim Lng_DataMax As Long:                Lng_DataMax = ThisWorkbook.Names("DataMax").RefersToRange.Value
    Dim Rng_QuestionCols As Range:          Set Rng_QuestionCols = ThisWorkbook.Names("QuestionCols").RefersToRange
    Dim Rng_StartCols As Range:             Set Rng_StartCols = ThisWorkbook.Names("StartCols").RefersToRange
    Dim Rng_DataArea As Range:              Set Rng_DataArea = ThisWorkbook.Names("DataArea").RefersToRange
    Dim Rng_Cell As Range
    Dim Lng_Record As Long
    Dim Lng_StartIndex As Long
    Dim Lng_Position As Long
    Dim Str_UniqueRef As String
    Dim Lng_ColIndex As Long
    Dim Bln_HasData As Boolean
    Dim Lng_CheckIndex As Long
    Dim Lng_ImportedCount As Long:          Lng_ImportedCount = 0
    Dim Str_FileName As String:             Str_FileName = Mid(Str_FilePath, InStrRev(Str_FilePath, "\") + 1)

    '--Find next empty paste row on Home sheet (append mode)
    Dim Lng_PasteRow As Long
    Lng_PasteRow = Wsh_Home.Cells(Wsh_Home.Rows.Count, Rng_DataArea.Column).End(xlUp).Row
    If Lng_PasteRow < Rng_DataArea.Row Then
        Lng_PasteRow = Rng_DataArea.Row
    Else
        Lng_PasteRow = Lng_PasteRow + 1
    End If

    '--Open source file
    Set Wbk_Source = Workbooks.Open(Str_FilePath, ReadOnly:=True)
    DoEvents

    Set Wsh_Source = Wbk_Source.Worksheets(Str_SheetName)

    '--Import records
    If Str_Orientation = "Columns" Then

        '--Columns orientation: one patient per column, questions in rows
        '--DataStart is a column letter; StartCols holds the source row number for each field
        Lng_StartIndex = Range(Str_DataStart & "1").Column

        For Lng_Record = Lng_StartIndex To Lng_StartIndex + Lng_DataMax - 1

            '--Check whether this patient column has at least one non-blank response
            '--Iterates question positions only (StartCols index 2 onwards; index 1 is unique ref)
            Bln_HasData = False
            For Lng_CheckIndex = 2 To Rng_StartCols.Cells.Count
                If Wsh_Source.Cells(Rng_StartCols.Cells(1, Lng_CheckIndex).Value, Lng_Record).Value <> "" Then
                    Bln_HasData = True
                    Exit For
                End If
            Next Lng_CheckIndex

            If Bln_HasData Then

                '--Read unique reference using the row number stored in StartCols(1)
                Str_UniqueRef = Wsh_Source.Cells(Rng_StartCols.Cells(1, 1).Value, Lng_Record).Value

                '--Write org name, submission name, submission ID, and unique reference to Home
                Wsh_Home.Cells(Lng_PasteRow, Rng_DataArea.Column - 4).Value = Str_OrgName
                Wsh_Home.Cells(Lng_PasteRow, Rng_DataArea.Column - 3).Value = Str_SubName
                Wsh_Home.Cells(Lng_PasteRow, Rng_DataArea.Column - 2).Value = Lng_SubmissionID
                Wsh_Home.Cells(Lng_PasteRow, Rng_DataArea.Column).Value = Str_UniqueRef

                '--Write question responses using StartCols to find source row for each question
                Lng_ColIndex = 2 '--Start at second cell in StartCols (K onwards = questions)
                For Each Rng_Cell In Rng_QuestionCols
                    If Rng_Cell.Column > Rng_DataArea.Column Then
                        Lng_Position = Rng_StartCols.Cells(1, Lng_ColIndex).Value
                        If Lng_Position > 0 Then
                            Wsh_Home.Cells(Lng_PasteRow, Rng_Cell.Column).Value = Wsh_Source.Cells(Lng_Position, Lng_Record).Value
                        End If
                        Lng_ColIndex = Lng_ColIndex + 1
                    End If
                Next Rng_Cell

                Lng_PasteRow = Lng_PasteRow + 1
                Lng_ImportedCount = Lng_ImportedCount + 1

            End If

        Next Lng_Record

    ElseIf Str_Orientation = "Rows" Then

        '--Rows orientation: one patient per row, questions in columns
        '--DataStart is a row number; StartCols holds the source column number for each field
        Lng_StartIndex = CLng(Str_DataStart)

        For Lng_Record = Lng_StartIndex To Lng_StartIndex + Lng_DataMax - 1

            '--Check whether this patient row has at least one non-blank response
            '--Iterates question positions only (StartCols index 2 onwards; index 1 is unique ref)
            Bln_HasData = False
            For Lng_CheckIndex = 2 To Rng_StartCols.Cells.Count
                If Wsh_Source.Cells(Lng_Record, Rng_StartCols.Cells(1, Lng_CheckIndex).Value).Value <> "" Then
                    Bln_HasData = True
                    Exit For
                End If
            Next Lng_CheckIndex

            If Bln_HasData Then

                '--Read unique reference using the column number stored in StartCols(1)
                Str_UniqueRef = Wsh_Source.Cells(Lng_Record, Rng_StartCols.Cells(1, 1).Value).Value

                '--Write org name, submission name, submission ID, and unique reference to Home
                Wsh_Home.Cells(Lng_PasteRow, Rng_DataArea.Column - 4).Value = Str_OrgName
                Wsh_Home.Cells(Lng_PasteRow, Rng_DataArea.Column - 3).Value = Str_SubName
                Wsh_Home.Cells(Lng_PasteRow, Rng_DataArea.Column - 2).Value = Lng_SubmissionID
                Wsh_Home.Cells(Lng_PasteRow, Rng_DataArea.Column).Value = Str_UniqueRef

                '--Write question responses using StartCols to find source column for each question
                Lng_ColIndex = 2 '--Start at second cell in StartCols (K onwards = questions)
                For Each Rng_Cell In Rng_QuestionCols
                    If Rng_Cell.Column > Rng_DataArea.Column Then
                        Lng_Position = Rng_StartCols.Cells(1, Lng_ColIndex).Value
                        If Lng_Position > 0 Then
                            Wsh_Home.Cells(Lng_PasteRow, Rng_Cell.Column).Value = Wsh_Source.Cells(Lng_Record, Lng_Position).Value
                        End If
                        Lng_ColIndex = Lng_ColIndex + 1
                    End If
                Next Rng_Cell

                Lng_PasteRow = Lng_PasteRow + 1
                Lng_ImportedCount = Lng_ImportedCount + 1

            End If

        Next Lng_Record

    End If

    '--Close source file
    Wbk_Source.Close SaveChanges:=False

    '--Report outcome: warn if nothing was imported
    If Lng_ImportedCount = 0 Then
        MsgBox "No patient data was found in this file." & vbCrLf & vbCrLf & _
               "File: " & Str_FileName & vbCrLf & vbCrLf & _
               "All " & Lng_DataMax & " patient positions were blank across all question rows. " & _
               "If the file structure has changed, please review manually.", _
               vbExclamation, "No Patient Data"
    End If

End Sub
