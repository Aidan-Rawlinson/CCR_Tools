Attribute VB_Name = "B1_Importer"
Option Explicit

' ============================================================
' B1_Importer
' ------------------------------------------------------------
' Pure data transfer. Called by B4_Process_Folder once a file
' has passed validation (B5) and been matched to a submission.
'
' Accepts file path and submission ID as parameters.
' No validation logic -- B5 owns that.
'
' Finds the next empty row on the Home sheet and appends;
' does not clear existing data (clearing is handled by B4
' at the start of the process run).
'
' Hard fails if the first expected patient position is blank
' (i.e. the file contains no patient data).
' ============================================================

Sub FileImporter(ByVal Str_FilePath As String, ByVal Lng_SubmissionID As Long)

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

        '--Empty file check: probe unique reference cell of first expected patient
        '--Note: this check looks at the first patient position only
        If Wsh_Source.Cells(Rng_StartCols.Cells(1, 1).Value, Lng_StartIndex).Value = "" Then
            MsgBox "No patient data was found in this file." & vbCrLf & vbCrLf & _
                   "File: " & Str_FileName & vbCrLf & vbCrLf & _
                   "This check looks at the first expected patient position only. " & _
                   "If the file structure has changed, please review manually." & vbCrLf & vbCrLf & _
                   "File will be skipped.", _
                   vbExclamation, "No Patient Data"
            Wbk_Source.Close SaveChanges:=False
            Exit Sub
        End If

        For Lng_Record = Lng_StartIndex To Lng_StartIndex + Lng_DataMax - 1

            '--Read unique reference using the row number stored in StartCols(1)
            Str_UniqueRef = Wsh_Source.Cells(Rng_StartCols.Cells(1, 1).Value, Lng_Record).Value

            If Str_UniqueRef <> "" Then

                '--Write submission ID to column H and unique reference to column J on Home
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

            End If

        Next Lng_Record

    ElseIf Str_Orientation = "Rows" Then

        '--Rows orientation: one patient per row, questions in columns
        '--DataStart is a row number; StartCols holds the source column number for each field
        Lng_StartIndex = CLng(Str_DataStart)

        '--Empty file check: probe unique reference cell of first expected patient
        '--Note: this check looks at the first patient position only
        If Wsh_Source.Cells(Lng_StartIndex, Rng_StartCols.Cells(1, 1).Value).Value = "" Then
            MsgBox "No patient data was found in this file." & vbCrLf & vbCrLf & _
                   "File: " & Str_FileName & vbCrLf & vbCrLf & _
                   "This check looks at the first expected patient position only. " & _
                   "If the file structure has changed, please review manually." & vbCrLf & vbCrLf & _
                   "File will be skipped.", _
                   vbExclamation, "No Patient Data"
            Wbk_Source.Close SaveChanges:=False
            Exit Sub
        End If

        For Lng_Record = Lng_StartIndex To Lng_StartIndex + Lng_DataMax - 1

            '--Read unique reference using the column number stored in StartCols(1)
            Str_UniqueRef = Wsh_Source.Cells(Lng_Record, Rng_StartCols.Cells(1, 1).Value).Value

            If Str_UniqueRef <> "" Then

                '--Write submission ID to column H and unique reference to column J on Home
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

            End If

        Next Lng_Record

    End If

    '--Close source file
    Wbk_Source.Close SaveChanges:=False

End Sub
