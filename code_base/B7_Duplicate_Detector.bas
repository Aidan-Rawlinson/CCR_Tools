Attribute VB_Name = "B7_Duplicate_Detector"
Option Explicit

' ============================================================
' B7_Duplicate_Detector
' ------------------------------------------------------------
' Runs after B6_Response_Validator on the same row range.
' Called by B4_Process_Folder with the first and last row
' numbers written during the current run.
'
' For each newly imported row, checks whether the combination
' of Submission ID (column J) + Unique Reference (column L)
' already exists in any row above Lng_FirstRow on the Home
' sheet. If a match is found, the row is treated as a
' duplicate of a previously imported record.
'
' On duplicate found:
'   - Column F set to "No"
'   - All response cells coloured green
'   - Cell-level comparison run against the matching earlier row:
'       LS  -- compare response text directly
'       All other types -- compare cell value directly
'     Mismatching cells coloured orange (overrides green)
'   - One log entry written for the duplicate record
'   - One log entry written per mismatching response cell
'
' Error Log sheet columns (appended after B6 entries):
'   Row | Unique Ref | Question ID | Question No. | Question Type | Detail
'
' The Error Log is NOT cleared by B7 -- B6 clears it at the
' start of the run. B7 appends to whatever B6 has written.
'
' Summary MsgBox on completion: count of duplicates found.
' ============================================================

Sub DetectDuplicates(Lng_FirstRow As Long, Lng_LastRow As Long)

    Dim Wsh_Home As Worksheet:          Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Wsh_ErrorLog As Worksheet:      Set Wsh_ErrorLog = ThisWorkbook.Worksheets("Error Log")

    Dim Rng_DataArea As Range:          Set Rng_DataArea = ThisWorkbook.Names("DataArea").RefersToRange
    Dim Rng_FullDataArea As Range:      Set Rng_FullDataArea = ThisWorkbook.Names("FullDataArea").RefersToRange
    Dim Rng_QuestionCols As Range:      Set Rng_QuestionCols = ThisWorkbook.Names("QuestionCols").RefersToRange
    Dim Rng_TypeCols As Range:          Set Rng_TypeCols = ThisWorkbook.Names("TypeCols").RefersToRange

    '--Column positions derived from named ranges
    Dim Lng_ColF As Long:               Lng_ColF = Rng_FullDataArea.Column                 ' Process toggle
    Dim Lng_ColJ As Long:               Lng_ColJ = Rng_DataArea.Column - 2                 ' Sub ID
    Dim Lng_ColL As Long:               Lng_ColL = Rng_DataArea.Column                     ' Unique Ref

    '--Colours
    Dim Lng_Green As Long:              Lng_Green = RGB(0, 176, 80)
    Dim Lng_Orange As Long:             Lng_Orange = RGB(255, 192, 0)

    '--Find next empty row on Error Log for appending
    Dim Lng_LogRow As Long
    Lng_LogRow = Wsh_ErrorLog.Cells(Wsh_ErrorLog.Rows.Count, 1).End(xlUp).Row + 1

    '--Duplicate counter
    Dim Lng_DupCount As Long:           Lng_DupCount = 0

    '--First data row (row above Lng_FirstRow is the earliest existing row to check against)
    Dim Lng_DataStart As Long:          Lng_DataStart = Rng_FullDataArea.Row

    '--Loop over newly imported rows
    Dim Lng_NewRow As Long
    For Lng_NewRow = Lng_FirstRow To Lng_LastRow

        Dim Str_NewSubID As String
        Dim Str_NewUniqueRef As String

        Str_NewSubID = CStr(Wsh_Home.Cells(Lng_NewRow, Lng_ColJ).Value)
        Str_NewUniqueRef = CStr(Wsh_Home.Cells(Lng_NewRow, Lng_ColL).Value)

        '--Skip if unique ref is blank
        If Str_NewUniqueRef = "" Then GoTo NextNewRow

        '--Search all rows above Lng_FirstRow for same Sub ID + Unique Ref
        Dim Lng_ExistingRow As Long
        Dim Lng_MatchRow As Long:       Lng_MatchRow = 0

        For Lng_ExistingRow = Lng_DataStart To Lng_FirstRow - 1
            If CStr(Wsh_Home.Cells(Lng_ExistingRow, Lng_ColJ).Value) = Str_NewSubID And _
               CStr(Wsh_Home.Cells(Lng_ExistingRow, Lng_ColL).Value) = Str_NewUniqueRef Then
                Lng_MatchRow = Lng_ExistingRow
                Exit For
            End If
        Next Lng_ExistingRow

        '--If a match was found, process as duplicate
        If Lng_MatchRow > 0 Then

            Lng_DupCount = Lng_DupCount + 1

            '--Set column F to No
            Wsh_Home.Cells(Lng_NewRow, Lng_ColF).Value = "No"

            '--Log the duplicate record
            Wsh_ErrorLog.Cells(Lng_LogRow, 1).Value = Lng_NewRow
            Wsh_ErrorLog.Cells(Lng_LogRow, 2).Value = Str_NewUniqueRef
            Wsh_ErrorLog.Cells(Lng_LogRow, 3).Value = ""
            Wsh_ErrorLog.Cells(Lng_LogRow, 4).Value = ""
            Wsh_ErrorLog.Cells(Lng_LogRow, 5).Value = ""
            Wsh_ErrorLog.Cells(Lng_LogRow, 6).Value = "Duplicate record -- matches row " & Lng_MatchRow
            Lng_LogRow = Lng_LogRow + 1

            '--Cell-level comparison across question columns
            Dim Rng_QCell As Range
            For Each Rng_QCell In Rng_QuestionCols

                Dim Lng_QCol As Long:       Lng_QCol = Rng_QCell.Column
                Dim Str_TypeCode As String
                Str_TypeCode = Rng_TypeCols.Cells(1, Lng_QCol - Rng_QuestionCols.Column + 1).Value

                Dim Str_NewVal As String
                Dim Str_ExistingVal As String

                Str_NewVal = CStr(Wsh_Home.Cells(Lng_NewRow, Lng_QCol).Value)
                Str_ExistingVal = CStr(Wsh_Home.Cells(Lng_MatchRow, Lng_QCol).Value)

                '--Skip blank cells in both rows
                If Str_NewVal = "" And Str_ExistingVal = "" Then GoTo NextQCell

                '--Default to green; override to orange on mismatch
                If Str_NewVal = Str_ExistingVal Then
                    Wsh_Home.Cells(Lng_NewRow, Lng_QCol).Interior.Color = Lng_Green
                Else
                    Wsh_Home.Cells(Lng_NewRow, Lng_QCol).Interior.Color = Lng_Orange
                    '--Log the mismatch
                    Wsh_ErrorLog.Cells(Lng_LogRow, 1).Value = Lng_NewRow
                    Wsh_ErrorLog.Cells(Lng_LogRow, 2).Value = Str_NewUniqueRef
                    Wsh_ErrorLog.Cells(Lng_LogRow, 3).Value = Rng_QCell.Value
                    Wsh_ErrorLog.Cells(Lng_LogRow, 4).Value = Wsh_Home.Cells(Rng_TypeCols.Row - 1, Lng_QCol).Value
                    Wsh_ErrorLog.Cells(Lng_LogRow, 5).Value = Str_TypeCode
                    Wsh_ErrorLog.Cells(Lng_LogRow, 6).Value = "Mismatch -- new: """ & Str_NewVal & """ existing: """ & Str_ExistingVal & """"
                    Lng_LogRow = Lng_LogRow + 1
                End If

NextQCell:
            Next Rng_QCell

        End If

NextNewRow:
    Next Lng_NewRow

    '--Summary message
    If Lng_DupCount = 0 Then
        MsgBox "Duplicate check complete -- no duplicates found.", _
               vbInformation, "Duplicate Detection"
    Else
        MsgBox "Duplicate check complete -- " & Lng_DupCount & " duplicate record(s) found." & vbCrLf & vbCrLf & _
               "Duplicate rows have been set to No and highlighted green." & vbCrLf & _
               "Orange cells indicate a response that differs from the earlier import." & vbCrLf & _
               "See the Error Log sheet for full details.", _
               vbExclamation, "Duplicate Detection"
    End If

End Sub
