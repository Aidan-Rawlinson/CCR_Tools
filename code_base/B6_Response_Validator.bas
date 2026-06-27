Attribute VB_Name = "B6_Response_Validator"
Option Explicit

' ============================================================
' B6_Response_Validator
' ------------------------------------------------------------
' Validates responses on the Home sheet for rows imported in
' the current run. Called by B4_Process_Folder after all files
' have been processed, passing the first and last row numbers
' written during the run.
'
' Validation rules:
'   LS  -- response text must match a valid option in the Drop
'           downs sheet for that question; no match → orange
'   YN  -- value must be exactly "Yes" or "No"; anything else → orange
'   N   -- cell value must be numeric; fails → orange
'   TX  -- must not be numeric; any non-blank text string is valid
'   DT  -- must be a numeric Excel date serial within the
'           valid range (1 Jun–31 Aug 2026: serials 46174–46269);
'           non-numeric or out-of-range → orange
'
' Orange cells must be corrected by the user before posting.
'
' Drop downs sheet structure:
'   Row 1:    QIDs in odd columns (A, C, E...)
'   Row 2:    Question labels in odd columns
'   Row 3+:   List item IDs in odd columns; response text in
'             adjacent even columns (B, D, F...)
'   Scan column = QID match column + 1 (the text column)
'
' Error Log sheet:
'   Cleared and re-headed at the start of each run.
'   One row per error: Row | Unique Ref | Question ID |
'                      Question No. | Question Type | Invalid Value
'
' Summary MsgBox on completion: count of errors found.
' ============================================================

Sub ValidateResponses(Lng_FirstRow As Long, Lng_LastRow As Long)

    Dim Wsh_Home As Worksheet:          Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Wsh_Dropdowns As Worksheet:     Set Wsh_Dropdowns = ThisWorkbook.Worksheets("Drop downs")
    Dim Wsh_ErrorLog As Worksheet:      Set Wsh_ErrorLog = ThisWorkbook.Worksheets("Error Log")

    Dim Rng_QuestionCols As Range:      Set Rng_QuestionCols = ThisWorkbook.Names("QuestionCols").RefersToRange
    Dim Rng_TypeCols As Range:          Set Rng_TypeCols = ThisWorkbook.Names("TypeCols").RefersToRange
    Dim Rng_DropDownQs As Range:        Set Rng_DropDownQs = Wsh_Dropdowns.Range("DropDownQs")
    Dim Rng_DataArea As Range:          Set Rng_DataArea = ThisWorkbook.Names("DataArea").RefersToRange

    '--DT valid serial range: 1 Jun 2026 = 46174, 31 Aug 2026 = 46269
    Const Lng_DT_Min As Long = 46174
    Const Lng_DT_Max As Long = 46269

    '--Error log setup: clear sheet and write headers
    Wsh_ErrorLog.Cells.ClearContents
    Wsh_ErrorLog.Cells.Interior.ColorIndex = xlNone
    With Wsh_ErrorLog
        .Cells(1, 1).Value = "Row"
        .Cells(1, 2).Value = "Unique Ref"
        .Cells(1, 3).Value = "Question ID"
        .Cells(1, 4).Value = "Question No."
        .Cells(1, 5).Value = "Question Type"
        .Cells(1, 6).Value = "Invalid Value"
    End With
    Dim Lng_LogRow As Long:             Lng_LogRow = 2

    '--Error counter
    Dim Lng_ErrorCount As Long:         Lng_ErrorCount = 0

    '--Orange colour for invalid cells
    Dim Lng_Orange As Long:             Lng_Orange = RGB(255, 192, 0)

    '--Loop over imported rows
    Dim Lng_Row As Long
    For Lng_Row = Lng_FirstRow To Lng_LastRow

        '--Read unique reference for this row (DataArea column = col K)
        Dim Str_UniqueRef As String
        Str_UniqueRef = Wsh_Home.Cells(Lng_Row, Rng_DataArea.Column).Value

        '--Loop over question columns in parallel with type columns
        Dim Rng_QCell As Range
        For Each Rng_QCell In Rng_QuestionCols

            '--Get type code for this column (TypeCols is same width, same column order)
            Dim Str_TypeCode As String
            Str_TypeCode = Rng_TypeCols.Cells(1, Rng_QCell.Column - Rng_QuestionCols.Column + 1).Value

            '--Get the response value for this cell
            Dim Str_Response As String
            Str_Response = Wsh_Home.Cells(Lng_Row, Rng_QCell.Column).Value

            '--Skip blank responses -- no validation needed
            If Str_Response = "" Then GoTo NextQuestion

            '--Flag to track whether this cell needs logging
            Dim Bln_Error As Boolean
            Bln_Error = False

            Select Case Str_TypeCode

            Case "LS"

                '--Look up QID in DropDownQs row 1 to find the correct column pair
                '--Odd column = list item IDs; even column (+1) = response text
                Dim Rng_QIDMatch As Range
                Set Rng_QIDMatch = Nothing
                On Error Resume Next
                Set Rng_QIDMatch = Rng_DropDownQs.Find(what:=Rng_QCell.Value, LookIn:=xlValues, Lookat:=xlWhole)
                On Error GoTo 0

                Dim Bln_Valid As Boolean
                Bln_Valid = False

                If Not Rng_QIDMatch Is Nothing Then
                    '--Scan the response text column (QID column + 1) from row 3 downwards
                    Dim Lng_TextCol As Long:    Lng_TextCol = Rng_QIDMatch.Column + 1
                    Dim Lng_ScanRow As Long:    Lng_ScanRow = 3
                    Do While Wsh_Dropdowns.Cells(Lng_ScanRow, Lng_TextCol).Value <> ""
                        If Trim(Wsh_Dropdowns.Cells(Lng_ScanRow, Lng_TextCol).Value) = Trim(Str_Response) Then
                            Bln_Valid = True
                            Exit Do
                        End If
                        Lng_ScanRow = Lng_ScanRow + 1
                    Loop
                End If

                If Not Bln_Valid Then Bln_Error = True

            Case "YN"

                '--Must be exactly "Yes" or "No"
                If Str_Response <> "Yes" And Str_Response <> "No" Then Bln_Error = True

            Case "N"

                '--Must be numeric
                If Not IsNumeric(Str_Response) Then Bln_Error = True

            Case "TX"

                '--Must not be numeric; any non-blank text string is valid
                If IsNumeric(Str_Response) Then Bln_Error = True

            Case "DT"

                '--Must be a numeric Excel date serial within 1 Jun–31 Aug 2026
                If Not IsNumeric(Str_Response) Then
                    Bln_Error = True
                ElseIf CLng(Str_Response) < Lng_DT_Min Or CLng(Str_Response) > Lng_DT_Max Then
                    Bln_Error = True
                End If

            End Select

            '--Log and colour any failed cell
            If Bln_Error Then
                Wsh_Home.Cells(Lng_Row, Rng_QCell.Column).Interior.Color = Lng_Orange
                Wsh_ErrorLog.Cells(Lng_LogRow, 1).Value = Lng_Row
                Wsh_ErrorLog.Cells(Lng_LogRow, 2).Value = Str_UniqueRef
                Wsh_ErrorLog.Cells(Lng_LogRow, 3).Value = Rng_QCell.Value
                Wsh_ErrorLog.Cells(Lng_LogRow, 4).Value = Wsh_Home.Cells(Rng_TypeCols.Row - 1, Rng_QCell.Column).Value
                Wsh_ErrorLog.Cells(Lng_LogRow, 5).Value = Str_TypeCode
                Wsh_ErrorLog.Cells(Lng_LogRow, 6).Value = Str_Response
                Lng_LogRow = Lng_LogRow + 1
                Lng_ErrorCount = Lng_ErrorCount + 1
            End If

NextQuestion:
        Next Rng_QCell

    Next Lng_Row

    '--Summary message
    If Lng_ErrorCount = 0 Then
        MsgBox "Validation complete — no errors found.", _
               vbInformation, "Response Validation"
    Else
        MsgBox "Validation complete — " & Lng_ErrorCount & " error(s) found." & vbCrLf & vbCrLf & _
               "Invalid cells are highlighted orange on the Home sheet." & vbCrLf & _
               "See the Error Log sheet for full details.", _
               vbExclamation, "Response Validation"
    End If

End Sub
