Attribute VB_Name = "B1_Importer"
Option Explicit

Sub FileImporter()

    Dim Wsh_Home As Worksheet:              Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Str_FolderPath As String:           Str_FolderPath = Wsh_Home.Range("SubmissionFolder").Value
    Dim Str_FileName As String
    Dim Wbk_Source As Workbook
    Dim Wsh_Source As Worksheet
    Dim Str_SourceOrg As String
    Dim Str_SourceSubID As String
    Dim Lng_LastRow As Long
    Dim Lng_ColLastRow As Long
    Dim Lng_PasteStartRow As Long
    Dim i As Long
    Dim Rng_Copy As Range
    Dim Rng_Paste As Range
    Dim Rng_Validation As Range

    'Ensure folder path ends with "\"
    If Right(Str_FolderPath, 1) <> "\" Then
        Str_FolderPath = Str_FolderPath & "\"
    End If

    'Clear previous contents once at the start
    Wsh_Home.Range("F6:CN100000").Clear

    'Start pasting at row 5
    Lng_PasteStartRow = 6

    'Loop through all Excel files in folder
    Str_FileName = Dir(Str_FolderPath & "*.xls*")

    Do While Str_FileName <> ""

        Set Wbk_Source = Workbooks.Open(Str_FolderPath & Str_FileName)
        DoEvents
        Set Wsh_Source = Wbk_Source.Worksheets("Bed based CCR")
        Str_SourceOrg = Wsh_Source.Range("B5").Value
        Str_SourceSubID = Wsh_Source.Range("B6").Value
        
        Lng_LastRow = 0

        'Find last row across columns B:S
        For i = 2 To 88 'B:CE

            Lng_ColLastRow = Wsh_Source.Cells(Wsh_Source.Rows.Count, i).End(xlUp).Row

            If Lng_ColLastRow > Lng_LastRow Then
                Lng_LastRow = Lng_ColLastRow
            End If

        Next i

        'Only process if data exists
        If Lng_LastRow >= 11 Then

            'Define copy range
            Set Rng_Copy = Wsh_Source.Range("A11:CJ" & Lng_LastRow)

            'Define paste range
            Set Rng_Paste = Wsh_Home.Range("J" & Lng_PasteStartRow).Resize(Rng_Copy.Rows.Count, Rng_Copy.Columns.Count)

            'Paste values
            Rng_Paste.Value = Rng_Copy.Value

            'Populate Source Org in column G
            Wsh_Home.Range("G" & Lng_PasteStartRow).Resize(Rng_Copy.Rows.Count, 1).Value = Str_SourceOrg

            'Populate Source Submission ID in column H
            Wsh_Home.Range("H" & Lng_PasteStartRow).Resize(Rng_Copy.Rows.Count, 1).Value = Str_SourceSubID

            'Full formatting range (F:CN)
            Set Rng_Paste = Wsh_Home.Range(Wsh_Home.Cells(Lng_PasteStartRow, 6), Wsh_Home.Cells(Lng_PasteStartRow + Rng_Copy.Rows.Count - 1, 97))

            'Formatting
            With Rng_Paste
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Font.Size = 9
            End With

            'Validation range (column F)
            Set Rng_Validation = Wsh_Home.Range("F" & Lng_PasteStartRow) _
                .Resize(Rng_Copy.Rows.Count, 1)

            'Apply Yes/No validation
            With Rng_Validation.Validation
                .Delete
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, _
                     Formula1:="Yes,No"

                .IgnoreBlank = True
                .InCellDropdown = True
            End With

            'Borders on validation column
            With Rng_Validation.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With

            'Move next paste position down
            Lng_PasteStartRow = Lng_PasteStartRow + Rng_Copy.Rows.Count

        End If

        'Close source workbook
        Wbk_Source.Close SaveChanges:=False

        'Next file
        Str_FileName = Dir()

    Loop

    MsgBox "Process Complete", vbInformation, "Process Complete"

End Sub

Function CaseCodeProcessed(Str_SubmissionId As String) As Boolean '--We are checking to see if the Unique Ref value on the worksheet exists in a casecode note field against the submission

    Dim Wsh_Home As Worksheet:              Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Lng_LastRow As Long:                Lng_LastRow = Wsh_Home.Cells(Wsh_Home.Rows.Count, 8).End(xlUp).Row
    Dim Rng_DataRows As Range:              Set Rng_DataRows = Wsh_Home.Range(Wsh_Home.Cells(5, 8), Wsh_Home.Cells(Lng_LastRow, 8))
    Dim Rng_Cell As Range
    Dim Var_Array() As Variant
    Dim Str_CaseCodeString As String
    Dim i As Long
    
    CaseCodeProcessed = False
    
    Var_Array = GetCaseCodeNote(Str_SubmissionId)
    
    If Not Var_Array(1, 1) = "Error" Then
    
        For Each Rng_Cell In Rng_DataRows
            Str_CaseCodeString = ""
            For i = LBound(Var_Array) To UBound(Var_Array)
                If Rng_Cell.Value = CInt(Var_Array(i, 1)) Then
                    If Str_CaseCodeString = "" Then
                        Str_CaseCodeString = Var_Array(i, 2)
                    Else
                        Str_CaseCodeString = Str_CaseCodeString & "^" & Var_Array(i, 2)
                    End If
                    Rng_Cell.Offset(0, -1).Value = Str_CaseCodeString
                End If
            Next i
            If Rng_Cell.Offset(0, -1).Value <> "" Then
                Rng_Cell.Offset(0, -2).Value = "No"
            Else
                Rng_Cell.Offset(0, -2).Value = "Yes"
            End If
        Next Rng_Cell
        CaseCodeProcessed = True
    Else
    CaseCodeProcessed = False
    End If
    
End Function

Sub QuestionResponseMatcher(Str_SubmissionId As String) 'Once we have established whether the unique ref is on the database, we will see if the responses match

    Dim Wsh_Home As Worksheet:              Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Lng_LastRow As Long:                Lng_LastRow = Wsh_Home.Cells(Wsh_Home.Rows.Count, 8).End(xlUp).Row
    Dim Rng_DataRows As Range:              Set Rng_DataRows = Wsh_Home.Range(Wsh_Home.Cells(5, 8), Wsh_Home.Cells(Lng_LastRow, 8))
    Dim Lng_ProjectId As Long:              Lng_ProjectId = Wsh_Home.Range("ProjectID").Value
    Dim Rng_Cell As Range
    Dim Rng_Cell2 As Range
    Dim Var_Array() As Variant
    Dim Str_CaseCodeString As String
    Dim i As Long
    Dim Rng_QuestionCols As Range:          Set Rng_QuestionCols = Wsh_Home.Range("QuestionCols")
    Dim Str_QuestionId As String
    Dim Str_ResponseId As String
    Dim Str_Response As String
    Dim Wsh_Dropdowns As Worksheet:         Set Wsh_Dropdowns = ThisWorkbook.Worksheets("Drop downs")
    Dim Rng_DropDownQs As Range:            Set Rng_DropDownQs = Wsh_Dropdowns.Range("DropDownQs")
    Dim Lng_QuestionLookupCol As Long
    Dim Rng_LookupRange As Range
    Dim Str_QResponse As String
    
    
    For Each Rng_Cell In Rng_DataRows
        Str_CaseCodeString = Rng_Cell.Offset(0, -1).Value
        If Str_CaseCodeString <> "" And InStr(Str_CaseCodeString, "^") < 1 Then 'we won't run this block if the row either has no case code or has more than one
            Var_Array = GetCaseCodeResponses(Str_SubmissionId, Lng_ProjectId, Str_CaseCodeString)
            
            For Each Rng_Cell2 In Rng_QuestionCols
                For i = LBound(Var_Array) To UBound(Var_Array)
                    If Var_Array(i, 1) = Rng_Cell2.Value Then
                        Str_QuestionId = Var_Array(i, 1)
                        Str_ResponseId = Var_Array(i, 2)
                        Exit For
                    End If
                Next i
                Lng_QuestionLookupCol = Rng_DropDownQs.Find(what:=Str_QuestionId, LookIn:=xlValues, Lookat:=xlWhole).Column + 1
                Set Rng_LookupRange = Wsh_Dropdowns.Range(Wsh_Dropdowns.Cells(4, Lng_QuestionLookupCol), Wsh_Dropdowns.Cells(20, Lng_QuestionLookupCol))
                Str_QResponse = Rng_LookupRange.Find(what:=Str_ResponseId, LookIn:=xlValues, Lookat:=xlWhole).Offset(0, -1).Value
                If LCase(Trim(Str_QResponse)) = LCase(Trim(Wsh_Home.Cells(Rng_Cell.Row, Rng_Cell2.Column).Value)) Then
                    Wsh_Home.Cells(Rng_Cell.Row, Rng_Cell2.Column).Interior.Color = RGB(193, 240, 200)
                Else
                    Wsh_Home.Cells(Rng_Cell.Row, Rng_Cell2.Column).Interior.Color = RGB(241, 169, 131)
                End If
            Next Rng_Cell2
            
        End If
    Next Rng_Cell

End Sub

Sub ResponseValidator()

    Dim Wsh_Home As Worksheet:                  Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Wsh_Dropdowns As Worksheet:             Set Wsh_Dropdowns = ThisWorkbook.Worksheets("Drop downs")
    Dim Rng_Questions As Range:                 Set Rng_Questions = Wsh_Home.Range("QuestionCols")
    Dim Rng_QuestionOptions As Range:           Set Rng_QuestionOptions = Wsh_Dropdowns.Range("DropDownQs")
    Dim Lng_LookupQuestionCol As Long
    Dim Rng_QuestionLookupRange As Range
    Dim Rng_Cell As Range
    Dim Rng_Cell2 As Range
    Dim Lng_QuestionId As Long
    Dim Lng_LastRow As Long
    Dim Rng_QuestionResponseRange As Range
    Dim Str_Response As String
    Dim Str_FoundResponse As String
    Dim Bln_ValidResponse As Boolean:           Bln_ValidResponse = False
    
    For Each Rng_Cell In Rng_Questions
        Lng_QuestionId = Rng_Cell.Value
        If Lng_QuestionId = 150434 Then
            Debug.Print Lng_QuestionId
        End If
        Lng_LookupQuestionCol = Rng_QuestionOptions.Find(what:=Lng_QuestionId, LookIn:=xlValues, Lookat:=xlWhole).Column
        Set Rng_QuestionLookupRange = Wsh_Dropdowns.Range(Wsh_Dropdowns.Cells(4, Lng_LookupQuestionCol), Wsh_Dropdowns.Cells(20, Lng_LookupQuestionCol))
        Lng_LastRow = Wsh_Home.Cells(Wsh_Home.Rows.Count, Rng_Cell.Column).End(xlUp).Row
        Set Rng_QuestionResponseRange = Wsh_Home.Range(Wsh_Home.Cells(5, Rng_Cell.Column), Wsh_Home.Cells(Lng_LastRow, Rng_Cell.Column))
        For Each Rng_Cell2 In Rng_QuestionResponseRange
            Str_FoundResponse = ""
            Str_Response = Trim(Rng_Cell2.Value)
            On Error Resume Next
            Str_FoundResponse = Trim(Rng_QuestionLookupRange.Find(what:=Str_Response, LookIn:=xlValues, Lookat:=xlWhole).Value)
            On Error GoTo 0
            If Str_FoundResponse = "" Then
                Rng_Cell2.Interior.Color = RGB(233, 113, 50)
            End If
        Next Rng_Cell2
    Next Rng_Cell
    
End Sub
