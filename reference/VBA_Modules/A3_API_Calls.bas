Attribute VB_Name = "A3_API_Calls"
Option Explicit
Option Base 1

Sub PostSurveyData()

    Dim Str_CaseCode As String
    Dim Bln_ValidatedCaseCode As Boolean
    Dim Wsh_Home As Worksheet:                  Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Wsh_Dropdowns As Worksheet:             Set Wsh_Dropdowns = ThisWorkbook.Worksheets("Drop downs")
    Dim Rng_DropDownQs As Range:                Set Rng_DropDownQs = Wsh_Dropdowns.Range("DropDownQs")
    Dim Lng_QuestionLookupCol As Long
    Dim Rng_LookupRange As Range
    Dim Str_QResponse As String
    Dim Var_LookupResponseId As Variant
    Dim Lng_LastRow As Long:                    Lng_LastRow = Wsh_Home.Cells(Wsh_Home.Rows.Count, 8).End(xlUp).Row
    Dim Rng_DataRows As Range:                  Set Rng_DataRows = Wsh_Home.Range(Wsh_Home.Cells(6, 10), Wsh_Home.Cells(Lng_LastRow, 10))
    Dim Rng_QuestionCols As Range:              Set Rng_QuestionCols = Wsh_Home.Range("QuestionCols")
    Dim Lng_OrgId As Long
    Dim Str_SubID As String
    Dim Str_ServiceId As String:                Str_ServiceId = Wsh_Home.Range("ServiceID").Value
    Dim Rng_Cell As Range
    Dim Rng_Cell2 As Range
    Dim i As Long
    Dim Var_ResponsesArray() As Variant
    Dim Var_FinalArray() As Variant
    Dim Lng_Row As Long
    Dim Var_Result As VbMsgBoxResult:           Var_Result = vbYes

    '--We take the rows that need to be imported and create an array of question ids and list item id responses

    For Each Rng_Cell In Rng_DataRows
        Var_Result = vbYes
        If Rng_Cell.Offset(0, -4).Value = "Yes" Then
            If Rng_Cell.Offset(0, -1).Value <> "" Then
                Var_Result = MsgBox("Row " & Rng_Cell.Row & " contains a value in the case code column. Do you wish to import the row?", vbYesNo, "Confirm Import")
            End If
            If Var_Result = vbYes Then
            '--Reset array for each row
                ReDim Var_ResponsesArray(1 To 3, 1 To 1)
                i = 0
                For Each Rng_Cell2 In Rng_QuestionCols
                    If Rng_Cell2.Value <> "QID (hide)" Then
                        Lng_OrgId = CInt(Trim(Split(Rng_Cell.Offset(0, -3).Value, "-")(0)))
                        Str_SubID = Rng_Cell.Offset(0, -2).Value
                        
                        Select Case Rng_Cell2.Offset(-1, 0).Value
                        
                        Case "LS"
                        
                        Str_QResponse = Wsh_Home.Cells(Rng_Cell.Row, Rng_Cell2.Column).Value
                        If Str_QResponse <> "" Then
                            i = i + 1
                            '--Resize LAST dimension (allowed with Preserve)
                            ReDim Preserve Var_ResponsesArray(1 To 3, 1 To i)
                            Var_ResponsesArray(1, i) = Wsh_Home.Cells(4, Rng_Cell2.Column).Value
                            
                            Lng_QuestionLookupCol = Rng_DropDownQs.Find(what:=Rng_Cell2.Value, LookIn:=xlValues, Lookat:=xlWhole).Column
                            Set Rng_LookupRange = Wsh_Dropdowns.Range(Wsh_Dropdowns.Cells(2, Lng_QuestionLookupCol), Wsh_Dropdowns.Cells(50, Lng_QuestionLookupCol))
                            Var_LookupResponseId = Application.WorksheetFunction.XLookup(Trim(Str_QResponse), Rng_LookupRange, Rng_LookupRange.Offset(0, 1), "")
                            Var_ResponsesArray(2, i) = Var_LookupResponseId
                            Var_ResponsesArray(3, i) = "list"
                        End If
                        
                        Case "YN"
                        
                        Str_QResponse = Wsh_Home.Cells(Rng_Cell.Row, Rng_Cell2.Column).Value
                        If Str_QResponse <> "" Or Str_QResponse = "Y" Or Str_QResponse = "N" Then
                            i = i + 1
                            '--Resize LAST dimension (allowed with Preserve)
                            ReDim Preserve Var_ResponsesArray(1 To 3, 1 To i)
                            Var_ResponsesArray(1, i) = Wsh_Home.Cells(4, Rng_Cell2.Column).Value
                            If Str_QResponse = "Yes" Then
                                Str_QResponse = """Y"""
                            Else
                                Str_QResponse = """N"""
                            End If
                            Var_ResponsesArray(2, i) = Str_QResponse
                            Var_ResponsesArray(3, i) = "yn"
                        End If
                            
                        Case "N"
                        
                        Str_QResponse = Wsh_Home.Cells(Rng_Cell.Row, Rng_Cell2.Column).Value
                        If Str_QResponse <> "" Then
                            i = i + 1
                            '--Resize LAST dimension (allowed with Preserve)
                            ReDim Preserve Var_ResponsesArray(1 To 3, 1 To i)
                            Var_ResponsesArray(1, i) = Wsh_Home.Cells(4, Rng_Cell2.Column).Value
                            Var_ResponsesArray(2, i) = Str_QResponse
                            Var_ResponsesArray(3, i) = "number"
                        End If
                        End Select
                    End If
                 
                    
                Next Rng_Cell2
            End If
        '--Each stage only proceeds if the previous one was successful
        '--We API to the database to get the next available casecode which is then created
        If Rng_Cell.Offset(0, -4).Value = "Yes" And Var_Result = vbYes Then
            Str_CaseCode = RetrieveNextCaseCode(Lng_OrgId, Str_SubID)
            If Str_CaseCode = "" Then
                MsgBox "Unable to retrieve next casecode, please ensure you have selected a submission", vbCritical, "API Failure"
                Exit Sub
            Else
                'application transpose is causing weirdness when the multi array only has one response.
                'It converts the array to a single dimension which causes subsequent rows to fail
                Dim x As Long

                ReDim Var_FinalArray(1 To i, 1 To 3)
                
                For x = 1 To i
                    Var_FinalArray(x, 1) = Var_ResponsesArray(1, x)
                    Var_FinalArray(x, 2) = Var_ResponsesArray(2, x)
                    Var_FinalArray(x, 3) = Var_ResponsesArray(3, x)
                Next x
                
                If APISurveyData(Str_CaseCode, Str_SubID, Str_ServiceId, Var_FinalArray) = True Then
                    '--Next, we mark the case code as complete
                    If API_CloseCaseCode(Str_SubID, Str_CaseCode) = True Then
                        Rng_Cell.Offset(0, -1).Value = Str_CaseCode
                    Else
                        MsgBox "Unable to set case code to Completed", vbCritical, "API Failure"
                        Exit Sub
                    End If
                Else
                    MsgBox "Your case code has been created and validated but unable to submit survey data through API", vbCritical, "API Failure"
                    Exit Sub
                End If
            End If
        End If
    End If
    Rng_Cell.Offset(0, -4).Value = "No"
    Next Rng_Cell
    MsgBox "Import Complete", vbInformation, "Import Complete"

End Sub

Function RetrieveNextCaseCode(Lng_OrgId As Long, Str_SubmissionId As String) As String

    Dim Wsh_Home As Worksheet:              Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Lng_ProjectId As Long:              Lng_ProjectId = Wsh_Home.Range("ProjectID").Value
    'Dim Rng_SubmissionList As Range:        Set Rng_SubmissionList = Wsh_Home.Range("Submissions")
    Dim Str_CaseCode As String
    
    'Str_SubmissionId = Rng_SubmissionList.Value
    If Str_SubmissionId <> "" Then
        'Str_SubmissionId = Split(Rng_SubmissionList.Value, "-")(0)
    Else
        Exit Function
    End If
    
    Str_CaseCode = API_GetNextCaseCode(Str_SubmissionId, Lng_OrgId)
    
    RetrieveNextCaseCode = Str_CaseCode
    
End Function


Function APISurveyData(Str_CaseCode As String, Str_SubID As String, Str_ServiceId As String, Var_ResponsesArray() As Variant) As Boolean
    
    If API_PostSurvey(Str_CaseCode, Str_SubID, Str_ServiceId, Var_ResponsesArray) = True Then
        APISurveyData = True
    End If
    
End Function

Function APICloseCaseCode(Str_SubID As String, Str_CaseCode As String) As Boolean
    
    If API_CloseCaseCode(Str_SubID, Str_CaseCode) = True Then
        APICloseCaseCode = True
    End If
    
End Function
