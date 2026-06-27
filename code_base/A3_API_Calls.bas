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
    Dim Rng_FullDataArea As Range:              Set Rng_FullDataArea = ThisWorkbook.Names("FullDataArea").RefersToRange
    Dim Rng_DataRows As Range:                  Set Rng_DataRows = Rng_FullDataArea.Columns(1).Cells
    Dim Rng_QuestionCols As Range:              Set Rng_QuestionCols = ThisWorkbook.Names("QuestionCols").RefersToRange
    Dim Rng_TypeCols As Range:                  Set Rng_TypeCols = ThisWorkbook.Names("TypeCols").RefersToRange
    Dim Lng_OrgId As Long
    Dim Str_SubID As String
    Dim Str_ServiceId As String:                Str_ServiceId = ThisWorkbook.Names("ServiceID").RefersToRange.Value
    Dim Rng_Cell As Range
    Dim Rng_Cell2 As Range
    Dim Rng_TypeCell As Range
    Dim i As Long
    Dim Var_ResponsesArray() As Variant
    Dim Var_FinalArray() As Variant
    Dim Lng_Row As Long
    Dim Var_Result As VbMsgBoxResult:           Var_Result = vbYes

    For Each Rng_Cell In Rng_DataRows
        If Rng_Cell.Cells(1).Value = "" Then Exit For
        Var_Result = vbYes
        If Rng_Cell.Cells(1).Value = "Yes" Then
            If Rng_Cell.Cells(1).Offset(0, 5).Value <> "" Then
                Var_Result = MsgBox("Row " & Rng_Cell.Cells(1).Row & " contains a value in the case code column. Do you wish to import the row?", vbYesNo, "Confirm Import")
            End If
            If Var_Result = vbYes Then
                ReDim Var_ResponsesArray(1 To 3, 1 To 1)
                i = 0

                Lng_OrgId = CLng(Rng_Cell.Cells(1).Offset(0, 3).Value)
                Str_SubID = CStr(Rng_Cell.Cells(1).Offset(0, 4).Value)

                For Each Rng_Cell2 In Rng_QuestionCols
                    Set Rng_TypeCell = Rng_TypeCols.Cells(1, Rng_Cell2.Column - Rng_QuestionCols.Column + 1)

                    If Rng_Cell2.Value <> "QID (hide)" Then

                        Select Case Rng_TypeCell.Value

                        Case "LS"

                        '--CStr() used on both response and lookup values to handle numeric list items
                        '--e.g. Rockwood scores which Excel may store as numbers
                        Str_QResponse = CStr(Wsh_Home.Cells(Rng_Cell.Cells(1).Row, Rng_Cell2.Column).Value)
                        If Str_QResponse <> "" Then
                            i = i + 1
                            ReDim Preserve Var_ResponsesArray(1 To 3, 1 To i)
                            Var_ResponsesArray(1, i) = Rng_Cell2.Value
                            Lng_QuestionLookupCol = Rng_DropDownQs.Find(what:=Rng_Cell2.Value, LookIn:=xlValues, Lookat:=xlWhole).Column
                            Set Rng_LookupRange = Wsh_Dropdowns.Range(Wsh_Dropdowns.Cells(3, Lng_QuestionLookupCol + 1), Wsh_Dropdowns.Cells(200, Lng_QuestionLookupCol + 1))
                            Var_LookupResponseId = Application.WorksheetFunction.XLookup(Trim(Str_QResponse), Rng_LookupRange, Rng_LookupRange.Offset(0, -1), "")
                            Var_ResponsesArray(2, i) = Var_LookupResponseId
                            Var_ResponsesArray(3, i) = "list"
                        End If

                        Case "YN"

                        Str_QResponse = Wsh_Home.Cells(Rng_Cell.Cells(1).Row, Rng_Cell2.Column).Value
                        If Str_QResponse <> "" Then
                            i = i + 1
                            ReDim Preserve Var_ResponsesArray(1 To 3, 1 To i)
                            Var_ResponsesArray(1, i) = Rng_Cell2.Value
                            If Str_QResponse = "Yes" Then
                                Str_QResponse = """Y"""
                            Else
                                Str_QResponse = """N"""
                            End If
                            Var_ResponsesArray(2, i) = Str_QResponse
                            Var_ResponsesArray(3, i) = "yn"
                        End If

                        Case "N"

                        Str_QResponse = Wsh_Home.Cells(Rng_Cell.Cells(1).Row, Rng_Cell2.Column).Value
                        If Str_QResponse <> "" Then
                            i = i + 1
                            ReDim Preserve Var_ResponsesArray(1 To 3, 1 To i)
                            Var_ResponsesArray(1, i) = Rng_Cell2.Value
                            Var_ResponsesArray(2, i) = Str_QResponse
                            Var_ResponsesArray(3, i) = "number"
                        End If

                        Case "TX"

                        Str_QResponse = Wsh_Home.Cells(Rng_Cell.Cells(1).Row, Rng_Cell2.Column).Value
                        If Str_QResponse <> "" Then
                            i = i + 1
                            ReDim Preserve Var_ResponsesArray(1 To 3, 1 To i)
                            Var_ResponsesArray(1, i) = Rng_Cell2.Value
                            Var_ResponsesArray(2, i) = """" & Str_QResponse & """"
                            Var_ResponsesArray(3, i) = "text"
                        End If

                        Case "DT"

                        '--Cell holds YYYY-MM-DD 00:00:00.000 string written by B6a_DT_Converter
                        '--Pass directly as a quoted string; skip if blank
                        Str_QResponse = Wsh_Home.Cells(Rng_Cell.Cells(1).Row, Rng_Cell2.Column).Value
                        If Str_QResponse <> "" Then
                            i = i + 1
                            ReDim Preserve Var_ResponsesArray(1 To 3, 1 To i)
                            Var_ResponsesArray(1, i) = Rng_Cell2.Value
                            Var_ResponsesArray(2, i) = """" & Str_QResponse & """"
                            Var_ResponsesArray(3, i) = "date"
                        End If

                        End Select
                    End If

                Next Rng_Cell2
            End If

            If Rng_Cell.Cells(1).Value = "Yes" And Var_Result = vbYes Then
                Str_CaseCode = RetrieveNextCaseCode(Lng_OrgId, Str_SubID)
                If Str_CaseCode = "" Then
                    MsgBox "Unable to retrieve next casecode, please ensure you have selected a submission", vbCritical, "API Failure"
                    Exit Sub
                Else
                    Dim x As Long
                    ReDim Var_FinalArray(1 To i, 1 To 3)
                    For x = 1 To i
                        Var_FinalArray(x, 1) = Var_ResponsesArray(1, x)
                        Var_FinalArray(x, 2) = Var_ResponsesArray(2, x)
                        Var_FinalArray(x, 3) = Var_ResponsesArray(3, x)
                    Next x

                    If APISurveyData(Str_CaseCode, Str_SubID, Str_ServiceId, Var_FinalArray) = True Then
                        If API_CloseCaseCode(Str_SubID, Str_CaseCode) = True Then
                            Rng_Cell.Cells(1).Offset(0, 5).Value = Str_CaseCode
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
        If Rng_Cell.Cells(1).Value = "Yes" Then Rng_Cell.Cells(1).Value = "No"
    Next Rng_Cell
    MsgBox "Import Complete", vbInformation, "Import Complete"

End Sub

Function RetrieveNextCaseCode(Lng_OrgId As Long, Str_SubmissionId As String) As String

    Dim Lng_ProjectId As Long:  Lng_ProjectId = ThisWorkbook.Names("ProjectID").RefersToRange.Value
    Dim Str_CaseCode As String

    If Str_SubmissionId <> "" Then
        Str_CaseCode = API_GetNextCaseCode(Str_SubmissionId, Lng_OrgId)
    End If

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
