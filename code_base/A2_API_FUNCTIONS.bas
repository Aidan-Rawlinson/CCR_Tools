Attribute VB_Name = "A2_API_FUNCTIONS"
Option Explicit
Option Compare Text
Option Private Module
Option Base 1

'###########################################
'# USES TOKEN FROM A1_API_SUPPORT          #
'# USES JSON CONVERTER FROM A1_API_SUPPORT #
'###########################################

Public Function API_GetSubmissions(Lng_ProjectId As Long, Lng_OrgId As Long) As Variant

    Dim Str_TESTURLBase As String:  Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/list?projectId=[ProjectID]&year=[Year]"
    Dim Str_LIVEURLBase As String:  Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/list?projectId=[ProjectID]&year=[Year]"
    Dim Str_URL As String
    Dim Obj_JSONData As Object
    Dim Str_APIResponse As String
    Dim Obj_SubmissionList As Object
    Dim Str_JsonFilename As String: Str_JsonFilename = "SubmissionData"
    Dim Var_Output() As Variant
    Dim Lng_SubmissionCount As Long: Lng_SubmissionCount = 0
    Dim i As Long
    Dim Str_Database As String:     Str_Database = ThisWorkbook.Names("Toggle").RefersToRange.Value
    Dim Str_Year As String:         Str_Year = CStr(ThisWorkbook.Names("SubmissionYear").RefersToRange.Value)

    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[ProjectID]", Lng_ProjectId)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[ProjectID]", Lng_ProjectId)
    End If
    Str_URL = Replace(Str_URL, "[Year]", Str_Year)

    Str_APIResponse = APICall(Str_URL)

    'Call SaveJSONToFile(Str_APIResponse, Str_JsonFilename)

    Set Obj_JSONData = ParseJson(Str_APIResponse)("data")
    Set Obj_SubmissionList = Obj_JSONData("submissionList")(Str_Year)

    For i = 1 To Obj_SubmissionList.Count
        If Obj_SubmissionList(i)("organisationId") = Lng_OrgId Then
            Lng_SubmissionCount = Lng_SubmissionCount + 1
            ReDim Preserve Var_Output(1 To 2, 1 To Lng_SubmissionCount)
            Var_Output(1, Lng_SubmissionCount) = Obj_SubmissionList(i)("submissionId")
            Var_Output(2, Lng_SubmissionCount) = Obj_SubmissionList(i)("submissionName")
        End If
    Next i

    If Lng_SubmissionCount = 0 Then
        API_GetSubmissions = "There are no projects!! Call the police!!"
    Else
        API_GetSubmissions = Application.Transpose(Var_Output)
    End If

End Function

'###########################################
'# USES TOKEN FROM A1_API_SUPPORT          #
'# USES JSON CONVERTER FROM A1_API_SUPPORT #
'###########################################

Public Function API_GetNextCaseCode(Str_SubmissionId As String, Lng_OrgId As Long) As String

    Dim Str_TESTURLBase As String:  Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/addCnrCodes"
    Dim Str_LIVEURLBase As String:  Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/addCnrCodes"
    Dim Str_URL As String
    Dim Str_Payload As String
    Dim Obj_JSONData As Object
    Dim Str_APIResponse As String
    Dim Str_CaseCode As String
    Dim Str_Database As String:     Str_Database = ThisWorkbook.Names("Toggle").RefersToRange.Value

    If Str_Database = "Live" Then
        Str_URL = Str_LIVEURLBase
    Else
        Str_URL = Str_TESTURLBase
    End If

    Str_Payload = "{""newCodeCount"":1}"
    Str_URL = Replace(Str_URL, "[SubmissionId]", Str_SubmissionId)
    Str_APIResponse = APIPost(Str_URL, Str_Payload)

    Set Obj_JSONData = ParseJson(Str_APIResponse)("data")("newCnrCodes")
    Str_CaseCode = Obj_JSONData(1)("caseCode")

    API_GetNextCaseCode = Str_CaseCode

End Function

'###########################################
'# USES TOKEN FROM A1_API_SUPPORT          #
'# USES JSON CONVERTER FROM A1_API_SUPPORT #
'###########################################

Public Function API_PostSurvey(Str_CaseCode As String, Str_SubID As String, Str_Service As String, Var_ResponsesArray As Variant) As Boolean

    Dim Str_TESTURLBase As String:  Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/projects/questions?submissionId=[SubmissionId]&serviceId=[ServiceID]&submissionCaseCode=[CaseCode]"
    Dim Str_LIVEURLBase As String:  Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/projects/questions?submissionId=[SubmissionId]&serviceId=[ServiceID]&submissionCaseCode=[CaseCode]"
    Dim Str_URL As String
    Dim Str_PayloadTemplate As String
    Dim Str_Payload As String
    Dim Str_Item As String
    Dim i As Long
    Dim Str_APIResponse As String
    Dim Str_Database As String:     Str_Database = ThisWorkbook.Names("Toggle").RefersToRange.Value

    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[SubmissionId]", Str_SubID)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[SubmissionId]", Str_SubID)
    End If
    Str_URL = Replace(Str_URL, "[ServiceID]", Str_Service)
    Str_URL = Replace(Str_URL, "[CaseCode]", Str_CaseCode)

    Str_PayloadTemplate = "{""questionId"":""[QuestionID]"",""questionPart"":1,""questionType"":""[Type]"",""value"":[Value]}"
    Str_Payload = "["

    For i = LBound(Var_ResponsesArray, 1) To UBound(Var_ResponsesArray, 1)
        Str_Item = Str_PayloadTemplate
        If Var_ResponsesArray(i, 2) <> "" Then
            Str_Item = Replace(Str_Item, "[QuestionID]", Var_ResponsesArray(i, 1))
            Str_Item = Replace(Str_Item, "[Value]", Var_ResponsesArray(i, 2))
            Str_Item = Replace(Str_Item, "[Type]", Var_ResponsesArray(i, 3))
            Str_Payload = Str_Payload & Str_Item
            If i < UBound(Var_ResponsesArray, 1) Then
                Str_Payload = Str_Payload & ","
            End If
        End If
    Next i

    Str_Payload = Str_Payload & "]"
    Str_APIResponse = APIPost(Str_URL, Str_Payload)

    If Str_APIResponse = "{""success"":true}" Then
        API_PostSurvey = True
    End If

End Function

Public Function API_CloseCaseCode(Str_SubID As String, Str_CaseCode As String) As Boolean

    Dim Str_TESTURLBase As String:  Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/setCaseCodeCompleted"
    Dim Str_LIVEURLBase As String:  Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/setCaseCodeCompleted"
    Dim Str_URL As String
    Dim Str_PayloadTemplate As String
    Dim Str_Payload As String
    Dim Str_APIResponse As String
    Dim Str_Database As String:     Str_Database = ThisWorkbook.Names("Toggle").RefersToRange.Value

    Str_PayloadTemplate = "{""caseCode"": ""[CaseCode]"", ""dataSubmitted"": ""Y""}"
    Str_Payload = Replace(Str_PayloadTemplate, "[CaseCode]", Str_CaseCode)

    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[SubmissionId]", Str_SubID)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[SubmissionId]", Str_SubID)
    End If

    Str_APIResponse = APIPost(Str_URL, Str_Payload)

    If Str_APIResponse = "{""success"":true}" Then
        API_CloseCaseCode = True
    End If

End Function

Public Function GetCaseCodeNote(Str_SubmissionId As String) As Variant

    Dim Str_TESTURLBase As String:  Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/caseCodes?allCaseCodes=true"
    Dim Str_LIVEURLBase As String:  Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/caseCodes?allCaseCodes=true"
    Dim Str_URL As String
    Dim Obj_JSONData As Object
    Dim Str_APIResponse As String
    Dim Obj_caseNoteCodes As Object
    Dim Str_CaseCodeNote As String
    Dim Var_Output() As Variant
    Dim Lng_CaseNoteCount As Long:  Lng_CaseNoteCount = 0
    Dim i As Long
    Dim Str_Database As String:     Str_Database = ThisWorkbook.Names("Toggle").RefersToRange.Value

    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[SubmissionId]", Str_SubmissionId)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[SubmissionId]", Str_SubmissionId)
    End If

    Str_APIResponse = APICall(Str_URL)

    'Call SaveJSONToFile(Str_APIResponse, Str_JsonFilename)

    Set Obj_JSONData = ParseJson(Str_APIResponse)("data")
    Set Obj_caseNoteCodes = Obj_JSONData("caseNoteCodes")

    For i = 1 To Obj_caseNoteCodes.Count
        If InStr(Obj_caseNoteCodes(i)("caseCodeNotes"), "externalCode") Then
            If Obj_caseNoteCodes(i)("dataSubmitted") = "True" And Obj_caseNoteCodes(i)("completionStatus") = "Completed" Then
                Lng_CaseNoteCount = Lng_CaseNoteCount + 1
                ReDim Preserve Var_Output(1 To 2, 1 To Lng_CaseNoteCount)
                Str_CaseCodeNote = Obj_caseNoteCodes(i)("caseCodeNotes")
                Var_Output(1, Lng_CaseNoteCount) = Split(Split(Str_CaseCodeNote, """externalCode"":""")(1), """")(0)
                Var_Output(2, Lng_CaseNoteCount) = Obj_caseNoteCodes(i)("caseCode")
            End If
        End If
    Next i

    If Lng_CaseNoteCount = 0 Then
        Dim Var_Error(1, 1) As Variant
        Var_Error(1, 1) = "Error"
        GetCaseCodeNote = Var_Error
    Else
        GetCaseCodeNote = Application.Transpose(Var_Output)
    End If

End Function

Public Function GetCaseCodeResponses(Str_SubmissionId As String, Lng_ProjectId As Long, Str_CaseCodeString As String) As Variant

    Dim Str_TESTURLBase As String:  Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/projects/[ProjectId]/responses?year=[Year]&submissionId=[SubmissionId]"
    Dim Str_LIVEURLBase As String:  Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/projects/[ProjectId]/responses?year=[Year]&submissionId=[SubmissionId]"
    Dim Str_URL As String
    Dim Obj_JSONData As Object
    Dim Str_APIResponse As String
    Dim Obj_Responses As Object
    Dim Var_Output() As Variant
    Dim Lng_ResponsesCount As Long: Lng_ResponsesCount = 0
    Dim i As Long
    Dim Str_Database As String:     Str_Database = ThisWorkbook.Names("Toggle").RefersToRange.Value
    Dim Str_Year As String:         Str_Year = CStr(ThisWorkbook.Names("SubmissionYear").RefersToRange.Value)

    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[ProjectId]", Lng_ProjectId)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[ProjectId]", Lng_ProjectId)
    End If
    Str_URL = Replace(Str_URL, "[Year]", Str_Year)
    Str_URL = Replace(Str_URL, "[SubmissionId]", Str_SubmissionId)

    Str_APIResponse = APICall(Str_URL)

    'Call SaveJSONToFile(Str_APIResponse, Str_JsonFilename)

    Set Obj_JSONData = ParseJson(Str_APIResponse)("data")
    Set Obj_Responses = Obj_JSONData("responseList")

    For i = 1 To Obj_Responses.Count
        If Obj_Responses(i)("caseCode") = Str_CaseCodeString Then
            Lng_ResponsesCount = Lng_ResponsesCount + 1
            ReDim Preserve Var_Output(1 To 2, 1 To Lng_ResponsesCount)
            Var_Output(1, Lng_ResponsesCount) = Obj_Responses(i)("questionId")
            Var_Output(2, Lng_ResponsesCount) = Obj_Responses(i)("itemId")
        End If
    Next i

    If Lng_ResponsesCount = 0 Then
        GetCaseCodeResponses = Var_Output
    Else
        GetCaseCodeResponses = Application.Transpose(Var_Output)
    End If

End Function

Function APICall(Str_URL As String) As String

    Dim Http As Object:     Set Http = CreateObject("MSXML2.XMLHTTP")
    Dim Token As String:    Token = GetToken()
    Dim Str_Response As String

    Http.Open "GET", (Str_URL), False
    Http.setRequestHeader "Accept", "application/json"
    Http.setRequestHeader "Token", Token
    Http.Send
    Str_Response = Http.responseText

    APICall = Str_Response

End Function

Function APIPost(Str_URL As String, Optional Str_JsonPayload As String) As String

    Dim Http As Object:     Set Http = CreateObject("MSXML2.XMLHTTP")
    Dim Token As String:    Token = GetToken()
    Dim Str_Response As String

    Http.Open "POST", Str_URL, False
    Http.setRequestHeader "Accept", "application/json"
    Http.setRequestHeader "Content-Type", "application/json"
    Http.setRequestHeader "Token", Token
    Http.Send Str_JsonPayload
    Str_Response = Http.responseText

    APIPost = Str_Response

End Function
