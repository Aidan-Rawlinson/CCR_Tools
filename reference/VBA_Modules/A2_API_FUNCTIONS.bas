Attribute VB_Name = "A2_API_FUNCTIONS"
Option Explicit 'forces variables to be declared properly
Option Compare Text 'allows text with non matching capitalistion to be viewed as being the same (default VBA behaviours seem mixed)
Option Private Module 'allows macros to be used across modules, but won't show in macro list
Option Base 1 'value 1 is the first value in an array (the default VBA behaviour is value 0 is the first value)

'###########################################
'# USES TOKEN FROM A1_API_SUPPORT          #
'# USES JSON CONVERTER FROM A1_API_SUPPORT #
'###########################################

Public Function API_GetSubmissions(Lng_ProjectId As Long, Lng_OrgId As Long) As Variant

    'API Variables
    
    Dim Str_TESTURLBase As String:              Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/list?projectId=[ProjectID]&year=2026"
    Dim Str_LIVEURLBase As String:              Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/list?projectId=[ProjectID]&year=2026"
    Dim Str_URL As String
    Dim Obj_JSONData As Object
    Dim Str_APIResponse As String
    Dim Obj_SubmissionList As Object
    Dim Str_JsonFilename As String:             Str_JsonFilename = "SubmissionData"
    Dim Var_Output() As Variant
    Dim Lng_SubmissionCount As Long:            Lng_SubmissionCount = 0
    Dim i As Long
    Dim Str_Database As String:                 Str_Database = ThisWorkbook.Worksheets("Orgs").Range("Toggle").Value
    
    'Update URL
    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[ProjectID]", Lng_ProjectId)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[ProjectID]", Lng_ProjectId)
    End If

    Str_APIResponse = APICall(Str_URL)
    
    'OPTIONAL - OUTPUT THE DATA FOR REVIEW -
    
    'Call SaveJSONToFile(Str_APIResponse, Str_JsonFilename)

    'Convert the JSON - this will have two items "success" and "data" - we'll assume success to be true!
    Set Obj_JSONData = ParseJson(Str_APIResponse)("data")

    'Drill down - here knowledge of the JSON structure is useful, so use the SaveJSONToFile above to review
    Set Obj_SubmissionList = Obj_JSONData("submissionList")("2026")
    
    'populate output array
    For i = 1 To Obj_SubmissionList.Count 'cycle
        If Obj_SubmissionList(i)("organisationId") = Lng_OrgId Then 'check is visible (i.e. current)
            Lng_SubmissionCount = Lng_SubmissionCount + 1
            ReDim Preserve Var_Output(1 To 2, 1 To Lng_SubmissionCount) 'Add "column" to array (redim preserve can only add to last dimension)
            Var_Output(1, Lng_SubmissionCount) = Obj_SubmissionList(i)("submissionId")
            Var_Output(2, Lng_SubmissionCount) = Obj_SubmissionList(i)("submissionName")
        End If
    Next i

    'Output array or no services message
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

    'API Variables
    
    Dim Str_TESTURLBase As String:          Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/addCnrCodes"
    Dim Str_LIVEURLBase As String:          Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/addCnrCodes"
    Dim Str_URL As String
    Dim Str_Payload As String
    Dim Obj_JSONData As Object
    Dim Str_APIResponse As String
    Dim Str_CaseCode As String
    Dim Obj_SubmissionList As Object
    Dim Str_JsonFilename As String:         Str_JsonFilename = "SubmissionData"
    Dim Var_Output() As Variant
    Dim Str_Database As String:             Str_Database = ThisWorkbook.Worksheets("Orgs").Range("Toggle").Value
    
    'Update URL
    If Str_Database = "Live" Then
        Str_URL = Str_LIVEURLBase
    Else
        Str_URL = Str_TESTURLBase
    End If
        
    'Update URL
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

    'API Variables
    Dim Str_TESTURLBase As String:          Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/projects/questions?submissionId=[SubmissionId]&serviceId=[ServiceID]&submissionCaseCode=[CaseCode]"
    Dim Str_LIVEURLBase As String:          Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/projects/questions?submissionId=[SubmissionId]&serviceId=[ServiceID]&submissionCaseCode=[CaseCode]"
    Dim Str_URL As String
    Dim Str_PayloadTemplate As String
    Dim Str_Payload As String
    Dim Str_Item As String
    Dim i As Long
    Dim Str_APIResponse As String
    Dim Str_Database As String:             Str_Database = ThisWorkbook.Worksheets("Orgs").Range("Toggle").Value
    
    'Update URL
    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[SubmissionId]", Str_SubID)
        Str_URL = Replace(Str_URL, "[ServiceID]", Str_Service)
        Str_URL = Replace(Str_URL, "[CaseCode]", Str_CaseCode)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[SubmissionId]", Str_SubID)
        Str_URL = Replace(Str_URL, "[ServiceID]", Str_Service)
        Str_URL = Replace(Str_URL, "[CaseCode]", Str_CaseCode)
    End If

    'Template for ONE object
    Str_PayloadTemplate = "{""questionId"":""[QuestionID]"",""questionPart"":1,""questionType"":""[Type]"",""value"":[Value]}"

    'Start JSON array
    Str_Payload = "["

    For i = LBound(Var_ResponsesArray, 1) To UBound(Var_ResponsesArray, 1)
    
        Str_Item = Str_PayloadTemplate
        
        If Var_ResponsesArray(i, 2) <> "" Then
            Str_Item = Replace(Str_Item, "[QuestionID]", Var_ResponsesArray(i, 1))
            Str_Item = Replace(Str_Item, "[Value]", Var_ResponsesArray(i, 2))
            Str_Item = Replace(Str_Item, "[Type]", Var_ResponsesArray(i, 3))
            
            Str_Payload = Str_Payload & Str_Item
        
            'Add comma except for last item
            If i < UBound(Var_ResponsesArray, 1) Then
                Str_Payload = Str_Payload & ","
            End If
        End If
    Next i

    'Close JSON array
    Str_Payload = Str_Payload & "]"
    'Run the API
    Str_APIResponse = APIPost(Str_URL, Str_Payload)
    If Str_APIResponse = "{""success"":true}" Then
        API_PostSurvey = True
    End If
    
End Function

Public Function API_CloseCaseCode(Str_SubID As String, Str_CaseCode As String) As Boolean

    'API Variables
    Dim Str_TESTURLBase As String:          Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/setCaseCodeCompleted"
    Dim Str_LIVEURLBase As String:          Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/setCaseCodeCompleted"
    Dim Str_URL As String
    Dim Str_PayloadTemplate As String
    Dim Str_Payload As String
    Dim Str_APIResponse As String
    Dim Str_Database As String:             Str_Database = ThisWorkbook.Worksheets("Orgs").Range("Toggle").Value
    
    'Template for ONE object
    Str_PayloadTemplate = "{""caseCode"": ""[CaseCode]"", ""dataSubmitted"": ""Y""}"
    Str_Payload = Replace(Str_PayloadTemplate, "[CaseCode]", Str_CaseCode)
    
    'Update URL
    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[SubmissionId]", Str_SubID)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[SubmissionId]", Str_SubID)
    End If

    'Run the API
    Str_APIResponse = APIPost(Str_URL, Str_Payload)
    If Str_APIResponse = "{""success"":true}" Then
        API_CloseCaseCode = True
    End If
End Function

Public Function GetCaseCodeNote(Str_SubmissionId As String) As Variant

    'API Variables
    
    Dim Str_TESTURLBase As String:              Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/caseCodes?allCaseCodes=true"
    Dim Str_LIVEURLBase As String:              Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/[SubmissionId]/caseCodes?allCaseCodes=true"
    Dim Str_URL As String
    Dim Obj_JSONData As Object
    Dim Str_APIResponse As String
    Dim Obj_caseNoteCodes As Object
    Dim Str_CaseCodeNote As String
    Dim Str_JsonFilename As String:         Str_JsonFilename = "CaseNoteData"
    Dim Var_Output() As Variant
    Dim Lng_CaseNoteCount As Long:          Lng_CaseNoteCount = 0
    Dim i As Long
    Dim Str_Database As String:             Str_Database = ThisWorkbook.Worksheets("Orgs").Range("Toggle").Value
    
    'Update URL
    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[SubmissionId]", Str_SubmissionId)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[SubmissionId]", Str_SubmissionId)
    End If

    Str_APIResponse = APICall(Str_URL)
    
    'OPTIONAL - OUTPUT THE DATA FOR REVIEW -
    
    'Call SaveJSONToFile(Str_APIResponse, Str_JsonFilename)

    'Convert the JSON - this will have two items "success" and "data" - we'll assume success to be true!
    Set Obj_JSONData = ParseJson(Str_APIResponse)("data")

    'Drill down - here knowledge of the JSON structure is useful, so use the SaveJSONToFile above to review
    Set Obj_caseNoteCodes = Obj_JSONData("caseNoteCodes")

    'populate output array
    For i = 1 To Obj_caseNoteCodes.Count 'cycle
        If InStr(Obj_caseNoteCodes(i)("caseCodeNotes"), "externalCode") Then
            If Obj_caseNoteCodes(i)("dataSubmitted") = "True" And Obj_caseNoteCodes(i)("completionStatus") = "Completed" Then
                Lng_CaseNoteCount = Lng_CaseNoteCount + 1
                ReDim Preserve Var_Output(1 To 2, 1 To Lng_CaseNoteCount) 'Add "column" to array (redim preserve can only add to last dimension)
                Str_CaseCodeNote = Obj_caseNoteCodes(i)("caseCodeNotes")
                Var_Output(1, Lng_CaseNoteCount) = Split(Split(Str_CaseCodeNote, """externalCode"":""")(1), """")(0)
                Var_Output(2, Lng_CaseNoteCount) = Obj_caseNoteCodes(i)("caseCode")
            End If
        End If
    Next i

    'Output array or no services message
    If Lng_CaseNoteCount = 0 Then
        Dim Var_Error(1, 1) As Variant
        Var_Error(1, 1) = "Error"
        GetCaseCodeNote = Var_Error
    Else
        GetCaseCodeNote = Application.Transpose(Var_Output)
    End If

End Function

Public Function GetCaseCodeResponses(Str_SubmissionId As String, Lng_ProjectId As Long, Str_CaseCodeString As String) As Variant

    'API Variables
    
    Dim Str_TESTURLBase As String:              Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/projects/[ProjectId]/responses?year=2026&submissionId=[SubmissionId]"
    Dim Str_LIVEURLBase As String:              Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/projects/[ProjectId]/responses?year=2026&submissionId=[SubmissionId]"
    Dim Str_URL As String
    Dim Obj_JSONData As Object
    Dim Str_APIResponse As String
    Dim Obj_Responses As Object
    Dim Str_CaseCodeNote As String
    Dim Str_JsonFilename As String:         Str_JsonFilename = "CaseNoteResponses"
    Dim Var_Output() As Variant
    Dim Lng_ResponsesCount As Long:         Lng_ResponsesCount = 0
    Dim i As Long
    Dim Str_Database As String:             Str_Database = ThisWorkbook.Worksheets("Orgs").Range("Toggle").Value
    
    'Update URL
    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[ProjectID]", Lng_ProjectId)
        Str_URL = Replace(Str_URL, "[SubmissionId]", Str_SubmissionId)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[ProjectID]", Lng_ProjectId)
        Str_URL = Replace(Str_URL, "[SubmissionId]", Str_SubmissionId)
    End If

    Str_APIResponse = APICall(Str_URL)
    
    'OPTIONAL - OUTPUT THE DATA FOR REVIEW -
    
    'Call SaveJSONToFile(Str_APIResponse, Str_JsonFilename)

    'Convert the JSON - this will have two items "success" and "data" - we'll assume success to be true!
    Set Obj_JSONData = ParseJson(Str_APIResponse)("data")

    'Drill down - here knowledge of the JSON structure is useful, so use the SaveJSONToFile above to review
    Set Obj_Responses = Obj_JSONData("responseList")
    'Info here is and item for each service, when then breaks down to:
    'serviceItemId, serviceItemName, serviceGroupId, serviceGroupName, projectId, isVisible, displaySequence, isVisibleInYear
    'serviceGroup is a new one on me and I think not relevant. I think we need serviceItemName, serviceItemID, and isVisible

    'populate output array
    For i = 1 To Obj_Responses.Count 'cycle
        If Obj_Responses(i)("caseCode") = Str_CaseCodeString Then
                Lng_ResponsesCount = Lng_ResponsesCount + 1
                ReDim Preserve Var_Output(1 To 2, 1 To Lng_ResponsesCount) 'Add "column" to array (redim preserve can only add to last dimension)
                Var_Output(1, Lng_ResponsesCount) = Obj_Responses(i)("questionId")
                Var_Output(2, Lng_ResponsesCount) = Obj_Responses(i)("itemId")
        End If
    Next i

    'Output array or no services message
    If Lng_ResponsesCount = 0 Then
        GetCaseCodeResponses = Var_Output
    Else
        GetCaseCodeResponses = Application.Transpose(Var_Output)
    End If

End Function

Function APICall(Str_URL As String) As String
    
    Dim Http As Object:                     Set Http = CreateObject("MSXML2.XMLHTTP")
    Dim Token As String:                    Token = GetToken()
    Dim Str_Response As String
    
    'Access API Data
    Http.Open "GET", (Str_URL), False
    Http.setRequestHeader "Accept", "application/json"
    Http.setRequestHeader "Token", Token
    Http.Send
    Str_Response = Http.responseText

    APICall = Str_Response

End Function

Function APIPost(Str_URL As String, Optional Str_JsonPayload As String) As String
    
    Dim Http As Object: Set Http = CreateObject("MSXML2.XMLHTTP")
    Dim Token As String: Token = GetToken()
    Dim Str_Response As String
    
    'Open POST request
    Http.Open "POST", Str_URL, False
    
    'Set headers
    Http.setRequestHeader "Accept", "application/json"
    Http.setRequestHeader "Content-Type", "application/json"  'IMPORTANT for JSON payloads
    Http.setRequestHeader "Token", Token
    
    'Send payload
    Http.Send Str_JsonPayload
    
    'Get response
    Str_Response = Http.responseText
    
    APIPost = Str_Response
    
End Function
