Attribute VB_Name = "B3_SUBMISSIONS"
Option Explicit
Option Compare Text
Option Base 1

'###########################################
'# USES TOKEN FROM A1_API_SUPPORT          #
'# USES JSON CONVERTER FROM A1_API_SUPPORT #
'# USES APICall FROM A2_API_FUNCTIONS      #
'###########################################

Public Sub PopulateSubmissions()

    Dim Str_TESTURLBase As String:  Str_TESTURLBase = "https://membersapidev.nhsbenchmarking.nhs.uk/submissions/list?projectId=[ProjectID]&year=[Year]"
    Dim Str_LIVEURLBase As String:  Str_LIVEURLBase = "https://membersapi.nhsbenchmarking.nhs.uk/submissions/list?projectId=[ProjectID]&year=[Year]"
    Dim Str_URL As String
    Dim Str_APIResponse As String
    Dim Obj_JSONData As Object
    Dim Obj_SubmissionList As Object
    Dim Rng_Submissions As Range
    Dim Rng_DataStart As Range
    Dim Var_Output() As Variant
    Dim Lng_Count As Long:          Lng_Count = 0
    Dim i As Long

    Dim Lng_ProjectId As Long:      Lng_ProjectId = CLng(ThisWorkbook.Names("ProjectID").RefersToRange.Value)
    Dim Str_Year As String:         Str_Year = CStr(ThisWorkbook.Names("SubmissionYear").RefersToRange.Value)
    Dim Str_Database As String:     Str_Database = ThisWorkbook.Names("Toggle").RefersToRange.Value

    ' Build URL
    If Str_Database = "Live" Then
        Str_URL = Replace(Str_LIVEURLBase, "[ProjectID]", Lng_ProjectId)
    Else
        Str_URL = Replace(Str_TESTURLBase, "[ProjectID]", Lng_ProjectId)
    End If
    Str_URL = Replace(Str_URL, "[Year]", Str_Year)

    ' Call API
    Str_APIResponse = APICall(Str_URL)

    ' Parse response
    Set Obj_JSONData = ParseJson(Str_APIResponse)("data")
    Set Obj_SubmissionList = Obj_JSONData("submissionList")(Str_Year)

    ' Clear existing data below header
    Set Rng_Submissions = ThisWorkbook.Names("Submissions").RefersToRange
    Set Rng_DataStart = Rng_Submissions.Offset(Rng_Submissions.Rows.Count, 0).Resize(1, Rng_Submissions.Columns.Count)
    Rng_DataStart.Resize(10000, Rng_Submissions.Columns.Count).ClearContents

    ' Build output array — one row per submission, four columns: Org ID, Org Name, Submission Name, Submission ID
    Lng_Count = Obj_SubmissionList.Count
    If Lng_Count = 0 Then
        MsgBox "No submissions found for Project " & Lng_ProjectId & " (" & Str_Year & ").", vbExclamation
        Exit Sub
    End If

    ReDim Var_Output(1 To Lng_Count, 1 To 4)

    For i = 1 To Lng_Count
        Var_Output(i, 1) = Obj_SubmissionList(i)("organisationId")
        Var_Output(i, 2) = Obj_SubmissionList(i)("organisationName")
        Var_Output(i, 3) = Obj_SubmissionList(i)("submissionName")
        Var_Output(i, 4) = Obj_SubmissionList(i)("submissionId")
    Next i

    ' Write to sheet in one operation
    Rng_DataStart.Resize(Lng_Count, 4).Value = Var_Output

    MsgBox "Submissions loaded: " & Lng_Count & " records.", vbInformation

End Sub
