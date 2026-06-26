Attribute VB_Name = "B5_File_Validator"
Option Explicit

' ============================================================
' B5_File_Validator
' ------------------------------------------------------------
' Validates a submitted questionnaire file before org/submission
' matching proceeds. Receives an already-open Workbook object
' from B4_Process_Folder and returns True (valid) or False
' (invalid) with a descriptive message.
'
' Three checks in sequence:
'   Check 1 - Mandatory sheets present
'             Config: MandatorySheets  e.g. "CCR^Support"
'   Check 2 - Support sheet field validation
'             XLookup checks Project ID, Submission Period,
'             Spec Type, and Organisation Name against Config
'             values and expected constants
'   Check 3 - Spot checks on CCR sheet
'             Config: SpotChecks  e.g. "CCR!E4:Patient 1^CCR!BB4:Patient 50"
'             Each token: SheetName!CellRef:ExpectedValue
'             Inner delimiters: ! between sheet and cell, : between cell and value
'
' A failed check surfaces a MsgBox to the user and returns False.
' All checks use the ^ delimiter to split lists.
' ============================================================

Public Function ValidateFile(ByVal Wbk_Source As Workbook, _
                             ByVal Str_FileName As String) As Boolean

    ValidateFile = False     '--Assume invalid until all checks pass

    ' ----------------------------------------------------------
    ' CHECK 1: Mandatory sheets present
    ' ----------------------------------------------------------
    Dim Str_MandatorySheets As String
    Str_MandatorySheets = Trim(CStr(ThisWorkbook.Names("MandatorySheets").RefersToRange.Value))

    If Str_MandatorySheets = "" Then
        MsgBox "Configuration error: MandatorySheets is not configured." & vbCrLf & vbCrLf & _
               "File: " & Str_FileName & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Configuration Error"
        Exit Function
    End If

    Dim Arr_Sheets() As String
    Arr_Sheets = Split(Str_MandatorySheets, "^")

    Dim i As Integer
    For i = 0 To UBound(Arr_Sheets)
        Dim Str_SheetName As String
        Str_SheetName = Trim(Arr_Sheets(i))
        If Str_SheetName = "" Then GoTo NextSheet

        Dim Wsh_Test As Worksheet
        Set Wsh_Test = Nothing
        On Error Resume Next
        Set Wsh_Test = Wbk_Source.Worksheets(Str_SheetName)
        On Error GoTo 0

        If Wsh_Test Is Nothing Then
            MsgBox "Check 1 failed: mandatory sheet not found." & vbCrLf & vbCrLf & _
                   "File:          " & Str_FileName & vbCrLf & _
                   "Missing sheet: " & Str_SheetName & vbCrLf & vbCrLf & _
                   "File will be skipped.", _
                   vbExclamation, "Invalid File"
            Exit Function
        End If

NextSheet:
    Next i

    ' ----------------------------------------------------------
    ' CHECK 2: Support sheet field validation
    ' ----------------------------------------------------------
    Dim Wsh_Support As Worksheet
    Set Wsh_Support = Wbk_Source.Worksheets("Support")

    Dim Str_ProjID As String
    Dim Str_SubYear As String
    Dim Str_ProjIDExpected As String
    Dim Str_SubYearExpected As String

    Str_ProjIDExpected = Trim(CStr(ThisWorkbook.Names("ProjectID").RefersToRange.Value))
    Str_SubYearExpected = Trim(CStr(ThisWorkbook.Names("SubmissionYear").RefersToRange.Value))

    '--XLookup: Project ID
    On Error Resume Next
    Str_ProjID = Trim(CStr(WorksheetFunction.XLookup( _
                    "Project ID", _
                    Wsh_Support.Columns(1), _
                    Wsh_Support.Columns(2), _
                    "NOT FOUND")))
    On Error GoTo 0

    If Str_ProjID = "NOT FOUND" Then
        MsgBox "Check 2 failed: 'Project ID' field not found in Support sheet." & vbCrLf & vbCrLf & _
               "File: " & Str_FileName & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Invalid File"
        Exit Function
    End If

    If Str_ProjID <> Str_ProjIDExpected Then
        MsgBox "Check 2 failed: Project ID in file does not match this tool." & vbCrLf & vbCrLf & _
               "File:     " & Str_FileName & vbCrLf & _
               "Expected: " & Str_ProjIDExpected & vbCrLf & _
               "Found:    " & Str_ProjID & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Invalid File"
        Exit Function
    End If

    '--XLookup: Submission Period
    Dim Str_SubPeriod As String
    On Error Resume Next
    Str_SubPeriod = Trim(CStr(WorksheetFunction.XLookup( _
                    "Submission period", _
                    Wsh_Support.Columns(1), _
                    Wsh_Support.Columns(2), _
                    "NOT FOUND")))
    On Error GoTo 0

    If Str_SubPeriod = "NOT FOUND" Then
        MsgBox "Check 2 failed: 'Submission period' field not found in Support sheet." & vbCrLf & vbCrLf & _
               "File: " & Str_FileName & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Invalid File"
        Exit Function
    End If

    If Str_SubPeriod <> Str_SubYearExpected Then
        MsgBox "Check 2 failed: Submission period in file does not match this tool." & vbCrLf & vbCrLf & _
               "File:     " & Str_FileName & vbCrLf & _
               "Expected: " & Str_SubYearExpected & vbCrLf & _
               "Found:    " & Str_SubPeriod & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Invalid File"
        Exit Function
    End If

    '--XLookup: Spec Type
    Dim Str_SpecType As String
    On Error Resume Next
    Str_SpecType = Trim(CStr(WorksheetFunction.XLookup( _
                    "Spec Type", _
                    Wsh_Support.Columns(1), _
                    Wsh_Support.Columns(2), _
                    "NOT FOUND")))
    On Error GoTo 0

    If Str_SpecType = "NOT FOUND" Then
        MsgBox "Check 2 failed: 'Spec Type' field not found in Support sheet." & vbCrLf & vbCrLf & _
               "File: " & Str_FileName & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Invalid File"
        Exit Function
    End If

    If Str_SpecType <> "Clinical Case Review" Then
        MsgBox "Check 2 failed: Spec Type in file is not 'Clinical Case Review'." & vbCrLf & vbCrLf & _
               "File:  " & Str_FileName & vbCrLf & _
               "Found: " & Str_SpecType & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Invalid File"
        Exit Function
    End If

    '--XLookup: Organisation Name (must be non-blank)
    Dim Str_OrgName As String
    On Error Resume Next
    Str_OrgName = Trim(CStr(WorksheetFunction.XLookup( _
                    "Organisation Name", _
                    Wsh_Support.Columns(1), _
                    Wsh_Support.Columns(2), _
                    "NOT FOUND")))
    On Error GoTo 0

    If Str_OrgName = "NOT FOUND" Then
        MsgBox "Check 2 failed: 'Organisation Name' field not found in Support sheet." & vbCrLf & vbCrLf & _
               "File: " & Str_FileName & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Invalid File"
        Exit Function
    End If

    If Str_OrgName = "" Or Str_OrgName = "0" Then
        MsgBox "Check 2 failed: Organisation Name has not been completed in the submitted file." & vbCrLf & vbCrLf & _
               "File: " & Str_FileName & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Invalid File"
        Exit Function
    End If

    ' ----------------------------------------------------------
    ' CHECK 3: Spot checks
    ' ----------------------------------------------------------
    Dim Str_SpotChecks As String
    Str_SpotChecks = Trim(CStr(ThisWorkbook.Names("SpotChecks").RefersToRange.Value))

    If Str_SpotChecks = "" Then
        MsgBox "Configuration error: SpotChecks is not configured." & vbCrLf & vbCrLf & _
               "File: " & Str_FileName & vbCrLf & vbCrLf & _
               "File will be skipped.", _
               vbExclamation, "Configuration Error"
        Exit Function
    End If

    Dim Arr_Checks() As String
    Arr_Checks = Split(Str_SpotChecks, "^")

    Dim j As Integer
    For j = 0 To UBound(Arr_Checks)

        Dim Str_Token As String
        Str_Token = Trim(Arr_Checks(j))
        If Str_Token = "" Then GoTo NextCheck

        '--Parse token: SheetName!CellRef:ExpectedValue
        Dim Int_Bang As Integer:    Int_Bang = InStr(Str_Token, "!")
        Dim Int_Colon As Integer:   Int_Colon = InStr(Str_Token, ":")

        If Int_Bang = 0 Or Int_Colon = 0 Or Int_Colon < Int_Bang Then
            MsgBox "Configuration error: SpotChecks token is malformed." & vbCrLf & vbCrLf & _
                   "Token: " & Str_Token & vbCrLf & vbCrLf & _
                   "Expected format: SheetName!CellRef:ExpectedValue" & vbCrLf & vbCrLf & _
                   "File will be skipped.", _
                   vbExclamation, "Configuration Error"
            Exit Function
        End If

        Dim Str_SC_Sheet As String:     Str_SC_Sheet = Trim(Left(Str_Token, Int_Bang - 1))
        Dim Str_SC_Cell As String:      Str_SC_Cell = Trim(Mid(Str_Token, Int_Bang + 1, Int_Colon - Int_Bang - 1))
        Dim Str_SC_Expected As String:  Str_SC_Expected = Trim(Mid(Str_Token, Int_Colon + 1))

        '--Get the sheet
        Dim Wsh_SC As Worksheet
        Set Wsh_SC = Nothing
        On Error Resume Next
        Set Wsh_SC = Wbk_Source.Worksheets(Str_SC_Sheet)
        On Error GoTo 0

        If Wsh_SC Is Nothing Then
            MsgBox "Check 3 failed: spot check sheet not found." & vbCrLf & vbCrLf & _
                   "File:  " & Str_FileName & vbCrLf & _
                   "Sheet: " & Str_SC_Sheet & vbCrLf & vbCrLf & _
                   "File will be skipped.", _
                   vbExclamation, "Invalid File"
            Exit Function
        End If

        '--Get the cell value
        Dim Str_SC_Actual As String
        On Error Resume Next
        Str_SC_Actual = Trim(CStr(Wsh_SC.Range(Str_SC_Cell).Value))
        On Error GoTo 0

        If Str_SC_Actual <> Str_SC_Expected Then
            MsgBox "Check 3 failed: spot check value does not match." & vbCrLf & vbCrLf & _
                   "File:     " & Str_FileName & vbCrLf & _
                   "Cell:     " & Str_SC_Sheet & "!" & Str_SC_Cell & vbCrLf & _
                   "Expected: " & Str_SC_Expected & vbCrLf & _
                   "Found:    " & Str_SC_Actual & vbCrLf & vbCrLf & _
                   "The template structure may have been modified. File will be skipped.", _
                   vbExclamation, "Invalid File"
            Exit Function
        End If

NextCheck:
    Next j

    ' ----------------------------------------------------------
    ' All checks passed
    ' ----------------------------------------------------------
    ValidateFile = True

End Function
