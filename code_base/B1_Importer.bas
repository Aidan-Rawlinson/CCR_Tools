Attribute VB_Name = "B1_Importer"
Option Explicit

Sub FileImporter()

    Dim Wsh_Home As Worksheet:              Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Wbk_Source As Workbook
    Dim Wsh_Source As Worksheet
    Dim Str_FilePath As String:             Str_FilePath = ThisWorkbook.Names("SubmissionFilePath").RefersToRange.Value
    Dim Str_SheetName As String:            Str_SheetName = ThisWorkbook.Names("DataSheetName").RefersToRange.Value
    Dim Str_Orientation As String:          Str_Orientation = ThisWorkbook.Names("Orientation").RefersToRange.Value
    Dim Str_DataStart As String:            Str_DataStart = ThisWorkbook.Names("DataStart").RefersToRange.Value
    Dim Lng_DataMax As Long:                Lng_DataMax = ThisWorkbook.Names("DataMax").RefersToRange.Value
    Dim Rng_QuestionCols As Range:          Set Rng_QuestionCols = ThisWorkbook.Names("QuestionCols").RefersToRange
    Dim Rng_StartCols As Range:             Set Rng_StartCols = ThisWorkbook.Names("StartCols").RefersToRange
    Dim Rng_DataArea As Range:              Set Rng_DataArea = ThisWorkbook.Names("DataArea").RefersToRange
    Dim Rng_FullDataArea As Range:          Set Rng_FullDataArea = ThisWorkbook.Names("FullDataArea").RefersToRange
    Dim Rng_Cell As Range
    Dim Lng_PasteRow As Long:               Lng_PasteRow = Rng_DataArea.Row
    Dim Lng_Record As Long
    Dim Lng_StartIndex As Long
    Dim Lng_Position As Long
    Dim Str_UniqueRef As String
    Dim Lng_ColIndex As Long

    '--Validate file path
    If Str_FilePath = "" Then
        MsgBox "Please enter a file path in the Submission File Path field on the Config sheet.", vbExclamation, "No File Path"
        Exit Sub
    End If

    If Dir(Str_FilePath) = "" Then
        MsgBox "The file at the specified path could not be found. Please check the Submission File Path on the Config sheet.", vbExclamation, "File Not Found"
        Exit Sub
    End If

    '--Clear previous data
    Rng_FullDataArea.ClearContents

    '--Open source file
    Set Wbk_Source = Workbooks.Open(Str_FilePath)
    DoEvents

    '--Check data sheet exists
    On Error Resume Next
    Set Wsh_Source = Wbk_Source.Worksheets(Str_SheetName)
    On Error GoTo 0

    If Wsh_Source Is Nothing Then
        MsgBox "The sheet '" & Str_SheetName & "' could not be found in the submitted file. Please check the Data Sheet Name on the Config sheet.", vbCritical, "Sheet Not Found"
        Wbk_Source.Close SaveChanges:=False
        Exit Sub
    End If

    '--Import records
    If Str_Orientation = "Columns" Then

        '--Columns orientation: one patient per column, questions in rows
        '--DataStart is a column letter; StartCols holds the source row number for each field
        Lng_StartIndex = Range(Str_DataStart & "1").Column

        For Lng_Record = Lng_StartIndex To Lng_StartIndex + Lng_DataMax - 1

            '--Read unique reference using the row number stored in StartCols(J)
            Str_UniqueRef = Wsh_Source.Cells(Rng_StartCols.Cells(1, 1).Value, Lng_Record).Value

            If Str_UniqueRef <> "" Then

                '--Write unique reference to column J on Home
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

        For Lng_Record = Lng_StartIndex To Lng_StartIndex + Lng_DataMax - 1

            '--Read unique reference using the column number stored in StartCols(J)
            Str_UniqueRef = Wsh_Source.Cells(Lng_Record, Rng_StartCols.Cells(1, 1).Value).Value

            If Str_UniqueRef <> "" Then

                '--Write unique reference to column J on Home
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

    Else
        MsgBox "Orientation value '" & Str_Orientation & "' is not recognised. Expected 'Columns' or 'Rows'.", vbCritical, "Invalid Orientation"
        Wbk_Source.Close SaveChanges:=False
        Exit Sub
    End If

    '--Close source file
    Wbk_Source.Close SaveChanges:=False

    MsgBox "Import complete.", vbInformation, "Import Complete"

End Sub
