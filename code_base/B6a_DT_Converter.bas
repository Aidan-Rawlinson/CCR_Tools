Attribute VB_Name = "B6a_DT_Converter"
Option Explicit

' ============================================================
' B6a_DT_Converter
' ------------------------------------------------------------
' Converts DT question cells on the Home sheet from their raw
' imported values (text DD/MM/YYYY or Excel date serials) into
' the API-ready string format YYYY-MM-DD 00:00:00.000.
'
' Called by B4_Process_Folder after all files have been
' imported and before B6_Response_Validator runs, passing the
' first and last row numbers written during the run.
'
' For each DT column (identified via TypeCols named range):
'   1. TextToColumns with xlDMYFormat is run on the full
'      column range for the imported rows. This forces Excel
'      to parse text dates as DD/MM/YYYY and convert them to
'      date serials regardless of regional settings. Values
'      already stored as serials are unaffected.
'   2. Each cell is formatted as Text (@) before the string
'      is written back, preventing Excel from re-interpreting
'      the YYYY-MM-DD 00:00:00.000 string as a date serial.
'   3. If not numeric after TextToColumns, the raw value is
'      left as-is and B6 will flag it orange.
' ============================================================

Sub ConvertDTColumns(Lng_FirstRow As Long, Lng_LastRow As Long)

    Dim Wsh_Home As Worksheet:          Set Wsh_Home = ThisWorkbook.Worksheets("Home")
    Dim Rng_QuestionCols As Range:      Set Rng_QuestionCols = ThisWorkbook.Names("QuestionCols").RefersToRange
    Dim Rng_TypeCols As Range:          Set Rng_TypeCols = ThisWorkbook.Names("TypeCols").RefersToRange

    Dim Rng_QCell As Range
    Dim Lng_Col As Long
    Dim Rng_DTRange As Range
    Dim Lng_Row As Long
    Dim Var_Val As Variant

    For Each Rng_QCell In Rng_QuestionCols

        '--Check whether this question column is DT type
        If Rng_TypeCols.Cells(1, Rng_QCell.Column - Rng_QuestionCols.Column + 1).Value = "DT" Then

            Lng_Col = Rng_QCell.Column

            '--Build the range covering the imported rows for this column
            Set Rng_DTRange = Wsh_Home.Range(Wsh_Home.Cells(Lng_FirstRow, Lng_Col), _
                                             Wsh_Home.Cells(Lng_LastRow, Lng_Col))

            '--Run TextToColumns to parse text DD/MM/YYYY values as date serials
            Rng_DTRange.TextToColumns _
                Destination:=Rng_DTRange.Cells(1), _
                DataType:=xlDelimited, _
                ConsecutiveDelimiter:=False, _
                Tab:=False, _
                Semicolon:=False, _
                Comma:=False, _
                Space:=False, _
                Other:=False, _
                FieldInfo:=Array(1, xlDMYFormat)

            '--Convert any resulting numeric serials to YYYY-MM-DD 00:00:00.000 strings
            For Lng_Row = Lng_FirstRow To Lng_LastRow
                Var_Val = Wsh_Home.Cells(Lng_Row, Lng_Col).Value
                If Var_Val <> "" Then
                    If IsNumeric(Var_Val) Then
                        On Error Resume Next
                        Dim Dt_Date As Date
                        Dt_Date = CDate(Var_Val)
                        If Err.Number = 0 Then
                            '--Format cell as Text first to prevent Excel re-interpreting the string as a date
                            Wsh_Home.Cells(Lng_Row, Lng_Col).NumberFormat = "@"
                            Wsh_Home.Cells(Lng_Row, Lng_Col).Value = Format(Dt_Date, "YYYY-MM-DD") & " 00:00:00.000"
                        End If
                        On Error GoTo 0
                    End If
                End If
            Next Lng_Row

        End If

    Next Rng_QCell

End Sub
