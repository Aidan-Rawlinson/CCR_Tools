Attribute VB_Name = "B2_Toggle"
Option Explicit

Sub ToggleButton()
    'Controls toggle button for Multi Email Templates functionality
    
    Dim Wsh_Home As Worksheet:                  Set Wsh_Home = ThisWorkbook.Sheets("Home")
    Dim Wrk_Hidden_Process As Worksheet:        Set Wrk_Hidden_Process = ThisWorkbook.Sheets("Orgs")
    Dim Shp_Shape As Shape:                     Set Shp_Shape = ActiveSheet.Shapes(Application.Caller)

    If Shp_Shape.Name = "DatabaseToggle" Then
        If Shp_Shape.TextFrame.Characters.Text = "TEST" Then
            ' Change to Live
            Shp_Shape.TextFrame.Characters.Text = "LIVE"
            Shp_Shape.Fill.ForeColor.RGB = RGB(240, 187, 41) ' Yellow
            Wrk_Hidden_Process.Range("Toggle").Value = "Live"
        Else
            ' Change to Test
            Shp_Shape.TextFrame.Characters.Text = "TEST"
            Shp_Shape.Fill.ForeColor.RGB = RGB(127, 154, 228) ' Purple
            Range("Toggle").Value = "Test"
        End If
    End If
    
End Sub

