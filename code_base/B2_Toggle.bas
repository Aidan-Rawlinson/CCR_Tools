Attribute VB_Name = "B2_Toggle"
Option Explicit

Sub ToggleButton()

    Dim Shp_Shape As Shape: Set Shp_Shape = ActiveSheet.Shapes(Application.Caller)

    If Shp_Shape.Name = "DatabaseToggle" Then
        If Shp_Shape.TextFrame.Characters.Text = "TEST" Then
            Shp_Shape.TextFrame.Characters.Text = "LIVE"
            Shp_Shape.Fill.ForeColor.RGB = RGB(240, 187, 41) ' Yellow
            ThisWorkbook.Names("Toggle").RefersToRange.Value = "Live"
        Else
            Shp_Shape.TextFrame.Characters.Text = "TEST"
            Shp_Shape.Fill.ForeColor.RGB = RGB(127, 154, 228) ' Blue
            ThisWorkbook.Names("Toggle").RefersToRange.Value = "Test"
        End If
    End If

End Sub
