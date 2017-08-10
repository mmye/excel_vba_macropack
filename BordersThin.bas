Attribute VB_Name = "BordersThin"
Option Explicit

Sub DrawBordersThin()
Selection.Borders.LineStyle = True

    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub


