Attribute VB_Name = "MarkApplication"
Option Explicit

Sub MarkApplication()
Dim i
Dim EndRow As Long

Dim AppCol As Long
AppCol = Range("application").column

Dim r
Set r = ActiveSheet.UsedRange
EndRow = r.Rows.Count

Dim SrcCol As Long
SrcCol = Range("source").column

For i = 2 To EndRow
    Dim src
    src = Cells(i, SrcCol).Value
    If InStr(src, "dairy") > 0 Then Cells(i, AppCol).Value = "Daily"
    If InStr(src, "pharma") > 0 Then Cells(i, AppCol).Value = "Pharm"
    If InStr(src, "printing") > 0 Then Cells(i, AppCol).Value = "Plastic"
    If InStr(src, "cosmetics") > 0 Then Cells(i, AppCol).Value = "Cosmetics"
Next i

Set r = Nothing

End Sub

