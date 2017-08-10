Attribute VB_Name = "CopyRangesMultiBooks"
Option Explicit

Sub CopyRanges()

Dim n As Workbook
Set n = Workbooks.Add
Dim tgt As Worksheet
Set tgt = n.Sheets(1)

Dim wks, wkb
For Each wkb In Workbooks
    If wkb.Name <> n.Name Then
        Dim src As Worksheet
        Set src = wkb.Sheets(1)

        Dim r As Range
        Set r = src.UsedRange
        Dim v As Variant
        v = r

        Dim h As Long
        h = UBound(v) - LBound(v)

        'target sheet first cell
        Dim f As Range
        Set f = tgt.UsedRange(tgt.UsedRange.Count)

        'Copy cells
        tgt.Range(Cells(f.Row, 1), Cells(f.Row + h, 2)) = v
    End If
Next wkb

Set wks = Nothing
Set wkb = Nothing

End Sub

