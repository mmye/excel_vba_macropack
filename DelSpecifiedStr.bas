Attribute VB_Name = "DelSpecifiedStr"
Option Explicit

Sub DelSpecifiedStr()
'   ### 開いたブックをtxtに変換するスクリプト
'   ### 保存前に全ブックで１列目を削除する処理が入っている

    Dim wkb As Workbook

    For Each wkb In Workbooks
        Dim wks As Worksheet
        Set wks = wkb.Sheets(1)
        Dim r As Range
        
        Dim k
        For k = 1 To wks.UsedRange.Rows.Count
            Set r = wks.Cells(k, 1)
            
            Dim buf
            buf = r.Value
            On Error Resume Next
            If Left$(buf, 1) = """" Then
                r.Value = Mid$(buf, 2, Len(buf) - 1)
            End If
            
            If Right$(buf, 1) = """" Then
                r.Value = Left$(buf, Len(buf) - 1)
            End If
        Next k
    Next wkb

End Sub

Sub test()
Dim str
str = ActiveCell
Debug.Print str

If Left$(ActiveCell.Value, 1) = """" Then
    ActiveCell.Value = Mid$(ActiveCell, 2, Len(ActiveCell) - 1)
End If

If Right$(ActiveCell.Value, 1) = """" Then
    ActiveCell.Value = Left$(ActiveCell, Len(ActiveCell) - 1)
End If

End Sub

