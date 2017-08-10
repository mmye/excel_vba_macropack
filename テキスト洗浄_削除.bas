Attribute VB_Name = "テキスト洗浄_削除"
Option Explicit

Sub 一字の文節を削除()
    Dim i As Long
    Dim lEndRow As Long
    Dim CurrCol As Long
    Application.ScreenUpdating = False
    If Selection.Count = 1 Then
        CurrCol = Selection.column
    Else
        MsgBox "１つのセルのみを選択してから再実行してください"
        Exit Sub
    End If
    lEndRow = Cells(Rows.Count, CurrCol).End(xlUp).Row
    For i = 1 To lEndRow
        On Error Resume Next
        If Len(Cells(i, CurrCol).Value) < 3 Then
            Cells(i, CurrCol).Value = Empty
        End If
        On Error GoTo 0
    Next i
    Application.ScreenUpdating = True
End Sub
