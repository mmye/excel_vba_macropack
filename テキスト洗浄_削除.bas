Attribute VB_Name = "�e�L�X�g���_�폜"
Option Explicit

Sub �ꎚ�̕��߂��폜()
    Dim i As Long
    Dim lEndRow As Long
    Dim CurrCol As Long
    Application.ScreenUpdating = False
    If Selection.Count = 1 Then
        CurrCol = Selection.column
    Else
        MsgBox "�P�̃Z���݂̂�I�����Ă���Ď��s���Ă�������"
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
