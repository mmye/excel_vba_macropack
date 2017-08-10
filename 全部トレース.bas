Attribute VB_Name = "�S���g���[�X"
Option Explicit

Sub �Q�ƌ��̃g���[�X_�I��͈͓�()
    Dim i As Long
    Dim buf As String
    Call ScreenUpdatingSwitch
    If TypeName(Selection) = "Range" Then
        For i = 1 To Selection.Count
            If Not IsError(Selection(i)) Then
                buf = Selection(i)
                If Len(buf) > 0 Then
                    If IsNumeric(buf) Then Selection(i).ShowPrecedents
                End If
            End If
        Next i
    End If
    Call ScreenUpdatingSwitch
End Sub

Sub �Q�ƌ��̃g���[�X_�I�𒆂̃V�[�g()
    Dim buf As String
    Dim rng As Range
    Dim r As Range
    Dim r1 As Range, r2 As Range
    
    Call ScreenUpdatingSwitch
    Set r1 = ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas)
    Set r2 = ActiveSheet.Cells.SpecialCells(xlCellTypeConstants)
    
    If Not r1 Is Nothing Then
        Set rng = Union(r1, r2)
    Else
        Set rng = r2
    End If
    
    For Each r In rng
        DoEvents
        If Not IsError(r.Value) Then
            buf = r.Value
            If Len(r.Value) > 0 Then
                If IsNumeric(r.Value) Then
                    Debug.Print r.Value
                    r.ShowPrecedents
                End If
            End If
        End If
    Next r
    
    Call ScreenUpdatingSwitch
    Set rng = Nothing
    Set r1 = Nothing
    Set r2 = Nothing
End Sub
Sub �Q�Ɛ�̃g���[�X_�I�𒆂̃V�[�g()
    Dim buf As String
    Dim rng As Range
    Dim r As Range
    Dim r1 As Range, r2 As Range
    
    Call ScreenUpdatingSwitch
    Set r1 = ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas)
    Set r2 = ActiveSheet.Cells.SpecialCells(xlCellTypeConstants)
    
    If Not r1 Is Nothing Then
        Set rng = Union(r1, r2)
    Else
        Set rng = r2
    End If
    
    For Each r In rng
        DoEvents
        If Not IsError(r.Value) Then
            buf = r.Value
            If Len(r.Value) > 0 Then
                If IsNumeric(r.Value) Then
                    r.ShowDependents
                End If
            End If
        End If
    Next r
    
    Call ScreenUpdatingSwitch
    Set rng = Nothing
    Set r1 = Nothing
    Set r2 = Nothing
End Sub
Sub �g���[�X���̏���()
Attribute �g���[�X���̏���.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.ClearArrows
End Sub

Private Sub ScreenUpdatingSwitch()
    Application.ScreenUpdating = Not Application.ScreenUpdating
'    MsgBox "ScreenUpdating:" & Application.ScreenUpdating
End Sub


