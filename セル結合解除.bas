Attribute VB_Name = "�Z����������"
Option Explicit
'�I��͈͓��̃Z���������������A�������ꂽ�Z���ɓ����e�L�X�g��}������
'�쐬���F20161107
'Power BI�ɓ���邽�߂Ƀf�[�^�𐮌`����̂Ɏg���B

Sub UnmergeSelectionandFillStr()
Attribute UnmergeSelectionandFillStr.VB_ProcData.VB_Invoke_Func = "U\n14"
    Dim r As Range, rMergeArea As Range, r2 As Range
    Dim str As String

    ScreenUpdatingSwitch
    If Selection.Count = 0 Then Exit Sub
    
    For Each r In Selection
        If Not IsError(r.Value) Then
            If r.Value <> "" Then
                If Not IsError(r.Value) Then
                    If r.MergeCells Then
                        Set rMergeArea = r.MergeArea
                        str = r.Value
                        r.UnMerge
                        For Each r2 In rMergeArea
                            r2.Value = str
                        Next r2
                    End If
                End If
            End If
        End If
    Next r

    Set r = Nothing
    Set r2 = Nothing
    Set rMergeArea = Nothing
    ScreenUpdatingSwitch
End Sub

Private Sub ScreenUpdatingSwitch()
    Application.ScreenUpdating = Not Application.ScreenUpdating
End Sub
