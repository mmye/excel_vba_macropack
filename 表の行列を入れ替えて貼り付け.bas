Attribute VB_Name = "�\�̍s������ւ��ē\��t��"
Option Explicit

Sub Pastespecial()
Attribute Pastespecial.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim TblCnt As Long
    Dim TblName As String
    
    TblCnt = ActiveSheet.ListObjects.Count
    TblName = "�e�[�u��" & CStr(TblCnt + 1)

'   �\�̍s������ւ��ē\��t��
    Selection.Pastespecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Application.CutCopyMode = False

'  �����Z�����������ăf�[�^��⊮����
    �Z�����������ƕ�����⊮
    
'  �\���e�[�u���ɕϊ�����
    ActiveSheet.ListObjects.Add _
            (xlSrcRange, Range(ActiveCell, Selection.Item(Selection.Count)), , xlYes).Name _
        = TblName
'   �͂��߂̃Z����I��
    Selection.Item(1).Select
End Sub
Private Sub �Z�����������ƕ�����⊮()
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

