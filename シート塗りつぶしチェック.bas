Attribute VB_Name = "�V�[�g�h��Ԃ��`�F�b�N"
Option Explicit


'---------------------------------------------------------------------------------------
' Method : CheckByInteriorColor
' Author : m.maeyama
' Date   : 2017/03/28
' Purpose: �V�[�gA�AB������Ƃ��A�X�V�ŃV�[�gA'�AB'�𓾂āA�X�V���ꂽ�Z���������h��Ԃ�
'---------------------------------------------------------------------------------------
Sub CheckByInteriorColor()
    Dim rFrom As Range
    Dim rTo As Range
    
    Dim wksFrom As Worksheet
    Dim wksTo As Worksheet
    
    Set wksFrom = Sheets(1)
    Set wksTo = Sheets(2)
    
    Dim i As Long
    
    Set rFrom = wksFrom.UsedRange
    
    Dim r As Range
    For Each r In rFrom
        '��r���̃Z�����h��Ԃ���Ă��邩������ׂ�
        If r.Interior.color = 65535 Then
            Dim Row, col
            Row = r.Row
            col = r.column
            
            Dim color As Long
            color = r.Interior.color
            
            '��r��V�[�g�𓯂��F�œh��Ԃ�
            wksTo.Cells(Row, col).Interior.color = color
        End If
    Next r
    
    Set wksFrom = Nothing
    Set wksTo = Nothing
    Set rFrom = Nothing
    Set rTo = Nothing
End Sub
