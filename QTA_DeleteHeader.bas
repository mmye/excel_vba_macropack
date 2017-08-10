Attribute VB_Name = "QTA_DeleteHeader"
Option Explicit
Dim mbExitProc As Boolean
Dim mbCancelEvent As Boolean

'---------------------------------------------------------------------------------------
' Method : GetLastRow
' Author : mokoo
' Date   : 2016/02/20
' Purpose: �V�[�g�̍ŉ��s���擾����
'---------------------------------------------------------------------------------------
Private Function GetLastRow(ByVal st As Worksheet) As Long
    Dim rngPrintArea As Range
    
    With st
        On Error GoTo ErrHandler
        Set rngPrintArea = .Range(.PageSetup.PrintArea)
        On Error GoTo 0
        GetLastRow = rngPrintArea.Item(rngPrintArea.Count).Row
    End With
 Exit Function
 
ErrHandler:
 
 mbCancelEvent = True
 
 End Function
 
 Private Function GetLastCol(ByVal st As Worksheet) As Long
    Dim rngPrintArea As Range
    
    With st
        Set rngPrintArea = .Range(.PageSetup.PrintArea)
        GetLastCol = rngPrintArea.Item(rngPrintArea.Count).column
    End With
 
 End Function
'---------------------------------------------------------------------------------------
' Method : CountTotalPages
' Author : temporary3
' Date   : 2016/04/07
' Purpose: �A�N�e�B�u�V�[�g�̌��Ϗ��Ƃ��Ẵy�[�W�����J�E���g����i60�s/�y�[�W�Ōv�Z�j
'---------------------------------------------------------------------------------------
Private Function GetPageCount() As Long
    Dim lPageCount As Long
    Dim rowCount As Long
    Dim lPageMargin As Long
    
    With ActiveWorkbook.ActiveSheet
        
        rowCount = .UsedRange.Rows.Count

        If rowCount < 60 Then
            MsgBox "1�y�[�W��������܂���"
            End    '�s����1�y�[�W�����Ȃ�I��
        End If
        
        '�K��s�����傤�ǂ��܂ރy�[�W��
        lPageCount = (rowCount / 60)
        '�]��y�[�W
        If (rowCount Mod 60) <> 0 Then lPageMargin = 1
        
        '���v�y�[�W��
        GetPageCount = lPageCount + lPageMargin
    End With
    
End Function
'---------------------------------------------------------------------------------------
' Method : DeleteHeader
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �w�b�_�[���폜����
'---------------------------------------------------------------------------------------
Sub DeleteHeader()
'*TODO:�w�b�_�[�̍폜�c�����N���邱�Ƃ�����

    Dim i As Long    '�J�E���^1
    Dim f As Long    '�J�E���^2
    Dim Pages As Long    '�y�[�W��
    Dim NumPage As Long
    Dim lLastRow As Long
    
    'Application.ScreenUpdating = False
    
    lLastRow = GetLastRow(ActiveWorkbook.ActiveSheet)    '�V�[�g�̃f�[�^������ŏI�s���擾
    NumPage = GetPageCount

    Call EraseBorders(NumPage)
    
    If mbExitProc Then
        mbExitProc = False
        MsgBox "�L�����Z�����܂���", vbOKOnly + vbInformation, "�L�����Z��"
        Exit Sub
    Else
        Call DeleteLogo(NumPage)
    End If
    
    'MsgBox lCountDeletedHeader & "�y�[�W���̃w�b�_�[���폜���܂����B", vbOKOnly Or vbInformation, "�w�b�_�[�폜����"
    
Exit Sub

NoHeadertoDelete:
    MsgBox "�폜����w�b�_�[������܂���B���̌��Ϗ���1�y�[�W�܂ł����Ȃ��悤�ł��B", vbOKOnly Or vbInformation, "�w�b�_�[�Ȃ�"

End Sub

Private Sub EraseBorders(ByVal NumPage As Long)

    Dim i As Long
    Dim j As Long
    Dim b1 As Border    '��1�r��
    Dim b2 As Border    '��2�r��
    Dim CurrentActivecell   As Range
    Dim lRowTop As Long
    Dim lRowBottm As Long
    Dim rngToDelete As Range
    Dim lCountDeletedHeader As Long
    Dim myYesNo As VbMsgBoxResult
    Dim lHeaderCount As Long
    Dim sUserMessage As String
    
    '�r����Bold�̏ꍇ�̏���
    For i = 60 To NumPage * 60    '�ŏI�Z���̔ԍ����疖�������߂�

        '�r���̎Q��
        With Cells(i, "C")
            Set b1 = .Borders(xlEdgeBottom)
            Set b2 = .offset(1, 0).Borders(xlEdgeBottom)
        End With
        
        If b1.LineStyle = xlContinuous And _
           b1.Weight = xlThick And _
           b2.LineStyle = xlContinuous Then

            '�w�b�_�[�̏�[�s���n�[�h�R�[�f�B���O�B
            lRowTop = i - 4
            
            '�w�b�_�[�̉��[�s���������Bj��Medium�����r���Ƃ̉���������
            For j = 1 To 10
                If Cells(i, "C").offset(j, 0).Borders(xlEdgeBottom).LineStyle = xlContinuous Then lRowBottm = i + j
            Next j
            
            If rngToDelete Is Nothing Then
                Set rngToDelete = Rows(lRowTop & ":" & lRowBottm)
                lHeaderCount = lHeaderCount + 1
            Else
                Set rngToDelete = Union(rngToDelete, Rows(lRowTop & ":" & lRowBottm))
                lHeaderCount = lHeaderCount + 1
            End If
            
            lCountDeletedHeader = lCountDeletedHeader + 1
        End If
    Next i

    '�r����Medium�̏ꍇ�̏���
    For i = 60 To (NumPage * 60)    '�ŏI�Z���̔ԍ����疖�������߂�

        With Cells(i, "C")
            Set b1 = .Borders(xlEdgeBottom)
            Set b2 = .offset(1, 0).Borders(xlEdgeBottom)
        End With

        '��Ж��̉���Medium�����r����������
        If b1.LineStyle = xlContinuous And _
           b1.Weight = xlMedium And _
           b2.LineStyle = xlContinuous Then

            '�w�b�_�[�̏�[�s���n�[�h�R�[�f�B���O�B
            lRowTop = i - 4
            
            '�w�b�_�[�̉��[�s���������Bj��Medium�����r���Ƃ̉���������
            For j = 1 To 10
                If Cells(i, "C").offset(j, 0).Borders(xlEdgeBottom).LineStyle = xlContinuous Then lRowBottm = i + j
            Next j
            
            If rngToDelete Is Nothing Then
                Set rngToDelete = Rows(lRowTop & ":" & lRowBottm)
                lHeaderCount = lHeaderCount + 1
            Else
                Set rngToDelete = Union(rngToDelete, Rows(lRowTop & ":" & lRowBottm))
                lHeaderCount = lHeaderCount + 1
            End If
            
        End If
    Next i
    
    If TypeName(Selection) = "Range" Then Set CurrentActivecell = ActiveCell
    
    If rngToDelete Is Nothing Then
        MsgBox "�w�b�_�[���݂���܂���ł���", vbOKOnly + vbInformation, "�w�b�_�[�Ȃ�"
        Exit Sub
    Else
    
        '�擾�����w�b�_�[�͈̔͂��폜���Ă悢�����[�U�[�Ɋm�F����
        rngToDelete.Select
        sUserMessage = "�w�b�_�[��" & lHeaderCount & "�݂���܂����B���ݑI������Ă���s���폜���Ă���낵���ł����H"
        myYesNo = MsgBox(sUserMessage, vbYesNo + vbQuestion, "�폜�͈͂̊m�F")
        
        If myYesNo = vbYes Then
            On Error GoTo ErrHandler
            rngToDelete.Delete
            Cells(1, 1).Select
            MsgBox "�w�b�_�[�̍폜���������܂���", vbOKOnly + vbInformation, "�폜����"
        Else
            If Not CurrentActivecell Is Nothing Then CurrentActivecell.Select '�A�N�e�B�u�Z���������̏�Ԃɖ߂�
            mbExitProc = True
            Exit Sub
        End If
    End If

Exit Sub

ErrHandler:
MsgBox "�폜���ɃG���[���N���܂���", vbOKOnly + vbInformation

End Sub

'---------------------------------------------------------------------------------------
' Method : DeleteLogo
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �w�b�_�[�Ɋ܂܂�郍�S�摜���폜����
'---------------------------------------------------------------------------------------
Private Sub DeleteLogo(ByVal NumPage As Long)

    Dim shp As Shape
    Dim i As Long
    Dim rng_shp As Range
    Dim rng As Range

    '���S������
    For i = 37 To NumPage * 60 Step 2   '�ŏ�����Ō�܂�

        Set rng = Range(Cells(i, 1), Cells(i, 20))  '���S�����肻���ȗ�

        For Each shp In ActiveSheet.Shapes
            Set rng_shp = Range(shp.TopLeftCell, shp.BottomRightCell)

            If Not (Intersect(rng_shp, rng) Is Nothing) Then    '��������Range��Shape���d�Ȃ�����ȉ��̏����B
                shp.Delete
            End If
        Next
    Next i

    '�e�L�X�g�{�b�N�X�i�Ж�����j������
    For i = 37 To NumPage * 60 Step 2   '�ŏ�����Ō�܂�

        Set rng = Cells(i, 16)    '�e�L�X�g�{�b�N�X�����肻���ȗ�

        For Each shp In ActiveSheet.Shapes
        
            Set rng_shp = Range(shp.TopLeftCell, shp.BottomRightCell)
            If Not (Intersect(rng_shp, rng) Is Nothing) Then
                shp.Delete
            End If
        Next shp
    Next i
    
    Set rng_shp = Nothing
    Set rng = Nothing
    
End Sub

