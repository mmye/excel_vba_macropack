Attribute VB_Name = "ZebraPint"
Option Explicit
 
Private Declare Function ChooseColor Lib "comdlg32.dll" _
    Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
 
Private Type ChooseColor
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
 
Private Const CC_RGBINIT = &H1
Private Const CC_LFULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SHOWHELP = &H8

Sub ZebraPint()
    Const DefaultColor = 15921906
    Dim InteriorColor As Long
    Dim r As Integer
    Dim lastCol As Integer
    Dim FirstRow As Integer
    Dim firstCol As Integer
    Dim LastRow As Integer

    '�I��͈͂̍s�����݂ɔ��D�œh��Ԃ�
    Selection.Interior.ColorIndex = xlNone '�܂��h��Ԃ��Ȃ�

    InteriorColor = GetColorDlg(DefaultColor) '����̐F���I�����ꂽ��ԂŃ_�C�A���O���J��
    FirstRow = Selection(1).Row
    firstCol = Selection(1).column
    LastRow = Selection(Selection.Count).Row
    lastCol = Selection(Selection.Count).column
    
    For r = FirstRow To LastRow Step 2
        Range(Cells(r, firstCol), _
        Cells(r, lastCol)).Interior.color = InteriorColor '�萔�œh��Ԃ�������`
    Next r
End Sub

Private Function GetColorDlg(lngDefColor As Long) As Long
 
  Dim udtChooseColor As ChooseColor
  Dim lngRet As Long
 
  With udtChooseColor '�_�C�A���O�̐ݒ�
    .lStructSize = Len(udtChooseColor)
    .lpCustColors = String$(64, Chr$(0))
    .flags = CC_RGBINIT + CC_LFULLOPEN
    .rgbResult = lngDefColor
    
    lngRet = ChooseColor(udtChooseColor) '�_�C�A���O��\��
    
    If lngRet <> 0 Then '�_�C�A���O����̖߂�l���`�F�b�N
      If .rgbResult > RGB(255, 255, 255) Then
        GetColorDlg = -2 '�G���[�̏ꍇ
      Else
        GetColorDlg = .rgbResult '�߂�l��RGB�l����
      End If
    Else
      GetColorDlg = -1 '�L�����Z�����ꂽ�ꍇ
    End If
   End With
 
End Function


