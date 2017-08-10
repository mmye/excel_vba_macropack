VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �O���[���J�[�h���̓c�[�� 
   Caption         =   "�O���[���J�[�h���̓��[�e�B���e�B"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15180
   OleObjectBlob   =   "�O���[���J�[�h���̓c�[��.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�O���[���J�[�h���̓c�[��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Times() As Single, Days() As String
    Dim OverworkTimes() As Single
    Dim wksGreenCard As Worksheet
    Dim wksSource As Worksheet
    Dim CleanLists As Variant

Private Sub btnGetOverworkTime_Click()
    Call Main
End Sub
Private Sub Main()
    Dim ItemCnt As Long
    Dim OvertimeLists As Variant
    Dim DayLists As Variant
    Dim TimeandDayLists() As Variant
    Dim i As Long
    
    If ActiveWorkbook.Name Like "*monthly*" Then
        Set wksSource = ActiveSheet
    Else
        MsgBox "�^�C���J�[�h�̃G�N�Z���u�b�N���A�N�e�B�u�ɂ��Ă���ēx���s���Ă�������", vbInformation
    End If
        
    
'  �c�Ǝ��ԂƓ��t��GET���ă��X�g�{�b�N�X�ɓ����
    OvertimeLists = GetOverTime(wksSource, DayLists)
    ItemCnt = UBound(DayLists) - LBound(DayLists) + 1
    ReDim TimeandDayLists(ItemCnt - 1, 0 To 1) As Variant
    For i = LBound(DayLists) To UBound(DayLists)
        TimeandDayLists(i - 1, 0) = DayLists(i, 1) 'DayList��2�����ڂ�1����͂��܂�
        TimeandDayLists(i - 1, 1) = OvertimeLists(i)
    Next i
    
    CleanLists = RemoveEmptyRow(TimeandDayLists)
    Me.lbList.List = CleanLists
    
    Set wksSource = Nothing
End Sub

Private Function GetOverTime(wks, DayLists) As Variant
    Dim r As Range, rng As Range, rDays As Range
    Dim c As Long
    Dim v As Variant, vTimes As Variant
    Dim i As Long
    
    Set r = wks.UsedRange.Find("���l").offset(1, -1)
    If r Is Nothing Then
        MsgBox "�G���[�F�^�C���J�[�h��F���ł��܂���"
        Exit Function
    End If
    Set rng = Range(r, wks.Cells(Rows.Count, r.column).End(xlUp))
    
    Set rDays = wks.UsedRange.Find("���t").offset(1, 0)
    If r Is Nothing Then
        MsgBox "�G���[�F�^�C���J�[�h��F���ł��܂���"
        Exit Function
    End If
    Set rDays = Range(rDays, wks.Cells(Rows.Count, rDays.column).End(xlUp))
    DayLists = rDays
    vTimes = rng
    ReDim vdsays(rDays.Count) As String
    ReDim Times(rDays.Count) As Single
    
    Dim Time As Single
    For Each v In vTimes
        Time = v - 8.75
        If Time <= 0 Then
            Times(c) = 0
        Else
            Times(c) = Time
        End If
        c = c + 1
    Next v
    
    ReDim OverworkTimes(rDays.Count) As Single
    For i = LBound(Times) To UBound(Times)
        If Times(i) > 0 Then OverworkTimes(i) = Times(i)
    Next i
    
    GetOverTime = OverworkTimes

    Set rng = Nothing
    Set rDays = Nothing
    Set r = Nothing

End Function

Private Sub btnQuit_Click()
    Unload Me
End Sub

Private Sub btnWriteOverworkTime_Click()
    Call WriteGreenCard
End Sub

Sub ManualWrite()
    Selection = CleanLists
End Sub

Private Sub WriteGreenCard()
'   �O���[���J�[�h�̃��[�N�V�[�g�Ɏ擾���ă��X�g�{�b�N�X�ɓ����Ă���c�Ǝ��Ԃ�]�L���鏈��
    Dim r As Range, rng As Range
    Dim f As Boolean
    Dim i As Long, c As Long
    Dim r2 As Range
    Dim msg As String
    
    msg = "�Ɏc�Ǝ��Ԃ��L�����܂��B��낵���ł����H" & vbCrLf & _
        "�Ⴄ�V�[�g���A�N�e�B�u�ɂȂ��Ă���ꍇ�A�O���[���J�[�h�̃V�[�g��I�����Ă���ēx���s���Ă�������"
    f = MsgBox(ActiveSheet.Name & msg, vbYesNo + vbQuestion)
    Set wksGreenCard = ActiveSheet
    On Error GoTo ERR
    Set r = wksGreenCard.UsedRange.Find("Over-time").offset(1, 0)
    On Error GoTo 0
    If r Is Nothing Then
        MsgBox "�G���[�F�^�C���J�[�h��F���ł��܂���"
        Exit Sub
    End If
    Set rng = Range(r, Cells(Rows.Count, r.column).End(xlUp))
    
    rng = OverworkTimes
    
    For Each r2 In rng
        If r2.Value = 0 Then rng.Item(c) = Empty
        c = c + 1
    Next r2
        
    Set rng = Nothing
    Set r = Nothing

    MsgBox "�O���[���J�[�h�Ɏc�Ǝ��Ԃ���͂��܂����B", vbInformation
Exit Sub
ERR:
    MsgBox "�O���[���J�[�h�̔F���Ɏ��s���܂����B�������V�[�g���A�N�e�B�u�ɂȂ��Ă��邱�Ƃ��m�F���āA�Ď��s���Ă�������", vbInformation
End Sub

Private Sub UserForm_Initialize()
'    Call RegisterWkstoCMB
End Sub

Private Sub UserForm_Terminate()
    Set wksGreenCard = Nothing
    Set wksSource = Nothing
End Sub

Private Function RemoveEmptyRow(Lists)
    Dim i As Long
    Dim ListCnt As Long
    Dim Lists2() As Variant
    Dim c As Long
    
    ListCnt = UBound(Lists) - LBound(Lists) + 1
    ReDim Lists2(ListCnt, 0 To 1) As Variant
    
    For i = LBound(Lists) To UBound(Lists)
        If Lists(i, 0) <> "" Then
            Lists2(c, 0) = Lists(i, 0)
            Lists2(c, 1) = Lists(i, 1)
            c = c + 1
        End If
    Next i
    RemoveEmptyRow = Lists2

End Function

Private Sub RegisterWkstoCMB()
    Dim wkb As Workbook
    Dim wks As Worksheet
    
    For Each wkb In Workbooks
        For Each wks In wkb
        Next
    Next
End Sub
