Attribute VB_Name = "GetStringsFromMultiBooks"

Option Explicit

Sub GetStringsFromBooks()

'### �G�N�Z���u�b�N�Ɋ܂܂�錴���E�󕶂��e�L�X�g�t�@�C���ɋ�؂���ł܂Ƃ߂�X�N���v�g
'### �Z�����̉��s�̓G�X�P�[�v���āA1�Z����1�s�ɕϊ����Ă���B

Dim p As String
p = "C:\000_Kaketsuken_Japanese_HMI\combined\"
'p = "C:\000_Kaketsuken_Japanese_HMI\target\"
'p = "C:\000_Kaketsuken_Japanese_HMI\source\"

Dim outPath As String '�S�u�b�N���ׂē��e���e�L�X�g�ɏo�͂���
outPath = p & "integrity_test_src.txt"
On Error Resume Next
Kill outPath ' Append ����̂ōŏ��ɏ����Ă���
On Error GoTo 0

Dim Path As Variant
Path = GetFileNames.GetFileNames(p)

BulkBookUtil.OpenAllBooks Path, p
'EventSwitch

Dim wks As Worksheet, wkb As Workbook
For Each wkb In Workbooks
    Dim srcSt As Worksheet
    Set srcSt = wkb.Sheets(1)

    '�S�͈͂�z��ɓ����
    Dim r As Range
    Set r = srcSt.UsedRange
    Dim vSrc As Variant, vTgt As Variant
    vSrc = r.Columns(1) ' ����
    vTgt = r.Columns(2) ' �a��
    
    '��̃��[�N�u�b�N���΂��iBook1�Ƃ����Ђ炢���܂܎��s���ăG���[�ɂȂ肪���j
    If (Not IsArray(vSrc)) Or (Not IsArray(vTgt)) Then GoTo NextWkb
    
    Dim i As Long
    Dim cSrc As Long
    Dim cTgt As Long '�s���J�E���g
    
    cSrc = cSrc + (UBound(vSrc) - LBound(vSrc))
    cTgt = cTgt + (UBound(vTgt) - LBound(vTgt))
    
    Debug.Print "wkb name: " & wkb.Name & "   src line count: " & cSrc & "   tgt line count: " & cTgt
    
    For i = LBound(vSrc) To UBound(vSrc)
        vSrc(i, 1) = Replace(vSrc(i, 1), vbLf, "\n")
    Next
    For i = LBound(vTgt) To UBound(vTgt)
        vTgt(i, 1) = Replace(vTgt(i, 1), vbLf, "\n")
    Next

    '�����Ɩ󕶂̍s�����`�F�b�N
    If (UBound(vSrc) = UBound(vTgt)) Then
        ' Line count matche => OK!
    Else
        Dim b As VbMsgBoxResult
        b = MsgBox("�����Ɩ󕶂̍s�����قȂ�܂��B�����܂����H", vbYesNo + vbQuestion)
        If b = vbNo Then Exit Sub
    End If

    Dim v As Variant
    Const Separator As String = "|"

    Open outPath For Append As #1
    For i = LBound(vSrc) To UBound(vSrc)
        Dim src As String, tgt As String
        src = vSrc(i, 1)
        tgt = vTgt(i, 1)
        Print #1, src & Separator & tgt
    Next i
    Close #1
NextWkb:
Next wkb

Timers.WaitForSeconds (1) 'Wait���Ȃ��ƃu�b�N����鎞�ɃG���[�ɂȂ�
BulkBookUtil.CloseAllBooks
Set wks = Nothing
Set wkb = Nothing

'EventSwitch

Dim msg As String
msg = "�������܂���" & vbCrLf & "�����s���F" & cSrc & vbCrLf _
       & "�󕪍s��" & cTgt
      
MsgBox msg, vbInformation

outPath = """" & outPath & """" 'file open���邽�߂ɕK�v�ȃG�X�P�[�v
CreateObject("Wscript.Shell").Run outPath '����̃G�f�B�^�ŏo�̓t�@�C�����J��

End Sub


Private Sub EventSwitch()
    With Application
        .DisplayAlerts = Not .Application.DisplayAlerts
        .ScreenUpdating = Not .ScreenUpdating
        .Visible = Not .Visible
    End With
End Sub


