Attribute VB_Name = "EscapeExcelLineBreak"
Option Explicit

Sub DeescapeBreaksBooks()

'### �^�[�Q�b�g�f�B���N�g���̂��ׂẴu�b�N�ŃG�X�P�[�v����
'### �Z�����̉��s�����ɖ߂��B

Application.DisplayAlerts = False
Dim p As String
p = "C:\000_Kaketsuken_Japanese_HMI\combined\"

Dim Path As Variant, tPath As Variant
Path = GetFileNames.GetFileNames(p)

If UBound(Path) = 0 Then
    MsgBox "�t�@�C�������擾�ł��܂���B�f�B���N�g�������������Ƃ��Ċm�F���Ă�������"
    Exit Sub
End If

BulkBookUtil.OpenAllBooks Path, p

Dim wks, wkb
For Each wkb In Workbooks
    If wkb.Name <> "PERSONAL.XLSB" Then 'Surface�����Ƃ���ŃG���[�ɂȂ�
        Dim src As Worksheet
        Set src = wkb.Sheets(1)
        
        Dim i As Long, col As Long
        For i = 1 To src.Cells(Rows.Count, 1).End(xlUp).Row
            For col = 1 To 2
                Dim buf As String
                buf = src.Cells(i, col).Value
                src.Cells(i, col).Value = Replace(buf, "\n", vbLf)
            Next
        Next
    End If
    
    Set src = Nothing
    wkb.Save
    wkb.Close
Next wkb

Application.DisplayAlerts = True
Set wkb = Nothing
Set src = Nothing

End Sub

Sub EscapeBreaksBooks_()

'### �J���Ă��邷�ׂẴu�b�N�ŃG�X�P�[�v����
'### �Z�����̉��s�����ɖ߂��B

Application.DisplayAlerts = False

Dim wks, wkb
For Each wkb In Workbooks
    If wkb.Name <> "PERSONAL.XLSB" Then 'Surface�����Ƃ���ŃG���[�ɂȂ�
        Dim src As Worksheet
        Set src = wkb.Sheets(1)
        
        Dim i As Long, col As Long
        For i = 1 To src.Cells(Rows.Count, 1).End(xlUp).Row
            For col = 1 To 2
                Dim buf As String
                buf = src.Cells(i, col).Value
                src.Cells(i, col).Value = Replace(buf, vbLf, "\n")
            Next
        Next
    End If
    
    Set src = Nothing
    wkb.Save
    wkb.Close
Next wkb

Application.DisplayAlerts = True
Set wkb = Nothing
Set src = Nothing

End Sub
