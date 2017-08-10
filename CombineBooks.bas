Attribute VB_Name = "CombineBooks"
Option Explicit

Const src_path As String = "c:\wd\�������A���[�����b�Z�[�W\omegat_kaketsuken_alarm_message_list\tm_source\"     '�����t�@�C��
Const tgt_path As String = "c:\wd\�������A���[�����b�Z�[�W\omegat_kaketsuken_alarm_message_list\tm_source\\"     '�󕶃t�@�C��
Const store_path As String = "c:\wd\�������A���[�����b�Z�[�W\omegat_kaketsuken_alarm_message_list\tm_source\result\" '�e�L�X�g�����������u�b�N�̕ۑ���

Sub start()
'�����Ɩ󕶂��������A�e�L�X�g�t�@�C���ɏo�͂���
BulkBookUtil.CloseAllBooks '���s���Ƀt�@�C�����J���Ă���Ƃւ�ɂȂ�
Application.DisplayAlerts = False
Application.ScreenUpdating = False
PasteTranslationToSrcBooks
EscapeExcelLineBreak.DeescapeBreaksBooks

'�G���[�ɂȂ� Ubound(sarr) ���傫����
'GetStringsFromMultiBooks.GetStringsFromBooks
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Private Sub PasteTranslationToSrcBooks()

'### �J�������ׂẴu�b�N���ЂƂ̃u�b�N�ɂ܂Ƃ߂�
'### �Z�����̉��s�̓G�X�P�[�v���āA1�Z����1�s�ɕϊ����Ă���B
Dim sFilePath As Variant, tfilePath As Variant
sFilePath = GetFileNames.GetFileNames(src_path)
tfilePath = GetFileNames.GetFileNames(tgt_path)

Application.DisplayAlerts = False

'�����Ɩ󕶂̃t�@�C������v�H
If UBound(sFilePath) <> UBound(tfilePath) Then
    MsgBox "�����Ɩ󕶂̃t�@�C��������v���܂���B", vbCritical
    Exit Sub
End If

BulkBookUtil.OpenAllBooks sFilePath, src_path

Dim fileCount As Long
fileCount = UBound(tfilePath) 'Files.CountFileNumber(src_path)

Dim sarr() As Variant
ReDim sarr(fileCount) As Variant
Dim tarr() As Variant
ReDim tarr(fileCount) As Variant
Dim c: c = 0

Dim wks  As Worksheet, wkb As Workbook
Dim i As Long, l As Long
For l = LBound(sFilePath) To UBound(sFilePath)
Set wkb = Workbooks(sFilePath(l))
    If wkb.Name <> "PERSONAL.XLSB" Then 'Surface�����Ƃ���ŃG���[�ɂȂ�
        Dim src As Worksheet
        Set src = wkb.Sheets(1)

        '�S�͈͂�z��ɓ����
        Dim r As Range
        Set r = src.UsedRange
        Dim var As Variant
        var = r.Columns(1) '1��ڂ���
        sarr(c) = var
        c = c + 1
        On Error GoTo 0
        Erase var
        Set src = Nothing
    End If
Next l

Timers.WaitForSeconds (1) 'Wait���Ȃ��ƃu�b�N����鎞�ɃG���[�ɂȂ�
BulkBookUtil.CloseAllBooks
BulkBookUtil.OpenAllBooks tfilePath, tgt_path

c = 0
For l = LBound(sFilePath) To UBound(sFilePath)
    Set wkb = Workbooks(tfilePath(l))

    If wkb.Name <> "PERSONAL.XLSB" And wkb.Name <> "Book1" Then
        Set src = wkb.Sheets(1)

        '�S�͈͂�z��ɓ����
        Set r = src.UsedRange
        var = r.Columns(2) '2��ڂ���
        tarr(c) = var
        c = c + 1
        Erase var
        Set src = Nothing
    End If
Next l

Dim vv As Variant
Dim k As Long

Dim fi: fi = 0
For k = LBound(sarr) To UBound(sarr)
    Dim fname As String

    fname = tfilePath(fi) 'file name
    fi = fi + 1

    Dim newbook As Workbook
    Set newbook = Workbooks.Add

    Dim st As Worksheet
    Set st = newbook.Sheets(1)

    Dim x
    For x = 1 To UBound(sarr(k))
        st.Cells(x, 1).Value = sarr(k)(x, 1)
        st.Cells(x, 2).Value = tarr(k)(x, 1)
    Next x

    newbook.Save
    'newbook.SaveAs FileName:=store_path & fname
Next k

Application.DisplayAlerts = True
Timers.WaitForSeconds (1) 'Wait���Ȃ��ƃu�b�N����鎞�ɃG���[�ɂȂ�
BulkBookUtil.CloseAllBooks

Set wks = Nothing
Set wkb = Nothing

End Sub


