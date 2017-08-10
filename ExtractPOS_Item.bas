Attribute VB_Name = "ExtractPOS_Item"
Option Explicit

Sub GetPos_Item()
    Dim i As Long
    Dim c As Long: c = 1

    Dim EndRow As Long
    EndRow = InputBox("�Ō�̍��ڂ̍s�ԍ�����͂��Ă�������...")
    If EndRow = 0 Then Exit Sub
    If Not IsNumeric(EndRow) Then
        MsgBox "���l�Ŏw�肵�Ă�������"
        Exit Sub
    End If

    Dim POSCol As String
    POSCol = InputBox("POS�̗񖼂��A���t�@�x�b�g�œ��͂��Ă�������...")
    If StrConv(POSCol, vbHiragana) <> POSCol Then Exit Sub

    Dim ItemCol As String
    ItemCol = InputBox("�i���̗񖼂��A���t�@�x�b�g�œ��͂��Ă�������...")
    If StrConv(ItemCol, vbHiragana) <> ItemCol Then Exit Sub

    Dim st As Worksheet
    Set st = ActiveWorkbook.Worksheets.Add

    For i = 40 To EndRow
        If Cells(i, POSCol).Value <> "" Then
            Dim POS As String
            POS = Cells(i, POSCol).Value
            Dim Item As String
            Item = Cells(i, ItemCol).Value
            st.Cells(c, "A").Value = POS
            st.Cells(c, "B").Value = Item
            c = c + 1
         End If
    Next i

    st.Select
    st.Cells(1, 1).Select
    MsgBox "POS�ƕi���𒊏o���܂���"

End Sub


