Attribute VB_Name = "�V�[�g����"
Option Explicit

Public Sub CopyActiveSheetToLastPosition()
    Dim st As Worksheet
    Dim Copied As Worksheet
    
    Set st = ActiveSheet
    st.Copy After:=Sheets(Sheets.Count)
    Set Copied = Sheets(Sheets.Count)
    Copied.Select
    Copied.Range("a1").Activate
    MsgBox st.Name & "��" & Copied.Name & "�ɃR�s�[���܂���"

End Sub

