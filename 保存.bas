Attribute VB_Name = "�ۑ�"
Option Explicit

Sub SavePDF2Desktop()
'�����N���b�N�ŃA�N�e�B�u�V�[�g��PDF�Ńf�X�N�g�b�v�ɕۑ�����
    Dim WSH As Variant
    Set WSH = CreateObject("WScript.Shell")
    Dim Path As String
    Path = WSH.SpecialFolders("Desktop") & "\"
    Path = Path & ActiveSheet.Name & ".pdf"
    On Error GoTo ERR
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=Path
    Set WSH = Nothing
    Exit Sub
ERR:
If ERR.Number = 1004 Then MsgBox "�󔒂̃V�[�g�ł�"
End Sub
