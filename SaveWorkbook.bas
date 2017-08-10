Attribute VB_Name = "SaveWorkbook"
Option Explicit

Sub SaveActiveworkbook()
    Dim myPath As String
    Dim myName As String
    Dim myFileName As String
    Dim wScriptHost As Object, strInitDir As String

    '�}�C�h�L�������g�̃p�X���擾
    Set wScriptHost = CreateObject("WScript.Shell")
    myPath = "C:\Documents" 'wScriptHost.SpecialFolders("MyDocuments")
    
    '�t�@�C�����̓��͂����߂�
    Do While myName = Empty
        myName = InputBox("�ۑ�����u�b�N�̖��O����͂��Ă�������")
        If myName = "" Then Exit Sub
    Loop
    
    '�p�X���쐬
    myFileName = myPath & "\" & myName
    '�}�N�����܂ޏꍇ��.xlsm�ŕۑ�����悤�ɏ�������
    If ActiveWorkbook.HasVBProject Then
        ActiveWorkbook.SaveAs fileName:=myFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        MsgBox myFileName & ".xlsm��ۑ����܂���", vbInformation + vbOKOnly, "���m�点"
    Else: ActiveWorkbook.SaveAs fileName:=myFileName, FileFormat:=xlWorkbookDefault
        MsgBox myFileName & ".xlsx��ۑ����܂���", vbInformation + vbOKOnly, "���m�点"
    End If
    
    Set wScriptHost = Nothing
End Sub

Sub DuplicateActiveSheet()
    Dim Name As String
    Name = ActiveSheet.Name
    ActiveSheet.Copy
    ActiveWorkbook.Sheets(1).Name = Name & "_�R�s�["
    SaveActiveworkbook
    MsgBox Name & "�̃R�s�[��ۑ����܂����B", vbOKOnly, "���m�点"
End Sub
