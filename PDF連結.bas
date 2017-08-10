Attribute VB_Name = "PDF�A��"
Option Explicit
Const fileName As String = "mergedpd"
Const OutputFolderPath As String = "d:\buf"

'���̃T�C�g����R�s�y���Ă����B
'PDFtk���g����PDF��A������B
'2017/02/09

Public Sub fncMergeFile()

    Dim WSH As Object, cmd As String, ret

    Const ExecPath = "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe"

    Set WSH = CreateObject("WScript.Shell")

    '�ꎞ�t�H���_���쐬
    Dim TmpDir As String
    TmpDir = fncCreateTmpFolder

    cmd = Chr(34) & ExecPath & Chr(34) & _
            TmpDir & "\*.pdf " & _
            "cat " & _
            "output " & OutputFolderPath & "\" & fileName & ".pdf"

    ret = WSH.Run(cmd, 0, True)

    Set WSH = Nothing

    If ret <> 0 Then GoTo FSO_ERR

    '�ꎞ�t�H���_���폜
    fncDeleteTmpFolder TmpDir
    Exit Sub

FSO_ERR:

    MsgBox ERR.Description
    fncDeleteTmpFolder TmpDir

End Sub


'PDF�p�ꎞ�t�H���_�[ �쐬�֐�
' - PDF�t�@�C���Q��%USERPROFILE%\AppData\Local�z���ɍ쐬�����
' - �t�H���_����RadXXXX
Public Function fncCreateTmpFolder()
    Dim FSO As Object, TempName As String

    On Error GoTo FSO_ERR

            Set FSO = CreateObject("Scripting.FileSystemObject")

            With FSO
                TempName = .GetSpecialFolder(2) & "\" & .GetBaseName(.GetTempName)
                FSO.CreateFolder (TempName)
            End With

            Set FSO = Nothing

            fncCreateTmpFolder = TempName

            Exit Function

FSO_ERR:

    Debug.Print ERR.Description
    fncCreateTmpFolder = "-1"
    
End Function

'PDF�p�ꎞ�t�H���_�[ �폜�֐�
Public Sub fncDeleteTmpFolder(ByVal FolderName As String)
    Dim FSO As Object
    
    On Error GoTo FSO_ERR
    
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        FSO.DeleteFolder (FolderName)
        
        Set FSO = Nothing
        
        Exit Sub

FSO_ERR:

    Debug.Print ERR.Description

End Sub

