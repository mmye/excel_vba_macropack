Attribute VB_Name = "PDF連結"
Option Explicit
Const fileName As String = "mergedpd"
Const OutputFolderPath As String = "d:\buf"

'このサイトからコピペしてきた。
'PDFtkを使ってPDFを連結する。
'2017/02/09

Public Sub fncMergeFile()

    Dim WSH As Object, cmd As String, ret

    Const ExecPath = "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe"

    Set WSH = CreateObject("WScript.Shell")

    '一時フォルダを作成
    Dim TmpDir As String
    TmpDir = fncCreateTmpFolder

    cmd = Chr(34) & ExecPath & Chr(34) & _
            TmpDir & "\*.pdf " & _
            "cat " & _
            "output " & OutputFolderPath & "\" & fileName & ".pdf"

    ret = WSH.Run(cmd, 0, True)

    Set WSH = Nothing

    If ret <> 0 Then GoTo FSO_ERR

    '一時フォルダを削除
    fncDeleteTmpFolder TmpDir
    Exit Sub

FSO_ERR:

    MsgBox ERR.Description
    fncDeleteTmpFolder TmpDir

End Sub


'PDF用一時フォルダー 作成関数
' - PDFファイル群は%USERPROFILE%\AppData\Local配下に作成される
' - フォルダ名はRadXXXX
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

'PDF用一時フォルダー 削除関数
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

