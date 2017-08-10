Attribute VB_Name = "保存"
Option Explicit

Sub SavePDF2Desktop()
'ワンクリックでアクティブシートをPDFでデスクトップに保存する
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
If ERR.Number = 1004 Then MsgBox "空白のシートです"
End Sub
