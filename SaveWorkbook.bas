Attribute VB_Name = "SaveWorkbook"
Option Explicit

Sub SaveActiveworkbook()
    Dim myPath As String
    Dim myName As String
    Dim myFileName As String
    Dim wScriptHost As Object, strInitDir As String

    'マイドキュメントのパスを取得
    Set wScriptHost = CreateObject("WScript.Shell")
    myPath = "C:\Documents" 'wScriptHost.SpecialFolders("MyDocuments")
    
    'ファイル名の入力を求める
    Do While myName = Empty
        myName = InputBox("保存するブックの名前を入力してください")
        If myName = "" Then Exit Sub
    Loop
    
    'パスを作成
    myFileName = myPath & "\" & myName
    'マクロを含む場合は.xlsmで保存するように条件分岐
    If ActiveWorkbook.HasVBProject Then
        ActiveWorkbook.SaveAs fileName:=myFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        MsgBox myFileName & ".xlsmを保存しました", vbInformation + vbOKOnly, "お知らせ"
    Else: ActiveWorkbook.SaveAs fileName:=myFileName, FileFormat:=xlWorkbookDefault
        MsgBox myFileName & ".xlsxを保存しました", vbInformation + vbOKOnly, "お知らせ"
    End If
    
    Set wScriptHost = Nothing
End Sub

Sub DuplicateActiveSheet()
    Dim Name As String
    Name = ActiveSheet.Name
    ActiveSheet.Copy
    ActiveWorkbook.Sheets(1).Name = Name & "_コピー"
    SaveActiveworkbook
    MsgBox Name & "のコピーを保存しました。", vbOKOnly, "お知らせ"
End Sub
