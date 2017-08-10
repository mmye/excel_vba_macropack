Attribute VB_Name = "ToggleBRtoMetaChar"
Option Explicit

Dim Mode As String
'Excelセル内の改行をメタ文字に置換、またはその逆をやるスクリプト
'2017/04/18


Sub Run()
    Dim Path As String
    Path = SelectFolderDialog("Excelブックがあるフォルダを選択")
    If Path = "" Then Exit Sub
    Mode = InputBox("Enter 1:Char(10)⇒\n 2:\n⇒Char(10)")
    If Mode = "" Then Exit Sub
    ToggleBRtoMetaChar Path
End Sub

Private Sub ToggleBRtoMetaChar(Path)
    Dim File As String
    Dim f As Object
    Dim wkb As Workbook

    File = Dir(Path & Application.PathSeparator & "*.xls*")
    Do While File <> ""
        If Left$(File, 1) <> "~" Then 'バックアップファイルや隠しファイルみたいなものは避ける
            Set wkb = Workbooks.Open(Path & Application.PathSeparator & File, ReadOnly:=False)
            ActiveWindow.Visible = True
            SwitchBRtoMetaChar wkb
            wkb.Save
            wkb.Close
        End If
        File = Dir()
    Loop

    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            ToggleBRtoMetaChar (f.Path)
        Next f
    End With
    Set wkb = Nothing
    Set f = Nothing
End Sub

Sub SwitchBRtoMetaChar(ByRef wkb As Excel.Workbook)
    Dim st As Worksheet
    Select Case Mode
        Case 1 ' Char(10)-> \n
            For Each st In wkb.Sheets
                st.Cells.Replace What:=Chr(10), Replacement:="\n", SearchOrder:=xlByRows
            Next
        Case 2 '\n-> Char(10)
            For Each st In wkb.Sheets
                st.Cells.Replace What:="\n", Replacement:=Chr(10), SearchOrder:=xlByRows
            Next
    End Select
End Sub

Function SelectFolderDialog(Optional title As String, Optional buttonName As String, _
    Optional initialFileName As String) As String
    Dim fDialog As FileDialog, result As Integer, it As Variant
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    'Properties
    If buttonName <> vbNullString Then fDialog.buttonName = buttonName
    If initialFileName <> vbNullString Then fDialog.initialFileName = initialFileName
    If title <> vbNullString Then fDialog.title = title
    'Show
    If fDialog.Show = -1 Then
        SelectFolderDialog = fDialog.SelectedItems(1)
    End If
End Function


