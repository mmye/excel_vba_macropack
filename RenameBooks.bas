Attribute VB_Name = "RenameBooks"
Option Explicit

Sub RenameBooks()

'### 指定フォルダにあるブック名を使用して別フォルダのファイル名を変更する
' ## ファイル取得の順番がおかしくなる　例の連番が正しくならない問題　1,10,11,みたいな

Dim src_path As String
src_path = "C:\000_Kaketsuken_Japanese_HMI\source\"

'テキストを結合したブックの保存先
Dim tgt_path As String
tgt_path = "C:\000_Kaketsuken_Japanese_HMI\combined\"

Dim new_name As Variant
new_name = GetFileNames.GetFileNames(src_path)

Dim old_name As Variant
old_name = GetFileNames.GetFileNames(tgt_path)

Dim v

For Each v In new_name
    Name tgt_path & old_name As tgt_path & new_name
Next v


End Sub
