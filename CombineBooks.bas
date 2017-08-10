Attribute VB_Name = "CombineBooks"
Option Explicit

Const src_path As String = "c:\wd\化血研アラームメッセージ\omegat_kaketsuken_alarm_message_list\tm_source\"     '原文ファイル
Const tgt_path As String = "c:\wd\化血研アラームメッセージ\omegat_kaketsuken_alarm_message_list\tm_source\\"     '訳文ファイル
Const store_path As String = "c:\wd\化血研アラームメッセージ\omegat_kaketsuken_alarm_message_list\tm_source\result\" 'テキストを結合したブックの保存先

Sub start()
'原文と訳文を結合し、テキストファイルに出力する
BulkBookUtil.CloseAllBooks '実行時にファイルが開いているとへんになる
Application.DisplayAlerts = False
Application.ScreenUpdating = False
PasteTranslationToSrcBooks
EscapeExcelLineBreak.DeescapeBreaksBooks

'エラーになる Ubound(sarr) が大きすぎ
'GetStringsFromMultiBooks.GetStringsFromBooks
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Private Sub PasteTranslationToSrcBooks()

'### 開いたすべてのブックをひとつのブックにまとめる
'### セル内の改行はエスケープして、1セル＝1行に変換している。
Dim sFilePath As Variant, tfilePath As Variant
sFilePath = GetFileNames.GetFileNames(src_path)
tfilePath = GetFileNames.GetFileNames(tgt_path)

Application.DisplayAlerts = False

'原文と訳文のファイル数一致？
If UBound(sFilePath) <> UBound(tfilePath) Then
    MsgBox "原文と訳文のファイル数が一致しません。", vbCritical
    Exit Sub
End If

BulkBookUtil.OpenAllBooks sFilePath, src_path

Dim fileCount As Long
fileCount = UBound(tfilePath) 'Files.CountFileNumber(src_path)

Dim sarr() As Variant
ReDim sarr(fileCount) As Variant
Dim tarr() As Variant
ReDim tarr(fileCount) As Variant
Dim c: c = 0

Dim wks  As Worksheet, wkb As Workbook
Dim i As Long, l As Long
For l = LBound(sFilePath) To UBound(sFilePath)
Set wkb = Workbooks(sFilePath(l))
    If wkb.Name <> "PERSONAL.XLSB" Then 'Surface環境だとこれでエラーになる
        Dim src As Worksheet
        Set src = wkb.Sheets(1)

        '全範囲を配列に入れる
        Dim r As Range
        Set r = src.UsedRange
        Dim var As Variant
        var = r.Columns(1) '1列目だけ
        sarr(c) = var
        c = c + 1
        On Error GoTo 0
        Erase var
        Set src = Nothing
    End If
Next l

Timers.WaitForSeconds (1) 'Waitしないとブックを閉じる時にエラーになる
BulkBookUtil.CloseAllBooks
BulkBookUtil.OpenAllBooks tfilePath, tgt_path

c = 0
For l = LBound(sFilePath) To UBound(sFilePath)
    Set wkb = Workbooks(tfilePath(l))

    If wkb.Name <> "PERSONAL.XLSB" And wkb.Name <> "Book1" Then
        Set src = wkb.Sheets(1)

        '全範囲を配列に入れる
        Set r = src.UsedRange
        var = r.Columns(2) '2列目だけ
        tarr(c) = var
        c = c + 1
        Erase var
        Set src = Nothing
    End If
Next l

Dim vv As Variant
Dim k As Long

Dim fi: fi = 0
For k = LBound(sarr) To UBound(sarr)
    Dim fname As String

    fname = tfilePath(fi) 'file name
    fi = fi + 1

    Dim newbook As Workbook
    Set newbook = Workbooks.Add

    Dim st As Worksheet
    Set st = newbook.Sheets(1)

    Dim x
    For x = 1 To UBound(sarr(k))
        st.Cells(x, 1).Value = sarr(k)(x, 1)
        st.Cells(x, 2).Value = tarr(k)(x, 1)
    Next x

    newbook.Save
    'newbook.SaveAs FileName:=store_path & fname
Next k

Application.DisplayAlerts = True
Timers.WaitForSeconds (1) 'Waitしないとブックを閉じる時にエラーになる
BulkBookUtil.CloseAllBooks

Set wks = Nothing
Set wkb = Nothing

End Sub


