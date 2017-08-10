Attribute VB_Name = "GetStringsFromMultiBooks"

Option Explicit

Sub GetStringsFromBooks()

'### エクセルブックに含まれる原文・訳文をテキストファイルに区切りつきでまとめるスクリプト
'### セル内の改行はエスケープして、1セル＝1行に変換している。

Dim p As String
p = "C:\000_Kaketsuken_Japanese_HMI\combined\"
'p = "C:\000_Kaketsuken_Japanese_HMI\target\"
'p = "C:\000_Kaketsuken_Japanese_HMI\source\"

Dim outPath As String '全ブックすべて内容をテキストに出力する
outPath = p & "integrity_test_src.txt"
On Error Resume Next
Kill outPath ' Append するので最初に消しておく
On Error GoTo 0

Dim Path As Variant
Path = GetFileNames.GetFileNames(p)

BulkBookUtil.OpenAllBooks Path, p
'EventSwitch

Dim wks As Worksheet, wkb As Workbook
For Each wkb In Workbooks
    Dim srcSt As Worksheet
    Set srcSt = wkb.Sheets(1)

    '全範囲を配列に入れる
    Dim r As Range
    Set r = srcSt.UsedRange
    Dim vSrc As Variant, vTgt As Variant
    vSrc = r.Columns(1) ' 原文
    vTgt = r.Columns(2) ' 和訳
    
    '空のワークブックを飛ばす（Book1とかをひらいたまま実行してエラーになりがち）
    If (Not IsArray(vSrc)) Or (Not IsArray(vTgt)) Then GoTo NextWkb
    
    Dim i As Long
    Dim cSrc As Long
    Dim cTgt As Long '行数カウント
    
    cSrc = cSrc + (UBound(vSrc) - LBound(vSrc))
    cTgt = cTgt + (UBound(vTgt) - LBound(vTgt))
    
    Debug.Print "wkb name: " & wkb.Name & "   src line count: " & cSrc & "   tgt line count: " & cTgt
    
    For i = LBound(vSrc) To UBound(vSrc)
        vSrc(i, 1) = Replace(vSrc(i, 1), vbLf, "\n")
    Next
    For i = LBound(vTgt) To UBound(vTgt)
        vTgt(i, 1) = Replace(vTgt(i, 1), vbLf, "\n")
    Next

    '原文と訳文の行数をチェック
    If (UBound(vSrc) = UBound(vTgt)) Then
        ' Line count matche => OK!
    Else
        Dim b As VbMsgBoxResult
        b = MsgBox("原文と訳文の行数が異なります。続けますか？", vbYesNo + vbQuestion)
        If b = vbNo Then Exit Sub
    End If

    Dim v As Variant
    Const Separator As String = "|"

    Open outPath For Append As #1
    For i = LBound(vSrc) To UBound(vSrc)
        Dim src As String, tgt As String
        src = vSrc(i, 1)
        tgt = vTgt(i, 1)
        Print #1, src & Separator & tgt
    Next i
    Close #1
NextWkb:
Next wkb

Timers.WaitForSeconds (1) 'Waitしないとブックを閉じる時にエラーになる
BulkBookUtil.CloseAllBooks
Set wks = Nothing
Set wkb = Nothing

'EventSwitch

Dim msg As String
msg = "完了しました" & vbCrLf & "原文行数：" & cSrc & vbCrLf _
       & "訳分行数" & cTgt
      
MsgBox msg, vbInformation

outPath = """" & outPath & """" 'file openするために必要なエスケープ
CreateObject("Wscript.Shell").Run outPath '既定のエディタで出力ファイルを開く

End Sub


Private Sub EventSwitch()
    With Application
        .DisplayAlerts = Not .Application.DisplayAlerts
        .ScreenUpdating = Not .ScreenUpdating
        .Visible = Not .Visible
    End With
End Sub


