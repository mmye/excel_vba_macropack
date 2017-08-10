Attribute VB_Name = "ZebraPint"
Option Explicit
 
Private Declare Function ChooseColor Lib "comdlg32.dll" _
    Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
 
Private Type ChooseColor
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
 
Private Const CC_RGBINIT = &H1
Private Const CC_LFULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SHOWHELP = &H8

Sub ZebraPint()
    Const DefaultColor = 15921906
    Dim InteriorColor As Long
    Dim r As Integer
    Dim lastCol As Integer
    Dim FirstRow As Integer
    Dim firstCol As Integer
    Dim LastRow As Integer

    '選択範囲の行を交互に薄灰で塗りつぶし
    Selection.Interior.ColorIndex = xlNone 'まず塗りつぶしなし

    InteriorColor = GetColorDlg(DefaultColor) '既定の色が選択された状態でダイアログが開く
    FirstRow = Selection(1).Row
    firstCol = Selection(1).column
    LastRow = Selection(Selection.Count).Row
    lastCol = Selection(Selection.Count).column
    
    For r = FirstRow To LastRow Step 2
        Range(Cells(r, firstCol), _
        Cells(r, lastCol)).Interior.color = InteriorColor '定数で塗りつぶしいろを定義
    Next r
End Sub

Private Function GetColorDlg(lngDefColor As Long) As Long
 
  Dim udtChooseColor As ChooseColor
  Dim lngRet As Long
 
  With udtChooseColor 'ダイアログの設定
    .lStructSize = Len(udtChooseColor)
    .lpCustColors = String$(64, Chr$(0))
    .flags = CC_RGBINIT + CC_LFULLOPEN
    .rgbResult = lngDefColor
    
    lngRet = ChooseColor(udtChooseColor) 'ダイアログを表示
    
    If lngRet <> 0 Then 'ダイアログからの戻り値をチェック
      If .rgbResult > RGB(255, 255, 255) Then
        GetColorDlg = -2 'エラーの場合
      Else
        GetColorDlg = .rgbResult '戻り値にRGB値を代入
      End If
    Else
      GetColorDlg = -1 'キャンセルされた場合
    End If
   End With
 
End Function


