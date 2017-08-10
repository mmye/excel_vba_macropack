VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "置換"
   ClientHeight    =   1770
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7680
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Counter As Long
Dim bookNum As String

Private Sub CommandButton1_Click() '置換実行ボタン

Application.ScreenUpdating = False
Call ActivateTargetBook
Call ResetCounter
Call CheckBeforeReplacement
Call NormalReplacement
Call ShowReplacementCounts
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton4_Click() '「閉じる」ボタン
Unload Me
End Sub

Private Sub CommandButton2_Click() '「新規」ボタン

Call CreateListBook '新規リストブック作成

ComboBox1.Clear
Call GetListBook '再読み込み
Call WriteListBook
End Sub
Sub WriteListBook()
ComboBox1.Text = GetNumListbook
End Sub
Private Sub CommandButton3_Click() 'リストブック選択ボタン

Dim bk As Object

    'リストブックが選択されていない場合
    If ComboBox1.Text = "" Then
        Dim wb As Workbook
        Dim bName As String
        bName = GetNumListbook
        ComboBox1.AddItem (bName)
        Workbooks(bName).Activate
    Else
    'リストブックが選択されている場合
        
        'リストブックが存在するかチェック
        For Each wb In Workbooks
            If wb.Name = ComboBox1.Text Then
                Dim flag As Boolean
                flag = True
            End If
        Next
            '存在しない場合
            If flag = False Then
                ComboBox1.Clear
                Call GetListBook
                MsgBox "そのリストブックは開かれていないようです"
            Else
            '存在する場合
                Workbooks(ComboBox1.Text).Activate
                
            End If
    End If
        
End Sub
Private Sub CommandButton5_Click()

   '置換するブックが選択されていない場合
    If ComboBox2.Text = "" Then
        Exit Sub
    Else
    
    '選択されている場合
       
    'ブックが存在するかチェック
        Dim wb As Workbook
        For Each wb In Workbooks
            If wb.Name = ComboBox2.Text Then
                Dim flag As Boolean
                flag = True
            End If
        Next
            '結果、存在しない場合
            If flag = False Then
                ComboBox2.Clear
                Call GetOpenbook
                MsgBox "そのブックは開かれていないようです"
            Else
            '存在する場合
                VBA.AppActivate Excel.Application.Caption
                'Windows(ComboBox2.Text).Activate
            End If
        End If
    End Sub

Private Sub UserForm_Initialize()

ComboBox1.SetFocus 'コマンドボタンにフォーカス

'TIPS
ComboBox1.ControlTipText = "置換用語句のブック"
ComboBox2.ControlTipText = "置換したいブック"
CommandButton1.ControlTipText = "置換を実行"
OptionButton1.ControlTipText = "文字列を赤に"
OptionButton2.ControlTipText = "文字列を青に"
OptionButton3.ControlTipText = "文字列を緑に"
OptionButton4.ControlTipText = "文字列の色そのまま"

StartUpPosition = 1 'フォームをエクセルの中央に表示

'タブオーダー
ComboBox1.TabIndex = 0
ComboBox2.TabIndex = 1
CommandButton2.TabIndex = 2
OptionButton1.TabIndex = 3
OptionButton2.TabIndex = 4
OptionButton3.TabIndex = 5
CommandButton1.TabIndex = 6

'コンボボックススタイル
ComboBox1.Style = fmStyleDropDownList
ComboBox2.Style = fmStyleDropDownList

'オプションボックス初期選択
OptionButton1 = True

Call ClearCombobox 'コンボボックスをクリア
Call GetOpenbook '開いているブックをすべて取る
Call GetListBook 'リストブックを取得

End Sub
Sub ShowReplacementCounts()

MsgBox Counter & "件置換しました"
End Sub
Sub ResetCounter()
Counter = 0 'モジュール変数をクリア。これがないと置換件数が累積されてしまう。
End Sub
Sub ActivateTargetBook()

'置換するブックをアクティブにしないと置換されない
If ComboBox2.Text <> "" Then
    Workbooks(ComboBox2.Text).Activate
End If
End Sub

Sub ClearCombobox()
    
    ComboBox1.Clear
    ComboBox2.Clear
End Sub
Sub GetOpenbook()
    
'開いているブックをすべて取得。コンボボックス２に入れる。
       
    Dim wb As Workbook
    For Each wb In Workbooks
        With wb.Sheets(1)
            If .Cells(1, 1) <> "検索する文字列" And .Cells(1, 2) <> "置換後の文字列" Then
                ComboBox2.AddItem wb.Name
            End If
        End With
    Next wb
End Sub
Sub GetListBook()

'リストを取得。開始行の文字列で判定。コンボボックス１に入れる。
    Dim i As Long, wb As Workbook
   
    For Each wb In Workbooks
        With wb.Sheets(1)
            If .Cells(1, 1) = "検索する文字列" And .Cells(1, 2) = "置換後の文字列" Then
                ComboBox1.AddItem wb.Name
            End If
        End With
    Next wb
End Sub
Function GetNumListbook()

'開いているブックを数え、最終ブック数を返す
Dim bkNum As Integer
bkNum = Workbooks.Count
Dim bkName As String
GetNumListbook = Workbooks(bkNum).Name

End Function
Sub CheckBeforeReplacement()

Dim flag As Boolean
flag = cDocFalse '関数で誤ったブック選択があるか検査
    If flag Then
        Exit Sub '判定ありで脱出
    End If

Dim TargetSt As Worksheet
Set TargetSt = Workbooks(ComboBox2.Text).Sheets(1)
Dim n As Long
n = CountCells(TargetSt) '空ではないセルの個数を返す

Dim flag2 As Boolean
flag2 = JudgeCellCounts(n)

    If flag2 = True Then
        Call PromptUsertoResume
    End If
End Sub
Sub PromptUsertoResume()

Dim flag As Integer
flag = MsgBox("処理に時間がかかる可能性があります。" & vbCrLf _
            & "続けますか？" _
        , vbYesNo + vbQuestion, "処理実行前の注意")

    If flag = vbYes Then
        Exit Sub
    Else
        End
    End If
End Sub

Sub NormalReplacement()
'リストとターゲットブックを変数に格納し、
'検索する語句と置換後の語句を変数に格納する
'そして、置換実行プロシージャに渡す
'戻り先はイベントプロシージャ（コマンドボタンクリック）

Dim listBook As Workbook
Dim TargetBook As Workbook
Dim strWhat As String
Dim strReplacement As String
Dim i As Integer

On Error GoTo ErrHandler

Set listBook = Workbooks(ComboBox1.Text)
Set TargetBook = Workbooks(ComboBox2.Text)

    With listBook.Sheets(1)
        
        For i = 2 To .Range("A10000").End(xlUp).Row
            strWhat = .Cells(i, 1).Value
            strReplacement = .Cells(i, 2).Value
            Call doReplace(strWhat, strReplacement, TargetBook) '置換実行
        Next i
    End With
Exit Sub

ErrHandler:
    Dim myMsg As String
    myMsg = "エラー番号：" & ERR.Number & vbCrLf & _
                "エラー内容：" & ERR.Description
    MsgBox myMsg
End Sub
Sub doReplace(ByRef strWhat As String, strReplacement As String, TargetBook As Workbook)

    With TargetBook.ActiveSheet
        
        Dim r As Range
        Dim start As Integer
        Dim ExplorerRange As Range
        Set ExplorerRange = .Range("A1", ActiveCell.SpecialCells(xlLastCell)) 'データがあるセル範囲を取得
    
        Dim FontColor As Integer
        FontColor = GetFontColor 'オプションボタンをしらべる
            
        For Each r In ExplorerRange
        
            start = 1
            Do While InStr(start, r.Value, strWhat) > 0
                start = InStr(start, r.Value, strWhat)
                r.Characters(start, Len(strWhat)).Delete
                r.Characters(start, 0).insert strReplacement
                r.Characters(start, Len(strReplacement)).Font.ColorIndex = FontColor
                Counter = Counter + 1
                start = start + Len(strReplacement)
                'Debug.Print Counter
            Loop
        Next
    End With
End Sub
Function GetFontColor()

    If OptionButton1 Then
        GetFontColor = 3 '赤
    End If
    If OptionButton2 Then
        GetFontColor = 5 '青
    End If
    If OptionButton3 Then
        GetFontColor = 50 '緑
    End If
    If OptionButton4 Then
        GetFontColor = 1 '黒
    End If
End Function
Function cDocFalse()

Dim List As String
Dim bk As String

List = ComboBox1.Text
bk = ComboBox2.Text

    If List = bk Then '１）同一の文書が選択されている場合
        MsgBox "置換リストとターゲットが同じです。"
        cDocFalse = True
        Exit Function
    End If
    
    If List = "" And bk = "" Then '２）いずれかの文書が選択されていない場合
        MsgBox "置換リストと置換するブックが選択されていません。"
        cDocFalse = True
        Exit Function
    End If
    
    If List <> "" And bk = "" Then '２）いずれかの文書が選択されていない場合
        MsgBox "置換するブックが選択されていません。"
        cDocFalse = True
        Exit Function
    End If
    
    If List = "" And bk <> "" Then '２）いずれかの文書が選択されていない場合
        MsgBox "置換リストが選択されていません。"
        cDocFalse = True
        Exit Function
    End If
    
    
End Function

Sub CreateListBook()

Dim wb As Workbook
Application.ScreenUpdating = False

Set wb = Workbooks.Add

    With wb.Sheets(1)
        .Cells(1, 1) = "検索する文字列"
        .Cells(1, 2) = "置換後の文字列"
        .Cells(2, 1).Select
        .Name = "1"
    End With

Call ResizeBook(wb)
Call LockPane(wb)

bookNum = GetNumListbook

Application.ScreenUpdating = True

End Sub

Sub ResizeBook(ByRef wb As Workbook)

'対象のワークブックを通常表示にします
    wb.Activate
    ActiveWindow.WindowState = xlNormal
'リボンを非表示
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
'スクロール範囲の制限
    wb.Sheets(1).ScrollArea = "A1:B1000"
'ワークブックのウィンドウサイズを変更します
    ActiveWindow.Height = 200
    ActiveWindow.Width = 300
    ActiveWindow.Top = 250
    ActiveWindow.Left = 250

    Call LayoutBook(wb)
End Sub
Sub LockPane(ByRef wb As Workbook)

'ウインドウ枠を先頭行で固定
    wb.Activate
    ActiveWindow.FreezePanes = True
End Sub
Sub LayoutBook(ByRef wb As Workbook)

'フォント
    With Sheets(1).Range("A1:B1")
        .Font.Name = "MS ゴシック"
        .Font.Size = 10
        .Font.Bold = True
        .Font.ColorIndex = 2
'塗りつぶし
        .Interior.ColorIndex = 1
'ラベル中央揃え
        .HorizontalAlignment = xlCenter
    End With
    
    With Sheets(1).Range("A2:B100")
        .Font.Name = "MS ゴシック"
        .Font.Size = 10
    End With

'罫線
    With Sheets(1).Range("A1:B100")
        .Borders.LineStyle = True
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With

'列幅
    With Sheets(1).Columns("A:B")
        .ColumnWidth = 23
    End With
End Sub

Function CountCells(st As Worksheet) As Long

'シート内のデータが入ったセル数を返す

With Workbooks(ComboBox2.Text).Sheets(1)

    Dim ActiveRng As Range
    Set ActiveRng = .Range("a1", ActiveCell.SpecialCells(xlLastCell))
    
    Dim rng As Range
    For Each rng In ActiveRng
        If rng.Value <> "" Then
            Dim n As Long
            n = n + 1
        End If
    Next
End With
    
CountCells = n

End Function

Function JudgeCellCounts(n As Long) As Boolean

Dim MaxLimit As Long
MaxLimit = 1500 '注意上限

If n > MaxLimit Then
    JudgeCellCounts = True
End If

End Function
