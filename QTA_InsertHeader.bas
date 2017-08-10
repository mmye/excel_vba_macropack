Attribute VB_Name = "QTA_InsertHeader"
#If VBA7 Then
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Option Explicit

Dim mbCancel As Boolean
Dim mlLastCol As Long
Dim mlLastRow As Long
Dim mbCancelEvent As Boolean

'---------------------------------------------------------------------------------------
' Method : InsertHeader
' Author : temporary3
' Date   : 2016/02/10
' Purpose: アクティブシートにヘッダーを挿入する。
'---------------------------------------------------------------------------------------
Sub InsertHeader()
    Dim CurrentActivecell   As Range
    Dim myLastRow As Long
    Dim myPageCount As Long
    Dim wks As Worksheet
    Dim myHeaderSpace As Range
    
    Set wks = ActiveWorkbook.ActiveSheet
    
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    mlLastRow = GetLastRow(wks) 'シートの引数を後で替える
    mlLastCol = GetLastCol(wks)
    myPageCount = GetPageBreak(wks)
    
    If TypeName(Selection) = "Range" Then Set CurrentActivecell = Selection
    
    Set myHeaderSpace = InsertRows(myPageCount)
    
    Call DrawLines(myPageCount)
    Call InsertDocNo(myPageCount)
    Call PasteLabels(myPageCount)
    Call InsertWincklerText(myPageCount)
    Call AddWincklerlogo(myPageCount)
    Call BreakPages(myPageCount)                'ヘッダー挿入後の改ページ設定
    Call RowsHeightAdjustment(myPageCount)
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If Not CurrentActivecell Is Nothing Then CurrentActivecell.Select
    
    MsgBox "ヘッダーを" & myPageCount & "ページまで追加しました", vbOKOnly + vbInformation, "ヘッダー挿入完了"

End Sub

'---------------------------------------------------------------------------------------
' Method : GetLastRow
' Author : mokoo
' Date   : 2016/02/20
' Purpose: シートの最下行を取得する
'---------------------------------------------------------------------------------------
Private Function GetLastRow(ByVal st As Worksheet) As Long
    Dim rngPrintArea As Range
    
    On Error GoTo ErrHandler
    Set rngPrintArea = st.Range(st.PageSetup.PrintArea)
    On Error GoTo 0
    GetLastRow = rngPrintArea.Item(rngPrintArea.Count).Row

Exit Function
 
ErrHandler:
 
 mbCancelEvent = True
 
 End Function
 
 Private Function GetLastCol(ByVal st As Worksheet) As Long
    Dim rngPrintArea As Range

    On Error GoTo ErrHandler
    Set rngPrintArea = st.Range(st.PageSetup.PrintArea)
    On Error GoTo 0
    GetLastCol = rngPrintArea.Item(rngPrintArea.Count).column

Exit Function
ErrHandler:
    MsgBox "ヘッダーがありません"
    
End Function
'---------------------------------------------------------------------------------------
' Method : GetPageCount
' Author : temporary3
' Date   : 2016/04/07
' Purpose: アクティブシートの見積書としてのページ数をカウントする（60行/ページで計算）
'---------------------------------------------------------------------------------------
Private Function GetPageCount() As Long
    Dim lPageCount As Long
    Dim rowCount As Long
    Dim lPageMargin As Long
    
        
    rowCount = ActiveWorkbook.ActiveSheet.UsedRange.Rows.Count

    If rowCount < 60 Then
        MsgBox "1ページしかありません"
        End    '行数が1ページ未満なら終了
    End If
    
    '規定行数ちょうどを含むページ数
    lPageCount = (rowCount / 60)
    '余りページ
    If (rowCount Mod 60) <> 0 Then lPageMargin = 1
    
    '合計ページ数
    GetPageCount = lPageCount + lPageMargin
    
End Function

Private Function GetPageBreak(ByVal wks As Worksheet) As Long
    
    GetPageBreak = wks.HPageBreaks.Count

End Function

'---------------------------------------------------------------------------------------
' Method : InsertRows
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ヘッダー用スペースとなる行を追加する。（12行/ヘッダー）
'---------------------------------------------------------------------------------------
Private Function InsertRows(ByVal myPageCount As Long) As Range

    Dim i As Long
    Dim iRowCountHeader As Long
    Dim rngRows As Range
    
    iRowCountHeader = 12    'ヘッダーの行数

    Set rngRows = Rows("1000:1012")

    For i = 60 To myPageCount * 60 Step 60    '最終セルの番号から末尾を決める
        DoEvents
        rngRows.Copy
        Rows(i).insert
        
        If InsertRows Is Nothing Then
             Set InsertRows = Rows(i & ":" & i + 11)
        Else
            Set InsertRows = Union(InsertRows, Rows(i & ":" & i + 11))
        End If
    Next i
                                                                                                                                              
End Function
'---------------------------------------------------------------------------------------
' Method : DrawLines
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ヘッダーの罫線を引く。
'---------------------------------------------------------------------------------------
Private Sub DrawLines(ByVal myPageCount As Long)

    Dim i As Long

    For i = 64 To myPageCount * 64 Step 60

        'ヘッダーの罫線を4本ひく
        '太い、細、細、細

        With Range(Cells(i, 1), Cells(i, mlLastCol)).Borders(xlEdgeBottom)  'ロゴの下のはじめの罫線は太い
            .LineStyle = xlContinuous
            .Weight = xlMedium    '中くらい
            .ColorIndex = xlAutomatic
        End With
        With Range(Cells(i, 1).offset(1, 0), Cells(i, mlLastCol).offset(1, 0)).Borders(xlEdgeBottom)  '二番目の罫線は細い
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Range(Cells(i, 1).offset(5, 0), Cells(i, mlLastCol).offset(5, 0)).Borders(xlEdgeBottom)  '三番目の罫線も細い
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Range(Cells(i, 1).offset(7, 0), Cells(i, mlLastCol).offset(7, 0)).Borders(xlEdgeBottom)    '一番下の罫線も細い
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next i

End Sub

'---------------------------------------------------------------------------------------
' Method : InsertDocNo
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ヘッダーごとに見積書のページ番号を書き込む
'---------------------------------------------------------------------------------------
Private Sub InsertDocNo(ByVal myPageCount As Long)

    Dim i As Long
    Dim Pages As Long
    Dim rngLabel As Range
    Dim Border1 As Border
    Const strLeftEdgePageNo As String = "CN"
    Const strRightEdgePageNo As String = "DF"
    Dim rngQuotationNo As Range
    Dim rngLabelQTNo As Range
    Dim sPageNumbering As String
    
    Pages = 2    'ページの開始は2

    For i = 61 To (myPageCount * 60 + 1) Step 60
        'ページ番号＆見積書番号のセル範囲
        Set rngLabel = Range(Cells(i, strLeftEdgePageNo), Cells(i, strRightEdgePageNo))

    '   1ページにある見積書番号をRangeにセット
        On Error GoTo NoNotFound
        Dim r
        Set r = Rows("1:6").Find("Nagoya")
        Set rngQuotationNo = Cells(r.Row - 1, r.column).Value
        Set rngLabelQTNo = Cells(i, strLeftEdgePageNo)
        
        rngQuotationNo.Copy Destination:=rngLabelQTNo
        rngQuotationNo.HorizontalAlignment = xlLeft
        
        On Error GoTo 0
      
        '見積書番号の2列右にページ数を入れる
        sPageNumbering = "'" & CStr(Pages) & "/" & CStr(myPageCount + 1)
        Cells(i, strRightEdgePageNo).offset(0, -2).Value = sPageNumbering
        Pages = Pages + 1
        
        'フォントの設定
        With rngLabel.Font
            .Name = "MS 明朝"
            .Size = 10
        End With

        '罫線
        Set Border1 = rngLabel.Borders(xlEdgeBottom)
        With Border1
            .LineStyle = xlDash
            .Weight = xlThin
        End With
    Next i
Exit Sub

NoNotFound:
    MsgBox "見積書番号が取得できませんでした。", vbOKOnly + vbCritical

End Sub

Private Sub PasteLabels(ByVal myPageCount As Long)
'ヘッダーのテキストを入れます。

    Dim i As Long
    'テキストの範囲を指定する変数
    Dim Item As Range
    Dim Description As Range
    Dim POS As Range
    Dim QTY As Range
    Dim UnitPrice As Range
    Dim TotalPrice As Range
    Dim Komoku As Range
    Dim Naiyo As Range
    Dim Designation As Range
    Dim rngHeaderWidth As Range
    Dim rngHeaderArea As Range

    Dim myCol_Item As Long
    Dim myCol_POS As Long
    Dim myCol_Total As Long
    
    '見積書の各要素の列番号を取得する
    myCol_Item = 見積書項目の列番号を取得(myCol_POS, myCol_Total)
    
    For i = 67 To (myPageCount * 60 + 7) Step 60
 
    ''取得した列番号を使用して見出しテキストの範囲を指定
    Set Item = Cells(i, myCol_Item)
    Set POS = Cells(i, myCol_POS)
    Set TotalPrice = Cells(i, myCol_Total)
    Set Komoku = Cells(i, "C")
    Set Naiyo = Cells(i, "BJ")
    Set Description = Cells(i, "bj").offset(1, 0)
    Set QTY = Cells(i, "BQ").offset(4, 0)
    Set UnitPrice = Cells(i, "BY").offset(4, 0)
    Set Designation = Cells(i, "S").offset(4, 0)


    'ヘッダテキストのフォントとサイズ
    Set rngHeaderWidth = Range(Cells(i - 7, "A"), Cells(i - 7, mlLastCol))
    Set rngHeaderArea = Range(rngHeaderWidth, rngHeaderWidth.offset(11, 0))
    
    With rngHeaderArea
        '.HorizontalAlignment = xlLeft
        .Font.Name = "ＭＳ Ｐ明朝"
        .Font.Size = 10
        .Interior.color = RGB(238, 238, 238)
    End With

    'ヘッダテキストの書き込みと位置調整
    '①
    With Komoku
        .Value = "項　　　目"
        .Font.Bold = False
    End With

    Range(Cells(i, "C"), Cells(i, "U")).HorizontalAlignment = xlCenterAcrossSelection
    '②
    With Item
        .Value = "Item"
        .Font.Bold = False
    End With
    Range(Cells(i, "C").offset(1, 0), Cells(i, "U").offset(1, 0)).HorizontalAlignment = xlCenterAcrossSelection
    
    '③
    With Naiyo
        .Value = "内　　　　容"
        .Font.Bold = False
    End With
    Range(Cells(i, "BJ"), Cells(i, "CV")).HorizontalAlignment = xlCenterAcrossSelection
    
    '④
    With Description
        .Value = "Description"
        .Font.Bold = False
    End With
    Range(Cells(i, "BJ").offset(1, 0), Cells(i, "CV").offset(1, 0)).HorizontalAlignment = xlCenterAcrossSelection
    
    '⑤
    With POS
        .Value = "Pos"
        .Font.Bold = False
    End With
    Range(Cells(i, "C").offset(4, 0), Cells(i, "F").offset(4, 0)).HorizontalAlignment = xlCenterAcrossSelection
    '⑥
    With Designation
        .Value = "品　  　名"
        .Font.Bold = False
    End With
    Range(Cells(i + 4, "X"), Cells(i + 4, "AF")).HorizontalAlignment = xlCenterAcrossSelection
    '⑦
    With QTY
        .Value = "数　量"
        .Font.Bold = False
    End With
    Range(Cells(i + 4, "BQ"), Cells(i + 4, "BW")).HorizontalAlignment = xlCenterAcrossSelection
    '⑧
    With UnitPrice
        .Value = "単　　価"
        .Font.Bold = False
    End With
    Range(Cells(i + 4, "BY"), Cells(i + 4, "CO")).HorizontalAlignment = xlCenterAcrossSelection
    '⑨
    With TotalPrice
        .Value = "価　　格"
        .Font.Bold = False
    End With
    Range(Cells(i + 4, "CR"), Cells(i + 4, "DK")).HorizontalAlignment = xlCenterAcrossSelection

    Next i

End Sub

'---------------------------------------------------------------------------------------
' Method : InsertWincklerText
' Author : temporary3
' Date   : 2016/02/10
' Purpose: 会社名のテキストを書き込む
'---------------------------------------------------------------------------------------
Private Sub InsertWincklerText(ByVal myPageCount As Long)

    Dim i As Long

    For i = 61 To (myPageCount * 61) Step 60
        With Cells(i, "M")
            .Value = "ウインクレル株式会社"
            .Font.Size = 14
            .Font.Name = "ＭＳ Ｐ明朝"
            .Font.Bold = False
        End With
        With Cells(i, "M").offset(1, 0)
            .Value = "WINCKLER & CO, LTD"
            .Font.Size = 14
            .Font.Name = "ＭＳ Ｐ明朝"
            .Font.Bold = False
        End With
    Next i

End Sub

'---------------------------------------------------------------------------------------
' Method : RowsHeightAdjustment
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ヘッダーの行高を設定する
'---------------------------------------------------------------------------------------
Private Sub RowsHeightAdjustment(ByVal myPageCount As Long)

    Dim i As Long

    For i = 60 To (myPageCount * 60) Step 60
        Rows(i).RowHeight = 12.5
        Rows(i + 1).RowHeight = 15.5    '「ウインクレル株式会社」
        Rows(i + 2).RowHeight = 15    'Winckler & Co, Ltd.
        Rows(i + 3).RowHeight = 12.5
        Rows(i + 4).RowHeight = 12.5
        Rows(i + 5).RowHeight = 3
        Rows(i + 6).RowHeight = 9.5
        Rows(i + 7).RowHeight = 14    '項目
        Rows(i + 8).RowHeight = 14    'Item
        Rows(i + 9).RowHeight = 9.5
        Rows(i + 10).RowHeight = 12.5
        Rows(i + 11).RowHeight = 12    'Pos
    Next i

End Sub
'---------------------------------------------------------------------------------------
' Method : AddWincklerlogo
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ウインクレルのロゴを貼り付ける
'---------------------------------------------------------------------------------------
Private Sub AddWincklerlogo(ByVal myPageCount As Long)

'ウインクレルのロゴを所定の位置に貼り付けます。

    Dim myFileName As String
    Dim myPic As Shape
    Dim i As Long
    Dim myFileSheet As Worksheet
    Dim spLogo As Shape
    Dim sShapePositionLEFT As Double
    Dim sShapePositionTOP As Double
    
    'ロゴのファイル名
    myFileName = "winckler_logo"
    'ロゴのパス
    Set myPic = ThisWorkbook.Sheets("pictures").Shapes("winckler_logo")
    'Set myFileSheet = ThisWorkbook.Worksheets("pictures")    '編集中のブックと同じディレクトリにある画像ファイルを指定しています。
    
    'ロゴをクリップボードに入れる
    myPic.Copy
    'With myFileSheet
    '    Set spLogo = .Shapes(myFileName)
    '    spLogo.Copy
    'End With

    On Error Resume Next    'なぜかエラーがでるから。
    For i = 61 To (myPageCount * 61) Step 60
        
        With ActiveWorkbook.ActiveSheet
            .Cells(i, "C").Select
            On Error GoTo ErrHandler
TryAgain:
            .Pastespecial Format:="図 (PNG)", Link:=False, DisplayAsIcon:=False '実行時エラー１００４が発生
            
            'たったいま貼り付けた図の位置を取得して、変更する
            sShapePositionLEFT = .Shapes(.Shapes.Count).Left
            sShapePositionTOP = .Shapes(.Shapes.Count).Top
        
            With .Shapes(.Shapes.Count) '図の位置を微調整
                .Left = sShapePositionLEFT - 4.5
                .Top = sShapePositionTOP + 2
            End With
        End With
        On Error GoTo 0
    Next i
    
Exit Sub

ErrHandler:
    If ERR.Number = 1004 Then GoTo TryAgain

End Sub
'---------------------------------------------------------------------------------------
' Method : BreakPages
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ヘッダー挿入後の改ページ設定
'---------------------------------------------------------------------------------------
Private Sub BreakPages(ByVal myPageCount As Long)
    Dim i As Long
    Dim Pages As Long
    Dim rng As Range
    Dim RowsPerPage As Long

    RowsPerPage = 60    '1ページの行数
    Pages = 0
    Application.PrintCommunication = False
    With ActiveSheet
        '現在の改ページをすべてリセット
        .ResetAllPageBreaks
        For i = 60 To (myPageCount * 60) Step RowsPerPage
            .HPageBreaks.Add before:=.Cells(i, 1)
        Next i
    End With
    Application.PrintCommunication = True
End Sub

Private Function 見積書項目の列番号を取得(ByRef myPosCol As Long, myTotalCol As Long) As Long
    Dim bkTemp As Workbook
    Dim wksQuotation As Worksheet
    Dim wksTemp As Worksheet
    Dim wksEval As Worksheet
    Dim myArray As Variant
    Dim myItemCol As Long
    
    Const bkName As String = "Temp"
    
    Set wksQuotation = ActiveWorkbook.ActiveSheet   'アクティブシート＝対象見積書とする
    Set bkTemp = Workbooks.Add                          '分析データを展開する一時シートをつくる
    Set wksTemp = bkTemp.Sheets(1)
    wksTemp.Name = bkName
    
    Set wksEval = データ件数抽出とランク付け(wksQuotation)
    If mbCancel Then Exit Function 'アクティブシートが見積書でなければ終了する
    myArray = GetExtractedCols(wksEval)  'データ件数の上位列を取り出す
    見積書項目の列番号を取得 = GetItemCol(wksTemp) '品名列をしらべる
    myPosCol = GetPosCol(wksTemp) 'Pos列をしらべる
    myTotalCol = GetTotalCol(wksQuotation, wksTemp)
    
    MsgBox "品名の列番号は" & 見積書項目の列番号を取得 & "です。"
    MsgBox "POS列は" & myPosCol & "です。"
    MsgBox "Total列は" & myTotalCol & "です。"
    
    Application.DisplayAlerts = False
    bkTemp.Close
    Set bkTemp = Nothing
    Application.DisplayAlerts = True
    
    Set wksQuotation = Nothing
    Set wksTemp = Nothing
    
End Function

Private Function データ件数抽出とランク付け(ByRef wks As Worksheet) As Worksheet
 '指定範囲のデータ件数をかぞえる
    
    Dim lCol As Long
    Dim myRow As Long
    Dim lEndCol As Long
    Dim lFirstRow As Long
    Dim lEndRow As Long
    Dim myCountA As Long
    Dim myPrintRange As Range
    Dim lCount As Long
    Const lStartRow As Long = 36
    
    '見積書の印刷範囲を取得
    On Error GoTo ErrHandler
    Set myPrintRange = wks.Range(wks.PageSetup.PrintArea)
    On Error GoTo 0
    lEndCol = myPrintRange.Item(myPrintRange.Count).column

    
    '一時シートに抽出結果を書き込む
    myRow = 1
    For lCol = 1 To lEndCol
        
        'データ件数をしらべる
        myCountA = WorksheetFunction.CountA(myPrintRange.Range(myPrintRange.Cells(lStartRow, _
                        lCol), myPrintRange.Cells(lEndRow, lCol)))

        If myCountA > 0 Then
            With Workbooks(Workbooks.Count).Worksheets(1)
                .Cells(1, 1).offset(myRow, 0).Value = lCol
                .Cells(1, 1).offset(myRow, 1).Value = myCountA
            End With
            myRow = myRow + 1
        End If
        
        '平均文字長さを調べる
        If myCountA > 0 Then Call 文字長さ平均を計算(lCol, wks)
    Next lCol
    
'   一時シートの書式をととのえる
    With Workbooks(Workbooks.Count).Worksheets(1)
        .Cells(1, 1).Value = "列番"
        .Cells(1, 2).Value = "データ件数"
        .Cells(1, 3).Value = "平均文字長さ"
        .Columns("A:C").EntireColumn.AutoFit
        .Range("A1").CurrentRegion.Sort _
            key1:=Range("b2"), order1:=xlDescending
    End With

    Set データ件数抽出とランク付け = Workbooks(Workbooks.Count).Worksheets(1)

Exit Function

ErrHandler:
MsgBox "無効なシートです"
mbCancel = True
End Function

Private Function GetExtractedCols(ByVal wks As Worksheet) As Variant
    Dim myArray() As Long
    Dim i As Long
    Dim myRow As Long
    
    Const myExtractingRows As Long = 5 '抽出する列数
    Const strCol As String = "B"            '？？？
    
    ReDim myArray(myExtractingRows) As Long
    
    With wks
        i = 0
        myRow = 2  '見出し行を除く
        For i = 0 To myExtractingRows
            myArray(i) = .Cells(myRow, strCol).Value
            myRow = myRow + 1
        Next i
    End With

    GetExtractedCols = myArray

End Function

Private Function GetItemCol(ByRef wks As Worksheet) As Long
    Dim i As Long
    Dim rngAveStringLen As Range
    Dim rng As Range
    Dim lMaxStringLen As Long
    
    Set rngAveStringLen = wks.Range(wks.Cells(2, 3), wks.Cells(2, 3).End(xlDown))
    For Each rng In rngAveStringLen
        If rng.Value > lMaxStringLen Then
            lMaxStringLen = rng.Value
            GetItemCol = rng.offset(0, -2) '平均文字長さが最長の列の番号を返す
        End If
    Next rng

    Set rngAveStringLen = Nothing
    Set rng = Nothing
    
End Function

Private Function GetPosCol(ByRef wks As Worksheet)
    Dim i As Long
    Dim rngColNum As Range
    Dim rng As Range
    Dim lMinColNum As Long
    Dim buf As Long
    
    Set rngColNum = wks.Range(wks.Cells(2, 1), wks.Cells(2, 1).End(xlDown))
    
    lMinColNum = 10
    For Each rng In rngColNum
        If rng.Value < lMinColNum Then
            lMinColNum = rng.Value
            GetPosCol = rng.Value '最小の列番号を返す
        End If
    Next rng
    
    Set rngColNum = Nothing
    Set rng = Nothing
    
End Function

Private Function GetTotalCol(ByRef wksQuotation As Worksheet, ByRef wksTemp As Worksheet)
    Dim i As Long
    Dim rngColNum As Range
    Dim rng As Range
    Dim lMaxColNum As Long
    Dim buf As Long
    Dim rngPrintArea As Range
    Dim myPrintAreaRightEdgeCol As Long
    
    '印刷範囲の右端の切れ目の列
    Set rngPrintArea = wksQuotation.Range(wksQuotation.PageSetup.PrintArea)
    myPrintAreaRightEdgeCol = rngPrintArea.Item(rngPrintArea.Count).column
    
    Set rngColNum = wksTemp.Range(wksTemp.Cells(2, 1), wksTemp.Cells(2, 1).End(xlDown))
    
    For Each rng In rngColNum
        If rng.Value > lMaxColNum And rng.Value < myPrintAreaRightEdgeCol Then
            lMaxColNum = rng.Value
            GetTotalCol = rng.Value '印刷範囲内で最大の列番号を返す。データ件数は考えない。
        End If
    Next rng
    
    Set rngPrintArea = Nothing
    Set rngColNum = Nothing
    Set rng = Nothing
    
End Function

Private Sub セルの表示形式をしらべる()
'表示形式をしらべる。単価と合計は、
'両方の表示形式が同じであることで判定できると思う。
    
    Dim i As Long
    Dim myFormat As String
    myFormat = ActiveCell.NumberFormat
    
End Sub
Private Sub 文字長さ平均を計算(ByVal lCol As Long, ByRef wksQuotation As Worksheet)
'CountAで割り出した列あたりのデータ数上位数件のうち、_
'もっとも平均文字長さが長い列が品名の列だ。
    Dim i As Long
    Dim lRow As Long
    Dim lEndRow As Long
    Dim lEndRowTemp As Long
    Dim myItemCntUsedRange As Long
    Dim rng As Range
    Dim sum As Long
    Dim myAverageLen As Double
    Dim myEndRow As Long
    
    '最終行
    Dim UsedRng
    Set UsedRng = wksQuotation.UsedRange
    lEndRow = UsedRng.UsedRange.Item(UsedRng.UsedRange.Count).Row
    myItemCntUsedRange = UsedRng.Count
    
    
'   注釈は計算に含めない：合計金額より以下の行のデータは無視する
    myEndRow = GetGrandTotalRow(wksQuotation)  '合計金額の行を取得
            
    For lRow = 39 To lEndRow
        If Len(wksQuotation.Cells(lRow, lCol).Value) > 0 And lRow < myEndRow Then  '文字列があって合計金額行より上の行のみ計算する
            sum = Len(wksQuotation.Cells(lRow, lCol)) + sum
            i = i + 1
        End If
    Next lRow
    
    If i = 0 Then Exit Sub '空の列なら次へ
    
    myAverageLen = sum / i  '文字長さの合計を件数でわる
    
    Dim a As Worksheet
    Set a = ActiveSheet
'   一時シートを参照
    lEndRowTemp = a.Cells(Rows.Count, 1).End(xlUp).Row + 1
    Set rng = a.Range(a.Cells(2, 1), a.Cells(lEndRowTemp, 1)).Find(lCol).offset(0, 2) '該当列の「平均文字長さ」の入力セルを取得
    rng.Value = myAverageLen
    
    Debug.Print lCol & "の平均文字長さ：" & myAverageLen
    Set rng = Nothing
    Set a = Nothing

End Sub

Private Function GetGrandTotalRow(ByRef wksQuotation As Worksheet) As Long
    Dim rngNumerics As Range
    Dim rng As Range
    Dim lMaxNum As Long
    
    With wksQuotation
        Set rngNumerics = .Cells(1, 1).SpecialCells(xlCellTypeFormulas, 1)
        For Each rng In rngNumerics
            '範囲内での最大値
            If rng.Value > lMaxNum Then
                GetGrandTotalRow = rng.Row '最大値がある行を返す
            End If
        Next rng
    End With
    
    Set rngNumerics = Nothing
    Set rng = Nothing

End Function

Private Sub CountMergeColumn()
    Dim i As Long
    Dim buf As String
    
    '品名以外はおそらくセルが結合されているのでラベルの位置をはっきりと特定できる
    With ActiveCell.MergeArea
        buf = buf & "セルの個数：" & .Columns.Count & vbCrLf
        buf = buf & "左端の列：" & .Item(1).column & vbCrLf
        buf = buf & "右端の列：" & .Item(.Count).column
    End With
    
    MsgBox buf
    
End Sub
