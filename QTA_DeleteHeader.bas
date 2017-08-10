Attribute VB_Name = "QTA_DeleteHeader"
Option Explicit
Dim mbExitProc As Boolean
Dim mbCancelEvent As Boolean

'---------------------------------------------------------------------------------------
' Method : GetLastRow
' Author : mokoo
' Date   : 2016/02/20
' Purpose: シートの最下行を取得する
'---------------------------------------------------------------------------------------
Private Function GetLastRow(ByVal st As Worksheet) As Long
    Dim rngPrintArea As Range
    
    With st
        On Error GoTo ErrHandler
        Set rngPrintArea = .Range(.PageSetup.PrintArea)
        On Error GoTo 0
        GetLastRow = rngPrintArea.Item(rngPrintArea.Count).Row
    End With
 Exit Function
 
ErrHandler:
 
 mbCancelEvent = True
 
 End Function
 
 Private Function GetLastCol(ByVal st As Worksheet) As Long
    Dim rngPrintArea As Range
    
    With st
        Set rngPrintArea = .Range(.PageSetup.PrintArea)
        GetLastCol = rngPrintArea.Item(rngPrintArea.Count).column
    End With
 
 End Function
'---------------------------------------------------------------------------------------
' Method : CountTotalPages
' Author : temporary3
' Date   : 2016/04/07
' Purpose: アクティブシートの見積書としてのページ数をカウントする（60行/ページで計算）
'---------------------------------------------------------------------------------------
Private Function GetPageCount() As Long
    Dim lPageCount As Long
    Dim rowCount As Long
    Dim lPageMargin As Long
    
    With ActiveWorkbook.ActiveSheet
        
        rowCount = .UsedRange.Rows.Count

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
    End With
    
End Function
'---------------------------------------------------------------------------------------
' Method : DeleteHeader
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ヘッダーを削除する
'---------------------------------------------------------------------------------------
Sub DeleteHeader()
'*TODO:ヘッダーの削除残しが起きることがある

    Dim i As Long    'カウンタ1
    Dim f As Long    'カウンタ2
    Dim Pages As Long    'ページ数
    Dim NumPage As Long
    Dim lLastRow As Long
    
    'Application.ScreenUpdating = False
    
    lLastRow = GetLastRow(ActiveWorkbook.ActiveSheet)    'シートのデータがある最終行を取得
    NumPage = GetPageCount

    Call EraseBorders(NumPage)
    
    If mbExitProc Then
        mbExitProc = False
        MsgBox "キャンセルしました", vbOKOnly + vbInformation, "キャンセル"
        Exit Sub
    Else
        Call DeleteLogo(NumPage)
    End If
    
    'MsgBox lCountDeletedHeader & "ページ分のヘッダーを削除しました。", vbOKOnly Or vbInformation, "ヘッダー削除完了"
    
Exit Sub

NoHeadertoDelete:
    MsgBox "削除するヘッダーがありません。この見積書は1ページまでしかないようです。", vbOKOnly Or vbInformation, "ヘッダーなし"

End Sub

Private Sub EraseBorders(ByVal NumPage As Long)

    Dim i As Long
    Dim j As Long
    Dim b1 As Border    '第1罫線
    Dim b2 As Border    '第2罫線
    Dim CurrentActivecell   As Range
    Dim lRowTop As Long
    Dim lRowBottm As Long
    Dim rngToDelete As Range
    Dim lCountDeletedHeader As Long
    Dim myYesNo As VbMsgBoxResult
    Dim lHeaderCount As Long
    Dim sUserMessage As String
    
    '罫線がBoldの場合の処理
    For i = 60 To NumPage * 60    '最終セルの番号から末尾を決める

        '罫線の参照
        With Cells(i, "C")
            Set b1 = .Borders(xlEdgeBottom)
            Set b2 = .offset(1, 0).Borders(xlEdgeBottom)
        End With
        
        If b1.LineStyle = xlContinuous And _
           b1.Weight = xlThick And _
           b2.LineStyle = xlContinuous Then

            'ヘッダーの上端行をハードコーディング。
            lRowTop = i - 4
            
            'ヘッダーの下端行をさがす。jはMedium太さ罫線との下方向差分
            For j = 1 To 10
                If Cells(i, "C").offset(j, 0).Borders(xlEdgeBottom).LineStyle = xlContinuous Then lRowBottm = i + j
            Next j
            
            If rngToDelete Is Nothing Then
                Set rngToDelete = Rows(lRowTop & ":" & lRowBottm)
                lHeaderCount = lHeaderCount + 1
            Else
                Set rngToDelete = Union(rngToDelete, Rows(lRowTop & ":" & lRowBottm))
                lHeaderCount = lHeaderCount + 1
            End If
            
            lCountDeletedHeader = lCountDeletedHeader + 1
        End If
    Next i

    '罫線がMediumの場合の処理
    For i = 60 To (NumPage * 60)    '最終セルの番号から末尾を決める

        With Cells(i, "C")
            Set b1 = .Borders(xlEdgeBottom)
            Set b2 = .offset(1, 0).Borders(xlEdgeBottom)
        End With

        '会社名の下のMedium太さ罫線をさがす
        If b1.LineStyle = xlContinuous And _
           b1.Weight = xlMedium And _
           b2.LineStyle = xlContinuous Then

            'ヘッダーの上端行をハードコーディング。
            lRowTop = i - 4
            
            'ヘッダーの下端行をさがす。jはMedium太さ罫線との下方向差分
            For j = 1 To 10
                If Cells(i, "C").offset(j, 0).Borders(xlEdgeBottom).LineStyle = xlContinuous Then lRowBottm = i + j
            Next j
            
            If rngToDelete Is Nothing Then
                Set rngToDelete = Rows(lRowTop & ":" & lRowBottm)
                lHeaderCount = lHeaderCount + 1
            Else
                Set rngToDelete = Union(rngToDelete, Rows(lRowTop & ":" & lRowBottm))
                lHeaderCount = lHeaderCount + 1
            End If
            
        End If
    Next i
    
    If TypeName(Selection) = "Range" Then Set CurrentActivecell = ActiveCell
    
    If rngToDelete Is Nothing Then
        MsgBox "ヘッダーがみつかりませんでした", vbOKOnly + vbInformation, "ヘッダーなし"
        Exit Sub
    Else
    
        '取得したヘッダーの範囲を削除してよいかユーザーに確認する
        rngToDelete.Select
        sUserMessage = "ヘッダーが" & lHeaderCount & "個みつかりました。現在選択されている行を削除してもよろしいですか？"
        myYesNo = MsgBox(sUserMessage, vbYesNo + vbQuestion, "削除範囲の確認")
        
        If myYesNo = vbYes Then
            On Error GoTo ErrHandler
            rngToDelete.Delete
            Cells(1, 1).Select
            MsgBox "ヘッダーの削除が完了しました", vbOKOnly + vbInformation, "削除完了"
        Else
            If Not CurrentActivecell Is Nothing Then CurrentActivecell.Select 'アクティブセルを初期の状態に戻す
            mbExitProc = True
            Exit Sub
        End If
    End If

Exit Sub

ErrHandler:
MsgBox "削除中にエラーが起きました", vbOKOnly + vbInformation

End Sub

'---------------------------------------------------------------------------------------
' Method : DeleteLogo
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ヘッダーに含まれるロゴ画像を削除する
'---------------------------------------------------------------------------------------
Private Sub DeleteLogo(ByVal NumPage As Long)

    Dim shp As Shape
    Dim i As Long
    Dim rng_shp As Range
    Dim rng As Range

    'ロゴを消す
    For i = 37 To NumPage * 60 Step 2   '最初から最後まで

        Set rng = Range(Cells(i, 1), Cells(i, 20))  'ロゴがありそうな列

        For Each shp In ActiveSheet.Shapes
            Set rng_shp = Range(shp.TopLeftCell, shp.BottomRightCell)

            If Not (Intersect(rng_shp, rng) Is Nothing) Then    '検索するRangeとShapeが重なったら以下の処理。
                shp.Delete
            End If
        Next
    Next i

    'テキストボックス（社名入り）を消す
    For i = 37 To NumPage * 60 Step 2   '最初から最後まで

        Set rng = Cells(i, 16)    'テキストボックスがありそうな列

        For Each shp In ActiveSheet.Shapes
        
            Set rng_shp = Range(shp.TopLeftCell, shp.BottomRightCell)
            If Not (Intersect(rng_shp, rng) Is Nothing) Then
                shp.Delete
            End If
        Next shp
    Next i
    
    Set rng_shp = Nothing
    Set rng = Nothing
    
End Sub

