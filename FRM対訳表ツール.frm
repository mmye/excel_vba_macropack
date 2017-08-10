VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM対訳表ツール 
   Caption         =   "対訳表ツール"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5010
   OleObjectBlob   =   "FRM対訳表ツール.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRM対訳表ツール"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    
    If cmbOriginalWkb.Text = cmbTranslatedWkb Then
        MsgBox "原文と訳文は別のブックを選択してください", vbOKOnly, "お知らせ"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Call 対訳表作成
End Sub

Private Sub CommandButton2_Click()
 Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim wkb As Workbook
    
    With cmbOriginalWkb
        For Each wkb In Workbooks
            .AddItem wkb.Name
        Next wkb
    End With
    
    With cmbTranslatedWkb
        For Each wkb In Workbooks
            .AddItem wkb.Name
        Next wkb
    End With
End Sub

Private Sub 対訳表作成()
    Dim wkbOriginal As Workbook
    Dim wkbTranslation As Workbook
    Dim wkbNew As Workbook
    Dim wksCompareTBL As Worksheet
    Dim buf As String
    Dim r As Long
    Dim c As Long
    Dim lEndRow As Long
    Dim lEndCol As Long
    Dim wks As Worksheet
    Dim wksIndex As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim i As Long
    Dim sOriginalWkb As String
    Dim sTranslatedWkb As String
    Const mySeparator As String = ""
    Const GrayFontColor As Long = 5855577
    
    Application.ScreenUpdating = False
    sOriginalWkb = cmbOriginalWkb.Text
    sTranslatedWkb = cmbTranslatedWkb.Text

    Set wkbOriginal = Workbooks(sOriginalWkb)
    Set wkbTranslation = Workbooks(sTranslatedWkb)
    
    Set wkbNew = Workbooks.Add
    
    For Each wks In wkbOriginal.Sheets
        'If i > 1 Then Exit Sub
        wksIndex = wksIndex + 1
        Application.StatusBar = wksIndex & ":" & wks.Name & "の対訳を作成しています..."
        If wksIndex > 1 Then
            
            'シートが無い場合は飛ばす
            On Error GoTo ERR
            wkbTranslation.Sheets(wksIndex).Copy After:=wkbNew.Sheets(wkbNew.Sheets.Count)
            Set wksCompareTBL = wkbNew.Sheets(wkbNew.Sheets.Count)
            
            On Error Resume Next '既に使用されているシート名を避ける
            wksCompareTBL.Name = wkbTranslation.Sheets(wksIndex).Name
            If ERR.Number <> 0 Then
                wksCompareTBL.Name = "foo_" & wkbTranslation.Sheets(wksIndex).Name
            End If
            ERR.Clear
            On Error GoTo 0
        
        Else
            wkbTranslation.Sheets(wksIndex).Copy After:=wkbNew.Sheets(wkbNew.Sheets.Count)
            Set wksCompareTBL = wkbNew.Sheets(wkbNew.Sheets.Count)
            wksCompareTBL.Name = wkbTranslation.Sheets(wksIndex).Name
        End If
        
        lEndRow = wks.UsedRange.Item(wks.UsedRange.Count).Row
        lEndCol = wks.UsedRange.Item(wks.UsedRange.Count).column
        For r = 1 To lEndRow
            For c = 1 To lEndCol
                With wksCompareTBL
                DoEvents
                    With .Cells(r, c)
                        If .Value <> "" Then
                            .Font.Name = "Arial" '文字列のフォントをそろえる
                            '原文同一判定：JAブックの原文は追記しない
                            If .Value <> wkbOriginal.Sheets(wksIndex).Cells(r, c).Value Then
                                lStart = Len(.Value) + Len(mySeparator) + 1
                                buf = wkbOriginal.Sheets(wksIndex).Cells(r, c).Value
                                .Value = .Value & vbCrLf & mySeparator & buf
                                lEnd = Len(.Value)
                                With .Characters(lStart, lEnd).Font
                                    If wksCompareTBL.Cells(r, c).Interior.color <> vbWhite Then
                                        If wks.Cells(r, c).Interior.color = 13434828 Then
                                            .color = GrayFontColor
                                        Else
                                            .color = GrayFontColor ' vbWhite
                                        End If
                                    Else
                                        .color = GrayFontColor '灰色
                                    End If
                                    .Name = "ＭＳ Ｐゴシック"
                                    .Size = 0.9 * wkbOriginal.Sheets(wksIndex).Cells(r, c).Font.Size
                                End With
                            End If
        '                    i = i + 1
                        End If
                    End With
                End With
            Next c
        Next r
    Next wks
    
    '余分なシートを削除して選択をA1にする
    With wkbNew.Sheets(1)
        Application.DisplayAlerts = False
        .Delete
        Application.DisplayAlerts = True
        wkbNew.Sheets(1).Activate
        Range("a1").Select
    End With
    
    Application.ScreenUpdating = True
    
    Set wkbNew = Nothing
    Set wkbOriginal = Nothing
    Set wkbTranslation = Nothing
    Application.StatusBar = False
    MsgBox "完了しました", vbOKOnly + vbInformation, "お知らせ"
    
    Exit Sub

ERR:
    'wksIndexが不正な値を示す場合のエラー処理（存在しないシートインデックスの参照）
    If ERR.Number = 9 Then Resume Next
End Sub

Sub WksNumberring()
    Dim c As Long
    Dim wks As Worksheet
    
    c = 0
    For Each wks In ActiveWorkbook.Sheets

        wks.Name = c & "_" & wks.Name
            c = c + 1
    Next wks


End Sub
