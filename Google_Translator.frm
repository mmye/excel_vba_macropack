VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Google_Translator 
   Caption         =   "Google Translate"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   OleObjectBlob   =   "Google_Translator.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Google_Translator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const Afrikaans = "af"
Const Irish = "ga"
Const Albanian = "sq"
Const Italian = "it"
Const Arabic = "ar"
Const Japanese = "ja"
Const Azerbaijani = "az"
Const Kannada = "kn"
Const Basque = "eu"
Const Korean = "ko"
Const Bengali = "bn"
Const Latin = "la"
Const Belarusian = "be"
Const Latvian = "lv"
Const Bulgarian = "bg"
Const Lithuanian = "lt"
Const Catalan = "ca"
Const Macedonian = "mk"
Const Chinese_Simplified = "zh-cn"
Const Malay = "ms"
Const Chinese_Traditional = "zh-TW"
Const Maltese = "mt"
Const Croatian = "hr"
Const Norwegian = "no"
Const Czech = "cs"
Const Persian = "fa"
Const Danish = "da"
Const Polish = "pl"
Const Dutch = "nl"
Const Portuguese = "pt"
Const English = "en"
Const Romanian = "ro"
Const Esperanto = "eo"
Const Russian = "ru"
Const Estonian = "et"
Const Serbian = "sr"
Const Filipino = "tl"
Const Slovak = "sk"
Const Finnish = "fi"
Const Slovenian = "sl"
Const French = "fr"
Const Spanish = "es"
Const Galician = "gl"
Const Swahili = "sw"
Const Georgian = "ka"
Const Swedish = "sv"
Const German = "de"
Const Tamil = "ta"
Const Greek = "el"
Const Telugu = "te"
Const Gujarati = "gu"
Const Thai = "th"
Const Haitian_Creole = "ht"
Const Turkish = "tr"
Const Hebrew = "iw"
Const Ukrainian = "uk"
Const Hindi = "hi"
Const Urdu = "ur"
Const Hungarian = "hu"
Const Vietnamese = "vi"
Const Icelandic = "is"
Const Welsh = "cy"
Const Indonesian = "id"
Const Yiddish = "yi"

Dim CtrlPresshed As Boolean
Dim objHTTP As Object
Dim Lists() As Variant
Dim v As Variant
Dim TranslationLanguages As String
Dim CtrlPressed As Boolean
Dim OriginalWordLists() As String
Dim TranslationLists() As String
Dim AddressLists() As String
Dim mbCancelProc As Boolean
Dim mbCancelEvent As Boolean

Private Function Translate(str As String, translateFrom As String, translateTo As String) As String
    Dim getParam As String, trans As String, url As String
    If LenB(str) = 0 Then Exit Function
    getParam = ConvertToGet(str)
    url = "https://translate.google.pl/m?hl=" & translateFrom & "&sl=" & translateFrom & "&tl=" & translateTo & "&ie=UTF-8&prev=_m&q=" & getParam
    objHTTP.Open "GET", url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    On Error GoTo ConnectionFailed
    objHTTP.send ("")
    On Error GoTo 0
    If InStr(objHTTP.responseText, "div dir=""ltr""") > 0 Then
        trans = RegularExpressions.RegexExecuteGet(objHTTP.responseText, "div[^""]*?""ltr"".*?>(.+?)</div>", 0, 0)
        Translate = Clean(trans)
    Else
        ERR.Raise 0, "Translate", "No connection or other error"
    End If
    Exit Function
ConnectionFailed:
    MsgBox "インターネット接続がないようです。接続を確認してください。"
    Exit Function
End Function
Private Function ConvertToGet(Val As String)
    Val = Replace(Val, " ", "+")
    Val = Replace(Val, vbNewLine, "+")
    Val = Replace(Val, "(", "%28")
    Val = Replace(Val, ")", "%29")
    ConvertToGet = Val
End Function
Private Function Clean(Val As String)
    Val = Replace(Val, "&quot;", """")
    Val = Replace(Val, "%2C", ",")
    Val = Replace(Val, "&#39;", "'")
    Clean = Val
End Function


Private Sub btnNext_Click()
    Dim SelItem As Long
    If lbWordList.ListCount = 0 Then Exit Sub
    SelItem = GetSelIndex
    If SelItem = -1 Then Exit Sub
    If SelItem = lbWordList.ListCount - 1 Then Exit Sub
    lbWordList.ListIndex = lbWordList.ListIndex + 1
End Sub

Private Sub btnPrevious_Click()
    Dim SelItem As Long
    SelItem = GetSelIndex
    If SelItem = -1 Then Exit Sub
    If SelItem = 0 Then Exit Sub
    lbWordList.ListIndex = lbWordList.ListIndex - 1
End Sub

Private Sub btnWriteToSheet_Click()
    Dim i As Long
    Dim wks As Worksheet
    Dim Address As String
    If lbWordList.ListCount = 0 Then Exit Sub
    Call ScreenUpdatingSwitch
    Set wks = ActiveWorkbook.Sheets(cmbWks.Text)
    wks.Select
    For i = 0 To lbWordList.ListCount - 1
        Address = AddressLists(i)
        wks.Range(Address).Value = lbWordList.List(i, 2)
    Next
   Set wks = Nothing
   Call ScreenUpdatingSwitch
   MsgBox "訳文をシートに入力しました。"
End Sub



Private Sub CommandButton1_Click()
    Dim SelIndex   As Long
    SelIndex = GetSelIndex
    lbWordList.List(SelIndex, 2) = txtTranslation.Text
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub lbWordList_Click()
    Dim SelIndex   As Long
    SelIndex = GetSelIndex
    Sheets(cmbWks.Text).Select
    If SelIndex = -1 Then Exit Sub
    lbOriginalBody.Caption = lbWordList.List(SelIndex, 1)
    txtTranslation.Text = lbWordList.List(SelIndex, 2)
    If Len(AddressLists(SelIndex)) > 0 Then _
    ActiveSheet.Range(AddressLists(SelIndex)).Select
    
End Sub

Sub UserForm_Initialize()
    Dim wks As Worksheet
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    For Each wks In ActiveWorkbook.Sheets
        cmbWks.AddItem wks.Name
    Next
    Set wks = Nothing
End Sub

Private Sub btnExecTranslation_Click()
    Dim Lists As Variant
    Dim ItemCount As Long
    Dim rng As Range
    Dim wksName As String
    Dim frmProgress As FProgressBar
    
    wksName = cmbWks.Text
    If wksName = "" Then
        MsgBox "シートを選択してください", vbInformation
        Exit Sub
    End If
    
    Set rng = GetTargetRange(wksName)
    If mbCancelEvent Then Exit Sub
    ItemCount = GetItemCount(rng)
    If ItemCount = 0 Then
        MsgBox "シートが空です。", vbCritical
        Exit Sub
    End If
        
'   言語選択チェック
    Dim LanguageSelected As Boolean
    LanguageSelected = CheckLanguageSelect
    If LanguageSelected = False Then
        MsgBox "言語を選択してください", vbInformation
        Exit Sub
    End If
    
    SetLaunguages
    InitializeProgressBar frmProgress, ItemCount
    Lists = GetWordList(rng, ItemCount, frmProgress)
    If IsArrayEx(Lists) <> -1 Then lbSegmentCount = UBound(Lists) + 1 Else: Exit Sub
    ThrowDataIntoListbox Lists
    lbWordList.List = Lists
    lbWordList.ListIndex = 0
    TerminateProgressBar frmProgress
End Sub
Private Function CheckLanguageSelect() As Boolean
    Dim IsSelected As Boolean
    If tglDJ.Value Then IsSelected = True
    If tglEJ.Value Then IsSelected = True
    If tglJE.Value Then IsSelected = True
    If IsSelected = False Then CheckLanguageSelect = False Else: CheckLanguageSelect = True
End Function
Private Function GetTargetRange(wksName) As Range
    Dim rConstants As Range, rFormulas As Range
    On Error GoTo ERR
    Set rConstants = ActiveWorkbook.Sheets(wksName).UsedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    On Error Resume Next
    Set rFormulas = ActiveWorkbook.Sheets(wksName).UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rFormulas Is Nothing Then
        Set GetTargetRange = rConstants
    Else
        Set GetTargetRange = Union(rConstants, rFormulas)
    End If
Exit Function
ERR:
    If ERR.Number = 1004 Then MsgBox "シートに文字が一つもありません"
    mbCancelEvent = True
End Function
Private Sub SetLaunguages()
    If tglJE.Value Then TranslationLanguages = "JE"
    If tglDJ.Value Then TranslationLanguages = "DJ"
    If tglEJ.Value Then TranslationLanguages = "EJ"
End Sub

Sub InitializeProgressBar(frmProgress, ItemCount)
    'Initialise the progress bar
    Set frmProgress = New FProgressBar
    frmProgress.title = "機械翻訳ツール"
    frmProgress.Text = "翻訳を取得しています…"
    frmProgress.Min = 1
    frmProgress.max = ItemCount

    'Show the progress bar
    frmProgress.ShowForm

End Sub

Private Sub TerminateProgressBar(frmProgress As FProgressBar)
    Unload frmProgress
End Sub
Private Sub UpdateProgressBar(lLoop, frmProgress)
    'Check if the user cancelled
    If frmProgress.Cancelled Then
        mbCancelProc = True
        Exit Sub
    End If
    'Update the progress
    frmProgress.progress = lLoop
End Sub

Private Function GetWordList(TargetRange, ItemCount, frmProgress As FProgressBar) As Variant
    Dim wks As Worksheet, r As Range
    Dim c As Long
    Dim buf As String
    Dim IndexLists() As Long

    ReDim IndexLists(ItemCount - 1) As Long
    ReDim OriginalWordLists(ItemCount - 1) As String
    ReDim TranslationLists(ItemCount - 1) As String
    ReDim AddressLists(ItemCount - 1) As String

    For Each r In TargetRange
        buf = r.Value
        If LenB(buf) > 0 Then
        If Not IsNumeric(buf) Then
        If Not IsError(buf) Then
            UpdateProgressBar c, frmProgress
            IndexLists(c) = c + 1
            OriginalWordLists(c) = buf
            Select Case TranslationLanguages
                Case "JE"
                    TranslationLists(c) = Translate(Trim$(buf), Translations.Japanese, Translations.English)
                Case "EJ"
                    TranslationLists(c) = Translate(Trim$(buf), Translations.English, Translations.Japanese)
                Case "DJ"
                    TranslationLists(c) = Translate(Trim$(buf), Translations.German, Translations.Japanese)
            End Select
            AddressLists(c) = r.Address
            c = c + 1
        End If
        End If
        End If
    Next

    Dim joinedLists() As Variant
    joinedLists = JoinLists(IndexLists, OriginalWordLists, TranslationLists)
    GetWordList = joinedLists
    Set r = Nothing

End Function

Private Function JoinLists(Index, Originals, Translations) As Variant
    Dim i As Long, c As Long
    Dim Lists() As Variant
    ReDim Lists(UBound(Index), 2) As Variant

    For i = 0 To UBound(Index)
        Lists(c, 0) = Index(c)
        Lists(c, 1) = Originals(c)
        Lists(c, 2) = Translations(c)
        c = c + 1
    Next i
    JoinLists = Lists
End Function

Private Sub ThrowDataIntoListbox(Lists)
    lbSegmentCount = UBound(Lists) + 1
    lbWordList.List = Lists
    lbWordList.ListIndex = 0
End Sub

Private Sub btnQuit_Click()
    Unload Me
End Sub
Private Function GetSelIndex() As Long
    Dim i
    For i = 0 To lbWordList.ListCount - 1
        If lbWordList.Selected(i) Then
            GetSelIndex = i
            Exit Function
        End If
    Next
End Function
Private Function GetItemCount(rng) As Long
    Dim c As Long
    Dim r As Range
    Dim buf As String
    For Each r In rng
        buf = r.Value
        If Not IsError(buf) Then
            If LenB(buf) > 0 Then
                If Not IsNumeric(buf) Then
                    c = c + 1
                End If
            End If
        End If
    Next
    GetItemCount = c
End Function

Private Sub ScreenUpdatingSwitch()
    Application.ScreenUpdating = Not Application.ScreenUpdating
End Sub
'***********************************************************
' 機能   : 引数が配列か判定し、配列の場合は空かどうかも判定する
' 引数   : varArray  配列
' 戻り値 : 判定結果（1:配列/0:空の配列/-1:配列じゃない）
'***********************************************************
Private Function IsArrayEx(varArray As Variant) As Long
    On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_:
    If ERR.Number = 9 Then
        IsArrayEx = 0
    End If
End Function

Private Sub cmbWks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmbWks.DropDown
End Sub

Private Sub txtTranslation_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim SelIndex As Long
    SelIndex = GetSelIndex
    If KeyCode = 13 Then lbWordList.List(SelIndex, 2) = txtTranslation.Text
End Sub
Private Sub txtTranslation_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 17 Then CtrlPressed = False
    Select Case KeyCode
        Case 78 'N
            If CtrlPressed Then btnNext_Click
            Exit Sub
        Case 80 'P
            If CtrlPressed Then btnPrevious_Click
            Exit Sub
    End Select
End Sub
Private Sub txtOriginalText_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 17 Then CtrlPresshed = False
    Select Case KeyCode
        Case 78 'N
            If CtrlPresshed Then btnNext_Click
            Exit Sub
        Case 80 'P
            If CtrlPresshed Then btnPrevious_Click
            Exit Sub
    End Select
End Sub
Private Sub tglDJ_Click()
    If mbCancelEvent Then Exit Sub
    mbCancelEvent = True
    If tglDJ.Value Then tglEJ.Value = False
    If tglDJ.Value Then tglJE.Value = False
    mbCancelEvent = False
End Sub

Private Sub tglEJ_Click()
    If mbCancelEvent Then Exit Sub
    mbCancelEvent = True
    If tglEJ.Value Then tglDJ.Value = False
    If tglEJ.Value Then tglJE.Value = False
    mbCancelEvent = False
End Sub

Private Sub tglJE_Click()
    If mbCancelEvent Then Exit Sub
    mbCancelEvent = True
    If tglJE.Value Then tglDJ.Value = False
    If tglJE.Value Then tglEJ.Value = False
    mbCancelEvent = False
End Sub

