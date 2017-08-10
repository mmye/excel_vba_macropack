Attribute VB_Name = "選択範囲内をGoogle翻訳"
Option Explicit

'2017/01/26
'このモジュールを使用するには、TranslationsおよびRegularExpressionsモジュールが必要です。

Sub 選択範囲内をGoogle翻訳()

    If Selection.Count = 0 Then Exit Sub

    Dim Sel As Range
    Set Sel = Selection

    Dim SourceLang As String
    SourceLang = InputBox("原文の言語を指定してください...(日本語：ja  英語:en)")
    Select Case SourceLang
        Case "ja"
            SourceLang = "ja"
        Case "en"
            SourceLang = "en"
        Case ""
            Exit Sub
        Case Else
            '何もしない
    End Select

    Dim TargetLang As String
    TargetLang = InputBox("訳文の言語を指定してください...(日本語：ja  英語:en)")
    Select Case TargetLang
        Case "ja"
            TargetLang = "ja"
        Case "en"
            TargetLang = "en"
        Case ""
            Exit Sub
        Case Else
            '何もしない
    End Select

    Dim r As Range
    For Each r In Sel.Cells
        If Not IsError(r) Or IsNumeric(r) Then
            Dim buf As String
            buf = r.Value
            If Len(buf) > 0 Then
                Dim Translation As String
                Translation = GetTranslations(buf, SourceLang, TargetLang)
                r.Value = Translation
            End If
        End If
    Next r

    Set Sel = Nothing
    MsgBox "選択範囲内をGoogle翻訳しました"
End Sub

Private Function GetTranslations(buf As String, SourceLang As String, TargetLang As String)
    Dim msg As String
    msg = "言語は日本語または英語から選択してください"
    
    Select Case SourceLang
        Case "ja"
            If TargetLang = "en" Then
                GetTranslations = Translate(buf, Translations.Japanese, Translations.English)
            Else
                MsgBox (msg)
                End
            End If
        Case "en"
            If TargetLang = "ja" Then
                GetTranslations = Translate(buf, Translations.English, Translations.Japanese)
            Else
                MsgBox (msg)
                End
            End If
    End Select

End Function

