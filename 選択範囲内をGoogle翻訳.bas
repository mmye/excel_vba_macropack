Attribute VB_Name = "�I��͈͓���Google�|��"
Option Explicit

'2017/01/26
'���̃��W���[�����g�p����ɂ́ATranslations�����RegularExpressions���W���[�����K�v�ł��B

Sub �I��͈͓���Google�|��()

    If Selection.Count = 0 Then Exit Sub

    Dim Sel As Range
    Set Sel = Selection

    Dim SourceLang As String
    SourceLang = InputBox("�����̌�����w�肵�Ă�������...(���{��Fja  �p��:en)")
    Select Case SourceLang
        Case "ja"
            SourceLang = "ja"
        Case "en"
            SourceLang = "en"
        Case ""
            Exit Sub
        Case Else
            '�������Ȃ�
    End Select

    Dim TargetLang As String
    TargetLang = InputBox("�󕶂̌�����w�肵�Ă�������...(���{��Fja  �p��:en)")
    Select Case TargetLang
        Case "ja"
            TargetLang = "ja"
        Case "en"
            TargetLang = "en"
        Case ""
            Exit Sub
        Case Else
            '�������Ȃ�
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
    MsgBox "�I��͈͓���Google�|�󂵂܂���"
End Sub

Private Function GetTranslations(buf As String, SourceLang As String, TargetLang As String)
    Dim msg As String
    msg = "����͓��{��܂��͉p�ꂩ��I�����Ă�������"
    
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

