Attribute VB_Name = "Translations"
Public Const Afrikaans = "af"
Public Const Irish = "ga"
Public Const Albanian = "sq"
Public Const Italian = "it"
Public Const Arabic = "ar"
Public Const Japanese = "ja"
Public Const Azerbaijani = "az"
Public Const Kannada = "kn"
Public Const Basque = "eu"
Public Const Korean = "ko"
Public Const Bengali = "bn"
Public Const Latin = "la"
Public Const Belarusian = "be"
Public Const Latvian = "lv"
Public Const Bulgarian = "bg"
Public Const Lithuanian = "lt"
Public Const Catalan = "ca"
Public Const Macedonian = "mk"
Public Const Chinese_Simplified = "zh-cn"
Public Const Malay = "ms"
Public Const Chinese_Traditional = "zh-TW"
Public Const Maltese = "mt"
Public Const Croatian = "hr"
Public Const Norwegian = "no"
Public Const Czech = "cs"
Public Const Persian = "fa"
Public Const Danish = "da"
Public Const Polish = "pl"
Public Const Dutch = "nl"
Public Const Portuguese = "pt"
Public Const English = "en"
Public Const Romanian = "ro"
Public Const Esperanto = "eo"
Public Const Russian = "ru"
Public Const Estonian = "et"
Public Const Serbian = "sr"
Public Const Filipino = "tl"
Public Const Slovak = "sk"
Public Const Finnish = "fi"
Public Const Slovenian = "sl"
Public Const French = "fr"
Public Const Spanish = "es"
Public Const Galician = "gl"
Public Const Swahili = "sw"
Public Const Georgian = "ka"
Public Const Swedish = "sv"
Public Const German = "de"
Public Const Tamil = "ta"
Public Const Greek = "el"
Public Const Telugu = "te"
Public Const Gujarati = "gu"
Public Const Thai = "th"
Public Const Haitian_Creole = "ht"
Public Const Turkish = "tr"
Public Const Hebrew = "iw"
Public Const Ukrainian = "uk"
Public Const Hindi = "hi"
Public Const Urdu = "ur"
Public Const Hungarian = "hu"
Public Const Vietnamese = "vi"
Public Const Icelandic = "is"
Public Const Welsh = "cy"
Public Const Indonesian = "id"
Public Const Yiddish = "yi"
Public Function Translate(str As String, translateFrom As String, translateTo As String) As String
    Dim getParam As String, trans As String
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
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
    Val = Replace(Val, vbLf, "+") 'Add
    Val = Replace(Val, "(", "%28")
    Val = Replace(Val, ")", "%29")
    ConvertToGet = Val
End Function
Private Function Clean(Val As String)
    Val = Replace(Val, "&quot;", """")
    Val = Replace(Val, "%2C", ",")
    Val = Replace(Val, "&#39;", "'")
    Val = Replace(Val, "&gt;", ">") 'Add
    Val = Replace(Val, "&lt;", "<") 'Add
    Val = Replace(Val, "+", vblr) 'Add
    Clean = Val
End Function


