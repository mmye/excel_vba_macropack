Attribute VB_Name = "PDFWord変換"
Option Explicit
 
Private Enum Conv
    TypeDoc = 0
    TypeDocx = 1
    TypeEps = 2
    TypeHtml = 3
    TypeJpeg = 4
    TypeJpf = 5
    TypePdfA = 6
    TypePdfE = 7
    TypePdfX = 8
    TypePng = 9
    TypePs = 10
    TypeRft = 11
    TypeTiff = 12
    TypeTxtA = 13
    TypeTxtP = 14
    TypeXlsx = 15
    TypeSpreadsheet = 16
    TypeXml = 17
End Enum

Public Sub start()
    Dim myFile As String
    Dim ConvertedDoc As document
    
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show Then myFile = .SelectedItems(1)
    End With
    If myFile = "" Then Exit Sub
    Set ConvertedDoc = Convert2PDF(myFile, TypeDocx)
    ScreenUpdatingSwitch
    ConvertToNarrowChar ConvertedDoc
    フォント種類を設定する ConvertedDoc, "ＭＳ Ｐ明朝", "Arial"
    行頭行末の空白文字を削除 ConvertedDoc
    ScreenUpdatingSwitch
    Set ConvertedDoc = Nothing
End Sub
 
Private Function Convert2PDF(ByVal TargetFilePath As String, _
                       ByVal TargetConvType As Conv) As Object
'   PDFを他のファイル形式に変換
    Dim jso As Object
    Dim convid As String
    Dim ext As String
    Dim fp As String, fn As String, File As String
     
'   フォルダパスとファイル名取得
    With CreateObject("Scripting.FileSystemObject")
        fp = AddPathSeparator(.GetParentFolderName(TargetFilePath))
        fn = .GetBaseName(TargetFilePath)
    End With
    convid = GetConvID(TargetConvType)
    ext = GetExtension(TargetConvType)
    File = fp & fn & "." & ext
    With CreateObject("AcroExch.PDDoc")
        If .Open(TargetFilePath) = True Then
          Set jso = .GetJSObject
          CallByName jso, "saveAs", VbMethod, _
                     File, convid
          .Close
        End If
    End With
    Set Convert2PDF = Documents.Open(File)

End Function
 
Private Function GetConvID(ByVal ConvType As Conv) As String
'cConvID取得
  Dim v As Variant
   
  v = Array("com.adobe.acrobat.doc", "com.adobe.acrobat.docx", "com.adobe.acrobat.eps", _
            "com.adobe.acrobat.html", "com.adobe.acrobat.jpeg", "com.adobe.acrobat.jp2k", _
            "com.callas.preflight.pdfa", "com.callas.preflight.pdfe", "com.callas.preflight.pdfx", _
            "com.adobe.acrobat.png", "com.adobe.acrobat.ps", "com.adobe.acrobat.rtf", _
            "com.adobe.acrobat.tiff", "com.adobe.acrobat.accesstext", "com.adobe.acrobat.plain-text", _
            "com.adobe.acrobat.xlsx", "com.adobe.acrobat.spreadsheet", "com.adobe.acrobat.xml-1-00")
  GetConvID = v(ConvType)
End Function
 
Private Function GetExtension(ByVal ConvType As Conv) As String
'拡張子取得
  Dim v As Variant
   
  v = Array("doc", "docx", "eps", "html", "jpeg", "jpf", "pdf", "pdf", "pdf", "png", _
            "ps", "rft", "tiff", "txt", "txt", "xlsx", "xml", "xml")
  GetExtension = v(ConvType)
End Function
 
Private Function AddPathSeparator(ByVal s As String)
  If Right(s, 1) <> ChrW(92) Then s = s & ChrW(92)
  AddPathSeparator = s
End Function

Private Sub ConvertToNarrowChar(Doc)
    '全角英数字を半角英数字へ一括変換
    Dim rng As Range
    Set rng = Doc.Range(0, 0)
    With rng.Find
        .Text = "[０-９Ａ-Ｚａ-ｚ]{1,}"  '対象の設定
        .MatchWildcards = True
        Do While .Execute = True
            rng.Collapse wdCollapseEnd
        Loop
    End With
    Set rng = Nothing
End Sub

Private Sub 行頭行末の空白文字を削除(Doc)
    Dim rng As Range
    Dim vWhat As Variant
    Dim vReplace As Variant
    Dim i As Long

    '置換語句をここにベタ打ち。連続する半角・全角スペース・タブを一個にする
    vWhat = Array("^13([ 　^t]{1,})", "([ 　^t]{1,})^13")
    vReplace = Array("^p", "^p")

    Set rng = Doc.Range(0, 0)
        
    For i = 0 To UBound(vWhat)
         With rng.Find
          .Text = vWhat(i)             '検索する文字列
          .Replacement.Text = vReplace(i)    '置換後の文字列
          .Forward = True             '検索方向
          .Wrap = wdFindContinue      '検索対象のオブジェクトの末尾での操作
          .Format = False             '書式
          .MatchCase = False          '大文字と小文字の区別する
          .MatchWholeWord = False     '完全に一致する単語だけを検索する
          .MatchByte = False          '半角と全角を区別する
          .MatchAllWordForms = False  '英単語の異なる活用形を検索する
          .MatchSoundsLike = False    'あいまい検索（英）
          .MatchFuzzy = False         'あいまい検索（日）
          .MatchWildcards = True      'ワイルドカードを使用する
        End With
        rng.Find.Execute Replace:=wdReplaceAll
    Next i
    Set rng = Nothing
End Sub

Private Sub フォント種類を設定する(Doc, fontFarEast, FontAscii)
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = fontFarEast
        .NameAscii = FontAscii
        .NameOther = FontAscii
        .Name = ""
    End With
    Selection.Collapse
End Sub

Private Sub ScreenUpdatingSwitch()
    Application.ScreenUpdating = Not Application.ScreenUpdating
End Sub


