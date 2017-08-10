Attribute VB_Name = "PDFWord�ϊ�"
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
    �t�H���g��ނ�ݒ肷�� ConvertedDoc, "�l�r �o����", "Arial"
    �s���s���̋󔒕������폜 ConvertedDoc
    ScreenUpdatingSwitch
    Set ConvertedDoc = Nothing
End Sub
 
Private Function Convert2PDF(ByVal TargetFilePath As String, _
                       ByVal TargetConvType As Conv) As Object
'   PDF�𑼂̃t�@�C���`���ɕϊ�
    Dim jso As Object
    Dim convid As String
    Dim ext As String
    Dim fp As String, fn As String, File As String
     
'   �t�H���_�p�X�ƃt�@�C�����擾
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
'cConvID�擾
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
'�g���q�擾
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
    '�S�p�p�����𔼊p�p�����ֈꊇ�ϊ�
    Dim rng As Range
    Set rng = Doc.Range(0, 0)
    With rng.Find
        .Text = "[�O-�X�`-�y��-��]{1,}"  '�Ώۂ̐ݒ�
        .MatchWildcards = True
        Do While .Execute = True
            rng.Collapse wdCollapseEnd
        Loop
    End With
    Set rng = Nothing
End Sub

Private Sub �s���s���̋󔒕������폜(Doc)
    Dim rng As Range
    Dim vWhat As Variant
    Dim vReplace As Variant
    Dim i As Long

    '�u�����������Ƀx�^�ł��B�A�����锼�p�E�S�p�X�y�[�X�E�^�u����ɂ���
    vWhat = Array("^13([ �@^t]{1,})", "([ �@^t]{1,})^13")
    vReplace = Array("^p", "^p")

    Set rng = Doc.Range(0, 0)
        
    For i = 0 To UBound(vWhat)
         With rng.Find
          .Text = vWhat(i)             '�������镶����
          .Replacement.Text = vReplace(i)    '�u����̕�����
          .Forward = True             '��������
          .Wrap = wdFindContinue      '�����Ώۂ̃I�u�W�F�N�g�̖����ł̑���
          .Format = False             '����
          .MatchCase = False          '�啶���Ə������̋�ʂ���
          .MatchWholeWord = False     '���S�Ɉ�v����P�ꂾ������������
          .MatchByte = False          '���p�ƑS�p����ʂ���
          .MatchAllWordForms = False  '�p�P��̈قȂ銈�p�`����������
          .MatchSoundsLike = False    '�����܂������i�p�j
          .MatchFuzzy = False         '�����܂������i���j
          .MatchWildcards = True      '���C���h�J�[�h���g�p����
        End With
        rng.Find.Execute Replace:=wdReplaceAll
    Next i
    Set rng = Nothing
End Sub

Private Sub �t�H���g��ނ�ݒ肷��(Doc, fontFarEast, FontAscii)
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


