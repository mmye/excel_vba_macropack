Attribute VB_Name = "Z_Translation"
Option Explicit

Type WordList
    Index As Long
    OriginalStr As String
    Translation As String
    Address As String
End Type
Dim Lists() As WordList

Sub ScrapeWks()
    Dim rng As Range, r As Range
    Dim rCnt As Long, c As Long
    
    rCnt = ActiveSheet.UsedRange.Count
    Set rng = ActiveSheet.UsedRange
    ReDim Lists(rCnt) As WordList
    
    For Each r In rng
        If r.Value <> "" Then
            Lists(c).Index = c
            Lists(c).OriginalStr = r.Value
            Lists(c).Translation = Translations.Translate(r.Value, Translations.English, Translations.Japanese)
            Lists(c).Address = r.Address
            Debug.Print "Index:" & Lists(c).Index & vbCr & "Text:" & Lists(c).OriginalStr & vbCr & _
                        "Translation:" & Lists(c).Translation & vbCr & "Address:" & Lists(c).Address
            c = c + 1
        End If
'        Debug.Print r.Value
    Next
    Set rng = Nothing
    Set r = Nothing
    
End Sub
Sub TranslateExampleWithQuotationJAtoEN()
    Dim r As Range
    For Each r In ActiveSheet.UsedRange
        On Error Resume Next
        If r.Value <> "" Then r.Value = Translations.Translate(r.Value, Translations.Japanese, Translations.English)
    Next
    On Error GoTo 0
End Sub
