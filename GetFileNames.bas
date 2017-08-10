Attribute VB_Name = "GetFileNames"
Option Explicit

Public Function GetFileNames(Path As String) As Variant
    Dim buf As String
    Dim c As Long: c = 0
    Dim fs() As String

    If Path = "" Then Exit Function
    'hogeˆÈ‰º‚É‚ ‚éƒtƒ@ƒCƒ‹ˆê——‚ðŽæ“¾‚µ‚½‚¢
    'Const Path As String = "D:\buf\"
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    buf = Dir(Path & "*.xls*")
    Do While buf <> ""
        ReDim Preserve fs(c) As String
        fs(c) = buf
        c = c + 1
        buf = Dir()
    Loop

    GetFileNames = fs
End Function
