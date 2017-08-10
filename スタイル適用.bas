Attribute VB_Name = "スタイル適用"
Option Explicit


Sub ApplyStyleGothic()
    Dim i As Long
    Dim r As Range
    Dim r1 As Range, r2 As Range, r3 As Range
    Dim rngHasData As Range
    Dim objChar As Object
    
    Set r1 = Cells.SpecialCells(xlCellTypeConstants)
    Set r2 = Cells.SpecialCells(xlCellTypeFormulas)
    If Not r2 Is Nothing Then
        Set rngHasData = Union(r1, r2, Range("a1"))
    Else
        Set rngHasData = Union(r1, Range("a1"))
    End If
    
    Application.ScreenUpdating = False
    For Each r In rngHasData
        If Not IsError(r.Value) Then
        If r.Value <> "" Then
        DoEvents
        On Error Resume Next
        For i = 1 To r.Characters.Count
            Set objChar = r.Characters(i, 1)
            If LenB(StrConv(objChar.Text, vbFromUnicode)) = 1 Then
                objChar.Font.Name = "Arial"
            Else
                objChar.Font.Name = "ＭＳ Ｐゴシック"
            End If
        Next i
        End If
        End If
    Next r
    
    Application.ScreenUpdating = True
    
    Set rngHasData = Nothing
    Set r1 = Nothing
    Set r2 = Nothing
    Set objChar = Nothing
    
End Sub

