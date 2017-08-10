Attribute VB_Name = "シート塗りつぶしチェック"
Option Explicit


'---------------------------------------------------------------------------------------
' Method : CheckByInteriorColor
' Author : m.maeyama
' Date   : 2017/03/28
' Purpose: シートA、Bがあるとき、更新版シートA'、B'を得て、更新されたセルを示す塗りつぶし
'---------------------------------------------------------------------------------------
Sub CheckByInteriorColor()
    Dim rFrom As Range
    Dim rTo As Range
    
    Dim wksFrom As Worksheet
    Dim wksTo As Worksheet
    
    Set wksFrom = Sheets(1)
    Set wksTo = Sheets(2)
    
    Dim i As Long
    
    Set rFrom = wksFrom.UsedRange
    
    Dim r As Range
    For Each r In rFrom
        '比較元のセルが塗りつぶされているかをしらべる
        If r.Interior.color = 65535 Then
            Dim Row, col
            Row = r.Row
            col = r.column
            
            Dim color As Long
            color = r.Interior.color
            
            '比較先シートを同じ色で塗りつぶす
            wksTo.Cells(Row, col).Interior.color = color
        End If
    Next r
    
    Set wksFrom = Nothing
    Set wksTo = Nothing
    Set rFrom = Nothing
    Set rTo = Nothing
End Sub
