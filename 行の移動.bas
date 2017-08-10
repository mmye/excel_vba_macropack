Attribute VB_Name = "�s�̈ړ�"
Option Explicit

Sub GoDownRow()
    Dim UpperRow As Range
    Dim LowerRow As Range
    Dim Upper As Variant
    Dim Lower As Variant
    Dim Cols As Long
    Dim SelRow As Long
    Dim RowOffset As Long
    
    Cols = ActiveSheet.UsedRange.Columns.Count
    SelRow = ActiveCell.Row
    RowOffset = Selection.Rows.Count
    
    'Selection�ɑΉ������͈͕t���ɂ��邱��
    Set UpperRow = Intersect(Range(Cells(SelRow, 1), Cells(SelRow, Cols).offset(RowOffset - 1, 0)), _
                        ActiveSheet.UsedRange)
    Set LowerRow = Intersect(Range(Cells(SelRow, 1).offset(1, 0), Cells(SelRow, Cols).offset(RowOffset, 0)), _
                        ActiveSheet.UsedRange)
    If UpperRow Is Nothing Then Exit Sub
    If LowerRow Is Nothing Then Exit Sub
    '��ɂ悯��K�v������͈̂�ԉ��̍s����
    Lower = UpperRow.Rows(UpperRow.Rows.Count).offset(1, 0)
    Upper = UpperRow
    
    Dim MoveTo As Range, EscapeTo As Range
    Set MoveTo = UpperRow.offset(1, 0)
    Set EscapeTo = UpperRow.Rows(1)
    EscapeTo = Lower
    MoveTo = Upper
    
    Selection.offset(1, 0).Select
    
    Set UpperRow = Nothing
    Set LowerRow = Nothing
    Set MoveTo = Nothing
    Set EscapeTo = Nothing
    Upper = Empty
    Lower = Empty
    
End Sub

Sub LiftRow()
    Dim UpperRow As Range
    Dim LowerRow As Range
    Dim Upper As Variant
    Dim Lower As Variant
    Dim Cols As Long
    Dim SelRow As Long
    Dim RowOffset As Long
    
    If ActiveCell.Row - Selection.Rows.Count < 1 Then Exit Sub
    
    Cols = ActiveSheet.UsedRange.Columns.Count
    SelRow = ActiveCell.Row
    RowOffset = Selection.Rows.Count
    
    'Selection�ɑΉ������͈͕t���ɂ��邱��
    Set UpperRow = Intersect(Range(Cells(SelRow, 1).offset(-1, 0), Cells(SelRow, Cols).offset(RowOffset - 2, 0)), _
                        ActiveSheet.UsedRange)
    Set LowerRow = Intersect(Range(Cells(SelRow, 1).offset(RowOffset - 1, 0), Cells(SelRow, Cols).offset(RowOffset - 1, 0)), _
                        ActiveSheet.UsedRange)

    '���ɂ悯��K�v������͈̂�ԏ�̍s����
    Upper = UpperRow.Rows(1)
    Lower = UpperRow.offset(1)
    
    Dim MoveTo As Range, EscapeTo As Range
    Set MoveTo = UpperRow
    Set EscapeTo = LowerRow
    
    If MoveTo Is Nothing Or EscapeTo Is Nothing Then
        MsgBox "�Z���g�p�͈͊O�ɂ͈ړ��ł��܂���"
        Exit Sub
    End If
    EscapeTo = Upper
    MoveTo = Lower
    
    Selection.offset(-1, 0).Select
    
    Set UpperRow = Nothing
    Set LowerRow = Nothing
    Set MoveTo = Nothing
    Set EscapeTo = Nothing
    Upper = Empty
    Lower = Empty
    
End Sub
