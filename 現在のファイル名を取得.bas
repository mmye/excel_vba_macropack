Attribute VB_Name = "現在のファイル名を取得"
Option Explicit

Sub GetActiveWorkbookName()
    Dim buf2 As String
    Dim CB As New DataObject
    Dim buf
    
    buf = ActiveWorkbook.Name
    
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
        buf2 = .GetteXt     ''DataObjectのデータを変数に取得する
    End With
    MsgBox "ファイル名をクリップボードにコピーしました。" & vbCrLf & buf
    
    Set CB = Nothing
End Sub
