Attribute VB_Name = "ExtractPOS_Item"
Option Explicit

Sub GetPos_Item()
    Dim i As Long
    Dim c As Long: c = 1

    Dim EndRow As Long
    EndRow = InputBox("最後の項目の行番号を入力してください...")
    If EndRow = 0 Then Exit Sub
    If Not IsNumeric(EndRow) Then
        MsgBox "数値で指定してください"
        Exit Sub
    End If

    Dim POSCol As String
    POSCol = InputBox("POSの列名をアルファベットで入力してください...")
    If StrConv(POSCol, vbHiragana) <> POSCol Then Exit Sub

    Dim ItemCol As String
    ItemCol = InputBox("品名の列名をアルファベットで入力してください...")
    If StrConv(ItemCol, vbHiragana) <> ItemCol Then Exit Sub

    Dim st As Worksheet
    Set st = ActiveWorkbook.Worksheets.Add

    For i = 40 To EndRow
        If Cells(i, POSCol).Value <> "" Then
            Dim POS As String
            POS = Cells(i, POSCol).Value
            Dim Item As String
            Item = Cells(i, ItemCol).Value
            st.Cells(c, "A").Value = POS
            st.Cells(c, "B").Value = Item
            c = c + 1
         End If
    Next i

    st.Select
    st.Cells(1, 1).Select
    MsgBox "POSと品名を抽出しました"

End Sub


