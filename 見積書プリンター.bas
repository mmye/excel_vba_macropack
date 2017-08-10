Attribute VB_Name = "見積書プリンター"
Option Explicit

Sub ExtractContents()
'  指定した列に含まれる文字列をコンソールに出力する。

    Dim Lists1 As Variant
    Dim Lists2 As Variant
    Dim Lists3 As Variant
    Dim i As Long
    Dim str As String
    Const Delimeter As String = vbTab
    
    Lists1 = GetteXt
    
    For i = LBound(Lists1) To UBound(Lists1)
        If Lists1(i) <> "" Then
            str = str & Lists1(i) & vbCrLf
        End If
    Next i
    
    Debug.Print str
    
End Sub

Private Function GetteXt() As Variant
    Dim i As Long, j As Long, c As Long
    Dim StartRow As Long, EndRow As Long
    Dim UserInput As String, v As Variant
    Dim TargetCol1 As String, TargetCol2 As String, TargetCol3 As String
    Dim buf
    Dim str As String
    Dim Lists() As String, List As String
    Const Delimeter As String = vbTab
    
    UserInput = InputBox("取り出す列をカンマで区切ってアルファベットで指定して下さい…　例.C, L, BP")
    v = Split(UserInput, ",")
    If IsEmpty(v) Then Exit Function
    TargetCol1 = v(0)
    StartRow = 38
    EndRow = Cells(Rows.Count, v(0)).End(xlUp).Row
    ReDim Lists(EndRow - StartRow) As String
    For i = StartRow To EndRow
        buf = Cells(i, v(0))
        For j = LBound(v) To UBound(v)
            If buf <> "" And Not IsError(buf) Then
                List = List & Trim$(Cells(i, v(j))) & Delimeter
            End If
        Next j
        Lists(c) = List
        List = Empty
        c = c + 1
    Next i
    
'  行数を書き出す
    ReDim Preserve Lists(UBound(Lists) + 1) As String
    Lists(c) = "行数：" & CStr((c))
    GetteXt = Lists
    
End Function

