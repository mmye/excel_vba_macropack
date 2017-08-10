Attribute VB_Name = "使い捨てマクロ"
Option Explicit

Sub ReplaceErrwithNA()
    Dim rng As Range, r As Range, rErr As Range, r2 As Range
    Const strErr As String = "N/A"
    
    Set rng = ActiveSheet.UsedRange
    For Each r In rng
        If IsError(r.Value) Then
            If rErr Is Nothing Then
                Set rErr = r
            Else
                Set rErr = Union(rErr, r)
            End If
        End If
    Next r
    
    If rErr Is Nothing Then Exit Sub
    rErr.Value = Empty
    
    For Each r2 In rErr
        r2.Value = strErr
    Next r2
    
    Set r = Nothing
    Set r2 = Nothing
    Set rErr = Nothing
    
End Sub
Sub ReplaceErrwithNABook()
    Dim rng As Range, r As Range, rErr As Range, r2 As Range
    Dim wks As Worksheet
    Const strErr As String = "N/A"
    
    For Each wks In ActiveWorkbook.Sheets
        wks.Select
        Set rng = ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas)
        If Not rng Is Nothing Then
            For Each r In rng
                If IsError(r.Value) Then
                    If rErr Is Nothing Then
                        Set rErr = r
                    Else
                        Set rErr = Union(rErr, r)
                    End If
                End If
            Next r
            
            If rErr Is Nothing Then Exit Sub
            rErr.Value = Empty
            
            For Each r2 In rErr
                r2.Value = strErr
            Next r2
        End If
    Next wks
    
    Set r = Nothing
    Set r2 = Nothing
    Set rErr = Nothing
    
End Sub

Sub FillNA()
    Dim rng As Range
    Dim r As Range, rFormulas As Range
    Dim v As Variant
    
    Set rFormulas = ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas)
    If Not rFormulas Is Nothing Then
        For Each r In rFormulas
            DoEvents
            If IsError(r.Value) Then
                On Error Resume Next
                r.Value = "N/A"
                On Error GoTo 0
            End If
        Next r
    End If
    
    Set rFormulas = Nothing
    Set r = Nothing
    
End Sub

Sub SplitText()
    Dim i As Long
    Dim r As Range
    Dim colName As String
    Dim StartPos As Long
    Dim WantedStr As String
    Dim RemainedStr As String
    Dim buf As String
    Dim str As String
    colName = "ネスト本数"
    str = "/"
    
    For i = 2 To Cells(Rows.Count, Range(colName).column).End(xlUp).Row
        DoEvents
        Set r = Cells(i, Range(colName).column)
        buf = r.Value
        If InStr(buf, str) > 0 Then
            StartPos = InStr(buf, str) + 1
            WantedStr = Trim(Right$(buf, Len(buf) - StartPos + 1))
             Cells(i, Range("バイアル径").column).Value = WantedStr
'            Debug.Print "All:" & buf & vbTab & "SplitStr：" & WantedStr & vbTab & "Remained：" & Left$(buf, StartPos - 2)
            RemainedStr = Left$(buf, StartPos - 2)
            r.Value = RemainedStr
        End If
    Next i
End Sub

Sub ExtractContents()
    Dim Lists1 As Variant
    Dim Lists2 As Variant
    Dim Lists3 As Variant
    Dim i As Long
    Dim str As String
    Const Delimeter As String = vbTab
    
    Lists1 = FetchContentsFromExcelHouganshi
    Lists2 = FetchContentsFromExcelHouganshi2
    Lists3 = FetchContentsFromExcelHouganshi3
    
    For i = LBound(Lists1) To UBound(Lists1)
        If Lists1(i) <> "" Then
            str = str & Lists1(i) & Delimeter & _
                        Lists2(i) & Delimeter & _
                        Lists3(i) & Delimeter & vbCrLf
        End If
    Next i
    
    Debug.Print str
    
End Sub

Private Function FetchContentsFromExcelHouganshi() As Variant
    Dim i As Long, c As Long
    Dim StartRow As Long, EndRow As Long
    Dim TargetCol1 As String, TargetCol2 As String, TargetCol3 As String
    Dim buf
    Dim str As String
    Dim Lists() As String
    TargetCol1 = "C"
    TargetCol2 = ""
    
    StartRow = 38
    EndRow = Cells(Rows.Count, TargetCol1).End(xlUp).Row
    ReDim Lists(EndRow - StartRow) As String
    For i = StartRow To EndRow
        buf = Cells(i, TargetCol1)
'        If buf <> "" And Not IsError(buf) And Left$(buf, 1) <> "*" Then
        If buf <> "" And Not IsError(buf) Then
            Lists(c) = buf
            c = c + 1
        End If
    Next i
    FetchContentsFromExcelHouganshi = Lists
    
End Function
Private Function FetchContentsFromExcelHouganshi2() As Variant
    Dim i As Long, c As Long
    Dim StartRow As Long, EndRow As Long
    Dim TargetCol1 As String, TargetCol2 As String, TargetCol3 As String
    Dim buf
    Dim str As String
    Dim Lists() As String
    TargetCol1 = "L"
    TargetCol2 = ""
    
    StartRow = 38
    EndRow = Cells(Rows.Count, TargetCol1).End(xlUp).Row
    ReDim Lists(EndRow - StartRow) As String
    
    For i = StartRow To EndRow
        buf = Cells(i, TargetCol1)
        If buf <> "" And Not IsError(buf) Then
            Lists(c) = buf
            c = c + 1
        End If
    Next i
    FetchContentsFromExcelHouganshi2 = Lists
    
End Function
Private Function FetchContentsFromExcelHouganshi3() As Variant
    Dim i As Long, c As Long
    Dim StartRow As Long, EndRow As Long
    Dim TargetCol1 As String, TargetCol2 As String, TargetCol3 As String
    Dim buf
    Dim Lists() As String
    TargetCol1 = "BP"
    TargetCol2 = ""
    
    StartRow = 38
    EndRow = Cells(Rows.Count, TargetCol1).End(xlUp).Row
    ReDim Lists(EndRow - StartRow) As String
    For i = StartRow To EndRow
        buf = Cells(i, TargetCol1)
        If buf <> "" And Not IsError(buf) Then
            Lists(c) = buf
            c = c + 1
        End If
    Next i
    FetchContentsFromExcelHouganshi3 = Lists
    
End Function
