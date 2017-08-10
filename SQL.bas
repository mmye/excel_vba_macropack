Attribute VB_Name = "SQL"
Function GetRecords(sqlQuery As String, Optional fileName As String = vbNullString, Optional includesColumnNames As Boolean = True) As Object
'Run an SQL query on current or external file, and return a RecordSet of records
'fileName - if null than assumes current file, otherwise allows you to query external files
'includesColumnNames - if true assumes first row contains column headers
    If fileName = vbNullString Then fileName = ThisWorkbook.FullName
    Dim cn As Object, rs As Object
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileName _
        & ";Extended Properties=""Excel 12.0;HDR=" & IIf(includesColumnNames, "Yes", "No") & ";IMEX=1"";"
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    cn.Open strCon
    rs.Open sqlQuery, cn
    Set GetRecords = rs
End Function
Function GetRecordsArray(sqlQuery As String, Optional fileName As String = vbNullString, Optional includesColumnNames As Boolean = True) As Variant()
'Run an SQL query on current or external file, and return a variant array of records
'fileName - if null than assumes current file, otherwise allows you to query external files
'includesColumnNames - if true assumes first row contains column headers
    If fileName = vbNullString Then fileName = ThisWorkbook.FullName
    Dim cn As Object, rs As Object
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileName _
        & ";Extended Properties=""Excel 12.0;HDR=" & IIf(includesColumnNames, "Yes", "No") & ";IMEX=1"";"
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    cn.Open strCon
    rs.CursorLocation = 3
    rs.Open sqlQuery, cn
    Dim retArr() As Variant, Row As Long, column As Long, i As Long
    ReDim retArr(rs.RecordCount - 1, rs.Fields.Count - 1) As Variant
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF = True
            For i = 0 To rs.Fields.Count - 1
                retArr(Row, i) = rs(i)
            Next i
            Row = Row + 1
            rs.MoveNext
        Loop
    End If
    GetRecordsArray = retArr
End Function

