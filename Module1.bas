Attribute VB_Name = "Module1"

Dim ret As Long
Dim i As Long
Dim rowCount As Long
Dim colCount As Long
Dim colType As Long
Dim sql As String
'Dim dbPath As String
Dim dbh As Long
Dim stmtHandle As Long
Dim Rows() As Variant

'DB、テーブルは既に作成済みとします。
Const dbPath As String = "C:\Users\m.maeyama\Desktop\payment_conditions.sqlite3"

Function get_all_records() As Variant

    '＜SELECT＞
    sql = "SELECT * FROM payment_conditions;"
    
    ret = SQLite3Open(dbPath, dbh)
    ret = SQLite3PrepareV2(dbh, sql, stmtHandle)
    ret = SQLite3Step(stmtHandle)
    rowCount = 0
    'RecordsetオブジェクトのGetRowsメソッドの取得と同じ
    'Rows = ***.GetRows()
    Do While ret <> SQLITE_DONE
        If rowCount = 0 Then
            colCount = SQLite3ColumnCount(stmtHandle)
            ReDim Rows(colCount - 1, rowCount)
        Else
            ReDim Preserve Rows(colCount - 1, rowCount)
        End If
        For i = 0 To colCount - 1
            colType = SQLite3ColumnType(stmtHandle, i)
            Rows(i, rowCount) = getColValue(stmtHandle, i, colType)
        Next
        ret = SQLite3Step(stmtHandle)
        rowCount = rowCount + 1
    Loop
    
    ret = SQLite3Finalize(stmtHandle)
    ret = SQLite3Close(dbh)
    
    get_all_records = Rows

End Function

Sub insert()

    sql = "INSERT INTO SAMPLE VALUES('1', 'ABCD', '9999')"
    
    ret = SQLite3Open(dbPath, dbh)
    ret = SQLite3PrepareV2(dbh, sql, stmtHandle)
    
    If ret <> SQLITE_DONE Then
        Debug.Print "SQL error: " & SQLite3ErrMsg(dbh)
    End If
    
    ret = SQLite3Step(stmtHandle)
    ret = SQLite3Finalize(stmtHandle)
    ret = SQLite3Close(dbh)

End Sub


'Sqlite3Demo.basより
Private Function getColValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            getColValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            getColValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            getColValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            getColValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            getColValue = Null
    End Select
End Function
