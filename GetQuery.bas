Attribute VB_Name = "GetQuery"
Option Explicit

Sub GetWebTable()
'+++++++++++++++++++++++++++++++++++
'追加したシートに書き込もうとするとエラーになる。
'アクティブシートを使えば問題ない。20161028
'+++++++++++++++++++++++++++++++++++

    Dim TempWks As Worksheet
    Dim qt As QueryTable
    Dim r As Range
    Dim myEURO As String
    
    Set TempWks = ActiveSheet
    
    On Error GoTo ERR
    Set qt = TempWks.QueryTables.Add(Connection:= _
                "URL;http://www.bk.mufg.jp/gdocs/rate/real_01.html", Destination:=TempWks.Range("a1"))
    qt.Name = "RealTime_EURO"
    qt.WebFormatting = xlWebFormattingNone
    qt.Refresh


    Set r = TempWks.UsedRange.Find("EUR (ユーロ)").offset(0, 1)
    myEURO = r.Value
    MsgBox "ユーロ為替レートは" & myEURO & "です", vbInformation, "UFJ最新為替レート"
    
    Set TempWks = Nothing
    Set r = Nothing
    Set qt = Nothing
    
Exit Sub

ERR:
'    Dim msg: msg = "データの取得に失敗しました"
    MsgBox ERR.Number & ERR.Description, vbCritical
    Set TempWks = Nothing
    Set r = Nothing
End Sub


