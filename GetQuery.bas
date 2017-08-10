Attribute VB_Name = "GetQuery"
Option Explicit

Sub GetWebTable()
'+++++++++++++++++++++++++++++++++++
'�ǉ������V�[�g�ɏ����������Ƃ���ƃG���[�ɂȂ�B
'�A�N�e�B�u�V�[�g���g���Ζ��Ȃ��B20161028
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


    Set r = TempWks.UsedRange.Find("EUR (���[��)").offset(0, 1)
    myEURO = r.Value
    MsgBox "���[���בփ��[�g��" & myEURO & "�ł�", vbInformation, "UFJ�ŐV�בփ��[�g"
    
    Set TempWks = Nothing
    Set r = Nothing
    Set qt = Nothing
    
Exit Sub

ERR:
'    Dim msg: msg = "�f�[�^�̎擾�Ɏ��s���܂���"
    MsgBox ERR.Number & ERR.Description, vbCritical
    Set TempWks = Nothing
    Set r = Nothing
End Sub


