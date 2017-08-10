Attribute VB_Name = "GetTTS"
Option Explicit
Dim tmpSt As Worksheet

Sub ShowTTS()
    Application.ScreenUpdating = False
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.application")
    IE.Visible = False
    IE.navigate "http://www.bk.mufg.jp/gdocs/rate/real_01.html"
    Do While IE.busy Or IE.readystate < 4 ' READYSTATE_COMPLETE�̒l
        DoEvents
    Loop
    
    Dim CurrentSt As Worksheet
    Set CurrentSt = ActiveSheet
    
    Call MakeList(IE)
    Dim Rate As String
    Rate = GetTTS
    Call DeleteSheet '�폜
    IE.Quit 'IE�����i�Еt���j
    Application.ScreenUpdating = True
       
    CurrentSt.Select
    Set CurrentSt = Nothing
    MsgBox "���݂�TTS���[�g�F" & Rate
    
End Sub
Sub MakeList(objIE As Object)

'�ꎞ�V�[�g�ɃE�F�u�y�[�W�̓��e���������݂܂�
    Dim n As Long
    Dim r As Long
    Dim Doc As Object
    Dim tmpSt As Worksheet
    Dim objTD As Object
    Dim objTag As Object
    n = 0
    r = 0
    
    Set tmpSt = Sheets.Add
    With tmpSt
        With .Range("a1", ActiveCell.SpecialCells(xlLastCell))
            .ClearContents
            .NumberFormatLocal = "G/�W��"
        End With
        
        Set Doc = objIE.document
        Set objTD = Doc.getElementsByTagName("TD")
        For Each objTag In objTD
            r = r + 1
            .Cells(r + 1, 1) = objTag.tagName
            .Cells(r + 1, 2) = n
            .Cells(r + 1, 3) = r
            .Cells(r + 1, 4) = objTag.innerText
        Next objTag
    End With
End Sub

Private Function GetTTS() As String
'�ꎞ�V�[�g����TTS�̏������o���܂�
    With tmpSt
        Dim i As Long
        For i = 1 To Range("A10000").End(xlUp).Row
            If Range("D" & i) = "EUR (���[��)" Then
                GetTTS = Range("D" & i).offset(1, 0).Value
                Exit Function '�K�v�ȃl�^���������o�܂�
            End If
        Next
    End With
End Function

Sub DeleteSheet()
'�ꎞ�V�[�g���폜���܂�
    Application.DisplayAlerts = False '�u�{���ɍ폜���܂����H�v��\�������Ȃ�
    On Error Resume Next
    tmpSt.Delete
    Application.DisplayAlerts = True '���ɖ߂�
    On Error GoTo 0
End Sub
