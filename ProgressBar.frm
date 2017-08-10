VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Progress"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim progress As Double, maxProgress As Double, maxWidth As Long, startTime As Double
Public Sub Initialize(title As String, Optional max As Long = 100)
'Initialize and shor progress bar
    Me.Caption = title
    maxProgress = max:  maxWidth = lBar.Width:    lBar.Width = 0
    lProgress.Caption = "0%"
    Me.Show False
    startTime = Time
End Sub
Public Sub AddProgress(Optional inc As Long = 1)
'Increase progress by an increment
    Dim tl As Double, tlMin As Integer, tlSec As Integer, tlHour As Integer, tlTotal As Integer, tlTotalSec, tlTotalMin, tlTotalHour
    progress = progress + inc
    If progress > maxProgress Then progress = maxProgress
    lBar.Width = CLng(CDbl(progress) / maxProgress * maxWidth)
    DoEvents
    tl = Time - startTime
    tlSec = Second(tl) + Minute(tl) * 60 + Hour(tl) * 3600
    tlTotal = tlSec
    If progress = 0 Then
        tlSec = 0
    Else
        tlSec = (tlSec / progress) * (maxProgress - progress)
    End If
    tlHour = Floor(tlSec / 3600)
    tlTotalHour = Floor(tlTotal / 3600)
    tlSec = tlSec - 3600 * tlHour
    tlTotal = tlTotal - 3600 * tlTotalHour
    tlMin = Floor(tlSec / 60)
    tlTotalMin = Floor(tlTotal / 60)
    tlSec = tlSec - 60 * tlMin
    tlTotal = tlTotal - 60 * tlTotalMin
    If tlSec > 0 Then
        tlMin = tlMin + 1
    End If
    'Captions
    lProgress.Caption = "" & CLng(CDbl(progress) / maxProgress * 100) & "%"
    lTimeLeft.Caption = "" & tlHour & " hours, " & tlMin & " minutes"
    lTimePassed.Caption = "" & tlTotalHour & " hours, " & tlTotalMin & " minutes, " & tlTotal & " seconds"
    'Hide if finished
    If progress = maxProgress Then Me.Hide
End Sub
Public Function Floor(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    Floor = Int(x / Factor) * Factor
End Function
