VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FProgressBar 
   Caption         =   "Professional Excel Development"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   OleObjectBlob   =   "FProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "FProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



' Description:  Displays a modeless progress bar on the screen
'
' Authors:      Stephen Bullen, www.oaltd.co.uk
'               Rob Bovey, www.appspro.com
'

Option Explicit

' **************************************************************
' Module-level variables to store the property values
' **************************************************************
Dim mdMin As Double
Dim mdMax As Double
Dim mdProgress As Double
Dim mdLastPerc As Double
Dim mlhWnd As Long
Dim mbCancelled As Boolean

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000

' **************************************************************
' Module DLL Declarations Follow
' **************************************************************

'Windows API calls to remove the [x] from the top-right of the form
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Windows API calls to bring the progress bar form to the front of other modeless forms
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Initialise the form to show a blank text and 0% complete
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Private Sub UserForm_Initialize()

    On Error Resume Next

    'Get the form's window handle, for use in API calls
    Me.Caption = "翻訳を取得しています…"
    mlhWnd = FindWindow(vbNullString, Me.Caption)

    'Assume an initial progress of 0-100
    mdMin = 0
    mdMax = 100

    lblMessage.Caption = ""
    Me.Caption = ""
    RemoveCloseButton
    Me.progress = 54

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Let the calling routine set/get the caption of the form
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Public Property Let title(ByVal RHS As String)
    Me.Caption = RHS
    RemoveCloseButton
End Property

Public Property Get title() As String
    title = Me.Caption
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Let the calling routine set/get the descriptive text on the form
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Public Property Let Text(ByVal RHS As String)

    If RHS <> lblMessage.Caption Then
        lblMessage.Caption = RHS

        'Refresh the form if it's being shown
        If Me.Visible Then
            Me.Repaint
            BringToFront
        End If
    End If

End Property

Public Property Get Text() As String
    Text = lblMessage.Caption
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Let the calling routine set/get the Minimum scale for the progress bar
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Public Property Let Min(ByVal RHS As Double)
    mdMin = RHS
End Property

Public Property Get Min() As Double
    Min = mdMin
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Let the calling routine set the Maximum scale for the progress bar
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Public Property Let max(ByVal RHS As Double)
    mdMax = RHS
End Property

Public Property Get max() As Double
    max = mdMax
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Whether the Cancel button has been clicked
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Public Property Get Cancelled() As Boolean
    Cancelled = mbCancelled
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Show the progress form
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Public Sub ShowForm()

    'Remove the [x] close button on the form
    RemoveCloseButton

    'Show the form modelessly
    Me.Show vbModeless

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Handle clicking the Cancel button
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Private Sub btnCancel_Click()
    lblMessage.Caption = "処理をキャンセルしています…"
    mbCancelled = True
    btnCancel.Enabled = False
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Set the progress bar's progress
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Public Property Let progress(ByVal RHS As Double)

    Dim dPerc As Double

    mdProgress = RHS

    'Calculate the progress percentage
    If mdMax = mdMin Then
        dPerc = 0
    Else
        dPerc = Abs((RHS - mdMin) / (mdMax - mdMin))
    End If

    'Only update the form every 0.5% change
    If Abs(dPerc - mdLastPerc) > 0.005 Then
        mdLastPerc = dPerc

        'Set the width of the inside frame, rounding to the nearest pixel
        fraInside.Width = Int(lblBack.Width * dPerc / 0.75 + 1) * 0.75

        'Set the captions for the blue-on-white and white-on-blue texts.
        lblBack.Caption = Format(dPerc, "0%")
        lblFront.Caption = Format(dPerc, "0%")

        'Refresh the form if it's being shown
        If Me.Visible Then
            Me.Repaint
            BringToFront
        End If

        'Allow Cancel click to be processed at each update
        DoEvents
    End If

End Property

Public Property Get progress() As Double
    progress = mdProgress
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Ignore clicking the [x] on the dialog (which shouldn't be visible anyway!)
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then Cancel = True

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Routine to remove the [x] from the top-right of the form
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Private Sub RemoveCloseButton()

    Dim lStyle As Long

    'Remove the close button on the form
    lStyle = GetWindowLong(mlhWnd, GWL_STYLE)

    If lStyle And WS_SYSMENU > 0 Then
        SetWindowLong mlhWnd, GWL_STYLE, (lStyle And Not WS_SYSMENU)
    End If

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Routine to bring the form to the front of other modeless forms
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 04 Jun 04     Stephen Bullen  Initial version
'
Private Sub BringToFront()

    Dim lFocusThread As Long
    Dim lThisThread As Long

    'Does the window being viewed by the user belong to the same thread as the progress bar?
    '(i.e. are they still looking at this instance of Excel)
    lFocusThread = GetWindowThreadProcessId(GetForegroundWindow(), 0)
    lThisThread = GetWindowThreadProcessId(mlhWnd, 0)

    If lFocusThread = lThisThread Then
        'The threads are the same, so force the progress bar in front of other modeless dialogs
        SetForegroundWindow mlhWnd
    Else
        'Not looking at Excel, so yield control to the app they're using
        '(and allow them to switch back to Excel)
        DoEvents
    End If

End Sub
