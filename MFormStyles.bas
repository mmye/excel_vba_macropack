Attribute VB_Name = "MFormStyles"
'
' Description:  Module to modify a userform's window styles
'
' Authors:      Stephen Bullen, www.oaltd.co.uk
'               Rob Bovey, www.appspro.com
'

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows API Declarations and Constants Follow
''''''''''''''''''''''''''''''''''''''''''''''''''

'Windows API calls to do all the dirty work
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

'Window API constants
Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20)         'The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar
Private Const WS_SYSMENU As Long = &H80000        'Style to add a system menu
Private Const WS_THICKFRAME As Long = &H40000     'Style to add a sizable frame
Private Const WS_MINIMIZEBOX As Long = &H20000    'Style to add a Minimize box on the title bar
Private Const WS_MAXIMIZEBOX As Long = &H10000    'Style to add a Maximize box to the title bar
Private Const WS_EX_DLGMODALFRAME As Long = &H1   'Controls if the window has an icon
Private Const WS_EX_TOOLWINDOW As Long = &H80     'Tool Window: small titlebar
Private Const SC_CLOSE As Long = &HF060           'Close menu item


''''''''''''''''''''''''''''''''''''''''''''''''''
' Module Enumerations Follow
''''''''''''''''''''''''''''''''''''''''''''''''''

'Public enum of our userform styles
Public Enum UserformWindowStyles
    uwsNoTitleBar = 0
    uwsHasTitleBar = 1
    uwsHasSmallTitleBar = 2
    uwsHasMaxButton = 4
    uwsHasMinButton = 8
    uwsHasCloseButton = 16
    uwsHasIcon = 32
    uwsCanResize = 64
    uwsDefault = uwsHasTitleBar Or uwsHasCloseButton
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Routine to set a userform's window style, called from
'           the Userform_Initialize event
'
' Arguments:    frmForm     The userform to change the style for
'               lStyles     An enumeration value of the style(s) to set.
'                           The enumeration values can be added together
'                           to set multiple styles in one call.
'               sIconPath   If the uwsHasIcon style is set, this is the path
'                           to the icon file to use for the form
'
' Date          Developer       Action
' --------------------------------------------------------------
' 05 Jun 04     Stephen Bullen  Created
'
Sub SetUserformAppearance(ByRef frmForm As Object, ByVal lStyles As UserformWindowStyles, Optional ByVal sIconPath As String)

    Dim sCaption As String
    Dim hwnd As Long
    Dim lStyle As Long
    Dim hMenu As Long

    'Find the window handle of the form
    sCaption = frmForm.Caption
    frmForm.Caption = "FindThis" & Rnd
    hwnd = FindOurWindow("ThunderDFrame", frmForm.Caption)
    frmForm.Caption = sCaption

    'If we want a small title bar, we can't have an icon, max or min buttons as well
    If lStyles And uwsHasSmallTitleBar Then
        lStyles = lStyles And Not (uwsHasMaxButton Or uwsHasMinButton Or uwsHasIcon)
    End If

    'Get the normal windows style bits
    lStyle = GetWindowLong(hwnd, GWL_STYLE)

    'Update the normal style bits appropriately

    'If we want and icon or Max, Min or Close buttons, we have to have a system menu
    ModifyStyles lStyle, lStyles, uwsHasIcon Or uwsHasMaxButton Or uwsHasMinButton Or uwsHasCloseButton, WS_SYSMENU

    'Most things need a title bar!
    ModifyStyles lStyle, lStyles, uwsHasIcon Or uwsHasMaxButton Or uwsHasMinButton Or uwsHasCloseButton Or uwsHasTitleBar Or uwsHasSmallTitleBar, WS_CAPTION

    ModifyStyles lStyle, lStyles, uwsHasMaxButton, WS_MAXIMIZEBOX
    ModifyStyles lStyle, lStyles, uwsHasMinButton, WS_MINIMIZEBOX
    ModifyStyles lStyle, lStyles, uwsCanResize, WS_THICKFRAME

    'Update the window with the normal style bits
    SetWindowLong hwnd, GWL_STYLE, lStyle

    'Get the extended style bits
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)

    'Modify them appropriately
    ModifyStyles lStyle, lStyles, uwsHasSmallTitleBar, WS_EX_TOOLWINDOW

    'The icon is different to the rest - we set a bit to turn it off, not on!
    If lStyles And uwsHasIcon Then
        lStyle = lStyle And Not WS_EX_DLGMODALFRAME

        'Set the icon, if given
        If Len(sIconPath) > 0 Then
            SetIcon hwnd, sIconPath
        End If
    Else
        lStyle = lStyle Or WS_EX_DLGMODALFRAME
    End If

    'Update the window with the extended style bits
    SetWindowLong hwnd, GWL_EXSTYLE, lStyle

    'The Close button is handled by removing it from the  control menu, not through a window style bit
    If lStyles And uwsHasCloseButton Then
        'We want it, so reset the control menu
        hMenu = GetSystemMenu(hwnd, 1)
    Else
        'We don't want it, so delete it from the control menu
        hMenu = GetSystemMenu(hwnd, 0)
        DeleteMenu hMenu, SC_CLOSE, 0&
    End If

    'Refresh the window with the changes
    DrawMenuBar hwnd

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Helper routine to check if one of our style bits is set
'           and set/clear the corresponding Windows style bit
'
' Date          Developer       Action
' --------------------------------------------------------------
' 05 Jun 04     Stephen Bullen  Created
'
Private Sub ModifyStyles(ByRef lFormStyle As Long, ByVal lStyleSet As Long, ByVal lChoice As UserformWindowStyles, ByVal lWS_Style As Long)

    If lStyleSet And lChoice Then
        lFormStyle = lFormStyle Or lWS_Style
    Else
        lFormStyle = lFormStyle And Not lWS_Style
    End If

End Sub

