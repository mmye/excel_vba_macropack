Attribute VB_Name = "MWindows"
'
' Description:  Contains API constants, variables, declarations and procedures
'               to demonstrate API routines related to the windows
'
' Authors:      Stephen Bullen, www.oaltd.co.uk
'               Rob Bovey, www.appspro.com
'

Option Explicit
Option Private Module

' **************************************************************
' Declarations for the ApphWnd and FindOurWindow example functions
' **************************************************************
'Get the handle to the desktop window
Private Declare Function GetDesktopWindow Lib "user32" () As Long

'Find a child window with a given class name and caption
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'Get the process ID of this instance of Excel
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'Get the ID of the process that the window belongs to
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long


' **************************************************************
' Declarations for the WorkbookWindowhWnd example function
' The WorkbookWindowhWnd procedure uses FindWindowEx, defined above
' **************************************************************


' **************************************************************
' Declarations for the SetNameDropdownWidth example procedure
' **************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants used in the SendMessage call
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const CB_SETDROPPEDWIDTH As Long = &H160&     'from winuser.h

''''''''''''''''''''''''''''''''''''''''''''''''''
' Function Declarations
' The SetNameDropdownWidth procedure also uses FindWindowEx, defined above
''''''''''''''''''''''''''''''''''''''''''''''''''
'Send a message to a window
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


' **************************************************************
' Declarations for the SetIcon example procedure
' **************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants used in the SendMessage call
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const WM_SETICON As Long = &H80

''''''''''''''''''''''''''''''''''''''''''''''''''
' Function Declarations
' The SetIcon procedure also uses SendMessage, defined above
''''''''''''''''''''''''''''''''''''''''''''''''''
'Get a handle to an icon from a file
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Foolproof way to find the main Excel window handle
'
' Arguments:    None
'
' Returns:      Long        The handle of Excel's main window
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
Function ApphWnd() As Long

    If Val(Application.version) >= 10 Then
        ApphWnd = Application.hwnd
    Else
        ApphWnd = FindOurWindow("XLMAIN", Application.Caption)
    End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Finds a top-level window of the given class and
'           caption that belongs to this instance of Excel,
'           by matching the process IDs
'
' Arguments:    sClass      The window class name to look for
'               sCaption    The window caption to look for
'
' Returns:      Long        The handle of Excel's main window
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
Function FindOurWindow(Optional ByVal sClass As String = vbNullString, Optional ByVal sCaption As String = vbNullString)

    Dim hWndDesktop As Long
    Dim hwnd As Long
    Dim hProcThis As Long
    Dim hProcWindow As Long

    'All top-level windows are children of the desktop,
    'so get that handle first
    hWndDesktop = GetDesktopWindow

    'Get the ID of this instance of Excel, to match
    hProcThis = GetCurrentProcessId

    Do
        'Find the next child window of the desktop that
        'matches the given window class and/or caption.
        'The first time in, hWnd will be zero, so we'll get
        'the first matching window.  Each call will pass the
        'handle of the window we found the last time, thereby
        'getting the next one (if any)
        hwnd = FindWindowEx(hWndDesktop, hwnd, sClass, sCaption)

        'Get the ID of the process that owns the window we found
        GetWindowThreadProcessId hwnd, hProcWindow

        'Loop until the window's process matches this process,
        'or we didn't find the window
    Loop Until hProcWindow = hProcThis Or hwnd = 0

    'Return the handle we found
    FindOurWindow = hwnd

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Find the handle of a given workbook's Window
'
' Arguments:    None
'
' Returns:      Long        The handle of the given Window
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
Function WorkbookWindowhWnd(ByRef wndWindow As Window) As Long

    Dim hWndExcel As Long
    Dim hWndDesk As Long

    'Get the main Excel window
    hWndExcel = ApphWnd

    'Find the desktop
    hWndDesk = FindWindowEx(hWndExcel, 0, "XLDESK", vbNullString)

    'Find the workbook window
    WorkbookWindowhWnd = FindWindowEx(hWndDesk, 0, "EXCEL7", wndWindow.Caption)

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Make the Name dropdown list 200 pixels wide
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
Sub SetNameDropdownWidth()

    Dim hWndExcel As Long
    Dim hWndFormulaBar As Long
    Dim hWndNameCombo As Long

    'Get the main Excel window
    hWndExcel = ApphWnd

    'Get the handle for the formula bar window
    hWndFormulaBar = FindWindowEx(hWndExcel, 0, "EXCEL;", vbNullString)

    'Get the handle for the Name combobox
    hWndNameCombo = FindWindowEx(hWndFormulaBar, 0, "combobox", vbNullString)

    'Set the dropdown list to be 200 pixels wide
    SendMessage hWndNameCombo, CB_SETDROPPEDWIDTH, 200, 0

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Set a window's icon
'           When workbook windows are maximised, Excel doesn't update
'           the icon (shown on the left of the menu bar) until the window
'           is minimised/restored.
'
' Arguments:    hwnd        The handle of the window to change the icon for
'               sIcon       The path of the icon file
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
Sub SetIcon(ByVal hwnd As Long, ByVal sIcon As String)

    Dim hIcon As Long

    'Get the icon handle
    hIcon = ExtractIcon(0, sIcon, 0)

    'Set the big (32x32) and small (16x16) icons
    SendMessage hwnd, WM_SETICON, True, hIcon
    SendMessage hwnd, WM_SETICON, False, hIcon

End Sub
