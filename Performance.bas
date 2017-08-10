Attribute VB_Name = "Performance"
Dim prevCalc, prevEvents, prevScreen, prevPageBreaks
Dim execTimer As HighResPerformanceTimer, currTime As Double
#If VBA7 Then
    Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal LpThreadAttributes As Long, ByVal DwStackSize As Long, ByVal LpStartAddress As Long, ByVal LpParameter As Long, ByVal dwCreationFlags As Long, ByRef LpThreadld As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal HANDLE As Long) As Long
#Else
    Private Declare Function CreateThread Lib "kernel32" (ByVal LpThreadAttributes As Long, ByVal DwStackSize As Long, ByVal LpStartAddress As Long, ByVal LpParameter As Long, ByVal dwCreationFlags As Long, ByRef LpThreadld As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal HANDLE As Long) As Long
#End If

'***********************OPTIMIZE VBA CODE**********************************
Sub OptimizeOn()
    'Optimize VBA exectution
    prevCalc = Application.Calculation: Application.Calculation = xlCalculationManual
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    prevScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    prevPageBreaks = ActiveSheet.DisplayPageBreaks: ActiveSheet.DisplayPageBreaks = False
End Sub
Sub OptimizeOff()
    'Turn off VBA optimization
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    ActiveSheet.DisplayPageBreaks = prevPageBreaks
End Sub
'*********************MEASURE EXECUTION TIME*******************************
'*****High Res (milliseconds)*****
Sub StartHighResPerformanceTimer()
    Set execTimer = New HighResPerformanceTimer
    execTimer.StartCounter
End Sub
Function StopHighResPerformanceTimer() As Double
    StopHighResPerformanceTimer = execTimer.TimeElapsed
    Set execTimer = Nothing
End Function
'*****Low Res (seconds)*****
Sub StartLowResPerformanceTimer()
    currTime = Timer
End Sub
Function StopLowResPerformanceTimer() As Double
    StopLowResPerformanceTimer = Timer - currTime
End Function
'*********************CREATE ADDITIONAL VBA THREAD*************************
'WARNING: USE AT OWN RISK!
'COMMENT: ONLY 1 ADDITIONAL THREAD SHOULD RUN AT ANY GIVEN MOMENT! OTHERWISE VBA PROJECT WILL MOST DEFINITELY CRASH
Function RunThread(subAddress) As Long
'Runs thread provided by Sub address. E.g. RunThread(AddressOf MySub)
'Returns threadId
    RunThread = CreateThread(nil, 0, subAddress, nil, 0, nil)
End Function
Sub KillThread(threadId As Long)
'Be sure to run this after / before the VBA Thread has completed running
    CloseHandle threadId
End Sub
