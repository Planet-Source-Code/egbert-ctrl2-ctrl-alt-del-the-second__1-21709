Attribute VB_Name = "Security"
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function SystemParametersInfo2 Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
#If Win32 Then
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
Public Declare Sub SetWindowPos Lib "User" (ByVal Hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#End If

Public Const SPI_SCREENSAVERRUNNING = 97
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Sub Hide1()
Dim pid As Long
Dim reserv As Long
pid = GetCurrentProcessId()
regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub

Public Sub Disable1()
Dim syssend As Long
syssend& = SystemParametersInfo2(SPI_SCREENSAVERRUNNING, True, False, 0)
End Sub

Public Sub Enable1()
Dim syssend As Long
syssend& = SystemParametersInfo2(SPI_SCREENSAVERRUNNING, False, True, 0)
End Sub

Function OnTop(Hwnd As Long)
#If Win32 Then
SetWindowPos Hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
#Else
SetWindowPos Hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
#End If
End Function

Function OffTop(Hwnd As Long)
#If Win32 Then
SetWindowPos Hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
#Else
SetWindowPos Hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
#End If
End Function


