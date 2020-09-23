Attribute VB_Name = "Declares"
Option Explicit

Private PID As Long
Public IsResond As String

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const WM_NULL = &H0
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, pdwResult As Long) As Long

Private Function fEnumWindowsCallBack(ByVal hWnd As Long, ByVal lpData As Long) As Long
Dim lThreadId  As Long
Dim lProcessId As Long

fEnumWindowsCallBack = 1
lThreadId = GetWindowThreadProcessId(hWnd, lProcessId)

If lProcessId = PID Then
    Call strCheck(hWnd)
    fEnumWindowsCallBack = 0
End If

End Function

Public Function fEnumWindows(clsPID As Long) As Boolean
Dim hWnd As Long

PID = clsPID

Call EnumWindows(AddressOf fEnumWindowsCallBack, hWnd)
End Function
    

Private Function strCheck(ByVal lhwnd As Long)
Dim lResult As Long
Dim lReturn As Long
Dim strRunning As String

If lhwnd = 0 Then Exit Function

lReturn = SendMessageTimeout(lhwnd, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG And SMTO_BLOCK, 1000, lResult)

If lReturn Then
    IsResond = "Responding"
Else
    IsResond = "Not Responding"
End If
End Function
