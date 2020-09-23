Attribute VB_Name = "Hotkeys"
Option Explicit

Public Declare Function RegisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = -4

Public Const MOD_CTRL = &H2
Public Const MOD_SHFT = &H4
Public Const MOD_ALT = &H1

Public Const VK_ADD = &H6B

Public glWinRet As Long

Public Function CallbackMsgs(ByVal wHwnd As Long, ByVal wmsg As Long, ByVal wp_id As Long, ByVal lp_id As Long) As Long
If wmsg = WM_HOTKEY Then
Call DoFunctions(wp_id)
CallbackMsgs = 1
Exit Function
End If
CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wmsg, wp_id, lp_id)
End Function

Public Sub DoFunctions(ByVal vKeyID As Byte)
Select Case vKeyID
Case 0
Disable1
Form2.ShowMe
Form1.ShowMe
End Select
End Sub

Function RegKey(Hwnd As Long)
Dim Ret As Boolean

Ret = RegisterHotKey(Hwnd, 0, MOD_CTRL, VK_ADD)
If Ret = False Then
MsgBox "Can not register the hotkey CTRL and + this key is already registered by some other running applications.", vbCritical + vbOKOnly, "Error"
End
End If
glWinRet = SetWindowLong(Hwnd, GWL_WNDPROC, AddressOf CallbackMsgs)
End Function

Function UnRegKey(Hwnd As Long)
UnregisterHotKey Hwnd, 0
End Function

