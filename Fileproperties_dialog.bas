Attribute VB_Name = "Fileproperties_dialog"
Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hWnd As Long
  lpVerb As String
  lpFile As String
  lpParameters As String
  lpDirectory As String
  nShow As Long
  hInstApp As Long
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Public Function ShowProperties(filename As String, OwnerhWnd As Long) As Long
Dim SEI As SHELLEXECUTEINFO, R As Long

With SEI
  .cbSize = Len(SEI)
  .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
  .hWnd = OwnerhWnd
  .lpVerb = "properties"
  .lpFile = filename
  .lpParameters = vbNullChar
  .lpDirectory = vbNullChar
  .nShow = 0
  .hInstApp = 0
  .lpIDList = 0
End With
R = ShellExecuteEX(SEI)
ShowProperties = SEI.hInstApp
End Function

