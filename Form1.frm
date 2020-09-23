VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Close Program"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   3480
   End
   Begin VB.PictureBox Work 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2880
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Display name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full name"
         Object.Width           =   8148
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usage"
         Object.Width           =   1878
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CloseMe 
      Caption         =   "&Close me"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   3600
      Width           =   1830
   End
   Begin VB.CommandButton Killit 
      Caption         =   "&End Task"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu terminate 
         Caption         =   "Terminate now"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim CurrentClicked As ListItem
Dim NumOfProcess As Long
Dim objActiveProcess As GetPro

Private Sub CloseMe_Click()
Hideme
End Sub

Private Sub Form_Load()
Set objActiveProcess = New GetPro
Hide1
Me.Hide
DoEvents
RegKey Me.Hwnd
'If Format(Command, "<") <> "background" Then
'Disable1
'Form2.ShowMe
'ShowMe
'Else
'End If
End Sub

Function MakeAList()
On Error Resume Next
Dim OneChecked As Boolean
Dim Ret

Killit.Enabled = False

Set ListView1.SmallIcons = Nothing

ListView1.ListItems.Clear
ImageList1.ListImages.Clear

NumOfProcess = objActiveProcess.GetActiveProcess

ImageList1.ListImages.Add , , Work.Image

Set ListView1.SmallIcons = ImageList1

For i = 1 To NumOfProcess
Ret = GetIcon(objActiveProcess.szExeFile(i), Work)
ImageList1.ListImages.Add i + 1, , Work.Image

ListView1.ListItems.Add , , Ret, , i + 1
ListView1.Refresh
ListView1.ListItems.Item(i).ListSubItems.Add , , objActiveProcess.szExeFile(i)
ListView1.ListItems.Item(i).ListSubItems.Add , , objActiveProcess.Usage(i) & "%"

If fEnumWindows(objActiveProcess.th32ProcessID(i)) = 0 Then
ListView1.ListItems.Item(i).ListSubItems.Add , , "Working..."
Else
ListView1.ListItems.Item(i).ListSubItems.Add , , "Stucked!"
ListView1.ListItems.Item(i).ForeColor = vbRed
ListView1.ListItems.Item(i).ListSubItems.Item(1).ForeColor = vbRed
ListView1.ListItems.Item(i).ListSubItems.Item(2).ForeColor = vbRed
OneChecked = True
ListView1.ListItems.Item(i).Checked = True
End If

Next i

If Not OneChecked Then ListView1.ListItems.Item(2).Checked = True
End Function

Private Sub Form_Unload(Cancel As Integer)
Set objActiveProcess = Nothing
UnRegKey Me.Hwnd
End
End Sub

Private Sub Killit_Click()
Dim lProcess As Long
Dim lReturn As Long
Dim Ret As VbMsgBoxResult

Ret = MsgBox("Are you sure you want to terminate it?", vbExclamation + vbYesNo, "Warning")

If Ret = vbYes Then
lProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, objActiveProcess.th32ProcessID(CurrentClicked.Index))
lReturn = TerminateProcess(lProcess, 0&)
Sleep 1000
MakeAList
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Killit.Enabled = True
Set CurrentClicked = Item
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Menu1
End If
End Sub

Private Sub openfolder_Click() ' i think this is a hard way but it works ;-)
Dim Ret1 As Long
Dim Ret2 As Long
Dim Ret3 As String

Do
Ret1 = InStr(Ret2 + 1, objActiveProcess.szExeFile(CurrentClicked.Index), "\", vbTextCompare)
If Ret1 = 0 Then
Ret3 = Left(objActiveProcess.szExeFile(CurrentClicked.Index), Ret2)
Exit Do
Else
Ret2 = Ret1
End If
Loop

Shell "rundll32.exe url.dll,FileProtocolHandler " & Ret3, vbNormalFocus
End Sub

Private Sub properties_Click()
ShowProperties objActiveProcess.szExeFile(CurrentClicked.Index), Me.Hwnd
End Sub

Private Sub terminate_Click()
Killit_Click
End Sub

Private Sub Timer1_Timer()
If GetKeyState(vbKeyControl) < -10 And GetKeyState(18) < -10 And GetKeyState(vbKeyAdd) < -10 Then
MsgBox ""
ShowMe
End If
End Sub

Function ShowMe()
Me.Show
Me.ZOrder 0
MakeAList
OnTop Me.Hwnd
End Function

Function Hideme()
Me.Hide
ListView1.ListItems.Clear ' to clear up some mem
Form2.Hide
OffTop Me.Hwnd
Enable1
End Function
