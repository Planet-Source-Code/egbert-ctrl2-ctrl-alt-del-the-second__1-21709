VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Hyper Protection"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9735
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ShowMe()
Me.Cls
BitBlt Me.hDC, 0, 0, Screen.Width, Screen.Height, GetDC(0), 0, 0, vbSrcCopy
Me.Show
End Function

