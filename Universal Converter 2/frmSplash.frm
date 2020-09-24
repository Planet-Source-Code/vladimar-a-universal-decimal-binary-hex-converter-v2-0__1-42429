VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   4185
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   7500
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3500
      Left            =   6960
      Top             =   3720
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_DblClick()
frmUniversalConverter2.Show
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Timer1_Timer()
frmUniversalConverter2.Show
Unload Me
End Sub
