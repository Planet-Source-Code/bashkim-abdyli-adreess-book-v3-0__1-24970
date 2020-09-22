VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   2400
   End
   Begin VB.Image Image2 
      Height          =   2850
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   4005
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Event Loaded()
Public Event UnLoaded()

Private Sub Form_Load()
   Debug.Print "Splash Screen Loaded"
   RaiseEvent Loaded
   DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent UnLoaded
End Sub

Private Sub Timer1_Timer()
   Unload Me
   DoEvents
End Sub
