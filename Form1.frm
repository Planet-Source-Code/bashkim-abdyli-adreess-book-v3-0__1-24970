VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReminder 
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameAddress 
      Height          =   4095
      Left            =   10
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   45
         TabIndex        =   1
         Top             =   120
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5318
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  frameReminder.Visible = False
  TabStrip1.Tabs(1).Selected = True
  TabStrip1.Refresh
  TabStrip1.Tabs (1)
End Sub

Private Sub TabStrip1_Click()
  If TabStrip1.Tabs(1).Selected = True Then
     frameReminder.Visible = True
    Else
     frameReminder.Visible = False
  End If
End Sub
