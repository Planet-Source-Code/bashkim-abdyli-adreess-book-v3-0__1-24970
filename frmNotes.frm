VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNotes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   2160
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   310
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   850
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   310
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   850
   End
   Begin VB.Frame NotesFrame 
      Height          =   2020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6330
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00C0FFFF&
         Height          =   1815
         Left            =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   120
         Width           =   6220
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   2550
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7011
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2118
            MinWidth        =   2118
            TextSave        =   "5/6/00"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "1:19 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   6000
      Picture         =   "frmNotes.frx":014A
      Top             =   2200
      Width           =   240
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
'================================================================
'                     TO GOD BE THE GLORY
'================================================================
'================================================================
'*** For any Questions or Comments concerning this program    ***
'*** homepage : http://www.omarswan.cjb.net                   ***
'*** Email    : omarswan@yahoo.com                            ***
'================================================================
'================================================================
'* Deducated to SmileyOrange -> http://www.smileyorange.cjb.net *
'================================================================
'================================================================

Option Explicit



Private Sub btnCancel_Click()
   Unload Me
End Sub

Private Sub btnOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Select Case frmMain.Edit_Mode
      Case True:
         btnOK.Enabled = True
         btnCancel.Enabled = True
         txtNotes.Locked = False

      Case False:
         btnOK.Enabled = True
         btnCancel.Enabled = False
         txtNotes.Locked = True
   End Select
End Sub

Private Sub Timer1_Timer()
   StatusBar1.Panels(3).Text = Format$(Now, "H:MM:SS AMPM")
   StatusBar1.Panels(2).Text = Format$(Now, "MM-DD-YYYY")
End Sub
