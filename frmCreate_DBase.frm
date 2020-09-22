VERSION 5.00
Begin VB.Form frmCreate_DBase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Family Address Book v3.0 Database Creator ..."
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmCreate_DBase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   10
      TabIndex        =   0
      Top             =   0
      Width           =   5160
      Begin VB.CommandButton btnExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton btnCreate 
         Caption         =   "&Create Database"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCreate_DBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
'================================================================
'***                  TO GOD BE THE GLORY                     ***
'================================================================
'================================================================
'*** For any Questions or Comments concerning this program    ***
'*** Homepage : http://www.omarswan.cjb.net                   ***
'*** Email    : omarswan@yahoo.com                            ***
'*** AOL      : smileyomar  or omarsmiley                     ***
'================================================================
'================================================================
'* Deducated to SmileyOrange -> http://www.smileyorange.cjb.net *
'================================================================
'================================================================

Option Explicit

Public MYDBASE As clsCreateDbase


Private Sub btnCreate_Click()
   Set MYDBASE = New clsCreateDbase
   If MYDBASE.RECREATE_DATABASE = True Then
      If MYDBASE.OPEN_CONNECTION = True Then
         If MYDBASE.ADD_NEW_USER("Admin", "Admin", "Your Full Name", "Administrator") = True Then
            MsgBox "Database Created", vbInformation + vbOKOnly
            Call btnExit_Click
         Else
            MsgBox "ERROR - While Creating Database", vbCritical + vbOKOnly
         End If
      Else
         MsgBox "ERROR - Unable to open Database Connection", vbCritical + vbOKOnly
      End If
   Else
      MsgBox "Unable to properly create the database", vbCritical + vbOKOnly
   End If
End Sub

Private Sub btnExit_Click()
   Set MYDBASE = Nothing
   Unload Me
   End
End Sub

Private Sub Form_Load()
   '
End Sub
