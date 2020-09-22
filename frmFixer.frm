VERSION 5.00
Begin VB.Form frmFixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Family Address Book v3.0 Database Fixer"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmFixer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5040
      Begin VB.CommandButton btnExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton btnLogout 
         Caption         =   "&Logout"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton btnRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox UserList 
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   ":: USER'S LOGGED IN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1860
   End
End
Attribute VB_Name = "frmFixer"
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


Private Sub btnExit_Click()
   Set frmFixer = Nothing
   Set frmFixLogin = Nothing
   Set TmpRecordset = Nothing
   Set PUBLIC_DATABASE = Nothing
   Unload Me
   End
End Sub

Private Sub btnLogout_Click()
   If UserList.ListCount > 0 Then
      If UserList.Text <> "" Then
         If modFix.LOGOUT_USER(UserList.Text) = False Then
            MsgBox "Unable to properly logout " & UserList.Text, vbCritical + vbOKOnly
         End If
         Call btnRefresh_Click
      Else
         MsgBox "Please Select A USER TO Logout"
      End If
   Else
      btnLogout.Enabled = False
   End If
End Sub

Private Sub btnRefresh_Click()
   TmpSQL = ""
   Set TmpRecordset = New ADODB.Recordset
   TmpSQL = "SELECT USER_LOGIN_NAME, USER_LOCKED FROM " & USER_TABLENAME & _
         " WHERE USER_LOCKED = TRUE"
   TmpRecordset.Open TmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   UserList.Clear
   If TmpRecordset.RecordCount > 0 Then
      Do While Not TmpRecordset.EOF
         UserList.AddItem TmpRecordset.Fields("USER_LOGIN_NAME").Value
         TmpRecordset.MoveNext
      Loop
      UserList.ListIndex = 0
      btnLogout.Enabled = True
   Else
      btnLogout.Enabled = False
   End If
   TmpSQL = ""
   Set TmpRecordset = Nothing
End Sub

Private Sub Form_Load()
   Unload frmFixLogin
   Set frmFixLogin = Nothing
   Call btnRefresh_Click
End Sub
