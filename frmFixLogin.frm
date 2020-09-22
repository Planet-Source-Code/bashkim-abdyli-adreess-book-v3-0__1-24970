VERSION 5.00
Begin VB.Form frmFixLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Family Address Book v3.0 Database Fixer"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmFixLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   20
      TabIndex        =   4
      Top             =   0
      Width           =   5400
      Begin VB.CommandButton btnExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4080
         TabIndex        =   3
         Top             =   830
         Width           =   1095
      End
      Begin VB.CommandButton btnLogin 
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   900
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Login Name"
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   440
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   ":: Family Address Book v3.0 Database Fixer ::"
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
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   4095
   End
End
Attribute VB_Name = "frmFixLogin"
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


Public Sub btnExit_Click()
   Set frmFixer = Nothing
   Set frmFixLogin = Nothing
   Set TmpRecordset = Nothing
   Set PUBLIC_DATABASE = Nothing
   Unload Me
   End
End Sub

Private Sub btnLogin_Click()
   If Len(TextBox(0).Text) < 1 Then
      MsgBox "Please enter your Login Name.", vbInformation + vbOKOnly
      TextBox(0).SetFocus
      Exit Sub
   End If

   If Len(TextBox(1).Text) < 1 Then
      MsgBox "Please enter your Password.", vbInformation + vbOKOnly
      TextBox(1).SetFocus
      Exit Sub
   End If

   TextBox(0).Text = modFix.Apostrophe(TextBox(0).Text)
   TextBox(1).Text = modFix.Apostrophe(TextBox(1).Text)

   TmpSQL = ""
   TmpSQL = "SELECT USER_LOGIN_NAME,USER_PASSWORD FROM " & USER_TABLENAME
   TmpSQL = TmpSQL & " WHERE USER_LOGIN_NAME = '" & TextBox(0).Text & "'"
   TmpSQL = TmpSQL & " and USER_PASSWORD = '" & TextBox(1).Text & "'"

   Set TmpRecordset = New ADODB.Recordset
   TmpRecordset.Open TmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   
   If TmpRecordset.RecordCount > 0 Then
      Set TmpRecordset = Nothing
      Load frmFixer
      frmFixer.Show
     Else
      Set TmpRecordset = Nothing
      MsgBox "Invalid Login Name or Password.", vbInformation + vbOKOnly
   End If
End Sub

Private Sub Form_Load()
   Call modFix.Main
End Sub

Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
   'If The Enter Key Is Pressed
   If KeyAscii = 13 Then
      Call btnLogin_Click
   End If
End Sub
