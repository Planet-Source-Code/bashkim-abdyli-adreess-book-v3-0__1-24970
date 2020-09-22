VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::  User(s) Profile(s) "
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   20
      ScaleHeight     =   255
      ScaleWidth      =   8100
      TabIndex        =   27
      Top             =   0
      Width           =   8100
      Begin Family_v3.Label3D Label3D2 
         Height          =   255
         Left            =   45
         TabIndex        =   28
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " :: Profile ..."
      End
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Height          =   310
      Left            =   7200
      TabIndex        =   11
      Top             =   4050
      Width           =   850
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Edit"
      Height          =   310
      Left            =   6240
      TabIndex        =   10
      Top             =   4050
      Width           =   855
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "&Add"
      Height          =   310
      Left            =   5280
      TabIndex        =   9
      Top             =   4050
      Width           =   850
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   2960
      TabIndex        =   17
      Top             =   180
      Width           =   5160
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   0
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "Full Name Textbox"
         Top             =   240
         Width           =   3495
      End
      Begin VB.ComboBox ComboBox 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   1
         Left            =   1550
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "&Cancel"
         Height          =   310
         Left            =   4180
         TabIndex        =   13
         Top             =   1320
         Width           =   850
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "&Save"
         Height          =   310
         Left            =   3200
         TabIndex        =   12
         Top             =   1320
         Width           =   850
      End
      Begin VB.ComboBox ComboBox 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   6
         Left            =   1550
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2860
         Width           =   3495
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   5
         Left            =   1550
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2470
         Width           =   3495
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   4
         Left            =   1550
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2100
         Width           =   3495
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   3
         Left            =   1550
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1720
         Width           =   3495
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   2
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   2
         Tag             =   "Password Textbox"
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   1
         Tag             =   "UserName Textbox"
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Locked"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   7
         Left            =   100
         TabIndex        =   25
         Top             =   3320
         Width           =   600
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Date Added"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   6
         Left            =   100
         TabIndex        =   24
         Top             =   2960
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "SMTP Server"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   5
         Left            =   105
         TabIndex        =   23
         Top             =   2570
         Width           =   1125
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Homepage (URL)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   4
         Left            =   100
         TabIndex        =   22
         Top             =   2180
         Width           =   1380
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Em@il Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   3
         Left            =   100
         TabIndex        =   21
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Access Level"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   2
         Left            =   100
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   1
         Left            =   100
         TabIndex        =   19
         Top             =   1070
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4560
         Picture         =   "frmUsers.frx":0ABA
         Top             =   3210
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   0
         Left            =   100
         TabIndex        =   18
         Top             =   700
         Width           =   900
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   310
      Left            =   120
      TabIndex        =   14
      Top             =   4050
      Width           =   850
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   20
      TabIndex        =   16
      Top             =   180
      Width           =   2895
      Begin MSComctlLib.TreeView tvUsers 
         Height          =   3500
         Left            =   80
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   150
         Width           =   2750
         _ExtentX        =   4868
         _ExtentY        =   6165
         _Version        =   393217
         Indentation     =   706
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   29
      Top             =   4440
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9155
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "7/13/2001"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "11:34 AM"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public ChangesMade      As Boolean
Public CurrentState     As String
Public ChangeOwnInfo   As Boolean

'Temporary Storage For UserName
Private tmpUserName As String

'USER_PROFILE Class
Public WithEvents MyUser As clsUser
Attribute MyUser.VB_VarHelpID = -1


'================================================================
'================================================================
Private Sub btnAdd_Click()
   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      CurrentState = "Add"
      Call Setup_TextBox(Me, True, False)
      Call Setup_ComboBox(Me, False, False)
      Call Change_Button(False, False, False, False, True, True)
      TextBox(6).Text = Format$(Now, "Long Date")
      TextBox(6).Locked = True
      'Disable The Treeview
      tvUsers.Enabled = False
      'Setfocus
      TextBox(0).SetFocus
   End If
End Sub
'================================================================


'================================================================
'================================================================
Private Sub btnCancel_Click()
   CurrentState = ""
   tmpUserName = ""

   Call Setup_TextBox(Me, True, True)
   Call Setup_ComboBox(Me, True, False)

   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      Call Change_Button(False, False, True, True, False, False)
      tvUsers.Enabled = True
      Me.Load_UserList (ComboBox(0).Text)
   ElseIf CURRENT_USER.ACCESS_LEVEL = User Then
      tvUsers.Enabled = True
      'Load The List
      Call Me.Load_UserList("User")
      tvUsers.Nodes("User").Expanded = True
      Call tvUsers_NodeClick(tvUsers.Nodes("User_" & CURRENT_USER.LOGIN_NAME))
      tvUsers.Nodes("User_" & CURRENT_USER.LOGIN_NAME).Selected = True
      'Lock The Treeview So That the "USER" cannot view other user's info
      tvUsers.Enabled = False
      Call Me.Change_Button(True, False, False, True, False, False)
   End If
End Sub
'================================================================


'================================================================
Private Sub btnDelete_Click()
   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      'Check if the user is trying to delete his/her own record
      If TextBox(1).Text = CURRENT_USER.LOGIN_NAME Then
         'Check it amount of administrators available is > 1
         If FAMILY.ADMINISTRATOR_COUNT > 1 Then
            MsgBox CURRENT_USER.FULL_NAME & ", please ask another Administrator to remove your account from the Database.", vbInformation + vbOKOnly
            Exit Sub
         Else
            MsgBox CURRENT_USER.FULL_NAME & ", you are the only Administrator. Your account cannot be removed.", vbInformation + vbOKOnly
            Exit Sub
         End If
      End If

      TmpMsgResult = MsgBox(CURRENT_USER.FULL_NAME & ", Are you sure that you want to remove " & TextBox(1).Text & "'s account from the Database.", vbQuestion + vbYesNo)
      If TmpMsgResult = vbYes Then
         Debug.Print "Removing " & TextBox(1).Text & " from the Database"
         'Call The Remove User Function
         Call FAMILY.DELETE_USER(Trim$(TextBox(1).Text))
         'Reload Userlist
         Me.Load_UserList (ComboBox(0).Text)
      End If
   End If
End Sub
'================================================================



'================================================================
Private Sub btnEdit_Click()
   CurrentState = "Edit"
   'Store The User Name
   tmpUserName = TextBox(1).Text

   'Disable the treview until finished editting
   tvUsers.Enabled = False

   'Check If the current user is trying to edit his/her own user Info
   If tmpUserName = CURRENT_USER.LOGIN_NAME Then
      ChangeOwnInfo = True

      'Check The AccessLevel
      If CURRENT_USER.ACCESS_LEVEL = Administrator Then
         ':: Administrator ::
         'Check The amount af Administrator
         If FAMILY.ADMINISTRATOR_COUNT > 1 Then
            Call Setup_ComboBox(Me, False, False)
            'Allow the Administrator to change his/her Accesslevel
            'Since there is more than one Administrators
            ComboBox(0).Locked = False
            ComboBox(1).Locked = False
            Call Setup_TextBox(Me, False, False)
            Call Change_Button(False, False, False, False, True, True)
            TextBox(6).Locked = True
         Else
            'Don't Allow the Administrator to change his/her
            'Accesslevel Since there is no more Administrators
            Call Setup_TextBox(Me, False, False)
            Call Setup_ComboBox(Me, False, False)
            ComboBox(0).Locked = True
            ComboBox(1).Locked = True
            Call Change_Button(False, False, False, False, True, True)
            TextBox(6).Locked = True
         End If

      ElseIf CURRENT_USER.ACCESS_LEVEL = User Then
         ':: User ::
         'A "User" is try to change his/her own information
         Call Setup_ComboBox(Me, True, False)
         Call Setup_TextBox(Me, False, False)
         ComboBox(0).Locked = True
         ComboBox(1).Locked = True
         Call Setup_TextBox(Me, False, False)
         Call Change_Button(False, False, False, False, True, True)
         TextBox(6).Locked = True
      End If

   Else
      ':: An Administrator Changing another User's Information ::
      'Make Sure that the current User is an Administrator
      Debug.Print "Changing another User's Info " & tmpUserName
      ChangeOwnInfo = False
      If CURRENT_USER.ACCESS_LEVEL = Administrator Then
         Call Setup_TextBox(Me, False, False)
         Call Change_Button(False, False, False, False, True, True)
         ComboBox(0).Locked = False
         ComboBox(1).Locked = False
         TextBox(6).Locked = True
      Else
         CurrentState = ""
         tmpUserName = ""
         Call Change_Button(False, False, False, False, True, True)
         MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", but you are not allowed to chage another User's information", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   'Setfocus
   TextBox(0).SetFocus
End Sub
'================================================================



'================================================================
'================================================================
Private Sub btnExit_Click()
   'Unload The Form
   Unload Me
End Sub
'================================================================


'================================================================
'================================================================
Private Sub btnSave_Click()
   Dim i As Byte
   Dim i2 As Byte
   Dim i3 As Byte
   On Error GoTo SAVE_ERROR
    
    
   TextBox(1).Text = Trim$(TextBox(1).Text)

   'Check For Spaces
   If InStr(1, TextBox(1).Text, " ") > 0 Then
      MsgBox CURRENT_USER.FULL_NAME & ", you have space in the " & TextBox(1).Tag, vbCritical + vbOKOnly
      Exit Sub
   End If

   'Check For Invalid Characters
   For i = LBound(IllegalChars) To UBound(IllegalChars)
      For i2 = 0 To 2

         If Len(TextBox(i2).Text) < MIN_USER_NAME_SIZE Then
            MsgBox "Hi " & CURRENT_USER.FULL_NAME & ", the " & TextBox(i2).Tag & " should be less than or equal to " & Str$(MAX_USER_NAME_SIZE) & "." & vbNewLine & _
                  "But greater than or equal to " & Str$(MIN_USER_NAME_SIZE) & ".", vbInformation + vbOKOnly
            Exit Sub
         End If

         For i3 = 1 To Len(TextBox(i2))
            If (i2 = 1) Or (i2 = 2) Then
               If Chr$(IllegalChars(i)) = Mid$(TextBox(i2), i3, 1) Then
                  MsgBox CURRENT_USER.FULL_NAME & ", you have an invalid charater [ " & Chr$(IllegalChars(i)) & " ] , in the " & TextBox(i2).Tag, vbCritical + vbOKOnly
                  Exit Sub
               End If
            End If
         Next i3

      Next i2
   Next i


   Select Case CurrentState
      Case "Add"   'Adding A New User
         'First Check If The User Name Exist
         If FAMILY.FORM_USERS.MyUser.USER_EXIST(TextBox(1).Text) Then
            MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", but that user name [ " & TextBox(1).Text & " ] already exist.", vbInformation + vbOKOnly
            Exit Sub
         End If

         Call FAMILY.ADD_NEW_USER( _
               Trim$(TextBox(1).Text), _
               Trim$(TextBox(2).Text), _
               Trim$(TextBox(0).Text), _
               ComboBox(0).Text, _
               Trim$(TextBox(3).Text), _
               Trim$(TextBox(4).Text), _
               Trim$(TextBox(5).Text))

      Case "Edit"   '- Edit User

         'Check IF The OLD User Name Was Changed
         If tmpUserName <> TextBox(1).Text Then
            'Since It Was Changed Check To See If The User Name Exist Already
            If FAMILY.FORM_USERS.MyUser.USER_EXIST(TextBox(1).Text) Then
               MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", but the user name [ " & TextBox(1).Text & " ] already exist.", vbInformation + vbOKOnly
               Exit Sub
            End If
         End If

         TmpString = ""
         TmpString = "UPDATE " & USER_TABLENAME & " SET " & _
               "FULL_NAME = '" & Apostrophe(ProperCase(Trim$(TextBox(0).Text))) & "', " & _
               "USER_LOGIN_NAME = '" & Apostrophe(Trim$(TextBox(1).Text)) & "', " & _
               "USER_PASSWORD = '" & Apostrophe(Trim$(TextBox(2).Text)) & "', " & _
               "USER_ACCESS_LEVEL = '" & ComboBox(0).Text & "', " & _
               "USER_EMAIL_ADDDRESS = '" & Apostrophe(Trim$(TextBox(3).Text)) & "', " & _
               "USER_HOMEPAGE_URL = '" & Apostrophe(Trim$(TextBox(4).Text)) & "', " & _
               "USER_SMTP_SERVER = '" & Apostrophe(Trim$(TextBox(5).Text)) & "' " & _
               "WHERE USER_LOGIN_NAME = '" & tmpUserName & "'"

         PUBLIC_DATABASE.CONNECTION.Execute TmpString
         DoEvents

         'Rename/Update the user's Data  - i.e. Contact, Reminders, etc..
         If FAMILY.RENAME_USER_DATA(tmpUserName, Apostrophe(Trim$(TextBox(1).Text))) = False Then
            Debug.Print "ERROR : Unable to Properly Update " & TextBox(1).Text & "'s Data."
            MsgBox "ERROR : Unable to Properly Update " & TextBox(1).Text & "'s Data."
            Exit Sub
         End If

         'Check if the user changed his/her own info
         If ChangeOwnInfo = True Then
            If FAMILY.RELOAD_CURRENT_USER(Apostrophe(Trim$(TextBox(1).Text))) = True Then
               ChangesMade = True
               Debug.Print CURRENT_USER.FULL_NAME & " has changed his/her own information"
            End If
         End If

         TmpString = ""
         tmpUserName = ""
         CurrentState = ""
   End Select


   TmpString = ""
   tmpUserName = ""
   CurrentState = ""

   Call Setup_TextBox(Me, True, True)
   Call Setup_ComboBox(Me, True, False)

   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      Call Change_Button(False, False, True, True, False, False)
      tvUsers.Enabled = True
      Me.Load_UserList (ComboBox(0).Text)
   ElseIf CURRENT_USER.ACCESS_LEVEL = User Then
      tvUsers.Enabled = True
      Call Me.Change_Button(False, False, False, True, False, False)
      'Load The List
      Call Me.Load_UserList("User")
      tvUsers.Nodes("User").Expanded = True
      Call tvUsers_NodeClick(tvUsers.Nodes("User_" & CURRENT_USER.LOGIN_NAME))
      tvUsers.Nodes("User_" & CURRENT_USER.LOGIN_NAME).Selected = True
      tvUsers.Enabled = False
   End If
   
   
SAVE_ERROR:
   If Err.Number <> 0 Then
      MsgBox "Error : frmUsers.btnSave " & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Err #" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Sub
'================================================================



'================================================================
'================================================================
Private Sub Form_Load()
   Call modPublic.RemoveMenus(Me)
   CurrentState = ""
   ChangesMade = False

   'Setup The Text Box and Combobox
   Call Setup_TextBox(Me, True, True)
   Call Setup_ComboBox(Me, True, False)

   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      Call Me.Change_Button(False, False, True, True, False, False)
      'Load The List
      Call Me.Load_UserList("Administrator")

      tvUsers.Nodes("Administrator").Expanded = True
      Call tvUsers_NodeClick(tvUsers.Nodes("Administrator_" & CURRENT_USER.LOGIN_NAME))
      tvUsers.Nodes("Administrator_" & CURRENT_USER.LOGIN_NAME).Selected = True
      tvUsers.Enabled = True
   ElseIf CURRENT_USER.ACCESS_LEVEL = User Then
      Call Me.Change_Button(False, False, False, True, False, False)
      'Load The List
      Call Me.Load_UserList("User")
      tvUsers.Nodes("User").Expanded = True
      Call tvUsers_NodeClick(tvUsers.Nodes("User_" & CURRENT_USER.LOGIN_NAME))
      tvUsers.Nodes("User_" & CURRENT_USER.LOGIN_NAME).Selected = True
      tvUsers.Enabled = False
   Else
      MsgBox "Invalid Access Level", vbCritical + vbOKOnly
      Unload Me
   End If

   TextBox(0).MaxLength = MAX_USER_NAME_SIZE
   TextBox(1).MaxLength = MAX_USER_NAME_SIZE
   TextBox(2).MaxLength = MAX_USER_NAME_SIZE
End Sub
'================================================================


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If FAMILY.FORM_USERS.CurrentState <> "" Then
      Cancel = True
   Else
      'Close The USER_PROFILE_RECORDSET
      Set FAMILY.FORM_USERS.MyUser.USER_PROFILE = Nothing

      If ChangesMade = True Then
         Call FAMILY.SETUP_FORM_MAIN
      End If

      'SetUp FrmMain Texboxes and Comboboxes
      Call Setup_TextBox(FAMILY.FORM_MAIN, True, True)
      Call Setup_ComboBox(FAMILY.FORM_MAIN, True, True)
      Call FAMILY.FORM_MAIN.Change_Button(False, False, True, True, False, False)
   End If
End Sub

'================================================================
'================================================================
Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub
'================================================================


'================================================================
'================================================================
Public Function Load_UserList(ByVal Expand_This_Node As String) As Boolean

   Set MyUser = New clsUser
   Set MyUser.USER_PROFILE = New ADODB.Recordset

   Call modTreeview.LOAD_USERS_TREEVIEW(Me, tvUsers, FAMILY.FORM_MAIN.ImageList1)
   DoEvents
   ' MsgBox FAMILY.FORM_USERS.MyUser.USER_EXIST(CURRENT_USER.LOGIN_NAME)

   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      'Change the color of node that represents the USER currently logged in
      tvUsers.Nodes("Administrator_" & CURRENT_USER.LOGIN_NAME).ForeColor = vbBlue
      tvUsers.Nodes("Administrator_" & CURRENT_USER.LOGIN_NAME).BackColor = vbGreen
   ElseIf CURRENT_USER.ACCESS_LEVEL = User Then
      'Change the color of node that represents the USER currently logged in
      tvUsers.Nodes("User_" & CURRENT_USER.LOGIN_NAME).ForeColor = vbBlue
      tvUsers.Nodes("User_" & CURRENT_USER.LOGIN_NAME).BackColor = vbGreen
   End If

   'Clear The Combobox
   ComboBox(0).Clear
   ComboBox(0).AddItem "Administrator"
   ComboBox(0).AddItem "User"
   ComboBox(0).ListIndex = 0
   ComboBox(1).Clear
   ComboBox(1).AddItem "False"
   ComboBox(1).AddItem "True"
   ComboBox(1).ListIndex = 0

   'Expand The node ("User" or "Administrator")
   tvUsers.Nodes(Expand_This_Node).Expanded = True

End Function
'================================================================




Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
   'Validate Character Entered
   If (Index = 1) Or (Index = 2) Then
      If modPublic.MyChar(Chr(KeyAscii), "-") = False Then
         If (KeyAscii <> 32) Or (KeyAscii <> 8) Then
            KeyAscii = 0
         End If
      End If
   End If
End Sub

'================================================================
'================================================================
Private Sub tvUsers_Collapse(ByVal Node As MSComctlLib.Node)
   'Clear The The Textboxes
   Call Setup_TextBox(Me, True, True)
   Call Setup_ComboBox(Me, True, False)

   If Node.Tag = "Parent" Then
      ComboBox(0).Text = Node.Text
   End If

   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      Call Change_Button(False, False, True, True, False, False)
   Else
      Call Change_Button(False, False, False, True, False, False)
   End If

End Sub
'================================================================



'================================================================
'================================================================
Private Sub tvUsers_Expand(ByVal Node As MSComctlLib.Node)
   'Clear The The Textboxes
   Call Setup_TextBox(Me, True, True)
   Call Setup_ComboBox(Me, True, False)

   If Node.Tag = "Parent" Then
      ComboBox(0).Text = Node.Text
   End If

   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      Call Change_Button(False, False, True, True, False, False)
   Else
      Call Change_Button(False, False, False, True, False, False)
   End If

End Sub
'================================================================



'================================================================
'================================================================
Private Sub tvUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Store the Value of the Button in the Tag For Later Use
   tvUsers.Tag = Button
End Sub
'================================================================



'================================================================
'================================================================
Private Sub tvUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim tvNode As Node

   'This is used to track the Node That Mouse is over
   Set tvNode = tvUsers.HitTest(X, Y)

   If tvNode Is Nothing Then
      StatusBar1.Panels(1).Text = ""
      Exit Sub
   Else
      StatusBar1.Panels(1).Text = tvNode.Text
   End If

End Sub
'================================================================


'================================================================
'================================================================
Private Sub tvUsers_NodeClick(ByVal Node As MSComctlLib.Node)
   Dim sPos           As Integer   'Seperator Position "_"
   Dim tmpUserName    As String   'Last Name
   Dim tmpAccessLevel As String   'Relation or Category
   Dim tmpSQL         As String   'SQL String

   Debug.Print "Node Name : " & Node.Key & "  - Node Tag " & Node.Tag

   Select Case Node.Tag
      Case "Child"
         tmpSQL = "SELECT FULL_NAME,USER_LOGIN_NAME,USER_PASSWORD,USER_EMAIL_ADDDRESS,USER_HOMEPAGE_URL,USER_SMTP_SERVER,DATE_ADDED,USER_ACCESS_LEVEL,USER_LOCKED FROM " & USER_TABLENAME & _
               " WHERE USER_LOGIN_NAME = '" & Node.Text & "'" & _
               " AND USER_ACCESS_LEVEL = '" & Node.Parent.Text & "'"

         Set MyUser.USER_PROFILE = New ADODB.Recordset
         MyUser.USER_PROFILE.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

         If CURRENT_USER.ACCESS_LEVEL = Administrator Then
            Call Change_Button(True, True, True, True, False, False)
         Else
            Call Change_Button(True, False, False, True, False, False)
         End If

         Call DISPLAY_CURRENT_RECORD(MyUser.USER_PROFILE)

      Case "Parent"
         'Clear The The Textboxes
         Call Setup_TextBox(Me, True, True)
         Call Setup_ComboBox(Me, True, False)
         ComboBox(0).Text = Node.Text

         If CURRENT_USER.ACCESS_LEVEL = Administrator Then
            Call Change_Button(False, False, True, True, False, False)
         Else
            Call Change_Button(False, False, False, True, False, False)
         End If


      Case "ROOT"
         'Clear The The Textboxes
         Call Setup_TextBox(Me, True, True)
         Call Setup_ComboBox(Me, True, False)

         If CURRENT_USER.ACCESS_LEVEL = Administrator Then
            Call Change_Button(False, False, True, True, False, False)
         Else
            Call Change_Button(False, False, False, True, False, False)
         End If
   End Select
End Sub
'================================================================


'================================================================
'================================================================
Public Sub Change_Button(ByVal EDIT_BUTTON As Boolean, _
                          ByVal DELETE_BUTTON As Boolean, _
                          ByVal ADD_BUTTON As Boolean, _
                          ByVal EXIT_BUTTON As Boolean, _
                          ByVal Save_Button As Boolean, _
                          ByVal Cancel_Button As Boolean)

   Me.btnEdit.Enabled = EDIT_BUTTON
   Me.btnAdd.Enabled = ADD_BUTTON
   Me.btnDelete.Enabled = DELETE_BUTTON
   Me.btnExit.Enabled = EXIT_BUTTON
   Me.btnSave.Enabled = Save_Button
   Me.btnCancel.Enabled = Cancel_Button
End Sub
'================================================================


'================================================================
'================================================================
Private Function FINDRecord(SQL_String As String) As Boolean

End Function
'================================================================



'=====================================================================
'Display The Current Record
'=====================================================================
Public Function DISPLAY_CURRENT_RECORD(ByVal tRecordSet As ADODB.Recordset) As Boolean
   If (Not tRecordSet.BOF) And (Not tRecordSet.EOF) Then
      TextBox(0).Text = tRecordSet.Fields("FULL_NAME")
      TextBox(1).Text = tRecordSet.Fields("USER_LOGIN_NAME")
      TextBox(2).Text = tRecordSet.Fields("USER_PASSWORD")
      TextBox(3).Text = tRecordSet.Fields("USER_EMAIL_ADDDRESS")
      TextBox(4).Text = tRecordSet.Fields("USER_HOMEPAGE_URL")
      TextBox(5).Text = tRecordSet.Fields("USER_SMTP_SERVER")
      TextBox(6).Text = Format$(tRecordSet.Fields("DATE_ADDED"), "Long Date")
      ComboBox(0).Text = tRecordSet.Fields("USER_ACCESS_LEVEL")
      ComboBox(1).Text = tRecordSet.Fields("USER_LOCKED")
      DISPLAY_CURRENT_RECORD = True
   Else
      DISPLAY_CURRENT_RECORD = False
   End If
End Function
'================================================================

