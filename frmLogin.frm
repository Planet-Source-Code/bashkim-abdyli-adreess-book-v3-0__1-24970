VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::  Family Address Book v3.0   (Login System)"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   5325
      TabIndex        =   8
      Top             =   0
      Width           =   5325
      Begin Family_v3.Label3D Label3D1 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " :: Please enter you Login Name and Password ..."
         BackColor       =   -2147483637
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   " Exit "
      Top             =   1680
      Width           =   970
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   5265
      Begin VB.TextBox Login_Box 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1500
         MaxLength       =   25
         TabIndex        =   1
         Top             =   740
         Width           =   2295
      End
      Begin VB.CommandButton btnClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4000
         TabIndex        =   3
         ToolTipText     =   " Clear Fields "
         Top             =   730
         Width           =   1000
      End
      Begin VB.CommandButton btnLogin 
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4000
         TabIndex        =   2
         ToolTipText     =   " Login "
         Top             =   350
         Width           =   1000
      End
      Begin VB.TextBox Login_Box 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1500
         MaxLength       =   25
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
      Begin Family_v3.Label3D Label3D1 
         Height          =   255
         Index           =   1
         Left            =   250
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   800
         Width           =   975
         _ExtentX        =   1720
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
         ForeColor1      =   -2147483634
         ForeColor2      =   16711680
         Caption         =   "Password"
         Phase           =   3
      End
      Begin Family_v3.Label3D Label3D1 
         Height          =   255
         Index           =   0
         Left            =   250
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   405
         Width           =   1095
         _ExtentX        =   1931
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
         ForeColor1      =   -2147483634
         ForeColor2      =   16711680
         Caption         =   "Login Name"
         Phase           =   3
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4921
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "7/13/2001"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "11:34 AM"
         EndProperty
      EndProperty
   End
   Begin Family_v3.Label3D Label3D1 
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   12648447
      ForeColor2      =   16711680
      Caption         =   "User : admin; pass : admin"
      Phase           =   1
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub btnClear_Click()
   Login_Box(0).Text = ""
   Login_Box(1).Text = ""
End Sub

Private Sub btnClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Clear Fields"
End Sub

Private Sub btnExit_Click()
   'Close Up and ShutDown
   Unload Me
   Call modPublic.ShutDown
   End
End Sub

Private Sub btnExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Exit"
End Sub

Private Sub btnLogin_Click()
   Dim i         As Byte
   Dim i2        As Byte
   Dim i3        As Byte
   On Error GoTo LOGIN_ERROR

   Debug.Print Me.Name & " : btnLogin_Click"
   'Check for illegal characters
   For i = LBound(IllegalChars) To UBound(IllegalChars)
      For i2 = 0 To 1
         For i3 = 1 To Len(Login_Box(i2))
            If Chr$(IllegalChars(i)) = Mid$(Login_Box(i2), i3, 1) Then
               Debug.Print IllegalChars(i)
               MsgBox "You have an invalid charater [ " & Chr$(IllegalChars(i)) & " ] , in the " & Login_Box(i2).Tag, vbCritical + vbOKOnly
               Exit Sub
            End If
         Next i3
      Next i2
   Next i

   tmpSQL = ""
   tmpSQL = "SELECT USER_LOGIN_NAME,USER_PASSWORD,FULL_NAME,USER_EMAIL_ADDDRESS,USER_SMTP_SERVER,USER_ACCESS_LEVEL,USER_LOCKED FROM " & USER_TABLENAME
   tmpSQL = tmpSQL & " WHERE USER_LOGIN_NAME = '" & Apostrophe(Login_Box(0).Text) & "'"
   tmpSQL = tmpSQL & " and USER_PASSWORD = '" & Apostrophe(Login_Box(1).Text) & "'"

   Set FAMILY.FORM_USERS.MyUser.USER_PROFILE = New ADODB.Recordset
   FAMILY.FORM_USERS.MyUser.USER_PROFILE.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   FAMILY.FORM_USERS.MyUser.USER_PROFILE.Requery

   If FAMILY.FORM_USERS.MyUser.USER_PROFILE.RecordCount > 0 Then

      'Check If The User is already logged in
'      If FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_LOCKED") = False Then
         CURRENT_USER.LOGIN_NAME = FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_LOGIN_NAME")
         CURRENT_USER.FULL_NAME = ProperCase(FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("FULL_NAME"))
         CURRENT_USER.EMAIL_ADDRESS = FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_EMAIL_ADDDRESS")
         CURRENT_USER.PASSWORD = FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_PASSWORD")
         CURRENT_USER.SMTP_SERVER = FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_SMTP_SERVER")

         StatusBar1.Panels(1).Text = "Login Accepted"
         If FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_ACCESS_LEVEL") = "Administrator" Then
            CURRENT_USER.ACCESS_LEVEL = Administrator
         Else
            CURRENT_USER.ACCESS_LEVEL = User
         End If

         'Login The User
         FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_LOCKED") = True
         FAMILY.FORM_USERS.MyUser.USER_PROFILE.Update
         DoEvents
         Set FAMILY.FORM_USERS.MyUser.USER_PROFILE = Nothing

         StatusBar1.Panels(1).Text = "Loading The Database..."

         Debug.Print "Access Granted to : " & CURRENT_USER.FULL_NAME & ". Now Setting up Form; Main"
         'Update the INI File
         WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "Last-Logged-In", Format(Now, "Long Date")

         'Set The Main Form
         FAMILY.SETUP_FORM_MAIN
         'Show the main form
         FAMILY.FORM_MAIN.Show
         DoEvents

'      Else

 '        MsgBox "Sorry, but the user [" & Login_Box(0).Text & "] is currently listed as logged in." & vbNewLine & _
 '              "If you are sure that " & Login_Box(0).Text & " is currently not Logged In, contact an Administrator or view readme.txt", vbInformation + vbOKOnly
         Set FAMILY.FORM_USERS.MyUser.USER_PROFILE = Nothing
 '     End If

   Else   ' Invalid entry
      MsgBox "Invalid Entry", vbCritical + vbInformation
   End If
   Exit Sub

LOGIN_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR - frmLogin.btnLogin" & vbNewLine & _
            "Description - " & Err.Description & vbNewLine & _
            "Error#" & Str$(Err.Number), vbInformation + vbOKOnly
      Err.Clear
   End If
End Sub

Private Sub btnLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Login"
End Sub


Private Sub Form_Load()
   Dim hSysMenu As Long
   ' Get handle of system menu
   hSysMenu = GetSystemMenu(hWnd, 0&)
   'Append separator
   Call AppendMenu(hSysMenu, MF_SEPARATOR, 0&, 0&)
   'Append About
   Call AppendMenu(hSysMenu, MF_STRING, IDM_ABOUT, App.ProductName & " v" & App.Major & "." & App.Minor)
   'Append separator
   Call AppendMenu(hSysMenu, MF_SEPARATOR, 0&, 0&)


   'Check The Databse Connection
   If PUBLIC_DATABASE.CONNECTION.STATE <> 1 Then
      MsgBox "DATABASE CONNECTION ERROR : You are not not connected to the database.", vbCritical + vbOKOnly
      'Close Down
      Set PUBLIC_DATABASE = Nothing
      Set FAMILY = Nothing
      Unload Me
      End
   End If

   Debug.Print "Login Form Loaded"
   'Call modPublic.RemoveMenus(Me)
   Call Setup_TextBox(Me, True, False)
   Login_Box(0).Tag = "User Name Text Box"
   Login_Box(1).Tag = "Password Text Box"
   DoEvents
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RemoveOldProc Me.hWnd, "OldProc"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Label3D3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Family Address Book v3.0"
End Sub

Private Sub txtbx_LoginName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Login Name"
End Sub

Private Sub txtbx_Password_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Password"
End Sub


Private Sub Tester(Optional MSG As String = "Tester")
   MsgBox MSG
End Sub

Private Sub login_box_GotFocus(Index As Integer)
   Login_Box(Index).SelStart = 0
   Login_Box(Index).SelLength = Len(Login_Box(Index).Text)
End Sub

Private Sub login_box_KeyPress(Index As Integer, KeyAscii As Integer)
   If (KeyAscii = 38) Or (KeyAscii = 39) Or (KeyAscii = 34) Or (KeyAscii = 95) Or (KeyAscii = 96) Or (KeyAscii = 45) Then
      MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
      KeyAscii = 0
   End If

   'If The Enter Key Is Pressed
   If KeyAscii = 13 Then
      Call btnLogin_Click
   End If
End Sub
