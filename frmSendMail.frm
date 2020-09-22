VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Email System ..."
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnNewMessage 
      Caption         =   "&New Message"
      Height          =   310
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   " :: New Message :: "
      Top             =   4260
      Width           =   1695
   End
   Begin VB.TextBox txtRecipientName 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   " :: The Recipient's Name :: "
      Top             =   120
      Width           =   2760
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "C&lose"
      Height          =   300
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   " :: Close :: "
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send Now"
      Height          =   310
      Left            =   6120
      TabIndex        =   5
      ToolTipText     =   " :: Send Message Now :: "
      Top             =   4260
      Width           =   1695
   End
   Begin VB.TextBox txtSubject 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   960
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   " :: The Subject Of This Message :: "
      Top             =   480
      Width           =   6735
   End
   Begin VB.TextBox txtRecipientEmail 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   960
      MaxLength       =   50
      TabIndex        =   0
      ToolTipText     =   " :: The Recipient's Email Address :: "
      Top             =   120
      Width           =   3000
   End
   Begin VB.Frame Frame2 
      Height          =   3345
      Left            =   25
      TabIndex        =   9
      Top             =   800
      Width           =   7815
      Begin VB.PictureBox Picture2 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   40
         ScaleHeight     =   270
         ScaleWidth      =   7710
         TabIndex        =   15
         Top             =   3030
         Width           =   7710
         Begin Family_v3.Label3D Label3D1 
            Height          =   240
            Left            =   1050
            TabIndex        =   16
            Top             =   25
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor1      =   12582912
            Caption         =   " :: Powered by vbSendMail 2.61 from http://www.freevbcode.com ::"
            BackColor       =   -2147483637
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   40
         ScaleHeight     =   255
         ScaleWidth      =   7710
         TabIndex        =   13
         Top             =   160
         Width           =   7710
         Begin Family_v3.Label3D Label3D3 
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1425
            _ExtentX        =   2514
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
            Caption         =   " :: Message ..."
            BackColor       =   -2147483636
         End
      End
      Begin VB.CommandButton btnOpenFile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   4
         ToolTipText     =   " :: Select A File To Send :: "
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtAttach 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   " :: The File Location :: "
         Top             =   2640
         Width           =   6375
      End
      Begin VB.TextBox txtMessage 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         ToolTipText     =   " :: The Message :: "
         Top             =   450
         Width           =   7695
      End
      Begin Family_v3.Label3D LABEL 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   16777215
         ForeColor2      =   16711680
         Caption         =   "File :"
         Phase           =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   18
      Top             =   4650
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8705
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
   Begin Family_v3.Label3D LABEL 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777215
      ForeColor2      =   16711680
      Caption         =   "Name :"
      Phase           =   1
   End
   Begin Family_v3.Label3D LABEL 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777215
      ForeColor2      =   16711680
      Caption         =   "Subject :"
      Phase           =   1
   End
   Begin Family_v3.Label3D LABEL 
      CausesValidation=   0   'False
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777215
      ForeColor2      =   16711680
      Caption         =   "Email :"
      Phase           =   1
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu fd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUUEncode 
         Caption         =   "&UUEncode Encode"
      End
      Begin VB.Menu dfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHtml 
         Caption         =   "Send Message In &HTML Format"
      End
      Begin VB.Menu yhu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigureServer 
         Caption         =   "&Configure Email Server"
         Shortcut        =   ^C
      End
      Begin VB.Menu vbghg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "C&lose"
         Shortcut        =   ^L
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

'Public WithEvents SendEmail       As vbSendMail.clsSendMail
'Public MyEncodeType               As vbSendMail.ENCODE_METHOD
Public bAuthLogin                 As Boolean
Public bHtml                      As Boolean
Public RecipientEmailAddress      As String
Public RecipientName              As String
Public EMAIL_SUBJECT              As String
Public EMAIL_BODY                 As String
'================================================================
'================================================================

'objForm.Width = 8790
'objForm.Height = 5890

Private Sub btnCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Cancel"
End Sub

Private Sub btnClose_Click()
   'SetUp FrmMain Texboxes and Comboboxes
   Call Setup_TextBox(FAMILY.FORM_MAIN, True, True)
   Call Setup_ComboBox(FAMILY.FORM_MAIN, True, True)
   Call FAMILY.FORM_MAIN.Change_Button(False, False, True, True, False, False)
   'Unload The Form
   Unload Me
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Close"
End Sub



Private Sub btnConfigureServer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Configure Email Server (SMTP Server)"
End Sub

Private Sub btnNewMessage_Click()
   txtRecipientEmail.Text = RecipientEmailAddress
   txtRecipientName.Text = RecipientName
   txtSubject.Text = EMAIL_SUBJECT
   txtMessage.Text = EMAIL_BODY
   txtAttach.Text = ""
End Sub



Private Sub btnNewMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "New Message"
End Sub

Private Sub btnOpenFile_Click()
   Dim tmpFileName As String

   ' // call the open Procedure
   tmpFileName = modPublic.Open_File(Me.hWnd)

   ' // Check the return value
   If tmpFileName <> "" Then
      txtAttach.Text = tmpFileName
   Else
      txtAttach.Text = ""
   End If
End Sub

Private Sub btnOpenFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Select File"
End Sub

Private Sub btnSend_Click()
   'first check if the email adddress provided is valid
   If Len(txtRecipientEmail.Text) < 1 Then
      MsgBox CURRENT_USER.FULL_NAME & ", please enter the recipient's email address.", vbInformation + vbOKOnly
      txtRecipientEmail.SetFocus
      Exit Sub
   End If
   
   If modPublic.IsValidEmail(txtRecipientEmail.Text) = False Then
      TmpMsgResult = MsgBox("Hi " & CURRENT_USER.FULL_NAME & ", is this email valid [" & txtRecipientEmail.Text & "] ?", vbQuestion + vbYesNo)
      If TmpMsgResult = vbNo Then
         txtRecipientEmail.SetFocus
         Exit Sub
      End If
   End If

   'check the user's email adddress provided is valid
   If modPublic.IsValidEmail(CURRENT_USER.EMAIL_ADDRESS) = False Then
      TmpMsgResult = MsgBox("Hi " & CURRENT_USER.FULL_NAME & ", is your email address valid [" & CURRENT_USER.EMAIL_ADDRESS & "] ?", vbQuestion + vbYesNo)
      If TmpMsgResult = vbNo Then
         MsgBox CURRENT_USER.FULL_NAME & ", you can change you email address by " & vbNewLine & _
               "Openning the Option Menu (top) and then click Configure Email Server", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If


  'Check for message subject
  If Len(Trim(txtSubject.Text)) < 1 Then
      MsgBox CURRENT_USER.FULL_NAME & ", please enter the subject of this message.", vbInformation + vbOKOnly
      txtSubject.SetFocus
      Exit Sub
  End If
  
  
  'Check for message subject
  If Len(Trim(txtMessage.Text)) < 1 Then
      MsgBox CURRENT_USER.FULL_NAME & ", please enter the message that you want to send.", vbInformation + vbOKOnly
      txtMessage.SetFocus
      Exit Sub
  End If
  

   Screen.MousePointer = vbHourglass


   Screen.MousePointer = vbDefault

End Sub


Private Sub btnSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Send Message"
End Sub

Private Sub Form_Load()
   Call mnuOptions_Click
   Call btnNewMessage_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call btnClose_Click
End Sub

Private Sub mnuConfigureServer_Click()
   FAMILY.FORM_USERS.Show vbModal
End Sub

Private Sub Send_Email_Progress(PercentComplete As Long)
   StatusBar1.Panels(1).Text = Str$(PercentComplete) & "%"
End Sub

Private Sub Send_Email_SendFailed(Explanation As String)
   StatusBar1.Panels(1).Text = "Unable to send your message"
End Sub

Private Sub Send_Email_SendSuccesful()
   StatusBar1.Panels(1).Text = "Your message was sent succesfully"
End Sub

Private Sub Send_Email_Status(Status As String)
   StatusBar1.Panels(1).Text = "Status : " & Status
End Sub



Private Sub mnuHtml_Click()
   If mnuHtml.Checked = True Then
      mnuHtml.Checked = False
      bHtml = False
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "SHTML", "False"
   Else
      mnuHtml.Checked = True
      bHtml = True
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "SHTML", "True"
   End If
End Sub

Private Sub mnuOptions_Click()
   Dim Tmpstr1 As String

   Tmpstr1 = ""
   Tmpstr1 = ReadIniFile(App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "UUEncode", "")
   If Tmpstr1 = "True" Then
      mnuUUEncode.Checked = True
'      MyEncodeType = UU_ENCODE
   Else
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "UUEncode", "False"
      mnuUUEncode.Checked = False
'      MyEncodeType = MIME_ENCODE
   End If

   Tmpstr1 = ""
   Tmpstr1 = ReadIniFile(App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "SHTML", "")
   If Tmpstr1 = "True" Then
      mnuHtml.Checked = True
      bHtml = True
   Else
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "SHTML", "False"
      mnuHtml.Checked = False
      bHtml = False
   End If
End Sub

Private Sub mnuUUEncode_Click()
   If mnuUUEncode.Checked = True Then
      mnuUUEncode.Checked = False
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "UUEncode", "False"
'      MyEncodeType = MIME_ENCODE
   Else
      mnuUUEncode.Checked = True
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "UUEncode", "True"
 '     MyEncodeType = UU_ENCODE
   End If
End Sub

Private Sub SendEmail_Progress(PercentComplete As Long)
   Debug.Print "Progress %" & Str$(PercentComplete)
   StatusBar1.Panels(2).Text = "Progress %" & Str$(PercentComplete)
End Sub

Private Sub SendEmail_SendFailed(Explanation As String)
   StatusBar1.Panels(1).Text = "SendFailed : " & Explanation
   MsgBox "SendFailed : " & Explanation, vbCritical + vbOKOnly
End Sub

Private Sub SendEmail_SendSuccesful()
   Debug.Print "Message Sent Succesfully"
   StatusBar1.Panels(1).Text = "Message Sent Succesfully"
   MsgBox "Hi " & CURRENT_USER.FULL_NAME & ", your message was sent succesfully", vbInformation + vbOKOnly
End Sub

Private Sub SendEmail_Status(Status As String)
   StatusBar1.Panels(1).Text = "Message Status : " & Status
End Sub

Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(2).ToolTipText = " " & Format(Now, "LONG DATE") & " "
   StatusBar1.Panels(3).ToolTipText = " " & Format(Now, "LONG TIME") & " "
End Sub

