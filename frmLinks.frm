VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmLinks 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::  Internet Links"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmLinks.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   7500
      TabIndex        =   13
      Top             =   0
      Width           =   7500
      Begin Family_v3.Label3D Label3D3 
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   2625
         _ExtentX        =   4630
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
         Caption         =   " ::  Internet Links ..."
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   3160
      TabIndex        =   6
      Top             =   3000
      Width           =   4335
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   3360
         TabIndex        =   9
         ToolTipText     =   " :: Delete :: "
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   2400
         TabIndex        =   8
         ToolTipText     =   " :: Edit :: "
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   " :: Add :: "
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   " :: Close :: "
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   3160
      TabIndex        =   1
      Top             =   180
      Width           =   4335
      Begin VB.TextBox Link_Textbox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   3360
         TabIndex        =   5
         ToolTipText     =   " :: Cancel :: "
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   2400
         TabIndex        =   4
         ToolTipText     =   " :: Save :: "
         Top             =   2160
         Width           =   855
      End
      Begin Family_v3.Label3D Label3D2 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   -2147483634
         ForeColor2      =   16711680
         Caption         =   "Link"
         Phase           =   1
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3720
         Picture         =   "frmLinks.frx":08CA
         Top             =   1320
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   3120
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3255
         Left            =   60
         TabIndex        =   2
         Top             =   165
         Width           =   2990
         _ExtentX        =   5265
         _ExtentY        =   5741
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   5
         SingleSel       =   -1  'True
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   15
      Top             =   4320
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8043
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
   Begin Family_v3.Label3D Label3D1 
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   3900
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   -2147483634
      ForeColor2      =   16711680
      Caption         =   ":: Your Internet Links (http://www.yourlinks.com) ::"
      Phase           =   1
   End
End
Attribute VB_Name = "frmLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents MyUser          As clsLinks
Attribute MyUser.VB_VarHelpID = -1
Private CURRENT_STATE             As String
Private tmpOldLink                As String


Private Sub btnAdd_Click()
   TreeView1.Enabled = False
   Call Change_Button(False, False, False, False, True, True)
   CURRENT_STATE = "Add"
   Link_Textbox.Locked = False
   Link_Textbox.Text = "http://"
End Sub

Private Sub btnAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Add New Link"
End Sub

Private Sub btnCancel_Click()
   TreeView1.Enabled = True
   CURRENT_STATE = ""
   Link_Textbox.Locked = True
   Link_Textbox.Text = ""

   If FAMILY.FORM_LINKS.MyUser.LinkCount(CURRENT_USER.LOGIN_NAME) >= MAX_LINKS_ALLOWED Then
      Call Change_Button(False, False, False, True, False, False)
   Else
      Call Change_Button(False, False, True, True, False, False)
   End If
End Sub

Private Sub btnCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Cancel"
End Sub

Private Sub btnClose_Click()
   'Unload The Form
   Unload Me
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Close"
End Sub

Private Sub btnDelete_Click()
   TmpMsgResult = MsgBox("Do you want to delete the selected link [" & Link_Textbox & "]", vbQuestion + vbYesNo + vbDefaultButton2)

   If TmpMsgResult = vbYes Then
      Call FAMILY.FORM_LINKS.MyUser.DELETE_LINK(CURRENT_USER.LOGIN_NAME, Link_Textbox.Text)
      'Reload load  the links list
      Call modTreeview.LOAD_LINKS_TREEVIEW(FAMILY.FORM_LINKS, FAMILY.FORM_LINKS.TreeView1, FAMILY.FORM_MAIN.ImageList1)
   End If

   Link_Textbox.Text = ""
   Call Change_Button(False, False, True, True, False, False)
   If FAMILY.FORM_LINKS.MyUser.LinkCount(CURRENT_USER.LOGIN_NAME) >= MAX_LINKS_ALLOWED Then
      btnAdd.Enabled = False
   End If
End Sub


Private Sub btnDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Delete Link"
End Sub

Private Sub btnEdit_Click()
   TreeView1.Enabled = False
   Call Change_Button(False, False, False, False, True, True)
   CURRENT_STATE = "Edit"
   tmpOldLink = Link_Textbox.Text
   Link_Textbox.Locked = False
End Sub

Private Sub btnEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Edit Link"
End Sub

Private Sub btnSave_Click()
   Link_Textbox.Text = ADD_HTTP(Trim$(Link_Textbox))

   Select Case CURRENT_STATE
      Case "Add"
         If FAMILY.FORM_LINKS.MyUser.LinkExist(CURRENT_USER.LOGIN_NAME, Link_Textbox.Text) = False Then
            Call FAMILY.FORM_LINKS.MyUser.ADD_LINK(CURRENT_USER.LOGIN_NAME, Link_Textbox.Text)
         Else
            MsgBox "Sorry, " & CURRENT_USER.FULL_NAME & ", but that link already exist.", vbInformation + vbOKOnly
            Exit Sub
         End If

      Case "Edit"
         If tmpOldLink <> Link_Textbox.Text Then
            If FAMILY.FORM_LINKS.MyUser.LinkExist(CURRENT_USER.LOGIN_NAME, Link_Textbox.Text) = False Then
               Call FAMILY.FORM_LINKS.MyUser.EDIT_LINK(CURRENT_USER.LOGIN_NAME, tmpOldLink, Link_Textbox.Text)
            Else
               MsgBox "Sorry, " & CURRENT_USER.FULL_NAME & ", but that link already exist.", vbInformation + vbOKOnly
               Exit Sub
            End If
         End If

   End Select

   'load  the links list
   Call modTreeview.LOAD_LINKS_TREEVIEW(Me, Me.TreeView1, FAMILY.FORM_MAIN.ImageList1)

   If FAMILY.FORM_LINKS.MyUser.LinkCount(CURRENT_USER.LOGIN_NAME) >= MAX_LINKS_ALLOWED Then
      Call Change_Button(False, False, False, True, False, False)
   Else
      Call Change_Button(False, False, True, True, False, False)
   End If

   TreeView1.Enabled = True
   CURRENT_STATE = ""
   tmpOldLink = ""
   Link_Textbox.Text = ""
   Link_Textbox.Locked = True

End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Save Link"
End Sub

Private Sub Form_Load()
   Call modPublic.RemoveMenus(Me)
   TreeView1.Enabled = True
   Link_Textbox.Locked = True
   Link_Textbox.Text = ""
   Link_Textbox.MaxLength = MAX_LINKS_SIZE
   CURRENT_STATE = ""
   tmpOldLink = ""

   'load  the links list
   Call modTreeview.LOAD_LINKS_TREEVIEW(Me, Me.TreeView1, FAMILY.FORM_MAIN.ImageList1)
   Call Change_Button(False, False, True, True, False, False)

   If FAMILY.FORM_LINKS.MyUser.LinkCount(CURRENT_USER.LOGIN_NAME) >= MAX_LINKS_ALLOWED Then
      Call Change_Button(False, False, False, True, False, False)
   Else
      Call Change_Button(False, False, True, True, False, False)
   End If

End Sub



'====================================================================
'Used to enable and disable the buttons
'====================================================================
Private Sub Change_Button(ByVal EDIT_BUTTON As Boolean, _
                          ByVal DELETE_BUTTON As Boolean, _
                          ByVal ADD_BUTTON As Boolean, _
                          ByVal CLOSE_BUTTON As Boolean, _
                          ByVal Save_Button As Boolean, _
                          ByVal Cancel_Button As Boolean)

   Me.btnEdit.Enabled = EDIT_BUTTON
   Me.btnAdd.Enabled = ADD_BUTTON
   Me.btnDelete.Enabled = DELETE_BUTTON
   Me.btnClose.Enabled = CLOSE_BUTTON
   Me.btnSave.Enabled = Save_Button
   Me.btnCancel.Enabled = Cancel_Button
End Sub
'====================================================================


'====================================================================
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub
'====================================================================


'====================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If CURRENT_STATE <> "" Then
      Cancel = True
   Else
      'Set Up The Link Menus
      FAMILY.FORM_MAIN.CheckLinkMenus
      FAMILY.FORM_LINKS.MyUser.LOAD_LINKS
      DoEvents
      Unload Me
   End If
End Sub
'====================================================================


'====================================================================
Private Function FIND_LINK(ByVal USERNAME As String, TheLink As String) As Boolean
   On Error GoTo FIND_LINK_ERROR

   Set TmpRecordSet = New ADODB.Recordset
   tmpSQL = "SELECT LINK FROM " & LINKS_TABLENAME & _
         " WHERE USER_NAME = '" & USERNAME & "'" & _
         " AND LINK = '" & TheLink & "'"

   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   If TmpRecordSet.RecordCount > 0 Then
      Link_Textbox.Text = TmpRecordSet.Fields("LINK")
      FIND_LINK = True
   Else
      FIND_LINK = False
   End If

   TmpRecordSet.Close
   Set TmpRecordSet = Nothing

FIND_LINK_ERROR:
   If Err.Number <> 0 Then
      FIND_LINK = False
      MsgBox "Error : frmLinks.FIND_LINK_ERROR " & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Err #" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'====================================================================


Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
   Link_Textbox.Text = ""
   If FAMILY.FORM_LINKS.MyUser.LinkCount(CURRENT_USER.LOGIN_NAME) >= MAX_LINKS_ALLOWED Then
      Call Change_Button(False, False, False, True, False, False)
   Else
      Call Change_Button(False, False, True, True, False, False)
   End If
End Sub


Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
   Link_Textbox.Text = ""
   If FAMILY.FORM_LINKS.MyUser.LinkCount(CURRENT_USER.LOGIN_NAME) >= MAX_LINKS_ALLOWED Then
      Call Change_Button(False, False, False, True, False, False)
   Else
      Call Change_Button(False, False, True, True, False, False)
   End If
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim TmpNode As Node

   Set TmpNode = TreeView1.HitTest(X, Y)

   If TmpNode Is Nothing Then
      StatusBar1.Panels(1).Text = ""
   Else
      StatusBar1.Panels(1).Text = TmpNode.Text
   End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
   If Node.Tag = "LINK" Then
      If FIND_LINK(CURRENT_USER.LOGIN_NAME, Node.Text) = True Then
         If FAMILY.FORM_LINKS.MyUser.LinkCount(CURRENT_USER.LOGIN_NAME) >= MAX_LINKS_ALLOWED Then
            Call Change_Button(True, True, False, True, False, False)
         Else
            Call Change_Button(True, True, True, True, False, False)
         End If
      Else
         Link_Textbox.Text = ""
      End If
   Else
      Link_Textbox.Text = ""
      If FAMILY.FORM_LINKS.MyUser.LinkCount(CURRENT_USER.LOGIN_NAME) >= MAX_LINKS_ALLOWED Then
         Call Change_Button(False, False, False, True, False, False)
      Else
         Call Change_Button(False, False, True, True, False, False)
      End If
   End If
End Sub
