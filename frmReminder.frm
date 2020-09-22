VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmReminder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " :: Your Daily Reminders ..."
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmReminder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAdd 
      Caption         =   "&Add A New Reminder"
      Height          =   310
      Left            =   5760
      TabIndex        =   5
      ToolTipText     =   " Add A New Reminder "
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   " Close "
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame frameAddress 
      Height          =   3255
      Left            =   10
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   50
         ScaleHeight     =   255
         ScaleWidth      =   7935
         TabIndex        =   6
         Top             =   120
         Width           =   7940
         Begin Family_v3.Label3D Label3D2 
            Height          =   255
            Left            =   50
            TabIndex        =   7
            Top             =   10
            Width           =   2985
            _ExtentX        =   5265
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
            Caption         =   ":: Your Daily Reminders ..."
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReminder.frx":030A
               Key             =   "closed"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReminder.frx":0466
               Key             =   "open"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvReminders 
         Height          =   2775
         Left            =   45
         TabIndex        =   1
         Top             =   390
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   10239
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2118
            MinWidth        =   2118
            TextSave        =   "7/13/2001"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:34 AM"
         EndProperty
      EndProperty
   End
   Begin Family_v3.Label3D Label3D1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777215
      ForeColor2      =   16711680
      Caption         =   "To View, Edit or Delete A Reminder, Just Double Mouse Click or Right Mouse Click On The Reminder...."
      Phase           =   1
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu ery 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add A New Reminder"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit The Selected Reminder"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View The Selected Reminder"
      End
      Begin VB.Menu fdrt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete The Selected Reminder"
      End
      Begin VB.Menu drte 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh Reminders"
      End
      Begin VB.Menu hghf 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Public_LV_Item         As ListItem
Public Selected_Record_Key    As String
Public WithEvents MyUser      As clsReminder
Attribute MyUser.VB_VarHelpID = -1
Public DateString             As String
Public CURRENT_STATE          As Integer
'1 - Adding
'2 - Editing
'3 - Viewing


Private Sub btnAdd_Click()
   'Add
   CURRENT_STATE = 1
   Call mnuAdd_Click
End Sub

Private Sub btnAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Add A New Reminder"
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
   StatusBar1.Panels(1).Text = "Close Your Daily Reminders"
End Sub


Private Sub Form_Load()
   If FAMILY.FORM_REMINDERS.LOAD_DATA = False Then
      MsgBox "Unable to load data.", vbCritical + vbOKOnly
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CURRENT_STATE = 0
End Sub

Private Sub Label3D1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub

Private Sub lvReminders_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   With lvReminders
      If .SortKey <> ColumnHeader.Index - 1 Then
         .SortKey = ColumnHeader.Index - 1
         .SortOrder = lvwAscending
      Else
         If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
         Else
            .SortOrder = lvwAscending
         End If
      End If
      .Sorted = True
   End With
End Sub

Private Sub lvReminders_DblClick()
   If Public_LV_Item Is Nothing Then
      Exit Sub
   Else
      Call mnuView_Click
   End If
End Sub

Private Sub lvReminders_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Selected_Record_Key = Item.Key
End Sub

Private Sub lvReminders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Set Public_LV_Item = lvReminders.HitTest(X, Y)

   'Check if a record was selected
   If Public_LV_Item Is Nothing Then

      If lvReminders.ListItems.Count > 0 Then
         lvReminders.SelectedItem.Selected = False
      End If

      If Button = 2 Then
         'Setup Menu Items
         mnuAdd.Enabled = True
         mnuEdit.Enabled = False
         mnuDelete.Enabled = False
         mnuView.Enabled = False
         PopupMenu mnuOptions, , , , mnuAdd
      End If
      Exit Sub
   Else
      Public_LV_Item.Selected = True
      If Button = 2 Then
         Selected_Record_Key = Public_LV_Item.Key
         'Setup Menu Items
         mnuAdd.Enabled = True
         mnuEdit.Enabled = True
         mnuDelete.Enabled = True
         mnuView.Enabled = True
         PopupMenu mnuOptions, , , , mnuView
      End If
   End If
End Sub

Private Sub lvReminders_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim TmpLvi As ListItem

   Set TmpLvi = lvReminders.HitTest(X, Y)

   'Display Which Item The Mouse Is Over
   If TmpLvi Is Nothing Then
      StatusBar1.Panels(1).Text = ""
      Exit Sub
   Else
      StatusBar1.Panels(1).Text = TmpLvi.Text
   End If
End Sub

Private Sub mnuAdd_Click()
   'Add
   CURRENT_STATE = 1
   FAMILY.FORM_REMINDER_NOTES.Show vbModal
End Sub

Private Sub mnuDelete_Click()
   On Error GoTo DELETE_ERROR
   
   TmpMsgResult = MsgBox("Are you sure you want to Delete the selected record?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Record")

   If TmpMsgResult = vbYes Then

      Dim sPos As Byte
      Dim tmpDate As String
      Dim tmpSubject As String
      Dim tmpSQL As String

      sPos = InStr(1, Selected_Record_Key, "_")
      If sPos > 0 Then
         tmpSubject = Apostrophe$(Mid$(Selected_Record_Key, 1, sPos - 1))
         tmpDate = Apostrophe$(Mid$(Selected_Record_Key, sPos + 1))

         tmpSQL = ""
         tmpSQL = "DELETE * FROM " & REMINDERS_TABLENAME
         tmpSQL = tmpSQL & " WHERE DATE_ENTERED = #" & CDate(tmpDate) & "#"   ' Compare The Dates
         tmpSQL = tmpSQL & " and SUBJECT = '" & tmpSubject & "'"
         tmpSQL = tmpSQL & " and USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'"

         Set FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER = New ADODB.Recordset

         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

      Else
         MsgBox "ERROR : The Selected Record Was Not Deleted", vbCritical + vbOKOnly
      End If

      If FAMILY.FORM_REMINDERS.LOAD_DATA = False Then
         MsgBox "Unable to load data.", vbCritical + vbOKOnly
      End If
   End If
   
   
DELETE_ERROR:
   If Err.Number <> 0 Then
      MsgBox "Error : frmReminder.mnuDelete " & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Err #" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Sub

Private Sub mnuEdit_Click()
   'edit
   CURRENT_STATE = 2
   FAMILY.FORM_REMINDER_NOTES.Show vbModal
End Sub

Private Sub mnuRefresh_Click()
   Call Me.LOAD_DATA
End Sub

Private Sub mnuView_Click()
   'Viewing
   CURRENT_STATE = 3
   FAMILY.FORM_REMINDER_NOTES.Show vbModal
End Sub




'================================================================
' USED To Load Data
'================================================================
Public Function LOAD_DATA() As Boolean
   On Error GoTo LOAD_DATA_ERROR

   LOAD_DATA = False

   'Set The listView's View Type To lvReport
   lvReminders.View = lvwReport
   'Load The ColumnHeaders
   Call Load_Reminders_ColumnHeaders(lvReminders)

   'book1
   tmpSQL = ""
   tmpSQL = "SELECT USER_NAME,DATE_ENTERED,DATE_EXPIRED,SUBJECT FROM " & REMINDERS_TABLENAME & _
         " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
         " ORDER BY DATE_EXPIRED DESC"

   Set FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER = New ADODB.Recordset
   FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Requery

   If FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.RecordCount > 0 Then
      Dim lvListItems As ListItem

      lvReminders.ListItems.Clear

      Do While Not FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.EOF
         Dim tmp_key As String

         tmp_key = FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("SUBJECT") & "_" & _
               FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_ENTERED")

         If FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED") >= Now Then
            Set lvListItems = lvReminders.ListItems.Add(, tmp_key, FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("SUBJECT").Value, "closed", "closed")
            lvListItems.SubItems(1) = Format$(FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED").Value, "Long Date")
            lvListItems.SubItems(2) = Format$(FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED").Value, "H:MM AMPM")
         Else
            Set lvListItems = lvReminders.ListItems.Add(, tmp_key, FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("SUBJECT").Value, "open", "open")
            lvListItems.SubItems(1) = Format$(FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED").Value, "Long Date")
            lvListItems.SubItems(2) = Format$(FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED").Value, "H:MM AMPM")
         End If
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.MoveNext
      Loop
   End If
   LOAD_DATA = True

LOAD_DATA_ERROR:
   If Err.Number <> 0 Then
      LOAD_DATA = False
      MsgBox "ERROR - frmReminder" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Error #" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'================================================================

