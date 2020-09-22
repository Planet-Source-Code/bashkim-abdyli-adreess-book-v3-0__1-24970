VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmReminderNotes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReminderNotes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame b 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   20
      TabIndex        =   4
      Top             =   0
      Width           =   6960
      Begin VB.Timer Timer1 
         Left            =   6120
         Top             =   2160
      End
      Begin VB.VScrollBar dayVScroll 
         Height          =   290
         Left            =   3120
         Max             =   -1
         Min             =   -31
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   862
         Value           =   -1
         Width           =   250
      End
      Begin VB.ComboBox ComboBox 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   5535
      End
      Begin VB.OptionButton Opt_PM 
         Alignment       =   1  'Right Justify
         Caption         =   "PM"
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
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton Opt_AM 
         Alignment       =   1  'Right Justify
         Caption         =   "AM"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   1125
         Index           =   3
         Left            =   70
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1560
         Width           =   6800
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
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
         Index           =   2
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "12:12"
         ToolTipText     =   " :: The Time Should Look Like This -    12:58      :: "
         Top             =   830
         Width           =   520
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   830
         Width           =   610
      End
      Begin VB.ComboBox ComboBox 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin Family_v3.Label3D Label3D3 
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   -2147483634
         ForeColor2      =   16711680
         Caption         =   "Time"
         Phase           =   1
      End
      Begin Family_v3.Label3D Label3D2 
         Height          =   255
         Left            =   100
         TabIndex        =   10
         Top             =   1200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   16777215
         ForeColor2      =   16711680
         Caption         =   "Notes"
         Phase           =   1
      End
      Begin Family_v3.Label3D Label3D1 
         Height          =   255
         Left            =   100
         TabIndex        =   9
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   -2147483634
         ForeColor2      =   16711680
         Caption         =   "Subject :"
         Phase           =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3720
         TabIndex        =   8
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   345
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   3330
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7594
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "7/13/2001"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "11:34 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   20
      TabIndex        =   14
      Top             =   2700
      Width           =   6960
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   310
         Left            =   5880
         TabIndex        =   22
         ToolTipText     =   " Close "
         Top             =   200
         Width           =   850
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   310
         Left            =   3600
         TabIndex        =   21
         ToolTipText     =   " Add A New Reminder "
         Top             =   200
         Width           =   850
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Height          =   310
         Left            =   4560
         TabIndex        =   20
         ToolTipText     =   " Edit The Current Reminder "
         Top             =   200
         Width           =   850
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "&Save"
         Height          =   310
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   " Save and Close "
         Top             =   200
         Width           =   850
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "&Cancel"
         Height          =   310
         Left            =   1130
         TabIndex        =   15
         ToolTipText     =   " Cancel and Close "
         Top             =   200
         Width           =   850
      End
   End
End
Attribute VB_Name = "frmReminderNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub btnAdd_Click()
   FAMILY.FORM_REMINDERS.CURRENT_STATE = 1

   TextBox(0).Locked = False
   TextBox(1).Locked = False
   TextBox(2).Locked = False
   TextBox(3).Locked = False
   ComboBox(0).Enabled = True
   ComboBox(1).Enabled = True
   Opt_AM.Enabled = True
   Opt_PM.Enabled = True

   dayVScroll.Enabled = True

   TextBox(0).Text = ""
   TextBox(1).Text = Day(Now)
   TextBox(2).Text = MyTime_TIME(Time)
   TextBox(3).Text = ""

   dayVScroll.Enabled = True
   dayVScroll.Value = -CInt(TextBox(1).Text)

   If MyTime_AMPM(Time) = "PM" Then
      Opt_PM.Value = True
   Else
      Opt_AM.Value = True
   End If

   btnSave.Enabled = True
   btnCancel.Enabled = True
   btnEdit.Enabled = False
   btnAdd.Enabled = False
   btnClose.Enabled = False
End Sub

Private Sub btnAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Add A New Reminder"
End Sub

Private Sub btnCancel_Click()
   FAMILY.FORM_REMINDERS.CURRENT_STATE = 0
   Unload Me
End Sub

Private Sub btnCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Cancel And Close"
End Sub

Private Sub btnClose_Click()
   FAMILY.FORM_REMINDERS.CURRENT_STATE = 0
   Unload Me
End Sub


Private Sub btnEdit_Click()

   If FAMILY.FORM_REMINDER_NOTES.FIND_RECORD(FAMILY.FORM_REMINDERS.Selected_Record_Key) = True Then

      If (Not FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.BOF) And (Not FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.EOF) Then
         TextBox(0).Locked = False
         TextBox(1).Locked = False
         TextBox(2).Locked = False
         TextBox(3).Locked = False
         ComboBox(0).Enabled = True
         ComboBox(1).Enabled = True
         Opt_AM.Enabled = True
         Opt_PM.Enabled = True

         dayVScroll.Enabled = True

         btnSave.Enabled = True
         btnCancel.Enabled = True
         btnEdit.Enabled = False
         btnAdd.Enabled = False
         btnClose.Enabled = False

         FAMILY.FORM_REMINDERS.CURRENT_STATE = 2
      End If
   Else
      MsgBox "Unable to edit record", vbCritical + vbOKOnly, "Record not found!"
      FAMILY.FORM_REMINDERS.CURRENT_STATE = 0
      Unload Me
   End If

End Sub

Private Sub btnSave_Click()
   On Error GoTo SAVE_ERROR

   If Len(Trim$(TextBox(0).Text)) < 1 Then
      MsgBox CURRENT_USER.FULL_NAME & ", you need to put something in the SUBJECT box", vbInformation + vbOKOnly
      Exit Sub
   End If

   If Len(Trim$(TextBox(3).Text)) < 1 Then
      MsgBox CURRENT_USER.FULL_NAME & ", you need to put something in the NOTES box", vbInformation + vbOKOnly
      Exit Sub
   End If

   '90/90/2000
   If Opt_AM.Value = True Then
      If IsDate(Str$(ComboBox(0).ListIndex + 1) & "/" & Str$(TextBox(1).Text) & "/" & _
            Str$(ComboBox(1).Text) & " " & Trim$(TextBox(2).Text) & " " & "AM") = False Then
         MsgBox "Invalid Date or Time"
         Exit Sub
      End If
   Else
      If IsDate(Str$(ComboBox(0).ListIndex + 1) & "/" & Str$(TextBox(1).Text) & "/" & _
            Str$(ComboBox(1).Text) & " " & Trim$(TextBox(2).Text) & " " & "PM") = False Then
         MsgBox CURRENT_USER.FULL_NAME & ", you have an Invalid Date or Invalid Time", vbInformation + vbOKOnly
         Exit Sub
      End If

   End If

   Select Case FAMILY.FORM_REMINDERS.CURRENT_STATE
         '1 - Adding '2 - Edditing  '3 - Viewing

      Case 1   '1 - Adding
         TmpString = ""
         TmpString = "SELECT USER_NAME,DATE_ENTERED,DATE_EXPIRED,SUBJECT,TODO FROM " & REMINDERS_TABLENAME
         TmpString = TmpString & " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'"

         'Create A New Recordset
         Set FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER = New ADODB.Recordset
         'Open A New Recordset
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Open TmpString, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

         'Add New
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.AddNew

         'Get The Values
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("USER_NAME").Value = CURRENT_USER.LOGIN_NAME
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_ENTERED").Value = Now

         If Opt_AM.Value = True Then
            FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED").Value = modPublic.sDate(ComboBox(0).ListIndex + 1, Trim$(TextBox(1).Text), ComboBox(1).Text, Trim$(TextBox(2).Text), True)
         Else
            FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED").Value = modPublic.sDate(ComboBox(0).ListIndex + 1, Trim$(TextBox(1).Text), ComboBox(1).Text, Trim$(TextBox(2).Text), False)
         End If

         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("SUBJECT").Value = Trim$(Apostrophe(TextBox(0).Text))
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("TODO").Value = Apostrophe(TextBox(3).Text)
         'Update The Database
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Update
         'Requery
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Requery

         'Clear FORM_REMINDER'S Current_State
         FAMILY.FORM_REMINDERS.CURRENT_STATE = 0


      Case 2   'Edit
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("USER_NAME").Value = CURRENT_USER.LOGIN_NAME
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_ENTERED").Value = Now

         If Opt_AM.Value = True Then
            FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED").Value = sDate(ComboBox(0).ListIndex + 1, Trim$(TextBox(1).Text), ComboBox(1).Text, Trim$(TextBox(2).Text), True)
         Else
            FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED").Value = sDate(ComboBox(0).ListIndex + 1, Trim$(TextBox(1).Text), ComboBox(1).Text, Trim$(TextBox(2).Text), False)
         End If

         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("SUBJECT").Value = Trim$(Apostrophe(TextBox(0).Text))
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("TODO").Value = Apostrophe(TextBox(3).Text)
         'Update The Database
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Update
         'Requery
         FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Requery
         'Clear FORM_REMINDER'S Current_State
         FAMILY.FORM_REMINDERS.CURRENT_STATE = 0
      Case Else
         MsgBox CURRENT_USER.LOGIN_NAME & ", I'm unable to properly process your request", vbInformation + vbOKOnly, "Invalid Request"

   End Select

   Call FAMILY.FORM_REMINDERS.LOAD_DATA
   Unload Me

SAVE_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR frmReminderNotes.btnSave" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Error#" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear

      Call FAMILY.FORM_REMINDERS.LOAD_DATA
      Unload Me
   End If
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Save The Changes Made and Close"
End Sub


Private Sub ComboBox_Click(Index As Integer)
   Select Case ComboBox(0).Text
      Case Is = "September", "April", "June", "November"
         dayVScroll.Min = (-30)
      Case Is = "February"
         If IsLeapYear(ComboBox(1).Text) = True Then
            dayVScroll.Min = (-29)
            MsgBox "Leap Year"
         Else
            dayVScroll.Min = (-28)
         End If
      Case Else
         dayVScroll.Min = (-31)
   End Select
End Sub


Private Sub dayVScroll_Change()
   TextBox(1).Text = (-dayVScroll.Value)
End Sub

Private Sub Form_Load()
   Call modPublic.RemoveMenus(Me)

   ComboBox(0).Clear
   ComboBox(0).AddItem "January"
   ComboBox(0).AddItem "February"
   ComboBox(0).AddItem "March"
   ComboBox(0).AddItem "April"
   ComboBox(0).AddItem "May"
   ComboBox(0).AddItem "June"
   ComboBox(0).AddItem "July"
   ComboBox(0).AddItem "August"
   ComboBox(0).AddItem "September"
   ComboBox(0).AddItem "October"
   ComboBox(0).AddItem "November"
   ComboBox(0).AddItem "December"

   ComboBox(0).ListIndex = Month(Now) - 1

   For TmpByte = 0 To 10   'Add 10 years to the current year
      ComboBox(1).AddItem CInt(Format$(Now, "YYYY")) + TmpByte
   Next TmpByte

   ComboBox(1).ListIndex = 0

   Select Case FAMILY.FORM_REMINDERS.CURRENT_STATE
      Case 1   'Add
         Call btnAdd_Click

      Case 2   'Edit
         Call btnEdit_Click

      Case 3   'View
         If FAMILY.FORM_REMINDER_NOTES.FIND_RECORD(FAMILY.FORM_REMINDERS.Selected_Record_Key) = True Then
            TextBox(0).Locked = True
            TextBox(1).Locked = True
            TextBox(2).Locked = True
            TextBox(3).Locked = True
            ComboBox(0).Enabled = False
            ComboBox(1).Enabled = False
            Opt_AM.Enabled = False
            Opt_PM.Enabled = False
            dayVScroll.Enabled = False

            btnSave.Enabled = False
            btnCancel.Enabled = False
            btnEdit.Enabled = True
            btnAdd.Enabled = True
            btnClose.Enabled = True
         Else
            MsgBox "Unable to find record", vbCritical + vbOKOnly, "Record not found!"
            Unload Me
         End If
   End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FAMILY.FORM_REMINDERS.CURRENT_STATE = 0
End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub

Private Sub TextBox_GotFocus(Index As Integer)
   TextBox(Index).SelStart = 0
   TextBox(Index).SelLength = Len(TextBox(Index).Text)
End Sub


Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         If (KeyAscii = 39) Or (KeyAscii = 34) Or (KeyAscii = 95) Then
            MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
            KeyAscii = 0
         End If
      Case 1, 2
         Select Case KeyAscii
            Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 8
            Case Else
               KeyAscii = 0
               Exit Sub
         End Select
   End Select
End Sub


'====================================================================
'Locate the Record
'====================================================================
Public Function FIND_RECORD(ByVal Record_String As String) As Boolean
   Dim sPos As Byte
   Dim tmpDate As String
   Dim tmpSubject As String
   On Error GoTo FIND_RECORD_ERROR

   sPos = InStr(1, Record_String, "_")
   If sPos > 0 Then
      tmpSubject = Apostrophe$(Mid$(Record_String, 1, sPos - 1))
      tmpDate = Apostrophe$(Mid$(Record_String, sPos + 1))

      tmpSQL = ""
      tmpSQL = "SELECT DATE_EXPIRED,SUBJECT,TODO,USER_NAME,DATE_ENTERED From " & REMINDERS_TABLENAME
      tmpSQL = tmpSQL & " WHERE DATE_ENTERED = #" & CDate(tmpDate) & "#"   ' Compare The Dates
      tmpSQL = tmpSQL & " and SUBJECT = '" & tmpSubject & "'"
      tmpSQL = tmpSQL & " and USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'"


      Set FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER = New ADODB.Recordset

      FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

      If FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.RecordCount > 0 Then
         ComboBox(0).ListIndex = Month(FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED")) - 1
         TextBox(0).Text = FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("SUBJECT")
         TextBox(1).Text = Day(FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED"))
         TextBox(2).Text = MyTime_TIME(FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED"))

         If MyTime_AMPM(FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("DATE_EXPIRED")) = "PM" Then
            Opt_PM.Value = True
         Else
            Opt_AM.Value = True
         End If

         TextBox(3).Text = FAMILY.FORM_REMINDERS.MyUser.USER_REMINDER.Fields("TODO")

         dayVScroll.Value = -CInt(TextBox(1).Text)
         FIND_RECORD = True
      Else
         FIND_RECORD = False
      End If
   End If
   Exit Function

FIND_RECORD_ERROR:
   If Err.Number <> 0 Then
      FIND_RECORD = False
      MsgBox "ERROR frmReminderNotes.FIND_RECORD" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Error#" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'====================================================================
