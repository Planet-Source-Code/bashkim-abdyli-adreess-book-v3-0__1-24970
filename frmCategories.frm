VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::  Your Categories/Relationship Groups"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmCategories.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4605
      TabIndex        =   12
      Top             =   0
      Width           =   4600
      Begin Family_v3.Label3D Label3D3 
         Height          =   255
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   15
         Width           =   3600
         _ExtentX        =   6350
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
         Caption         =   " :: Categories/Relationship Groups ..."
      End
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   3570
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   4313
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   3540
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3120
      Width           =   980
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   50
      TabIndex        =   3
      Top             =   3120
      Width           =   980
   End
   Begin VB.TextBox txtGroup 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   50
      TabIndex        =   2
      Tag             =   "Category Name Text Box"
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Height          =   2160
      Left            =   0
      TabIndex        =   1
      Top             =   180
      Width           =   4605
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3550
         TabIndex        =   11
         ToolTipText     =   " Edit Add A New Category/Relationship Group "
         Top             =   720
         Width           =   980
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3550
         TabIndex        =   10
         Top             =   1200
         Width           =   980
      End
      Begin VB.CommandButton btnExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3550
         TabIndex        =   9
         Top             =   1680
         Width           =   980
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   3550
         TabIndex        =   8
         ToolTipText     =   " Add A New Category/Relationship Group "
         Top             =   240
         Width           =   980
      End
      Begin VB.ListBox CategoryList 
         BackColor       =   &H00C0FFFF&
         Height          =   1815
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
   End
   Begin Family_v3.Label3D Label3D1 
      Height          =   255
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   ":: Category Name ::"
      Phase           =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3840
      Picture         =   "frmCategories.frx":030A
      Top             =   2640
      Width           =   480
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'3100 - height1  - normal form height
'4230 - height2  - extended form height

'This will be used to tell if any changes were made
Dim ChangesMade As Boolean
'Dim MY_RECORDSET As ADODB.Recordset
Dim CurrentMode As String

Public MyUser As clsCategories



'=====================================================================
Private Sub btnAdd_Click()
   'Change the form height
   FAMILY.FORM_CATEGORIES.Height = 4230

   'Change Up Buttons
   Call ChangeButtons(False, False, False, False)

   CurrentMode = "Add"

   txtGroup.Text = "NewCategory"
   txtGroup.SetFocus

End Sub
'=====================================================================



'=====================================================================
Private Sub btnApply_Click()
   Dim i As Byte
   Dim i3 As Byte
   Dim tmpPos As Byte

   txtGroup.Text = Trim$(txtGroup.Text)
   If Len(txtGroup.Text) < MINIMUM_CATEGORY_NAME_SIZE Then
      MsgBox "Insufficient amount of character", vbInformation + vbOKOnly
      Exit Sub
   End If

   'Check For Spaces
   If InStr(1, txtGroup.Text, " ") > 0 Then
      MsgBox CURRENT_USER.FULL_NAME & ", you have space in the " & txtGroup.Tag, vbCritical + vbOKOnly
      Exit Sub
   End If

   'Check For Invalid Characters
   For i = LBound(IllegalChars) To UBound(IllegalChars)
      For i3 = 1 To Len(txtGroup.Text)
         If Chr$(IllegalChars(i)) = Mid$(txtGroup.Text, i3, 1) Then
            MsgBox CURRENT_USER.FULL_NAME & ", you have an invalid charater [ " & Chr$(IllegalChars(i)) & " ] , in the " & txtGroup.Tag, vbCritical + vbOKOnly
            Exit Sub
         End If
      Next i3
   Next i


   'Check Is The Name Was or was not changed
   If txtGroup.Text = CategoryList.Text Then
      CurrentMode = ""
      Call btnCancel_Click
      'since it was not changed we exit the sub
      Exit Sub
   End If

   'First Check If The Category Already Exist
   If (FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY_EXIST(txtGroup.Text, CURRENT_USER.LOGIN_NAME) = True) Then
      MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", but [ " & txtGroup.Text & " ] already exists.", vbInformation + vbOKOnly
      Exit Sub

   Else

      txtGroup.Text = Apostrophe(Trim$(txtGroup.Text))

      Select Case CurrentMode
         Case "Add"   'Add A New Category
            Call FAMILY.FORM_CATEGORIES.MyUser.ADD_CATEGORY(txtGroup.Text, CURRENT_USER.LOGIN_NAME)

         Case "Edit"
            If FAMILY.FORM_CATEGORIES.MyUser.EDIT_CATEGORY(CategoryList.Text, txtGroup.Text, CURRENT_USER.LOGIN_NAME) = False Then
               MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", the program was unable to properly edit the category.", vbInformation + vbOKOnly
            End If
      End Select

      'Set this to True since we know that changes were made
      ChangesMade = True

      'Reload Categorylist
      Call Load_CategoryListbox

      'Select The Last Item That Was Modified
      CategoryList.Text = Trim$(txtGroup.Text)

      'Change Up Buttons
      Call ChangeButtons(True, True, True, True)
      CurrentMode = ""
      FAMILY.FORM_CATEGORIES.Height = 3100
   End If
End Sub
'=====================================================================



'=====================================================================
Private Sub btnCancel_Click()
   'Change Up Buttons
   Call ChangeButtons(True, True, True, True)
   CurrentMode = ""
   Me.Height = 3100
End Sub
'=====================================================================



'=====================================================================
Private Sub btnDelete_Click()
   If Len(CategoryList.Text) > 0 Then
      CurrentMode = "Delete"

      TmpMsgResult = MsgBox(CURRENT_USER.LOGIN_NAME & ", are you sure that you want to remove (" & CategoryList.Text & ") and all it's contents ?", vbYesNo + vbOKOnly + vbDefaultButton2 + vbQuestion)

      If TmpMsgResult = vbYes Then
         'Delete Category
         If FAMILY.FORM_CATEGORIES.MyUser.DELETE_CATEGORY(CategoryList.Text, CURRENT_USER.LOGIN_NAME) = False Then
            MsgBox "Unable to properly delete category", vbInformation + vbOKOnly
         End If

         'Reload the list
         Call Load_CategoryListbox
         'set chiangesmade to true
         ChangesMade = True
      End If
      CurrentMode = ""
   End If
End Sub
'=====================================================================



'=====================================================================
Private Sub btnEdit_Click()
   'Set The Form's Height
   Me.Height = 4230

   'Change Up Buttons
   Call ChangeButtons(False, False, False, False)

   CurrentMode = "Edit"

   txtGroup.Text = CategoryList.Text
   txtGroup.SetFocus
End Sub
'=====================================================================



'=====================================================================
Private Sub btnExit_Click()
   'Check If Changes Were Made
   If ChangesMade = True Then
      'Set Up FrmMain
      Call FAMILY.SETUP_FORM_MAIN
   End If

   'SetUp FrmMain Texboxes and Comboboxes
   Call Setup_TextBox(FAMILY.FORM_MAIN, True, True)
   Call Setup_ComboBox(FAMILY.FORM_MAIN, True, True)
   Call FAMILY.FORM_MAIN.Change_Button(False, False, True, True, False, False)
   'Unload The Form
   Unload Me
End Sub
'=====================================================================



'=====================================================================
Private Sub Form_Load()
   'Set The For Height
   FAMILY.FORM_CATEGORIES.Height = 3100
   'Remove The Colose Button
   Call modPublic.RemoveMenus(Me)
   'Set MaxLength
   txtGroup.MaxLength = MAX_CATEGORY_NAME_SIZE
   'Load the listbox with all the categories
   Call FAMILY.FORM_CATEGORIES.Load_CategoryListbox
   'Set ChangesMade = False (Changes were made?)
   ChangesMade = False
End Sub
'=====================================================================


'=====================================================================
'Load_CategoryListbox
'=====================================================================
Public Sub Load_CategoryListbox()
   CategoryList.Clear
   Debug.Print "Load_CategoryListbox"

   TmpString = ""
   TmpString = "SELECT CATEGORY_NAME FROM " & CATEGORIES_TABLENAME & _
         " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
         " ORDER BY CATEGORY_NAME ASC"

   'Set Recordset to new
   Set FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY = New ADODB.Recordset

   FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.Open TmpString, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   'Refresh
   FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.Requery

   If FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.RecordCount > 0 Then
      Do While Not FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.EOF
         CategoryList.AddItem FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.Fields("CATEGORY_NAME").Value
         FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.MoveNext
      Loop
   End If
   FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.Close

   If CategoryList.ListCount > 0 Then
      CategoryList.Text = FAMILY.FORM_MAIN.ComboBox(1).Text
   End If

End Sub
'=====================================================================


'=====================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If CurrentMode <> "" Then
      Cancel = True
   Else
      Call btnExit_Click
   End If
End Sub
'=====================================================================


'=====================================================================
Private Sub txtGroup_GotFocus()
   txtGroup.SelStart = 0
   txtGroup.SelLength = Len(txtGroup.Text)
End Sub
'=====================================================================


'=====================================================================
Private Sub txtGroup_KeyPress(KeyAscii As Integer)
   'Validate Character Entered
   If modPublic.MyChar(Chr(KeyAscii), "-", True) = False Then
      If (KeyAscii <> vbKeyBack) Then
         KeyAscii = 0
      End If
   End If
End Sub
'=====================================================================


'=====================================================================
'ChangeButtons
'=====================================================================
Private Sub ChangeButtons(ByVal ADD_BUTTON As Boolean, _
                          ByVal EDIT_BUTTON As Boolean, _
                          ByVal DELETE_BUTTON As Boolean, _
                          ByVal EXIT_BUTTON As Boolean)

   btnAdd.Enabled = ADD_BUTTON
   btnEdit.Enabled = EDIT_BUTTON
   btnDelete.Enabled = DELETE_BUTTON
   btnExit.Enabled = EXIT_BUTTON
End Sub
'=====================================================================
