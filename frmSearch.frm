VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::  Search / Printing - Records"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7560
      Top             =   3120
   End
   Begin Family_v3.TrayArea TrayArea1 
      Left            =   7560
      Top             =   2760
      _ExtentX        =   635
      _ExtentY        =   397
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   15
      Top             =   4950
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9260
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
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   7200
      TabIndex        =   14
      Top             =   4500
      Width           =   920
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CheckBox chkExact 
         Caption         =   "Search Exact Phrase"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4750
         TabIndex        =   21
         Top             =   880
         Width           =   1935
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7440
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearch.frx":0442
               Key             =   "person1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearch.frx":0D1E
               Key             =   "person2"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearch.frx":15FA
               Key             =   "mail"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearch.frx":1A4E
               Key             =   "group"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   50
         ScaleHeight     =   255
         ScaleWidth      =   3045
         TabIndex        =   19
         Top             =   120
         Width           =   3045
         Begin Family_v3.Label3D Label3D2 
            Height          =   255
            Left            =   45
            TabIndex        =   20
            Top             =   40
            Width           =   3090
            _ExtentX        =   5450
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
            Caption         =   ":: Search / Printing - Records ..."
         End
      End
      Begin VB.ComboBox ComboBox 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Index           =   1
         Left            =   4750
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "Sex Combobox"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   50
         TabIndex        =   16
         Tag             =   "Last Name Textbox"
         Top             =   800
         Width           =   2650
      End
      Begin VB.CheckBox chkListAll 
         Caption         =   "List All Records"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4750
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkRelCat 
         Caption         =   "Relation / Category"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2720
         TabIndex        =   8
         Top             =   1300
         Width           =   1905
      End
      Begin VB.CheckBox chkSex 
         Caption         =   "Sex"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6240
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox ComboBox 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Index           =   0
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Relation / Category Combobox"
         Top             =   1200
         Width           =   2650
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "&Search"
         Height          =   345
         Left            =   7150
         TabIndex        =   5
         Top             =   1200
         Width           =   915
      End
      Begin VB.CheckBox chkLastName 
         Caption         =   "Last Name"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2720
         TabIndex        =   4
         Top             =   880
         Width           =   1335
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   0
         Left            =   50
         TabIndex        =   3
         Tag             =   "First Name Textbox"
         Top             =   400
         Width           =   2650
      End
      Begin VB.CheckBox chkFirstName 
         Caption         =   "First Name"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2720
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   2580
         Left            =   45
         TabIndex        =   1
         Top             =   1680
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   4551
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   8454143
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   7400
         Picture         =   "frmSearch.frx":1FEA
         Top             =   550
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Printing Records"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   30
      TabIndex        =   10
      Top             =   4320
      Width           =   4335
      Begin VB.OptionButton opt_All_Listed 
         Caption         =   "All Listed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2040
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton opt_All_Checked 
         Caption         =   "All Checked"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin Family_v3.Label3D Label3D1 
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   4560
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   ":: Search / Printing Record(s) ::"
      Phase           =   1
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim TheKey As String
Dim lvListItems As ListItem

Private Sub btnClose_Click()
   'SetUp FrmMain Texboxes and Comboboxes
   Call Setup_TextBox(FAMILY.FORM_MAIN, True, True)
   Call Setup_ComboBox(FAMILY.FORM_MAIN, True, True)
   Call FAMILY.FORM_MAIN.Change_Button(False, False, True, True, False, False)

   'Unload The Form
   Unload Me
End Sub

Private Sub btnPrint_Click()

   If opt_All_Checked = True Then   'IF The User wants to print only the records selected

      Dim i As Integer
      Dim i2 As Integer
      Dim ItemChecked As Boolean

      ItemChecked = False

      For i = 1 To lvSearch.ListItems.Count
         If lvSearch.ListItems(i).Checked = True Then
            ItemChecked = True
            Exit For
         End If
      Next i

      If ItemChecked = True Then
         Printer.Font = "Times New Roman"
         Printer.FontBold = False
         Printer.FontUnderline = True
         Printer.FontSize = 10
         Printer.Print vbNewLine
         PrintCenter (CURRENT_USER.FULL_NAME & "'s " & App.ProductName)
         Printer.FontUnderline = False
         Printer.FontBold = False
         Printer.Print vbNewLine

         i2 = 0

         For i = 1 To lvSearch.ListItems.Count
            If lvSearch.ListItems(i).Checked = True Then
               i2 = i2 + 1
               Printer.Print Space(6) & "Record #" & Str$(i2)
               Printer.Print Space(6) & "Name : " & ProperCase(lvSearch.ListItems(i).Text & " " & lvSearch.ListItems(i).ListSubItems(1).Text)
               Printer.Print Space(6) & "Sex : " & lvSearch.ListItems(i).ListSubItems(2).Text
               Printer.Print Space(6) & "Telephone : " & lvSearch.ListItems(i).ListSubItems(3).Text
               Printer.Print Space(6) & "Address : " & lvSearch.ListItems(i).ListSubItems(4).Text
               Printer.Print Space(6) & "City-State-ZipCode : " & lvSearch.ListItems(i).ListSubItems(5).Text & " " & lvSearch.ListItems(i).ListSubItems(6).Text
               Printer.Print Space(6) & "Email Address : " & lvSearch.ListItems(i).ListSubItems(7).Text
               Printer.Print vbNewLine
            End If
         Next i
         Printer.EndDoc
      End If

   Else   ' The User Wants To Prin All The Records Listed Selected

      If lvSearch.ListItems.Count > 0 Then
         Printer.Font = "Times New Roman"
         Printer.FontBold = False
         Printer.FontUnderline = True
         Printer.FontSize = 10
         Printer.Print vbNewLine
         PrintCenter (CURRENT_USER.FULL_NAME & "'s " & App.ProductName)
         Printer.FontUnderline = False
         Printer.FontBold = False
         Printer.Print vbNewLine

         For i = 1 To lvSearch.ListItems.Count
            Printer.Print Space(6) & "Record #" & Str$(i)
            Printer.Print Space(6) & "Name : " & ProperCase(lvSearch.ListItems(i).Text & " " & lvSearch.ListItems(i).ListSubItems(1).Text)
            Printer.Print Space(6) & "Sex : " & lvSearch.ListItems(i).ListSubItems(2).Text
            Printer.Print Space(6) & "Telephone : " & lvSearch.ListItems(i).ListSubItems(3).Text
            Printer.Print Space(6) & "Address : " & lvSearch.ListItems(i).ListSubItems(4).Text
            Printer.Print Space(6) & "City-State-ZipCode : " & lvSearch.ListItems(i).ListSubItems(5).Text & " " & lvSearch.ListItems(i).ListSubItems(6).Text
            Printer.Print Space(6) & "Email Address : " & lvSearch.ListItems(i).ListSubItems(7).Text
            Printer.Print vbNewLine
         Next i
         Printer.EndDoc
      End If
   End If

End Sub



Private Sub btnSearch_Click()
   Dim tmpSearch   As String
   On Error GoTo SEARCH_ERROR

   tmpSearch = ""
   lvSearch.ListItems.Clear

   If chkListAll.Value = 1 Then
      tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & " ORDER BY LastName" & _
            " AND USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'"

   Else

      Select Case chkExact.Value

         Case Is = 0   'The User does not want to search for the exact value

            If chkFirstName.Value = 1 Then   'If First Name was checked
               If Len(Trim$(TextBox(0).Text)) > 0 Then
                  tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & _
                        " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
                        " AND FirstName LIKE '%" & Apostrophe(TextBox(0).Text) & "%'"
               Else   'If First Name was not checked
                  tmpSearch = ""
                  MsgBox CURRENT_USER.FULL_NAME & ", you will have to enter a value in the " & TextBox(0).Tag & " since you have it checked.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If

            If chkLastName.Value = 1 Then   'If Last Name was checked
               If Len(Trim$(TextBox(1).Text)) > 0 Then   'Check If values were previously entered
                  If Len(tmpSearch) > 0 Then
                     tmpSearch = tmpSearch & " AND LastName LIKE '%" & Apostrophe(TextBox(1).Text) & "%'"
                  Else
                     tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & _
                           " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
                           " AND LastName LIKE '%" & Apostrophe(TextBox(1).Text) & "%'"
                  End If

               Else
                  tmpSearch = ""
                  MsgBox CURRENT_USER.FULL_NAME & ", you will have to enter a value in the " & TextBox(1).Tag & " since you have it checked.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If


            If chkRelCat.Value = 1 Then   'Checck If Relationsip/Category was checked
               If Len(ComboBox(0).Text) > 0 Then   'Check If values were previously entered
                  If Len(tmpSearch) > 0 Then
                     tmpSearch = tmpSearch & " AND CATEGORY_NAME = '" & Apostrophe(ComboBox(0).Text) & "'"
                  Else
                     tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & _
                           " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
                           " AND CATEGORY_NAME = '" & Apostrophe(ComboBox(0).Text) & "'"
                  End If

               Else
                  tmpSearch = ""
                  MsgBox CURRENT_USER.FULL_NAME & ", you will have to enter a value in the " & ComboBox(0).Tag & " since you have it checked.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If


            If chkSex.Value = 1 Then   'Check If Sex wax selected
               If Len(ComboBox(1).Text) > 0 Then   'Check If values were previously entered
                  If Len(tmpSearch) > 0 Then
                     tmpSearch = tmpSearch & " AND Sex = '" & Apostrophe(ComboBox(1).Text) & "'"
                  Else
                     tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & _
                           " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
                           " AND Sex = '" & Apostrophe(ComboBox(1).Text) & "'"
                  End If

               Else
                  tmpSearch = ""
                  MsgBox CURRENT_USER.FULL_NAME & ", you will have to enter a value in the " & ComboBox(1).Tag & " since you have it checked.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If



         Case Is = 1   'The User Want The Search Result To Be Exact

            If chkFirstName.Value = 1 Then   'Check If First Name was Checked
               If Len(Trim$(TextBox(0).Text)) > 0 Then   'Check If values were previously entered
                  tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & _
                        " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
                        " AND FirstName = '" & Apostrophe(TextBox(0).Text) & "'"

               Else
                  tmpSearch = ""
                  MsgBox CURRENT_USER.FULL_NAME & ", you will have to enter a value in the " & TextBox(0).Tag & " since you have it checked.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If

            If chkLastName.Value = 1 Then   'Check IF Last Name was Checked
               If Len(Trim$(TextBox(1).Text)) > 0 Then   'Check If values were previously entered
                  If Len(tmpSearch) > 0 Then
                     tmpSearch = tmpSearch & " AND LastName = '" & Apostrophe(TextBox(1).Text) & "'"
                  Else
                     tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & _
                           " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
                           " AND LastName = '" & Apostrophe(TextBox(1).Text) & "'"
                  End If

               Else
                  tmpSearch = ""
                  MsgBox CURRENT_USER.FULL_NAME & ", you will have to enter a value in the " & TextBox(1).Tag & " since you have it checked.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If


            If chkRelCat.Value = 1 Then   'Check If Relationship/Category was selected
               If Len(ComboBox(0).Text) > 0 Then
                  If Len(tmpSearch) > 0 Then   'Check If values were previously entered
                     tmpSearch = tmpSearch & " AND CATEGORY_NAME = '" & Apostrophe(ComboBox(0).Text) & "'"
                  Else
                     tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & _
                           " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
                           " AND CATEGORY_NAME = '" & Apostrophe(ComboBox(0).Text) & "'"
                  End If

               Else
                  tmpSearch = ""
                  MsgBox CURRENT_USER.FULL_NAME & ", you will have to enter a value in the " & ComboBox(0).Tag & " since you have it checked.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If


            If chkSex.Value = 1 Then   'Check if Sex was selected
               If Len(ComboBox(1).Text) > 0 Then
                  If Len(tmpSearch) > 0 Then   'Check If values were previously entered
                     tmpSearch = tmpSearch & " AND Sex = '" & Apostrophe(ComboBox(1).Text) & "'"
                  Else
                     tmpSearch = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,EmailAddress FROM " & CONTACTS_TABLENAME & _
                           " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
                           " AND Sex = '" & Apostrophe(ComboBox(1).Text) & "'"
                  End If

               Else
                  tmpSearch = ""
                  MsgBox CURRENT_USER.FULL_NAME & ", you will have to enter a value in the " & ComboBox(1).Tag & " since you have it checked.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If


      End Select

   End If



   If Len(tmpSearch) < 1 Then
      Exit Sub
   End If


   Set TmpRecordSet = New ADODB.Recordset
   TmpRecordSet.Open tmpSearch, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic


   If TmpRecordSet.RecordCount > 0 Then
      Do While Not TmpRecordSet.EOF
         TheKey = TmpRecordSet.Fields("USER_NAME") & "_" & _
               ProperCase(TmpRecordSet.Fields("FirstName")) & "_" & _
               ProperCase(TmpRecordSet.Fields("LastName")) & "_" & _
               ProperCase(TmpRecordSet.Fields("CATEGORY_NAME"))

         If TmpRecordSet.Fields("Sex") = "Male" Then
            Set lvListItems = lvSearch.ListItems.Add(, TheKey, ProperCase(TmpRecordSet.Fields("FirstName")), "person1", "person1")
         Else
            Set lvListItems = lvSearch.ListItems.Add(, TheKey, ProperCase(TmpRecordSet.Fields("FirstName")), "person2", "person2")
         End If

         'LastName
         lvListItems.SubItems(1) = ProperCase(TmpRecordSet.Fields("LastName"))
         'Sex
         lvListItems.SubItems(2) = ProperCase(TmpRecordSet.Fields("Sex"))
         'Telephone
         lvListItems.SubItems(3) = TmpRecordSet.Fields("Telephone")
         'Address
         lvListItems.SubItems(4) = ProperCase(TmpRecordSet.Fields("Address"))
         'City, State
         lvListItems.SubItems(5) = ProperCase(TmpRecordSet.Fields("City_State"))
         'Zip Code
         lvListItems.SubItems(6) = TmpRecordSet.Fields("ZipCode")
         'Email Address
         lvListItems.SubItems(7) = TmpRecordSet.Fields("EmailAddress")
         'Category
         lvListItems.SubItems(8) = ProperCase(TmpRecordSet.Fields("CATEGORY_NAME"))

         TmpRecordSet.MoveNext
      Loop
      Set TmpRecordSet = Nothing
   Else
      StatusBar1.Panels(1).Text = "No Records Found"
   End If
   Exit Sub


SEARCH_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR frmSearch.btnSearch" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Error#" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Sub

Private Sub chkFirstName_Click()
   If chkFirstName.Value = 1 Then
      chkListAll.Value = 0
   End If
End Sub

Private Sub chkLastName_Click()
   If chkLastName.Value = 1 Then
      chkListAll.Value = 0
   End If
End Sub

Private Sub chkListAll_Click()
   If chkListAll.Value = 1 Then
      chkFirstName.Value = 0
      chkLastName.Value = 0
      chkRelCat.Value = 0
   End If
End Sub

Private Sub chkRelCat_Click()
   If chkRelCat.Value = 1 Then
      chkListAll.Value = 0
   End If
End Sub



Private Sub Form_Load()
On Error GoTo FORM_LOAD_ERROR

   'Select option All Listed
   opt_All_Listed.Value = True

   lvSearch.View = lvwReport
   Call modListview.Load_Search_Lisview(lvSearch)

   tmpSQL = ""
   tmpSQL = "SELECT CATEGORY_NAME FROM " & CATEGORIES_TABLENAME & _
         " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "' ORDER BY CATEGORY_NAME ASC"

   Set TmpRecordSet = New ADODB.Recordset
   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   ComboBox(0).Clear
   If TmpRecordSet.RecordCount > 0 Then
      Do While Not TmpRecordSet.EOF
         ComboBox(0).AddItem TmpRecordSet.Fields("CATEGORY_NAME").Value
         TmpRecordSet.MoveNext
      Loop
      ComboBox(0).ListIndex = 0
   End If
   tmpSQL = ""
   Set TmpRecordSet = Nothing
   ComboBox(1).Clear
   ComboBox(1).AddItem "Female"
   ComboBox(1).AddItem "Male"
   ComboBox(1).ListIndex = 0
   
FORM_LOAD_ERROR:
      If Err.Number <> 0 Then
      MsgBox "ERROR frmSearch.Form_Load" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Error#" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If

End Sub


Private Sub lvSearch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   'Used to sort the items in the list
   'Found this somewhere on the web

   With lvSearch
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


Private Sub lvSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim tmpItem As ListItem
   Set tmpItem = lvSearch.HitTest(X, Y)

   If tmpItem Is Nothing Then
      StatusBar1.Panels(1).Text = ""
   Else
      StatusBar1.Panels(1).Text = tmpItem.Text
   End If
End Sub
