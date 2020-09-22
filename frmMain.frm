VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Family Address Book v3.0"
   ClientHeight    =   5235
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   640
      Left            =   50
      TabIndex        =   24
      Top             =   4200
      Width           =   3030
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tlbBtnReminders"
               Object.ToolTipText     =   " Reminders ... "
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tlbBtnSearch"
               Object.ToolTipText     =   " Search/Print Records ... "
               ImageIndex      =   10
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tlbBtnEmail"
               Object.ToolTipText     =   " Send Email ... "
               ImageIndex      =   9
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tlbBtnUserProfile"
               Object.ToolTipText     =   " Click To View/Edit User Profile ... "
               ImageIndex      =   11
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   " ::  Update Your Personal Internet Links  :: "
               ImageIndex      =   14
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame NotesFrame 
      Height          =   855
      Left            =   8760
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame FieldFrame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3260
      TabIndex        =   19
      Top             =   420
      Width           =   5300
      Begin VB.CommandButton btnCancel 
         Caption         =   "&Cancel"
         Height          =   310
         Left            =   4080
         TabIndex        =   10
         ToolTipText     =   " ::  Cancel Changes  :: "
         Top             =   960
         Width           =   900
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "&Save"
         Height          =   310
         Left            =   3000
         TabIndex        =   9
         ToolTipText     =   " ::  Save Changes  :: "
         Top             =   960
         Width           =   900
      End
      Begin VB.ComboBox ComboBox 
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
         Height          =   345
         Index           =   1
         Left            =   3200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3150
         Width           =   2000
      End
      Begin VB.TextBox TextBox 
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
         Height          =   285
         Index           =   6
         Left            =   1100
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2760
         Width           =   4100
      End
      Begin VB.TextBox TextBox 
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
         Height          =   285
         Index           =   5
         Left            =   1100
         MaxLength       =   11
         TabIndex        =   6
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox TextBox 
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
         Height          =   285
         Index           =   4
         Left            =   1100
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2040
         Width           =   4100
      End
      Begin VB.TextBox TextBox 
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
         Height          =   285
         Index           =   3
         Left            =   1100
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1680
         Width           =   4100
      End
      Begin VB.TextBox TextBox 
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
         Height          =   285
         Index           =   2
         Left            =   1100
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1320
         Width           =   4100
      End
      Begin VB.TextBox TextBox 
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
         Height          =   285
         Index           =   1
         Left            =   1100
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Last Name Textbox"
         Top             =   600
         Width           =   4100
      End
      Begin VB.ComboBox ComboBox 
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
         Height          =   345
         Index           =   0
         Left            =   1100
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   930
         Width           =   1600
      End
      Begin VB.TextBox TextBox 
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
         Height          =   285
         Index           =   0
         Left            =   1100
         MaxLength       =   30
         TabIndex        =   0
         Tag             =   "First Name Textbox"
         Top             =   240
         Width           =   4100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Relation/Category"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   8
         Left            =   1100
         TabIndex        =   35
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Em@il"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   7
         Left            =   100
         TabIndex        =   32
         Top             =   2850
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Zip Code"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   6
         Left            =   100
         TabIndex        =   31
         Top             =   2500
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "City, State"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   5
         Left            =   100
         TabIndex        =   30
         Top             =   2140
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   4
         Left            =   100
         TabIndex        =   29
         Top             =   1790
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telephone"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   3
         Left            =   100
         TabIndex        =   28
         Top             =   1400
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   2
         Left            =   100
         TabIndex        =   27
         Top             =   1080
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Name"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   1
         Left            =   100
         TabIndex        =   26
         Top             =   700
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "First Name"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   20
         Top             =   300
         Width           =   855
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4100
      Left            =   3150
      TabIndex        =   18
      Top             =   70
      Width           =   5500
      _ExtentX        =   9710
      _ExtentY        =   7223
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contacts"
            Key             =   "Contacts"
            Object.ToolTipText     =   " ::  Contacts  :: "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Memo/Notes"
            Key             =   "memo"
            Object.ToolTipText     =   " ::  Memo/Notes  :: "
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "people1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EDE
            Key             =   "person1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17BA
            Key             =   "person2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2096
            Key             =   "book1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21F2
            Key             =   "open_book"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":234E
            Key             =   "closed_book"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":307E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3ED2
            Key             =   "users"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":499E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":567A
            Key             =   "ie1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F56
            Key             =   "ie2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   17
      Top             =   4935
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   10636
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4150
      Left            =   30
      TabIndex        =   15
      Top             =   20
      Width           =   3050
      Begin Family_v3.TrayArea TrayArea1 
         Left            =   840
         Top             =   3840
         _ExtentX        =   635
         _ExtentY        =   397
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   2925
         TabIndex        =   33
         Top             =   140
         Width           =   2920
         Begin Family_v3.Label3D Label3D1 
            Height          =   255
            Left            =   30
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   20
            Width           =   1305
            _ExtentX        =   2302
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
            Caption         =   ":: People ..."
            BackColor       =   -2147483637
         End
      End
      Begin MSComctlLib.TreeView UsersTreeView 
         Height          =   3690
         Left            =   50
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   405
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   6509
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   640
      Left            =   3260
      TabIndex        =   21
      Top             =   4200
      Width           =   5295
      Begin VB.CommandButton btnExit 
         Caption         =   "E&xit"
         Height          =   310
         Left            =   4200
         TabIndex        =   14
         ToolTipText     =   " ::  Exit  :: "
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Height          =   320
         Left            =   130
         TabIndex        =   11
         ToolTipText     =   " ::  Edit The Current Record  :: "
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Height          =   320
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   " ::  Remove Current Record  :: "
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   320
         Left            =   1200
         TabIndex        =   12
         ToolTipText     =   " ::  Add A New Record  :: "
         Top             =   240
         Width           =   950
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu we 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrganizeCategory 
         Caption         =   "&Organize Relation/Category ..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuReminder 
         Caption         =   "&Reminder"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSearchPrint 
         Caption         =   "&Search/Print Record(s) ..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuUserProfile 
         Caption         =   "User &Profile .."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuRelation_Category 
         Caption         =   "Relation/Category"
         Visible         =   0   'False
         Begin VB.Menu dfg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGroup 
            Caption         =   "Add, Rename or Delete  &Group"
            Shortcut        =   ^G
         End
         Begin VB.Menu rtyr 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditRecord 
            Caption         =   "Edit &This Record"
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuAddRecord 
            Caption         =   "Add A &New Record"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuSendEmail1 
            Caption         =   "Send &Email ..."
         End
         Begin VB.Menu kugb 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuEmails 
         Caption         =   "Email"
         Begin VB.Menu jtrufu 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSendEmail 
            Caption         =   "Send &Email ..."
            Shortcut        =   ^E
         End
         Begin VB.Menu gy 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCMail 
            Caption         =   "&Configure Email ..."
         End
         Begin VB.Menu jiugmgk 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuConfigurations 
         Caption         =   "Configurations"
         Begin VB.Menu cdfhdfhfgn 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTray 
            Caption         =   "&Minimize To System Tray"
         End
         Begin VB.Menu ujyfujvu 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuLinks 
         Caption         =   "&Personal Links"
         Begin VB.Menu qwedf 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUpdateLinks 
            Caption         =   "&Update Links ..."
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuLink 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   1
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   2
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   3
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   4
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   5
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   6
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   7
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   8
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   9
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link"
            Index           =   10
         End
         Begin VB.Menu ertert 
            Caption         =   "-"
         End
      End
      Begin VB.Menu yui 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
      Begin VB.Menu yjtyju 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuTreeView 
      Caption         =   "TV"
      Visible         =   0   'False
      Begin VB.Menu aa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd_Category 
         Caption         =   "&Add Category"
      End
      Begin VB.Menu a2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove_Category 
         Caption         =   "&Remove Category"
      End
      Begin VB.Menu a3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd_Record 
         Caption         =   "Add Record"
      End
      Begin VB.Menu a4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove_Record 
         Caption         =   "Re&move Record"
      End
      Begin VB.Menu a5 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnHlp 
      Caption         =   "&Help"
      Begin VB.Menu qwe 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu yfh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmailAuthor 
         Caption         =   "Email The &Author"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuVistMyWeb 
         Caption         =   "&Visit The Author's Webpage"
         Shortcut        =   ^V
      End
      Begin VB.Menu ftytty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportIt 
         Caption         =   "&Report Problems"
      End
      Begin VB.Menu sdsfdf 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuTrayPop 
      Caption         =   "mnuTrayPop"
      Visible         =   0   'False
      Begin VB.Menu drftgr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu fthrt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout1 
         Caption         =   "&About Family Address Book v3.0 "
      End
      Begin VB.Menu srwer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "E&xit"
      End
      Begin VB.Menu erter 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'===================================================================
'Variable Declaration
'===================================================================
Public EDIT_MODE                   As Boolean   'EDIT MODE
Public CURRENT_STATE               As String   'CURRENT STATE
Public RECORD_SELECTED             As Boolean   'Tells Is A Record is Selected
'
Public Last_Parent_Node_Clicked    As String   'Last Parent Node Clicked
Public Last_Node_Clicked           As String   'Last Node Clicked
'
Public WithEvents MyUser           As clsContacts   ' THE USER'S CLASS
Attribute MyUser.VB_VarHelpID = -1

Public OLD_FIRST_NAME              As String
Public OLD_LAST_NAME               As String
Public OLD_CATEGORY                As String

'===================================================================
'===================================================================



'===================================================================
Private Sub btnAdd_Click()

   'This Checks Is The User Has One Or More Categories
   If UsersTreeView.Nodes("ROOT").Children < 1 Then
      TmpMsgResult = MsgBox("Hi " & CURRENT_USER.LOGIN_NAME & ", before you can add someone to you Contacts Database" & vbCrLf & _
            "you must to have one (1)  or more Category or Relationship Groups." & vbCrLf & vbCrLf & _
            "So " & CURRENT_USER.LOGIN_NAME & ", would like to create a new Category or Relationship Group ?", vbInformation + vbYesNo)
      If TmpMsgResult = vbYes Then
         Call mnuGroup_Click
         Exit Sub
      Else
         Exit Sub
      End If
   End If

   'Set Current_State = "Adding"
   CURRENT_STATE = "Adding"
   'Set Making_Changes To True
   Making_Changes True
   'Set Edit_Mode TO True
   EDIT_MODE = True
   'Set Record_Selected to False
   RECORD_SELECTED = False

   'Unlock The Textboxes and Comboboxes
   Call Setup_TextBox(Me, True, False)
   Call Setup_ComboBox(Me, False, False)
   Call Change_Button(False, False, False, False, True, True)

   'Disable User Access To The TreeView
   UsersTreeView.Enabled = False

   If TabStrip1.Tabs(2).Selected = True Then
      TabStrip1.Tabs(1).Selected = True
      TabStrip1_Click
   End If
   'Set Focus To The Textbox(0) - FirstName
   TextBox(0).SetFocus

   'Disable The Toolbar
   Toolbar1.Enabled = False
End Sub
'===================================================================


'===================================================================
Private Sub btnAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Click to add a new record"
End Sub
'===================================================================


'===================================================================
Private Sub btnCancel_Click()
   CURRENT_STATE = ""
   EDIT_MODE = False
   Making_Changes False
   'Lock the Textbox and comboboxes
   Call Setup_TextBox(Me, True, True)
   Call Setup_ComboBox(Me, True, False)
   Call Change_Button(False, False, True, True, False, False)

   UsersTreeView.Enabled = True
   'Enable The Toolbar
   Toolbar1.Enabled = True
End Sub
'===================================================================


'===================================================================
Private Sub btnCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Click to cancel changes made"
End Sub
'===================================================================


'===================================================================
Private Sub btnEdit_Click()
   'Make Sure That A Record Is Selected
   If RECORD_SELECTED = True Then

      'Set Current_State = "Editing"
      CURRENT_STATE = "Editing"
      EDIT_MODE = True
      Making_Changes True

      'Unlock the Textbox and comboboxes
      Call Setup_TextBox(Me, False, False)
      Call Setup_ComboBox(Me, False, False)
      Call Change_Button(False, False, False, False, True, True)

      If TabStrip1.Tabs(2).Selected = True Then
         TabStrip1.Tabs(1).Selected = True
         TabStrip1_Click
      End If

      TextBox(0).SetFocus
      Toolbar1.Enabled = False
      UsersTreeView.Enabled = False

      OLD_FIRST_NAME = TextBox(0).Text
      OLD_LAST_NAME = TextBox(1).Text
      OLD_CATEGORY = ComboBox(1).Text
   Else
      MsgBox "No Record Was Selected.", vbInformation + vbOKOnly

      'Reload TreeVIEW
      Call LOAD_frmMAIN_TREEVIEW(Me, Me.UsersTreeView, ImageList1, Last_Parent_Node_Clicked)
      DoEvents

      'Lock The Texboxes and Comboboxes
      Call Setup_TextBox(Me, True, True)
      Call Setup_ComboBox(Me, True, True)
      Call Change_Button(False, False, True, True, False, False)
   End If
End Sub
'===================================================================


'===================================================================
Private Sub btnEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Click To Edit The Current Record"
End Sub
'===================================================================


'===================================================================
Public Sub btnExit_Click()
   If FAMILY.FORM_MAIN.CURRENT_STATE = "" Then
      If FAMILY.FORM_USERS.MyUser.LOGOUT_USER(CURRENT_USER.LOGIN_NAME) = False Then
         MsgBox CURRENT_USER.FULL_NAME & ", I was unable to properly log you out. Read the file Readme.doc", vbCritical + vbOKOnly
      Else
         Debug.Print "You were logged out properly. ;-)"
      End If

      'ShutDown
      Unload Me
      Call modPublic.ShutDown
      End
   End If
End Sub
'===================================================================


'===================================================================
Private Sub btnExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Click to Exit"
End Sub
'===================================================================




'===================================================================
Private Sub btnDelete_Click()
   TmpMsgResult = MsgBox("Hi " & CURRENT_USER.FULL_NAME & ", do you want go to delete [ " & TextBox(0).Text & " " & TextBox(1).Text & " ] from your database?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Record")
   If TmpMsgResult = vbYes Then
      'Delete The Current Record
      MyUser.USER_CONTACTS.Delete
      'Refresh it
      MyUser.USER_CONTACTS.Requery

      'Reload TreeVIEW
      If LOAD_frmMAIN_TREEVIEW(FAMILY.FORM_MAIN, FAMILY.FORM_MAIN.UsersTreeView, FAMILY.FORM_MAIN.ImageList1, FAMILY.FORM_MAIN.Last_Parent_Node_Clicked) <> True Then
         MsgBox "Hi " & CURRENT_USER.FULL_NAME & ", the TreeView was not properly reloaded.", vbCritical + vbOKOnly, "Reload Error"
      End If

      'Setup Textboxes,comboboxes and buttons
      Call Setup_TextBox(Me, True, True)
      Call Setup_ComboBox(Me, True, False)
      Call Change_Button(False, False, True, True, False, False)
   End If
End Sub
'===================================================================


'===================================================================
Private Sub btnDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Click To Delete The Current Record"
End Sub
'===================================================================


'===================================================================
Private Sub btnSave_Click()
   Dim i As Byte
   Dim i2 As Byte
   Dim i3 As Byte
   'Validate Data Entered
   If Len(Trim$(TextBox(0).Text)) < MIN_FNAME_SIZE Then
      MsgBox "Hi " & CURRENT_USER.FULL_NAME & ", the First Name box should be less than or equal to " & Str$(MAX_FNAME_SIZE) & "." & vbNewLine & _
            "But greater than or equal to " & Str$(MIN_FNAME_SIZE) & ".", vbInformation + vbOKOnly
      Exit Sub
   End If

   'Check The Last Name
   If Len(Trim$(TextBox(1).Text)) < 1 Then
      TextBox(1).Text = "NotSpecified"
   End If

   'Check First Name and Last Name for invalid characters
   'Note : I have already stored some information the first and
   '       Last Name Textbox
   For i = LBound(IllegalChars) To UBound(IllegalChars)
      For i2 = 0 To 1
         For i3 = 1 To Len(TextBox(i2))
            If Chr$(IllegalChars(i)) = Mid$(TextBox(i2), i3, 1) Then
               MsgBox CURRENT_USER.FULL_NAME & ", you have an invalid charater [ " & Chr$(IllegalChars(i)) & " ] , in the " & TextBox(i2).Tag, vbCritical + vbOKOnly
               Exit Sub
            End If
         Next i3
      Next i2
   Next i


   'Set Last_Parent_Node_Clicked to the tmem in combobox(1)- category
   Last_Parent_Node_Clicked = ComboBox(1).Text

   Select Case CURRENT_STATE
      Case "Adding"

         'First Check If the record already exist
         If FAMILY.FORM_MAIN.MyUser.RECORD_EXIST( _
               TextBox(0).Text, _
               TextBox(1).Text, _
               ComboBox(1).Text, _
               CURRENT_USER.LOGIN_NAME) = True Then

            MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", that record already exist.", vbInformation + vbOKOnly
            Exit Sub
         End If


         If FAMILY.FORM_MAIN.MyUser.ADD_RECORD( _
               TextBox(0).Text, _
               TextBox(1).Text, _
               ComboBox(0).Text, _
               TextBox(2).Text, _
               TextBox(3).Text, _
               TextBox(4).Text, _
               TextBox(5).Text, _
               TextBox(6).Text, _
               ComboBox(1).Text, _
               txtNotes.Text, _
               CURRENT_USER.LOGIN_NAME) = True Then
            Debug.Print "Record Added Successfully"
         Else
            MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", but I wasn unable to add a new record.", vbInformation + vbOKOnly
         End If


      Case "Editing"

         'Make Sure That A Record Is Selected
         If RECORD_SELECTED = True Then
            'Check If Information was changed
            If (OLD_FIRST_NAME <> TextBox(0).Text) Or (OLD_LAST_NAME <> TextBox(1).Text) Or (OLD_CATEGORY <> ComboBox(1).Text) Then
               'Check If The Record Already Exist
               If FAMILY.FORM_MAIN.MyUser.RECORD_EXIST( _
                     TextBox(0).Text, _
                     TextBox(1).Text, _
                     ComboBox(1).Text, _
                     CURRENT_USER.LOGIN_NAME) = True Then

                  MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", but that record already exist.", vbInformation + vbOKOnly
                  Exit Sub
               End If
            End If


            'MyFamily.USER_CONTACTS
            Debug.Print "Editmode"

            'First Name
            MyUser.USER_CONTACTS.Fields("FirstName") = "" & ProperCase(Trim$(TextBox(0).Text))
            'Last Name
            MyUser.USER_CONTACTS.Fields("LastName") = "" & ProperCase(Trim$(TextBox(1).Text))
            'Sex
            MyUser.USER_CONTACTS.Fields("Sex") = Trim$(ComboBox(0).Text)
            'Telephone
            MyUser.USER_CONTACTS.Fields("Telephone") = "" & Trim$(TextBox(2).Text)
            'Address
            MyUser.USER_CONTACTS.Fields("Address") = "" & ProperCase(Trim$(TextBox(3).Text))
            'City-State
            MyUser.USER_CONTACTS.Fields("City_State") = "" & ProperCase(Trim$(TextBox(4).Text))
            'Zip Code
            MyUser.USER_CONTACTS.Fields("ZipCode") = "" & Trim$(TextBox(5).Text)
            'Email
            MyUser.USER_CONTACTS.Fields("EmailAddress") = "" & Trim$(TextBox(6).Text)
            'Relation/Category
            MyUser.USER_CONTACTS.Fields("CATEGORY_NAME") = Trim$(ComboBox(1).Text)
            'Notes
            MyUser.USER_CONTACTS.Fields("Notes") = Trim$(txtNotes.Text)
            'User_Name
            MyUser.USER_CONTACTS.Fields("USER_NAME") = "" & CURRENT_USER.LOGIN_NAME

            'Update Database
            MyUser.USER_CONTACTS.Update

            OLD_FIRST_NAME = ""
            OLD_LAST_NAME = ""
            OLD_CATEGORY = ""
         Else
            OLD_FIRST_NAME = ""
            OLD_LAST_NAME = ""
            OLD_CATEGORY = ""
            MsgBox "No Record Was Selected", vbInformation + vbOKOnly

            'Reload TreeVIEW
            Call LOAD_frmMAIN_TREEVIEW(Me, Me.UsersTreeView, ImageList1, Last_Parent_Node_Clicked)
            DoEvents

            'Lock The Texboxes and Comboboxes
            Call Setup_TextBox(Me, True, True)
            Call Setup_ComboBox(Me, True, True)
            Call Change_Button(False, False, True, True, False, False)
         End If

   End Select


   CURRENT_STATE = ""
   Making_Changes False
   EDIT_MODE = False
   RECORD_SELECTED = False

   'Reload TreeVIEW
   Call LOAD_frmMAIN_TREEVIEW(Me, Me.UsersTreeView, ImageList1, Last_Parent_Node_Clicked)
   DoEvents

   'Lock The Texboxes and Comboboxes
   Call Setup_TextBox(Me, True, True)
   Call Setup_ComboBox(Me, True, True)
   Call Change_Button(False, False, True, True, False, False)

   'Enable User Access To The Treeview
   UsersTreeView.Enabled = True

   'Enable The Toolbar
   Toolbar1.Enabled = True
   'Exit The Sub
   Exit Sub


btnSave_err:
   If Err.Number <> 0 Then
      MsgBox "Error : " & Err.Description & ". : #" & Str(Err.Number), vbCritical + vbOKOnly
      EDIT_MODE = False
      Err.Clear
      UsersTreeView.Enabled = True
      CURRENT_STATE = ""
   End If
End Sub
'===================================================================



'===================================================================
Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Click to save changes"
End Sub
'===================================================================


'===================================================================
Private Sub CurrentUser_ERR(ByVal ErrNum As Integer, ErrMsg As String)
   MsgBox (ErrMsg)
End Sub
'===================================================================


'===================================================================
Private Sub CurrentUser_MYDEBUG(ByVal Debug_Message As String)
   Debug.Print Debug_Message
End Sub
'===================================================================


Private Sub ComboBox_Change(Index As Integer)
   '
End Sub

'===================================================================
Private Sub FieldFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) And (CURRENT_STATE = "") Then
      PopupMenu mnuFile
   End If
End Sub
'===================================================================



'===================================================================
Private Sub Form_Load()
   Unload FAMILY.FORM_LOGIN

   'Check The Databse Connection State
   If PUBLIC_DATABASE.CONNECTION.STATE <> 1 Then
      MsgBox "DATABASE CONNECTION ERROR : You are not not connected to the database.", vbCritical + vbOKOnly
      Call btnExit_Click
   End If

   RECORD_SELECTED = False

   EDIT_MODE = False
   Making_Changes False

   Call Change_Button(False, False, True, True, False, False)
   'call modpublic.RemoveMenus(Me)

   '************
   'Set The Last_Parent_Node_Clicked to The First Item
   If ComboBox(1).ListCount > 0 Then
      Last_Parent_Node_Clicked = ComboBox(1).Text
   End If

   CURRENT_STATE = ""
End Sub
'===================================================================



'===================================================================
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) And (CURRENT_STATE = "") Then
      PopupMenu mnuFile
   End If
End Sub
'===================================================================


'===================================================================
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub
'===================================================================


'===================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If FAMILY.FORM_MAIN.CURRENT_STATE <> "" Then
      Cancel = True
   Else
      Call FAMILY.FORM_MAIN.btnExit_Click
   End If
End Sub
'===================================================================






'===================================================================
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub
'===================================================================


Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) And (CURRENT_STATE = "") Then
      If RECORD_SELECTED = True Then
         mnuEditRecord.Enabled = True
         PopupMenu mnuRelation_Category, , , , mnuEditRecord
      Else
         mnuEditRecord.Enabled = False
         PopupMenu mnuRelation_Category, , , , mnuAddRecord
      End If
   End If
End Sub
'===================================================================


'===================================================================
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub
'===================================================================


'===================================================================
Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) And (CURRENT_STATE = "") Then
      PopupMenu mnuFile
   End If
End Sub
'===================================================================


'===================================================================
Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub
'===================================================================



Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      PopupMenu mnuFile
   End If
End Sub


'===================================================================
Private Sub mnuAbout_Click()
   FAMILY.FORM_ABOUT.Show
End Sub
'===================================================================

'===================================================================
Private Sub mnuAbout1_Click()
   FAMILY.FORM_ABOUT.Show
End Sub
'===================================================================


'===================================================================
Private Sub mnuAddRecord_Click()
   Call btnAdd_Click
End Sub
'===================================================================


'===================================================================
Private Sub mnuCMail_Click()
   FAMILY.FORM_USERS.Show vbModal
End Sub
'===================================================================


Private Sub mnuEditRecord_Click()
   Call btnEdit_Click
End Sub

'===================================================================
Private Sub mnuExit_Click()
   Call btnExit_Click
End Sub
'===================================================================



Private Sub mnuExit2_Click()
   TrayArea1.Visible = False
   Call btnExit_Click
End Sub

Private Sub mnuGroup_Click()
   FAMILY.FORM_CATEGORIES.Show vbModal
End Sub

Private Sub mnuHelp_Click()
   MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", but no help is available at the moment.", vbInformation + vbOKOnly
End Sub



Private Sub mnuLink_Click(Index As Integer)
   'Execute The Link
   If (Len(mnuLink(Index).Caption) > 7) And (mnuLink(Index).Caption <> "Link") Then
      Call modPublic.OpenWebsite(mnuLink(Index).Caption)
   End If
End Sub

Private Sub mnuLinks_Click()
   Call CheckLinkMenus
End Sub

Private Sub mnuConfigurationS_Click()
   If MINIMIZE_TO_SYSTEM_TRAY = True Then
      mnuTray.Checked = True
   Else
      mnuTray.Checked = False
   End If
End Sub


Private Sub mnuOrganizeCategory_Click()
   Call mnuGroup_Click
End Sub

'===================================================================
Private Sub mnuReminder_Click()
   FAMILY.FORM_REMINDERS.Show vbModal
End Sub
'===================================================================


Private Sub mnuReportIt_Click()
'   FAMILY.FORM_SEND_EMAIL.RecipientName = AUTHOR_NAME
'   FAMILY.FORM_SEND_EMAIL.RecipientEmailAddress = AUTHOR_EMAIL_ADDRESS
'   FAMILY.FORM_SEND_EMAIL.EMAIL_SUBJECT = "ERROR FOUND IN " & App.ProductName & " ver" & Str$(App.Major) & "." & Str$(App.Minor)
'   FAMILY.FORM_SEND_EMAIL.EMAIL_BODY = "Hi, " & AUTHOR_NAME & vbNewLine & vbTab & "My name is " & CURRENT_USER.FULL_NAME & ". " & _
'         "I'm reporting an error that I found in " & App.ProductName & " ver" & Str$(App.Major) & "." & Str$(App.Minor) & " ."
'   FAMILY.FORM_SEND_EMAIL.Show
End Sub

Private Sub mnuRestore_Click()
   Call TrayArea1_DblClick
End Sub

'===================================================================
Private Sub mnuSearchPrint_Click()
   FAMILY.FORM_SEARCH.Show vbModal
End Sub
'===================================================================


'===================================================================
Private Sub mnuSendEmail_Click()
   If RECORD_SELECTED = True Then
      FAMILY.FORM_SEND_EMAIL.RecipientName = TextBox(0).Text & " " & TextBox(1).Text
      FAMILY.FORM_SEND_EMAIL.RecipientEmailAddress = TextBox(6).Text
      FAMILY.FORM_SEND_EMAIL.EMAIL_SUBJECT = "Hi " & TextBox(0).Text & ","
      FAMILY.FORM_SEND_EMAIL.EMAIL_BODY = ""
   Else
      FAMILY.FORM_SEND_EMAIL.RecipientName = ""
      FAMILY.FORM_SEND_EMAIL.RecipientEmailAddress = ""
      FAMILY.FORM_SEND_EMAIL.EMAIL_SUBJECT = ""
      FAMILY.FORM_SEND_EMAIL.EMAIL_BODY = ""
   End If
   '\\* Display The Email Form
   FAMILY.FORM_SEND_EMAIL.Show
End Sub
'===================================================================



Private Sub mnuSendEmail1_Click()
   Call mnuSendEmail_Click
End Sub

Private Sub mnuTray_Click()
   If mnuTray.Checked = True Then
      Call modPublic.SET_MINIMIZE_TO_SYSTEM_TRAY(False)
      mnuTray.Checked = False
   Else
      Call modPublic.SET_MINIMIZE_TO_SYSTEM_TRAY(True)
      mnuTray.Checked = True
   End If
End Sub

'===================================================================
Private Sub mnuUpdateLinks_Click()
   FAMILY.FORM_LINKS.Show vbModal
End Sub
'===================================================================


'===================================================================
Private Sub mnuUserProfile_Click()
   FAMILY.FORM_USERS.Show vbModal
End Sub
'===================================================================



'===================================================================
Private Sub MyFamily_Error(Description As String, ERROR_ID As Long)
   MsgBox Description, vbCritical
End Sub
'===================================================================


Private Sub mnuVistMyWeb_Click()
   Debug.Print "Visiting The Author's Web Page"

End Sub

Private Sub StatusBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) And (CURRENT_STATE = "") Then
      PopupMenu mnuFile
   End If
End Sub

'===================================================================
Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(2).ToolTipText = " " & Format(Now, "LONG DATE") & " "
   StatusBar1.Panels(3).ToolTipText = " " & Format(Now, "LONG TIME") & " "
End Sub
'===================================================================



'===================================================================
Private Sub TabStrip1_Click()

   If TabStrip1.Tabs(1).Selected = True Then
      FieldFrame.Visible = True
      NotesFrame.Visible = False
   End If

   If TabStrip1.Tabs(2).Selected = True Then
      FieldFrame.Visible = False
      NotesFrame.Visible = True
      txtNotes.Visible = True
   End If

End Sub
'===================================================================



'===================================================================
Private Sub TabStrip1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) And (CURRENT_STATE = "") Then
      PopupMenu mnuFile
   End If
End Sub
'===================================================================



'===================================================================
Private Sub TextBox_GotFocus(Index As Integer)
   If EDIT_MODE = True Then
      TextBox(Index).SelStart = 0
      TextBox(Index).SelLength = Len(TextBox(Index).Text)
   End If
End Sub
'===================================================================



Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
   '   'Validate Character Entered for First Name and Last Name
   '   If (Index = 0) Or (Index = 1) Then
   '      If modPublic.MyChar(Chr(KeyAscii), "-", True) = False Then
   '         If (KeyAscii <> vbKeyBack) Then
   '            KeyAscii = 0
   '         End If
   '      End If
   '   End If
   'Debug.Print Str(KeyAscii)
End Sub

'===================================================================
Private Sub TextBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) And (CURRENT_STATE = "") And (RECORD_SELECTED = True) Then
      PopupMenu mnuEmails, , , , mnuSendEmail
   ElseIf (Button = vbRightButton) And (CURRENT_STATE = "") And (RECORD_SELECTED = False) Then
      PopupMenu mnuFile
   End If
End Sub
'===================================================================





'===================================================================
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 2   'Reminders
         Call mnuReminder_Click
      Case 4   'Search/Print
         Call mnuSearchPrint_Click
      Case 6   'Send Email
         Call mnuSendEmail_Click
      Case 8   'User Profile
         Call mnuUserProfile_Click
      Case 10   'Update Personal Links
         Call mnuUpdateLinks_Click
   End Select
End Sub
'===================================================================



Private Sub Toolbar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) And (CURRENT_STATE = "") Then
      PopupMenu mnuFile
   End If
End Sub

'===================================================================
Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim btn As MSComctlLib.Button   'use ComctlLib.Button if you have VB5

   For Each btn In Toolbar1.Buttons
      If (X >= btn.Left And X <= btn.Left + btn.Width) And (Y >= btn.Top And Y <= btn.Top + btn.Height) Then
         'you could check the button style here because
         'you're probably not interested in seperator buttons
         Debug.Print "You are over button #" & btn.Index

         Select Case btn.Index
            Case 2
               'Reminders
               StatusBar1.Panels(1).Text = "Click For Reminders"
            Case 4
               'Search/Print
               StatusBar1.Panels(1).Text = "Click To Search/Print Records"
            Case 6
               'Email
               StatusBar1.Panels(1).Text = "Click To Send An Email"
            Case 8
               StatusBar1.Panels(1).Text = "Click For View/Edit User Profile"
            Case 10
               'Update INTERNET Links
               StatusBar1.Panels(1).Text = "Update Your Internet-Links"
         End Select
         'This Exits The For Loop
         Exit For
      End If
   Next
End Sub
'===================================================================



Private Sub TrayArea1_DblClick()
   TrayArea1.Visible = False
   FAMILY.FORM_MAIN.WindowState = 0
   FAMILY.FORM_MAIN.Show
   FAMILY.FORM_MAIN.Visible = True
End Sub

Private Sub TrayArea1_MouseDown(Button As Integer)
   PopupMenu mnuTrayPop, , , , mnuRestore
End Sub

Private Sub UsersTreeView_Collapse(ByVal Node As MSComctlLib.Node)
   If Node.Text = "" Then
      Last_Parent_Node_Clicked = Node.Text
      ComboBox(1).Text = Last_Parent_Node_Clicked
      Last_Node_Clicked = Node.Text
   End If
   'Clear The The Textboxes
   Call Setup_TextBox(Me, True, True)
   Call Change_Button(False, False, True, True, False, False)
   RECORD_SELECTED = False
End Sub

Private Sub UsersTreeView_Expand(ByVal Node As MSComctlLib.Node)
   If Node.Text = "" Then
      Last_Parent_Node_Clicked = Node.Text
      ComboBox(1).Text = Last_Parent_Node_Clicked
      Last_Node_Clicked = Node.Text
   End If

   'Clear The The Textboxes
   Call Setup_TextBox(Me, True, True)
   Call Change_Button(False, False, True, True, False, False)
   RECORD_SELECTED = False
End Sub

'===================================================================
Private Sub UsersTreeView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Store the Value of the Button in the Tag For Later Use
   UsersTreeView.Tag = Button
End Sub
'===================================================================




'===================================================================
Private Sub UsersTreeView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim tvNode As Node

   'This is used to track the Node That Mouse is over
   Set tvNode = UsersTreeView.HitTest(X, Y)

   If tvNode Is Nothing Then
      StatusBar1.Panels(1).Text = ""
      Exit Sub
   Else
      StatusBar1.Panels(1).Text = tvNode.Text
   End If
End Sub
'===================================================================



'===================================================================
Private Sub UsersTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
   Dim sPos          As Integer   'Seperator Position "_"
   Dim tmpFirstName  As String   'First Name
   Dim tmpLastName   As String   'Last Name
   Dim tmpRelation   As String   'Relation or Category
   Dim tmpSQL        As String   'SQL String

   'Use The Tag That was stored from UsersTreeView_MouseDown
   Select Case UsersTreeView.Tag

      Case vbRightButton   'Right Mouse Button Clicked

         Select Case Node.Key
               'Check is the node selected is the "ROOT"
            Case Is = "ROOT"
               Last_Parent_Node_Clicked = "ROOT"
               Last_Node_Clicked = Node.Text
               'Clear The The Textboxes
               Call Setup_TextBox(Me, True, True)
               Call Change_Button(False, False, True, True, False, False)
               'Disable the Menu Option "Edit This Record"
               mnuEditRecord.Enabled = False
               PopupMenu mnuRelation_Category, , , , mnuGroup

               ' It's either a category or a person
            Case Else

               'This is used to check if the node selected is a Category
               If Node.Parent.Key = "ROOT" Then
                  Last_Parent_Node_Clicked = "ROOT"
                  Last_Node_Clicked = Node.Text
                  'Clear The The Textboxes
                  Call Setup_TextBox(Me, True, True)
                  Call Change_Button(False, False, True, True, False, False)
                  ComboBox(1).Text = Last_Node_Clicked
                  'Disable the Menu Option "Edit This Record"
                  mnuEditRecord.Enabled = False
                  PopupMenu mnuRelation_Category, , , , mnuAddRecord
                  RECORD_SELECTED = False

               Else   ' It's a person from you contact list

                  Last_Parent_Node_Clicked = Node.Parent
                  Last_Node_Clicked = Node.Text

                  'Find The Record
                  'Used To Find The Position of
                  sPos = InStr(1, Node.Text, "_")
                  If sPos > 0 Then
                     tmpFirstName = Apostrophe$(Mid$(Node.Text, 1, sPos - 1))
                     tmpLastName = Apostrophe$(Mid$(Node.Text, sPos + 1))
                     tmpRelation = Node.Parent.Text

                     If FAMILY.FORM_MAIN.FIND_RECORD(tmpFirstName, tmpLastName, tmpRelation) = True Then
                        Debug.Print "Record Found"
                        If DISPLAY_RECORD(MyUser.USER_CONTACTS) = False Then
                           MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", the program was unable to properly display the record.", vbCritical + vbOKOnly
                        End If

                        Call Change_Button(True, True, True, True, False, False)
                        'Enable the Menu Option "Edit This Record"
                        mnuEditRecord.Enabled = True
                        'Display The Popupmenu
                        PopupMenu mnuRelation_Category, , , , mnuEditRecord
                        RECORD_SELECTED = True
                     Else
                        Debug.Print "Record Not Found"
                        RECORD_SELECTED = False
                        'Reload TreeVIEW
                        Call LOAD_frmMAIN_TREEVIEW(Me, Me.UsersTreeView, ImageList1, Last_Parent_Node_Clicked)
                        DoEvents

                        'Lock The Texboxes and Comboboxes
                        Call Setup_TextBox(Me, True, True)
                        Call Setup_ComboBox(Me, True, True)
                        Call Change_Button(False, False, True, True, False, False)
                        Exit Sub
                     End If

                  End If
               End If
         End Select


         'Left Mouse Button Clicked
      Case vbLeftButton

         Select Case Node.Key

               'check if it's the "ROOT"
            Case Is = "ROOT"
               Last_Parent_Node_Clicked = "ROOT"
               Last_Node_Clicked = Node.Text
               'Clear The The Textboxes
               Call Setup_TextBox(Me, True, True)
               Call Change_Button(False, False, True, True, False, False)
               RECORD_SELECTED = False

            Case Else   ' It's either a category or a person

               'This is used to check if the node selected is a Category
               If Node.Parent.Key = "ROOT" Then

                  Last_Parent_Node_Clicked = Node.Text
                  ComboBox(1).Text = Last_Parent_Node_Clicked
                  'Clear The The Textboxes
                  Call Setup_TextBox(Me, True, True)
                  Call Change_Button(False, False, True, True, False, False)
                  Last_Node_Clicked = Node.Text
                  RECORD_SELECTED = False

                  'Then It Is A Person For Your Contact List
               Else

                  Last_Parent_Node_Clicked = Node.Parent.Text
                  Last_Node_Clicked = Node.Text

                  'Find The Record
                  'This Locates The Position of The "_"
                  sPos = InStr(1, Node.Text, "_")
                  If sPos > 0 Then
                     'Find The First Name
                     tmpFirstName = Apostrophe$(Mid$(Node.Text, 1, sPos - 1))
                     'Find The Last Name
                     tmpLastName = Apostrophe$(Mid$(Node.Text, sPos + 1))
                     tmpRelation = Node.Parent.Text

                     If FAMILY.FORM_MAIN.FIND_RECORD(tmpFirstName, tmpLastName, tmpRelation) = True Then
                        RECORD_SELECTED = True
                        Debug.Print "Record Found"

                        If DISPLAY_RECORD(MyUser.USER_CONTACTS) = False Then
                           MsgBox "Sorry " & CURRENT_USER.FULL_NAME & ", the program was unable to properly display the record.", vbCritical + vbOKOnly
                        End If

                        Call Change_Button(True, True, True, True, False, False)
                     Else
                        Debug.Print "Record Not Found"
                        RECORD_SELECTED = False
                        'Reload TreeVIEW
                        Call LOAD_frmMAIN_TREEVIEW(Me, Me.UsersTreeView, ImageList1, Last_Parent_Node_Clicked)
                        DoEvents

                        'Lock The Texboxes and Comboboxes
                        Call Setup_TextBox(Me, True, True)
                        Call Setup_ComboBox(Me, True, True)
                        Call Change_Button(False, False, True, True, False, False)
                        Exit Sub
                     End If

                  End If

               End If

         End Select

   End Select

End Sub
'===================================================================



'===================================================================
Public Sub Making_Changes(ByVal STATE As Boolean)

   Select Case STATE
         'If edditing or adding
      Case True
         btnEdit.Enabled = False
         btnAdd.Enabled = False
         btnDelete.Enabled = False
         btnSave.Enabled = True
         btnCancel.Enabled = True
         ComboBox(0).Locked = False
         ComboBox(1).Locked = False
         'If not edditing or adding
      Case False
         btnEdit.Enabled = True
         btnAdd.Enabled = True
         btnDelete.Enabled = True
         btnSave.Enabled = False
         btnCancel.Enabled = False
         btnExit.Enabled = False
         ComboBox(0).Locked = True
         ComboBox(1).Locked = True
   End Select
End Sub
'====================================================================


'====================================================================
'Used To Display Record
'====================================================================
Public Function DISPLAY_RECORD(ByVal REC_SET As ADODB.Recordset) As Boolean
   'on error goto DISPLAY_RECORD_ERROR
   If (Not REC_SET.BOF) And (Not REC_SET.EOF) Then
      'First Name
      TextBox(0).Text = "" & REC_SET.Fields("FirstName").Value
      'Last Name
      TextBox(1).Text = "" & REC_SET.Fields("LastName").Value
      'Sex
      ComboBox(0).Text = REC_SET.Fields("Sex").Value
      'Telephone
      TextBox(2).Text = REC_SET.Fields("Telephone").Value
      'Address
      TextBox(3).Text = REC_SET.Fields("Address").Value
      'City-State
      TextBox(4).Text = REC_SET.Fields("City_State").Value
      'Zip Code
      TextBox(5).Text = REC_SET.Fields("ZipCode").Value
      'Email
      TextBox(6).Text = REC_SET.Fields("EmailAddress").Value
      'Relation/Category
      ComboBox(1).Text = REC_SET.Fields("CATEGORY_NAME").Value
      'Notes
      txtNotes.Text = "" & REC_SET.Fields("Notes").Value

      RECORD_SELECTED = True
      DISPLAY_RECORD = True
   Else
      DISPLAY_RECORD = False
      RECORD_SELECTED = False
   End If
   Exit Function

DISPLAY_RECORD_ERROR:
   If Err.Number <> 0 Then
      DISPLAY_RECORD = False
      RECORD_SELECTED = False
      MsgBox "ERROR - frmMain.DISPLAY_RECORD" & vbNewLine & _
            "Error : " & Err.Description & " - err#" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'====================================================================



'====================================================================
'====================================================================
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
'====================================================================


'====================================================================
'====================================================================
Public Sub CheckLinkMenus()
   'mnuLinks
   For TmpByte = 1 To 10
      If (mnuLink(TmpByte).Caption <> "link") And (mnuLink(TmpByte).Caption <> "http://") And (Len(mnuLink(TmpByte).Caption) > 7) Then
         mnuLink(TmpByte).Enabled = True
      Else
         mnuLink(TmpByte).Enabled = False
      End If
   Next TmpByte
End Sub
'====================================================================


'**********************************************************************
'**********************************************************************
Private Sub Form_Resize()
   'Minimized
   If FAMILY.FORM_MAIN.WindowState = vbMinimized Then
      If MINIMIZE_TO_SYSTEM_TRAY Then
         If FAMILY.FORM_MAIN.CURRENT_STATE = "" Then
            Set TrayArea1.Icon = FAMILY.FORM_MAIN.Icon
            TrayArea1.ToolTip = " Double-Click To Restore " & frmMain.Caption & " "
            TrayArea1.Visible = True
            FAMILY.FORM_MAIN.Hide
         Else
            MsgBox "Hi, " & CURRENT_USER.FULL_NAME & ", you have to finish what you are doing before you minimize to the system tray.", vbInformation + vbOKOnly
            FAMILY.FORM_MAIN.WindowState = 0
         End If
      End If
   End If
End Sub
'**********************************************************************



'**********************************************************************
'Used To Find A Specific Record
'**********************************************************************
Public Function FIND_RECORD(ByVal FNAME As String, ByVal LNAME As String, ByVal RELATIONSHIP As String) As Boolean
   On Error GoTo FIND_RECORD_ERROR

   FNAME = Apostrophe$(FNAME)
   LNAME = Apostrophe$(LNAME)

   If (Len(FNAME) < 1) Or (Len(LNAME) < 1) Then
      RECORD_SELECTED = False
      FIND_RECORD = False

      'Reload TreeVIEW
      Call LOAD_frmMAIN_TREEVIEW(Me, Me.UsersTreeView, ImageList1, Last_Parent_Node_Clicked)
      DoEvents

      'Lock The Texboxes and Comboboxes
      Call Setup_TextBox(Me, True, True)
      Call Setup_ComboBox(Me, True, True)
      Call Change_Button(False, False, True, True, False, False)

      Exit Function

   Else

      tmpSQL = ""
      tmpSQL = "SELECT FirstName,LastName,CATEGORY_NAME,USER_NAME,Sex,Telephone,Address,City_State,ZipCode,Notes,EmailAddress FROM " & CONTACTS_TABLENAME & _
            " WHERE FirstName = '" & FNAME & "'" & _
            " AND LastName = '" & LNAME & "'" & _
            " AND CATEGORY_NAME = '" & RELATIONSHIP & "'" & _
            " AND USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'"


      'Set Recordset to new
      Set FAMILY.FORM_MAIN.MyUser.USER_CONTACTS = New ADODB.Recordset
      FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

      If FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.RecordCount > 0 Then
         FIND_RECORD = True
         Debug.Print "Record Found"
      Else
         RECORD_SELECTED = False
         FIND_RECORD = False
         Debug.Print "record not found"
      End If

      tmpSQL = ""
      Exit Function

   End If

FIND_RECORD_ERROR:
   If Err.Number <> 0 Then
      RECORD_SELECTED = False
      FIND_RECORD = False
      MsgBox "ERROR : frmMain.FINDRecord" & vbNewLine & _
            "ERROR #" & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'**********************************************************************
