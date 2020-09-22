VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::  About - Family Address Book v3.0"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrProgTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   3600
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10
      ScaleHeight     =   255
      ScaleWidth      =   6375
      TabIndex        =   4
      Top             =   0
      Width           =   6370
      Begin Family_v3.Label3D Label3D2 
         Height          =   255
         Left            =   45
         TabIndex        =   5
         Top             =   15
         Width           =   4770
         _ExtentX        =   8414
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
         Caption         =   " ::  About - Family Address Book v3.0  ..."
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   20
      TabIndex        =   1
      Top             =   180
      Width           =   6375
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3615
         ScaleWidth      =   6135
         TabIndex        =   2
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4250
      Width           =   855
   End
   Begin Family_v3.Label3D Label3D1 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   4320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   -2147483634
      ForeColor2      =   16711680
      Caption         =   ":: About - Family Address Book v3.0 ::"
      Phase           =   1
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' TYPE STRUCTURES
Private Type tpeTextProperties
   cbSize As Long
   iTabLength As Long
   iLeftMargin As Long
   iRightMargin As Long
   uiLengthDrawn As Long
End Type
Private Type tpeRectangle
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

' CONSTANTS
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4

' API DECLARATIONS
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As tpeRectangle, ByVal un As Long, lpDrawTextParams As tpeTextProperties) As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
Private Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As tpeRectangle) As Long

Public strCharSpace As Integer




Private Sub btnOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Call modPublic.RemoveMenus(Me)
   Me.Caption = "About - Family Address Book v3.0 - " & CURRENT_USER.LOGIN_NAME
   Call Picture1_Click
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'SetUp FrmMain Texboxes and Comboboxes
   Call Setup_TextBox(FAMILY.FORM_MAIN, True, True)
   Call Setup_ComboBox(FAMILY.FORM_MAIN, True, True)
   Call FAMILY.FORM_MAIN.Change_Button(False, False, True, True, False, False)

   'Unload The Form
   Unload Me
End Sub

Private Sub Picture1_Click()
   ' Draw the text with a large space between the characters
   strCharSpace = 40
   Call doAnimationFX
   ' Start the timer
   tmrProgTimer.Enabled = True
End Sub

Private Sub tmrProgTimer_Timer()
   ' Take away one of the present value of the spacing
   strCharSpace = strCharSpace - 1
   Call doAnimationFX   ' Draw the new string
   ' Check the value of 'strCharSpace'
   If strCharSpace = 0 Then tmrProgTimer.Enabled = False
End Sub

Private Sub doAnimationFX()
   ' Procedure Scope Declarations
   Dim typeDrawRect As tpeRectangle
   Dim typeDrawParams As tpeTextProperties
   Dim strCaption As String
   ' Set the string which will be animated
   strCaption = "Family Address Book v3.0" & vbNewLine & _
         vbNewLine & _
         "Special Thank To Each and Every VB Programmer" & vbNewLine & _
         "Who Post Their Source Codes So That Other Young" & vbNewLine & _
         "Aspiring Programers Such As Myself Can Learn." & vbNewLine & _
         "Thank You All!!" & vbNewLine & _
         vbNewLine & vbNewLine & _
         ":: Long live www.Planet-Source-Code.com ::"


   ' Set the area in which the animation will take place.
   ' Needs to be a control which has the '.hwnd' property
   ' and can be refreshed and cleared easily. So a picture
   ' box is the best candidate
   GetClientRect Picture1.hWnd, typeDrawRect
   ' Now set the properties which will be used in the animation
   With typeDrawParams
      ' The size of the animation
      .cbSize = Len(typeDrawParams)
      ' The left and right margins
      .iLeftMargin = 0
      .iRightMargin = 0
   End With
   ' Clear the picture box
   Picture1.Cls
   ' Set the character spacing which will be used
   SetTextCharacterExtra Picture1.hdc, Val(strCharSpace)
   ' Draw the string of text, in the set area with the
   ' specified options
   DrawTextEx Picture1.hdc, strCaption, Len(strCaption), _
         typeDrawRect, SaveOptions, typeDrawParams
   ' Refresh the picture box which contains the animation
   Picture1.Refresh
End Sub

Private Function SaveOptions() As Long
   ' Procedure Scope Declaration
   Dim MyFlags As Long
   ' Set the options which will be used in the FX
   MyFlags = MyFlags Or DT_CENTER
   MyFlags = MyFlags Or DT_VCENTER
   ' Store the flags which we have set above
   SaveOptions = MyFlags
End Function


