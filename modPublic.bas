Attribute VB_Name = "modPublic"

Option Explicit

'Menu Stuff
Public OldProc As Long
Public OldProc1 As Long

Public Const IDM_ABOUT As Long = 1010

Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const WM_SYSCOMMAND = &H112
Public Const GWL_WNDPROC = (-4)
Public Const WM_CLOSE = &H10

Public Declare Function CallWindowProc Lib "user32" _
      Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
      ByVal hWnd As Long, ByVal MSG As Long, _
      ByVal wParam As Long, lParam As Any) As Long

Public Declare Function AppendMenu Lib "user32" _
      Alias "AppendMenuA" (ByVal hMenu As Long, _
      ByVal wFlags As Long, ByVal wIDNewItem As Long, _
      ByVal lpNewItem As Any) As Long

Public Declare Function SetWindowLong Lib "user32" _
      Alias "SetWindowLongA" (ByVal hWnd As Long, _
      ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" _
      (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Declare Function DeleteMenu Lib "user32" _
      (ByVal hMenu As Long, ByVal nPosition As Long, _
      ByVal wFlags As Long) As Long

Public Declare Function GetProp Lib "user32" _
      Alias "GetPropA" (ByVal hWnd As Long, _
      ByVal lpString As String) As Long

Public Declare Function RemoveProp Lib "user32" _
      Alias "RemovePropA" (ByVal hWnd As Long, _
      ByVal lpString As String) As Long

'=================================================================


Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Public Const SW_NORMAL = 1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETREDRAW = &HB


 Public Enum ACCESS_LEVEL_TYPE
   Administrator = 0
   User = 1
End Enum


Public Type USER_INFO
   FULL_NAME      As String
   LOGIN_NAME     As String
   PASSWORD       As String
   EMAIL_ADDRESS  As String
   SMTP_SERVER    As String
   ACCESS_LEVEL   As ACCESS_LEVEL_TYPE
End Type
Public CURRENT_USER As USER_INFO

 Public Enum MySwitch
   [Switch Off] = 0
   [Switch On] = 1
End Enum


'Maximum size of the user name
Public Const MAX_USER_NAME_SIZE         As Byte = 25
'Minimum size of the user name
Public Const MIN_USER_NAME_SIZE         As Byte = 3
'Max First Name or Last Size
Public Const MAX_FNAME_SIZE             As Byte = 30
'Minimum First Name Size
Public Const MIN_FNAME_SIZE             As Byte = 3
'Maximum FullName Size
Public Const MAX_FULLNAME_SIZE          As Byte = 50
'Mimimum FullName Size
Public Const MIN_FULLNAME_SIZE          As Byte = 3


'Links
Public Const MAX_LINKS_SIZE As Byte = 50
Public Const MAX_LINKS_ALLOWED As Byte = 10

'Stores the MAXIMUM AMOUNT OF CATEGORIES ALLOWED
Public Const MAX_CATEGORY_ALLOWED       As Byte = 10
'Maximum size of a Category Name
Public Const MAX_CATEGORY_NAME_SIZE     As Byte = 25
'Minimum CATEGORY NAME SIZE
Public Const MINIMUM_CATEGORY_NAME_SIZE As Byte = 5

Public Const TXTBOX_BGROUND_COLOR As Long = &HC0FFFF
Public Const TXTBOX_FGROUND_COLOR = vbBlack
'-----------------------------------------------------------------


'-----------------------------------------------------------------
' SOME DATABASE STUFF
'-----------------------------------------------------------------
Public Const DATABASE_FILENAME         As String = "FamilyDB.FM3"
Public Const DATABASE_PASSWORD         As String = "SmileyOmar"
'This Will Be Used To Store The Database
Public DATABASE_PATH                   As String
' This Stores The Database Vaesion
Public Const DATABASE_VERSION          As Byte = 3

'-----------------------------------------------------------------

'::--------------------------------------------------------------::
'Stores The Names of the Database Tables
'::--------------------------------------------------------------::
'Name of the table used to store all the categories
Public Const CATEGORIES_TABLENAME   As String = "CATEGORIES"
'Name of the table used to store all the reminders
Public Const REMINDERS_TABLENAME    As String = "REMINDERS"
'Name of the table used to store all the categories
Public Const CONTACTS_TABLENAME     As String = "CONTACTS"
'Name of the table used to store all the info about all users
Public Const USER_TABLENAME         As String = "USERS"
'Name of the table used to store all the info about all users
Public Const LINKS_TABLENAME         As String = "LINKS"
'::--------------------------------------------------------------::


'::--------------------------------------------------------------::
'::                       Temporary Variables                    ::
'::--------------------------------------------------------------::
Public TmpByte                         As Byte
Public TmpInt                          As Integer
Public TmpString                       As String
Public tmpSQL                          As String
Public TmpMsgResult                    As VbMsgBoxResult
Public TmpRecordSet                    As ADODB.Recordset
'::--------------------------------------------------------------::


'::--------------------------------------------------------------::
'::                 SOME INFORMATION FOR THE AUTHOR              ::
'::--------------------------------------------------------------::
'Stores The Author's Home Page
Public Const AUTHOR_HOME_PAGE        As String = "http://www.omarswa.cjb.net"
'Stores the Author's Email Address
Public Const AUTHOR_EMAIL_ADDRESS    As String = "omarswan@yahoo.com"
'Stores The Author's First Name
Public Const AUTHOR_NAME             As String = "Omar"

Public Const AUTHOR_SMTP_SERVER      As String = "smtp.yourisp.com"
'::--------------------------------------------------------------::


''::--------------------------------------------------------------::
''Stores the SMTP HOST
'Public Smtp_Host                    As String
''Stores the SMTP HOST PORT : Default 25
'Public SMTP_PORT                    As String
''Stores the User Name Used to access the SMTP HOST
'Public Email_User_Name              As String
''Stores The Sender's Name
'Public Email_Sender_Name           As String
''Stores The Sender's Email Address
'Public Email_Sender_Email_Address As String
''::--------------------------------------------------------------::

'::==============================================================::
'::--------------------------------------------------------------::
'::          These Control Most Of The Program Operation         ::
'::--------------------------------------------------------------::
Public FAMILY                   As clsFamily   'FAMILY CLASS
Public PUBLIC_DATABASE          As clsDBase   'DATABASE CLASS
'::--------------------------------------------------------------::
'::--------------------------------------------------------------::
'::==============================================================::
'::==============================================================::


'::==============================================================::
'USED TO STORE ILLEGAL CHARACTERS
Public IllegalChars() As Byte
'::==============================================================::



'::--------------------------------------------------------------::
'::--------------------------------------------------------------::
'Task: Create a multi-level directory structure using
'CreateDirectory API call Declarations
'::--------------------------------------------------------------::
Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'::--------------------------------------------------------------::
'::--------------------------------------------------------------::

'Declaration To Plaw Music
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'::--------------------------------------------------------------::
'::--------------------------------------------------------------::


'---------------------------------------------------------------
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

' // Type declarations
Public Type OPENFILENAME
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type
'---------------------------------------------------------------


'---------------------------------------------------------------
Public Function Open_File(hWnd As Long) As String
   '
   Dim OpenFileDialog As OPENFILENAME
   Dim rv As Long
   On Error GoTo Open_File_ERR

   ' // init dialog
   With OpenFileDialog
      .lStructSize = Len(OpenFileDialog)
      .hwndOwner = hWnd&
      .hInstance = App.hInstance
      .lpstrFilter = "All Files (*.*)" & Chr$(0) & "*.*" + Chr$(0)
      .lpstrFile = Space$(254)
      .nMaxFile = 255
      .lpstrFileTitle = Space$(254)
      .nMaxFileTitle = 255
      .lpstrInitialDir = CurDir
      .lpstrTitle = "Open File..."
      .flags = 0
   End With

   ' // call API to show the dialog that was just initialized
   rv& = GetOpenFileName(OpenFileDialog)

   If (rv&) Then
      Open_File = Trim$(OpenFileDialog.lpstrFile)
   Else
      Open_File = ""
   End If

Open_File_ERR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : modPublic.Open_File" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      Open_File = ""
   End If
End Function
'---------------------------------------------------------------






'===================================================================
Public Sub RemoveMenus(ByVal ObjForm As Form)
   Dim hMenu As Long
   ' Get the form's system menu handle.
   hMenu = GetSystemMenu(ObjForm.hWnd, False)
   DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub
'===================================================================


'::--------------------------------------------------------------::
'::         Used to Check if the year is a leap year             ::
'::--------------------------------------------------------------::
Function IsLeapYear(ByVal sYear As String) As Boolean
   If IsDate("02/29/" & sYear) Then
      IsLeapYear = True
   Else
      IsLeapYear = False
   End If
End Function
'::--------------------------------------------------------------::




'::--------------------------------------------------------------::
'::              Used to Check if Caption is on                  ::
'::--------------------------------------------------------------::
Function Caption_On(Optional Switch As MySwitch) As Boolean
   If Switch = [Switch On] Then
      MsgBox "Empty"
   End If
End Function
'::--------------------------------------------------------------::



'================================================================
'--------------------ClearAll Procedure -------------------------
'================================================================
Public Sub ClearAll(Optional frm As Form, _
                     Optional txtBox As Boolean = True, _
                     Optional txtCob As Boolean = True, _
                     Optional txtOpt As Boolean = True)
   '-------------------------------------------------------------
   Dim MyControl As Control
   '-------------------------------------------------
   For Each MyControl In frm.Controls
      '---------------------
      If txtBox Then
         If TypeOf MyControl Is TextBox Then
            MyControl.Text = ""
         End If
      End If
      '-----------------------
      If txtCob Then
         If TypeOf MyControl Is ComboBox Then
            MyControl.Text = ""
         End If
      End If
      '-----------------------
      If txtOpt Then
         If TypeOf MyControl Is OptionButton Then
            MyControl.Value = False
         End If
      End If
      '----------------------
   Next
   '-----------------------------------------
End Sub
'================================================================
'================================================================






'::--------------------------------------------------------------::
'::--------------------------------------------------------------::
Public Sub WriteToErrorLog(sFormName As Form, sRoutineName As String, sError As String, iErrorNumber As Long)
   Dim FileFree As Byte

   On Local Error Resume Next

   FileFree = FreeFile
   Open App.PATH & "\ErrorLog.Txt" For Append As #FileFree
   Print #FileFree, vbNewLine
   Print #FileFree, "Date : " & Format$(Now, "Long Date")
   Print #FileFree, sFormName.Name & " " & sRoutineName & " " & sError
   Close #FileFree
End Sub
'::--------------------------------------------------------------::


Public Function AppPath() As String
   If Right(App.PATH, 1) <> "\" Then AppPath = App.PATH & "\" Else AppPath = App.PATH
End Function



'::--------------------------------------------------------------::
'::           Used to clear all TextBoxes on a form              ::
'::--------------------------------------------------------------::
Public Sub Setup_TextBox(ByVal ObjForm As Form, _
                         ByVal Clear_It As Boolean, _
                         Optional ByVal Lock_It As Boolean = True, _
                         Optional ByVal FGROUND_COLOR As OLE_COLOR = TXTBOX_FGROUND_COLOR, _
                         Optional ByVal BGROUND_COLOR As OLE_COLOR = TXTBOX_BGROUND_COLOR)

   Dim Ctrl As Control
   'Loop to clear all Text Box 's.
   For Each Ctrl In ObjForm
      If TypeOf Ctrl Is TextBox Then

         If Clear_It = True Then
            Ctrl.Text = ""   'Empty the contents of the text box
         End If

         Ctrl.Locked = Lock_It   'Lock the TextBox

         Ctrl.BackColor = BGROUND_COLOR   'Default BackColor
         Ctrl.ForeColor = FGROUND_COLOR   'Default ForeColor
      End If
   Next
End Sub
'::--------------------------------------------------------------::



'::--------------------------------------------------------------::
'::    Set all the comboboxes listindex = 0  'first option       ::
'::--------------------------------------------------------------::
Public Sub Setup_ComboBox(ByVal ObjForm As Form, _
                          Optional ByVal Lock_It As Boolean = True, _
                          Optional ByVal Set_ListIndex As Boolean = True, _
                          Optional ByVal FGROUND_COLOR As OLE_COLOR = TXTBOX_FGROUND_COLOR, _
                          Optional ByVal BGROUND_COLOR As OLE_COLOR = TXTBOX_BGROUND_COLOR)

   Dim cmbBox As Control
   For Each cmbBox In ObjForm
      If TypeOf cmbBox Is ComboBox Then
         If Set_ListIndex = True Then
            If cmbBox.ListCount > 0 Then
               cmbBox.ListIndex = 0
            End If
         End If
         cmbBox.BackColor = BGROUND_COLOR   'Default BackColor
         cmbBox.ForeColor = FGROUND_COLOR   'Default ForeColor
         cmbBox.Locked = Lock_It
      End If
   Next
End Sub
'::--------------------------------------------------------------::



'===================================================================
Public Sub Mousebusy()
   Screen.MousePointer = vbHourglass
End Sub

Public Sub MouseDefault()
   Screen.MousePointer = vbNormal
End Sub
'===================================================================


'===================================================================
'Used To Check for Valid Characters
'===================================================================
Public Function MyChar(ByVal Character As String, Optional ByVal Extra_Chars As String = "", Optional ByVal Beep_on_Invalid As Boolean = False) As Boolean
   Character = Left$(Trim$(Character), 1)

   If Len(Extra_Chars) < 1 Then
      If (Character Like "[a-z]") Or (Character Like "[A-Z]") Or (Character Like "[0-9]") Or (Character Like Chr(vbKeyBack)) Then
         MyChar = True
      Else
         MyChar = False
         If Beep_on_Invalid = True Then
            Beep
         End If
      End If
   Else
      If (Character Like "[a-z]") Or (Character Like "[A-Z]") Or (Character Like "[0-9]") Or (Character Like "[" & Extra_Chars & "]") Or (Character Like Chr(vbKeyBack)) Then
         MyChar = True
      Else
         MyChar = False
         If Beep_on_Invalid = True Then
            Beep
         End If
      End If
   End If
End Function
'===================================================================


'::--------------------------------------------------------------::
'::--------------------------------------------------------------::
':: :: :: ::    This Is Where Everything Starts         :: :: :: ::
'::--------------------------------------------------------------::
'::--------------------------------------------------------------::
Public Sub Main()
   '   'Used To Check If The Application is already running
   '   If App.PrevInstance Then   'This checks if webserver is allready started
   '      MsgBox "Sorry, but " & App.ProductName & " is already running.", vbMsgBoxSetForeground + vbInformation
   '      End
   '   End If

   'Setup The Illegal Characters
   ReDim IllegalChars(13)
   IllegalChars(1) = 34   ' " " "
   IllegalChars(2) = 38   ' " & "
   IllegalChars(3) = 39   ' " ' "
   IllegalChars(4) = 95   ' " _ "
   IllegalChars(5) = 96   ' " ` "
   IllegalChars(6) = 44   ' " , "
   IllegalChars(7) = 91   ' " [ "
   IllegalChars(8) = 93   ' " ] "
   IllegalChars(9) = 37   ' " % "
   IllegalChars(10) = 42   ' " * "
   IllegalChars(11) = 59   ' " ; "
   IllegalChars(12) = 40   ' " ( "
   IllegalChars(13) = 41   ' " ) "

   Debug.Print "Setting Up Classes"
   'Initialize The Family Class (The Main Controler)
   Set FAMILY = New clsFamily
   'Initialize The Database   Class   (Controls The Database Connection)
   Set PUBLIC_DATABASE = New clsDBase
   DoEvents
   Debug.Print "Classes Loaded"

   If PUBLIC_DATABASE.OPEN_DATABASE_CONNECTION = False Then
      MsgBox "The database was not found or could not be loaded." & _
            vbNewLine & "Please recreate the database, and make sure that it is in it's correct path." & _
            vbNewLine & vbNewLine & PUBLIC_DATABASE.PATH_AND_FILENAME, vbCritical + vbOKOnly
      Set PUBLIC_DATABASE = Nothing
      Set FAMILY = Nothing
      End
   Else
      Debug.Print "The Database was opened successfully"
   End If

   'Load and Show The Splash Screen
   Load FAMILY.FORM_SPLASH
   FAMILY.FORM_SPLASH.Show
End Sub
'::--------------------------------------------------------------::



'::--------------------------------------------------------------::
'::              Original Author : Gaetan Savoie                 ::
':: Used To Format a SQL string incase it has an Apostrophe [']  ::
'::--------------------------------------------------------------::
Public Function Apostrophe(ByVal sFieldString As String) As String
   If InStr(sFieldString, "'") Then
      Dim iLen        As Integer
      Dim i           As Integer
      Dim apostr      As Integer
      iLen = Len(sFieldString)
      i = 1

      Do While i <= iLen
         If Mid$(sFieldString, i, 1) = "'" Then
            apostr = i
            sFieldString = Left$(sFieldString, apostr) & "'" & _
                  Right$(sFieldString, iLen - apostr)
            iLen = Len(sFieldString)
            i = i + 1
         End If
         i = i + 1
      Loop
   End If
   Apostrophe = sFieldString
End Function
'::--------------------------------------------------------------::



'::--------------------------------------------------------------::
':: Used To Convert Some String Into Date  [FORMAT - MM/DD/YY]   ::
'::--------------------------------------------------------------::
Public Function sDate(ByVal TheMonth As Byte, _
                      ByVal TheDay As Byte, _
                      ByVal TheYear As Integer, _
                      ByVal tTime As String, _
                      ByVal IsIt_AM As Boolean) As Date

   Dim tmpStr As String
   Dim TheTime As String

   If IsIt_AM = True Then
      TheTime = tTime & " AM"
   Else
      TheTime = tTime & " PM"
   End If

   tmpStr = CStr(TheMonth) & "/" & CStr(TheDay) & "/" & CStr(TheYear) & " " & TheTime

   'Check If Year As Passed
   'If TheYear < Year(Now) Then
   '   TheYear = Year(Now)
   'End If

   If IsDate(tmpStr) = True Then
      sDate = CDate(tmpStr)
   Else
      sDate = Now
   End If
End Function
'::--------------------------------------------------------------::



'::--------------------------------------------------------------::
'::            Extracts The Time Only. Example  02:40            ::
'::--------------------------------------------------------------::
Public Function MyTime_TIME(ByVal TheTime As Date) As String
   MyTime_TIME = Left$(Format$(TheTime, "HH:MM AMPM"), 5)
End Function
'::--------------------------------------------------------------::


'::--------------------------------------------------------------::
'::               Extracts The AM or PM from DatTime             ::
'::--------------------------------------------------------------::
Public Function MyTime_AMPM(ByVal TheTime As Date) As String
   MyTime_AMPM = Right$(Format$(TheTime, "HH:MM AMPM"), 2)
End Function
'::--------------------------------------------------------------::



'::--------------------------------------------------------------::
'::    Used to load the time into a combobox [01:00 to 12:59]    ::
'::--------------------------------------------------------------::
Public Function LOAD_TIME_LIST(ByVal TheComboBox As Object, Optional ByVal Current_Time As Boolean = True) As Boolean
   Dim Cnt1 As Byte
   Dim Cnt2 As Byte

   'Temporarily stop the repainting of the combo
   SendMessage TheComboBox.hWnd, WM_SETREDRAW, False, vbEmpty

   TheComboBox.Clear

   Dim tm_val As Single, hr As Integer, mi As Integer
   TheComboBox.Clear
   For hr = 1 To 12
      For mi = 0 To 59
         tm_val = hr / 24 + mi / 24 / 60
         TheComboBox.AddItem Format(tm_val, "hh:mm")
      Next mi
   Next hr


   'Repaint and refresh the combo
   SendMessage TheComboBox.hWnd, WM_SETREDRAW, True, vbEmpty

   '   For Cnt1 = 1 To 12
   '      For Cnt2 = 0 To 59
   '         TheComboBox.AddItem Format$(Str$(Cnt1), "00") & ":" & Format$(Str$(Cnt2), "00")
   '      Next Cnt2
   '   Next Cnt1
   '
   '   If Current_Time = True Then
   '      TheComboBox.Text = modPublic.MyTime_TIME(Time)
   '   Else
   '      TheComboBox.ListIndex = 0
   '   End If

   LOAD_TIME_LIST = True
End Function
'::--------------------------------------------------------------::




'::--------------------------------------------------------------::
'::              Sub Used to create new directory                ::
'::--------------------------------------------------------------::
Public Sub CreateNewDirectory(ByVal NewDirectory As String)
   Dim sDirTest As String
   Dim SecAttrib As SECURITY_ATTRIBUTES
   Dim bSuccess As Boolean
   Dim sPath As String
   Dim iCounter As Integer
   Dim sTempDir As String

   sPath = NewDirectory

   If Right(sPath, Len(sPath)) <> "\" Then
      sPath = sPath & "\"
   End If

   iCounter = 1

   Do Until InStr(iCounter, sPath, "\") = 0
      iCounter = InStr(iCounter, sPath, "\")
      sTempDir = Left(sPath, iCounter)
      sDirTest = Dir(sTempDir)
      iCounter = iCounter + 1
      'create directory
      SecAttrib.lpSecurityDescriptor = &O0
      SecAttrib.bInheritHandle = False
      SecAttrib.nLength = Len(SecAttrib)
      bSuccess = CreateDirectory(sTempDir, SecAttrib)
   Loop
End Sub
'::--------------------------------------------------------------::



'::--------------------------------------------------------------::
'::               Checks if file or directory exist              ::
'::--------------------------------------------------------------::
Public Function FileExist(DirPath As String) As Boolean
   FileExist = Dir$(DirPath) <> ""
End Function
'::--------------------------------------------------------------::



'============================================================================================================
'============================================================================================================
Public Sub PlayWav(Wav As String)
   Dim mc
   On Error Resume Next
   mc = sndPlaySound(Wav, 1)
End Sub
'============================================================================================================




'::--------------------------------------------------------------::
'::               Used To Proper Case a String                   ::
'::--------------------------------------------------------------::
Public Function ProperCase(ByVal vData As String) As String
   Dim i As Integer   'Used as a counter

   vData = Trim$(vData)   'Trim The String

   If Len(vData) < 1 Then   'Check the length
      ProperCase = vData
      Exit Function
   End If

   vData = UCase$(Left$(vData, 1)) & LCase$(Right$(vData, Len(vData) - 1))

   For i = 2 To Len(vData)
      If (Mid$(vData, i, 1) = " ") Or (Mid$(vData, i, 1) = "-") And (i + 1 <= Len(vData)) Then
         Mid$(vData, i + 1, 1) = UCase$(Mid$(vData, i + 1, 1))
      End If
   Next i

   ProperCase = vData
End Function
'::--------------------------------------------------------------::



'::--------------------------------------------------------------::
'::                    Used to add a BackSlash [\]               ::
'::--------------------------------------------------------------::
Public Function AddBackSlash(ByVal sPath As String) As String
   'Returns sPath with a trailing backslash if sPath does not
   'already have a trailing backslash. Otherwise, returns sPath.

   sPath = Trim$(sPath)
   If Len(sPath) > 0 Then
      sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
   End If
   AddBackSlash = sPath

End Function
'::--------------------------------------------------------------::




'::--------------------------------------------------------------::
'::                   Used to check for [ http://  ]             ::
'::--------------------------------------------------------------::
Public Function ADD_HTTP(ByVal sLink As String) As String

   sLink = Trim$(sLink)
   If Len(sLink) > 0 Then
      sLink = IIf(LCase$(Left$(sLink, 7)) <> "http://", "http://", "") & sLink
   End If
   ADD_HTTP = sLink
End Function
'::--------------------------------------------------------------::




'::--------------------------------------------------------------::
'Check The Minimize Option
'::--------------------------------------------------------------::
Public Function MINIMIZE_TO_SYSTEM_TRAY() As Boolean
   Dim TmpTray As String
   TmpTray = ReadIniFile(App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "TRAY", "")
   If TmpTray = "True" Then
      MINIMIZE_TO_SYSTEM_TRAY = True
   Else
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "TRAY", "False"
      MINIMIZE_TO_SYSTEM_TRAY = False
   End If
End Function
'::--------------------------------------------------------------::


'::--------------------------------------------------------------::
':: Used To Set Minimize To Tray Option
'::--------------------------------------------------------------::
Public Sub SET_MINIMIZE_TO_SYSTEM_TRAY(MinimizeIT As Boolean)
   If MinimizeIT = True Then
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "TRAY", "True"
   Else
      WriteIniFile App.PATH & "\Family3.ini", CURRENT_USER.LOGIN_NAME, "TRAY", "False"
   End If
End Sub
'::--------------------------------------------------------------::


'============================================================================================================
'\\ Checks if the em@il is valid
'============================================================================================================
Public Function IsValidEmail(ByVal EmailAddress As String) As Boolean
   IsValidEmail = EmailAddress Like "*@[A-Z,a-z,0-9]*.*"
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
' Used To Execute A Link A Open The Default Web Browser
'============================================================================================================
Public Function OPEN_LINK(ByVal url As String) As Boolean
   '
End Function
'============================================================================================================
'============================================================================================================


'**********************************************************************
'Print Text in the Center of the page
'**********************************************************************
Public Sub PrintCenter(PrintString$)
   'print the string in the center of the page
   Printer.CurrentX = (Printer.ScaleWidth / 2) - ((Printer.FontSize * _
         (Printer.TextWidth(PrintString$) / 8.28)) / 2)
   'where the 8.28 is the PC
   'default font size   (where the width of the letters comnes from)
   Printer.Print PrintString$
End Sub
'**********************************************************************
'**********************************************************************


'**********************************************************************
'Used To Open A Web Site
'**********************************************************************
Public Sub OpenWebsite(strWebsite As String)
   On Error GoTo OpenWebsite_Error
   If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL) < 33 Then
      ' Insert Error handling code here
   End If

OpenWebsite_Error:
   If Err.Number <> 0 Then
      MsgBox "OpenWebsite Error : " & Err.Description & " : " & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Sub
'**********************************************************************





'**********************************************************************
'**********************************************************************
Public Sub RemoveOldProc(inHWND, Proc)
   Dim tmpProc As Long
   tmpProc = GetProp(inHWND, Proc)
   If tmpProc = 0 Then
      Exit Sub
   End If
   RemoveProp inHWND, Proc
   SetWindowLong inHWND, GWL_WNDPROC, tmpProc
End Sub
'**********************************************************************


Public Sub ShutDown()
   Set PUBLIC_DATABASE = Nothing
   Set FAMILY = Nothing
   End
End Sub
