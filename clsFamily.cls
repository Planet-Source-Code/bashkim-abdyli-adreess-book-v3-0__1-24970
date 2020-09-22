VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'=================================================================
'Forms
'=================================================================
Public FORM_MAIN                    As frmMain
Public FORM_ABOUT                   As frmAbout
Public FORM_REMINDERS               As frmReminder
Public FORM_REMINDER_NOTES          As frmReminderNotes
Public FORM_LOGIN                   As frmLogin
Public FORM_SEND_EMAIL              As frmSendMail
Public FORM_SEARCH                  As frmSearch
Public FORM_CATEGORIES              As frmCategories
Public FORM_USERS                   As frmUsers
Public FORM_LINKS                   As frmLinks
Public WithEvents FORM_SPLASH       As frmSplash
Attribute FORM_SPLASH.VB_VarHelpID = -1
'=================================================================


'Stores the Connection String
'Private Conn_String                 As String
'Stores The Database Path
'Private Dbase_Path                  As String
'Stores the Name of the Database User
Private Dbase_User                  As String
'Stores The Datbase Name
'Private Dbase_Name                  As String
'Stores The Datbase Password
'Private Dbase_Password              As String


'======================================================================
'Public Events
'======================================================================
Public Event UserAdded(ByVal USER_NAME_ As String)
Public Event UserDeleted(ByVal USER_NAME_ As String)
Public Event CategoryAdded(ByVal CATEGORY_NAME As String)
Public Event CategoryRemoved(ByVal CATEGORY_NAME As String)
Public Event ERROR(ByVal ERROR_MESSAGE As String, ByVal ERR_NUM As Long)
'======================================================================
'======================================================================


'======================================================================
Private Sub Class_Initialize()
   On Error GoTo INIT_ERROR

   Set FORM_LOGIN = New frmLogin
   Set FORM_USERS = New frmUsers
   Set FORM_SEARCH = New frmSearch
   Set FORM_ABOUT = New frmAbout
   Set FORM_REMINDER_NOTES = New frmReminderNotes
   Set FORM_SPLASH = New frmSplash

   'Set FORM_SEND_EMAIL = New frmSendMail
   'Set FORM_SEND_EMAIL.SendEmail = New vbSendMail.clsSendMail

   Set FORM_REMINDERS = New frmReminder
   Set FORM_REMINDERS.MyUser = New clsReminder

   Set FORM_CATEGORIES = New frmCategories
   Set FORM_CATEGORIES.MyUser = New clsCategories

   Set FORM_USERS = New frmUsers
   Set FORM_USERS.MyUser = New clsUser

   Set FORM_LINKS = New frmLinks
   Set FORM_LINKS.MyUser = New clsLinks

   Set FORM_MAIN = New frmMain
   Set FORM_MAIN.MyUser = New clsContacts

INIT_ERROR:
   If Err.Number <> 0 Then
      MsgBox "The program is unable to Initialize." & vbNewLine & "clsFamily - Class Initialise ERROR" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "ERROR #     : " & Str$(Err.Number), vbCritical + vbOKOnly
      RaiseEvent ERROR(Err.Description, Err.Number)
      Err.Clear
      'Shutdown
      Call modPublic.ShutDown
      End
   End If
End Sub
'================================================================



'================================================================
Public Sub SETUP_FORM_MAIN()
   Call LOAD_frmMAIN_TREEVIEW(FAMILY.FORM_MAIN, FAMILY.FORM_MAIN.UsersTreeView, FAMILY.FORM_MAIN.ImageList1)
   DoEvents
   Call Setup_frmMain(FORM_MAIN)
   DoEvents
   FAMILY.FORM_MAIN.Caption = " ::  Family Address Book v3.0 - [ Welcome " & CURRENT_USER.FULL_NAME & " ]  ::"
   'FAMILY.FORM_MAIN.Cls

   'Set Up The Link Menus
   FAMILY.FORM_LINKS.MyUser.LOAD_LINKS
   FAMILY.FORM_MAIN.CheckLinkMenus
   DoEvents

   FAMILY.FORM_MAIN.RECORD_SELECTED = False
End Sub
'================================================================



'================================================================
'This Sub is Used To Setup frmMain
'*** Do not Make Any Changes if You Don't Understand It ***
'================================================================
Private Sub Setup_frmMain(ByVal ObjForm As Form)
   Debug.Print "clsFamily.Setup_frmMain : Setting Up frmMain"

   ObjForm.Width = 8790
   ObjForm.Height = 5890

   ObjForm.ComboBox(0).Clear
   ObjForm.ComboBox(0).AddItem "Female"
   ObjForm.ComboBox(0).AddItem "Male"
   ObjForm.ComboBox(0).ListIndex = 0
   ObjForm.ComboBox(0).Locked = True
   Debug.Print "Setup Main : Items Loaded To ComboBox(0)"

   'Set The Visibility of The Frame Surrounding txtNoes to False
   ObjForm.NotesFrame.Visible = False
   'txtNotes is The Textbox used to Store and Display the Extra
   'Notes For the Personal Contacts
   'Set txtNotes
   ObjForm.txtNotes.Visible = False
   ObjForm.FieldFrame.Visible = True

   ObjForm.TabStrip1.Top = 70
   ObjForm.TabStrip1.Left = 3150
   ObjForm.TabStrip1.Width = 5500
   ObjForm.TabStrip1.Height = 4100
   ObjForm.FieldFrame.Top = ObjForm.TabStrip1.Top + 350
   ObjForm.FieldFrame.Left = ObjForm.TabStrip1.Left + 100
   ObjForm.FieldFrame.Height = 3550
   ObjForm.FieldFrame.Width = 5300

   ObjForm.NotesFrame.Top = ObjForm.TabStrip1.Top + 350
   ObjForm.NotesFrame.Left = ObjForm.TabStrip1.Left + 100
   ObjForm.NotesFrame.Width = 5300
   ObjForm.NotesFrame.Height = 3550
   ObjForm.txtNotes.Left = 100
   ObjForm.txtNotes.Top = 200
   ObjForm.txtNotes.Width = ObjForm.NotesFrame.Width - 200
   ObjForm.txtNotes.Height = ObjForm.NotesFrame.Height - 250

   'Set Up Combo and Textboxes
   Call Setup_ComboBoxes(ObjForm, True, True)
   Call Setup_TextBoxes(ObjForm, True, True)
   '<<------------------------------------------------------------>>
   DoEvents
End Sub
'================================================================
'================================================================



'===================================================================
'Used to clear all TextBoxes on a form
'===================================================================
Public Sub Setup_TextBoxes(ByVal ObjForm As Form, _
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
'===================================================================
'===================================================================


'===================================================================
'Set all the comboboxes listindex = 0  'first option
'===================================================================
Public Sub Setup_ComboBoxes(ByVal ObjForm As Form, _
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
      End If
   Next
End Sub
'===================================================================
'===================================================================


'================================================================
Private Sub Class_Terminate()
   Call CLOSE_UP
End Sub
'================================================================





'================================================================
' Adds The New User
'================================================================
Public Function ADD_NEW_USER(ByVal pUSER_USERNAME As String, _
       ByVal pUSER_PASSWORD As String, _
       ByVal pFULL_NAME As String, _
       ByVal pUSER_ACCESSLEVEL As String, _
       Optional pUSER_EMAILADDRESS As String = "username@yourmail.com", _
       Optional pUSER_URL As String = AUTHOR_HOME_PAGE, _
       Optional pUSER_SMTP_SERVER As String = AUTHOR_SMTP_SERVER) As Boolean

   On Error GoTo ADD_NEW_USER_Error

   '::--------------------------------------------------------::
   '::- Method #1 For Adding A New User
   '::- To Add A new User Can Use This Method or the one below::
   '::--------------------------------------------------------::
   TmpString = ""
   TmpString = "INSERT INTO " & USER_TABLENAME & " (" & _
         "USER_LOGIN_NAME, " & _
         "USER_PASSWORD, " & _
         "FULL_NAME, " & _
         "USER_ACCESS_LEVEL, " & _
         "USER_EMAIL_ADDDRESS, " & _
         "USER_SMTP_SERVER, " & _
         "USER_HOMEPAGE_URL, " & _
         "DATE_ADDED, " & _
         "USER_LOCKED)" & _
         " VALUES (" & _
         "'" & Apostrophe(pUSER_USERNAME) & "', " & _
         "'" & Apostrophe(pUSER_PASSWORD) & "', " & _
         "'" & Apostrophe(pFULL_NAME) & "', " & _
         "'" & pUSER_ACCESSLEVEL & "', " & _
         "'" & pUSER_EMAILADDRESS & "', " & _
         "'" & pUSER_SMTP_SERVER & "', " & _
         "'" & pUSER_URL & "', " & _
         "#" & Now() & "#, " & _
         False & ")"

   PUBLIC_DATABASE.CONNECTION.Execute TmpString
   DoEvents
   '::--------------------------------------------------------::



   '::--------------------------------------------------------::
   '::- Method #2 For Adding A New User
   '::- To Add A new User Can Use This Method or the one Above::
   '::--------------------------------------------------------::
   '   Dim uRecordset    As ADODB.Recordset
   '   'On Error GoTo Add_To_USERS_TABLE_Error
   '
   '   ADD_NEW_USER = False
   '
   '   'Set uConnection = New ADODB.Connection
   '   Set uRecordset = New ADODB.Recordset
   '
   '   'Open The Recordset
   '   uRecordset.Open USER_TABLENAME, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   '
   '   Debug.Print "ADD_NEW_USER : Attempting To ADD A New Record"
   '   uRecordset.AddNew
   '   uRecordset.Fields("USER_LOGIN_NAME") = pUSER_USERNAME
   '   uRecordset.Fields("USER_PASSWORD") = pUSER_PASSWORD
   '   uRecordset.Fields("FULL_NAME") = ProperCase(Trim$(pFULL_NAME))
   '   uRecordset.Fields("USER_ACCESS_LEVEL") = pUSER_ACCESSLEVEL
   '   uRecordset.Fields("USER_EMAIL_ADDDRESS") = pUSER_EMAILADDRESS
   '   uRecordset.Fields("USER_SMTP_SERVER") = pUSER_SMTP_SERVER
   '   uRecordset.Fields("USER_HOMEPAGE_URL") = pUSER_URL
   '   uRecordset.Fields("DATE_ADDED") = Now
   '   uRecordset.Fields("USER_LOCKED") = False
   '   'Update the database
   '   uRecordset.Update
   '   Debug.Print "ADD_NEW_USER : USER Added"
   '   uRecordset.Requery
   '
   '   Debug.Print "ADD_TO_USERS_TABLE : Attempting To Close UP"
   '   'Close UP
   '   uRecordset.Close
   '   Set uRecordset = Nothing
   '   Debug.Print "ADD_TO_USERS_TABLE : Finished Closing Up"
   '


   Debug.Print "Adding default categories for " & pUSER_USERNAME
   'Add Default Categories
   ADD_NEW_USER_CATEGORY pUSER_USERNAME, "Acquaintance"
   ADD_NEW_USER_CATEGORY pUSER_USERNAME, "Co-Worker"
   ADD_NEW_USER_CATEGORY pUSER_USERNAME, "Family"
   ADD_NEW_USER_CATEGORY pUSER_USERNAME, "Friend"
   ADD_NEW_USER_CATEGORY pUSER_USERNAME, "Spouse"
   Debug.Print "Default Categories Added Succesfully for " & pUSER_USERNAME


   ADD_NEW_USER = True
   Exit Function

ADD_NEW_USER_Error:
   If Err.Number <> 0 Then
      ADD_NEW_USER = False
      MsgBox "clsFamily - ADD_NEW_USER_Error" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "ERROR #     : " & Str$(Err.Number), vbCritical + vbOKOnly
      RaiseEvent ERROR("Error (ADD_NEW_USER) : " & Err.Description & " : " & Str(Err.Number), 0)
      Err.Clear
   End If
End Function
'================================================================
'================================================================



'================================================================
'Used to add a new category
'================================================================
Public Function ADD_NEW_USER_CATEGORY(ByVal User_Name As String, ByVal CATEGORY_NAME As String) As Boolean
   Dim cRecordset       As ADODB.Recordset
   On Error GoTo ADD_NEW_USER_CATEGORY_ERROR

   ADD_NEW_USER_CATEGORY = False

   Set cRecordset = New ADODB.Recordset

   'cConnection.Open CONNECTION_STRING
   cRecordset.Open CATEGORIES_TABLENAME, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   Debug.Print "ADD_NEW_USER_CATEGORY : Connection and Recordset Opened, Trying to ADD New Record"

   cRecordset.AddNew
   cRecordset.Fields("USER_NAME") = User_Name
   cRecordset.Fields("CATEGORY_NAME") = CATEGORY_NAME
   cRecordset.Update
   Debug.Print "ADD_NEW_USER_CATEGORY : New User CATEGORY ADDED, Trying To Close UP"

   'Closing UP
   cRecordset.Close
   Set cRecordset = Nothing
   Debug.Print "ADD_NEW_USER_CATEGORY : Closing Completed"

   ADD_NEW_USER_CATEGORY = True
   Exit Function

ADD_NEW_USER_CATEGORY_ERROR:
   If Err.Number <> 0 Then
      ADD_NEW_USER_CATEGORY = False
      Debug.Print "ADD_NEW_USER_CATEGORY_ERROR : " & Err.Description
      MsgBox "clsFamily - Class ADD_NEW_USER_CATEGORY ERROR" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "ERROR #     : " & Str$(Err.Number), vbCritical + vbOKOnly

      RaiseEvent ERROR("ADD_NEW_USER_CATEGORY_ERROR : " & Err.Description, 0)
      Err.Clear
      Set cRecordset = Nothing
      Exit Function
   End If
End Function
'================================================================
'================================================================



'================================================================
' INITIALIZE
'================================================================
Public Function INITIALISE(ByVal USERNAME As String) As Boolean
   'Dim tmpStr As String
   'On Error GoTo INITIALISE_ERROR

   Dbase_User = USERNAME
INITIALISE_ERROR:
   If Err.Number <> 0 Then
      INITIALISE = False
      Debug.Print "INITAILISE ERROR : " & Err.Description & " " & Str$(Err.Number)
      RaiseEvent ERROR("INITAILISE ERROR : " & Err.Description, Err.Number)
      Err.Clear
   End If
End Function
'======================================================================
'======================================================================





'======================================================================
'CLOSE UP
'======================================================================
Public Sub CLOSE_UP()
   'Close Up The Classes
'   FORM_SEND_EMAIL.SendEmail.ShutDown
   Set FORM_MAIN.MyUser = Nothing
   Set FORM_REMINDERS.MyUser = Nothing
   Set FORM_CATEGORIES.MyUser = Nothing
'   Set FORM_SEND_EMAIL.SendEmail = Nothing

   'Close Up The Forms
   Set FORM_MAIN = Nothing
   Set FORM_ABOUT = Nothing
   Set FORM_REMINDERS = Nothing
   Set FORM_REMINDER_NOTES = Nothing
   Set FORM_LOGIN = Nothing

   Set FORM_CATEGORIES = Nothing
   Set FORM_SEND_EMAIL = Nothing
   Set FORM_USERS = Nothing
   Set FORM_SEARCH = Nothing
   Set FORM_LINKS = Nothing
   Set FORM_SPLASH = Nothing
End Sub
'=====================================================================



'=====================================================================
'This Function Is Used To Rename A User's Date
'=====================================================================
Public Function RENAME_USER_DATA(ByVal Old_User_Name As String, ByVal New_User_Name As String) As Boolean
   'To Do
   'Rename The User Contacts
   'Rename The User Categories
   'Rename The User Reminders

   On Error GoTo RENAME_USER_DATA_ERROR
   
   RENAME_USER_DATA = False

   '//::-------------------------------------------------------
   Debug.Print "Renaming User Contacts"
   tmpSQL = ""
   tmpSQL = "UPDATE " & CONTACTS_TABLENAME & " SET USER_NAME = '" & New_User_Name & "'" & _
         " WHERE USER_NAME = '" & Old_User_Name & "'"

   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   '//::-------------------------------------------------------


   '//::-------------------------------------------------------
   Debug.Print "Renaming The User Categories"
   tmpSQL = ""
   tmpSQL = "UPDATE " & CATEGORIES_TABLENAME & " SET USER_NAME = '" & New_User_Name & "'" & _
         " WHERE USER_NAME = '" & Old_User_Name & "'"

   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   '//::-------------------------------------------------------

   '//::-------------------------------------------------------
   Debug.Print "Renaming The User Reminders"
   tmpSQL = ""
   tmpSQL = "UPDATE " & REMINDERS_TABLENAME & " SET USER_NAME = '" & New_User_Name & "'" & _
         " WHERE USER_NAME = '" & Old_User_Name & "'"

   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   '//::-------------------------------------------------------

   RENAME_USER_DATA = True
   
RENAME_USER_DATA_ERROR:
   If Err.Number <> 0 Then
      RENAME_USER_DATA = False
      MsgBox "Error : clsFamily.RENAME_USER_DATA " & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "Err #" & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If

End Function
'=====================================================================




'=====================================================================
'This Function Is Used To Delete A User
'=====================================================================
Public Function DELETE_USER(ByVal User_Name As String) As Boolean
   On Error GoTo DELETE_USER_Error
   'TODO LIST
   '1 Remove From USER's Table
   '2 Remove CONTACTS
   '3 REMOVE CATEGORIES
   '4 Remove Reminders
   '5 Remove LINKS

   '::-------------------------------------------------------------::
   '# 1 Remove The User's Record From the USER'S TABLE
   '::-------------------------------------------------------------::
   tmpSQL = ""
   tmpSQL = "DELETE * FROM " & USER_TABLENAME
   tmpSQL = tmpSQL & " WHERE USER_LOGIN_NAME = '" & User_Name & "'"
   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   '::-------------------------------------------------------------::

   '::-------------------------------------------------------------::
   '# 2 Remove The User's Record From the CONTACTS TABLE
   '::-------------------------------------------------------------::
   tmpSQL = ""
   tmpSQL = "DELETE * FROM " & CONTACTS_TABLENAME
   tmpSQL = tmpSQL & " WHERE USER_NAME = '" & User_Name & "'"
   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   '::-------------------------------------------------------------::

   '::-------------------------------------------------------------::
   '# 3 Remove The User's Record From the CATEGORIES TABLE
   '::-------------------------------------------------------------::
   tmpSQL = ""
   tmpSQL = "DELETE * FROM " & CATEGORIES_TABLENAME
   tmpSQL = tmpSQL & " WHERE USER_NAME = '" & User_Name & "'"
   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   '::-------------------------------------------------------------::

   '::-------------------------------------------------------------::
   '# 4 Remove The User's Record From the REMINDERS TABLE
   '::-------------------------------------------------------------::
   tmpSQL = ""
   tmpSQL = "DELETE * FROM " & REMINDERS_TABLENAME
   tmpSQL = tmpSQL & " WHERE USER_NAME = '" & User_Name & "'"
   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   '::-------------------------------------------------------------::

   '::-------------------------------------------------------------::
   '# 5 Remove The User's Record From the LINKS TABLE
   '::-------------------------------------------------------------::
   tmpSQL = ""
   tmpSQL = "DELETE * FROM " & LINKS_TABLENAME
   tmpSQL = tmpSQL & " WHERE USER_NAME = '" & User_Name & "'"
   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   '::-------------------------------------------------------------::

DELETE_USER_Error:
   If Err.Number <> 0 Then
      RaiseEvent ERROR("clsFamily.DELETE_USER_Error:" & Err.Description, Err.Number)
      Err.Clear
   End If

End Function
'=====================================================================


'=====================================================================
'This Function Is Used To Counthe Number of Administrator
'=====================================================================
Public Function ADMINISTRATOR_COUNT() As Long
   Dim tmpRecset As ADODB.Recordset

   Debug.Print "clsFamily.ADMINISTRATOR_COUNT"

   tmpSQL = ""
   tmpSQL = "SELECT USER_ACCESS_LEVEL FROM " & USER_TABLENAME & _
         " WHERE USER_ACCESS_LEVEL = '" & "Administrator" & "'"

   Set tmpRecset = New ADODB.Recordset
   tmpRecset.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   tmpRecset.Requery

   Debug.Print "ADMINISTRATOR_COUNT = " & Str$(tmpRecset.RecordCount)
   ADMINISTRATOR_COUNT = tmpRecset.RecordCount

   'Clean Up
   tmpRecset.Close
   Set tmpRecset = Nothing
End Function
'=====================================================================



'=====================================================================
Private Sub FORM_SPLASH_UnLoaded()
   'Load And Show The Login Screen
   Debug.Print "Loading The Login Screen"
   Load FORM_LOGIN
   FORM_LOGIN.Show
End Sub
'=====================================================================



'::--------------------------------------------------------------::
'::--------------------------------------------------------------::
'::            Used To Reload The Current User Info              ::
'::--------------------------------------------------------------::
'::--------------------------------------------------------------::
Public Function RELOAD_CURRENT_USER(ByVal USERNAME As String) As Boolean
   Dim TmpRecordSet As ADODB.Recordset

   If Len(USERNAME) < 1 Then
      RELOAD_CURRENT_USER = False
      Exit Function
   End If

   tmpSQL = ""
   tmpSQL = "SELECT USER_LOGIN_NAME,FULL_NAME,USER_EMAIL_ADDDRESS,USER_PASSWORD,USER_SMTP_SERVER,USER_ACCESS_LEVEL FROM " & USER_TABLENAME
   tmpSQL = tmpSQL & " WHERE USER_LOGIN_NAME = '" & USERNAME & "'"

   Set TmpRecordSet = New ADODB.Recordset
   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   TmpRecordSet.Requery

   If TmpRecordSet.RecordCount > 0 Then
      CURRENT_USER.LOGIN_NAME = TmpRecordSet.Fields("USER_LOGIN_NAME")
      CURRENT_USER.FULL_NAME = ProperCase(TmpRecordSet.Fields("FULL_NAME"))
      CURRENT_USER.EMAIL_ADDRESS = TmpRecordSet.Fields("USER_EMAIL_ADDDRESS")
      CURRENT_USER.PASSWORD = TmpRecordSet.Fields("USER_PASSWORD")
      CURRENT_USER.SMTP_SERVER = TmpRecordSet.Fields("USER_SMTP_SERVER")

      If TmpRecordSet.Fields("USER_ACCESS_LEVEL") = "Administrator" Then
         CURRENT_USER.ACCESS_LEVEL = Administrator
      Else
         CURRENT_USER.ACCESS_LEVEL = User
      End If
      TmpRecordSet.Close
      Set TmpRecordSet = Nothing

      RELOAD_CURRENT_USER = True
      Exit Function
   Else
      RELOAD_CURRENT_USER = False
      Exit Function
   End If

End Function
'::--------------------------------------------------------------::


