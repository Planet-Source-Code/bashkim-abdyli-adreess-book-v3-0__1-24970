VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreateDbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================
'================================================================
'***                  TO GOD BE THE GLORY                     ***
'================================================================
'================================================================
'*** For any Questions or Comments concerning this program    ***
'*** Homepage : http://www.omarswan.cjb.net                   ***
'*** Email    : omarswan@yahoo.com or omarswan@freemail.org.mk***
'*** AOL      : smileyomar  or omarsmiley                     ***
'================================================================
'================================================================
'* Deducated to SmileyOrange -> http://www.smileyorange.cjb.net *
'================================================================
'================================================================

Option Explicit

Public PUBLIC_CONNECTION As ADODB.Connection
Public CONNECTION_STRING As String
Public Path              As String
Public FileName          As String



'================================================================
'Used to Create or Recreate the database(DAO)
'================================================================
Public Function RECREATE_DATABASE() As Boolean
   Dim NEW_DBASE                      As DAO.Database
   Dim NEW_USERS_TABLE                As DAO.TableDef
   Dim NEW_CATEGORIES_DBASE_TABLE     As DAO.TableDef
   Dim NEW_REMINDERS_DBASE_TABLE      As DAO.TableDef
   Dim NEW_USERS_PERSONAL_DBASE_TABLE As DAO.TableDef
   Dim TmpFields(1 To 12)             As DAO.Field   ' An Array Of Fields
   Dim MsgAns                         As VbMsgBoxResult
   Dim dbLocation                     As String
   Dim dBName                         As String

   On Error GoTo RECREATE_DATABASE_ERROR
   RECREATE_DATABASE = False

   dBName = DATABASE_PATH & "\" & DATABASE_FILENAME

   'Check If The Database File Exist
   If Dir$(dBName) <> "" Then
      MsgAns = MsgBox("Database - " & dBName & " already exist." & _
            vbNewLine & "Are you sure that you want to recreate it?", vbCritical + vbYesNo, "Create Database")

      If MsgAns = vbYes Then
         'Delete File
         Debug.Print "Deleting the File - " & dBName
         Call Kill(dBName)
      Else
         RECREATE_DATABASE = False
         Exit Function
      End If
   End If


   'Used To check If The Database Directory Exist
   If Dir$(DATABASE_PATH, vbDirectory) <> "" Then
      Debug.Print "The Database Directory Exist"
   Else   'If Directory Does Not Exist The Create
      Debug.Print "Creating A New Directory Exist"
      Call CreateNewDirectory(DATABASE_PATH)
   End If

   '::>>::---------------------------------------------------::<<::
   '::>>:: Create A New Database "Password Protected" (DAO)  ::<<::
   '::>>:: MS Access 97 Database                             ::<<::
   '::>>::---------------------------------------------------::<<::
   Debug.Print "clsFamily.Recreate : Creating A Password Protected Database : " & dBName
   Set NEW_DBASE = CreateDatabase(dBName, dbLangGeneral & ";pwd=" & DATABASE_PASSWORD)
   '::>>::---------------------------------------------------::<<::


   '::>>::---------------------------------------------------::<<::
   '::>>::      Create a new table for USER_TABLENAME        ::<<::
   ':: the value for USER_TABLENAME can be found in "modPublic"
   ':: This is used to store informations about ALL the Users
   '::>>::---------------------------------------------------::<<::

   Debug.Print "clsFamily.Recreate : Creating Table : " & USER_TABLENAME
   Set NEW_USERS_TABLE = NEW_DBASE.CreateTableDef(USER_TABLENAME)

   'Add These Fileds to the USER_TABLENAME Table
   'Note This is not the only method for creating fields
   Set TmpFields(1) = NEW_USERS_TABLE.CreateField("USER_LOGIN_NAME", DAO.dbText, MAX_USER_NAME_SIZE)
   TmpFields(1).AllowZeroLength = False
   NEW_USERS_TABLE.Fields.Append TmpFields(1)

   Set TmpFields(2) = NEW_USERS_TABLE.CreateField("USER_PASSWORD", DAO.dbText, MAX_USER_NAME_SIZE)
   TmpFields(2).AllowZeroLength = False
   NEW_USERS_TABLE.Fields.Append TmpFields(2)

   Set TmpFields(3) = NEW_USERS_TABLE.CreateField("USER_ACCESS_LEVEL", DAO.dbText, 13)
   TmpFields(3).AllowZeroLength = False
   NEW_USERS_TABLE.Fields.Append TmpFields(3)

   Set TmpFields(4) = NEW_USERS_TABLE.CreateField("FULL_NAME", DAO.dbText, MAX_FULLNAME_SIZE)
   TmpFields(4).AllowZeroLength = False
   NEW_USERS_TABLE.Fields.Append TmpFields(4)

   Set TmpFields(5) = NEW_USERS_TABLE.CreateField("USER_EMAIL_ADDDRESS", DAO.dbText, 50)
   TmpFields(5).AllowZeroLength = True
   NEW_USERS_TABLE.Fields.Append TmpFields(5)

   Set TmpFields(6) = NEW_USERS_TABLE.CreateField("USER_SMTP_SERVER", DAO.dbText, 50)
   TmpFields(6).AllowZeroLength = True
   NEW_USERS_TABLE.Fields.Append TmpFields(6)

   Set TmpFields(7) = NEW_USERS_TABLE.CreateField("USER_POP3_SERVER", DAO.dbText, 50)
   TmpFields(7).AllowZeroLength = True
   NEW_USERS_TABLE.Fields.Append TmpFields(7)

   Set TmpFields(8) = NEW_USERS_TABLE.CreateField("USER_HOMEPAGE_URL", DAO.dbText, 50)
   TmpFields(8).AllowZeroLength = True
   NEW_USERS_TABLE.Fields.Append TmpFields(8)

   Set TmpFields(9) = NEW_USERS_TABLE.CreateField("USER_LOCKED", DAO.dbBoolean)
   NEW_USERS_TABLE.Fields.Append TmpFields(9)

   Set TmpFields(10) = NEW_USERS_TABLE.CreateField("DATE_ADDED", DAO.dbDate)
   NEW_USERS_TABLE.Fields.Append TmpFields(10)


   'Append The Table To The Database
   NEW_DBASE.TableDefs.Append NEW_USERS_TABLE
   'Refresh The Database Tables
   NEW_DBASE.TableDefs.Refresh
   'Close Fields
   For TmpByte = 1 To 10
      Set TmpFields(TmpByte) = Nothing
   Next TmpByte
   'Close The Table
   Set NEW_USERS_TABLE = Nothing
   '::>>::---------------------------------------------------::<<::




   '::>>::---------------------------------------------------::<<::
   '::>>::     Create a new tabe called "CATEGORIES"         ::<<::
   '::>>::---------------------------------------------------::<<::
   '::Used to store all the CATEGORIES
   '::>>::---------------------------------------------------::<<::
   Debug.Print "clsFamily.Recreate : Creating Table : " & CATEGORIES_TABLENAME
   Set NEW_CATEGORIES_DBASE_TABLE = NEW_DBASE.CreateTableDef(CATEGORIES_TABLENAME)

   'Add These Fileds to the "CATEGORIES" Table
   'Note This is not the only method for creating fields

   Set TmpFields(1) = NEW_CATEGORIES_DBASE_TABLE.CreateField("USER_NAME", DAO.dbText, MAX_USER_NAME_SIZE)
   TmpFields(1).AllowZeroLength = False
   NEW_CATEGORIES_DBASE_TABLE.Fields.Append TmpFields(1)

   Set TmpFields(2) = NEW_CATEGORIES_DBASE_TABLE.CreateField("CATEGORY_NAME", DAO.dbText, MAX_CATEGORY_NAME_SIZE)
   TmpFields(2).AllowZeroLength = False
   NEW_CATEGORIES_DBASE_TABLE.Fields.Append TmpFields(2)

   'Append The Table To The Database
   NEW_DBASE.TableDefs.Append NEW_CATEGORIES_DBASE_TABLE
   'Refresh The Database Tables
   NEW_DBASE.TableDefs.Refresh
   'Close Fields
   For TmpByte = 1 To 2
      Set TmpFields(TmpByte) = Nothing
   Next TmpByte
   'Close The Table
   Set NEW_CATEGORIES_DBASE_TABLE = Nothing
   '::>>::---------------------------------------------------::<<::




   '::>>::---------------------------------------------------::<<::
   '::>>::     Create a new table for CONTACTS_TABLENAME     ::<<::
   '::>>::---------------------------------------------------::<<::
   ':: Note : The value for CONTACTS_TABLENAME can be found in "modPublic"
   ':: Used to store all the Personal CONTACTS
   '::>>::---------------------------------------------------::<<::
   Debug.Print "clsFamily.Recreate : Creating Table : " & CONTACTS_TABLENAME
   Set NEW_USERS_PERSONAL_DBASE_TABLE = NEW_DBASE.CreateTableDef(CONTACTS_TABLENAME)

   'Add Fields to CONTACTS_TABLENAME
   Set TmpFields(1) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("FirstName", DAO.dbText, MAX_FNAME_SIZE)
   TmpFields(1).AllowZeroLength = False
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(1)

   Set TmpFields(2) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("LastName", DAO.dbText, MAX_FNAME_SIZE)
   TmpFields(2).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(2)

   Set TmpFields(3) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("Sex", DAO.dbText, 6)
   TmpFields(3).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(3)

   Set TmpFields(4) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("Telephone", DAO.dbText, 20)
   TmpFields(4).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(4)

   Set TmpFields(5) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("Address", DAO.dbText, 50)
   TmpFields(5).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(5)

   Set TmpFields(6) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("City_State", DAO.dbText, 50)
   TmpFields(6).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(6)

   Set TmpFields(7) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("ZipCode", DAO.dbText, 11)
   TmpFields(7).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(7)

   Set TmpFields(8) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("EmailAddress", DAO.dbText, MAX_FULLNAME_SIZE)
   TmpFields(8).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(8)

   Set TmpFields(9) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("CATEGORY_NAME", DAO.dbText, MAX_CATEGORY_NAME_SIZE)
   TmpFields(9).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(9)

   Set TmpFields(10) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("LOCKED", DAO.dbBoolean)
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(10)

   Set TmpFields(11) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("Notes", DAO.dbMemo)
   TmpFields(11).AllowZeroLength = True
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(11)

   Set TmpFields(12) = NEW_USERS_PERSONAL_DBASE_TABLE.CreateField("USER_NAME", DAO.dbText, MAX_USER_NAME_SIZE)
   TmpFields(12).AllowZeroLength = False
   NEW_USERS_PERSONAL_DBASE_TABLE.Fields.Append TmpFields(12)


   'Append The Table To The Database
   Debug.Print "Trying to Append " & CONTACTS_TABLENAME & " to the database"
   NEW_DBASE.TableDefs.Append NEW_USERS_PERSONAL_DBASE_TABLE
   Debug.Print CONTACTS_TABLENAME & " has been Appended successfully to the Database"

   'Refresh The Database Tables
   Debug.Print "Refreshing Tables"
   NEW_DBASE.TableDefs.Refresh

   'Closing Table and Fields
   Debug.Print "Closing tables and fields"
   'Close Fields
   For TmpByte = 1 To 12
      Set TmpFields(TmpByte) = Nothing
   Next TmpByte
   'Close The Table
   Set NEW_USERS_PERSONAL_DBASE_TABLE = Nothing
   '::>>::---------------------------------------------------::<<::




   '::>>::---------------------------------------------------::<<::
   '::>>::   Create a new table called REMINDERS_TABLENAME   ::<<::
   '::>>::---------------------------------------------------::<<::
   ':: Note : The value for REMINDERS_TABLENAME can be found in "modPublic"
   ':: Used to store all the personal REMINDERS
   '::>>::---------------------------------------------------::<<::
   Debug.Print "Trying to create a new table called " & REMINDERS_TABLENAME
   Set NEW_REMINDERS_DBASE_TABLE = NEW_DBASE.CreateTableDef(REMINDERS_TABLENAME)
   Debug.Print REMINDERS_TABLENAME & " has been recreated successfully, trying to add fields"

   'Add Fields to REMINDERS_TABLENAME
   Set TmpFields(1) = NEW_REMINDERS_DBASE_TABLE.CreateField("USER_NAME", DAO.dbText, MAX_USER_NAME_SIZE)
   TmpFields(1).AllowZeroLength = False
   NEW_REMINDERS_DBASE_TABLE.Fields.Append TmpFields(1)

   'Date Entered
   Set TmpFields(2) = NEW_REMINDERS_DBASE_TABLE.CreateField("DATE_ENTERED", DAO.dbDate)
   NEW_REMINDERS_DBASE_TABLE.Fields.Append TmpFields(2)

   'Date Expired
   Set TmpFields(3) = NEW_REMINDERS_DBASE_TABLE.CreateField("DATE_EXPIRED", DAO.dbDate)
   NEW_REMINDERS_DBASE_TABLE.Fields.Append TmpFields(3)

   'Subject
   Set TmpFields(4) = NEW_REMINDERS_DBASE_TABLE.CreateField("SUBJECT", DAO.dbText, 50)
   TmpFields(4).AllowZeroLength = False
   NEW_REMINDERS_DBASE_TABLE.Fields.Append TmpFields(4)

   'TO DO
   Set TmpFields(5) = NEW_REMINDERS_DBASE_TABLE.CreateField("TODO", DAO.dbMemo)
   TmpFields(5).AllowZeroLength = False
   NEW_REMINDERS_DBASE_TABLE.Fields.Append TmpFields(5)

   'Append The Table To The Database
   Debug.Print "Trying to Append " & REMINDERS_TABLENAME & " to the database"
   NEW_DBASE.TableDefs.Append NEW_REMINDERS_DBASE_TABLE

   Debug.Print REMINDERS_TABLENAME & " has been successfully added to database, now trying to refresh the database"
   'Refresh
   NEW_DBASE.TableDefs.Refresh

   'Closing Table and Fields
   Debug.Print "Now trying to close field"
   For TmpByte = 1 To 5
      Set TmpFields(TmpByte) = Nothing
   Next TmpByte
   'Close The Table
   Set NEW_REMINDERS_DBASE_TABLE = Nothing
   '::>>::---------------------------------------------------::<<::



   '::>>::---------------------------------------------------::<<::
   '::>>          Create a new tabe called "LINKS"           ::<<::
   '::>>::---------------------------------------------------::<<::
   'Used to store all the CATEGORIES
   Debug.Print "clsFamily.Recreate : Creating Table : " & LINKS_TABLENAME
   Set NEW_CATEGORIES_DBASE_TABLE = NEW_DBASE.CreateTableDef(LINKS_TABLENAME)

   'Add These Fileds to the "LINKS" Table
   'Note This is not the only method for creating fields

   Set TmpFields(1) = NEW_CATEGORIES_DBASE_TABLE.CreateField("USER_NAME", DAO.dbText, MAX_USER_NAME_SIZE)
   TmpFields(1).AllowZeroLength = False
   NEW_CATEGORIES_DBASE_TABLE.Fields.Append TmpFields(1)

   Set TmpFields(2) = NEW_CATEGORIES_DBASE_TABLE.CreateField("LINK", DAO.dbText, MAX_LINKS_SIZE)
   TmpFields(2).AllowZeroLength = False
   NEW_CATEGORIES_DBASE_TABLE.Fields.Append TmpFields(2)

   'Append The Table To The Database
   NEW_DBASE.TableDefs.Append NEW_CATEGORIES_DBASE_TABLE
   'Refresh The Database Tables
   NEW_DBASE.TableDefs.Refresh
   'Close Fields
   For TmpByte = 1 To 2
      Set TmpFields(TmpByte) = Nothing
   Next TmpByte
   'Close The Table
   Set NEW_CATEGORIES_DBASE_TABLE = Nothing
   '::>>::---------------------------------------------------::<<::



   '::>>::---------------------------------------------------::<<::
   '::>>                    Close The Database               ::<<::
   '::>>::---------------------------------------------------::<<::
   Debug.Print "Closing Database : " & dBName
   NEW_DBASE.Close
   Set NEW_DBASE = Nothing
   RECREATE_DATABASE = True
   Debug.Print "DATABASE RECREATED AND CLOSED SUCCESSFULLY"
   Exit Function
   '::>>::---------------------------------------------------::<<::


RECREATE_DATABASE_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : clsDBase.RECREATE_DATABASE" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & vbNewLine & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      RECREATE_DATABASE = False
   End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::




'================================================================
'Used to add a new category
'================================================================
Public Function ADD_NEW_USER_CATEGORY(ByVal User_Name As String, ByVal CATEGORY_NAME As String) As Boolean
   Dim cRecordset       As ADODB.Recordset
   On Error GoTo ADD_NEW_USER_CATEGORY_ERROR

   ADD_NEW_USER_CATEGORY = False

   Set cRecordset = New ADODB.Recordset

   'cConnection.Open CONNECTION_STRING
   cRecordset.Open CATEGORIES_TABLENAME, PUBLIC_CONNECTION, adOpenKeyset, adLockOptimistic
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
      Debug.Print "ADD_NEW_USER_CATEGORY_ERROR : " & Err.Description
      Err.Clear
      Set cRecordset = Nothing
      Exit Function
   End If
End Function
'================================================================
'================================================================




'================================================================
' Adds The New User
'================================================================
Public Function ADD_NEW_USER(ByVal pUSER_USERNAME As String, _
       ByVal pUSER_PASSWORD As String, _
       ByVal pFULL_NAME As String, _
       ByVal pUSER_ACCESSLEVEL As String, _
       Optional pUSER_EMAILADDRESS As String = "username@yourmail.com", _
       Optional pUSER_URL As String = "http://www.justprograms.cjb.net", _
       Optional pUSER_SMTP_SERVER As String = "smtp.your_isp.com") As Boolean

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

   PUBLIC_CONNECTION.Execute TmpString
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
   Call ADD_NEW_USER_CATEGORY(pUSER_USERNAME, "Acquaintance")
   Call ADD_NEW_USER_CATEGORY(pUSER_USERNAME, "Co-Worker")
   Call ADD_NEW_USER_CATEGORY(pUSER_USERNAME, "Family")
   Call ADD_NEW_USER_CATEGORY(pUSER_USERNAME, "Friend")
   Call ADD_NEW_USER_CATEGORY(pUSER_USERNAME, "Spouse")

   Debug.Print "Default Categories Added Succesfully for " & pUSER_USERNAME


   ADD_NEW_USER = True
   Exit Function

ADD_NEW_USER_Error:
   If Err.Number <> 0 Then
      ADD_NEW_USER = False
      MsgBox "Error (ADD_NEW_USER) : " & Err.Description & " : " & Str(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'================================================================
'================================================================



'================================================================
'USED TO CONNECT TO THE DATABASE
'================================================================
Public Function OPEN_CONNECTION() As Boolean
   On Error GoTo OPEN_CONNECTION_ERROR

   OPEN_CONNECTION = False
   Set PUBLIC_CONNECTION = New ADODB.Connection

   'The Database Directory
   Path = AddBackSlash(App.Path) & DATABASE_PATH & "\"

   'Setup the Default Connection String
   CONNECTION_STRING = "DRIVER={Microsoft Access Driver (*.mdb)};" & _
         "pwd=" & DATABASE_PASSWORD & ";" & _
         "UID=Admin;" & _
         "DBQ=" & Path & DATABASE_FILENAME


   PUBLIC_CONNECTION.Open CONNECTION_STRING
   DoEvents
   
   If PUBLIC_CONNECTION.State = 1 Then
      OPEN_CONNECTION = True
   Else
      OPEN_CONNECTION = False
   End If


OPEN_CONNECTION_ERROR:
   If Err.Number <> 0 Then
      OPEN_CONNECTION = False
      MsgBox "Error (OPEN_CONNECTION) : " & Err.Description & " : " & Str(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'================================================================


Private Sub Class_Terminate()
   Set PUBLIC_CONNECTION = Nothing
End Sub


