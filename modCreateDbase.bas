Attribute VB_Name = "modCreateDbase"
'================================================================
'================================================================
'***                  TO GOD BE THE GLORY                     ***
'================================================================
'================================================================
'*** For any Questions or Comments concerning this program    ***
'*** Homepage : http://www.omarswan.cjb.net                   ***
'*** Email    : omarswan@yahoo.com                            ***
'*** AOL      : smileyomar  or omarsmiley                     ***
'================================================================
'================================================================
'* Deducated to SmileyOrange -> http://www.smileyorange.cjb.net *
'================================================================
'================================================================

Option Explicit

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


'-----------------------------------------------------------------
' SOME DATABASE STUFF
'-----------------------------------------------------------------
Public Const DATABASE_FILENAME         As String = "FamilyDB.FM3"
Public Const DATABASE_PASSWORD         As String = "SmileyOmar"
'This Will Be Used To Store The Database
Public Const DATABASE_PATH             As String = "Dbase"
' This Stores The Database Version
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
Public Const LINKS_TABLENAME        As String = "LINKS"
'::--------------------------------------------------------------::

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

'Maximum size of a Category Name
Public Const MAX_CATEGORY_NAME_SIZE     As Byte = 25

Public Const MAX_LINKS_SIZE As Byte = 50
'::--------------------------------------------------------------::
'::                       Temporary Variables                    ::
'::--------------------------------------------------------------::
Public TmpByte                         As Byte
Public TmpInt                          As Integer
Public TmpString                       As String
Public tmpSQL                          As String
Public TmpMsgResult                    As VbMsgBoxResult
Public myConnection                    As ADODB.Connection
Public CONNECTION_STRING               As String
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


