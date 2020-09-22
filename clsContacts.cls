VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'Recordsets
Public USER_CONTACTS                    As ADODB.Recordset

Public Event ERROR(Description As String, ERROR_ID As Long)


'================================================================
'USED TO ADD A NEW RECORD (CONTACTS)
'================================================================
Public Function ADD_RECORD(ByVal FirstName As String, _
                           ByVal LastName As String, _
                           ByVal Sex As String, _
                           ByVal Telephone As String, _
                           ByVal Address As String, _
                           ByVal City_State As String, _
                           ByVal ZipCode As String, _
                           ByVal EmailAddress As String, _
                           ByVal CATEGORY_NAME As String, _
                           ByVal Notes As String, _
                           ByVal USERNAME As String) As Boolean

   Dim tmpSQL As String
   On Error GoTo ADD_RECORD_ERROR

   tmpSQL = ""
   tmpSQL = "SELECT FirstName,LastName,Sex,Telephone,Address,City_State,ZipCode,EmailAddress,CATEGORY_NAME,Notes,USER_NAME FROM " & CONTACTS_TABLENAME
   tmpSQL = tmpSQL & " WHERE USER_NAME = '" & USERNAME & "'"

   Set FAMILY.FORM_MAIN.MyUser.USER_CONTACTS = New ADODB.Recordset
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.AddNew
   'First Name
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("FirstName") = "" & ProperCase(Trim$(FirstName))
   'Last Name
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("LastName") = "" & ProperCase(Trim$(LastName))
   'Sex
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("Sex") = Trim$(Sex)
   'Telephone
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("Telephone") = "" & Trim$(Telephone)
   'Address
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("Address") = "" & ProperCase(Trim$(Address))
   'City-State
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("City_State") = "" & ProperCase(Trim$(City_State))
   'Zip Code
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("ZipCode") = "" & Trim$(ZipCode)
   'Email
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("EmailAddress") = "" & Trim$(EmailAddress)
   'Relation/Category
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("CATEGORY_NAME") = Trim$(CATEGORY_NAME)
   'Notes
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("Notes") = Trim$(Notes)
   'User_Name
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("USER_NAME") = "" & USERNAME

   'Update Database
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Update
   DoEvents
   ADD_RECORD = True
   Debug.Print "clsContacts.ADD_RECORD - Record Added"
   Exit Function

ADD_RECORD_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : clsContacts.ADD_RECORD" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      ADD_RECORD = False
   End If
End Function
'================================================================



'================================================================
'USED TO REMOVE A RECORD (CONTACTS)
'================================================================
Public Function DELETE_RECORD(ByVal FirstName As String, _
                              ByVal LastName As String, _
                              ByVal CategoryName As String, _
                              ByVal User_Name As String) As Boolean
   Dim tmpSQL As String
   On Error GoTo DELETE_RECORD_ERROR

   DELETE_RECORD = False

   tmpSQL = ""
   tmpSQL = "DELETE * FROM " & CONTACTS_TABLENAME & _
         " WHERE FirstName = '" & FirstName & "'" & _
         " AND LastName = '" & LastName & "'" & _
         " AND CATEGORY_NAME = '" & CategoryName & "'" & _
         " AND USER_NAME = '" & User_Name & "'"

   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   DELETE_RECORD = True

DELETE_RECORD_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : clsContacts.DELETE_RECORD" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      DELETE_RECORD = False
   End If
End Function
'================================================================



'================================================================
'USED TO CHECK IF A RECORD EXIST (CONTACTS)
'================================================================
Public Function RECORD_EXIST(ByVal FirstName As String, _
                              ByVal LastName As String, _
                              ByVal CategoryName As String, _
                              ByVal User_Name As String) As Boolean
   Dim tmpSQL As String
   On Error GoTo RECORD_EXIST_ERROR

   tmpSQL = ""
   tmpSQL = "SELECT USER_NAME FROM " & CONTACTS_TABLENAME & _
         " WHERE FirstName = '" & FirstName & "'" & _
         " AND LastName = '" & LastName & "'" & _
         " AND CATEGORY_NAME = '" & CategoryName & "'" & _
         " AND USER_NAME = '" & User_Name & "'"

   Set TmpRecordSet = New ADODB.Recordset
   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   If TmpRecordSet.RecordCount > 0 Then
      RECORD_EXIST = True
      Debug.Print "Record exist. ;-("
   Else
      RECORD_EXIST = False
      Debug.Print "Record Does Not exist! Yes ;-)"
   End If

   tmpSQL = ""
   Set TmpRecordSet = Nothing

RECORD_EXIST_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : clsContacts.RECORD_EXIST" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      RECORD_EXIST = True
   End If
End Function
'================================================================

