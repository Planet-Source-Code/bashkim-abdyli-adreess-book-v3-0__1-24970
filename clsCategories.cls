VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


Option Explicit

Public USER_CATEGORY As ADODB.Recordset
Public Event ERROR(ByVal MSG As String)


Private Sub Class_Initialize()
   Set USER_CATEGORY = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Set USER_CATEGORY = Nothing
End Sub


'================================================================
'USED TH DELETE A CATEGORY
'================================================================
Public Function DELETE_CATEGORY(ByVal CategoryName As String, ByVal User_Name As String) As Boolean
   On Error GoTo DELETE_CATEGORY_ERROR

   'Remove all The Contacts Record from CONTACTS_TABLENAME
   TmpString = ""
   TmpString = "DELETE * FROM " & CONTACTS_TABLENAME & _
         " WHERE USER_NAME = '" & User_Name & "'" & _
         " AND CATEGORY_NAME = '" & CategoryName & "'"

   PUBLIC_DATABASE.CONNECTION.Execute TmpString
   DoEvents

   'Remove all The Contacts Record from CONTACTS_TABLENAME
   TmpString = ""
   TmpString = "DELETE * FROM " & CATEGORIES_TABLENAME & _
         " WHERE USER_NAME = '" & User_Name & "'" & _
         " AND CATEGORY_NAME = '" & CategoryName & "'"

   PUBLIC_DATABASE.CONNECTION.Execute TmpString
   DoEvents

   DELETE_CATEGORY = True
   Exit Function

DELETE_CATEGORY_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : clsCategories.DELETE_CATEGORY" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      DELETE_CATEGORY = False
      RaiseEvent ERROR(Err.Description & " : " & Str$(Err.Number))
   End If
End Function
'================================================================



'================================================================
'USED TO EDIT A CATEGORY
'================================================================
Public Function EDIT_CATEGORY(ByVal OldCategory As String, ByVal NewCategory As String, ByVal User_Name As String) As Boolean
   On Error GoTo EDIT_CATEGORY_ERROR

   'Change The Category Names of all the records that matches
   Debug.Print "Renaming The Category Names From " & CONTACTS_TABLENAME

   TmpString = ""
   TmpString = "UPDATE " & CONTACTS_TABLENAME & " SET CATEGORY_NAME = '" & Trim$(NewCategory) & "'" & _
         " WHERE USER_NAME = '" & User_Name & "'" & _
         " AND CATEGORY_NAME = '" & OldCategory & "'"

   PUBLIC_DATABASE.CONNECTION.Execute TmpString
   DoEvents


   'Rename Category Name from The CATEGORY TABLE
   Debug.Print "Renaming The Category Names From " & CATEGORIES_TABLENAME

   TmpString = ""
   TmpString = "UPDATE " & CATEGORIES_TABLENAME & " SET CATEGORY_NAME = '" & Trim$(NewCategory) & "'" & _
         " WHERE USER_NAME = '" & User_Name & "'" & _
         " AND CATEGORY_NAME ='" & OldCategory & "'"

   PUBLIC_DATABASE.CONNECTION.Execute TmpString
   DoEvents

   EDIT_CATEGORY = True

EDIT_CATEGORY_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : clsCategories.EDIT_CATEGORY" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      EDIT_CATEGORY = False
      RaiseEvent ERROR(Err.Description & " : " & Str$(Err.Number))
   End If
End Function
'================================================================




'================================================================
'USED TO ADD A New CATEGORY
'================================================================
Public Function ADD_CATEGORY(ByVal CategoryName As String, ByVal User_Name As String) As Boolean
   On Error GoTo ADD_CATEGORY_ERROR

   Debug.Print "Adding A New Category For"
   'Initialise The Recordset
   Set TmpRecordSet = New ADODB.Recordset

   tmpSQL = ""
   tmpSQL = "SELECT USER_NAME,CATEGORY_NAME FROM " & CATEGORIES_TABLENAME & _
         " WHERE USER_NAME = '" & User_Name & "'"

   'OPEN THE Categories Table
   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   TmpRecordSet.AddNew
   TmpRecordSet.Fields("USER_NAME") = CURRENT_USER.LOGIN_NAME
   TmpRecordSet.Fields("CATEGORY_NAME") = Trim$(CategoryName)
   TmpRecordSet.Update   'UPDATE
   TmpRecordSet.Requery   'REQUERY
   TmpRecordSet.Close   'CLOSE
   Set TmpRecordSet = Nothing

   ADD_CATEGORY = True
   Exit Function

ADD_CATEGORY_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : clsCategories.ADD_CATEGORY" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      ADD_CATEGORY = False
      RaiseEvent ERROR(Err.Description & " : " & Str$(Err.Number))
   End If
End Function
'================================================================



'================================================================
'Used to Check If a CATEGORY NAME ALREADY EXIST
'================================================================
Public Function USER_CATEGORY_EXIST(ByVal CATEGORY_NAME As String, ByVal User_Name As String) As Boolean
   Dim TmpRecordSet As ADODB.Recordset
   On Error GoTo USER_CATEGORY_EXIST_ERROR

   Set TmpRecordSet = New ADODB.Recordset

   tmpSQL = ""
   tmpSQL = "SELECT CATEGORY_NAME FROM " & CATEGORIES_TABLENAME & _
         " WHERE USER_NAME = '" & User_Name & "'" & _
         " AND CATEGORY_NAME = '" & CATEGORY_NAME & "'"

   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   DoEvents

   If TmpRecordSet.RecordCount > 0 Then
      USER_CATEGORY_EXIST = True
   Else
      USER_CATEGORY_EXIST = False
   End If

   tmpSQL = ""
   Set TmpRecordSet = Nothing
   Exit Function

USER_CATEGORY_EXIST_ERROR:
   If Err.Number <> 0 Then
      MsgBox "USER_CATEGORY_EXIST_ERROR : " & Err.Description & vbNewLine & " - Error# " & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
      Set TmpRecordSet = Nothing
      USER_CATEGORY_EXIST = True
      RaiseEvent ERROR(Err.Description & " : " & Str$(Err.Number))
   End If
End Function
'================================================================
'================================================================
