VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Event ERROR(ByVal Error_Msg As String)

Public USER_PROFILE As ADODB.Recordset


Private Sub Class_Initialize()
   Set USER_PROFILE = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Set USER_PROFILE = Nothing
End Sub


'=====================================================================
'Checks If A User User Exist
'=====================================================================
Public Function USER_EXIST(ByVal USERNAME As String) As Boolean
   Dim TmpRecordSet As ADODB.Recordset
   Dim tmpStr As String

   On Error GoTo USER_EXIST_ERROR

   Set TmpRecordSet = New ADODB.Recordset
   tmpStr = "SELECT USER_LOGIN_NAME FROM " & USER_TABLENAME & _
         " WHERE USER_LOGIN_NAME = '" & USERNAME & "'"

   TmpRecordSet.Open tmpStr, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockReadOnly

   TmpRecordSet.Requery

   If TmpRecordSet.RecordCount > 0 Then
      USER_EXIST = True
   Else
      USER_EXIST = False
   End If

   TmpRecordSet.Close
   Set TmpRecordSet = Nothing

USER_EXIST_ERROR:
   If Err.Number <> 0 Then
      RaiseEvent ERROR("USER_EXIST_ERROR : " & Err.Description & " " & Str$(Err.Number))
      MsgBox "ERROR : clsUSER.USER_EXIST" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      USER_EXIST = True
   End If
End Function
'=====================================================================



'=====================================================================
'Used To LOG_OUT  a user
'=====================================================================
Public Function LOGOUT_USER(ByVal User_Name As String) As Boolean
   On Error GoTo LOGOUT_USER_ERROR

   tmpSQL = ""
   tmpSQL = "SELECT USER_LOGIN_NAME, USER_LOCKED FROM " & USER_TABLENAME & _
         " WHERE USER_LOGIN_NAME = '" & User_Name & "'" & _
         " ORDER BY USER_LOGIN_NAME ASC"

   Set TmpRecordSet = New ADODB.Recordset
   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   If TmpRecordSet.RecordCount > 0 Then
      TmpRecordSet.Fields("USER_LOCKED") = False
      TmpRecordSet.Update
      LOGOUT_USER = True
      Debug.Print User_Name & " has been logged out"
   Else
      LOGOUT_USER = False
      Debug.Print User_Name & " not has been logged out"
   End If

   tmpSQL = ""
   Set TmpRecordSet = Nothing
   Exit Function

LOGOUT_USER_ERROR:
   If Err.Number <> 0 Then
      MsgBox "ERROR : clsUser.LOGOUT_USER" & vbNewLine & _
            "ERROR # " & Str$(Err.Number) & _
            "DESCRIPTION - " & Err.Description & vbNewLine, vbCritical + vbOKOnly
      Err.Clear
      LOGOUT_USER = False
   End If
End Function
'=====================================================================
