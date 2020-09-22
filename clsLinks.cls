VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Event ERROR(ByVal Error_Msg As String, ByVal ErrNum As Long)
Public USER_LINKS As ADODB.Recordset

'=====================================================================
'Used to add a link
'=====================================================================
Public Function ADD_LINK(ByVal USERNAME As String, ByVal TheLink As String) As Boolean
   On Error GoTo ADD_LINK_ERROR

   Debug.Print "Trying to add a new link"
   TheLink = modPublic.Apostrophe(Trim$(TheLink))
   TheLink = modPublic.ADD_HTTP(TheLink)

   Set FAMILY.FORM_LINKS.MyUser.USER_LINKS = New ADODB.Recordset
   FAMILY.FORM_LINKS.MyUser.USER_LINKS.Open LINKS_TABLENAME, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   FAMILY.FORM_LINKS.MyUser.USER_LINKS.AddNew
   FAMILY.FORM_LINKS.MyUser.USER_LINKS.Fields("LINK") = TheLink
   FAMILY.FORM_LINKS.MyUser.USER_LINKS.Fields("USER_NAME") = USERNAME
   'Update and Then Close
   FAMILY.FORM_LINKS.MyUser.USER_LINKS.Update
   FAMILY.FORM_LINKS.MyUser.USER_LINKS.Close
   Debug.Print "Link added"
   ADD_LINK = True
   Exit Function
ADD_LINK_ERROR:
   If Err.Number <> 0 Then
      Debug.Print "clsLinks.ADD_LINK_ERROR : " & Err.Description & " : " & Str$(Err.Number)
      MsgBox "clsLinks.ADD_LINK_ERROR : " & Err.Description & " : " & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'=====================================================================


'=====================================================================
'Used to Edit a link
'=====================================================================
Public Function EDIT_LINK(ByVal USERNAME As String, ByVal THE_OLD_LINK As String, ByVal THE_NEW_LINK As String) As Boolean
   Dim TMP_STR As String

   On Error GoTo EDIT_LINK_ERROR

   Debug.Print "Trying to EDIT a link"

   THE_NEW_LINK = modPublic.Apostrophe(Trim$(THE_NEW_LINK))

   THE_NEW_LINK = modPublic.ADD_HTTP(THE_NEW_LINK)

   TMP_STR = "UPDATE " & LINKS_TABLENAME & " SET LINK  = '" & THE_NEW_LINK & "'" & _
         " WHERE USER_NAME = '" & USERNAME & "'" & _
         " AND LINK = '" & THE_OLD_LINK & "'"

   PUBLIC_DATABASE.CONNECTION.Execute TMP_STR
   DoEvents
   EDIT_LINK = True
   Exit Function

EDIT_LINK_ERROR:
   If Err.Number <> 0 Then
      EDIT_LINK = False
      Debug.Print "clsLinks.EDIT_LINK_ERROR : " & Err.Description & " : " & Str$(Err.Number)
      MsgBox "clsLinks.ADD_LINK_ERROR : " & Err.Description & " : " & Str$(Err.Number), vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'=====================================================================



'=====================================================================
'Used To Check If A Link Exist
'=====================================================================
Public Function LinkExist(ByVal USERNAME As String, ByVal TheLink As String) As Boolean
   Dim TmpRecordSet As ADODB.Recordset
   On Error GoTo LinkExist_ERROR

   tmpSQL = "SELECT LINK FROM " & LINKS_TABLENAME & " " & _
         "WHERE USER_NAME = '" & USERNAME & "' " & _
         "AND LINK = '" & TheLink & "'"

   Set TmpRecordSet = New ADODB.Recordset
   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockPessimistic

   LinkExist = False
   If TmpRecordSet.RecordCount > 0 Then
      LinkExist = True
   Else
      LinkExist = False
   End If
   TmpRecordSet.Close
   Set TmpRecordSet = Nothing

LinkExist_ERROR:
   If Err.Number <> 0 Then
      MsgBox "LinkExist_ERROR : " & Err.Description & Str$(Err.Number), vbCritical + vbOKOnly
      RaiseEvent ERROR("LinkExist_ERROR : " & Err.Description, Err.Number)
      Err.Clear
   End If
End Function
'=====================================================================



'=====================================================================
'USED To Delete A Link
'=====================================================================
Public Function DELETE_LINK(ByVal USERNAME As String, ByVal THE_LINK As String) As Boolean
   On Error GoTo DELETE_LINK_ERROR

   tmpSQL = "DELETE * FROM " & LINKS_TABLENAME & _
         " WHERE LINK = '" & THE_LINK & "'" & _
         " AND USER_NAME = '" & USERNAME & "'"

   PUBLIC_DATABASE.CONNECTION.Execute tmpSQL
   DoEvents
   Exit Function

DELETE_LINK_ERROR:
   If Err.Number <> 0 Then
      MsgBox "LinkExist_ERROR : " & Err.Description & Str$(Err.Number), vbCritical + vbOKOnly
      RaiseEvent ERROR("LinkExist_ERROR : " & Err.Description, Err.Number)
      Err.Clear
   End If
End Function
'=====================================================================




'=====================================================================
'=====================================================================
Public Function LinkCount(ByVal USERNAME As String) As Long
   On Error GoTo LinkCount_ERROR

   tmpSQL = "SELECT LINK FROM " & LINKS_TABLENAME & " " & _
         "WHERE USER_NAME = '" & USERNAME & "'"

   Set TmpRecordSet = New ADODB.Recordset
   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockPessimistic

   LinkCount = TmpRecordSet.RecordCount

   TmpRecordSet.Close
   Set TmpRecordSet = Nothing

LinkCount_ERROR:
   If Err.Number <> 0 Then
      MsgBox "LinkCount_ERROR : " & Err.Description & Str$(Err.Number), vbCritical + vbOKOnly
      RaiseEvent ERROR("LinkCount_ERROR : " & Err.Description, Err.Number)
      Err.Clear
   End If
End Function
'=====================================================================


'=====================================================================
'USED TO THE USER'S LOAD_LINKS
'=====================================================================
Public Function LOAD_LINKS() As Boolean
   Dim TmpCnt As Long

   On Error GoTo LOAD_LINKS_ERROR

   Set TmpRecordSet = New ADODB.Recordset
   tmpSQL = "SELECT LINK FROM " & LINKS_TABLENAME & _
         " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'"

   TmpRecordSet.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   If TmpRecordSet.RecordCount > 0 Then
      TmpCnt = 0
      Do While Not TmpRecordSet.EOF
         TmpCnt = TmpCnt + 1
         FAMILY.FORM_MAIN.mnuLink(TmpCnt).Caption = TmpRecordSet.Fields("LINK")
         TmpRecordSet.MoveNext
         If TmpCnt >= 10 Then
            Exit Do
         End If
      Loop
   End If

LOAD_LINKS_ERROR:
   If Err.Number <> 0 Then
      MsgBox "LOAD_LINKS_ERROR : " & Err.Description & Str$(Err.Number), vbCritical + vbOKOnly
      RaiseEvent ERROR("LOAD_LINKS_ERROR : " & Err.Description, Err.Number)
      Err.Clear
   End If

End Function
'=====================================================================
