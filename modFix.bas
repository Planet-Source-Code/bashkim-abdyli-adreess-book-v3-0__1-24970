Attribute VB_Name = "modFix"
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

'The Password For Te Database
Public Const DATABASE_PASSWORD         As String = "SmileyOmar"
'Name of the table used to store all the info about all users
Public Const USER_TABLENAME            As String = "USERS"

Public PUBLIC_DATABASE                 As clsFixDbase
Public TmpRecordset                    As ADODB.Recordset
Public TmpSQL                          As String



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


'=====================================================================
'Used To LOG_OUT  a user
'=====================================================================
Public Function LOGOUT_USER(ByVal User_Name As String) As Boolean
   '  On Error GoTo LOGOUT_USER_ERROR

   TmpSQL = ""
   TmpSQL = "SELECT USER_LOGIN_NAME, USER_LOCKED FROM " & USER_TABLENAME & _
         " WHERE USER_LOGIN_NAME = '" & User_Name & "'" & _
         " ORDER BY USER_LOGIN_NAME ASC"

   Set TmpRecordset = New ADODB.Recordset
   TmpRecordset.Open TmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   If TmpRecordset.RecordCount > 0 Then
      TmpRecordset.Fields("USER_LOCKED") = False
      TmpRecordset.Update
      LOGOUT_USER = True
   Else
      LOGOUT_USER = False
   End If

   TmpSQL = ""
   Set TmpRecordset = Nothing
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



Public Sub Main()
   Set PUBLIC_DATABASE = New clsFixDbase
   If PUBLIC_DATABASE.OPEN_DATABASE_CONNECTION = False Then
      MsgBox "The database was not found or could not be loaded." & vbNewLine & "Read The File README.doc", vbCritical + vbOKOnly
      Call frmFixLogin.btnExit_Click
   End If
End Sub

