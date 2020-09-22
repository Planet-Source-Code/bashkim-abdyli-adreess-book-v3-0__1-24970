VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFixDbase"
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
'*** Email    : omarswan@yahoo.com                            ***
'*** AOL      : smileyomar  or omarsmiley                     ***
'================================================================
'================================================================
'* Deducated to SmileyOrange -> http://www.smileyorange.cjb.net *
'================================================================
'================================================================

Option Explicit

':://::-----------------------------------------------------::\\::
'The Public Database Connection
Public CONNECTION As ADODB.CONNECTION
':://::-----------------------------------------------------::\\::


'local variable(s) to hold property value(s)
Private mvarFILENAME As String   'local copy
Private mvarPATH As String   'local copy

'local variable(s) to hold property value(s)
Private mvarCONNECTION_STRING As String   'local copy
'local variable(s) to hold property value(s)
Private mvarPASSWORD As String   'local copy


Public Event CLOSED()
Public Event CONNECTED()
Public Event ERROR(ByVal The_Message As String, ByVal ErrNum As Long)
Public Event MESSAGE(ByVal The_Message As String)

Public Property Let PASSWORD(ByVal vData As String)
   mvarPASSWORD = vData
End Property


Public Property Get PASSWORD() As String
   PASSWORD = mvarPASSWORD
End Property



Public Property Let CONNECTION_STRING(ByVal vData As String)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.CONNECTION_STRING = 5
   mvarCONNECTION_STRING = vData
End Property


Public Property Get CONNECTION_STRING() As String
   CONNECTION_STRING = mvarCONNECTION_STRING
End Property


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'Used To Open The Database Connection
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Function OPEN_DATABASE_CONNECTION() As Boolean

   On Error GoTo OPEN_DATABASE_CONNECTION_ERROR

   'Connecting To The Database
   Debug.Print "clsDBAse.OPEN_DATABASE_CONNECTION - Connecting To The Database : "
   Set CONNECTION = New ADODB.CONNECTION

   CONNECTION.Open mvarCONNECTION_STRING
   DoEvents

   If CONNECTION.State = 1 Then
      RaiseEvent CONNECTED
      OPEN_DATABASE_CONNECTION = True
   Else
      OPEN_DATABASE_CONNECTION = False
   End If


OPEN_DATABASE_CONNECTION_ERROR:
   If Err.Number <> 0 Then
      OPEN_DATABASE_CONNECTION = False
      MsgBox "OPEN_DATABASE_CONNECTION_ERROR" & vbNewLine & _
            "Description : " & Err.Description & vbNewLine & _
            "ERROR #     : " & Str$(Err.Number) & vbNewLine & _
            "LOCATION    : " & mvarPATH & mvarFILENAME, vbCritical + vbOKOnly
   End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Property Let PATH(ByVal vData As String)
   mvarPATH = vData
End Property
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


Public Property Get PATH() As String
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.PATH
   PATH = mvarPATH
End Property


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Property Let FILENAME(ByVal vData As String)
   mvarFILENAME = vData
End Property
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Property Get FILENAME() As String
   FILENAME = mvarFILENAME
End Property
'::>>-------------------------------------------------------------<<::


'::>>-------------------------------------------------------------<<::
Private Sub Class_Initialize()
   'The Database Password
   mvarPASSWORD = "SmileyOmar"
   'The Database Directory
   mvarPATH = AddBackSlash(App.PATH) & AddBackSlash("Dbase")
   'The Database FileName
   mvarFILENAME = "FamilyDB.FM3"

   'Setup the Default Connection String
   mvarCONNECTION_STRING = "DRIVER={Microsoft Access Driver (*.mdb)};" & _
         "pwd=" & DATABASE_PASSWORD & ";" & _
         "UID=Admin;" & _
         "DBQ=" & PATH_AND_FILENAME

End Sub
'::>>-------------------------------------------------------------<<::


'::>>-------------------------------------------------------------<<::
Private Sub Class_Terminate()
   mvarPATH = ""
   mvarFILENAME = ""
   mvarCONNECTION_STRING = ""
   Set CONNECTION = Nothing
End Sub
'::>>-------------------------------------------------------------<<::


'::>>-------------------------------------------------------------<<::
Public Function PATH_AND_FILENAME() As String
   PATH_AND_FILENAME = AddBackSlash(mvarPATH) & mvarFILENAME
End Function
'::>>-------------------------------------------------------------<<::



