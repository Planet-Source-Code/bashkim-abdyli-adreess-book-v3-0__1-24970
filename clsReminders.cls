VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event ERROR(ByVal Error_Msg As String, ByVal ErrNum As Long)

'Recordset
Public USER_REMINDER                   As ADODB.Recordset


Private Sub Class_Initialize()
   'Create A New Recordset
   Set USER_REMINDER = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Set USER_REMINDER = Nothing
End Sub
