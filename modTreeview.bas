Attribute VB_Name = "modTreeview"

Option Explicit


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Enum TVMColorProps
   TVM_GETBKCOLOR = 4383
   TVM_GETTEXTCOLOR = 4384
   TVM_SETBKCOLOR = 4381
   TVM_SETTEXTCOLOR = 4382
End Enum

Public Const GWL_STYLE As Long = -16
Public Const TVIS_BOLD As Long = 16
Public Const TVIF_STATE As Long = 8
Public Const TVS_HASLINES As Long = 2



'================================================================
'================================================================
Public Function LOAD_frmMAIN_TREEVIEW(ByVal ObjForm As Form, _
                             ByVal TView As TreeView, _
                             ByVal IMG_List As ImageList, _
                             Optional ByVal Expand_This_Node As String = "ROOT" _
                             ) As Boolean

   Dim TmpNode        As Node
   Dim TmpString      As String
   Dim tmpNameString  As String
   Dim tmpSQL         As String
   On Error GoTo Load_frmMain_TreeView_error

   LOAD_frmMAIN_TREEVIEW = False

   'Disable The Form until the treeview has been loaded
   ObjForm.Enabled = False

   'Clear the treeview and node
   Set TmpNode = Nothing
   TView.Nodes.Clear


   'Set The TreeView Image List
   Set TView.ImageList = IMG_List

   'Call SetTVBackgroundColor(TXTBOX_BGROUND_COLOR, TView)

   'This is Used to Add The "ROOT" Node
   Set TmpNode = TView.Nodes.Add(, , "ROOT", "People", "closed_book", "closed_book")
   'Store Some Information In The Node's Tag - I will use this info. later
   TView.Nodes("ROOT").Tag = "ROOT"
   Debug.Print "Load Treeview Data"

   'Clear The Combox(1) - Category/Relation
   ObjForm.ComboBox(1).Clear

   Debug.Print "Load_frmMain_TreeView : Loading CATEGORY RECORDSET"
   tmpSQL = ""
   tmpSQL = "SELECT CATEGORY_NAME FROM " & CATEGORIES_TABLENAME & _
         " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
         " ORDER BY CATEGORY_NAME ASC"

   'Loading Category
   Set FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY = New ADODB.Recordset
   FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   'Requery The Recordset
   FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.Requery

   If FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.RecordCount < 1 Then
      ObjForm.Enabled = True
      LOAD_frmMAIN_TREEVIEW = True
      Exit Function
   End If

   'Loading The Categories To The Treeview
   Do While Not FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.EOF

      'Store The Category Name To tmpString
      TmpString = ProperCase(FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.Fields("CATEGORY_NAME").Value)

      'Add The Category Name To The Combobox(1)  - Category/Relation
      ObjForm.ComboBox(1).AddItem TmpString

      'Add the Category/Relation Nodes
      Set TmpNode = TView.Nodes.Add("ROOT", tvwChild, "cat_" & TmpString, TmpString, "users", "users")

      'Store Some Information In The Node's Tag - I will use this info. later
      TView.Nodes("cat_" & TmpString).Tag = "Category"

      'Move To The Next Record
      FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY.MoveNext
      DoEvents
   Loop


   'This Is Used To Set The Combobox to The Fist element
   If ObjForm.ComboBox(1).ListCount > 0 Then
      ObjForm.ComboBox(1).ListIndex = 0
   End If

   'Close And Clean Uo The CATEGORY_RS
   Set FAMILY.FORM_CATEGORIES.MyUser.USER_CATEGORY = Nothing

   Debug.Print "Load_frmMain_TreeView : Loading CONTACTS RECORDSET"
   tmpSQL = ""
   tmpSQL = "SELECT FirstName,LastName,CATEGORY_NAME,Sex FROM " & CONTACTS_TABLENAME & _
         " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
         " ORDER BY FirstName ASC, LastName ASC"

   'Loading the contacts recordset
   Set FAMILY.FORM_MAIN.MyUser.USER_CONTACTS = New ADODB.Recordset
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic
   'REQUERY The CONTACTS
   FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Requery

   If FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.RecordCount > 0 Then
      'Loading The Categories To The Treeview
      Do While Not FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.EOF
         'Store The Category Name To tmpString
         TmpString = FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("CATEGORY_NAME").Value

         'Store The First Name and LastName Of The Person
         'ex : John_Black  - the "_" will be used later to extract The First and Last Name Seperate
         tmpNameString = ProperCase(FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("FirstName").Value) & "_" & _
               ProperCase(FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("LastName").Value)

         If FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.Fields("Sex") = "Male" Then
            'Add Each Person To It's Category/Relation Nodes
            Set TmpNode = TView.Nodes.Add("cat_" & TmpString, tvwChild, "cat_" & TmpString & "_" & tmpNameString, tmpNameString, "person1", "person1")
            TView.Nodes("cat_" & TmpString & "_" & tmpNameString).Tag = "Child"
         Else
            'Add Each Person To It's Category/Relation Nodes
            Set TmpNode = TView.Nodes.Add("cat_" & TmpString, tvwChild, "cat_" & TmpString & "_" & tmpNameString, tmpNameString, "person2", "person2")
            TView.Nodes("cat_" & TmpString & "_" & tmpNameString).Tag = "Child"
         End If


         'Move To The NEXT Record
         FAMILY.FORM_MAIN.MyUser.USER_CONTACTS.MoveNext
      Loop
      DoEvents
   End If

   'Close And Clean Up CONTACTS_RS
   Set FAMILY.FORM_MAIN.MyUser.USER_CONTACTS = Nothing

   'This Is Used To Expand A Parent or ROOT Node
   If Expand_This_Node = "ROOT" Then
      TView.Nodes("ROOT").Expanded = True
   Else
      TView.Nodes("ROOT").Expanded = True
      TView.Nodes("cat_" & Expand_This_Node).Expanded = True
   End If


   'Set This Function = To True
   LOAD_frmMAIN_TREEVIEW = True
   'Enable the form
   ObjForm.Enabled = True

Load_frmMain_TreeView_error:
   If Err.Number <> 0 Then
      'Set This Function = To False
      LOAD_frmMAIN_TREEVIEW = False
      MsgBox "Error Loading Treeview : " & Err.Description & vbCrLf & _
            "Error # : " & Str$(Err.Number) & ".", vbCritical + vbOKOnly
      ObjForm.Enabled = True
      Err.Clear
   End If
End Function
'================================================================






'================================================================
'================================================================
Public Function LOAD_USERS_TREEVIEW(ByVal ObjForm As Form, _
                             ByVal TView As TreeView, _
                             ByVal IMG_List As ImageList, _
                             Optional ByVal Expand_This_Node As String = "ROOT" _
                             ) As Boolean

   Dim TmpNode        As Node
   Dim TmpString      As String
   Dim tmpNameString  As String
   Dim tmpSQL         As String

   On Error GoTo Load_USERS_TreeView_error

   'Clear the treeview and node
   Set TmpNode = Nothing
   TView.Nodes.Clear

   'Set The TreeView Image List
   Set TView.ImageList = IMG_List

   'This is Used to Add The "ROOT" Node
   Set TmpNode = TView.Nodes.Add(, , "ROOT", "People", "closed_book", "closed_book")
   'Store Some Information In The Node's Tag - I will use this info. later
   TView.Nodes("ROOT").Tag = "ROOT"
   Debug.Print "Load Treeview Data"

   Set TmpNode = TView.Nodes.Add("ROOT", tvwChild, "Administrator", "Administrator", "people1", "people1")
   TView.Nodes("Administrator").Tag = "Parent"
   Set TmpNode = TView.Nodes.Add("ROOT", tvwChild, "User", "User", "people1", "people1")
   TView.Nodes("User").Tag = "Parent"

   Debug.Print "Load_Users_TreeView : Loading USERS RECORDSET"
   tmpSQL = ""
   tmpSQL = "SELECT USER_LOGIN_NAME, USER_ACCESS_LEVEL FROM " & USER_TABLENAME & _
         " ORDER BY USER_LOGIN_NAME ASC"


   'Loading User
   Set FAMILY.FORM_USERS.MyUser.USER_PROFILE = New ADODB.Recordset
   'FAMILY.FORM_USERS.MyUser.USER_PROFILE.CursorLocation = adUseServer

   FAMILY.FORM_USERS.MyUser.USER_PROFILE.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   'Requery The Recordset
   FAMILY.FORM_USERS.MyUser.USER_PROFILE.Requery

   'MsgBox FAMILY.FORM_USERS.MyUser.USER_PROFILE.RecordCount

   If FAMILY.FORM_USERS.MyUser.USER_PROFILE.RecordCount < 1 Then
      Exit Function
   End If

   'Loading The Categories To The Treeview
   Do While Not FAMILY.FORM_USERS.MyUser.USER_PROFILE.EOF

      'Store The Accesslevel To tmpString
      TmpString = FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_ACCESS_LEVEL").Value
      'Store The User Name
      tmpNameString = FAMILY.FORM_USERS.MyUser.USER_PROFILE.Fields("USER_LOGIN_NAME").Value


      If TmpString = "Administrator" Then
         'Add the Nodes
         Set TmpNode = TView.Nodes.Add(TmpString, tvwChild, TmpString & "_" & tmpNameString, tmpNameString, "person1", "person1")
      Else
         'Add the Nodes
         Set TmpNode = TView.Nodes.Add(TmpString, tvwChild, TmpString & "_" & tmpNameString, tmpNameString, "person1", "person1")
      End If

      'Store Some Information In The Node's Tag - I will use this info. later
      TView.Nodes(TmpString & "_" & tmpNameString).Tag = "Child"
      'Set the Node Color
      TView.Nodes(TmpString & "_" & tmpNameString).ForeColor = vbBlack
      TView.Nodes(TmpString & "_" & tmpNameString).BackColor = vbWhite

      'Move To The Next Record
      FAMILY.FORM_USERS.MyUser.USER_PROFILE.MoveNext
   Loop

   'Close And Clean Uo The CATEGORY_RS
   FAMILY.FORM_USERS.MyUser.USER_PROFILE.Close
   Set FAMILY.FORM_USERS.MyUser.USER_PROFILE = Nothing

   'This Is Used To Expand A Parent or ROOT Node
   If Expand_This_Node = "ROOT" Then
      TView.Nodes("ROOT").Expanded = True
   Else
      TView.Nodes("ROOT").Expanded = True
      TView.Nodes("cat_" & Expand_This_Node).Expanded = True
   End If

   If CURRENT_USER.ACCESS_LEVEL = Administrator Then
      TView.Enabled = True
   Else
      TView.Enabled = False
   End If

   'Set This Function = To True
   LOAD_USERS_TREEVIEW = True

   DoEvents


Load_USERS_TreeView_error:
   If Err.Number <> 0 Then
      'Set This Function = To False
      LOAD_USERS_TREEVIEW = False
      MsgBox "Error Loading Treeview : " & Err.Description & vbCrLf & _
            "Error # : " & Str$(Err.Number) & ".", vbCritical + vbOKOnly
      Err.Clear
   End If
End Function
'================================================================




'================================================================
'Used to Change The Background Color Of The Treeview
'Thanks To  dennis wrenn
'================================================================
Public Function SetTVBackgroundColor(ByVal ColorValue As Long, ByVal TreeViewControl As TreeView) As Boolean
   On Error GoTo error_SetTVBackgroundColor
   Dim lngLineStyle As Long
   Dim lngRet As Long

   With TreeViewControl
      lngRet = SendMessage(.hWnd, TVM_SETBKCOLOR, 0, ByVal ColorValue)
      lngLineStyle = GetWindowLong(.hWnd, GWL_STYLE)
      If (lngLineStyle And TVS_HASLINES) = True Then
         lngRet = SetWindowLong(.hWnd, GWL_STYLE, lngLineStyle Xor TVS_HASLINES)
         lngRet = SetWindowLong(.hWnd, GWL_STYLE, lngLineStyle)
      End If

      SetTVBackgroundColor = True
   End With

   Exit Function

error_SetTVBackgroundColor:
   SetTVBackgroundColor = False

   'SetTVBackgroundColor RGB(0, 200, 255), TreeView1
End Function
'================================================================



'================================================================
'                       LOAD_LINKS_TREEVIEW
'================================================================
Public Function LOAD_LINKS_TREEVIEW(ByVal ObjForm As Form, _
                             ByVal TView As TreeView, _
                             ByVal IMG_List As ImageList, _
                             Optional ByVal Expand_This_Node As String = "ROOT" _
                             ) As Boolean


   Dim TmpNode        As Node
   Dim TmpString      As String
   Dim tmpNameString  As String
   Dim tmpSQL         As String

   On Error GoTo LOAD_LINKS_TREEVIEW_ERROR

   LOAD_LINKS_TREEVIEW = False

   'Disable The Form until the treeview has been loaded
   ObjForm.Enabled = False

   'Clear the treeview and node
   Set TmpNode = Nothing
   TView.Nodes.Clear


   'Set The TreeView Image List
   Set TView.ImageList = IMG_List

   'This is Used to Add The "ROOT" Node
   Set TmpNode = TView.Nodes.Add(, , "ROOT", CURRENT_USER.LOGIN_NAME, "ie1", "ie1")
   'Store Some Information In The Node's Tag - I will use this info. later
   TView.Nodes("ROOT").Tag = "ROOT"

   tmpSQL = ""
   tmpSQL = "SELECT LINK FROM " & LINKS_TABLENAME & _
         " WHERE USER_NAME = '" & CURRENT_USER.LOGIN_NAME & "'" & _
         " ORDER BY LINK ASC"

   'Loading Category
   Set FAMILY.FORM_LINKS.MyUser.USER_LINKS = New ADODB.Recordset
   FAMILY.FORM_LINKS.MyUser.USER_LINKS.Open tmpSQL, PUBLIC_DATABASE.CONNECTION, adOpenKeyset, adLockOptimistic

   'Requery The Recordset
   FAMILY.FORM_LINKS.MyUser.USER_LINKS.Requery
   TmpString = ""

   If FAMILY.FORM_LINKS.MyUser.USER_LINKS.RecordCount > 0 Then

      Set TmpNode = Nothing
      'Loading The Linkss To The Treeview
      Do While Not FAMILY.FORM_LINKS.MyUser.USER_LINKS.EOF
         'Store The Category Name To tmpString
         TmpString = FAMILY.FORM_LINKS.MyUser.USER_LINKS.Fields("LINK").Value

         'This is Used to Add The "ROOT" Node
         Set TmpNode = TView.Nodes.Add("ROOT", tvwChild, TmpString, TmpString, "ie2", "ie2")
         'Store Some Information In The Node's Tag - I will use this info. later
         TView.Nodes(TmpString).Tag = "LINK"

         FAMILY.FORM_LINKS.MyUser.USER_LINKS.MoveNext
      Loop

      DoEvents

      TView.Nodes("ROOT").Expanded = True
      Set FAMILY.FORM_LINKS.MyUser.USER_LINKS = Nothing
      LOAD_LINKS_TREEVIEW = True
      ObjForm.Enabled = True
      Exit Function

   Else

      'Enable the Form
      ObjForm.Enabled = True
      LOAD_LINKS_TREEVIEW = True
      Exit Function

   End If


LOAD_LINKS_TREEVIEW_ERROR:
   If Err.Number <> 0 Then
      MsgBox "LOAD_LINKS_TREEVIEW_ERROR : " & Err.Description & " : " & Str$(Err.Number), vbCritical + vbOKOnly
      ObjForm.Enabled = True
      Err.Clear
   End If
End Function
'================================================================



