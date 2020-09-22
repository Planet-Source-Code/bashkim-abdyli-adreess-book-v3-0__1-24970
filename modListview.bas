Attribute VB_Name = "modListview"

Option Explicit

'Used for enhance the listview control
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
      (ByVal hWnd As Long, _
      ByVal MSG As Long, _
      ByVal wParam As Long, _
      lParam As Any) As Long


Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

'USED FOR ENTIRE ROW SELECT
Public Const LVS_EX_FULLROWSELECT = &H20
'--end block--'


'======================================================================
'This is Used To Setup The ListView Control For frmReminder
'======================================================================
Public Function Load_Reminders_ColumnHeaders(ByVal List_View As ListView) As Boolean
   Dim ListviewHeader    As ColumnHeader   ' Used For Listview Header
   Dim lvListItems As ListItem   ' Used For Listview Items

   'Clear the Listview Control
   List_View.ListItems.Clear
   'Clear The ColumnHeaders
   List_View.ColumnHeaders.Clear

   Set ListviewHeader = Nothing

   'Start Adding The Listview Headers
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C1", "Subject", 4500, lvwColumnLeft)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C2", "Date", 2200, lvwColumnCenter)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C3", "Time", 1200, lvwColumnCenter)
End Function
'======================================================================



'======================================================================
'This is Used To Setup The ListView Control For frmSearch
'======================================================================
Public Function Load_Search_Lisview(ByVal List_View As ListView) As Boolean
   Dim ListviewHeader    As ColumnHeader   ' Used For Listview Header
   Dim lvListItems As ListItem   ' Used For Listview Items

   'Clear the Listview Control
   List_View.ListItems.Clear

   Set ListviewHeader = Nothing

   'Start Adding The Listview Headers
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C1", "First Name", 2500, lvwColumnLeft)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C2", "Last Name", 2500, lvwColumnLeft)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C3", "Sex", 800, lvwColumnLeft)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C4", "Telephone #", 1800, lvwColumnLeft)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C5", "Address", 2500, lvwColumnLeft)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C6", "City, State", 2500, lvwColumnLeft)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C7", "Zip Code", 1000, lvwColumnLeft)
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C8", "Em@il Address", 2500, lvwColumnLeft, "mail")
   Set ListviewHeader = List_View.ColumnHeaders.Add(, "C9", "Categories", 1500, lvwColumnLeft)
   List_View.View = lvwReport

   Load_Search_Lisview = True
End Function
'======================================================================

