Attribute VB_Name = "modGraphics"
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

'---------------------------------------------------------------
'Used To Remove And Append Menu
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
'---------------------------------------------------------------



'================================================================
' TYPE STRUCTURES
Private Type tpeTextProperties
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Private Type tpeRectangle
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' CONSTANTS
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4

' API DECLARATIONS
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As tpeRectangle, ByVal un As Long, lpDrawTextParams As tpeTextProperties) As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
Private Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As tpeRectangle) As Long

Public strCharSpace As Integer
'================================================================






'===================================================================
Public Sub RemoveMenus(ByVal ObjForm As Form)
   Dim hMenu As Long
   ' Get the form's system menu handle.
   hMenu = GetSystemMenu(ObjForm.hwnd, False)
   DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub
'===================================================================


'
' Public Type RECT
'   Left As Long
'   Top As Long
'   Right As Long
'   Bottom As Long
'End Type
'
'Public Const BDR_RAISEDOUTER = &H1
'Public Const BDR_SUNKENOUTER = &H2
'Public Const BDR_RAISEDINNER = &H4
'Public Const BDR_SUNKENINNER = &H8
'
'Public Const BDR_OUTER = &H3
'Public Const BDR_INNER = &HC
'Public Const BDR_RAISED = &H5
'Public Const BDR_SUNKEN = &HA
'
'Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
'Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
'Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
'Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
'
'Public Const BF_LEFT = &H1
'Public Const BF_TOP = &H2
'Public Const BF_RIGHT = &H4
'Public Const BF_BOTTOM = &H8
'
'Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
'Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
'Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
'Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
'Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'
'Public Const BF_DIAGONAL = &H10
'
''For diagonal lines, the BF_RECT flags specify the end point of
''the vector bounded by the rectangle parameter.
'Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP _
'      Or BF_RIGHT)
'Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP _
'      Or BF_LEFT)
'Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM _
'      Or BF_LEFT)
'Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM _
'      Or BF_RIGHT)
'
'Public Const BF_MIDDLE = &H800   ' Fill in the middle.
'Public Const BF_SOFT = &H1000   ' Use for softer buttons.
'Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
'Public Const BF_FLAT = &H4000   ' For flat rather than 3-D borders.
'Public Const BF_MONO = &H8000   ' For monochrome borders.
'
'Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
'      qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean


''---------------------------------------------------------------
''Used To Tile Image
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Const SRCCOPY = &HCC0020
''---------------------------------------------------------------
'
'
'
''================================================================
''Used to Tile an Image on a form
''================================================================
'Public Sub TileBitmap(Target As Form, Source As PictureBox)
'   Dim BackupInformation_ScaleMode As Byte
'   Dim BackupInformation_ScaleMode2 As Byte
'   Dim YDraw As Long
'   Dim XDraw As Long
'
'   BackupInformation_ScaleMode = Target.ScaleMode
'   BackupInformation_ScaleMode2 = Source.ScaleMode
'   Source.ScaleMode = 3
'   Target.ScaleMode = 3
'   Target.Cls
'   Target.AutoRedraw = True
'
'   For YDraw = 0 To Target.Height Step Source.ScaleHeight
'      For XDraw = 0 To Target.ScaleWidth Step Source.ScaleWidth
'         BitBlt Target.hdc, XDraw, YDraw, Source.ScaleWidth, Source.ScaleHeight, Source.hdc, 0, 0, SRCCOPY
'      Next XDraw
'   Next YDraw
'   Target.ScaleMode = BackupInformation_ScaleMode
'   Source.ScaleMode = BackupInformation_ScaleMode2
'End Sub
''================================================================






