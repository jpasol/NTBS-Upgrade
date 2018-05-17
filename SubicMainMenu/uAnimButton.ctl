VERSION 5.00
Begin VB.UserControl uAnimButton 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   120
   End
End
Attribute VB_Name = "uAnimButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' Name:     uAnimButton
' Author:   Steve McMahon
' Date:     24 January 1999
'
' A very simple animated button.  When the mouse moves over,
' it animates a picture strip.  When pressed or the mouse is
' not over, it shows the default image.
' ===========================================================================

Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
    Private Const BF_LEFT = &H1
    Private Const BF_BOTTOM = &H8
    Private Const BF_RIGHT = &H4
    Private Const BF_TOP = &H2
    Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    Private Const BDR_INNER = &HC
    Private Const BDR_OUTER = &H3
    Private Const BDR_RAISED = &H5
    Private Const BDR_RAISEDINNER = &H4
    Private Const BDR_RAISEDOUTER = &H1
    Private Const BDR_SUNKEN = &HA
    Private Const BDR_SUNKENINNER = &H8
    Private Const BDR_SUNKENOUTER = &H2
    Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Private m_pic As StdPicture
Private m_lXCells As Long, m_lYCells As Long
Private m_lCellWidth As Long, m_lCellHeight As Long
Private m_lWidth As Long, m_lHeight As Long
Private m_lCell As Long
Private m_lCellCount As Long
Private m_lDefaultCell As Long
Private m_bPressed As Boolean
Private m_bInterlock As Boolean
Private m_lCellSteps() As Long
Private m_lStep As Long

Public Event Click()

Public Property Get Interval() As Long
   Interval = tmrAnim.Interval
End Property

Public Property Let Interval(ByVal lInterval As Long)
   tmrAnim.Interval = lInterval
   PropertyChanged "Interval"
End Property

Public Property Get CellSteps(ByVal lCell As Long) As Long
   CellSteps = m_lCellSteps(lCell)
End Property

Public Property Let CellSteps(ByVal lCell As Long, ByVal lSteps As Long)
   m_lCellSteps(lCell) = lSteps
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal oColor As OLE_COLOR)
   UserControl.BackColor = oColor
   PropertyChanged "BackColor"
End Property

Public Property Get CellCount() As Long
   CellCount = m_lCellCount
End Property

Public Property Get DefaultCell() As Long
   DefaultCell = m_lDefaultCell
End Property

Public Property Let DefaultCell(ByVal lCell As Long)
   m_lDefaultCell = lCell
   m_lCell = lCell
   Draw
   PropertyChanged "DefaultCell"
End Property

Public Property Get XCells() As Long
   XCells = m_lXCells
End Property

Public Property Let XCells(ByVal lX As Long)
   m_lXCells = lX
   If (lX <> 0) Then
      m_lCellWidth = m_lWidth \ m_lXCells
      m_lCellCount = m_lXCells * m_lYCells
      If (m_lCellCount > 0) Then
         ReDim Preserve m_lCellSteps(0 To m_lCellCount - 1) As Long
      End If
      PropertyChanged "XCells"
   End If
End Property

Public Property Get YCells() As Long
   YCells = m_lYCells
End Property

Public Property Let YCells(ByVal lY As Long)
   m_lYCells = lY
   If (lY <> 0) Then
      m_lCellHeight = m_lHeight \ m_lYCells
      If (m_lCellCount > 0) Then
         ReDim Preserve m_lCellSteps(0 To m_lCellCount - 1) As Long
      End If
      m_lCellCount = m_lXCells * m_lYCells
      PropertyChanged "YCells"
   End If
End Property

Public Property Set Picture(ByRef s As StdPicture)
   Set m_pic = s
   If Not (m_pic Is Nothing) Then
      m_lWidth = UserControl.ScaleX(s.Width, vbHimetric, vbPixels)
      m_lHeight = UserControl.ScaleY(s.Height, vbHimetric, vbPixels)
      XCells = m_lXCells
      YCells = m_lYCells
      Draw
   End If
   PropertyChanged "Picture"
End Property

Public Property Get Cell() As Long
   Cell = m_lCell
   PropertyChanged "Cell"
End Property

Public Sub Step()
   m_lStep = m_lStep + 1
   If (m_lStep > m_lCellSteps(m_lCell)) Then
      m_lStep = 0
      m_lCell = m_lCell + 1
      If (m_lCell >= m_lCellCount) Then
         m_lCell = 0
      End If
      Draw
   End If
End Sub

Public Sub Draw()
Dim tR As RECT
Dim lEdge As Long
Dim lLeft As Long, lTop As Long
Dim lWidth As Long, lHeight As Long
Dim lSrcLeft As Long, lSrcTOp As Long
Static bPressed As Boolean

   If (m_bPressed <> bPressed) Then
      UserControl.Cls
   End If

   GetClientRect UserControl.hwnd, tR
   If (m_bPressed) Then
      lEdge = EDGE_BUMP
   Else
      lEdge = EDGE_RAISED
   End If
   DrawEdge UserControl.hDC, tR, lEdge, BF_RECT
   InflateRect tR, -1, -1
   
   lLeft = tR.left + (tR.Right - tR.left - m_lCellWidth) \ 2 + 1
   lTop = tR.tOp + (tR.Bottom - tR.tOp - m_lCellHeight) \ 2 + 1
   If (lLeft < tR.left) Then
      lLeft = tR.left
      lWidth = tR.Right - tR.left
   Else
      lWidth = m_lCellWidth
   End If
   If (lTop < tR.tOp) Then
      lTop = tR.tOp
      lHeight = tR.Bottom - tR.tOp
   Else
      lHeight = m_lCellHeight - 1
   End If
   
   If Not (m_pic Is Nothing) Then
      lSrcLeft = (m_lCell Mod m_lXCells) * m_lCellWidth
      lSrcTOp = (m_lCell \ m_lXCells) * m_lCellHeight
      If (m_bPressed) Then
         lLeft = lLeft + 1
         lTop = lTop + 1
      End If
      
      UserControl.PaintPicture m_pic, lLeft, lTop, lWidth, lHeight, lSrcLeft, lSrcTOp, lWidth, lHeight
   End If
   
   bPressed = m_bPressed
   
End Sub

Private Sub tmrAnim_Timer()
Dim tP As POINTAPI
Dim tR As RECT
   GetCursorPos tP
   GetWindowRect UserControl.hwnd, tR
   If (PtInRect(tR, tP.x, tP.y) = 0) Then
      tmrAnim.Enabled = False
      m_lCell = m_lDefaultCell
      If Not (m_bInterlock) Then
         Draw
      End If
      m_bInterlock = False
   Else
      If Not (m_bInterlock) Then
         Step
         Draw
      End If
   End If
End Sub

Private Sub UserControl_Initialize()
   m_lCellCount = 1
   m_lXCells = 1
   m_lYCells = 1
   m_lDefaultCell = 0
   m_lCell = 0
   ReDim m_lCellSteps(0 To 0) As Long
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeySpace) Then
      UserControl_MouseDown 1, 0, 10, 10
   End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeySpace) Then
      UserControl_MouseUp 1, 0, 10, 10
   End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   m_bPressed = True
   tmrAnim.Enabled = False
   m_lCell = m_lDefaultCell
   Draw
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If (Button = 0) Then
      If Not (tmrAnim.Enabled) Then
         tmrAnim.Enabled = True
      End If
   Else
      If (x > 0) And (x < UserControl.ScaleWidth) And (y > 0) And (y < UserControl.ScaleHeight) Then
         If Not (m_bPressed) Then
            m_bPressed = True
            Draw
         End If
      Else
         If (m_bPressed) Then
            m_bPressed = False
            Draw
         End If
      End If
   End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tP As POINTAPI
Dim tR As RECT
   m_bInterlock = True
   tmrAnim.Enabled = True
   m_bPressed = False
   If (x > 0) And (x < UserControl.ScaleWidth) And (y > 0) And (y < UserControl.ScaleHeight) Then
      RaiseEvent Click
   End If
   Draw
End Sub

Private Sub UserControl_Paint()
   Draw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Set Picture = PropBag.ReadProperty("Picture", Nothing)
   XCells = PropBag.ReadProperty("XCells", 1)
   YCells = PropBag.ReadProperty("YCells", 1)
   DefaultCell = PropBag.ReadProperty("DefaultCell", 0)
   Interval = PropBag.ReadProperty("Interval", 200)
End Sub

Private Sub UserControl_Terminate()
   tmrAnim.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "Picture", m_pic, Nothing
   PropBag.WriteProperty "XCells", XCells, 1
   PropBag.WriteProperty "YCells", YCells, 1
   PropBag.WriteProperty "DefaultCell", DefaultCell, 0
   PropBag.WriteProperty "Interval", Interval, 200
End Sub
