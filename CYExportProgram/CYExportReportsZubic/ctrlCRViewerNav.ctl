VERSION 5.00
Begin VB.UserControl ctrlCRViewerNav 
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9345
   KeyPreview      =   -1  'True
   ScaleHeight     =   810
   ScaleWidth      =   9345
   ToolboxBitmap   =   "ctrlCRViewerNav.ctx":0000
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10
      Picture         =   "ctrlCRViewerNav.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " View / Refresh "
      Top             =   10
      Width           =   2280
   End
   Begin VB.ComboBox cboPageSize 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "ctrlCRViewerNav.ctx":045C
      Left            =   6600
      List            =   "ctrlCRViewerNav.ctx":0475
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   " Zoom "
      Top             =   120
      Width           =   2085
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "&Last"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   5160
      Picture         =   "ctrlCRViewerNav.ctx":04A4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Last Page"
      Top             =   10
      Width           =   855
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   4320
      Picture         =   "ctrlCRViewerNav.ctx":05EE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Next Page"
      Top             =   10
      Width           =   855
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   3480
      Picture         =   "ctrlCRViewerNav.ctx":0738
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Previous Page"
      Top             =   10
      Width           =   855
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "&First"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2640
      Picture         =   "ctrlCRViewerNav.ctx":0882
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " First Page "
      Top             =   10
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Picture         =   "ctrlCRViewerNav.ctx":09CC
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   " Print "
      Top             =   10
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   -90
      Width           =   9255
   End
End
Attribute VB_Name = "ctrlCRViewerNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal Y As Long)
Private Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As Rect)
Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Type Rect
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
'Property Variables:
Dim m_CRViewerControl As CRViewer
Dim m_ReportObject As Object
'Dim m_CRViewerControl As CRViewer
Event PreviewReport()
Event PrintReport()

Private Sub cboPageSize_Change()
    Call lzResizePage
    Call cmdRefresh_Click
End Sub

Private Sub cmdPage_Click(Index As Integer)
Select Case Index
    Case 0
        m_CRViewerControl.ShowFirstPage
        cmdPage(0).SetFocus
        Call PositionCursor
    Case 1
        m_CRViewerControl.ShowPreviousPage
       cmdPage(1).SetFocus
       Call PositionCursor
    Case 2
        m_CRViewerControl.ShowNextPage
       cmdPage(2).SetFocus
       Call PositionCursor
    Case 3
        m_CRViewerControl.ShowLastPage
       cmdPage(3).SetFocus
       Call PositionCursor
End Select
End Sub
Private Sub cmdPrint_Click()
    m_CRViewerControl.PrintReport
    RaiseEvent PrintReport
    cmdPrint.SetFocus
End Sub
Public Sub cmdRefresh_Click()
    Call lzResizePage
    Call lzCursor2Viewer
    RaiseEvent PreviewReport
    cmdRefresh.SetFocus
End Sub
Private Sub lzResizePage()
Dim i, sz As Integer
    i = cboPageSize.ListIndex
    If (i < 5) Then
        sz = (left(cboPageSize.List(i), 3))
    ElseIf (i = 5) Then
        sz = CInt(1)
    ElseIf (i = 6) Then
        sz = CInt(2)
    End If
    On Error GoTo err_pagesize
tagRepeat:
    m_CRViewerControl.Zoom (sz)
    Exit Sub
err_pagesize:
    DoEvents
    GoTo tagRepeat
End Sub
Private Sub lzCursor2Viewer()
    Dim Rect As Rect
    Dim X As Integer
    Dim Y As Integer
    
    X = 900
    Y = 500
    
    SetCursorPos X, Y
    
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        SendKeys "{Tab}", True
    Case 38
        SendKeys "+{Tab}", True
End Select
End Sub

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=15,2,0,0
'Public Property Get CRViewerControl() As CRViewer
''    If Ambient.UserMode Then Err.Raise 393
'    Set CRViewerControl = m_CRViewerControl
'End Property
'
'Public Property Set CRViewerControl(ByVal New_CRViewerControl As CRViewer)
'    Set m_CRViewerControl = New_CRViewerControl
'    PropertyChanged "CRViewerControl"
'End Property
Private Sub UserControl_Initialize()
    cboPageSize.ListIndex = 3
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    Set m_CRViewerControl = PropBag.ReadProperty("CRViewerControl", Nothing)
    Set m_CRViewerControl = PropBag.ReadProperty("CRViewerControl", Nothing)
    Set m_ReportObject = PropBag.ReadProperty("ReportObject", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("CRViewerControl", m_CRViewerControl, Nothing)
    Call PropBag.WriteProperty("CRViewerControl", m_CRViewerControl, Nothing)
    Call PropBag.WriteProperty("ReportObject", m_ReportObject, Nothing)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,2,0,0
Public Property Get CRViewerControl() As CRViewer
    If Ambient.UserMode Then Err.Raise 393
    Set CRViewerControl = m_CRViewerControl
End Property

Public Property Set CRViewerControl(ByVal New_CRViewerControl As CRViewer)
    Set m_CRViewerControl = New_CRViewerControl
    PropertyChanged "CRViewerControl"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get ReportObject() As Object
    Set ReportObject = m_ReportObject
End Property

Public Property Set ReportObject(ByVal New_ReportObject As Object)
    Set m_ReportObject = New_ReportObject
    PropertyChanged "ReportObject"
End Property
Public Sub PositionCursor()
    lzCursor2Viewer
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

