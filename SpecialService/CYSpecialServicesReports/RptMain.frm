VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRptMain 
   Caption         =   " CY Special Services Reports"
   ClientHeight    =   3540
   ClientLeft      =   75
   ClientTop       =   645
   ClientWidth     =   13380
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RptMain.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   13380
   Begin VB.Frame fraParameter 
      Caption         =   " Parameters "
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1515
      Left            =   7125
      TabIndex        =   18
      Top             =   75
      Width           =   6165
      Begin MSMask.MaskEdBox txtFromDate 
         Height          =   390
         Left            =   1650
         TabIndex        =   2
         ToolTipText     =   " Start of date range "
         Top             =   600
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-##-##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtToDate 
         Height          =   390
         Left            =   1650
         TabIndex        =   3
         ToolTipText     =   " End of date range "
         Top             =   1050
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd"
         Mask            =   "####-##-##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtOther 
         Height          =   390
         Left            =   1650
         TabIndex        =   1
         ToolTipText     =   " User ID "
         Top             =   150
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFromTime 
         Height          =   390
         Left            =   5175
         TabIndex        =   4
         ToolTipText     =   " Start of time range "
         Top             =   600
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtToTime 
         Height          =   390
         Left            =   5175
         TabIndex        =   5
         ToolTipText     =   " End of date range "
         Top             =   1050
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "To Time"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         TabIndex        =   23
         Top             =   1125
         Width           =   1365
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "From Time"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         TabIndex        =   22
         Top             =   675
         Width           =   1365
      End
      Begin VB.Label lblOther 
         Alignment       =   1  'Right Justify
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   21
         Top             =   225
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   20
         Top             =   1125
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   19
         Top             =   675
         Width           =   1515
      End
   End
   Begin VB.Frame fraReport 
      Caption         =   " Report "
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   75
      TabIndex        =   17
      Top             =   75
      Width           =   6990
      Begin VB.ComboBox cboReport 
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "RptMain.frx":27A2
         Left            =   150
         List            =   "RptMain.frx":27A4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   " Select report from here "
         Top             =   225
         Width           =   6690
      End
   End
   Begin VB.Frame fraControl 
      Caption         =   " Control "
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   765
      Left            =   75
      TabIndex        =   16
      Top             =   825
      Width           =   6990
      Begin VB.CommandButton cmdPrint 
         Height          =   390
         Left            =   750
         Picture         =   "RptMain.frx":27A6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Print "
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   390
         Left            =   150
         Picture         =   "RptMain.frx":28F0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " View / Refresh "
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdPage 
         Height          =   390
         Index           =   1
         Left            =   1800
         Picture         =   "RptMain.frx":2A3A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   " Previous Page"
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdPage 
         Height          =   390
         Index           =   0
         Left            =   1350
         Picture         =   "RptMain.frx":2B84
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " First Page "
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdPage 
         Height          =   390
         Index           =   2
         Left            =   2250
         Picture         =   "RptMain.frx":2CCE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Next Page"
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdPage 
         Height          =   390
         Index           =   3
         Left            =   2700
         Picture         =   "RptMain.frx":2E18
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Last Page"
         Top             =   225
         Width           =   390
      End
      Begin VB.ComboBox cboPageSize 
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "RptMain.frx":2F62
         Left            =   4575
         List            =   "RptMain.frx":2F7B
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   " Zoom "
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton cmdGo2Page 
         Height          =   390
         Left            =   3975
         Picture         =   "RptMain.frx":2FAA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " Jump To Page "
         Top             =   225
         Width           =   390
      End
      Begin MSMask.MaskEdBox txtPageNo 
         Height          =   390
         Left            =   3300
         TabIndex        =   12
         ToolTipText     =   " Page No "
         Top             =   225
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###"
         PromptChar      =   " "
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   1785
      Left            =   75
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1650
      Width           =   13230
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Menu"
      Begin VB.Menu mnuReportChoose 
         Caption         =   "&Choose Report"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuReportParam 
         Caption         =   "Para&meters"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuReportRefresh 
         Caption         =   "&Refresh / View"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuReportPrint 
         Caption         =   "&Print"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuReportF1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportFirst 
         Caption         =   "&First Page"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuReportPrevious 
         Caption         =   "Pre&vious Page"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuReportNext 
         Caption         =   "&Next Page"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuReportLast 
         Caption         =   "&Last Page"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuReportGoTo 
         Caption         =   "&Go To Page No."
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuReportZoom 
         Caption         =   "&Zoom"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuReportF2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmRptMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCursorPos& Lib "user32" (ByVal x As Long, ByVal y As Long)
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

Dim repDailyColxn As rptDailyColxn
Dim repSOList As rptSOList
Dim repSOInquiry As rptSOInquiry
Dim repStoColxn As rptStoColxn
Dim repMonthly As rptMonthly

Public Sub cboPageSize_Click()
    Call lzResizePage
End Sub

Public Sub cboPageSize_GotFocus()
    cboPageSize.BackColor = vbInfoBackground
    SetMouseFocus cboPageSize
End Sub

Public Sub cboPageSize_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub cboPageSize_LostFocus()
    cboPageSize.BackColor = vbWindowBackground
End Sub

Public Sub cboReport_GotFocus()
    cboReport.BackColor = vbInfoBackground
    SetMouseFocus cboReport
End Sub

Public Sub cboReport_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case vbKeySpace
            SendKeys ("{F4}")
        Case vbKeyEscape
            Unload Me
        Case Else
    End Select
End Sub

Public Sub cboReport_LostFocus()
    cboReport.BackColor = vbWindowBackground
    Call lzSetParm
End Sub

Public Sub cmdGo2Page_Click()
    On Error GoTo err_page
    CRViewer1.ShowNthPage (CInt(txtPageNo.Text))
tagReturn:
    On Error GoTo err_wait
tagRepeat:
    txtPageNo.Text = CRViewer1.GetCurrentPageNumber
    Exit Sub
err_wait:
    DoEvents
    GoTo tagRepeat
    Exit Sub
err_page:
    On Error GoTo err_wait
    CRViewer1.ShowLastPage
    GoTo tagReturn
End Sub

Public Sub cmdGo2Page_GotFocus()
    SetMouseFocus cmdGo2Page
End Sub

Public Sub cmdGo2Page_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            CRViewer1.ShowNthPage (CInt(txtPageNo.Text))
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub cmdPage_Click(Index As Integer)
    On Error Resume Next
    With CRViewer1
        Select Case Index
            Case 0
                .ShowFirstPage
            Case 1
                .ShowPreviousPage
            Case 2
                .ShowNextPage
            Case 3
                .ShowLastPage
            Case Else
        End Select
        On Error GoTo err_wait
tagRepeat:
        txtPageNo.Text = .GetCurrentPageNumber
        SetMouseFocus cmdPage(Index)
    End With
    Exit Sub
err_wait:
    DoEvents
    GoTo tagRepeat
End Sub

Public Sub cmdPage_GotFocus(Index As Integer)
    SetMouseFocus cmdPage(Index)
End Sub

Public Sub cmdPage_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            cmdPage_Click (Index)
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub cmdPrint_Click()
    Call lzPrint
End Sub

Public Sub cmdPrint_GotFocus()
    SetMouseFocus cmdPrint
End Sub

Public Sub cmdPrint_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            Call lzPrint
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub cmdRefresh_Click()
    lzViewReport
End Sub

Public Sub cmdRefresh_GotFocus()
    SetMouseFocus cmdRefresh
End Sub

Public Sub cmdRefresh_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            lzViewReport
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub Combo1_GotFocus()
    SendKeys ("{F4}")
End Sub

Public Sub Form_Load()
    lzInitialize
End Sub

Public Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        CRViewer1.Height = ScaleHeight - 1800
        CRViewer1.Width = ScaleWidth - 150
    End If
End Sub

Public Sub Form_Unload(Cancel As Integer)
    If (MsgBox("Exit CY Special Services Reports ?", vbYesNo, "CYS Reports") = vbNo) Then
        Cancel = 1
    End If
End Sub

Public Sub mnuReportChoose_Click()
    cboReport.SetFocus
End Sub

Public Sub mnuReportExit_Click()
    Unload Me
End Sub

Public Sub mnuReportFirst_Click()
    Call cmdPage_Click(0)
End Sub

Public Sub mnuReportGoTo_Click()
    txtPageNo.SetFocus
End Sub

Public Sub mnuReportLast_Click()
    Call cmdPage_Click(3)
End Sub

Public Sub mnuReportNext_Click()
    Call cmdPage_Click(2)
End Sub

Public Sub mnuReportParam_Click()
    Call lzSetParm
End Sub

Public Sub mnuReportPrevious_Click()
    Call cmdPage_Click(1)
End Sub

Public Sub mnuReportPrint_Click()
    Call lzPrint
End Sub

Public Sub mnuReportRefresh_Click()
    lzViewReport
End Sub

Public Sub mnuReportZoom_Click()
    cboPageSize.SetFocus
End Sub

Public Sub txtFromDate_GotFocus()
    With txtFromDate
        .SelStart = 0
        .SelLength = .MaxLength
        .BackColor = vbInfoBackground
    End With
    SetMouseFocus txtFromDate
End Sub

Public Sub txtFromDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub txtFromDate_LostFocus()
    With txtFromDate
        .BackColor = vbWindowBackground
        .Text = right(Space(8) & .Text, 8)
    End With
End Sub

Public Sub txtFromTime_GotFocus()
    With txtFromTime
        .SelStart = 0
        .SelLength = .MaxLength
        .BackColor = vbInfoBackground
    End With
    SetMouseFocus txtFromTime
End Sub

Public Sub txtFromTime_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub txtFromTime_LostFocus()
    With txtFromTime
        .BackColor = vbWindowBackground
        .Text = right("0000" & .Text, 4)
    End With
End Sub

Public Sub txtPageNo_GotFocus()
    With txtPageNo
        .SelStart = 0
        .SelLength = .MaxLength
        .BackColor = vbInfoBackground
    End With
    SetMouseFocus txtPageNo
End Sub

Public Sub txtPageNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub txtPageNo_LostFocus()
    txtPageNo.BackColor = vbWindowBackground
End Sub

Public Sub txtToDate_GotFocus()
    With txtToDate
        .SelStart = 0
        .SelLength = .MaxLength
        .BackColor = vbInfoBackground
    End With
    SetMouseFocus txtToDate
End Sub

Public Sub txtToDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub txtToDate_LostFocus()
    With txtToDate
        .BackColor = vbWindowBackground
        .Text = right(Space(8) & .Text, 8)
    End With
End Sub

Public Sub txtToTime_GotFocus()
    With txtToTime
        .SelStart = 0
        .SelLength = .MaxLength
        .BackColor = vbInfoBackground
    End With
    SetMouseFocus txtToTime
End Sub

Public Sub txtToTime_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub txtToTime_LostFocus()
    With txtToTime
        .BackColor = vbWindowBackground
        .Text = right("0000" & .Text, 4)
    End With
End Sub

Public Sub txtOther_GotFocus()
    With txtOther
        .SelStart = 0
        .SelLength = .MaxLength
        .BackColor = vbInfoBackground
    End With
    SetMouseFocus txtOther
End Sub

Public Sub txtOther_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Public Sub txtOther_LostFocus()
    txtOther.BackColor = vbWindowBackground
End Sub

Public Sub lzResizePage()
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
    CRViewer1.Zoom (sz)
    Exit Sub
err_pagesize:
    DoEvents
    GoTo tagRepeat
End Sub

Public Sub lzEnable(ByRef c As Control)
    c.Enabled = True
End Sub

Public Sub lzDisable(ByRef c As Control)
    With c
        .Text = Space(.MaxLength): .Enabled = False
    End With
End Sub

Public Sub lzPrint()
    CRViewer1.PrintReport
End Sub

Public Sub SetMouseFocus(ByVal Obj As Object)
    Dim Rect As Rect

    'Get the bounding rectangle for window
    GetWindowRect Obj.hwnd, Rect

    SetCursorPos Rect.right - 5, _
                 Rect.bottom - ((Rect.bottom - Rect.top) / 2)
End Sub

Public Sub lzInitialize()
    With Me
        .top = 50: .left = 50: .Width = 15200: .Height = 11000
    End With
    
    txtOther.Text = gUserID: lblOther.Caption = "User"
    txtFromDate.Text = Format(Now, "yyyy-mm-dd")
    txtToDate.Text = Format(Now, "yyyy-mm-dd")
    txtFromTime.Text = Format("00:01", "hh:mm")
    txtToTime.Text = Format("23:59", "hh:mm")
   
    cboPageSize.ListIndex = 3
    With cboReport
        .AddItem "1 | Daily Collection Report"
        .AddItem "2 | Shutout Payments Report"
        .AddItem "3 | Shutout Payments for a Container"
        .AddItem "4 | Storage Collection Report"
        .AddItem "5 | Monthly Report"
        .ListIndex = gRpt
    End With
End Sub

Public Sub lzSetParm()
    Select Case cboReport.ListIndex
        Case 0  '1 | Daily Collection Report
            Call lzEnable(txtOther)
            Call lzEnable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzEnable(txtFromTime)
            Call lzEnable(txtToTime)
            lblOther.Caption = "User": txtOther.SetFocus
        Case 1  '2 | Shutout Payments Report
            Call lzDisable(txtOther)
            Call lzEnable(txtFromDate)
            Call lzEnable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            txtFromDate.SetFocus
        Case 2  '3 | Shutout Payments for a Container
            Call lzEnable(txtOther): txtOther.Text = ""
            Call lzDisable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            lblOther.Caption = "Container": txtOther.SetFocus
        Case 3  '4 | Storage Collection Report
            Call lzEnable(txtOther): txtOther.Text = "I"
            Call lzEnable(txtFromDate)
            Call lzEnable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            lblOther.Caption = "(I)mp/(E)xp": txtOther.SetFocus
        Case 4  '5 | Monthly Report
            Call lzEnable(txtOther): txtOther.Text = Format(Now(), "YYYY-MM")
            Call lzDisable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            lblOther.Caption = "Year-Mo ": txtOther.SetFocus
        Case Else
    End Select
End Sub

Public Sub lzCursor2Viewer()
    Dim Rect As Rect
    SetCursorPos 50, 175
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Sub lzViewReport()
Dim rcdSel As String
    On Error GoTo err_wait
    Screen.MousePointer = vbHourglass
    With CRViewer1
        ' select report
        Select Case cboReport.ListIndex
            Case 0
                Set repDailyColxn = Nothing
                Set repDailyColxn = New rptDailyColxn
                With repDailyColxn
                    .txtFromDate.SetText Format(txtFromDate.Text, "####/##/##")
                    .txtFromTime.SetText Format(txtFromTime.Text, "00:00")
                    .txtToTime.SetText Format(txtToTime.Text, "00:00")
                    .txtTeller.SetText (IIf(Len(Trim(txtOther.Text)) = 0, "ALL TELLERS", Trim(txtOther.Text)))
                    .txtFromDate2.SetText Format(txtFromDate.Text, "####/##/##")
                    .txtFromTime2.SetText Format(txtFromTime.Text, "00:00")
                    .txtToTime2.SetText Format(txtToTime.Text, "00:00")
                    .txtTeller2.SetText (IIf(Len(Trim(txtOther.Text)) = 0, "ALL TELLERS", Trim(txtOther.Text)))
                    .txtISORef.SetText (IIf(Len(Trim(txtOther.Text)) = 0, "SR-BAC-019", "SR-BAC-001"))
                    rcdSel = "(date({CCRDtl.sysdttm}) = date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(time({CCRDtl.sysdttm}) >= time(" & left(txtFromTime.Text, 2) & _
                             "," & right(txtFromTime.Text, 2) & ",00)) AND " & _
                             "(time({CCRDtl.sysdttm}) <= time(" & left(txtToTime.Text, 2) & _
                             "," & right(txtToTime.Text, 2) & ",59))"
                    If Len(Trim(txtOther.Text)) > 0 Then
                        rcdSel = "({CCRDtl.userid} = '" & Trim(txtOther.Text) & "') AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = repDailyColxn
            Case 1
                Set repSOList = Nothing
                Set repSOList = New rptSOList
                With repSOList
                    .txtFromDate.SetText Format(txtFromDate.Text, "####/##/##")
                    .txtToDate.SetText Format(txtToDate.Text, "####/##/##")
                    rcdSel = "(date({CCRDtl.sysdttm}) >= date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(date({CCRDtl.sysdttm}) <= date(" & left(txtToDate, 4) & "," & _
                             Mid(txtToDate, 5, 2) & "," & Mid(txtToDate, 7, 2) & ")) AND " & _
                             "({CCRDtl.chargetyp} = 'SOE' OR {CCRDtl.chargetyp} = 'SOF') AND " & _
                             "{CCRDtl.status} <> 'CAN'"
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = repSOList
            Case 2
                Set repSOInquiry = Nothing
                Set repSOInquiry = New rptSOInquiry
                With repSOInquiry
                    .txtContNo.SetText txtOther.Text
                    If Len(Trim(txtOther.Text)) > 0 Then
                        .txtContNo.SetText txtOther.Text
                        .RecordSelectionFormula = "({CCRDtl.cntnum} = '" & Trim(txtOther.Text) & "')"
                    Else
                        .txtContNo.SetText " ALL"
                    End If
                End With
                .ReportSource = repSOInquiry
            Case 3
                Set repStoColxn = Nothing
                Set repStoColxn = New rptStoColxn
                With repStoColxn
                    .txtFromDate.SetText Format(txtFromDate.Text, "####/##/##")
                    .txtToDate.SetText Format(txtToDate.Text, "####/##/##")
                    rcdSel = "(date({CCRDtl.sysdttm}) >= date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(date({CCRDtl.sysdttm}) <= date(" & left(txtToDate, 4) & "," & _
                             Mid(txtToDate, 5, 2) & "," & Mid(txtToDate, 7, 2) & "))"
                    .RecordSelectionFormula = rcdSel
                    If UCase(left(Trim(txtOther.Text), 1)) = "I" Then
                        .txtImpExp.SetText "- IMPORT"
                        .RecordSelectionFormula = "({CCRDtl.chargetyp} = 'IMST') AND ({CCRDtl.status} <> 'CAN') AND " & rcdSel
                    Else
                        .txtImpExp.SetText "- EXPORT"
                        .RecordSelectionFormula = "({CCRDtl.chargetyp} = 'EXST') AND ({CCRDtl.status} <> 'CAN') AND " & rcdSel
                    End If
                End With
                .ReportSource = repStoColxn
            Case 3
                Set repMonthly = Nothing
                Set repMonthly = New rptMonthly
                With repMonthly
'                    .txtFromDate.SetText Format(txtFromDate.Text, "####/##/##")
'                    rcdSel = "(date({CCRDtl.sysdttm}) >= date(" & left(txtFromDate, 4) & "," & _
'                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
'                             "(date({CCRDtl.sysdttm}) <= date(" & left(txtToDate, 4) & "," & _
'                             Mid(txtToDate, 5, 2) & "," & Mid(txtToDate, 7, 2) & "))"
'                    .RecordSelectionFormula = rcdSel
'                    If UCase(left(Trim(txtOther.Text), 1)) = "I" Then
'                        .txtImpExp.SetText "- IMPORT"
'                        .RecordSelectionFormula = "({CCRDtl.chargetyp} = 'IMST') AND ({CCRDtl.status} <> 'CAN') AND " & rcdSel
'                    Else
'                        .txtImpExp.SetText "- EXPORT"
'                        .RecordSelectionFormula = "({CCRDtl.chargetyp} = 'EXST') AND ({CCRDtl.status} <> 'CAN') AND " & rcdSel
'                    End If
                End With
                .ReportSource = repMonthly
        End Select
        ' view report
        .ViewReport
tagRepeat:
        txtPageNo.Text = .GetCurrentPageNumber
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo err_size
    Call lzResizePage
    Call lzCursor2Viewer
    Exit Sub
err_wait:
    DoEvents
    GoTo tagRepeat
err_size:
    DoEvents
    GoTo tagRepeat
End Sub
