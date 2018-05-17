VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptMain 
   Caption         =   "CY Import Reports"
   ClientHeight    =   9345
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   13125
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
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   1980
      Left            =   7125
      TabIndex        =   18
      Top             =   75
      Width           =   5790
      Begin VB.TextBox txtDay 
         Height          =   345
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkAcctBatch 
         Caption         =   "Acct.Batch"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   25
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chkICX 
         Caption         =   "ICX"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtFromDate 
         Height          =   390
         Left            =   1560
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
         Left            =   1575
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
         Left            =   1575
         TabIndex        =   1
         ToolTipText     =   " User ID "
         Top             =   150
         Width           =   3990
         _ExtentX        =   7038
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
         Left            =   4725
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
         Left            =   4725
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Day"
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
         Left            =   3960
         TabIndex        =   26
         Top             =   1560
         Width           =   645
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
         Left            =   3225
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
         Left            =   3225
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
         Width           =   1365
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
         Width           =   1365
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
         Width           =   1365
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
      Left            =   120
      TabIndex        =   16
      Top             =   1320
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
         Left            =   120
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
         Top             =   240
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
      Height          =   7050
      Left            =   75
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2130
      Width           =   12855
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
      EnableToolbar   =   0   'False
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
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

Dim curArrastre As Currency
Dim curArrastreVat As Currency
Dim curArrastreNoVat As Currency
Dim curArrastreTax As Currency
Dim curStorage As Currency
Dim curStorageVat As Currency
Dim curStorageNoVat As Currency
Dim curStorageTax As Currency
Dim curReefer As Currency
Dim curReeferVat As Currency
Dim curReeferNoVat As Currency
Dim curReeferTax As Currency
Dim curWeighing As Currency
Dim curWeighingVat As Currency
Dim curWeighingNoVat As Currency
Dim curWeighingTax As Currency
Dim curADRAmount As Currency
Dim curADRArrastre As Currency
Dim curADRStorage As Currency
Dim curADRWeighing As Currency
Dim curADRReefer As Currency

Dim dtmDateFrom As Date
Dim dtmDateTo As Date

Private Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal y As Long)
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
 Dim rptCYMPR18p As rptCYMPR18
 Dim rptCYMPR13p As rptCYMPR13
 Dim rptCYMPR14p As rptCYMPR14
 Dim rptCYMPR11p As rptCYMPR11
 Dim rptCYMPR05p As rptCYMPR05
 Dim rptCYMPR07p As rptCYMPR07
 Dim rptCYMPR24p As rptCYMPR24
 Dim rptCYMPR12p As rptCYMPR12
 Dim rptCYMSTORp As rptCYMSTOR
 Dim rptCYMCONTIp As rptContEntry
 Dim rptCYMIN10p As rptCYMIN10
 
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
    If chkAcctBatch.Value = 1 Then
        Call ExtractAcctBatch
    End If
    lzViewReport
End Sub

Private Sub ExtractAcctBatch()
    Dim curADRTotal As Currency
    Dim curNetArrastre As Currency
    Dim curNetStorage As Currency
    Dim curNetWeighing As Currency
    Dim curNetReefer As Currency
    Dim strFromDate As String
    Dim strFromTime As String
    Dim strToDate As String
    Dim strToTime As String
    curADRArrastre = 0
    curADRStorage = 0
    curADRWeighing = 0
    curADRReefer = 0
    
    curArrastre = 0
    curArrastreVat = 0
    curArrastreNoVat = 0
    curArrastreTax = 0
    curNetArrastre = 0
    
    curStorage = 0
    curStorageVat = 0
    curStorageNoVat = 0
    curStorageTax = 0
    curNetStorage = 0
    
    curWeighing = 0
    curWeighingVat = 0
    curWeighingNoVat = 0
    curWeighingTax = 0
    curNetWeighing = 0
    
    curReefer = 0
    curReeferVat = 0
    curReeferNoVat = 0
    curReeferTax = 0
    curNetReefer = 0
    
    strFromDate = Format(txtFromDate, "####-##-##")
    strFromTime = left(txtFromTime, 2) & ":" & right(txtFromTime, 2)
    strToTime = left(txtToTime, 2) & ":" & right(txtToTime, 2)
    
    
    dtmDateFrom = CDate(strFromDate & Space(1) & strFromTime & ":00")
    dtmDateTo = CDate(strFromDate & Space(1) & strToTime & ":59")
    
    curArrastre = gzGetArrastre(txtOther, dtmDateFrom, dtmDateTo)
    curArrastreVat = gzGetArrastreVat(txtOther, dtmDateFrom, dtmDateTo)
    curArrastreNoVat = gzGetArrastreNoVat(txtOther, dtmDateFrom, dtmDateTo)
    curArrastreTax = gzGetArrastreTax(txtOther, dtmDateFrom, dtmDateTo)
    curNetArrastre = curArrastre + curArrastreVat + curArrastreNoVat - curArrastreTax
    
    curStorage = gzGetStorage(txtOther, dtmDateFrom, dtmDateTo)
    curStorageVat = gzGetStorageVat(txtOther, dtmDateFrom, dtmDateTo)
    curStorageNoVat = gzGetStorageNoVat(txtOther, dtmDateFrom, dtmDateTo)
    curStorageTax = gzGetStorageTax(txtOther, dtmDateFrom, dtmDateTo)
    curNetStorage = curStorage + curStorageVat + curArrastreNoVat - curStorageTax
    
    curWeighing = gzGetWeighing(txtOther, dtmDateFrom, dtmDateTo)
    curWeighingVat = gzGetWeighingVat(txtOther, dtmDateFrom, dtmDateTo)
    curWeighingNoVat = gzGetWeighingNoVat(txtOther, dtmDateFrom, dtmDateTo)
    curWeighingTax = gzGetWeighingTax(txtOther, dtmDateFrom, dtmDateTo)
    curNetWeighing = curWeighing + curWeighingVat + curWeighingNoVat - curWeighingTax
    
    curReefer = gzGetReefer(txtOther, dtmDateFrom, dtmDateTo)
    curReeferVat = gzGetReeferVat(txtOther, dtmDateFrom, dtmDateTo)
    curReeferNoVat = gzGetReeferNoVat(txtOther, dtmDateFrom, dtmDateTo)
    curReeferTax = gzGetReeferTax(txtOther, dtmDateFrom, dtmDateTo)
    curNetReefer = curReefer + curReeferVat + curReeferNoVat - curReeferTax
    curADRAmount = gzGetADRPaid(txtOther, dtmDateFrom, dtmDateTo)
    curADRTotal = curADRAmount
    If curADRAmount > 0 Then
        If curNetArrastre >= curADRTotal Then
            curADRArrastre = curADRTotal
            curADRTotal = curADRTotal - curADRArrastre
        Else
            curADRArrastre = curNetArrastre
            curADRTotal = curADRTotal - curADRArrastre
        End If
        
        If curADRTotal <= 0 Then
            If curNetStorage >= curADRTotal Then
                curADRStorage = curADRTotal
                curADRTotal = curADRTotal - curADRStorage
            Else
                curADRStorage = curNetStorage
                curADRTotal = curADRTotal - curADRStorage
            End If
        End If
        
        If curADRTotal <= 0 Then
            If curNetWeighing >= curADRTotal Then
                curADRWeighing = curADRTotal
                curADRTotal = curADRTotal - curADRWeighing
            Else
                curADRWeighing = curNetWeighing
                curADRTotal = curADRTotal - curADRWeighing
            End If
        End If
        
        If curADRTotal <= 0 Then
            If curNetReefer >= curADRTotal Then
                curADRReefer = curADRTotal
                curADRTotal = curADRTotal - curADRReefer
            Else
                curADRReefer = curNetReefer
                curADRTotal = curADRTotal - curADRReefer
            End If
        End If
    End If

    Call SaveToSummary
End Sub

Private Sub SaveToSummary()
    Dim rstCYSummary As ADODB.Recordset
    Dim intResponse As Integer
    On Error GoTo ErrSaveSummary
        Set rstCYSummary = New ADODB.Recordset
        With rstCYSummary
            .LockType = adLockOptimistic
            .CursorType = adOpenDynamic
            .Open "CYSummary", gcnnBilling, , , adCmdTable

            .AddNew
            .Fields("arramt") = curArrastre
            .Fields("arrvat") = curArrastreVat
            .Fields("arrnov") = curArrastreNoVat
            .Fields("arrwtx") = curArrastreTax
            
            .Fields("stoamt") = curStorage
            .Fields("stovat") = curStorageVat
            .Fields("stonov") = curStorageNoVat
            .Fields("stowtx") = curStorageTax
            
            .Fields("wghamt") = curWeighing
            .Fields("wghvat") = curWeighingVat
            .Fields("wghnov") = curWeighingNoVat
            .Fields("wghwtx") = curWeighingTax
            
            .Fields("rframt") = curReefer
            .Fields("rfrvat") = curReeferVat
            .Fields("rfrnov") = curReeferNoVat
            .Fields("rfrwtx") = curReeferTax
            .Fields("adrnum") = Space(5)
            .Fields("adramt") = curADRAmount
            .Fields("adrarr") = curADRArrastre
            .Fields("adrsto") = curADRStorage
            .Fields("adrwgh") = curADRWeighing
            .Fields("adrrfr") = curADRReefer
            .Fields("impexp") = "I"
            .Fields("daycde") = txtDay
            .Fields("status") = Space(3)
            .Fields("userid") = txtOther
            .Fields("strdte") = dtmDateFrom
            .Fields("enddte") = dtmDateTo
            .Update
            .Close
        End With
    On Error GoTo 0
    Exit Sub

ErrSaveSummary:
    intResponse = MsgBox("Error writing in header...", vbExclamation + vbDefaultButton2 + vbAbortRetryIgnore, "Error!")
    If intResponse = vbAbort Then
        Unload Me
    ElseIf (intResponse = vbRetry) Or (intResponse = vbIgnore) Then
        Resume
    End If
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

Private Sub Form_Activate()
    chkAcctBatch.Value = 1
End Sub

Public Sub Form_Load()
    lzInitialize
End Sub

Public Sub Form_Resize()
    CRViewer1.Height = ScaleHeight - 2300
    CRViewer1.Width = ScaleWidth - 150
End Sub

Public Sub lzViewReport()
    Dim rcdSel As String
    Dim curArrastre As Currency

    On Error GoTo err_wait
    Screen.MousePointer = vbHourglass
    With CRViewer1
       'select report
        Select Case cboReport.ListIndex
           'Assessor/Teller Turnover Report
            Case 0
'                If chkSubmitToBatch.Value = 1 Then
''                    curArrastre = gzGetArrastre(Trim(Format(txtFromDate, "yyyy-mm-dd")) & Space(1) & Trim(Format(txtFromTime, "hh:mm")) & _
'                    "59.999", Trim(Format(txtToDate, "yyyy-mm-dd")) & Space(1) & Trim(Format(txtToTime, "hh:mm")) & "00.000")
'
'                    curArrastre = gzGetArrastre("1999-07-28", "1999-07-28")
'                End If
                Set rptCYMPR18p = Nothing
                Set rptCYMPR18p = New rptCYMPR18
                With rptCYMPR18p
                    .txtDate.SetText Format(txtFromDate, "####/##/##")
                    .txtFromTime.SetText Format(txtFromTime, "##:##")
                    .txtToTime.SetText Format(txtToTime, "##:##")
                    .txtTeller.SetText (IIf(Len(Trim(txtOther.Text)) = 0, "ALL TELLERS", Trim(txtOther.Text)))
                    
                    rcdSel = "(date({CYMgps.sysdte}) = date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(time({CYMgps.sysdte}) >= time(" & left(txtFromTime.Text, 2) & _
                             "," & right(txtFromTime.Text, 2) & ",00)) AND " & _
                             "(time({CYMgps.sysdte}) <= time(" & left(txtToTime.Text, 2) & _
                             "," & right(txtToTime.Text, 2) & ",59))"
                    If Len(Trim(txtOther.Text)) > 0 Then
                        rcdSel = "({CYMgps.userid} = '" & Trim(txtOther.Text) & "') AND " & rcdSel
                    End If
                    If chkICX.Value = 1 Then
                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
                    Else
                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = rptCYMPR18p
           'Daily Summary Report (per Reference)
            Case 1
                Set rptCYMPR13p = Nothing
                Set rptCYMPR13p = New rptCYMPR13
                With rptCYMPR13p
                    .txtDate.SetText Format(txtFromDate, "####/##/##")
                    .txtFromTime.SetText Format(txtFromTime, "##:##")
                    .txtToTime.SetText Format(txtToTime, "##:##")
                    .txtTeller.SetText (IIf(Len(Trim(txtOther.Text)) = 0, "ALL TELLERS", Trim(txtOther.Text)))
                    
                    rcdSel = "(date({CYMgps.sysdte}) = date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(time({CYMgps.sysdte}) >= time(" & left(txtFromTime.Text, 2) & _
                             "," & right(txtFromTime.Text, 2) & ",00)) AND " & _
                             "(time({CYMgps.sysdte}) <= time(" & left(txtToTime.Text, 2) & _
                             "," & right(txtToTime.Text, 2) & ",59))"
                    If Len(Trim(txtOther.Text)) > 0 Then
                        rcdSel = "({CYMgps.userid} = '" & Trim(txtOther.Text) & "') AND " & rcdSel
                    End If
                    If chkICX.Value = 1 Then
                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
                    Else
                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = rptCYMPR13p
           'Daily Summary Report (per CCR)
            Case 2
                Set rptCYMPR14p = Nothing
                Set rptCYMPR14p = New rptCYMPR14
                With rptCYMPR14p
                    .txtDate.SetText Format(txtFromDate, "####/##/##")
                    .txtFromTime.SetText Format(txtFromTime, "##:##")
                    .txtToTime.SetText Format(txtToTime, "##:##")
                    .txtTeller.SetText (IIf(Len(Trim(txtOther.Text)) = 0, "ALL TELLERS", Trim(txtOther.Text)))
                    
'                    rcdSel = "(date({CYMgps.sysdte}) = date(" & left(txtFromDate, 4) & "," & _
'                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
'                             "(time({CYMgps.sysdte}) >= time(" & left(txtFromTime.Text, 2) & _
'                             "," & right(txtFromTime.Text, 2) & ",00)) AND " & _
'                             "(time({CYMgps.sysdte}) <= time(" & left(txtToTime.Text, 2) & _
'                             "," & right(txtToTime.Text, 2) & ",59))"

                    rcdSel = "(date({CYMgps.sysdte}) = date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(time({CYMgps.sysdte}) >= time(" & left(txtFromTime.Text, 2) & _
                             "," & right(txtFromTime.Text, 2) & ",00)) AND " & _
                             "(time({CYMgps.sysdte}) <= time(" & left(txtToTime.Text, 2) & _
                             "," & right(txtToTime.Text, 2) & ",59)) AND " & _
                             "(date({CYMpay.sysdttm}) = date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(time({CYMpay.sysdttm}) >= time(" & left(txtFromTime.Text, 2) & _
                             "," & right(txtFromTime.Text, 2) & ",00)) AND " & _
                             "(time({CYMpay.sysdttm}) <= time(" & left(txtToTime.Text, 2) & _
                             "," & right(txtToTime.Text, 2) & ",59))"

                    If Len(Trim(txtOther.Text)) > 0 Then
                        rcdSel = "({CYMgps.userid} = '" & Trim(txtOther.Text) & "') AND " & rcdSel
                    End If
                    If chkICX.Value = 1 Then
                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
                    Else
                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = rptCYMPR14p
           'Wharfage Exemption Report
            Case 3
                Set rptCYMPR11p = Nothing
                Set rptCYMPR11p = New rptCYMPR11
                With rptCYMPR11p
                    .txtFromDate.SetText Format(txtFromDate, "####/##/##")
                    .txtToDate.SetText Format(txtToDate, "####/##/##")
                     rcdSel = "(date({CYMgps.sysdte}) >= date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(date({CYMgps.sysdte}) <= date(" & left(txtToDate, 4) & "," & _
                             Mid(txtToDate, 5, 2) & "," & Mid(txtToDate, 7, 2) & "))"
                   
                    If chkICX.Value = 1 Then
                       'rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND Trim({CYMgps.gpstyp})= '1' AND " & rcdSel
                    Else
                       'rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND Trim({CYMgps.gpstyp})= '1' AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel & " AND {Cymgps.status}<>'CAN' AND {Cymgps.whfcde}='Y'"
                End With
                .ReportSource = rptCYMPR11p
           'Import Gatepass Summary Report
            Case 4
                Set rptCYMPR05p = Nothing
                Set rptCYMPR05p = New rptCYMPR05
                With rptCYMPR05p
                    .txtFromDate.SetText Format(txtFromDate, "####/##/##")
                    .txtToDate.SetText Format(txtToDate, "####/##/##")
                    .txtFromDate2.SetText Format(txtFromDate, "####/##/##")
                    .txtToDate2.SetText Format(txtToDate, "####/##/##")
                    rcdSel = "(date({CYMgps.sysdte}) >= date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(date({CYMgps.sysdte}) <= date(" & left(txtToDate, 4) & "," & _
                             Mid(txtToDate, 5, 2) & "," & Mid(txtToDate, 7, 2) & "))"
                    
                    If chkICX.Value = 1 Then
                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
                    Else
                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel & " AND {Cymgps.status}<>'CAN'"
                End With
                .ReportSource = rptCYMPR05p
           'Cancelled Gatepass Report
            Case 5
                Set rptCYMPR12p = Nothing
                Set rptCYMPR12p = New rptCYMPR12
                With rptCYMPR12p
                    .txtFromDate.SetText Format(txtFromDate, "####/##/##")
                    .txtToDate.SetText Format(txtToDate, "####/##/##")
                    rcdSel = "(date({CYMgps.sysdte}) >= date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(date({CYMgps.sysdte}) <= date(" & left(txtToDate, 4) & "," & _
                             Mid(txtToDate, 5, 2) & "," & Mid(txtToDate, 7, 2) & "))"
                             
                    If chkICX.Value = 1 Then
                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
                    Else
                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel & " AND {CYMgps.status} = 'CAN'"
                End With
                .ReportSource = rptCYMPR12p
           'Storage Collection Report
            Case 6
                Set rptCYMSTORp = Nothing
                Set rptCYMSTORp = New rptCYMSTOR
                With rptCYMSTORp
                    .txtFromDate.SetText Format(txtFromDate, "####/##/##")
                    .txtToDate.SetText Format(txtToDate, "####/##/##")
                    rcdSel = "(date({CYMgps.sysdte}) >= date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(date({CYMgps.sysdte}) <= date(" & left(txtToDate, 4) & "," & _
                             Mid(txtToDate, 5, 2) & "," & Mid(txtToDate, 7, 2) & "))"
                    
                    If chkICX.Value = 1 Then
                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
                    Else
                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel & "AND {CYMGPS.status} <> 'CAN' AND {CYMGPS.stoday}>0"
                End With
                .ReportSource = rptCYMSTORp
           'Monthly Report
            Case 7
                Set rptCYMPR24p = Nothing
                Set rptCYMPR24p = New rptCYMPR24
                With rptCYMPR24p
                    .txtDate.SetText MonthName(Val(Mid(txtFromDate, 5, 2))) & " " & left(txtFromDate, 4)
                    
                    rcdSel = "year({cymgps.sysdte}) = " & Val(left(txtFromDate, 4)) _
                                    & " and  month({cymgps.sysdte}) = " & Val(Mid(txtFromDate, 5, 2))
                    
                    If chkICX.Value = 1 Then
                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
                    Else
                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
                    End If
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = rptCYMPR24p
           'Assessor/Teller Turnover Report (PPA Copy)
            Case 8
                Set rptCYMPR07p = Nothing
                Set rptCYMPR07p = New rptCYMPR07
                With rptCYMPR07p
                    .txtDate.SetText Format(txtFromDate, "####/##/##")
                    .txtFromTime.SetText Format(txtFromTime, "##:##")
                    .txtToTime.SetText Format(txtToTime, "##:##")
                    .txtTeller.SetText (IIf(Len(Trim(txtOther.Text)) = 0, "ALL TELLERS", Trim(txtOther.Text)))
                    
                    rcdSel = "(date({CYMgps.sysdte}) = date(" & left(txtFromDate, 4) & "," & _
                             Mid(txtFromDate, 5, 2) & "," & Mid(txtFromDate, 7, 2) & ")) AND " & _
                             "(time({CYMgps.sysdte}) >= time(" & left(txtFromTime.Text, 2) & _
                             "," & right(txtFromTime.Text, 2) & ",00)) AND " & _
                             "(time({CYMgps.sysdte}) <= time(" & left(txtToTime.Text, 2) & _
                             "," & right(txtToTime.Text, 2) & ",59))"
                    If Len(Trim(txtOther.Text)) > 0 Then
                        rcdSel = "({CYMgps.userid} = '" & Trim(txtOther.Text) & "') AND " & rcdSel
                    End If
                    
'                    If chkICX.Value = 1 Then
'                        rcdSel = "Trim({UserInfo.dptcde}) = 'ICX' AND " & rcdSel
'                    Else
'                        rcdSel = "Trim({UserInfo.dptcde}) <> 'ICX' AND " & rcdSel
'                    End If
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = rptCYMPR07p
            Case 9
                Set rptCYMCONTIp = Nothing
                Set rptCYMCONTIp = New rptContEntry
                With rptCYMCONTIp
                    rcdSel = "{ACOctn.ctn_ctnnum} = '" & Trim(txtOther.Text) & "'"
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = rptCYMCONTIp
            Case 10
                Set rptCYMIN10p = Nothing
                Set rptCYMIN10p = New rptCYMIN10
                With rptCYMIN10p
                    rcdSel = "{cymgps.gpsnum} = " & Val(txtOther.Text)
                    .RecordSelectionFormula = rcdSel
                End With
                .ReportSource = rptCYMIN10p
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

Public Sub Form_Unload(Cancel As Integer)
    If (MsgBox("Exit CY Import Reports ?", vbYesNo, "CYS Reports") = vbNo) Then
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

Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strDay As String
    Dim strFromDate As String
    If KeyCode = vbKeyReturn Then
        strFromDate = left(txtFromDate, 4) & "-" & Mid(txtFromDate, 5, 2) & "-" & right(txtFromDate, 2)
        strDay = DatePart("w", strFromDate)
        Select Case strDay
            Case "1"
                txtDay = "SU"
            Case "2"
                txtDay = "MO"
            Case "3"
                txtDay = "TU"
            Case "4"
                txtDay = "WE"
            Case "5"
                txtDay = "TH"
            Case "6"
                txtDay = "FR"
            Case "7"
                txtDay = "SA"
        End Select
    End If
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

Public Sub lzInitialize()
    txtOther.Text = gUserID: lblOther.Caption = "User"
    txtFromDate.Text = Format(Now, "yyyy-mm-dd")
    txtToDate.Text = Format(Now, "yyyy-mm-dd")
    txtFromTime.Text = Format(Now, "hh:mm")
    txtToTime.Text = Format(Now, "hh:mm")
   
    cboPageSize.ListIndex = 3
    With cboReport
        .AddItem "1 | Assessor/Teller Turnover Report"
        .AddItem "2 | Daily Summary Report (Auditors Copy) per Reference"
        .AddItem "3 | Daily Summary Report (Auditors Copy) per CCR"
        .AddItem "4 | Wharfage Exemption Report"
        .AddItem "5 | Import Gatepass Summary Report"
        .AddItem "6 | Cancelled Gatepass Report"
        .AddItem "7 | Storage Collection Report"
        .AddItem "8 | Monthly Report"
        .AddItem "9 | Assessor/Teller Turnover Report (PPA Copy)"
        .AddItem "10| Container Entry Inquiry"
        .AddItem "11| Inquire by Gatepass "
        .ListIndex = gRpt
    End With
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

Public Sub lzSetParm()
    Select Case cboReport.ListIndex
        Case 0
            Call lzEnable(txtOther)
            Call lzEnable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzEnable(txtFromTime)
            Call lzEnable(txtToTime)
            lblOther.Caption = "User": txtOther.SetFocus
        Case 1
            Call lzEnable(txtOther)
            Call lzEnable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzEnable(txtFromTime)
            Call lzEnable(txtToTime)
            lblOther.Caption = "User": txtOther.SetFocus
        Case 2
            Call lzEnable(txtOther)
            Call lzEnable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzEnable(txtFromTime)
            Call lzEnable(txtToTime)
            lblOther.Caption = "User": txtOther.SetFocus
        Case 3
            Call lzDisable(txtOther): txtOther.Text = ""
            Call lzEnable(txtFromDate)
            Call lzEnable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            txtFromDate.SetFocus
        Case 4
            Call lzDisable(txtOther): txtOther.Text = ""
            Call lzEnable(txtFromDate)
            Call lzEnable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            txtFromDate.SetFocus
        Case 5
            Call lzDisable(txtOther): txtOther.Text = ""
            Call lzEnable(txtFromDate)
            Call lzEnable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            txtFromDate.SetFocus
        Case 6
            Call lzDisable(txtOther): txtOther.Text = ""
            Call lzEnable(txtFromDate)
            Call lzEnable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            txtFromDate.SetFocus
        Case 7
            Call lzDisable(txtOther): txtOther.Text = ""
            Call lzEnable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            txtFromDate.SetFocus
        Case 8
            Call lzEnable(txtOther)
            Call lzEnable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzEnable(txtFromTime)
            Call lzEnable(txtToTime)
            lblOther.Caption = "User": txtOther.SetFocus
        Case 9
            Call lzEnable(txtOther)
            Call lzDisable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            lblOther.Caption = "Container": txtOther.SetFocus
        Case 10
            Call lzEnable(txtOther)
            Call lzDisable(txtFromDate)
            Call lzDisable(txtToDate)
            Call lzDisable(txtFromTime)
            Call lzDisable(txtToTime)
            lblOther.Caption = "Gatepass": txtOther.SetFocus
        Case Else
    End Select
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

Public Sub lzCursor2Viewer()
    Dim Rect As Rect
    SetCursorPos 50, 175
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Function gzGetArrastre(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetArrastre As ADODB.Command
    Dim prmGetArrastre As ADODB.Parameter

    ' create command
    Set cmdGetArrastre = New ADODB.Command
    Set prmGetArrastre = New ADODB.Parameter
    With cmdGetArrastre
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getarrastrewvat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetArrastre = 0
        Else
            gzGetArrastre = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetArrastreCYX(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetArrastre As ADODB.Command
    Dim prmGetArrastre As ADODB.Parameter

    ' create command
    Set cmdGetArrastre = New ADODB.Command
    Set prmGetArrastre = New ADODB.Parameter
    With cmdGetArrastre
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getarrastrewvatcyx"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetArrastreCYX = 0
        Else
            gzGetArrastreCYX = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetArrastreVat(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetArrastre As ADODB.Command
    Dim prmGetArrastre As ADODB.Parameter

    ' create command
    Set cmdGetArrastre = New ADODB.Command
    Set prmGetArrastre = New ADODB.Parameter
    With cmdGetArrastre
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getarrastrevat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetArrastreVat = 0
        Else
            gzGetArrastreVat = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetArrastreVatCYX(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetArrastre As ADODB.Command
    Dim prmGetArrastre As ADODB.Parameter

    ' create command
    Set cmdGetArrastre = New ADODB.Command
    Set prmGetArrastre = New ADODB.Parameter
    With cmdGetArrastre
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getarrastrevatcyx"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetArrastreVatCYX = 0
        Else
            gzGetArrastreVatCYX = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetArrastreNoVat(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetArrastre As ADODB.Command
    Dim prmGetArrastre As ADODB.Parameter

    ' create command
    Set cmdGetArrastre = New ADODB.Command
    Set prmGetArrastre = New ADODB.Parameter
    With cmdGetArrastre
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getarrastrenovat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetArrastreNoVat = 0
        Else
            gzGetArrastreNoVat = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetArrastreNoVatCYX(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetArrastre As ADODB.Command
    Dim prmGetArrastre As ADODB.Parameter

    ' create command
    Set cmdGetArrastre = New ADODB.Command
    Set prmGetArrastre = New ADODB.Parameter
    With cmdGetArrastre
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getarrastrenovatcyx"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetArrastreNoVatCYX = 0
        Else
            gzGetArrastreNoVatCYX = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetArrastreTax(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetArrastre As ADODB.Command
    Dim prmGetArrastre As ADODB.Parameter

    ' create command
    Set cmdGetArrastre = New ADODB.Command
    Set prmGetArrastre = New ADODB.Parameter
    With cmdGetArrastre
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getarrastretax"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetArrastreTax = 0
        Else
            gzGetArrastreTax = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetArrastreTaxCYX(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetArrastre As ADODB.Command
    Dim prmGetArrastre As ADODB.Parameter

    ' create command
    Set cmdGetArrastre = New ADODB.Command
    Set prmGetArrastre = New ADODB.Parameter
    With cmdGetArrastre
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getarrastretaxcyx"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetArrastreTaxCYX = 0
        Else
            gzGetArrastreTaxCYX = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetStorage(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetStorage As ADODB.Command
    Dim prmGetStorage As ADODB.Parameter

    ' create command
    Set cmdGetStorage = New ADODB.Command
    Set prmGetStorage = New ADODB.Parameter
    With cmdGetStorage
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getstoragewvat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetStorage = 0
        Else
            gzGetStorage = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetStorageVat(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetStorage As ADODB.Command
    Dim prmGetStorage As ADODB.Parameter

    ' create command
    Set cmdGetStorage = New ADODB.Command
    Set prmGetStorage = New ADODB.Parameter
    With cmdGetStorage
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getstoragevat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetStorageVat = 0
        Else
            gzGetStorageVat = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetStorageNoVat(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetStorage As ADODB.Command
    Dim prmGetStorage As ADODB.Parameter

    ' create command
    Set cmdGetStorage = New ADODB.Command
    Set prmGetStorage = New ADODB.Parameter
    With cmdGetStorage
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getstoragenovat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetStorageNoVat = 0
        Else
            gzGetStorageNoVat = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetStorageTax(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetStorage As ADODB.Command
    Dim prmGetStorage As ADODB.Parameter

    ' create command
    Set cmdGetStorage = New ADODB.Command
    Set prmGetStorage = New ADODB.Parameter
    With cmdGetStorage
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getstoragetax"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetStorageTax = 0
        Else
            gzGetStorageTax = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetWeighing(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetWeighing As ADODB.Command
    Dim prmGetWeighing As ADODB.Parameter

    ' create command
    Set cmdGetWeighing = New ADODB.Command
    Set prmGetWeighing = New ADODB.Parameter
    With cmdGetWeighing
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getweighingwvat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetWeighing = 0
        Else
            gzGetWeighing = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetWeighingVat(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetWeighing As ADODB.Command
    Dim prmGetWeighing As ADODB.Parameter

    ' create command
    Set cmdGetWeighing = New ADODB.Command
    Set prmGetWeighing = New ADODB.Parameter
    With cmdGetWeighing
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getweighingvat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetWeighingVat = 0
        Else
            gzGetWeighingVat = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetWeighingNoVat(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetWeighing As ADODB.Command
    Dim prmGetWeighing As ADODB.Parameter

    ' create command
    Set cmdGetWeighing = New ADODB.Command
    Set prmGetWeighing = New ADODB.Parameter
    With cmdGetWeighing
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getweighingnovat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetWeighingNoVat = 0
        Else
            gzGetWeighingNoVat = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetWeighingTax(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetWeighing As ADODB.Command
    Dim prmGetWeighing As ADODB.Parameter

    ' create command
    Set cmdGetWeighing = New ADODB.Command
    Set prmGetWeighing = New ADODB.Parameter
    With cmdGetWeighing
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getweighingtax"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetWeighingTax = 0
        Else
            gzGetWeighingTax = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetReefer(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetReefer As ADODB.Command
    Dim prmGetReefer As ADODB.Parameter

    ' create command
    Set cmdGetReefer = New ADODB.Command
    Set prmGetReefer = New ADODB.Parameter
    With cmdGetReefer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getreeferwvat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetReefer = 0
        Else
            gzGetReefer = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetReeferVat(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetReefer As ADODB.Command
    Dim prmGetReefer As ADODB.Parameter

    ' create command
    Set cmdGetReefer = New ADODB.Command
    Set prmGetReefer = New ADODB.Parameter
    With cmdGetReefer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getreefervat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetReeferVat = 0
        Else
            gzGetReeferVat = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetReeferNoVat(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetReefer As ADODB.Command
    Dim prmGetReefer As ADODB.Parameter

    ' create command
    Set cmdGetReefer = New ADODB.Command
    Set prmGetReefer = New ADODB.Parameter
    With cmdGetReefer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getreefernovat"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetReeferNoVat = 0
        Else
            gzGetReeferNoVat = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetReeferTax(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetReefer As ADODB.Command
    Dim prmGetReefer As ADODB.Parameter

    ' create command
    Set cmdGetReefer = New ADODB.Command
    Set prmGetReefer = New ADODB.Parameter
    With cmdGetReefer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getreefertax"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetReeferTax = 0
        Else
            gzGetReeferTax = .Parameters(4)
        End If
    End With
End Function

Public Function gzGetADRPaid(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetADRPaid As ADODB.Command
    Dim prmGetADRPaid As ADODB.Parameter

    ' create command
    Set cmdGetADRPaid = New ADODB.Command
    Set prmGetADRPaid = New ADODB.Parameter
    With cmdGetADRPaid
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getadrpaid"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetADRPaid = 0
        Else
            gzGetADRPaid = .Parameters(4)
        End If
    End With
End Function


Public Function gzGetADRPaidCYX(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetADRPaid As ADODB.Command
    Dim prmGetADRPaid As ADODB.Parameter

    ' create command
    Set cmdGetADRPaid = New ADODB.Command
    Set prmGetADRPaid = New ADODB.Parameter
    With cmdGetADRPaid
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getadrpaidcyx"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetADRPaidCYX = 0
        Else
            gzGetADRPaidCYX = .Parameters(4)
        End If
    End With
End Function


