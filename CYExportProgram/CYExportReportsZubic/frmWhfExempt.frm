VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWhfExempt 
   Caption         =   "Wharfage Exemption Report"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10890
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin zcCCRRpt.ctrlCRViewerNav Nav5 
      Height          =   855
      Left            =   2760
      TabIndex        =   6
      Top             =   8400
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1508
   End
   Begin zcCCRRpt.prvusrctrlTime rptFromtime 
      Height          =   420
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   8
   End
   Begin zcCCRRpt.prvusrctrlDate RepDte 
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox noCopy 
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   9000
      TabIndex        =   4
      Text            =   "3"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   15015
   End
   Begin VB.CommandButton cmdCriteria 
      Caption         =   "&Criteria"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   8400
      Width           =   2535
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6975
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1320
      Width           =   15015
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "F7 - Pre&view"
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      Top             =   9480
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   16
      Top             =   9240
      Width           =   15015
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "F11 - Pr&int"
      Height          =   615
      Left            =   11760
      Picture         =   "frmWhfExempt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9480
      Width           =   3375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "F3 - E&xit"
      Height          =   615
      Left            =   120
      Picture         =   "frmWhfExempt.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9480
      Width           =   2535
   End
   Begin VB.TextBox Teller 
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   3360
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Height          =   7095
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   15015
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   10515
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12277
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/11/00"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "6:59 AM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "CCRRPT"
            TextSave        =   "CCRRPT"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin zcCCRRpt.prvusrctrlDate RepDte1 
      Height          =   420
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin zcCCRRpt.prvusrctrlTime rptTotime 
      Height          =   420
      Left            =   7080
      TabIndex        =   3
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   8
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. Copy"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   9000
      TabIndex        =   20
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To"
      Height          =   420
      Left            =   2040
      TabIndex        =   18
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To"
      Height          =   420
      Left            =   6480
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time Range"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   4680
      TabIndex        =   14
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Teller"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   3360
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Range"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmWhfExempt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub noCopy_GotFocus()
    noCopy.BackColor = &HFFFFFF
    noCopy.SelStart = 0
    noCopy.SelLength = Len(noCopy.Text)
End Sub
Private Sub noCopy_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        SendKeys "{Tab}", True
    Case 38
        SendKeys "+{Tab}", True
End Select
End Sub
Private Sub noCopy_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If KeyAscii = 13 Then
            SendKeys "{Tab}", True
        Else
            If KeyAscii = 27 Then
                SendKeys "+{Tab}", True
            Else
                If KeyAscii < 48 Or KeyAscii > 57 Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        End If
    End If
End Sub

Private Sub noCopy_LostFocus()
    noCopy.BackColor = &H8000000F
    If Len(Trim(noCopy.Text)) = 0 Then
        noCopy.Text = "1"
    End If
End Sub

Private Sub cmdCriteria_Click()
    RepDte.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Function OutLiquidator(Preview As Boolean) As Boolean
    
    Dim LqRpt As New rptWhfExempt
    Dim rsCheck As Recordset
    Dim fromDte As Date
    Dim toDte As Date
    Dim TellerToProcess As String * 10
    
    ' ** Setting the default
    OutLiquidator = False
    fromDte = CDate(Trim(RepDte.Text) & " " & Trim(rptFromtime.Text))
    toDte = CDate(Trim(RepDte1.Text) & " " & Trim(rptTotime.Text))
    toDte = DateAdd("n", 1, toDte)
    TellerToProcess = Trim(Teller.Text)
    ' ** Setting the selection formula
    
    LqRpt.RecordSelectionFormula = "{CCrcyx.sysdttm} >= DATETIME(" & Format(fromDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
            & " and {CCRcyx.sysdttm} <= DATETIME(" & Format(toDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
            & " and {CCRcyx.status} <> 'WHF' and {CCRcyx.status} <> 'CAN' and ({CCRcyx.whfcde} = '1' or {CCRcyx.whfcde} = '2') "
    
    ' ** Check if there are records
    VE.getExemption fromDte, toDte  ', Trim(TellerToProcess)
    Set rsCheck = VE.rsgetExemption
    
    If rsCheck.RecordCount < 1 Then
        Beep
        MsgBox "There are no records extracted from the given Criteria, Please Try Again", vbExclamation + vbOKOnly, "No Records Found"
        rsCheck.Close
        Set rsCheck = Nothing
        CRViewer1.Visible = False
        cmdCriteria.Visible = False
        Nav5.Visible = False
        Exit Function
    End If
    
    ' ** Closing the checker
    rsCheck.Close
    Set rsCheck = Nothing
    
    ' ** putting the details to report file
    LqRpt.txtRange.SetText ("As of " & Trim(Format(fromDte, "mmmm dd, yyyy ")) & " - " & Trim(Format(toDte, "mmmm dd, yyyy ")))
        
    ' ** Determining the Output
    If Preview Then
        CRViewer1.Visible = True
        cmdCriteria.Visible = True
        Nav5.Visible = True
        LqRpt.PaperOrientation = crLandscape
        LqRpt.PaperSize = crPaperA4
        CRViewer1.ReportSource = LqRpt
        CRViewer1.ViewReport
        Nav5.cmdRefresh_Click
    Else
        LqRpt.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        LqRpt.PaperOrientation = crLandscape
        LqRpt.PaperSize = crPaperA4
        LqRpt.DisplayProgressDialog = False
              
        If IsNumeric(noCopy.Text) Then
            LqRpt.PrintOut True, CInt(noCopy.Text)
        Else
            LqRpt.PrintOut
        End If
        CRViewer1.Visible = False
        cmdCriteria.Visible = False
        Nav5.Visible = False
    End If
    RepDte.SetFocus
End Function
Private Sub cmdPreview_Click()
    Screen.MousePointer = vbHourglass
    Call OutLiquidator(True)
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdRun_Click()
    Screen.MousePointer = vbHourglass
    Call OutLiquidator(False)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdView_Click()
CRViewer1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Call cmdExit_Click
        Case vbKeyF11
            If cmdRun.Enabled Then
                Call cmdRun_Click
            End If
        Case vbKeyF7
            If cmdPreview.Enabled Then
                Call cmdPreview_Click
            End If
    End Select
'    cmdPreview.Enabled = cmdRun.Enabled
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbDefault
    cmdRun.Enabled = False
    Teller.Text = UCase(gUserid)
    CRViewer1.Visible = False
    cmdCriteria.Visible = False
    cmdPreview.Enabled = False
    RepDte.Text = Format(Now, "YYYY-MM-DD")
    RepDte1.Text = Format(Now, "YYYY-MM-DD")
    rptFromtime.Text = "00:01:00"
    rptTotime.Text = "23:58:59"
    Set Nav5.CRViewerControl = CRViewer1
    Nav5.Visible = False
    Call Initialize
End Sub
Private Sub Initialize()
Dim rsUsr As Recordset
    VE.getInformation
    Set rsUsr = VE.rsgetInformation
    SBar.Panels(1) = gUserid
    SBar.Panels(2) = rsUsr.Fields("workstation")
    rsUsr.Close
    Set rsUsr = Nothing
    SBar.Panels(3) = "Printer Device : " & Printer.DeviceName
End Sub

Private Sub RepDte1_Change()
If IsDate(RepDte.Text) And IsDate(RepDte1.Text) And IsDate(rptFromtime.Text) And IsDate(rptTotime.Text) Then
        If CDate(rptFromtime.Text) < CDate(rptTotime.Text) And CDate(RepDte.Text) <= CDate(RepDte1.Text) Then
            cmdRun.Enabled = True
        Else
            cmdRun.Enabled = False
        End If
    Else
        cmdRun.Enabled = False
    End If
    cmdPreview.Enabled = cmdRun.Enabled
    cmdCriteria.Enabled = cmdRun.Enabled
End Sub

Private Sub rptFromtime_Change()
    If IsDate(RepDte.Text) And IsDate(rptFromtime.Text) And IsDate(rptTotime.Text) Then
        If CDate(rptFromtime.Text) < CDate(rptTotime.Text) Then
            cmdRun.Enabled = True
        Else
            cmdRun.Enabled = False
        End If
    Else
        cmdRun.Enabled = False
    End If
    cmdPreview.Enabled = cmdRun.Enabled
    cmdCriteria.Enabled = cmdRun.Enabled
End Sub
Private Sub rptTotime_Change()
    If IsDate(RepDte.Text) And IsDate(rptFromtime.Text) And IsDate(rptTotime.Text) Then
        If CDate(rptFromtime.Text) < CDate(rptTotime.Text) Then
            cmdRun.Enabled = True
        Else
            cmdRun.Enabled = False
        End If
    Else
        cmdRun.Enabled = False
    End If
    cmdPreview.Enabled = cmdRun.Enabled
    cmdCriteria.Enabled = cmdRun.Enabled
End Sub

Private Sub Teller_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        SendKeys "{Tab}", True
    Case 38
        SendKeys "+{Tab}", True
End Select
End Sub
Private Sub Teller_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Teller_LostFocus()
    Teller.BackColor = &H8000000F
    If Len(Trim(Teller.Text)) = 0 Then
        Teller.Text = gUserid
    End If
End Sub
Private Sub Teller_GotFocus()
    Teller.BackColor = &HFFFFFF
    Teller.SelStart = 0
    Teller.SelLength = Len(Teller.Text)
End Sub
Private Sub RepDte_Change()
    If IsDate(RepDte.Text) And IsDate(RepDte1.Text) And IsDate(rptFromtime.Text) And IsDate(rptTotime.Text) Then
        If CDate(rptFromtime.Text) < CDate(rptTotime.Text) And CDate(RepDte.Text) <= CDate(RepDte1.Text) Then
            cmdRun.Enabled = True
        Else
            cmdRun.Enabled = False
        End If
    Else
        cmdRun.Enabled = False
    End If
    cmdPreview.Enabled = cmdRun.Enabled
    cmdCriteria.Enabled = cmdRun.Enabled
End Sub
