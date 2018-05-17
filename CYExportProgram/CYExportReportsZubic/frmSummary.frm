VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSummary 
   Caption         =   " Summary Collection Report"
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
   Begin zcCCRRpt.ctrlCRViewerNav Nav1 
      Height          =   855
      Left            =   2400
      TabIndex        =   7
      Top             =   8400
      Width           =   9375
      _extentx        =   16536
      _extenty        =   1508
   End
   Begin zcCCRRpt.prvusrctrlTime rptFromtime 
      Height          =   420
      Left            =   4440
      TabIndex        =   2
      Top             =   600
      Width           =   1455
      _extentx        =   2566
      _extenty        =   741
      font            =   "frmSummary.frx":0000
      maxlength       =   8
   End
   Begin zcCCRRpt.prvusrctrlDate RepDte 
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1815
      _extentx        =   3201
      _extenty        =   741
      font            =   "frmSummary.frx":0030
   End
   Begin VB.TextBox txtNoted 
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   9600
      TabIndex        =   5
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox noCopy 
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   7920
      TabIndex        =   4
      Text            =   "1"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdCriteria 
      Caption         =   "&Criteria"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   8400
      Width           =   2055
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
      Caption         =   "Pre&view"
      Height          =   615
      Left            =   8400
      TabIndex        =   9
      Top             =   9480
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   240
      TabIndex        =   16
      Top             =   9240
      Width           =   15015
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Pr&int"
      Height          =   615
      Left            =   11880
      Picture         =   "frmSummary.frx":0060
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9480
      Width           =   3375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   240
      Picture         =   "frmSummary.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9480
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   15015
   End
   Begin VB.TextBox Teller 
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Height          =   7095
      Left            =   120
      TabIndex        =   18
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
            TextSave        =   "10/19/00"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:34 AM"
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
   Begin zcCCRRpt.prvusrctrlTime rptTotime 
      Height          =   420
      Left            =   6360
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _extentx        =   2566
      _extenty        =   741
      font            =   "frmSummary.frx":02F4
      maxlength       =   8
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Noted By :"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   9600
      TabIndex        =   20
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. Copy"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   7920
      TabIndex        =   19
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To"
      Height          =   420
      Left            =   5880
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
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Teller"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   1920
      TabIndex        =   12
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Dim LqRpt As New rptTeller

Private Sub cmdCriteria_Click()
    RepDte.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Function OutLiquidator(Preview As Boolean) As Boolean
    
    Dim Mp As Recordset
    Dim AdrAmt As String
    Dim LqRpt As New rptSummary
    Dim rsCheck As Recordset
    Dim fromDte As Date
    Dim toDte As Date
    Dim TellerToProcess As String * 10
    
    Dim strRange As String
    Dim strType As String
    Dim strTeller As String
    Dim strComment As String
    Dim strLevel As String * 1
        
    ' ** Setting the default
    OutLiquidator = False
    fromDte = CDate(Trim(RepDte.Text) & " " & Trim(rptFromtime.Text))
    toDte = CDate(Trim(RepDte.Text) & " " & Trim(rptTotime.Text))
    TellerToProcess = Trim(UCase(Teller.Text))
    
    If Len(Trim(TellerToProcess)) = 0 Then
        VE.getTotalAdr fromDte, toDte
        Set Mp = VE.rsgetTotalAdr
        LqRpt.RecordSelectionFormula = "{CCRcyx.sysdttm} >= DATETIME(" & Format(fromDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
                & " and {CCRcyx.sysdttm} <= DATETIME(" & Format(toDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
                & " and  {CCRcyx.status} <> 'CAN' "
    Else
        VE.getTellerTotalAdr fromDte, toDte, TellerToProcess
        Set Mp = VE.rsgetTellerTotalAdr
        LqRpt.RecordSelectionFormula = "{CCRcyx.sysdttm} >= DATETIME(" & Format(fromDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
                                        & " and {CCRcyx.sysdttm} <= DATETIME(" & Format(toDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
                                        & " and {CCRcyx.userid} = '" & Trim(TellerToProcess) & "'" _
                                         & " and  {CCRcyx.status} <> 'CAN'  "
    End If
    If Mp.RecordCount > 0 Then
        If Not IsNull(Mp.Fields("TotalAdr")) Then
            AdrAmt = Mp.Fields("TotalAdr")
        Else
            AdrAmt = "0.00"
        End If
    Else
        AdrAmt = "0.00"
    End If
    Mp.Close
    Set Mp = Nothing
    
    LqRpt.ParameterFields(1).AddCurrentValue (AdrAmt)
    LqRpt.ParameterFields(2).AddCurrentValue (CDate(Trim(RepDte.Text)))
    LqRpt.ParameterFields(3).AddCurrentValue (TellerToProcess)
    LqRpt.ParameterFields(4).AddCurrentValue (Trim(txtNoted.Text))
    ' LqRpt.ParameterFields(5).AddCurrentValue ("")
    
    ' ** Determining the Output
    If Preview Then
        CRViewer1.Visible = True
        cmdCriteria.Visible = True
        Nav1.Visible = True
        LqRpt.PaperOrientation = crPortrait
        CRViewer1.ReportSource = LqRpt
        CRViewer1.ViewReport
        Nav1.cmdRefresh_Click
    Else
        LqRpt.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        LqRpt.PaperOrientation = crPortrait
        LqRpt.DisplayProgressDialog = False
        If IsNumeric(noCopy.Text) Then
            LqRpt.PrintOut False, CInt(noCopy)
        Else
            LqRpt.PrintOut
        End If
        CRViewer1.Visible = False
        cmdCriteria.Visible = False
        Nav1.Visible = False
    End If
    RepDte.SetFocus
End Function
Private Sub cmdPreview_Click()
    Screen.MousePointer = vbHourglass
    Call OutLiquidator(True)
    Nav1.PositionCursor
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdPreviewPPA_Click()
    Screen.MousePointer = vbHourglass
    Call OutLiquidator(True)
    Nav1.PositionCursor
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdRun_Click()
    OutLiquidator False
End Sub
Private Sub Form_Load()
    cmdRun.Enabled = False
    cmdPreview.Enabled = False
    CRViewer1.Visible = False
    cmdCriteria.Visible = False
    RepDte.Text = Format(Now, "YYYY-MM-DD")
    rptFromtime.Text = "00:00:01"
    rptTotime.Text = "23:58:59"
    Teller.Text = UCase(gUserid)
    Call Initialize
    Nav1.Visible = False
    Set Nav1.CRViewerControl = CRViewer1
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
If KeyAscii <> 8 Then
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End If
End Sub
Private Sub Teller_LostFocus()
    Teller.BackColor = &H8000000F
'    If Len(Trim(Teller.Text)) = 0 Then
'        Teller.Text = gUserid
'    End If
End Sub
Private Sub Teller_GotFocus()
    Teller.BackColor = &HFFFFFF
    Teller.SelStart = 0
    Teller.SelLength = Len(Teller.Text)
End Sub
Private Sub RepDte_Change()
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
Private Sub txtNoted_GotFocus()
    txtNoted.BackColor = &HFFFFFF
    txtNoted.SelStart = 0
    txtNoted.SelLength = Len(txtNoted.Text)
End Sub
Private Sub txtNoted_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        SendKeys "{Tab}", True
    Case 38
        SendKeys "+{Tab}", True
End Select
End Sub
Private Sub txtNoted_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    Else
        If KeyAscii = 27 Then
            SendKeys "+{Tab}", True
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub
Private Sub txtNoted_LostFocus()
    txtNoted.BackColor = &H8000000F
End Sub
