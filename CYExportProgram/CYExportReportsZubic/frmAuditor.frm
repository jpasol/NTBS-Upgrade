VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAuditor 
   Caption         =   "Auditor's Report"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
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
      Left            =   2760
      TabIndex        =   6
      Top             =   8400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1508
   End
   Begin zcCCRRpt.prvusrctrlTime rptFromtime 
      Height          =   420
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
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
            TextSave        =   "10/17/00"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "2:00 PM"
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
   Begin VB.TextBox noCopy 
      BackColor       =   &H8000000F&
      Height          =   420
      Index           =   0
      Left            =   5760
      TabIndex        =   3
      Text            =   "1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Pre&view"
      Height          =   615
      Left            =   9000
      TabIndex        =   8
      Top             =   9600
      Width           =   2655
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Pr&int"
      Height          =   615
      Left            =   11760
      Picture         =   "frmAuditor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9600
      Width           =   3375
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
      TabIndex        =   13
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
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   12
      Top             =   9240
      Width           =   15015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   120
      Picture         =   "frmAuditor.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9600
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   15015
   End
   Begin VB.Frame Frame3 
      Height          =   7095
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   15015
   End
   Begin VB.PictureBox RepDte2 
      Enabled         =   0   'False
      Height          =   420
      Left            =   8160
      ScaleHeight     =   360
      ScaleWidth      =   1755
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin zcCCRRpt.prvusrctrlTime rptTotime 
      Height          =   420
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To"
      Height          =   420
      Left            =   3480
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No. of Copy"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   5760
      TabIndex        =   15
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date/Time Range"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmAuditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim sngSumArr As Single
Dim sngSumArrVat As Single
Dim sngSumArrNVat As Single
Dim sngSumArrTax As Single
Dim sngSumAdrAmt As Single
Dim sngSumNetArr As Single

Dim fromDte As Date
Dim toDte As Date

Dim sngCash As Currency
Dim sngCheque As Currency
Dim sngADR  As Currency
Dim sngChange As Currency
Dim sngTotal As Currency


Private Sub cmdCriteria_Click()
    RepDte.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Function OutLiquidator(Preview As Boolean) As Boolean
    
    Dim Mp As Recordset
    
    Dim ChqAmt As String
    Dim AdrAmt As String
    Dim CshAmt As String
    Dim TotalAmt As String
   
    Dim LqRpt As New rptAuditor
    Dim rsCheck As Recordset
    Dim fromDte As Date
    Dim toDte As Date
    Dim TellerToProcess As String * 10
    
    Dim strRange As String
    Dim strType As String
        
    ' ** Setting the default
    OutLiquidator = False
    fromDte = CDate(Trim(RepDte.Text) & " " & Trim(rptFromtime.Text))
    toDte = CDate(Trim(RepDte.Text) & " " & Trim(rptTotime.Text))
    
    strRange = "From " & Format(fromDte, "yyyy-mm-dd hh:nn:ss") & " To " & Format(toDte, "yyyy-mm-dd hh:nn:ss")
    strType = "Daily Collection Report (Export Auditor's Copy)"
    toDte = DateAdd("n", 1, toDte)
    
    VE.getTotalAdr fromDte, toDte
    Set Mp = VE.rsgetTotalAdr
    If Mp.RecordCount > 0 Then
        If Not IsNull(Mp.Fields("TotalAdr")) Then
            Call ComputeCash(fromDte, toDte, TellerToProcess)
            AdrAmt = Mp.Fields("TotalAdr")
            'sngCash = sngCash - sngChange
            CshAmt = Format(sngCash, "###,###,###.00")
            ChqAmt = Format(sngCheque, "###,###,###.00")
            sngTotal = Format(sngCash + sngCheque, "###,###,###.00")
            TotalAmt = Format(sngTotal, "###,###,###.00")
        Else
            AdrAmt = "0.00"
            CshAmt = "0.00"
            ChqAmt = "0.00"
            TotalAmt = "0.00"
        End If
'        If (Not IsNull(Mp.Fields("TotalAdr"))) And (Not IsNull(Mp.Fields("TotalCsh"))) And (Not IsNull(Mp.Fields("TotalChq"))) Then
'            AdrAmt = Mp.Fields("TotalAdr")
'            CshAmt = Format(Mp.Fields("TotalCsh"), "###,###,###.00")
'            ChqAmt = Format(Mp.Fields("TotalChq"), "###,###,###.00")
'            TotalAmt = Format(Mp.Fields("Total"), "###,###,###.00")
'        Else
'            AdrAmt = "0.00"
'            CshAmt = "0.00"
'            ChqAmt = "0.00"
'            TotalAmt = "0.00"
'        End If
    Else
        AdrAmt = "0.00"
    End If
    Mp.Close
    Set Mp = Nothing
    
    ' ** Setting the selection formula
    LqRpt.RecordSelectionFormula = "{CCRcyx.sysdttm} >= DATETIME(" & Format(fromDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
                                        & " and {CCRcyx.sysdttm} <= DATETIME(" & Format(toDte, "yyyy,mm,dd,hh,mm,ss") & ")"
    LqRpt.ParameterFields(1).AddCurrentValue (AdrAmt)
    LqRpt.ParameterFields(2).AddCurrentValue (" ")
    LqRpt.ParameterFields(3).AddCurrentValue (strType)
    LqRpt.ParameterFields(4).AddCurrentValue (strRange)
    LqRpt.ParameterFields(5).AddCurrentValue ("DCR-E (Auditor)")
    LqRpt.ParameterFields(6).AddCurrentValue ("1")
    LqRpt.TotalCash.SetText Format(CshAmt, "####,###,###.00")
    LqRpt.TotalCheque.SetText Format(ChqAmt, "####,###,###.00")
    LqRpt.Total.SetText Format(TotalAmt, "####,###,###.00")
    ' ** Check if there are records
    VE.getConsLiquidator fromDte, toDte
    Set rsCheck = VE.rsgetConsLiquidator
    If rsCheck.RecordCount < 1 Then
        Beep
        MsgBox "There are no records extracted from the given Criteria, Please Try Again", vbExclamation + vbOKOnly, "No Records Found"
        rsCheck.Close
        Set rsCheck = Nothing
        CRViewer1.Visible = False
        cmdCriteria.Visible = False
        Nav1.Visible = False
        Exit Function
    End If
    
    ' ** Closing the checker
    rsCheck.Close
    Set rsCheck = Nothing
    
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
        LqRpt.DisplayProgressDialog = False
        LqRpt.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        LqRpt.PaperOrientation = crPortrait
        If IsNumeric(noCopy(0).Text) Then
            LqRpt.PrintOut True, CInt(noCopy(0).Text)
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
Private Sub cmdRun_Click()
    Screen.MousePointer = vbHourglass
    Call OutLiquidator(False)
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbDefault
    cmdRun.Enabled = False
    CRViewer1.Visible = False
    cmdCriteria.Visible = False
    Nav1.Visible = False
    cmdPreview.Enabled = False
    RepDte.Text = Format(Now, "YYYY-MM-DD")
'    RepDte2.Text = Format(Now, "YYYY-MM-DD")
    rptFromtime.Text = "00:01:00"
    rptTotime.Text = "23:58:59"
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

Private Sub noCopy_GotFocus(Index As Integer)
    noCopy(Index).BackColor = &HFFFFFF
    noCopy(Index).SelStart = 0
    noCopy(Index).SelLength = Len(noCopy(Index).Text)
End Sub
Private Sub noCopy_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        SendKeys "{Tab}", True
    Case 38
        SendKeys "+{Tab}", True
End Select
End Sub
Private Sub noCopy_KeyPress(Index As Integer, KeyAscii As Integer)
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
Private Sub noCopy_LostFocus(Index As Integer)
    noCopy(Index).BackColor = &H8000000F
    If Len(Trim(noCopy(Index).Text)) = 0 Then
        noCopy(Index).Text = "1"
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
Private Sub ComputeCash(tmpFrDate As Date, tmpToDate As Date, tmpTeller As String)
Dim rstCash As Recordset
Dim sngChk As Currency
Dim sngTmpCsh As Currency
Dim sngTmpChk As Currency
Dim sngTmpAdr As Currency
Dim sngTmpChange As Currency
Dim sngTmpTotal As Currency



sngCash = 0
sngCheque = 0
sngADR = 0
sngChange = 0
sngTotal = 0


VE.getCashAll tmpFrDate, tmpToDate
Set rstCash = VE.rsgetCashAll

With rstCash
    Do Until .EOF
        sngTmpCsh = 0
        sngTmpChk = 0
        sngTmpAdr = 0
        sngTmpChange = 0
        sngTmpTotal = 0
        
        sngTmpCsh = .Fields("cshamt")
        If .Fields("chkamt1") > 0 Then
        
        sngTmpChk = CSng(.Fields("chkamt1")) + CSng(.Fields("chkamt2")) + _
                            CSng(.Fields("chkamt3")) + CSng(.Fields("chkamt4")) + _
                            CSng(.Fields("chkamt5"))
        End If
        sngTmpChange = .Fields("chgamt")
        If sngTmpChk > 0 And sngTmpChange > 0 And sngTmpCsh > 0 Then
            sngTmpCsh = sngTmpCsh - sngTmpChange
        Else
            If sngTmpChk = 0 And sngTmpChange > 0 And sngTmpCsh > 0 Then
                sngTmpCsh = sngTmpCsh - sngTmpChange
            End If
        End If
        
        sngCash = sngCash + sngTmpCsh
        sngCheque = sngCheque + sngTmpChk
        
'        If Not IsNull(.Fields("chgamt")) Then
'            If .Fields("cshamt") > 0 And .Fields("chgamt") > 0 And _
'                .Fields("chkamt1") = 0 And .Fields("chkamt2") = 0 And _
'                .Fields("chkamt3") = 0 And .Fields("chkamt5") = 0 Then
'
'            End If
'        End If
        
    .MoveNext
    Loop
    sngTotal = sngCash + sngCheque
    .Close
Set rstCash = Nothing
End With

End Sub



