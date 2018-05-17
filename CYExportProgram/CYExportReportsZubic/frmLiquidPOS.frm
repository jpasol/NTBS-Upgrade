VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiquidPOS 
   Caption         =   "Liquidator's Cash & PPA Report"
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
   Begin VB.TextBox txtDay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   14520
      MaxLength       =   2
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtSubmit 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   14520
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "Y"
      Top             =   120
      Width           =   615
   End
   Begin zcCCRRpt.prvusrctrlTime rptFromtime 
      Height          =   420
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
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
   Begin zcCCRRpt.prvusrctrlDate RepDte 
      Height          =   420
      Left            =   2880
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
   Begin VB.TextBox noCopy 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   420
      Index           =   1
      Left            =   12840
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox noCopy 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   420
      Index           =   0
      Left            =   12840
      TabIndex        =   5
      Text            =   "2"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdPreviewPPA 
      Caption         =   "Previe&w PPA Copy"
      Height          =   615
      Left            =   7680
      TabIndex        =   13
      Top             =   9480
      Width           =   3855
   End
   Begin zcCCRRpt.ctrlCRViewerNav Nav1 
      Height          =   855
      Left            =   2280
      TabIndex        =   10
      Top             =   8400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1508
   End
   Begin VB.CommandButton cmdCriteria 
      Caption         =   "&Criteria"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   8400
      Width           =   2055
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6975
      Left            =   120
      TabIndex        =   20
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
      Caption         =   "Pre&view Cash Copy"
      Height          =   615
      Left            =   3480
      TabIndex        =   12
      Top             =   9480
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   19
      Top             =   9240
      Width           =   15015
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Pr&int"
      Height          =   615
      Left            =   11640
      Picture         =   "frmLiquidPOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9480
      Width           =   3495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   120
      Picture         =   "frmLiquidPOS.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9480
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   15015
   End
   Begin VB.TextBox Teller 
      BackColor       =   &H8000000F&
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Height          =   7095
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   15015
   End
   Begin zcCCRRpt.prvusrctrlTime rptTotime 
      Height          =   420
      Left            =   9000
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
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
   Begin zcCCRRpt.prvusrctrlDate RepDte2 
      Height          =   420
      Left            =   7080
      TabIndex        =   3
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
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
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
            Object.Width           =   12356
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/11/2000"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "9:25 AM"
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   13440
      TabIndex        =   25
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Submit "
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   13440
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No.CASH Copy"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   10440
      TabIndex        =   22
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To"
      Height          =   420
      Left            =   6360
      TabIndex        =   18
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Teller"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date/Time Range"
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmLiquidPOS"
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

Private Function OutLiquidator(Preview As Boolean, Level As Integer) As Boolean
    
    Dim Mp As Recordset
    
    Dim ChqAmt As String
    Dim AdrAmt As String
    Dim CshAmt As String
    Dim TotalAmt As String
    
    Dim LqRpt As New rptAuditorPOS
    Dim rsCheck As Recordset
    Dim TellerToProcess As String * 10
    
    Dim strRange As String
    Dim strType As String
    Dim strTeller As String
    Dim strComment As String
    Dim strLevel As String * 1
        
    ' ** Setting the default
    OutLiquidator = False
    fromDte = CDate(Trim(RepDte.Text) & " " & Trim(rptFromtime.Text))
    toDte = CDate(Trim(RepDte2.Text) & " " & Trim(rptTotime.Text))
    
    strRange = "From " & Format(fromDte, "yyyy-mm-dd hh:nn:ss") & " To " & Format(toDte, "yyyy-mm-dd hh:nn:ss")
    toDte = DateAdd("n", 1, toDte)
    TellerToProcess = Trim(UCase(Teller.Text))
    strTeller = "Teller  : " & Trim(TellerToProcess)
    
    If Level = 1 Then
        strType = "Assessor/Teller Collection Report (Export Cash Copy)"
        strComment = "ATCR-E (Cash)"
        strLevel = "1"
    Else
        strType = "Assessor/Teller Collection Report (Export PPA Copy)"
        strComment = "ATCR-E (PPA)"
        strLevel = "2"
    End If
    AdrAmt = "0.00"
    CshAmt = "0.00"
    ChqAmt = "0.00"
    TotalAmt = "0.00"
    VE.getTellerAmt fromDte, toDte, TellerToProcess
    Set Mp = VE.rsgetTellerAmt
    If Mp.RecordCount > 0 Then
'        If Not IsNull(Mp.Fields("TotalAdr")) Then
            Call ComputeCash(fromDte, toDte, TellerToProcess)
'            AdrAmt = Mp.Fields("TotalAdr")
            sngCash = sngCash - sngChange
            CshAmt = Format(sngCash, "###,###,###.00")
            ChqAmt = Format(sngCheque, "###,###,###.00")
            sngTotal = Format(sngCash + sngCheque, "###,###,###.00")
            TotalAmt = Format(sngTotal, "###,###,###.00")
    Else
        AdrAmt = "0.00"
    End If
    Mp.Close
    Set Mp = Nothing
    
    ' ** Setting the selection formula
'    LqRpt.RecordSelectionFormula = "{CCRcyx.sysdttm} >= DATETIME(" & Format(fromDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
'                                        & " and {CCRcyx.sysdttm} <= DATETIME(" & Format(toDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
'                                        & " and {CCRcyx.userid} = '" & Trim(TellerToProcess) & "'"
    LqRpt.RecordSelectionFormula = "{CCRpay.sysdttm} >= DATETIME(" & Format(fromDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
                                        & " and {CCRpay.sysdttm} <= DATETIME(" & Format(toDte, "yyyy,mm,dd,hh,mm,ss") & ")" _
                                        & " and {CCRpay.userid} = '" & Trim(TellerToProcess) & "'" _
                                        & " and {CCRpay.ftramt} > 0 and {CCRpay.adramt} = 0 and {CCRpay.refnum} = {CCRcyx.refnum}"

    
    LqRpt.ParameterFields(1).AddCurrentValue (AdrAmt)
    LqRpt.ParameterFields(2).AddCurrentValue (strTeller)
    LqRpt.ParameterFields(3).AddCurrentValue (strType)
    LqRpt.ParameterFields(4).AddCurrentValue (strRange)
    LqRpt.ParameterFields(5).AddCurrentValue (strComment)
    LqRpt.ParameterFields(6).AddCurrentValue (strLevel)
    
    LqRpt.TellerName.SetText (Teller.Text)
    LqRpt.TotalCash.SetText (CshAmt)
    LqRpt.TotalCheque.SetText (ChqAmt)
    LqRpt.Total.SetText (TotalAmt)
    
    ' ** Check if there are records
    
    VE.getLiquidator fromDte, toDte, Trim(TellerToProcess)
    Set rsCheck = VE.rsgetLiquidator
    
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
        If Level = 1 Then
            If Not IsNumeric(noCopy(0).Text) Or noCopy(0).Text = "0" Then
'                noCopy(0).Text = "1"
                GoTo gtEnd
            End If
            LqRpt.PrintOut True, CLng(noCopy(0).Text)
        Else
            If Not IsNumeric(noCopy(1).Text) Or noCopy(1).Text = "0" Then
'                noCopy(1).Text = "1"
                GoTo gtEnd
            End If
            LqRpt.PrintOut True, CLng(noCopy(1).Text)
        End If
gtEnd:
        CRViewer1.Visible = False
        cmdCriteria.Visible = False
        Nav1.Visible = False
    End If
    Teller.SetFocus
End Function

Private Sub cmdPreview_Click()
    Screen.MousePointer = vbHourglass
    Call OutLiquidator(True, 1)
    Nav1.PositionCursor
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPreviewPPA_Click()
    Screen.MousePointer = vbHourglass
    Call OutLiquidator(True, 2)
    Nav1.PositionCursor
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRun_Click()
OutLiquidator False, 1
OutLiquidator False, 2
If Trim(txtSubmit.Text) = "Y" Then
    Call SummarySave
End If
End Sub

Private Sub Form_Load()
    cmdRun.Enabled = False
    cmdPreview.Enabled = False
    CRViewer1.Visible = False
    cmdCriteria.Visible = False
    RepDte.Text = Format(Now, "YYYY-MM-DD")
    RepDte2.Text = Format(Now, "YYYY-MM-DD")
    rptFromtime.Text = "00:01:00"
    rptTotime.Text = "23:58:59"
    Teller.Text = UCase(gUserid)
    txtDay.Text = UCase(Format(Now, "ddd"))
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
        If KeyAscii = 27 Then
            SendKeys "+{Tab}", True
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End If
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

Private Sub txtDay_GotFocus()
    txtDay.BackColor = &HFFFFFF
    txtDay.SelStart = 0
    txtDay.SelLength = Len(txtSubmit.Text)
End Sub

Private Sub txtDay_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        SendKeys "{Tab}", True
    Case 38
        SendKeys "+{Tab}", True
End Select
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    Else
        If KeyAscii = 27 Then
            SendKeys "+{Tab}", True
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End If
End Sub

Private Sub txtDay_LostFocus()
    txtDay.BackColor = &H8000000F
End Sub

Private Sub txtSubmit_GotFocus()
    txtSubmit.BackColor = &HFFFFFF
    txtSubmit.SelStart = 0
    txtSubmit.SelLength = Len(txtSubmit.Text)
End Sub

Private Sub txtSubmit_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        SendKeys "{Tab}", True
    Case 38
        SendKeys "+{Tab}", True
End Select

End Sub

Private Sub txtSubmit_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    Else
        If KeyAscii = 27 Then
            SendKeys "+{Tab}", True
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End If
End Sub

Private Sub txtSubmit_LostFocus()
    txtSubmit.BackColor = &H8000000F
End Sub

Private Sub SummarySave()

Dim rstSummary As Recordset
sngSumArr = 0
sngSumArrVat = 0
sngSumArrNVat = 0
sngSumArrTax = 0
sngSumNetArr = 0
sngSumAdrAmt = 0

On Error GoTo ErrSummarySave

fromDte = CDate(Trim(RepDte.Text) & " " & Trim(rptFromtime.Text))
toDte = CDate(Trim(RepDte2.Text) & " " & Trim(rptTotime.Text))
VE.getArrCyx Trim(Teller.Text), fromDte, toDte, sngSumArr
VE.getArrVatCyx Trim(Teller.Text), fromDte, toDte, sngSumArrVat
VE.getArrNoVatCyx Trim(Teller.Text), fromDte, toDte, sngSumArrNVat
VE.getArrTaxCyx Trim(Teller.Text), fromDte, toDte, sngSumArrTax
VE.getAdrPaidCyx Trim(Teller.Text), fromDte, toDte, sngSumAdrAmt
sngSumNetArr = sngSumArr + sngSumArrVat + sngSumArrNVat - sngSumArrTax
 
VE.CyxSummary

Set rstSummary = VE.rsCyxSummary

With rstSummary
    .AddNew
    .Fields("arramt") = sngSumArr
    .Fields("arrvat") = sngSumArrVat
    .Fields("arrnov") = sngSumArrNVat
    .Fields("arrwtx") = sngSumArrTax
    
    .Fields("stoamt") = 0
    .Fields("stovat") = 0
    .Fields("stonov") = 0
    .Fields("stowtx") = 0
    
    .Fields("wghamt") = 0
    .Fields("wghvat") = 0
    .Fields("wghnov") = 0
    .Fields("wghwtx") = 0
    
    .Fields("rframt") = 0
    .Fields("rfrvat") = 0
    .Fields("rfrnov") = 0
    .Fields("rfrwtx") = 0
    
    .Fields("adrnum") = Space(5)
    .Fields("adramt") = sngSumAdrAmt
    .Fields("adrarr") = sngSumAdrAmt
    .Fields("adrsto") = 0
    .Fields("adrwgh") = 0
    .Fields("adrrfr") = 0
    
    .Fields("impexp") = "E"
    .Fields("daycde") = Trim(txtDay.Text)
    .Fields("status") = Space(3)
    .Fields("userid") = Trim(Teller.Text)
    .Fields("strdte") = fromDte
    .Fields("enddte") = toDte
    
    .Update
    .Close
    Set rstSummary = Nothing
End With
Exit Sub
ErrSummarySave:
    MsgBox "Error writing in header...", vbExclamation + vbOKOnly, "Error!"
End Sub

Private Sub ComputeCash(tmpFrDate As Date, tmpToDate As Date, tmpTeller As String)

Dim rstCash As Recordset
Dim sngChk As Currency

sngCash = 0
sngCheque = 0
sngADR = 0
sngChange = 0
sngTotal = 0

VE.getCash tmpFrDate, tmpToDate, tmpTeller

Set rstCash = VE.rsgetCash

With rstCash
    Do Until .EOF
        sngCash = sngCash + .Fields("cshamt")
        If .Fields("chkamt1") > 0 Then
        sngChk = 0
        sngChk = CSng(.Fields("chkamt1")) + CSng(.Fields("chkamt2")) + _
                            CSng(.Fields("chkamt3")) + CSng(.Fields("chkamt4")) + _
                            CSng(.Fields("chkamt5"))
        sngCheque = sngCheque + sngChk
        End If
        
        sngADR = sngADR + .Fields("adramt")
        If Not IsNull(.Fields("chgamt")) Then
            If .Fields("cshamt") > 0 And .Fields("chgamt") > 0 And _
                .Fields("chkamt1") = 0 And .Fields("chkamt2") = 0 And _
                .Fields("chkamt3") = 0 And .Fields("chkamt5") = 0 Then
                sngChange = sngChange + .Fields("chgamt")
            End If
        End If
    .MoveNext
    Loop
    .Close
Set rstCash = Nothing
End With

End Sub

