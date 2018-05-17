VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSubicINVReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Reports"
   ClientHeight    =   11145
   ClientLeft      =   90
   ClientTop       =   1410
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
   Icon            =   "frmSubicINVReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraEntries 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   4695
      Begin VB.TextBox txtStaff 
         BackColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin MSMask.MaskEdBox mskMonth 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEnd 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483628
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskStart 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483628
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Billing Staff"
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   2145
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         Height          =   300
         Left            =   1680
         TabIndex        =   15
         Top             =   2520
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         Height          =   300
         Left            =   840
         TabIndex        =   14
         Top             =   1320
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
         Height          =   300
         Left            =   1200
         TabIndex        =   13
         Top             =   1800
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   4455
   End
   Begin VB.PictureBox picChkMrk 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   360
      Picture         =   "frmSubicINVReports.frx":014A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picChkMrk 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   360
      Picture         =   "frmSubicINVReports.frx":058C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Preview"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Monthly Report"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Daily Collection Report"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   10605
      Left            =   4920
      TabIndex        =   8
      Top             =   240
      Width           =   10005
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
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmSubicINVReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intRptTyp As Integer

Private Sub cmdReport_Click(Index As Integer)
    Select Case Index
        Case 0      ' Collection Report
            intRptTyp = 0
            picChkMrk(0).Visible = True
            picChkMrk(1).Visible = False
            mskStart.Enabled = True: mskEnd.Enabled = True: mskMonth.Enabled = False
            mskStart.SetFocus
        Case 1      ' Summary Report
            intRptTyp = 1
            picChkMrk(1).Visible = True
            picChkMrk(0).Visible = False
            mskStart.Enabled = False: mskEnd.Enabled = False: mskMonth.Enabled = True
            mskMonth.SetFocus
        Case Else
    End Select
    txtStaff.SetFocus
End Sub

Private Sub Form_Load()
    txtStaff.Text = gUserID
    intRptTyp = 0
    picChkMrk(0).Visible = True
    mskStart.Text = Format(Now, "YYYY/MM/DD")
    mskEnd.Text = Format(Now, "YYYY/MM/DD")
    mskMonth.Text = Format(Now, "MM/YYYY")
    mskStart.Enabled = True: mskEnd.Enabled = True: mskMonth.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    Select Case intRptTyp
        Case 0
            If Not IsDate(mskStart) Or Not IsDate(mskEnd) Then GoTo endPrint
            PrintDaily
        Case 1
            If Not IsDate(mskMonth) Then GoTo endPrint
            PrintMonthly
        Case Else
    End Select
    
    Exit Sub
endPrint: MsgBox "Please specify valid entries.", vbInformation, "Error"
          txtStaff.SetFocus
End Sub

Private Sub mskEnd_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub mskMonth_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub mskStart_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtStaff_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtStaff_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub PrintDaily()
    Dim Report As New Daily
    
    If Trim(txtStaff) = "" Then
        Report.RecordSelectionFormula = "{INVcyb.sysdttm} >= DATE(" & Format(mskStart, "yyyy,mm,dd") & ")" _
                                        & " and {INVcyb.sysdttm} <= DATE(" & Format(mskEnd, "yyyy,mm,dd") & ")"
    Else
        Report.RecordSelectionFormula = "{INVcyb.sysdttm} >= DATE(" & Format(mskStart, "yyyy,mm,dd") & ")" _
                                        & " and {INVcyb.sysdttm} <= DATE(" & Format(mskEnd, "yyyy,mm,dd") & ")" _
                                        & " and {INVcyb.userid} = '" & Trim(txtStaff) & "'"
    End If

    '   Fill in date headings of report
    Report.rtxtStart.SetText (mskStart.Text)
    Report.rtxtEnd.SetText (mskEnd.Text)
    Report.rtxtStaff.SetText (txtStaff.Text)
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
End Sub

Private Sub PrintMonthly()
    Dim Report As New Monthly
    Dim strSelect As String
    Dim strSelect1 As String

    If Trim(txtStaff) = "" Then
        strSelect = "MONTH({INVict.invdttm})=(" & Format(mskMonth, "mm") & ")" _
                                         & " and YEAR({INVict.invdttm}) = (" & Format(mskMonth, "yyyy") & ")"
        strSelect1 = "MONTH({INVcyb.sysdttm})=(" & Format(mskMonth, "mm") & ")" _
                                         & " and YEAR({INVcyb.sysdttm}) = (" & Format(mskMonth, "yyyy") & ")" _
                                         & " and ({INVcyb.status} <> 'CAN')"
    Else
        strSelect = "MONTH({INVict.invdttm})=(" & Format(mskMonth, "mm") & ")" _
                                         & " and YEAR({INVict.invdttm}) = (" & Format(mskMonth, "yyyy") & ")" _
                                        & " and {INVict.userid} = '" & Trim(txtStaff) & "'"
        strSelect1 = "MONTH({INVcyb.sysdttm})=(" & Format(mskMonth, "mm") & ")" _
                                         & " and YEAR({INVcyb.sysdttm}) = (" & Format(mskMonth, "yyyy") & ")" _
                                        & " and {INVcyb.userid} = '" & Trim(txtStaff) & "'" _
                                        & " and ({INVcyb.status} <> 'CAN')"
    End If
    
    Report.RecordSelectionFormula = strSelect
    Report.Subreport1.OpenSubreport.RecordSelectionFormula = strSelect1
    '   Fill in date headings of report
    Screen.MousePointer = vbHourglass
    Report.rtxtStaff.SetText (txtStaff.Text)
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
End Sub
