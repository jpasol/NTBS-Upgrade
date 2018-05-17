VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTS"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   14535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   14295
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   6855
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14055
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   0   'False
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   0   'False
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
      End
   End
   Begin VB.Frame fraParameter 
      Caption         =   " Parameters "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1740
      Left            =   7200
      TabIndex        =   13
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox txtCusnam 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
      Begin MSMask.MaskEdBox txtToDate 
         Height          =   390
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   " End of date range "
         Top             =   960
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
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
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lblOther 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1845
      End
   End
   Begin VB.Frame fraControl 
      Caption         =   " Control "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   765
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   6990
      Begin VB.ComboBox cboPageSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRpt.frx":0000
         Left            =   3840
         List            =   "frmRpt.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   " Zoom "
         Top             =   240
         Width           =   1485
      End
      Begin VB.CommandButton cmdPage 
         Height          =   390
         Index           =   3
         Left            =   2700
         Picture         =   "frmRpt.frx":0048
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Last Page"
         Top             =   240
         Width           =   390
      End
      Begin VB.CommandButton cmdPage 
         Height          =   390
         Index           =   2
         Left            =   2250
         Picture         =   "frmRpt.frx":0192
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Next Page"
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdPage 
         Height          =   390
         Index           =   0
         Left            =   1350
         Picture         =   "frmRpt.frx":02DC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   " First Page "
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdPage 
         Height          =   390
         Index           =   1
         Left            =   1800
         Picture         =   "frmRpt.frx":0426
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Previous Page"
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   390
         Left            =   120
         Picture         =   "frmRpt.frx":0570
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " View / Refresh "
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   390
         Left            =   750
         Picture         =   "frmRpt.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Print "
         Top             =   225
         Width           =   390
      End
      Begin MSMask.MaskEdBox txtPageNo 
         Height          =   390
         Left            =   3120
         TabIndex        =   6
         ToolTipText     =   " Page No "
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "F3 - EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5400
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox cmbRpt 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuRef 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FRpt As FirstRpt
Dim SRpt As SecondRpt
Dim FlRpt As FinalRpt
Dim SORpt As SOArpt
Dim RptSum As SumRpt
Dim ArRpt As AccountsReceive
Dim SLrpt As SubLedger
Private Sub Fill_Cust()
Dim vRec As ADODB.Recordset
Set vRec = New ADODB.Recordset
vRec.Open "Select Cusnam from customer", gcnnBilling, adOpenStatic, adLockOptimistic, adCmdText
With vRec
    If Not (.BOF And .EOF) Then
        .MoveFirst
        While Not .EOF
            With txtCusnam
                .AddItem vRec.Fields("cusnam")
            End With
            .MoveNext
        Wend
    End If
End With
Set vRec = Nothing
End Sub
Public Sub lzViewReport()
    Dim rcdSel As String

    On Error GoTo err_wait
    Screen.MousePointer = vbHourglass
    With CRViewer1
       'select report
        Select Case cmbRpt.ListIndex
        
        '1st Notice
        Case 0
            Set FRpt = Nothing
            Set FRpt = New FirstRpt
            With FRpt
                .txtCustomer.SetText Trim(txtCusnam.Text)
                .txtNotice.SetText Mid(cmbRpt.Text, 5, 12)
                .SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                .PaperOrientation = crPortrait
                .PaperSize = crPaperLetter

             End With
            .ReportSource = FRpt
            
        '2ND NOTICE
        Case 1
            Set SRpt = Nothing
            Set SRpt = New SecondRpt
            With SRpt
                .txtCustomer.SetText Trim(txtCusnam.Text)
                .txtNotice.SetText Mid(cmbRpt, 5, 12)
                .SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                .PaperOrientation = crPortrait
                .PaperSize = crPaperLetter

            End With
            .ReportSource = SRpt
            
        
        'FINAL NOTICE
        Case 2
            Set FlRpt = Nothing
            Set FlRpt = New FinalRpt
            With FlRpt
                .txtCustomer.SetText Trim(txtCusnam.Text)
                .txtNotice.SetText Mid(cmbRpt, 5, 15)
                .txtTdate.SetText Format(txtToDate, "####/##/##")

                rcdSel = "((date({invict.invdttm}) <= date(" & Left(txtToDate, 4) & "," & _
                             Mid(txtToDate, 5, 2) & "," & Mid(txtToDate, 7, 2) & "))"

                If Len(Trim(txtCusnam.Text)) > 0 Then
                    rcdSel = rcdSel & " AND ({invict.cusnam} = '" & Trim(txtCusnam.Text) & "'))"
                End If
                .RecordSelectionFormula = rcdSel
                .SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                .PaperOrientation = crPortrait
                .PaperSize = crPaperLetter

            End With
            .ReportSource = FlRpt
        
        'Statement of accounts
        Case 3
            Set SORpt = Nothing
            Set SORpt = New SOArpt
            With SORpt
                .txtCustomer.SetText Trim(txtCusnam.Text)
                        
            If Len(Trim(txtCusnam.Text)) > 0 Then
                rcdSel = "({invict.cusnam} = '" & Trim(txtCusnam.Text) & "')"
            End If
            .RecordSelectionFormula = rcdSel
            .SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            .PaperOrientation = crPortrait
            .PaperSize = crPaperLetter

            End With
            .ReportSource = SORpt
        
        'Summary
        Case 4
            Set RptSum = Nothing
            Set RptSum = New SumRpt
            With RptSum
                .txtCustomer.SetText Trim(txtCusnam.Text)
            
            If Len(Trim(txtCusnam.Text)) > 0 Then
                rcdSel = "({invict.cusnam} = '" & Trim(txtCusnam.Text) & "')"
            End If
                .RecordSelectionFormula = rcdSel
                .SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                .PaperOrientation = crLandscape
                .PaperSize = crPaperA4
            End With
            .ReportSource = RptSum
        
        'Accounts Receivable
        Case 5
            Set ArRpt = Nothing
            Set ArRpt = New AccountsReceive
            With ArRpt
                .SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                .PaperOrientation = crLandscape
                .PaperSize = crPaperA4
    
            End With
            .ReportSource = ArRpt
            
        'Subsidiary ledger
        Case 6
            Set SLrpt = Nothing
            Set SLrpt = New SubLedger
            With SLrpt
                .txtCustomer.SetText Trim(txtCusnam.Text)
                
            If Len(Trim(txtCusnam.Text)) > 0 Then
                rcdSel = "({invict.cusnam} = '" & Trim(txtCusnam.Text) & "')"
            End If
                .RecordSelectionFormula = rcdSel
                .SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                .PaperOrientation = crPortrait
                .PaperSize = crPaperA4
            End With
            .ReportSource = SLrpt
        End Select
        .ViewReport
        
        
tagRepeat:
        txtPageNo.Text = .GetCurrentPageNumber
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo err_size
    Exit Sub
err_wait:
    DoEvents
    GoTo tagRepeat
err_size:
    DoEvents
    GoTo tagRepeat
End Sub

Private Sub cboPageSize_Click()
    Call lzResizePage
End Sub

Private Sub cboPageSize_GotFocus()
    cboPageSize.BackColor = &H80000018
End Sub

Private Sub cmbRpt_GotFocus()
    cmbRpt.BackColor = &H80000018
    cmbRpt.SelStart = 0
    cmbRpt.SelLength = Len(cmbRpt.Text)
End Sub

Private Sub cmbRpt_LostFocus()
    cmbRpt.BackColor = &H80000014
End Sub

Private Sub cmdPrint_Click()
     CRViewer1.PrintReport
End Sub

Private Sub cmdRefresh_Click()
    Call lzViewReport
End Sub

Private Sub Form_Load()
With cmbRpt
    .AddItem "1  |  1st NOTICE"
    .AddItem "2  |  2nd NOTICE"
    .AddItem "3  |  FINAL NOTICE"
    .AddItem "4  |  STATEMENT OF ACCOUNTS"
    .AddItem "5  |  SUMMARY"
    .AddItem "6  |  ACCOUNTS RECEIVABLE-TRADE"
    .AddItem "7  |  SUBSIDIARY LEDGER"
End With
Call Fill_Cust
txtToDate.Text = Format(Now, "yyyy-mm-dd")
End Sub

Public Sub lzResizePage()
Dim i, sz As Integer
    i = cboPageSize.ListIndex
    If (i < 5) Then
        sz = (Left(cboPageSize.List(i), 3))
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
      '  SetMouseFocus cmdPage(Index)
    End With
    Exit Sub
err_wait:
    DoEvents
    GoTo tagRepeat
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (MsgBox("Exit Generation of Reports ?", vbYesNo, "Subic Reports") = vbNo) Then
        Cancel = 1
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuprint_Click()
    Call cmdPrint_Click
End Sub

Private Sub mnuRef_Click()
    Call cmdRefresh_Click
End Sub

Private Sub txtCusnam_GotFocus()
    txtCusnam.BackColor = &H80000018
    txtCusnam.SelStart = 0
    txtCusnam.SelLength = Len(txtCusnam.Text)
End Sub

Private Sub txtCusnam_LostFocus()
    txtCusnam.BackColor = &H80000014
    txtCusnam.Text = UCase(txtCusnam.Text)
End Sub

Private Sub txtPageNo_GotFocus()
    txtPageNo.BackColor = &H80000018
    txtPageNo.SelStart = 0
    txtPageNo.SelLength = Len(txtPageNo.Text)
    End Sub

Private Sub txtPageNo_LostFocus()
    txtPageNo.BackColor = &H80000014
End Sub

Private Sub txtToDate_GotFocus()
    txtToDate.BackColor = &H80000018
    txtToDate.SelStart = 0
    txtToDate.SelLength = 10
End Sub

Private Sub txtToDate_LostFocus()
    txtToDate.BackColor = &H80000014
End Sub
