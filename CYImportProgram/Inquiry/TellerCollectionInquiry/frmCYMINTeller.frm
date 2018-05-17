VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCYMINTeller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CY Import Teller's Collection Inquiry"
   ClientHeight    =   11055
   ClientLeft      =   255
   ClientTop       =   525
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   15030
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1335
      Left            =   480
      TabIndex        =   13
      Top             =   9120
      Width           =   14415
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   735
         Left            =   8880
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   735
         Left            =   11640
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   10440
      TabIndex        =   10
      Top             =   0
      Width           =   4455
      Begin VB.Label Label4 
         Caption         =   "Teller ID"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblUserID 
         Alignment       =   2  'Center
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   480
      TabIndex        =   6
      Top             =   0
      Width           =   9975
      Begin MSMask.MaskEdBox mskDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483628
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTimeB 
         Height          =   375
         Left            =   8160
         TabIndex        =   2
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483628
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTimeA 
         Height          =   375
         Left            =   6120
         TabIndex        =   1
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483628
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Date "
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Time "
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7935
      Left            =   360
      TabIndex        =   14
      Top             =   1320
      Width           =   14415
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   0   'False
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmCYMINTeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblUserID.Caption = gUserid
    mskDate.Text = Format(gzGetSysDate, "yyyy/mm/dd")
    mskTimeA.Text = Format(gzGetSysDate, "hh:mm")
    mskTimeB.Text = Format(gzGetSysDate, "hh:mm")
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
'    CRViewer1.Visible = False
    mskDate.SetFocus
End Sub

Private Sub cmdView_Click()
    Dim Report As New rptCYMINTeller
    Dim fromDte As Date
    Dim toDte As Date
        
    If Not ValidReport Then
        MsgBox "Please specify valid entries.", vbInformation, "Error Message"
        mskDate.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    fromDte = CDate(mskDate.Text & " " & mskTimeA.Text)
    'fromDte = DateAdd("n", -1, fromDte)
    toDte = CDate(mskDate.Text & " " & mskTimeB.Text)
    toDte = DateAdd("n", 1, toDte)
    Report.ParameterFields(1).AddCurrentValue (fromDte)
    Report.ParameterFields(2).AddCurrentValue (toDte)
    Report.ParameterFields(3).AddCurrentValue (lblUserID.Caption)
    Report.itxtTeller.SetText (lblUserID.Caption)
    Report.itxtDate.SetText (mskDate.Text)
    Report.itxtTimeA.SetText (mskTimeA.Text)
    Report.itxtTimeB.SetText (mskTimeB.Text)
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
    CRViewer1.Visible = True
    cmdCancel.Enabled = True
End Sub

Private Function ValidReport() As Boolean
    ValidReport = True
    If Not IsDate(mskDate & " " & mskTimeA) Or Not IsDate(mskDate & " " & mskTimeA) Then
        ValidReport = False
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Msg As String
    
    Msg = "Do you want to exit the program?"
    If MsgBox(Msg, vbQuestion + vbYesNo, "Exit") = vbNo Then Cancel = True
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mskDate_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub mskDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        mskTimeA.SetFocus
    End If
End Sub

Private Sub mskTimeA_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub mskTimeA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        mskTimeB.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        mskDate.SetFocus
    End If
End Sub

Private Sub mskTimeB_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub mskTimeB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        cmdView.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        mskTimeA.SetFocus
    End If
End Sub
