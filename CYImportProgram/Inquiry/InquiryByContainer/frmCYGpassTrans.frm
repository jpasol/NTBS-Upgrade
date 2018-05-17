VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCYGpassTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CY Import Billing Gatepass Transactions for a Container"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
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
   ScaleHeight     =   10680
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   360
      TabIndex        =   5
      Top             =   9240
      Width           =   14655
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   735
         Left            =   6720
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   735
         Left            =   11880
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   735
         Left            =   9240
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8055
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   14655
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   7125
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   14235
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   0   'False
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   0   'False
         EnableProgressControl=   -1  'True
         EnableSearchControl=   0   'False
         EnableRefreshButton=   0   'False
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   0   'False
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   14655
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   12480
         Top             =   360
      End
      Begin MSMask.MaskEdBox mskCntNum 
         Height          =   375
         Left            =   3360
         TabIndex        =   0
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   12
         Mask            =   ">????-#######"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMessage 
         Height          =   375
         Left            =   7800
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Container Number "
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmCYGpassTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Private Sub cmdCancel_Click()
    CRViewer1.Visible = False
    cmdCancel.Enabled = False
    mskCntNum.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    Dim Report As New CrystalReport1
    Dim prmCntNum As String
    
    If InStr(mskCntNum, "_") > 0 Then
        MsgBox "Please specify a valid container number.", _
        vbInformation, "Error Message"
        mskCntNum.SetFocus
        Exit Sub
    End If
    prmCntNum = CStr(Left(mskCntNum, 4) & Right(mskCntNum, 7))
    Report.ParameterFields(1).AddCurrentValue (Trim(prmCntNum))
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    CRViewer1.Visible = True
    cmdCancel.Enabled = True
End Sub

Private Function ValidGPS(pGpass As String) As Boolean
    ValidGPS = False
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msg As String
    msg = "Do you want to exit the program?"
    If MsgBox(msg, vbQuestion + vbYesNo, "Exit") = vbNo Then Cancel = True
End Sub

Private Sub mskCntNum_KeyDown(KeyCode As Integer, Shift As Integer)
    lblMessage.Visible = False
End Sub
