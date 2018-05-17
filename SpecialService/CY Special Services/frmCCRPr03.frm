VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Begin VB.Form frmCCRPr03 
   Caption         =   "Report View"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14160
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
   LockControls    =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13935
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
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
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   9120
      Width           =   2055
   End
End
Attribute VB_Name = "frmCCRPr03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Dim CCR As rptCCRPr03
    Set CCR = New rptCCRPr03
    CCR.ParameterFields(1).AddCurrentValue (RefNum)
    CCR.ParameterFields(2).AddCurrentValue (SeqNum)
    CCR.ParameterFields(3).AddCurrentValue (Customer)
    CCR.ParameterFields(4).AddCurrentValue (strAdrAmt)
    CCR.ParameterFields(5).AddCurrentValue (strCshAmt)
'    CCR.ParameterFields(6).AddCurrentValue (strChqAmt1)
'    CCR.ParameterFields(7).AddCurrentValue (strChqAmt2)
'    CCR.ParameterFields(8).AddCurrentValue (strChqAmt3)
'    CCR.ParameterFields(9).AddCurrentValue (strChqAmt4)
'    CCR.ParameterFields(10).AddCurrentValue (strChqAmt5)
    CCR.ParameterFields(6).AddCurrentValue (strChqAmt)
    CCR.ParameterFields(7).AddCurrentValue (blnChkno1)
    CCR.ParameterFields(8).AddCurrentValue (blnChkno2)
    CCR.ParameterFields(9).AddCurrentValue (blnChkno3)
    CCR.ParameterFields(10).AddCurrentValue (blnChkno4)
    CCR.ParameterFields(11).AddCurrentValue (blnChkno5)
    CCR.ParameterFields(12).AddCurrentValue (strSupervisor)
'    CCR.TxtSupervisor.SetText strSupervisor
'    crviewer1.DataSource
    CRViewer1.ReportSource = CCR
    CRViewer1.ViewReport
End Sub
