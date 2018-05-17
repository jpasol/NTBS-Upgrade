VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmReprint 
   Caption         =   "Reprint Gatepass"
   ClientHeight    =   9405
   ClientLeft      =   -1950
   ClientTop       =   510
   ClientWidth     =   15240
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptCYMPR01 
      Left            =   9360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      Begin VB.CommandButton cmdReprint 
         Caption         =   "Reprint"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4440
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mskReference 
         Height          =   405
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskSequence 
         Height          =   405
         Left            =   3600
         TabIndex        =   3
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label lblMain 
         Alignment       =   1  'Right Justify
         Caption         =   "Reference:"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblMain 
         Alignment       =   1  'Right Justify
         Caption         =   "Sequence:"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intResponse As Integer

Private Sub cmdReprint_Click()
    rptCYMPR01.ReportFileName = App.Path & "\cympr01.rpt"
    rptCYMPR01.SelectionFormula = "{CYMgps.refnum} = " & mskReference & " and {CYMgps.seqnum} = " & mskSequence
    rptCYMPR01.PrintReport
End Sub

Private Sub Form_Activate()
    mskReference.SetFocus
End Sub

Private Sub mskReference_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskReference, mskSequence)
End Sub

Private Sub mskSequence_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskReference, cmdReprint)
End Sub

Private Sub FieldAdvance(pKeyCode As Integer, pPreviousControl As Control, pNextControl As Control)
    Select Case pKeyCode
        Case vbKeyDown
            pNextControl.SetFocus
        Case vbKeyReturn
            pNextControl.SetFocus
        Case vbKeyUp
            pPreviousControl.SetFocus
         Case vbKeyF3
            intResponse = MsgBox("Do you really want to Exit?", vbYesNo + vbCritical, "Quit Program")
            If intResponse = vbYes Then
                Unload Me
            End If
    End Select
End Sub
