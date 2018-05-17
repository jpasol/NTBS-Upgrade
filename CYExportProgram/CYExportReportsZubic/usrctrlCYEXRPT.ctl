VERSION 5.00
Begin VB.UserControl usrctrlCYEXRPT 
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
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
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   Begin zcCCRRpt.prvusrctrlPlain prvusrctrlPlain2 
      Height          =   420
      Left            =   960
      TabIndex        =   3
      Top             =   2040
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Monthly Report"
   End
   Begin zcCCRRpt.prvusrctrlPlain Rpt7 
      Height          =   420
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Collection Summary"
   End
   Begin zcCCRRpt.prvusrctrlPlain Rpt6 
      Height          =   420
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Assessors Report(Auditors Copy)                                "
   End
   Begin zcCCRRpt.prvusrctrlPlain Rpt4 
      Height          =   420
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Assessors Collection Report       "
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "F4 - Change Printer"
      Height          =   660
      Left            =   5160
      Picture         =   "usrctrlCYEXRPT.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9600
      Width           =   4935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "F3 - E&xit"
      Height          =   615
      Left            =   240
      Picture         =   "usrctrlCYEXRPT.ctx":014A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9720
      Width           =   4575
   End
   Begin VB.Frame Proxy 
      Height          =   10815
      Left            =   5040
      TabIndex        =   12
      Top             =   0
      Width           =   10215
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   120
         TabIndex        =   16
         Top             =   9360
         Width           =   9975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CCRRPT"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   8280
         TabIndex        =   18
         Top             =   10440
         Width           =   1935
      End
      Begin VB.Label lblProxy 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   10215
      End
      Begin VB.Label lblPrinter 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   10440
         Width           =   8295
      End
   End
   Begin VB.CommandButton cmdA 
      Appearance      =   0  'Flat
      Caption         =   "(&A)                       "
      Height          =   420
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   4575
   End
   Begin VB.CommandButton cmdC 
      Appearance      =   0  'Flat
      Caption         =   "(&B)                       "
      Height          =   420
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton cmdD 
      Appearance      =   0  'Flat
      Caption         =   "(&C)                       "
      Height          =   420
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "(&D)                       "
      Height          =   420
      Left            =   240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Height          =   9375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4815
      Begin zcCCRRpt.prvusrctrlPlain Rpt8 
         Height          =   420
         Left            =   840
         TabIndex        =   20
         Top             =   2520
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Under Guarantee Report"
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "(&E)                       "
         Height          =   420
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2520
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CY Export Reports Menu"
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   9360
      Width           =   4815
      Begin VB.Label lblTeller 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblWorkstation 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   1080
         Width           =   2415
      End
   End
End
Attribute VB_Name = "usrctrlCYEXRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Event InMenu()
    Event InTab()
    Event Closing()
Public Sub StartInitialize()
    Dim rsUsr As Recordset
    VE.getInformation
    Set rsUsr = VE.rsgetInformation
        lblTeller.Caption = gUserid
    lblWorkstation.Caption = rsUsr.Fields("workstation")
    lblPrinter.Caption = "Printer Device : " & Printer.DeviceName
    rsUsr.Close
    Set rsUsr = Nothing
End Sub
Private Sub cmdExit_Click()
    RaiseEvent Closing
End Sub
Private Sub cmdPrinter_Click()
    frmPrinter.Show vbModal
    lblPrinter.Caption = "Printer Device :" & Printer.DeviceName
End Sub

Private Sub prvusrctrlPlain1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        frmLiquidPOS.top = lblProxy.top + 1310 ' 1430
        frmLiquidPOS.left = Proxy.left
        frmLiquidPOS.Show vbModal
    End If
End Sub

Private Sub prvusrctrlPlain2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        frmMonthly.top = lblProxy.top + 1310 ' 1430
        frmMonthly.left = Proxy.left
        frmMonthly.Show vbModal
    End If
End Sub

Private Sub prvusrctrlPlain3_Click()

End Sub

Private Sub Rpt4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        frmLiquid.top = lblProxy.top + 1310 ' 1430
        frmLiquid.left = Proxy.left
        frmLiquid.Show vbModal
    End If
End Sub
Private Sub Rpt4_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Rpt5_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Rpt6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        frmAuditor.top = lblProxy.top + 1310 ' 1430
        frmAuditor.left = Proxy.left
        frmAuditor.Show vbModal
    End If
End Sub
Private Sub Rpt6_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Rpt7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        frmSummary.top = lblProxy.top + 1310 ' 1430
        frmSummary.left = Proxy.left
        frmSummary.Show vbModal
    End If
End Sub
Private Sub Rpt7_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Rpt8_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        frmUGuarantee.top = lblProxy.top + 1310  ' 1430
        frmUGuarantee.left = Proxy.left
        frmUGuarantee.Show vbModal
    End If
End Sub

Private Sub Rpt8_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            RaiseEvent Closing
        Case vbKeyF4
            Call cmdPrinter_Click
    End Select
End Sub
