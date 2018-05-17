VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReprintCCR 
   Caption         =   "CY Specical Services CCR Reprinting"
   ClientHeight    =   2385
   ClientLeft      =   4275
   ClientTop       =   3525
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReprintCCR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   4980
   Begin VB.CommandButton cmdReprint 
      Caption         =   "Re&print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   375
      TabIndex        =   2
      Top             =   1575
      Width           =   2265
   End
   Begin MSMask.MaskEdBox txtCCRRefNo 
      Height          =   465
      Left            =   2925
      TabIndex        =   0
      Top             =   300
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   820
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtCCRNum 
      Height          =   465
      Left            =   2925
      TabIndex        =   1
      ToolTipText     =   " Leave blank for all CCRs "
      Top             =   900
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   820
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99999999"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "CCR Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   390
      Left            =   375
      TabIndex        =   6
      Top             =   900
      Width           =   2340
   End
   Begin VB.Label lblExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   420
      Index           =   1
      Left            =   3675
      TabIndex        =   5
      Top             =   1650
      Width           =   1005
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Top             =   1650
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Reference No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   390
      Left            =   375
      TabIndex        =   3
      Top             =   300
      Width           =   2340
   End
End
Attribute VB_Name = "frmReprintCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsCCRReprint As Object

Private Sub cmdReprint_Click()
    With clsCCRReprint
'        On Error GoTo err_Reprint
        .CCRSupervisor vSupervisor
        .CCRNumber = CLng("0" & txtCCRNum)
        .PrintCCR CLng("0" & Trim(txtCCRRefNo))
        txtCCRRefNo = Space(txtCCRRefNo.MaxLength)
        txtCCRNum = Space(txtCCRNum.MaxLength)
    End With
'    Exit Sub
'err_Reprint:
'    On Error GoTo 0
'    MsgBox "Reference/CCR number " & Trim(txtCCRRefNo) & " not found", vbInformation
'    txtCCRRefNo.SetFocus
End Sub

Private Sub cmdReprint_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            SendKeys "{TAB}"
        Case Else
    End Select
End Sub

Private Sub Form_Load()
    Set clsCCRReprint = CreateObject("CCRPR03.clsCCRPR03")
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Set clsCCRReprint = Nothing
End Sub

Private Sub txtCCRNum_GotFocus()
    With txtCCRNum
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCCRNum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            SendKeys "{TAB}"
        Case Else
    End Select
End Sub

Private Sub txtCCRNum_LostFocus()
    With txtCCRNum
        .BackColor = vbWindowBackground
    End With
End Sub

Private Sub txtCCRRefNo_Change()
    cmdReprint.Enabled = (CLng("0" & txtCCRRefNo) > 0)
End Sub

Private Sub txtCCRRefNo_GotFocus()
    With txtCCRRefNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCCRRefNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            SendKeys "{TAB}"
        Case Else
    End Select
End Sub

Private Sub txtCCRRefNo_LostFocus()
    With txtCCRRefNo
        .BackColor = vbWindowBackground
    End With
End Sub
