VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNextCCR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Next CCR"
   ClientHeight    =   1890
   ClientLeft      =   4410
   ClientTop       =   4485
   ClientWidth     =   4005
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   540
      Left            =   2100
      TabIndex        =   3
      Top             =   1050
      Width           =   1740
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   540
      Left            =   225
      TabIndex        =   2
      Top             =   1050
      Width           =   1740
   End
   Begin MSMask.MaskEdBox txtCCRNum 
      Height          =   465
      Left            =   2250
      TabIndex        =   1
      Top             =   300
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Caption         =   "CCR Number"
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   375
      Width           =   1890
   End
End
Attribute VB_Name = "frmNextCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    vNextCCR = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    vNextCCR = CLng(txtCCRNum)
    Unload Me
End Sub

Private Sub Form_Load()
    txtCCRNum = Format(vNextCCR, txtCCRNum.Mask)
End Sub

Private Sub txtCCRNum_Change()
    cmdOK.Enabled = Trim(txtCCRNum) <> ""
End Sub

Private Sub txtCCRNum_GotFocus()
    With txtCCRNum
        .SelStart = 0
        .SelLength = Len(Trim(txtCCRNum))
        .BackColor = vbInfoBackground
    End With
End Sub

Private Sub txtCCRNum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            vNextCCR = CLng(txtCCRNum)
            Unload Me
        Case vbKeyEscape
            vNextCCR = 0
            Unload Me
        Case Else
    End Select
End Sub

Private Sub txtCCRNum_LostFocus()
    With txtCCRNum
        .Text = Trim(txtCCRNum)
        .BackColor = vbWindowBackground
    End With
End Sub
