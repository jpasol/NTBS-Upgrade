VERSION 5.00
Begin VB.Form frmTesting 
   Caption         =   "Reprint"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCCRnum 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "102"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtRefnum 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "212621"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "frmTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CYCCr As New clsCCRPr03
Private Sub cmdPreview_Click()
    CYCCr.CCRSupervisor "Noel"
    CYCCr.CCRNumber = CLng(txtCCRnum.Text)
    CYCCr.PreviewCCR (CLng(txtRefnum.Text))
End Sub

Private Sub cmdPrint_Click()
    CYCCr.CCRSupervisor = "Noel"
    CYCCr.CCRNumber = CLng(txtCCRnum.Text)
    CYCCr.PrintCCR (CLng(txtRefnum.Text))
End Sub

Private Sub Form_Load()
txtCCRnum.Text = 0
txtRefnum.Text = 0
End Sub
