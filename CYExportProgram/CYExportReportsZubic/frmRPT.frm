VERSION 5.00
Begin VB.Form frmRPT 
   Caption         =   "( SBMA SUBIC - ZCCCRRPT ) CY Export Report Generation"
   ClientHeight    =   11025
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin zcCCRRpt.usrctrlCYEXRPT usrctrlCYEXRPT1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   19711
   End
End
Attribute VB_Name = "frmRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    usrctrlCYEXRPT1.StartInitialize
End Sub
Private Sub usrctrlCYEXRPT1_Closing()
    Unload Me
End Sub

