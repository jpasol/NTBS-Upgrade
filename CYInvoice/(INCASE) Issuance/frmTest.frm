VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim x As New clsSubicINVDE01
    Dim gConnStr As String
    
    gConnStr = "Provider=sqloledb;Data Source=SBITCBilling;Initial Catalog=BILLING;Integrated Security=SSPI"
    x.ConnectByStr (gConnStr)
    x.Userid = "borillano"
    x.Execute
    x.Disconnect
    
End Sub




