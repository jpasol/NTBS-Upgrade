VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3285
   ClientTop       =   3240
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim x As New ClsCommon
    Dim gConnStr As String
    
    gConnStr = "Provider=sqloledb;Data Source=sbitcbilling;Initial Catalog=BILLING;Integrated Security=SSPI"
    x.ConnectByStr (gConnStr)
    x.Userid = "borillano"
    x.Execute
    x.Disconnect
End Sub
