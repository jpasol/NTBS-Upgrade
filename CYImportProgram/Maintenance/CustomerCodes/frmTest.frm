VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x As Object
    Dim gConnStr As String
    
    Set x = CreateObject("CustomerMaintenance.clsCustMaintenance")
    gConnStr = "Provider=sqloledb;Data Source=SBITCBILLING;Initial Catalog=BILLING;Integrated Security=SSPI"
    x.ConnectByStr (gConnStr)
    x.Userid = "BROSALES"
    x.Execute
    x.Disconnect
End Sub
