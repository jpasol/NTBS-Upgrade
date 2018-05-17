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
'    Dim x As New clsCCRAllocation
'    Dim gConnStr As String
'    gConnStr = "Provider=sqloledb;Data Source=BILLING_NT;Initial Catalog=BILLING;Integrated Security=SSPI"
'
'    x.ConnectByStr (gConnStr)
'    x.Userid = "RJTAMSI"
'    x.Execute
'    x.Disconnect

Dim cALLOC As Object
Dim gConnStr As String
    
    gConnStr = "Provider=sqloledb;Data Source=SBITCBILLING;Initial Catalog=BILLING;Integrated Security=SSPI"

    Set cALLOC = CreateObject("CCRAllocation.clsCCRAllocation")
    With cALLOC
        .ConnectByStr (gConnStr)
        .Userid = "RJTAMSI"
        .Execute
        .Disconnect
    End With
    Set cALLOC = Nothing
    
End Sub

