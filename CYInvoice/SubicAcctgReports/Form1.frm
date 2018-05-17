VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim c As cSubicRpt
Dim gconnstr As String

    gconnstr = "Provider=sqloledb" & _
                   ";Data Source=" & "SBITCBilling" & _
                   ";Initial Catalog=" & "BILLING" & _
                   ";Integrated Security=SSPI"
    
    Set c = New cSubicRpt
    With c
        .UserID = "borillano"
        .ConnectByStr (gconnstr)
        .Execute
        .Disconnect
    End With
    Set c = Nothing
End Sub

