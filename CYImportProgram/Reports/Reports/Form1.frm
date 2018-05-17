VERSION 5.00
Begin VB.Form Form1 
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
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gConnStr As String



Private Sub Command1_Click()
Dim c As cCYSRPT

    gConnStr = "Provider=sqloledb" & _
                   ";Data Source=" & "SBITCBILLING" & _
                   ";Initial Catalog=" & "BILLING" & _
                   ";Integrated Security=SSPI"
    
    Set c = New cCYSRPT
    With c
        .UserID = "BORILLANO"
        '.ConnectByStr gConnStr
        .Execute (0)
        '.Disconnect
    End With
    Set c = Nothing
End Sub
