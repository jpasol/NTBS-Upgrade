VERSION 5.00
Begin VB.Form frmTestManifestDE 
   Caption         =   "Test"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Run Manifest Data Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
End
Attribute VB_Name = "frmTestManifestDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gConnStr As String

Private Sub Command1_Click()
    Dim objManifest As prjManifestDE.clsManifestDE
      
    Set objManifest = CreateObject("prjManifestDE.clsManifestDE")
    
    gConnStr = "Provider=sqloledb" & _
        ";Data Source=" & Trim("SBITCBILLING") & _
        ";Initial Catalog=" & Trim("BILLING") & _
        ";Integrated Security=SSPI"
        
    With objManifest
        .ConnectByStr gConnStr
        .Execute
        .Disconnect
    End With
    
    Set objManifest = Nothing
End Sub


