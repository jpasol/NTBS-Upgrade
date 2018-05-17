VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Tariff Rates Maintenance"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "C L O S E"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "L O G I N "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gConnStr As String

Dim sqlConBilling As String
Dim sqlConNavis As String

Function ConnectToServer() As Boolean
   
    Dim x As Object
    
    Call ReadConfig
    gConnStr = sqlConBilling
    
'    gConnStr = "Provider=sqloledb" & _
'                ";Data Source=" & "MDC-VIRTUAL\MDCDEV" & _
'                ";Initial Catalog=" & "sbitc_billing" & _
'                ";Integrated Security=SSPI"
    
    Set x = CreateObject("CYRatesMaintenance.clsCYRates")
    With x
        .ConnectByStr (gConnStr)
        .Execute
        .Disconnect
    End With
    Set x = Nothing
End Function

'MDC (20131209)
'Exclude connection string on source code

Public Sub ReadConfig()
Dim Xcnt As Integer
Open App.Path & "\" & "Conn.cfg" For Binary Access Read As #1

Do While Not EOF(1)
    Xcnt = Xcnt + 1
    Select Case Xcnt
        Case 1
            Line Input #1, sqlConBilling
        Case 2
            Line Input #1, sqlConNavis
    End Select
Loop
End Sub


Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdLogin_Click()
    Call ConnectToServer
End Sub

