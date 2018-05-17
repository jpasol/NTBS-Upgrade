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
Public sqlConBilling As String
Public sqlConNavis As String
Public gConnStr As String




Private Sub Command1_Click()

    Dim x As New clsSubicINVDE01
      
      'added by MDC (20131206)
      'move connection string to conn.cfg text file for deployment of system
    Call ReadConfig
    
    gConnStr = sqlConBilling '"Provider=sqloledb;Data Source=SBITCBilling;Initial Catalog=BILLING;Integrated Security=SSPI"
    x.ConnectByStr (gConnStr)
    x.Userid = UCase(Environ("USERNAME")) 'zCurrentUser '"borillano"
    x.Execute
    x.Disconnect
    
End Sub

'Public Function zCurrentUser() As String
'Dim lpUserName As String * 64
'    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
'        zCurrentUser = ""
'    Else
'        zCurrentUser = Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)
'    End If
'End Function




Public Sub ReadConfig()
Dim Xcnt As Integer
Dim Udefa As String
Open App.Path & "\" & "Conn.cfg" For Binary Access Read As #1 ' Open file.

Do While Not EOF(1) ' Loop until end of file.
   Xcnt = Xcnt + 1
   Select Case Xcnt
      Case 1
        Line Input #1, sqlConBilling   ' Read line into variable.
      Case 2
        Line Input #1, sqlConNavis
   End Select
Loop
  sqlConBilling = Trim(sqlConBilling)
  sqlConNavis = Trim(sqlConNavis)
Close #1   ' Close file.
End Sub
