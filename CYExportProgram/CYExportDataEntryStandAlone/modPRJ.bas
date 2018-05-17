Attribute VB_Name = "modPRJ"
Option Explicit
Public gConnStr As String


Dim sqlConBilling As String
Dim sqlConNavis As String

Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Public Sub Main()
    Dim mp As clsCCRde06
    'sharon
    
    Call ReadConfig
    
    Call gzCurrentUser
    gConnStr = sqlConBilling
    
'
'    "Provider=sqloledb" & _
'        ";Data Source=sbitcbilling" & _
'        ";Initial Catalog=BILLING" & _
'        ";Integrated Security=SSPI"
'        ";UID=tosadmin; password=password"


         gConnStr = "Provider=sqloledb" & _
        ";Data Source=sbitc-dev" & _
        ";Initial Catalog=sbitcbilling" & _
        ";UID=sa_ictsi; password=Ictsi123"
        
       '";Integrated Security=SSPI"
    Set mp = New clsCCRde06
    mp.Userid = "HSISON" ' '"jsapinoso"
    mp.ConnectByStr gConnStr
    mp.Execute
    mp.Disconnect
    Set mp = Nothing
End Sub


Public Function gzCurrentUser() As String
Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        gzCurrentUser = ""
    Else
        gzCurrentUser = Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    End If
End Function

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
