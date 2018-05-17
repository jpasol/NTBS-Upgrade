Attribute VB_Name = "modTestReprint"
Dim gConnStr As String

Private Sub Main()
Dim c As clsCYMPR01

    gConnStr = "Provider=sqloledb" & _
                   ";Data Source=" & "SBITCBILLING" & _
                   ";Initial Catalog=" & "BILLING" & _
                   ";Integrated Security=SSPI"
    
    Set c = New clsCYMPR01
    With c
        '.UserID = "BORILLANO"
        .ConnectByStr gConnStr
        .Execute (0)
        '.Disconnect
    End With
    Set c = Nothing
End Sub

