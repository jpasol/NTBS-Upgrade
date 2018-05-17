Attribute VB_Name = "modTestReport"
Dim gConnStr As String

Private Sub Main()

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

