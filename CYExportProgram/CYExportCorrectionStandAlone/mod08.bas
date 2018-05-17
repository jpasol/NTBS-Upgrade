Attribute VB_Name = "mod08"
Option Explicit
Dim gConnStr As String
Private Sub main()
   Dim MS As Object
   gConnStr = "Provider=sqloledb" & _
        ";Data Source=" & Trim("sbitcbilling") & _
        ";Initial Catalog=" & Trim("BILLING") & _
        ";Integrated Security=SSPI"
    Set MS = CreateObject("CCRde08.clsCCRde08")
    MS.Userid = "borillano"
    MS.ConnectByStr gConnStr
    MS.Execute
    MS.Disconnect
    Set MS = Nothing
End Sub
