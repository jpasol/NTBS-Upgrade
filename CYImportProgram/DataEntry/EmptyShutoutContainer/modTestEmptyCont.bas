Attribute VB_Name = "modTestEmptyCont"
Option Explicit
Dim gConnStr As String
Private Sub main()
   Dim MS As Object
   gConnStr = "Provider=sqloledb" & _
        ";Data Source=" & Trim("SBITCBILLING") & _
        ";Initial Catalog=" & Trim("BILLING") & _
        ";Integrated Security=SSPI"
    Set MS = CreateObject("prjEmptyCont.clsCYEDE01")
    'MS.Userid = "borillano"
    MS.ConnectByStr gConnStr
    MS.Execute "borillano"
    MS.Disconnect
    Set MS = Nothing
End Sub
