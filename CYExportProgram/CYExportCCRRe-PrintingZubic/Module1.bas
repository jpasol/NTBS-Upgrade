Attribute VB_Name = "Module1"
Private Sub Main()
    Dim xp As Object
    Set xp = CreateObject("CCRCYREPRT.clsCYEXCCR")
    xp.Userid = "mpluelong"
    xp.Execute
    Set xp = Nothing
End Sub
