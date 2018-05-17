Attribute VB_Name = "zModule1"
Private Sub Main()
    Dim xp As Object
    Set xp = CreateObject("zCCRCYREPRT.zclsCYEXCCR")
    xp.Userid = "borillano"
    xp.Execute
    Set xp = Nothing
End Sub
