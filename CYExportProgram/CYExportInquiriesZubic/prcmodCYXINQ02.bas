Attribute VB_Name = "prcmodCYXINQ02"
Option Explicit
Private Sub main()
    Dim xp As clsCYXINQ01
    Set xp = New clsCYXINQ01
'    xp.Userid = "RGLEAN"
    xp.Userid = "borillano"
    xp.Execute
    Set xp = Nothing
End Sub
