Attribute VB_Name = "modManifest"
Option Explicit

Dim gConnStr As String

Private Sub Main()
    Dim objManifest As prjManifestDE.clsManifestDE
      
    'Set objManifest = New prjManifestDE.clsManifestDE
    
    Set objManifest = CreateObject("prjManifestDE.clsManifestDE")
    
    gConnStr = "Provider=sqloledb" & _
        ";Data Source=" & Trim("mis8bgr") & _
        ";Initial Catalog=" & Trim("BILLING") & _
        ";Integrated Security=SSPI"
        
    With objManifest
        .ConnectByStr gConnStr
        .Execute
        .Disconnect
    End With
    
    Set objManifest = Nothing
End Sub
