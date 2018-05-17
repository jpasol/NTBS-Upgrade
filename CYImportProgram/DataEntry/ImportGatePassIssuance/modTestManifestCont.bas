Attribute VB_Name = "modTestManifestCont"
Option Explicit
Dim gConnStr As String
Dim gconnnav As String
        '";Data Source=SBITCBILLING"
Private Sub Main()
   Dim MS As Object

'PROD
'   gConnStr = "Provider=sqloledb" & _
'        ";Data Source=SBITCBILLING" & _
'        ";Initial Catalog=billing" & _
'        ";UID=sa_ictsi; password=Ictsi123"
'    Set MS = CreateObject("prjManifestCont.clsCYMDE01")

'TEST
    gConnStr = "Provider=sqloledb" & _
        ";Data Source=SBITC-DEV" & _
        ";Initial Catalog=sbitcbilling" & _
        ";UID=sa_ictsi; password=Ictsi123"
    Set MS = CreateObject("prjManifestCont.clsCYMDE01")

    
''PRNH - Test
'    gConnStr = "Provider=sqloledb" & _
'        ";Data Source=192.168.11.155" & _
'        ";Initial Catalog=SBITCbilling" & _
'        ";UID=sa_ictsi; password=Ictsi123"
'    Set MS = CreateObject("prjManifestCont.clsCYMDE01")
'
'       gConnStr = "Provider=sqloledb" & _
'        ";Data Source=gitmdc-l" & _
'        ";Initial Catalog=sbitc_billing" & _
'        ";UID=sa; password=p@ssw0rd"
'    Set MS = CreateObject("prjManifestCont.clsCYMDE01")

        '";UID=tosadmin; password=password"
    'MS.Userid = "borillano"
    
    MS.ConnectByStr (gConnStr)
    MS.Execute "AGALLARDO"
    MS.Disconnect
    Set MS = Nothing
End Sub
