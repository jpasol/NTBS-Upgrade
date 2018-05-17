Attribute VB_Name = "modCaller"
Option Explicit

Sub Main()
Dim gConnStr As String
Dim gN4IP As String
Dim c As Object

    'gConnStr = "Provider=sqloledb; Data Source=itss01adc; Initial Catalog=sbitcBILLING; Integrated Security=SSPI"
'  gConnStr = "Provider=sqloledb" & _
'        ";Data Source=sbitc-dev" & _
'        ";Initial Catalog=sbitcbilling" & _
'        ";User Id=tosadmin;Password=password;"

        
        
  gConnStr = "Provider=sqloledb" & _
        ";Data Source=SBITC-DEV" & _
        ";Initial Catalog=sbitcbilling" & _
        ";User Id=SA_ICTSI;Password=Ictsi123;"
        
'   gConnStr = "Provider=sqloledb" & _
'        ";Data Source=sbitcbilling" & _
'        ";Initial Catalog=billing" & _
'        ";Integrated Security=SSPI"
        
        
    'Set MS = CreateObject("prjManifestCont.clsCYMDE01")
    Set c = CreateObject("SubicCYSCCR.cCYSCCR")
    With c
'        Call .ConnectByStr(gConnStr, "glacorte")
        Call .ConnectByStr(gConnStr, "HSISON")
        Call .CCRSuper("aherrera")
        Call .Execute
        Call .Disconnect
    End With
    Set c = Nothing
End Sub
