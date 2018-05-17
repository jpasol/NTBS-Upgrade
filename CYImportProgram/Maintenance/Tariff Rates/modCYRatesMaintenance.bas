Attribute VB_Name = "modCYRatesMaintenance"
Option Explicit

Public gUserid As String * 10
Public strCnn As String
Public cnnConnect As ADODB.Connection

Public Sub ConnectToServer()
    strCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI" & _
             ";Persist Security Info=False;Initial Catalog=billing;Data Source=BILLING_NT"
    Set cnnConnect = New ADODB.Connection
    cnnConnect.Open strCnn
End Sub
'--------------------------------------------------------------------
' Function      :   gzGetSysDate()
' Parameters    :   none
' Returns       :   DateTime    -> Server Date and Time
'--------------------------------------------------------------------
Public Function gzGetSysDate() As Date
Dim cmdGetSysDate As ADODB.Command
Dim prmGetSysDate As ADODB.Parameter
Dim X As Date
    
    ' create command
    Set cmdGetSysDate = New ADODB.Command
    Set prmGetSysDate = New ADODB.Parameter
    With cmdGetSysDate
        Set .ActiveConnection = cnnConnect
        .CommandText = "up_getsysdate"
        .CommandType = adCmdStoredProc
        Set prmGetSysDate = .CreateParameter("pDATE", adDate, adParamOutput)
        .Parameters.Append prmGetSysDate
        .Execute
        gzGetSysDate = .Parameters("pDATE")
    End With
End Function
