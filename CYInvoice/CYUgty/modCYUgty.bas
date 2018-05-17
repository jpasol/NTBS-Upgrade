Attribute VB_Name = "modCYUgty"
Option Explicit

Public gcnnBilling As ADODB.Connection
Public gbConnected As Boolean

Public Type tCustInfo
    cuscde As String * 6        '  customer code
    custyp As String * 3        '  customer type
    cusnam As String * 40       '  customer name
    careof As String * 40       '  agent
    adress As String * 40       '  customer address
    telfax As String * 30       '  telephone/fax
End Type

Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Public Sub ConnectToBilling()
Dim gConnStr As String
    gConnStr = "Provider=SQLOLEDB;Data Source=SBITCBILLING;Initial Catalog=Billing;Integrated Security=SSPI"
    'gConnStr = "Provider=SQLOLEDB;Data Source=SBITC-DEV;Initial Catalog=billing-ntbs-11162017;Integrated Security=SSPI"
    Set gcnnBilling = New ADODB.Connection
    gcnnBilling.Open gConnStr
    gbConnected = True
End Sub

Public Function zCurrentUser() As String
Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        zCurrentUser = ""
    Else
        zCurrentUser = Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    End If
    zCurrentUser = UCase(Trim(zCurrentUser))
End Function
'------------------------------------------------------------------------
' Function      :   gzGetRefNum()
' Parameters    :   pCtlTyp     -> control type
' Returns       :   reference number
'------------------------------------------------------------------------
Public Function gzGetRefNum(ByVal pCtlTyp As String) As Long
Dim cmdGetRefNum As ADODB.Command
Dim prmGetRefNum As ADODB.Parameter

    ' create command
    Set cmdGetRefNum = New ADODB.Command
    With cmdGetRefNum
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcontrolno"
        .CommandType = adCmdStoredProc

        ' set parameters then execute
        Set prmGetRefNum = .CreateParameter(, adChar, adParamInput, 3, pCtlTyp)
        .Parameters.Append prmGetRefNum
        Set prmGetRefNum = .CreateParameter("pCTLNUM", adInteger, adParamOutput)
        .Parameters.Append prmGetRefNum
        .Execute
        gzGetRefNum = .Parameters("pCTLNUM")
    End With
End Function
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
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getsysdate"
        .CommandType = adCmdStoredProc
        Set prmGetSysDate = .CreateParameter("pDATE", adDate, adParamOutput)
        .Parameters.Append prmGetSysDate
        .Execute
        gzGetSysDate = .Parameters("pDATE")
    End With
End Function
'--------------------------------------------------------------------
' Function      :   gzGetCustomerInfo()
' Parameters    :   pCustCde     -> customer code (string * 6)
' Returns       :   tCustInfo    -> customer information (see type declaration)
'--------------------------------------------------------------------
Public Function gzGetCustomerInfo(ByVal pCustCde As String) As tCustInfo
Dim cmdGetCustomer As ADODB.Command
Dim prmGetCustomer As ADODB.Parameter
    
    ' create command
    Set cmdGetCustomer = New ADODB.Command
    With cmdGetCustomer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcustomerinfo"
        .CommandType = adCmdStoredProc
        ' set parameters then execute
        Set prmGetCustomer = New ADODB.Parameter
        Set prmGetCustomer = .CreateParameter(, adChar, adParamInput, 6, pCustCde)
        .Parameters.Append prmGetCustomer
        Set prmGetCustomer = .CreateParameter("pTYPE", adVarChar, adParamOutput, 3)
        .Parameters.Append prmGetCustomer
        Set prmGetCustomer = .CreateParameter("pNAME", adVarChar, adParamOutput, 40)
        .Parameters.Append prmGetCustomer
        Set prmGetCustomer = .CreateParameter("pCAREOF", adVarChar, adParamOutput, 40)
        .Parameters.Append prmGetCustomer
        Set prmGetCustomer = .CreateParameter("pADDRESS", adVarChar, adParamOutput, 40)
        .Parameters.Append prmGetCustomer
        Set prmGetCustomer = .CreateParameter("pTELFAX", adVarChar, adParamOutput, 30)
        .Parameters.Append prmGetCustomer
        .Execute
    End With
    With gzGetCustomerInfo
        .cuscde = pCustCde
        .custyp = "" & cmdGetCustomer.Parameters("pTYPE")
        .cusnam = "" & cmdGetCustomer.Parameters("pNAME")
        .careof = "" & cmdGetCustomer.Parameters("pCAREOF")
        .adress = "" & cmdGetCustomer.Parameters("pADDRESS")
        .telfax = "" & cmdGetCustomer.Parameters("pTELFAX")
    End With
End Function
