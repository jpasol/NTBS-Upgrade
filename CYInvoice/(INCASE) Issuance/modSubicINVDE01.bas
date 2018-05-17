Attribute VB_Name = "modSubicINVDE01"
Option Explicit

Public gcnnBilling As ADODB.Connection
' Public Variables
Public gUserID As String * 10
Public gPassword As String * 10
Public gConnStr As String
Public gbConnected As Boolean
Public gsCaption As String
' INI variables
Public gINIServer As String * 20
Public gINIDatabase As String * 20
' API Declares
Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

' Type declarations
Public Type tVslInfo
    regnum As String * 12       '  registry number
    vstnum As Long              '  visit id
    vslcde As String * 7        '  vessel code
    voyage As String * 12       '  voyage number
    lstdch As Date              '  last discharge date
End Type
Public Type tCustInfo
    cuscde As String * 6        '  customer code
    custyp As String * 3        '  customer type
    cusnam As String * 40       '  customer name
    careof As String * 40       '  agent
    adress As String * 40       '  customer address
    telfax As String * 30       '  telephone/fax
End Type

Public Function zCurrentUser() As String
Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        zCurrentUser = ""
    Else
        zCurrentUser = Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    End If
End Function

Public Function zCurrentComputer() As String
Dim lSize As Long
Dim sBuffer As String
    sBuffer = Space$(15& + 1)
    lSize = Len(sBuffer)
    zCurrentComputer = ""
    If GetComputerName(sBuffer, lSize) Then
        zCurrentComputer = Left$(sBuffer, lSize)
    End If
End Function
'--------------------------------------------------------------------
' Function      :   gzGetVesselInfo()
' Parameters    :   pRegNum     -> registry number (string * 12)
' Returns       :   tVslInfo    -> vessel information (see type declaration)
'--------------------------------------------------------------------
Public Function gzGetVesselInfo(ByVal pRegNo As String) As tVslInfo
Dim cmdGetVessel As ADODB.Command
Dim prmGetVessel As ADODB.Parameter
    
    ' create command
    Set cmdGetVessel = New ADODB.Command
    With cmdGetVessel
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getvesselinfo"
        .CommandType = adCmdStoredProc
         ' set parameters then execute
        Set prmGetVessel = New ADODB.Parameter
        Set prmGetVessel = .CreateParameter(, adChar, adParamInput, 12, pRegNo)
        .Parameters.Append prmGetVessel
        Set prmGetVessel = .CreateParameter("pVISIT", adNumeric, adParamOutput)
        .Parameters.Append prmGetVessel
        Set prmGetVessel = .CreateParameter("pVSLCDE", adChar, adParamOutput, 7)
        .Parameters.Append prmGetVessel
        Set prmGetVessel = .CreateParameter("pVOYAGE", adChar, adParamOutput, 12)
        .Parameters.Append prmGetVessel
        Set prmGetVessel = .CreateParameter("pLSTDCH", adDate, adParamOutput)
        .Parameters.Append prmGetVessel
        .Execute
    End With
    With gzGetVesselInfo
        On Error Resume Next
        .regnum = pRegNo
        .vstnum = "" & cmdGetVessel.Parameters("pVISIT")
        .vslcde = "" & cmdGetVessel.Parameters("pVSLCDE")
        .voyage = "" & cmdGetVessel.Parameters("pVOYAGE")
        .lstdch = "" & cmdGetVessel.Parameters("pLSTDCH")
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
'------------------------------------------------------------------------
' Function      :   gzGetRefNum()
' Parameters    :   pUserID     -> control type
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
