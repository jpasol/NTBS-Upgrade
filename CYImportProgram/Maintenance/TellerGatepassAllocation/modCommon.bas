Attribute VB_Name = "modCommon"
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
' Function      :   gzChkCCRExists()
' Parameters    :   Start CCR, End CCR
' Returns       :   boolean (T - CCR cannot be allocated, F - CCR can be allocated)
'--------------------------------------------------------------------
Public Function gzChkCCRExists(ByVal pFrom As Long, pTo As Long) As Long
Dim cmdGetValidCCR As ADODB.Command
Dim prmGetValidCCR As ADODB.Parameter
    
    ' create command
    Set cmdGetValidCCR = New ADODB.Command
    With cmdGetValidCCR
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_ccrexists"
        .CommandType = adCmdStoredProc
         ' set parameters then execute
        Set prmGetValidCCR = New ADODB.Parameter
        Set prmGetValidCCR = .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append prmGetValidCCR
        Set prmGetValidCCR = .CreateParameter(, adInteger, adParamInput, 8, pFrom)
        .Parameters.Append prmGetValidCCR
        Set prmGetValidCCR = .CreateParameter(, adInteger, adParamInput, 8, pTo)
        .Parameters.Append prmGetValidCCR
        .Execute
        gzChkCCRExists = .Parameters("RETURN_VALUE")
    End With
End Function
'--------------------------------------------------------------------
' Function      :   gzChkGPSExists()
' Parameters    :   Start CCR, End CCR
' Returns       :   boolean (T - CCR cannot be allocated, F - CCR can be allocated)
'--------------------------------------------------------------------
Public Function gzChkGPSExists(ByVal pFrom As Long, pTo As Long, pTyp As Integer) As Long
Dim cmdGetValidGPS As ADODB.Command
Dim prmGetValidGPS As ADODB.Parameter
    
    ' create command
    Set cmdGetValidGPS = New ADODB.Command
    With cmdGetValidGPS
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_gpassexists"
        .CommandType = adCmdStoredProc
         ' set parameters then execute
        Set prmGetValidGPS = New ADODB.Parameter
        Set prmGetValidGPS = .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append prmGetValidGPS
        Set prmGetValidGPS = .CreateParameter(, adInteger, adParamInput, 8, pFrom)
        .Parameters.Append prmGetValidGPS
        Set prmGetValidGPS = .CreateParameter(, adInteger, adParamInput, 8, pTo)
        .Parameters.Append prmGetValidGPS
        Set prmGetValidGPS = .CreateParameter(, adInteger, adParamInput, 1, pTyp)
        .Parameters.Append prmGetValidGPS
        .Execute
        gzChkGPSExists = .Parameters("RETURN_VALUE")
    End With
End Function
'--------------------------------------------------------------------
' Function      :   gzChkUserInfo()
' Parameters    :   txtTeller(index)
' Returns       :   1 if user exists; else 0
'--------------------------------------------------------------------
Public Function gzChkUserInfo(ByVal pUserID As String) As Boolean
Dim cmdGetUserInfo As ADODB.Command
Dim prmGetUserInfo As ADODB.Parameter
    
    ' create command
    Set cmdGetUserInfo = New ADODB.Command
    With cmdGetUserInfo
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_chkuserinfo"
        .CommandType = adCmdStoredProc
         ' set parameters then execute
        Set prmGetUserInfo = New ADODB.Parameter
        Set prmGetUserInfo = .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append prmGetUserInfo
        Set prmGetUserInfo = .CreateParameter(, adChar, adParamInput, 10, pUserID)
        .Parameters.Append prmGetUserInfo
        .Execute
        gzChkUserInfo = .Parameters("RETURN_VALUE")
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
