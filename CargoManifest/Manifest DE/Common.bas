Attribute VB_Name = "modCommon"
Option Explicit

Public gcnnBilling As ADODB.Connection
' Public Variables
Public gUserID As String * 10
Public gPassword As String * 10
Public gConnStr As String
Public gbConnected As Boolean
Public gsCaption As String
Public gbSupervisor As String

' INI variables
Const cINIFile = "NTBS.INI"
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


Public Sub zGetINIVal()
'Added by LAT
Dim n, i As Integer
Dim f, s, s1, s2 As String
    
    gINIServer = "LOCAL"
    gINIDatabase = "DUMMY"
    
    n = FreeFile
    f = App.Path & "\" & cINIFile
    On Error GoTo err_INI
    Open f For Input As #n
    While Not EOF(n)
        Line Input #n, s
            i = InStr(1, s, "=")
            If i > 0 Then
                s1 = Trim(Left(s, i - 1))
                s2 = Trim(Mid(s, i + 1))
                Select Case s1
                    Case "SERVER"
                        gINIServer = Trim(s2)
                    Case "DATABASE"
                        gINIDatabase = s2
                    Case "BACKGROUND"
                        gBackground = s2
                End Select
            End If
    Wend
    Close #n
    Exit Sub
err_INI:
    MsgBox "Cannot read INI file", vbCritical
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
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getsysdate"
        .CommandType = adCmdStoredProc
        Set prmGetSysDate = .CreateParameter("pDATE", adDate, adParamOutput)
        .Parameters.Append prmGetSysDate
        .Execute
        gzGetSysDate = .Parameters("pDATE")
    End With
End Function
