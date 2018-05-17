Attribute VB_Name = "modCommon"
Option Explicit

Const cINIFile = "NTBS.INI"

Public gcnnBilling As ADODB.Connection

' Public Constants
Public Const vbWindowFocus = &H80000018
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TOOLWINDOW = &H80&

' Public Variables
Public gUserID As String * 10
Public gPassword As String * 10
Public gConnStr As String
Public gbConnected As Boolean
Public gbSupervisor As String

Public gComputer As String * 20
Public gBackground As String
Public gShutDown As Boolean

' INI variables
Public gINIServer As String * 20
Public gINIDatabase As String * 20

' API Declares
Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Function zCurrentUser() As String
Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        zCurrentUser = ""
    Else
        zCurrentUser = UCase(Trim(left(lpUserName, InStr(lpUserName, Chr(0)) - 1)))
    End If
End Function

Public Function zCurrentComputer() As String
Dim lSize As Long
Dim sBuffer As String
    sBuffer = Space$(15& + 1)
    lSize = Len(sBuffer)
    zCurrentComputer = ""
    If GetComputerName(sBuffer, lSize) Then
        zCurrentComputer = UCase(Trim(left$(sBuffer, lSize)))
    End If
End Function

Public Sub zGetINIVal(ByVal port As String)
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
                s1 = Trim(left(s, i - 1))
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

Public Function gzConnected(pServer As String, pDatabase As String, pTrusted As Byte, pUserID As String, pPassword As String) As Boolean
Dim errBilling As ADODB.Error
Dim lsErrStr As String

    If pTrusted = 0 Then
        gConnStr = "Provider=sqloledb" & _
                   ";Data Source=" & Trim(pServer) & _
                   ";Initial Catalog=" & Trim(pDatabase) & _
                   ";User Id=" & Trim(pUserID) & _
                   ";Password=" & Trim(pPassword)
        gUserID = UCase(pUserID)
        gPassword = pPassword
    Else
         gConnStr = "Provider=sqloledb" & _
                   ";Data Source=" & Trim(pServer) & _
                   ";Initial Catalog=" & Trim(pDatabase) & _
                   ";Integrated Security=SSPI"
        gUserID = zCurrentUser()
        gPassword = ""
    End If
    
    ' Open the database.
    On Error GoTo err_Connect
    Set gcnnBilling = New ADODB.Connection
    gcnnBilling.Open gConnStr
    gzConnected = True
    
    Exit Function
    
err_Connect:
    gzConnected = False
    For Each errBilling In gcnnBilling.Errors
        With errBilling
            lsErrStr = "Connection Error. " & .Description & vbLf & _
            "Verify Log On then retry.  Contact MIS for assistance."
        End With
        MsgBox lsErrStr, vbCritical
    Next
End Function

Public Sub zShowInTaskbar(Visible As Boolean, hwnd As Long)
Dim L As Long
    L = ShowWindow(hwnd, SW_HIDE)
    DoEvents
    L = SetWindowLong(hwnd, GWL_EXSTYLE, IIf(Visible, -WS_EX_TOOLWINDOW, WS_EX_TOOLWINDOW))
    DoEvents
    L = ShowWindow(hwnd, SW_SHOW)
End Sub

Public Sub gzEcho(s As String)
    frmMain.lblStatusMsg = " " & s
    frmMain.lblStatusMsg.Refresh
End Sub
