Attribute VB_Name = "modCommon"
Option Explicit

Public gcnnBilling As ADODB.Connection
Public gUserID As String
Public gPassword As String
Public gbConnected As Boolean

Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const MF_BYPOSITION = &H400&

Public Function gzCurrentUser() As String
Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        gzCurrentUser = ""
    Else
        gzCurrentUser = UCase(Left(lpUserName, InStr(lpUserName, Chr(0)) - 1))
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

Public Sub Main()
    gUserID = ""
    gPassword = ""
    gbConnected = False
End Sub

Public Sub RemoveCancelMenuItem(frm As Form)
    Dim hSysMenu As Long
    'get the system menu for this form
    hSysMenu = GetSystemMenu(frm.hWnd, 0)
    'remove the close item
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    'remove the separator that was over the close item
    Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

'--------------------------------------------------------------------
' Function      :   gzGetSysDate()
' Parameters    :   none
' Returns       :   DateTime    -> Server Date and Time
'--------------------------------------------------------------------
Public Function gzGetSysDate() As Date
Dim cmdGetSysDate As ADODB.Command
Dim prmGetSysDate As ADODB.Parameter
    
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
