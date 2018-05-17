Attribute VB_Name = "zprcCYEXRPT"
Option Explicit
'Private Sub Main()
'    Dim pprc As Object
'    Set pprc = CreateObject("cCCRRPT.clsCCRRpt")
'    pprc.Userid = "MPUELONG"
'    pprc.Execute
'    Set pprc = Nothing
'End Sub


Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function zCurrentUser() As String
Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        zCurrentUser = ""
    Else
        zCurrentUser = UCase(Trim(Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)))
    End If
End Function
Public Function zCurrentComputer() As String
Dim lSize As Long
Dim sBuffer As String
    sBuffer = Space$(15& + 1)
    lSize = Len(sBuffer)
    zCurrentComputer = ""
    If GetComputerName(sBuffer, lSize) Then
        zCurrentComputer = UCase(Trim(Left$(sBuffer, lSize)))
    End If
End Function
Private Sub Main()
    Dim pprc As Object
    Set pprc = CreateObject("zcCCRRPT.zclsCCRRpt")
    pprc.Userid = zCurrentUser
    pprc.Execute
    Set pprc = Nothing
End Sub


