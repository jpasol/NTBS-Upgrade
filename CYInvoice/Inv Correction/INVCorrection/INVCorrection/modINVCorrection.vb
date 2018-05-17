Option Explicit On 
Imports Microsoft.VisualBasic.Compatibility

Module modINVCorrection

    Public dsINVPayHdr As New DataSet
    Public dsINVPayDtl As New DataSet
    Public dsUser As New DataSet

    Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, ByRef lpnLength As Integer) As Integer

    Public Function zCurrentUser() As String
        Dim lpUserName As New VB6.FixedLengthString(64)
        If WNetGetUser("", lpUserName.Value, Len(lpUserName.Value)) Then
            zCurrentUser = ""
        Else
            zCurrentUser = UCase(Trim(Left(lpUserName.Value, InStr(lpUserName.Value, Chr(0)) - 1)))
        End If
    End Function

End Module
