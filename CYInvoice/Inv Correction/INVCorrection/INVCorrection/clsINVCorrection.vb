Option Explicit On 
Imports System.Data
Imports System.Data.SqlClient

Public Class clsINVCorrection

    'Connection String for Billing Database    
    Private ConnStr As String = _
        "Initial Catalog=Billing;" & _
        "Data Source=MIS8BGR;" & _
        "integrated security=sspi"

    Public Function RetriveINVPAYHDR(ByVal ORNum) As DataRow
        Dim dtab As DataTable
        dtab = dsINVPayHdr.Tables(0)
        Try
            Return dtab.Select("ORNum=" & ORNum)(0)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Sub PopulateDataSet(ByVal strSQL As String, ByVal strDta As String)
        Dim SQLConn As New SqlConnection(ConnStr)
        Dim SQLAdapter As New SqlDataAdapter
        Try
            SQLConn.Open()
            SQLAdapter.SelectCommand = New SqlCommand(strSQL, SQLConn)
            Select Case strDta
                Case "Header"
                    dsINVPayHdr.Reset()
                    SQLAdapter.Fill(dsINVPayHdr)

                Case "Detail"
                    dsINVPayDtl.Reset()
                    SQLAdapter.Fill(dsINVPayDtl)
            End Select

            SQLConn.Close()

        Catch ex As Exception
            SQLConn.Close()
            MsgBox("Application Cannot Retrive Data From Database", MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Function UpdateINVPAYHDR(ByVal strSQL) As Boolean
        Dim SQLConn As New SqlConnection(ConnStr)
        Dim SQLComm As SqlCommand
        Try
            SQLConn.Open()
            SQLComm = New SqlCommand(strSQL, SQLConn)

            Dim intAffectedRow As Int16
            intAffectedRow = SQLComm.ExecuteNonQuery()
            If intAffectedRow <> 1 Then
                SQLComm.Dispose()
                SQLConn.Close()
                SQLConn.Dispose()
                Return False
            End If
            SQLComm.Dispose()
            SQLConn.Close()
            SQLConn.Dispose()
            Return True
        Catch ex As Exception
            SQLComm.Dispose()
            SQLConn.Close()
            SQLConn.Dispose()
            Return False
        End Try
    End Function

    Public Function ValidateUser() As Boolean
        Dim strUserName As String = UCase(zCurrentUser())

        If strUserName <> "" Then
            Dim strSQL As String = "SELECT COUNT(*) FROM UserInfo WHERE usrpos = 'OIC Finance' AND UPPER(userid) = '" & strUserName & "'"

            Dim SQLConn As New SqlConnection(ConnStr)
            Dim SQLComm As SqlCommand

            Try
                SQLConn.Open()
                SQLComm = New SqlCommand(strSQL, SQLConn)

                Dim intAffectedRow As Int16
                intAffectedRow = SQLComm.ExecuteScalar()
                If intAffectedRow <> 1 Then
                    SQLComm.Dispose()
                    SQLConn.Close()
                    SQLConn.Dispose()
                    Return False
                End If
                SQLComm.Dispose()
                SQLConn.Close()
                SQLConn.Dispose()
                Return True

            Catch ex As Exception
                SQLComm.Dispose()
                SQLConn.Close()
                SQLConn.Dispose()
                Return False
            End Try
        Else
            Return False
        End If

    End Function

    Public Function GetCustomerName(ByVal strSQL) As String
        Dim SQLConn As New SqlConnection(ConnStr)
        Dim SQLComm As SqlCommand

        Try
            SQLConn.Open()
            SQLComm = New SqlCommand(strSQL, SQLConn)

            Dim strCusName As String
            strCusName = SQLComm.ExecuteScalar()
            If IsDBNull(strCusName) = True Then
                SQLComm.Dispose()
                SQLConn.Close()
                SQLConn.Dispose()
                Return ""
            End If
            SQLComm.Dispose()
            SQLConn.Close()
            SQLConn.Dispose()
            Return Trim(strCusName)

        Catch ex As Exception
            SQLComm.Dispose()
            SQLConn.Close()
            SQLConn.Dispose()
            Return ""
        End Try
    End Function


End Class
