Imports System.Data.SqlClient
Module modReports

    Private objConn As SqlConnection
    Private strConn As String = "Data Source=SBITCBILLING;Initial Catalog=BILLING;Integrated Security=SSPI"

    Private Function Connection() As Boolean
        Try
            If objConn Is Nothing Then
                objConn = New SqlConnection
                objConn.ConnectionString = Trim(strConn)
                objConn.Open()
            End If

            If objConn.State = ConnectionState.Closed Then
                objConn.Open()
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Connection Error")
            Return False
        End Try
    End Function

    Public Function Get_CCRCyx(ByVal dteStart As Date, ByVal dteEnd As Date) As DataTable
        Dim cmdCCRCyx As New SqlCommand
        Dim daCCRCyx As New SqlDataAdapter
        Dim dtabCCRCyx As New DataTable

        If Connection() = True Then
            With cmdCCRCyx
                .Connection = objConn
                .CommandText = "SELECT * FROM CCRCyx Where Status <> 'CAN' AND sysdttm >='" & dteStart & " 00:00:00 AM' AND sysdttm <='" & dteEnd & " 11:58:59 PM'"
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daCCRCyx.SelectCommand = cmdCCRCyx
                daCCRCyx.Fill(dtabCCRCyx)

                Return dtabCCRCyx
            End With
        End If
    End Function

    Public Function Get_CYMGps(ByVal dteStart As Date, ByVal dteEnd As Date) As DataTable
        Dim cmdCYMGps As New SqlCommand
        Dim daCYMGps As New SqlDataAdapter
        Dim dtabCYMGps As New DataTable

        If Connection() = True Then
            With cmdCYMGps
                .Connection = objConn
                .CommandText = "SELECT * FROM CYMGps Where Status <> 'CAN' AND sysdte >='" & dteStart & " 00:00:00 AM' AND sysdte <='" & dteEnd & " 11:58:59 PM'"
                .CommandType = CommandType.Text

                .ExecuteNonQuery()

                daCYMGps.SelectCommand = cmdCYMGps
                daCYMGps.Fill(dtabCYMGps)

                Return dtabCYMGps
            End With
        End If
    End Function

    Public Function DisConnect()
        If Not objConn Is Nothing Then
            objConn.Close()
            objConn = Nothing
        End If
    End Function
End Module
