Imports System.Data.SqlClient

Public Class clsUpload

    'Object
    Private objConn As SqlConnection

    'Variables
    Public strServer, strDatabase As String

    Public Sub main()
        Dim frmUpload As New frmUploading

        strSrvr = strServer
        strDB = strDatabase
        frmUpload.Show()
    End Sub

    Public Function Connect() As Boolean
        Try
            If objConn Is Nothing Then
                strConn = "Data Source='" & Trim(strSrvr) & "';Initial Catalog='" & Trim(strDB) & "';Integrated Security=SSPI"
                objConn = New SqlConnection
                objConn.ConnectionString = Trim(strConn)
                objConn.Open()
            End If

            If objConn.State = ConnectionState.Closed Then
                objConn.Open()
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, "NTBS Connection Error")
            objConn = Nothing
            Return False
        End Try
    End Function

    Public Function Check_Header(ByVal strBL As String) As Boolean
        Dim cmdCMH As New SqlCommand
        Dim daCMH As New SqlDataAdapter
        Dim dtabCMH As New DataTable

        If Connect() = True Then
            With cmdCMH
                .Connection = objConn
                .CommandText = "Select * From CargoMHead Where BilNum='" & Trim(strBL) & "'"
                .CommandType = CommandType.Text
                .ExecuteNonQuery()
            End With

            daCMH.SelectCommand = cmdCMH
            daCMH.Fill(dtabCMH)

            If dtabCMH.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Function Check_HeaderByRegNum(ByVal strRegNum As String) As Boolean
        Dim cmdCMH As New SqlCommand
        Dim daCMH As New SqlDataAdapter
        Dim dtabCMH As New DataTable

        If Connect() = True Then
            With cmdCMH
                .Connection = objConn
                .CommandText = "Select * From CargoMHead Where RegNum='" & Trim(strRegNum) & "'"
                .CommandType = CommandType.Text
                .ExecuteNonQuery()
            End With

            daCMH.SelectCommand = cmdCMH
            daCMH.Fill(dtabCMH)

            If dtabCMH.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Function Check_Detail(ByVal strBL As String, ByVal strCtnNo As String) As Boolean
        Dim cmdCMD As New SqlCommand
        Dim daCMD As New SqlDataAdapter
        Dim dtabCMD As New DataTable

        If Connect() = True Then
            With cmdCMD
                .Connection = objConn
                .CommandText = "Select * From CargoMDet Where BilNum='" & Trim(strBL) & "' And CtnNum='" & Trim(strCtnNo) & "'"
                .CommandType = CommandType.Text
                .ExecuteNonQuery()
            End With

            daCMD.SelectCommand = cmdCMD
            daCMD.Fill(dtabCMD)

            If dtabCMD.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Function Insert_Header(ByVal strBL As String, ByVal strCarCde As String, _
                                  ByVal strPO As String, ByVal strRegNum As String, _
                                  ByVal strConsignee As String, ByVal strCtnNamDescr As String, _
                                  ByVal strBroker As String, ByVal strUser As String) As Boolean

        Dim cmdCMH As New SqlCommand
        Dim strQry As String = ""

        strQry = "Insert Into CargoMHead(bilnum,carcde,po," & _
                 "regnum,consignee,ctnNameDesc,broker," & _
                 "sysdte,userid) Values('" & Trim(strBL) & "','" & _
                 Trim(strCarCde) & "','" & Trim(strPO) & "','" & _
                 Trim(strRegNum) & "','" & Trim(strConsignee) & "','" & _
                 Trim(strCtnNamDescr) & "','" & Trim(strBroker) & "','" & _
                 gzGetSysDate() & "','" & Trim(strUser) & "')"

        If Connect() = True Then
            Try
                With cmdCMH
                    .Connection = objConn
                    .CommandText = Trim(strQry)
                    .CommandType = CommandType.Text
                    .ExecuteNonQuery()

                    Return True
                End With
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Header Insertion Error")
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Function Update_Header(ByVal strRegNum As String, ByVal strVslName As String, _
                                  ByVal dteArrival As Date, ByVal strVoyNum As String) As Boolean

        Dim cmdCMH As New SqlCommand
        Dim strQry As String = ""

        strQry = "Update CargoMHead Set vslname ='" & Trim(strVslName) & "'," & _
                 "voynum ='" & Trim(strVoyNum) & "', arvdte='" & dteArrival.ToShortDateString & "' Where regnum ='" & Trim(strRegNum) & "'"

        If Connect() = True Then
            Try
                With cmdCMH
                    .Connection = objConn
                    .CommandText = Trim(strQry)
                    .CommandType = CommandType.Text
                    .ExecuteNonQuery()

                    Return True
                End With
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Header Insertion Error")
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Function Insert_Detail(ByVal strBL As String, ByVal strCtnNo As String, _
                                  ByVal strCtnType As String, ByVal intCtnSze As String, _
                                  ByVal strWeight As String, ByVal strFullEmpty As String, _
                                  ByVal strSilNo As String, ByVal strRegNum As String, _
                                  ByVal strUser As String) As Boolean
        Dim cmdCMD As New SqlCommand
        Dim strQry As String = ""

        strQry = "Insert Into CargoMDet(bilnum,ctnnum,ctntype,ctnsze," & _
                 "ctnweight,fullempty,silnum,regnum,sysdte,userid) " & _
                 "Values('" & Trim(strBL) & "','" & Trim(strCtnNo) & "','" & _
                 Trim(strCtnType) & "'," & intCtnSze & ",'" & Trim(strWeight) & "','" & _
                 Trim(strFullEmpty) & "','" & Trim(strSilNo) & "','" & Trim(strRegNum) & "','" & _
                 gzGetSysDate() & "','" & Trim(strUser) & "')"

        If Connect() = True Then
            Try
                With cmdCMD
                    .Connection = objConn
                    .CommandText = Trim(strQry)
                    .CommandType = CommandType.Text
                    .ExecuteNonQuery()

                    Return True
                End With
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Detail Insertion Error")
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Private Function gzGetSysDate() As Date
        Dim cmdDate As New SqlCommand

        With cmdDate
            .Connection = objConn
            .CommandText = "Select getdate()"
            .CommandType = CommandType.Text

            gzGetSysDate = .ExecuteScalar()
        End With
    End Function
End Class
