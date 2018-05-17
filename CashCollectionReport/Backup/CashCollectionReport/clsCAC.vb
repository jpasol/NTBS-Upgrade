Option Explicit On 
Imports System.Data
Imports System.Data.SqlClient

Public Class clsCAC

    'Connection String for GIP Database    
    Private ConnStr As String = _
        "Initial Catalog=Billing;" & _
        "Data Source=SBITCBILLING;" & _
        "integrated security=sspi"

    Public Function GetTurnOver(ByVal xID) As DataRow
        Dim dtab As DataTable
        dtab = dsTurnOverSlip.Tables(0)
        Try
            Return dtab.Select("ID=" & xID)(0)
        Catch ex As Exception
            MsgBox("Application Cannot Retrive Data From Database", MsgBoxStyle.Critical, "ERROR")
            Return Nothing
        End Try
    End Function

    Public Sub RetrieveCAC(ByVal strSQL)
        Dim SQLConn As New SqlConnection(ConnStr)
        Dim SQLAdapter As New SqlDataAdapter
        Try
            SQLConn.Open()
            SQLAdapter.SelectCommand = New SqlCommand(strSQL, SQLConn)
            dsCAC.Reset()
            SQLAdapter.Fill(dsCAC)
            SQLConn.Close()
        Catch ex As Exception
            MsgBox("Application Cannot Retrive Data From Database", MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub RetrieveTurnOverSlip(ByVal strSQL)
        Dim SQLConn As New SqlConnection(ConnStr)
        Dim SQLAdapter As New SqlDataAdapter
        Try
            SQLConn.Open()
            SQLAdapter.SelectCommand = New SqlCommand(strSQL, SQLConn)
            dsTurnOverSlip.Reset()
            SQLAdapter.Fill(dsTurnOverSlip)
            SQLConn.Close()
        Catch ex As Exception
            MsgBox("Application Cannot Retrive Data From Database", MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Function GetBatch(ByVal strSQL) As String
        Dim SQLConn As New SqlConnection(ConnStr)
        Dim SQLComm As SqlCommand
        Try
            SQLConn.Open()

            SQLComm = New SqlCommand(strSQL, SQLConn)

            Dim intBatch As Integer
            intBatch = SQLComm.ExecuteScalar()
            SQLConn.Close()
            SQLComm.Dispose()
            SQLConn.Dispose()
            Return CType(intBatch, String)
        Catch ex As Exception
            Return "Application Cannot Retrive Data From Database"
        End Try
    End Function

    Public Function SaveCAC(ByVal strSQL) As Boolean
        Dim SQLConn As New SqlConnection(ConnStr)
        Dim SQLComm As SqlCommand
        Try
            SQLConn.Open()

            SQLComm = New SqlCommand(strSQL, SQLConn)

            Dim intAffectedRows As Integer
            intAffectedRows = SQLComm.ExecuteNonQuery()
            If intAffectedRows <> 1 Then
                SQLConn.Close()
                SQLComm.Dispose()
                SQLConn.Dispose()
                Return False
            End If
            SQLConn.Close()
            SQLComm.Dispose()
            SQLConn.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function NumVal(ByVal xValue) As Boolean
        If IsNumeric(xValue) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function StrVal(ByVal xValue) As Boolean
        If xValue = "" Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function TimeVal(ByVal xTimeFrom, ByVal xTimeTo) As Boolean
        Try
            If TimeValue(xTimeFrom) < TimeValue(xTimeTo) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function getToString(ByRef strValue) As String
        If strValue.Trim = "" Then
            Return "NULL"
        Else
            Return "'" & strValue.Trim & "'"
        End If
    End Function

    Function isNull(ByRef xValue) As String
        If IsDBNull(xValue) Then
            Return ""
        Else
            Dim Value As String = xValue.Trim
            Return Value
        End If
    End Function

End Class
