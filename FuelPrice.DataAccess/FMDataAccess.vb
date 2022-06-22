Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Public Class FMDataAccess
    Dim Conn As New SqlConnection

    Public Function Connect(ByVal ConnectionName) As Boolean
        Dim result As Boolean

        Try
            'If Conn IsNot Nothing Then
            If Conn.State = ConnectionState.Open And Conn.ConnectionString = ConfigurationManager.ConnectionStrings(ConnectionName).ToString() Then
                result = True
            Else
                'Conn = New SqlConnection
                If Conn.State = ConnectionState.Open Then Conn.Close()
                Conn.ConnectionString = ConfigurationManager.ConnectionStrings(ConnectionName).ToString()
                Conn.Open()
                result = True
            End If
        Catch ex As Exception
            result = False
        End Try

        Return result
    End Function

    Public Function Consulta(ByVal Query, ByVal ConnectionName) As DataTable
        Dim result As New DataTable

        Try
            If Connect(ConnectionName) Then
                Dim comando As New SqlCommand(Query, Conn)
                Dim adapter As New SqlDataAdapter(comando)
                'Se agregó para dar más tiempo de respuesta a las consultas (ej. Fórmulas)
                adapter.SelectCommand.CommandTimeout = 300

                adapter.Fill(result)
                Conn.Close()
                adapter.Dispose()
            End If
        Catch ex As Exception
            result = New DataTable
        End Try

        Return result
    End Function

    Public Function NoQuery(ByVal Query, ByVal ConnectionName) As Integer
        Dim result As Integer

        Try
            If Connect(ConnectionName) Then
                Dim comando As New SqlCommand(Query, Conn)
                result = comando.ExecuteNonQuery()
            Else
                result = 0
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return result
    End Function
    Public Function IfExist(ByVal query, ByVal ConnectionName) As Boolean
        Dim result As Boolean

        Try
            If Connect(ConnectionName) Then
                Dim comando As New SqlCommand(query, Conn)
                Dim adapter As New SqlDataAdapter(comando)
                Dim dt As New DataTable()
                adapter.Fill(dt)
                Conn.Close()
                adapter.Dispose()
                If dt.Rows.Count > 0 Then
                    result = True
                Else
                    result = False
                End If
            End If
        Catch ex As Exception
            result = False

        End Try
        Return result
    End Function
End Class
