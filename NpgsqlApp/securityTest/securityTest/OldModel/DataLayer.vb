Imports Npgsql
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.Data.OleDb

Public MustInherit Class DataLayer
   
    Private Const DbConnectionString As String = "securityTest.My.MySettings.ConnectionString"
    Private Const ResourceBaseName As String = "security.Resources"
    Private Const DatabaseTypeString As String = "DatabaseType"

    Private Enum DatabaseType
        PostgreSql = 1
        Sql
        OleDb
        Odbc
    End Enum

    Private Sub New()
    End Sub

    Private Shared Function getConnectionObject() As IDbConnection
        Dim ConnectionString As String = My.Settings.ConnectionString
        Dim DbTypeString As String = My.Settings.DataBaseType

        If IsNothing(ConnectionString) Or ConnectionString.Length <= 0 Then
            Throw New DataLayerException("Missing Connection String")
        End If

        If IsNothing(DbTypeString) Or DbTypeString.Length <= 0 Then
            DbTypeString = DatabaseType.PostgreSql.ToString
        End If

        Select Case DbTypeString
            Case DatabaseType.PostgreSql.ToString
                Return DirectCast(New NpgsqlConnection(ConnectionString), IDbConnection)
            Case DatabaseType.Sql.ToString
                Return DirectCast(New SqlConnection(ConnectionString), IDbConnection)
            Case DatabaseType.OleDb.ToString
                Return DirectCast(New OleDbConnection(ConnectionString), IDbConnection)
            Case DatabaseType.Odbc.ToString
                Return DirectCast(New OdbcConnection(ConnectionString), IDbConnection)
            Case Else
                Return DirectCast(New NpgsqlConnection(ConnectionString), IDbConnection)
        End Select
    End Function

    Private Shared Function CreateCommandObject(ByVal DbConnection As IDbConnection, ByVal sqlStr As String)
        Dim dbCommand As IDbCommand = DbConnection.CreateCommand()
        dbCommand.CommandText = sqlStr
        dbCommand.CommandType = CommandType.Text
        If IsNothing(dbCommand.CommandText) Or dbCommand.CommandText.Length <= 0 Then
            Throw New DataLayerException("MissingSqlString")
        End If
        Return dbCommand
    End Function

    Private Shared Function AddParameter(ByVal dbCommand As IDbCommand, ByVal ParamName As String, ByVal Value As String)
        Dim dbParam As IDataParameter = dbCommand.CreateParameter()
        dbParam.ParameterName = ParamName
        dbParam.Value = Value
        dbCommand.Parameters.Add(dbParam)

        If IsNothing(dbParam.ParameterName) Or dbParam.ParameterName.Length <= 0 Then
            Throw New DataLayerException("Missing Param String")
        End If
        Return dbParam
    End Function

    Public Shared Function CheckUserNameAndPassword(ByVal UserName As String, ByVal Password As String, ByRef userid As Integer) As Boolean
        Dim returnvalue As Object
        Dim dbConnection = getConnectionObject()
        Dim Sqlstr = "Select username from users where username = @username"
        Dim dbCommand As IDbCommand = CreateCommandObject(dbConnection, Sqlstr)
        AddParameter(dbCommand, "username", UserName)
        Try
            dbConnection.Open()
            returnvalue = dbCommand.ExecuteScalar
        Catch ex As Exception
        Finally
            dbConnection.Close()
        End Try
        If IsNothing(returnvalue) Then
            userid = 0
            Return False
        Else
            Return True
        End If
    End Function

End Class

<Serializable()>
Public Class DataLayerException
    Inherits ApplicationException

    Public Sub New(ByVal ErrorMessage As String)
        MyBase.New(ErrorMessage)
    End Sub
End Class