Imports Npgsql
Imports HR.HelperClass
Public Class DbAdapter
    Dim _connectionstring As String
    Dim myTransaction As NpgsqlTransaction

    Public Property Connectionstring As String
        Get
            Return _connectionstring

        End Get
        Set(ByVal value As String)
            _connectionstring = value
        End Set
    End Property


    Public Sub New(ByVal Connectionstring As String)
        _connectionstring = Connectionstring
    End Sub
#Region "GetDataSet"
    Public Overloads Function TbgetDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter    
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region

#Region "GetDataSetTable"
    Public Overloads Function TbgetDataSet(ByVal TableName As String, ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                DataAdapter.Fill(DataSet, 0, 10, TableName)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region

#Region "SaveChanges"
    Public Function SaveChanges(ByRef DataSet As DataSet, ByVal sqlstr As String, Optional ByRef message As String = "") As Boolean        
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(ConnectionString)
                conn.Open()
                Dim cmd = New NpgsqlCommand(sqlstr)
                Dim cmdbuilder = New NpgsqlCommandBuilder(DataAdapter)
                DataAdapter.SelectCommand = cmd
                DataAdapter.SelectCommand.Connection = conn
                DataAdapter.Update(DataSet.Tables(0))
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region

#Region "TBProgram"

    Public Function TBProgramSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False
        Dim param As NpgsqlParameter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    'Select
                    sqlstr = "select * from tbprogram"
                    cmd = New NpgsqlCommand(sqlstr)
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                    'Delete
                    sqlstr = "Delete from tbprogram where programid = @programid"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("@programid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "programid").SourceVersion = DataRowVersion.Original

                    'Update
                    sqlstr = "Update tbprogram set parentid = @parentid,myorder = @myorder,description = @description, programname = @programname,isactive = @isactive,icon = @icon, iconindex = @iconindex,members = @members,formname = @formname" &
                             " where programid = @programid and latestupdate=@latestupdate;" & _
                             " select latestupdate from tbprogram where programid = @programid;"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("@programid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "programid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("@parentid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "parentid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@myorder", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myorder").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@description", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@programname", NpgsqlTypes.NpgsqlDbType.Text, 0, "programname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@isactive", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@icon", NpgsqlTypes.NpgsqlDbType.Text, 0, "icon").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@iconindex", NpgsqlTypes.NpgsqlDbType.Integer, 0, "iconindex").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@members", NpgsqlTypes.NpgsqlDbType.Text, 0, "members").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@formname", NpgsqlTypes.NpgsqlDbType.Text, 0, "formname").SourceVersion = DataRowVersion.Current
                    param = DataAdapter.UpdateCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate")
                    param.SourceVersion = DataRowVersion.Current
                    param.Direction = ParameterDirection.InputOutput

                    'insert
                    sqlstr = "insert into tbprogram(parentid,myorder,description,programname,isactive,icon,iconindex,members,applicationname,formname) values " & _
                             "(@parentid,@myorder,@description,@programname,@isactive,@icon,@iconindex,@members,@applicationname,@formname);" & _
                             " select currval('tbprogram_programid_seq') as programid,latestupdate from tbprogram where programid = currval('tbprogram_programid_seq');"

                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("@parentid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "parentid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@myorder", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myorder").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@description", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@programname", NpgsqlTypes.NpgsqlDbType.Text, 0, "programname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@isactive", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@icon", NpgsqlTypes.NpgsqlDbType.Text, 0, "icon").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@iconindex", NpgsqlTypes.NpgsqlDbType.Integer, 0, "iconindex").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@members", NpgsqlTypes.NpgsqlDbType.Text, 0, "members").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@formname", NpgsqlTypes.NpgsqlDbType.Text, 0, "formname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@programid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "programid").Direction = ParameterDirection.Output
                    param = DataAdapter.InsertCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate")
                    param.Direction = ParameterDirection.Output
                    RecordAffected = DataAdapter.Update(DataSet.Tables("TBProgram"))
                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region

#Region "TBRoles"
    Public Function TBRolesSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False
        Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    'Select
                    sqlstr = "select * from roles"
                    'sqlstr = "sp_selectuser"
                    cmd = New NpgsqlCommand(sqlstr)
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
                    'DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                    'Delete
                    sqlstr = "Delete from roles where rolename = @rolename and applicationname=@applicationname"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.DeleteCommand.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Text, 0, "rolename").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").Value = applicationname

                    'Update
                    sqlstr = "Update roles set rolename = @rolename,applicationname = @applicationname" &
                             " where rolename = @rolenameori and applicationname=@applicationnameori"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Text, 0, "rolename").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").Value = applicationname
                    DataAdapter.UpdateCommand.Parameters.Add("@rolenameori", NpgsqlTypes.NpgsqlDbType.Text, 0, "rolename").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("@applicationnameori", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").Value = applicationname

                    'insert
                    sqlstr = "insert into roles(rolename,applicationname) values " & _
                             "(@rolename,@applicationname)"

                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Text, 0, "rolename").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").Value = applicationname

                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))

                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region

#Region "User"
    Public Function UpdateUserRegion(ByVal email As String, ByVal regionid As Integer) As String
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_updateuserregion", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = email
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception

        End Try
        Return result
    End Function

    Public Function CreateUserDb(ByVal username As String, ByVal password As String) As String
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("create role " & username & " with login password '" & password & "'", conn)
                cmd.CommandType = CommandType.Text
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception

        End Try
        Return result
    End Function
    Public Function ChangePasswordDb(ByVal username As String, ByVal password As String) As String
        Dim result As Object = Nothing
        'Try
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("Alter role " & username & " with password '" & password & "'", conn)
            cmd.CommandType = CommandType.Text
            result = cmd.ExecuteScalar
        End Using
        'Catch ex As Exception

        'End Try
        Return result
    End Function


    Public Function DropUserDb(ByVal username As String) As String
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("drop role if exists " & username, conn)
                cmd.CommandType = CommandType.Text
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception

        End Try
        Return result
    End Function
#End Region

#Region "ImportFinanceInformation"
    Public Function getSapAccountid(ByVal sapaccount As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapaccountid", conn)
                cmd.CommandType = CommandType.StoredProcedure                
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapaccount
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Public Function getdeptid(ByVal dept As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getdeptid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dept
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Public Function getExpensesNatureId(ByVal expensesnature As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getexpensesnatureid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = expensesnature
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Public Function getSapAccNameId(ByVal sapaccname As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapaccnameid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapaccname
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getSapIndexid(ByVal sapaccountid As Integer, ByVal sapaccid As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapindexid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = sapaccountid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapaccid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getsapccid(ByVal sapcc As String, ByVal currency As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapccid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapcc
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = currency
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function getsapccid(ByVal costcenter As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapccid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = costcenter
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getIndexcostcenterId(ByVal sapindexid As Integer, ByVal costcenterid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getindexcostcenterid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = sapindexid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = costcenterid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getIndexCostCenterDeptId(ByVal indexcostcenterid As Integer, ByVal deptid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getindexcostcenterdeptid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = indexcostcenterid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = deptid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getAccExpensesId(ByVal sapaccnameid As Integer, ByVal expensesnatureid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getaccexpensesid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = sapaccnameid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesnatureid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function



    Function getExpensesDetailid(ByVal accexpensesid As Integer, ByVal indexcostcenterdeptid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getexpensesdetailid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = accexpensesid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = indexcostcenterdeptid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function setexpensesnaturefullyear(ByVal expensesnatureid As Integer, ByVal value As Boolean) As Boolean
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_setexpensesnaturefullyear", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesnatureid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = value

                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function insertexpensesnaturemonths(ByVal expensesnatureid As Integer, ByVal monthToCreate As String, ByVal year As Integer) As Boolean
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertexpensesnaturemonths", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesnatureid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = monthToCreate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = year

                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function


#End Region

#Region "HR"
    Function getcategoryId(ByVal category As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getcategoryid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = category
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = category.Substring(0, 2)
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function


    Function getcategorytypeId(ByVal categorytype As Object) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getcategorytypeid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = categorytype
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function gettitleid(ByVal title As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_gettitleid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = title
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getpersonid(ByVal personname As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getpersonid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = personname
                'cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = expat


                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function insertpersontitle(ByVal titleid As Integer, ByVal joindate As Date, ByVal personjoindatecategoryid As Integer) As Object
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertpersontitle", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = titleid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = joindate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindatecategoryid

                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getpersonjoindateid(ByVal personid As Integer, ByVal joindate As Date, ByVal othername As String, ByVal status As Boolean, ByVal enddate As Date) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getpersonjoindateid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = joindate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = othername
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = status
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = enddate
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getpersonjoindateid(ByVal personid As Integer, ByVal joindate As Date, ByVal othername As String, ByVal status As Boolean) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getpersonjoindateid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = joindate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = othername
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = status
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getpersonjoindatecategoryid(ByVal categoryid As Integer, ByVal personjoindateid As Integer, ByVal deptid As Integer, ByVal headcount As Double, ByVal expat As Boolean) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getpersonjoindatecategoryid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = categoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindateid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = deptid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = headcount
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = expat

                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function getpersonjoindatecategoryid(ByVal categoryid As Integer, ByVal personjoindateid As Integer, ByVal deptid As Integer, ByVal headcount As Double, ByVal expat As Boolean, ByVal effectivedatestart As Date, ByVal effectivedateend As Date?, ByVal bonusfactor As Integer?) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getpersonjoindatecategoryid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = categoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindateid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = deptid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = headcount
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = expat
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = effectivedatestart
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = IIf(IsNothing(effectivedateend), DBNull.Value, effectivedateend)
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = IIf(IsNothing(bonusfactor), DBNull.Value, bonusfactor)
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getExpensesDetailtxid(ByVal myyear As Integer, ByVal expensesdetailid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getexpensesdetailtxid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesdetailid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function insertfamilymemberid(ByVal personjoindateid As Integer, ByVal planname As String, ByVal count As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertfamilymemberid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindateid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = planname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = count
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertfamilymemberid(ByVal personjoindateid As Integer, ByVal planname As String, ByVal count As String, ByVal myyear As Integer, ByVal verid As Integer, ByVal regionid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertfamilymemberid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindateid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = planname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = count
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = verid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertpersonexpensesid(ByVal icdpjcid As Integer, ByVal expensesdetailtxid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertpersonexpensesid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = icdpjcid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesdetailtxid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function


    Function insertsalarytx(ByVal personjoindateid As Integer, ByVal value As Integer, ByVal startingdate As Date) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertsalarytx", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindateid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = value
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = startingdate

                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertsalarytx(ByVal category As String, ByVal costcenter As String, ByVal joindate As Date, ByVal personname As String, _
                            ByVal myvalue As Double, ByVal startingdate As Date, ByVal txtype As String, ByVal verid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertsalarytx", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = category
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = costcenter
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = joindate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = personname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = myvalue
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = startingdate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = txtype
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = verid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertsalarytx(ByVal category As String, ByVal costcenter As String, ByVal joindate As Date, ByVal personname As String, _
                            ByVal myvalue As Double, ByVal startingdate As Date, ByVal txtype As String, ByVal verid As Integer, ByVal regionid As Integer, ByVal myyear As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertsalarytx", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = category
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = costcenter
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = joindate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = personname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = myvalue
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = startingdate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = txtype
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = verid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear

                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertpersonexpensesdtl(ByVal personexpensesid As Integer, ByVal value As Double, ByVal startingdate As Date) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertpersonexpensesdtl", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personexpensesid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = value
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = startingdate
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function insertcategorydetail(ByVal expensesnatureid As Integer, ByVal categoryid As Integer, ByVal amount As Double, ByVal startingdate As Date, ByVal year As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertcategorydetail", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesnatureid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = amount
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = startingdate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = year
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertpersonexpensesdtl(ByVal sapaccname As String, ByVal expensesnature As String, ByVal sapaccount As String, ByVal category As String, _
                                     ByVal costcenter As String, ByVal dept As String, ByVal joindate As Date, ByVal staffname As String, ByVal amount As Double, ByVal startingdate As Date) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertpersonexpensesdtl", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapaccname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = expensesnature
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapaccount
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = category
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = costcenter
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dept
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = joindate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = staffname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Double, 0).Value = amount
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertcategorydtl(ByVal categoryid As Integer, ByVal categorytypeid As Integer, ByVal value As Double, ByVal myyear As Integer, ByVal verid As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertcategorydtl", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = categoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = categorytypeid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = value
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = verid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getpersonexpensesid(ByVal icdpjcid As Object, ByVal expensesdetailtxid As Object) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertpersonexpenses", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = icdpjcid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesdetailtxid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getRegionShortName(ByVal username As String) As String
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getregionshortname", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function getRegionID(ByVal username As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getregionid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function getplanid(ByVal planname As String) As Double
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getplanid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = planname

                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertCategorytxMonths(ByVal categoryid As Integer, ByVal categorytype As String, ByVal monthToInsert As String, ByVal myYear As Integer, ByVal myverid As Integer)
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertCategorytxMonths", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = categoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = categorytype
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = monthToInsert
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myYear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myverid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function insertCategorytxMonths(ByVal categoryid As Integer, ByVal categorytype As String, ByVal monthToInsert As String, ByVal myyear As Integer, ByVal myVerid As Integer, ByVal regionimport As Integer)
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertCategorytxMonths", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = categoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = categorytype
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = monthToInsert
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function geticdpjcid(ByVal personjoindatecategoryid As Integer, ByVal indexcostcenterdeptid As Integer, ByVal accexpensesid As Integer, ByVal myyear As Integer, ByVal verid As Integer, ByVal regionimport As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_geticdpjcid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindatecategoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = indexcostcenterdeptid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = accexpensesid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = verid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getRegionName(ByVal username As String) As String
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getregionname", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getVerId(ByVal versionname As String, ByVal versionReason As String, ByVal myyear As Integer) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getversionid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Versionname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = versionReason
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getRegionIDFromRegionName(ByVal regionname As String) As Integer
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getregionidfromregionname", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = regionname
                result = cmd.ExecuteScalar
            End Using

        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Sub deleteicdpjc(ByVal myyear As Integer, ByVal myVerid As Integer, ByVal regionimport As Integer)
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_deleteicdpjc", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                result = cmd.ExecuteScalar
            End Using

        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try

    End Sub

    Sub deletefamilymember(ByVal myyear As Integer, ByVal myVerid As Integer, ByVal regionimport As Integer)
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_deletefamilymemberplan", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                result = cmd.ExecuteScalar
            End Using

        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try

    End Sub
    Sub deletesalartytx(ByVal myyear As Integer, ByVal myVerid As Integer, ByVal regionimport As Integer)
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_deletesalarytx", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                result = cmd.ExecuteScalar
            End Using

        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Sub
    Sub deletepersontxmonth(ByVal myyear As Integer, ByVal myVerid As Integer, ByVal regionimport As Integer)
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_deletepersontxmonth", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                result = cmd.ExecuteScalar
            End Using

        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Sub

    Sub deletecategorytxmonth(ByVal myyear As Integer, ByVal myVerid As Integer, ByVal regionimport As Integer)
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_deletecategorytxmonths", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                result = cmd.ExecuteScalar
            End Using

        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Sub

#End Region



    Function TBParamDtSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False
        Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_getcurrency() as tb(paramname character varying,dvalue date,nvalue numeric,paramdtid integer,paramhdid integer) "
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey
                    DataAdapter.Fill(DataSet)

                    'Delete
                    sqlstr = "sp_deletecurrency"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paramdtid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updatecurrency"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paramdtid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paramname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "dvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertcurrency"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paramname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "dvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))

                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function



    Function TBParamDataAdapter(ByRef DataSet As DataSet, ByVal paramhdid As Integer, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False
        Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_getparam(:col1) as tb(paramname character varying,cvalue character varying,dvalue date,ivalue integer,nvalue numeric,paramdtid integer,paramhdid integer)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey
                    DataAdapter.SelectCommand.Parameters.Add("col1", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dbtools1.Region
                    DataAdapter.Fill(DataSet)

                    'Delete
                    sqlstr = "sp_deleteparam"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paramdtid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updateparam"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paramdtid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paramname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "dvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ivalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertparam"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paramname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "dvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ivalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("paramhdid", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = paramhdid                    
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))

                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Function TBPlanTxDataAdapter(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False
        Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_getplantx() as tb(plantxid integer,planid integer,plantypeid integer,nominal numeric,validfrom date)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey                    
                    DataAdapter.Fill(DataSet)

                    'Delete
                    sqlstr = "sp_deleteplantx"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plantxid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updateplantx"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plantxid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "planid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plantypeid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nominal").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validfrom").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertplantx"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "planid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plantypeid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nominal").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validfrom").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))

                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Sub insertsalarytx(ByVal categoryid As Integer, ByVal amount As Double, ByVal mydate As Date, ByVal txtype As String, ByVal myVerid As Integer)
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertsalarytx", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = categoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = amount
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = mydate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = txtype
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid

                result = cmd.ExecuteScalar
            End Using

        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Sub
    Sub insertsalarytx(ByVal categoryid As Integer, ByVal amount As Double, ByVal mydate As Date, ByVal txtype As String, ByVal myVerid As Integer, ByVal regionimport As Integer, ByVal myyear As Integer)
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertsalarytx", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = categoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = amount
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = mydate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = txtype
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                result = cmd.ExecuteScalar
            End Using

        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Sub
    Function insertcategory(ByVal category As String, ByVal description As String, ByVal regionid As Integer) As Integer
        Dim result As Object = Nothing

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertcategory", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = category
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = description
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function gettargetrate(ByVal personjoindatecategoryid As Integer, ByVal period As Date) As Decimal
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_gettargetrate", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindatecategoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = period
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getsapcc(ByVal personjoindatecategoryid As Integer) As String
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapcc", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindatecategoryid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Sub UpdateLocation(ByVal category As String, ByVal location As String)
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_updatelocation", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = category
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = location
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try

    End Sub

    Function insertPersonalTxMonths(ByVal personjoindatecategoryid As Integer, ByVal expensesnatureid As Integer, ByVal monthToInsert As String, ByVal myyear As Integer, ByVal myverid As Integer)
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertPersontxMonths", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindatecategoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesnatureid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = monthToInsert
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myverid
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function insertPersonalTxMonths(ByVal personjoindatecategoryid As Integer, ByVal expensesnatureid As Integer, ByVal monthToInsert As String, ByVal myyear As Integer, ByVal myVerid As Integer, ByVal regionimport As Integer)
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertPersontxMonths", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = personjoindatecategoryid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = expensesnatureid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = monthToInsert
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myVerid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = regionimport
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function
    Function validversion(ByVal versionid As Integer, ByVal myyear As Integer, ByVal paramname As String, ByVal regionname As String) As Boolean
        Dim result As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_checkcurrentversion", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = versionid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = paramname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = regionname
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getlastversion(ByVal myyear As Integer, ByVal paramname As String, ByVal regionname As String) As Integer
        Dim result As Integer
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getcurrentversion", conn)
                cmd.CommandType = CommandType.StoredProcedure                
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = paramname
                'cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = regionname
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function setcurrentversion(ByVal versionid As Object, ByVal myyear As Integer, ByVal paramname As String, ByVal regionname As String, ByVal versionname As String) As Boolean
        Dim result As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_setcurrentversion", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = versionid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = myyear
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = paramname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = regionname
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = versionname
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getSapAccountFid(ByVal sapaccountf As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapaccountfid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapaccountf
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getsapccfid(ByVal sapccf As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapccfid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapccf
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getSapIndexfid(ByVal sapaccountfid As Integer, ByVal sapccfid As Integer, ByVal sapindexf As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapindexfid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = sapaccountfid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = sapccfid
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapindexf
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function getSapAccNamefId(ByVal sapaccnamef As String) As Integer
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getsapaccnamefid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sapaccnamef
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function checkSapAccNameId(ByVal sapaccname As String) As Boolean
        Throw New NotImplementedException
    End Function

    Function createAccountName(ByVal accountname As String) As Long
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertaccountname", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = accountname
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return result
    End Function

    Function GroupCategoryAdapter(ByVal Dataset As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False
        Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_getgroupingtable() as tb(category character varying,sapaccname character varying,sapaccount character varying,groupingtableid integer)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey
                    DataAdapter.Fill(Dataset)

                    'Delete
                    sqlstr = "sp_deletegroupingtable"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "groupingtableid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updategroupingtable"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "groupingtableid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "category").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sapaccname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sapaccount").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertgroupingtable"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "category").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sapaccname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sapaccount").SourceVersion = DataRowVersion.Current
                    
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    RecordAffected = DataAdapter.Update(Dataset.Tables(0))

                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Sub deletesapindexf()
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("delete from sapindexf;select setval('sapindexf_sapindexfid_seq',1,false)", conn)
                cmd.CommandType = CommandType.Text
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try

    End Sub

    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Dim builder As New NpgsqlConnectionStringBuilder()
        builder.ConnectionString = "host=hon14nt;port=5432;database=LogisticDb;commandTimeout=1000;Timeout=1000;"
        builder.Add("User Id", "admin")
        builder.Add("password", "admin")
        Dim Connectionstring = builder.ConnectionString
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = "admin"
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function
   
End Class
