Imports Npgsql

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

#Region "Sample Store Procedure"
    'CREATE OR REPLACE FUNCTION insertanimal(animaltype character varying)
    '  RETURNS bigint AS
    '$BODY$insert into exanimal(exanimaltype) values($1);
    'select currval('exanimal_exanimalid_seq');
    '$BODY$
    '  LANGUAGE sql VOLATILE
    '  COST 100;
    'ALTER FUNCTION insertanimal(character varying) OWNER TO postgres;

    Public Function InsertAnimal(ByVal AnimalType As String) As Long
        Dim result As Object = Nothing
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("funcinsertanimal", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0).Value = AnimalType
                result = cmd.ExecuteScalar
            End Using
        Catch ex As Exception

        End Try
        Return result
    End Function
#End Region

#Region "RefCursonAnimal"
    Public Function GetRefCursorAnimal(ByRef dataset As DataSet) As Boolean
        Dim result As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim trans = conn.BeginTransaction
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("refcursoranimal", conn)
                cmd.CommandType = CommandType.StoredProcedure
                Dim DataAdapter = New NpgsqlDataAdapter(cmd)
                DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                DataAdapter.Fill(dataset)
                trans.Commit()
                result = True
            End Using
        Catch ex As Exception

        End Try
        Return result

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
                    sqlstr = "Update tbprogram set parentid = @parentid,myorder = @myorder,description = @description, programname = @programname,isactive = @isactive,icon = @icon, iconindex = @iconindex,members = @members" &
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
                    param = DataAdapter.UpdateCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate")
                    param.SourceVersion = DataRowVersion.Current
                    param.Direction = ParameterDirection.InputOutput

                    'insert
                    sqlstr = "insert into tbprogram(parentid,myorder,description,programname,isactive,icon,iconindex,members,applicationname) values " & _
                             "(@parentid,@myorder,@description,@programname,@isactive,@icon,@iconindex,@members,@applicationname);" & _
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
                    cmd = New NpgsqlCommand(sqlstr)
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
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

#Region "TBUsers"
    Public Function TBUsersSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
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
                    sqlstr = "select * from users"
                    cmd = New NpgsqlCommand(sqlstr)
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                    'Delete
                    sqlstr = "Delete from users where pkid=@pkid"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Text, 0, "username").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.Parameters.Add("@pkid", NpgsqlTypes.NpgsqlDbType.Uuid, 0, "pkid").SourceVersion = DataRowVersion.Original

                    'Update
                    sqlstr = "Update users set username = @username,applicationname = @applicationname,email=@email,comments=@comments,isapproved=@isapproved" &
                             " where pkid=@pkid"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Text, 0, "username").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").Value = applicationname
                    DataAdapter.UpdateCommand.Parameters.Add("@pkid", NpgsqlTypes.NpgsqlDbType.Uuid, 0, "pkid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("@email", NpgsqlTypes.NpgsqlDbType.Text, 0, "email").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@comments", NpgsqlTypes.NpgsqlDbType.Text, 0, "comments").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@isapproved", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isapproved").SourceVersion = DataRowVersion.Current

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

#Region "TBAnimal"
    Public Function TBAnimalSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Dim param As NpgsqlParameter
        Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    'Select
                    myTransaction = conn.BeginTransaction

                    'sqlstr = "select * from exanimal"                 
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.AcceptChangesDuringUpdate = False
                    DataAdapter.SelectCommand = New NpgsqlCommand("funcselectexanimal", conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                    'Delete

                    DataAdapter.DeleteCommand = New NpgsqlCommand("funcdeleteexanimal", conn)
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure
                    DataAdapter.DeleteCommand.Parameters.Add("@exanimalid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "exanimalid").SourceVersion = DataRowVersion.Original

                    'Update

                    DataAdapter.UpdateCommand = New NpgsqlCommand("funcupdexanimal", conn)
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                    DataAdapter.UpdateCommand.Parameters.Add("@exanimaltype", NpgsqlTypes.NpgsqlDbType.Text, 0, "exanimaltype").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@exanimalid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "exanimalid").SourceVersion = DataRowVersion.Original
                    param = DataAdapter.UpdateCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate")
                    param.SourceVersion = DataRowVersion.Original
                    param.Direction = ParameterDirection.InputOutput

                    'Insert()
                    DataAdapter.InsertCommand = New NpgsqlCommand("funcinsertexanimal", conn)
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure
                    DataAdapter.InsertCommand.Parameters.Add("@exanimaltype", NpgsqlTypes.NpgsqlDbType.Text, 0, "exanimaltype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@exanimalid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "exanimalid").Direction = ParameterDirection.Output
                    DataAdapter.InsertCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate").Direction = ParameterDirection.Output

                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))

                    myTransaction.Commit()
                Catch ex As Exception
                    message = ex.Message
                    myTransaction.Rollback()
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

#Region "TBPet"
    Public Function TBPetSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim param As NpgsqlParameter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    myTransaction = conn.BeginTransaction
                    'Select
                    sqlstr = "select * from expet"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.AcceptChangesDuringUpdate = False
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                    'Delete
                    sqlstr = "Delete from expet where expetid=@expetid"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("@expetid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "expetid").SourceVersion = DataRowVersion.Original

                    'Update
                    sqlstr = "Update expet set exanimalid = @exanimalid,firstname=@firstname,lastname=@lastname,weight=@weight" &
                             " where expetid=@expetid and latestupdate=@latestupdate;select latestupdate from expet where expetid = @expetid"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Parameters.Add("@exanimalid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "exanimalid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@firstname", NpgsqlTypes.NpgsqlDbType.Text, 0, "firstname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@lastname", NpgsqlTypes.NpgsqlDbType.Text, 0, "lastname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@weight", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "weight").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@expetid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "expetid").SourceVersion = DataRowVersion.Original
                    param = DataAdapter.UpdateCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate")
                    param.SourceVersion = DataRowVersion.Original
                    param.Direction = ParameterDirection.InputOutput

                    sqlstr = "insert into expet(exanimalid,firstname,lastname,weight) values " & _
                             "(@exanimalid,@firstname,@lastname,@weight);" & _
                             "select currval('expet_expetid_seq') as expetid,latestupdate from expet where expetid = currval('expet_expetid_seq') ;"

                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Parameters.Add("@exanimalid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "exanimalid").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("@firstname", NpgsqlTypes.NpgsqlDbType.Text, 0, "firstname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@lastname", NpgsqlTypes.NpgsqlDbType.Text, 0, "lastname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@weight", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "weight").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@expetid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "expetid").Direction = ParameterDirection.Output
                    DataAdapter.InsertCommand.Parameters.Add("@latesupadte", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate").Direction = ParameterDirection.Output
                    RecordAffected = DataAdapter.Update(DataSet.Tables("TBExPet"))
                    myTransaction.Commit()
                Catch ex As Exception
                    message = ex.Message
                    myTransaction.Rollback()
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

#Region "TBPetBelonging"
    Public Function TBPetBelongingSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim param As NpgsqlParameter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    'Select
                    sqlstr = "select * from expetbelonging"

                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.AcceptChangesDuringUpdate = False
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                    'Delete
                    sqlstr = "Delete from expetbelonging where expetbelongingid=@expetbelongingid"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Parameters.Add("@expetbelongingid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "expetbelongingid").SourceVersion = DataRowVersion.Original

                    'Update
                    'sqlstr = "Update expetbelonging set expetid = @expetid,price=@price,petname=@petname" &
                    '         " where expetbelongingid=@expetbelongingid and latestupdate=@latestupdate;select latestupdate from expetbelonging where expetbelongingid = @expetbelongingid"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.UpdateCommand.Transaction = myTransaction

                    DataAdapter.UpdateCommand = New NpgsqlCommand("funcupdexpetanimal", conn)
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Parameters.Add("@expetbelongingid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "expetbelongingid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("@price", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@petname", NpgsqlTypes.NpgsqlDbType.Text, 0, "petname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@expetid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "expetid").SourceVersion = DataRowVersion.Current
                    param = DataAdapter.UpdateCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate")
                    param.SourceVersion = DataRowVersion.Original
                    param.Direction = ParameterDirection.InputOutput


                    sqlstr = "insert into expetbelonging(expetid,price,petname) values " & _
                             "(@expetid,@price,@petname);select currval('expetbelonging_expetbelongingid_seq') as expetbelongingid,latestupdate from expetbelonging where expetbelongingid = currval('expetbelonging_expetbelongingid_seq');"

                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Parameters.Add("@expetid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "expetid").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("@price", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@petname", NpgsqlTypes.NpgsqlDbType.Text, 0, "petname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@expetbelongingid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "expetbelongingid").Direction = ParameterDirection.Output
                    DataAdapter.UpdateCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate").Direction = ParameterDirection.Output

                    RecordAffected = DataAdapter.Update(DataSet.Tables("TBExPetBelonging"))
                    myTransaction.Commit()
                Catch ex As Exception
                    myTransaction.Rollback()
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


End Class
