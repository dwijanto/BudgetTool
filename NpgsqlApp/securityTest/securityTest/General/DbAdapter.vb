Imports Npgsql

Partial Public Class DbAdapter


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
                    sqlstr = "Update users set username = @username,applicationname = @applicationname,email=@email,comments=@comments,isapproved=@isapproved,regionid=@regionid" &
                             " where pkid=@pkid"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Text, 0, "username").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").Value = applicationname
                    DataAdapter.UpdateCommand.Parameters.Add("@pkid", NpgsqlTypes.NpgsqlDbType.Uuid, 0, "pkid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("@email", NpgsqlTypes.NpgsqlDbType.Text, 0, "email").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@comments", NpgsqlTypes.NpgsqlDbType.Text, 0, "comments").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@isapproved", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isapproved").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@regionid", NpgsqlTypes.NpgsqlDbType.Integer, 0, "regionid").SourceVersion = DataRowVersion.Current

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
