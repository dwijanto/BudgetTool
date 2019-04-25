Imports System.Configuration
Imports System.Web.Security
Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.AppConfig
Imports securityTest.HelperClass
Public Class FormLogon

    Dim dbtools = New Dbtools("admin", "admin")

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogon.Click
        If (TextBox1.Text.Length <= 0 OrElse TextBox2.Text.Length <= 0) Then
            MessageBox.Show("No Valid User Name Or Password", "MissingInformation")
            TextBox2.Text = String.Empty
            Exit Sub
        End If
        Try
            If DJLib.AppConfig.MembershipService.ValidateUser(TextBox1.Text, TextBox2.Text) Then
                DJLib.AppConfig.Identity = DJLib.MyIdentity.CreateIdentity(TextBox1.Text)
                DJLib.AppConfig.Principal = DJLib.MyPrincipal.CreatePrincipal(DJLib.AppConfig.Identity)

                Dim connectionString As String = dbtools.getConnectionString.ToString
                Dim connectionstrings() As String = connectionString.Split(";")
                ConnectionStringCollections = New Collection
                For i = 0 To (connectionstrings.Length - 1)
                    Dim mystrs() As String = connectionstrings(i).Split("=")
                    ConnectionStringCollections.Add(mystrs(1), mystrs(0))
                Next i
                FormMenu.Show()
                Me.Close()
            Else
                MessageBox.Show("The user name or password provided is incorrect.", "Missing Information")
            End If
        Catch ex As Exception

            MessageBox.Show(ex.Message, "Missing Information")
            Me.Close()
        End Try

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        DJLib.AppConfig.dbTools = dbtools
        Dim MembershipService As New DJLib.AccountMembershipService
        Dim RoleAttribute As DJLib.NpgsqlRoleProvider
        RoleAttribute = New DJLib.NpgsqlRoleProvider("npgsqlProvider", DJLib.AppConfig.CreateConfig)
        DJLib.AppConfig.RoleAttribute = RoleAttribute
        DJLib.AppConfig.MembershipService = MembershipService

        DbAdapter1 = New DbAdapter(dbtools.getConnectionString.ToString)
        'if no user then create superuser
        checkTableUser()

       
    End Sub

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim status = DJLib.AppConfig.MembershipService.CreateUser("ari2", "ari2", "ari2@yahoo.com")
    '    Try
    '        For Each Str As String In DJLib.AppConfig.RoleAttribute.GetRolesForUser(DJLib.AppConfig.Identity.Name)
    '            MsgBox(Str)
    '        Next
    '        MsgBox(DJLib.AppConfig.Principal.IsInRole("members"))
    '    Catch exx As NullReferenceException
    '        MsgBox(exx.Message)
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try

    'End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub checkTableUser()
        Dim dataset As New DataSet
        DbAdapter1.TbgetDataSet("select * from users order by username", dataset)
        If dataset.Tables(0).Rows.Count = 0 Then
            Dim status = DJLib.AppConfig.MembershipService.CreateUser("admin", "admin", "dwijanto@yahoo.com")
        End If
    End Sub

End Class
