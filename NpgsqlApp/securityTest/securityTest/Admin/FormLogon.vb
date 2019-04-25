Imports System.Configuration
Imports System.Web.Security
Imports DJLib.AppConfig
Imports HR.HelperClass
Public Class FormLogon

    Private Sub btnLogon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogon.Click

        dbtools1 = New DJLib.Dbtools
        BudgetYear = CDate(Today.Year + 1 & "-" & Today.Month.ToString & "-" & Today.Day.ToString)
        dbtools1.Userid = TextBox1.Text
        dbtools1.Password = TextBox2.Text
        DJLib.AppConfig.dbTools = dbtools1
        Dim MembershipService As New DJLib.AccountMembershipService
        Dim RoleAttribute As DJLib.NpgsqlRoleProvider
        RoleAttribute = New DJLib.NpgsqlRoleProvider("npgsqlProvider", DJLib.AppConfig.CreateConfig)
        DJLib.AppConfig.RoleAttribute = RoleAttribute
        DJLib.AppConfig.MembershipService = MembershipService

        DbAdapter1 = New DbAdapter(dbtools1.getConnectionString.ToString)
        'if no user then create superuser
        'checkTableUser()

        If (TextBox1.Text.Length <= 0 OrElse TextBox2.Text.Length <= 0) Then
            MessageBox.Show("No Valid User Name Or Password", "MissingInformation")
            TextBox2.Text = String.Empty
            Exit Sub
        End If
        Try

            If DJLib.AppConfig.MembershipService.ValidateUser(TextBox1.Text, TextBox2.Text) Then
                HR.HelperClass.ChangeUser = True
                DJLib.AppConfig.Identity = DJLib.MyIdentity.CreateIdentity(TextBox1.Text)
                DJLib.AppConfig.Principal = DJLib.MyPrincipal.CreatePrincipal(DJLib.AppConfig.Identity)

                dbtools1.Region = DbAdapter1.getRegionShortName(TextBox1.Text)
                dbtools1.RegionId = DbAdapter1.getRegionID(TextBox1.Text)
                dbtools1.RegionName = DbAdapter1.getRegionName(TextBox1.Text)
                Dim connectionString As String = dbtools1.getConnectionString.ToString
                Dim connectionstrings() As String = connectionString.Split(";")
                ConnectionStringCollections = New Collection
                For i = 0 To (connectionstrings.Length - 1)
                    Dim mystrs() As String = connectionstrings(i).Split("=")
                    ConnectionStringCollections.Add(mystrs(1), mystrs(0))
                Next i

                getRegionShortname()

                Try
                    loglogin(TextBox1.Text)
                Catch ex As Exception
                End Try

                'check for group

                FormMenu.Show()
                Me.Close()
                Me.Dispose()
            Else
                HR.HelperClass.ChangeUser = False
                MessageBox.Show("The user name or password provided is incorrect.", "Missing Information")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Missing Information")
            Me.Close()
            Application.Exit()
        End Try

    End Sub

    Private Sub loglogin(ByVal userid As String)
        Dim folders = Application.StartupPath.Split("\")
        Dim applicationname As String = folders(folders.Length - 1)
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.


    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub checkTableUser()
        Dim dataset As New DataSet
        If Not DbAdapter1.TbgetDataSet("select * from users order by username", dataset) Then
            Dim status = DJLib.AppConfig.MembershipService.CreateUser("admin", "admin", "dwijanto@yahoo.com")
        End If
        If dataset.Tables(0).Rows.Count = 0 Then
            Dim status = DJLib.AppConfig.MembershipService.CreateUser("admin", "admin", "dwijanto@yahoo.com")
        End If
    End Sub

    Private Sub FormLogon_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'DJLib.AppConfig.dbTools = dbtools1
        'Dim MembershipService As New DJLib.AccountMembershipService
        'Dim RoleAttribute As DJLib.NpgsqlRoleProvider
        'RoleAttribute = New DJLib.NpgsqlRoleProvider("npgsqlProvider", DJLib.AppConfig.CreateConfig)
        'DJLib.AppConfig.RoleAttribute = RoleAttribute
        'DJLib.AppConfig.MembershipService = MembershipService


    End Sub

    Private Sub getRegionShortname()
        Dim DataSet1 As New DataSet
        Dim message As String = String.Empty
        If Not dbtools1.getDataSet("select regionid,regionshortname from region", DataSet1, message) Then
            Throw New System.Exception(message)
        End If
        RegionDict = New Dictionary(Of Integer, String)
        Dim query = From rec In DataSet1.Tables(0)
                    Select rec
                    Order By rec.Item("regionid")
        For Each dr In query
            RegionDict.Add(dr.Item("regionid"), dr.Item("regionshortname"))
        Next

        MonthToIntDict = New Dictionary(Of String, Integer)

        Dim mymonth() As String = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
        For i = 1 To 12
            MonthToIntDict.Add(mymonth(i - 1), i)
        Next

        CurrencyDict = New Dictionary(Of String, String)
        Dim region() As String = {"Hong Kong", "Shenzhen", "Taiwan", "Philippine"}
        Dim curtype() As String = {"HKD", "RMB", "NTD", "PHD"}
        For i = 0 To region.Count - 1
            CurrencyDict.Add(region(i), curtype(i))
        Next
    End Sub

End Class
