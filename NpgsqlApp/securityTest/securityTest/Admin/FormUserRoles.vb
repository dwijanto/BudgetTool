Imports System.Data
Imports HR.HelperClass
Public Class FormUserRoles
    Dim Dataset As DataSet
    Dim AllRoles() As String
    Dim inRoles() As String
    Dim CurrentUser As String
    Dim cm As CurrencyManager
    Dim myCheck As New ArrayList
    Dim deletedUser As New ArrayList
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        BindingSource1 = New BindingSource
        loadData()
        InitRoles()
    End Sub

    Private Sub loadData()
        Dataset = New DataSet
        BindingNavigator1.BindingSource = BindingSource1
        DbAdapter1.TbgetDataSet("select * from users order by username;select * from region order by regionname", Dataset)
        Dataset.Tables(0).TableName = "users"
        Dataset.Tables(1).TableName = "region"
        BindingSource1.DataSource = Dataset.Tables("users")
        bindingDataGridView()
        cm = CType(BindingContext(BindingSource1), CurrencyManager)

    End Sub

    Private Sub InitRoles()
        AllRoles = DJLib.AppConfig.RoleAttribute.GetAllRoles

        Array.Sort(AllRoles)

        For i = 0 To AllRoles.Length - 1
            Dim mycheckbox As New CheckBox
            mycheckbox.Text = AllRoles(i)
            myCheck.Add(mycheckbox)

            FlowLayoutPanel1.Controls.Add(myCheck(i))
            AddHandler CType(myCheck(i), CheckBox).Click, AddressOf onCheckboxClicked
        Next
    End Sub
    Private Sub bindingDataGridView()
        DataGridView1.Columns.Clear()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = BindingSource1


        Dim Col1 As New DataGridViewTextBoxColumn()
        With Col1
            .DataPropertyName = "username"
            .Name = "col1"
            .HeaderText = "User Name"
            .Width = 100
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With

        Dim Col2 As New DataGridViewTextBoxColumn()
        With Col2
            .DataPropertyName = "email"
            .Name = "col2"
            .HeaderText = "Email"
            .Width = 200
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With

        Dim Col3 As New DataGridViewTextBoxColumn()
        With Col3
            .DataPropertyName = "comments"
            .Name = "col3"
            .HeaderText = "Comments"
            .Width = 100
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With

        Dim col4 As New DataGridViewCheckBoxColumn()
        With col4
            .HeaderText = "Is Approved"
            .Name = "Col4"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .FlatStyle = FlatStyle.Standard
            '.CellTemplate = New DataGridViewCheckBoxCell()
            '.CellTemplate.Style.BackColor = Color.Beige
            .DataPropertyName = "isapproved"
        End With

        Dim col5 As DataGridViewComboBoxColumn
        col5 = CreateComboBoxColumn()
        SetAlternateChoicesUsingDataSource(col5)
        'col5.HeaderText = "Region Name"

        With DataGridView1
            .Columns.Insert(0, Col1)
            .Columns.Insert(1, Col2)
            .Columns.Insert(2, Col3)
            .Columns.Insert(3, Col4)
            .Columns.Insert(4, col5)
            '.Columns.Insert(5, Col6)
        End With

        
    End Sub
    Private Function CreateComboBoxColumn() _
        As DataGridViewComboBoxColumn
        Dim column As New DataGridViewComboBoxColumn()

        With column
            .DataPropertyName = "regionid"
            .HeaderText = "Region Name"
            .DropDownWidth = 160
            .Width = 160
            .MaxDropDownItems = 4
            .FlatStyle = FlatStyle.Flat
        End With
        Return column
    End Function

    Private Sub SetAlternateChoicesUsingDataSource( _
        ByVal comboboxColumn As DataGridViewComboBoxColumn)
        With comboboxColumn
            .DataSource = Dataset.Tables(1)
            .ValueMember = "regionid"
            .DisplayMember = "regionname"
        End With
    End Sub



    Private Sub bindingDataGridView1()
        DataGridView1.DataSource = BindingSource1

        DataGridView1.AutoGenerateColumns = True
        For i = 0 To 21
            DataGridView1.Columns(i).Visible = False
        Next
        DataGridView1.Columns(1).Visible = True
        DataGridView1.Columns(3).Visible = True
        DataGridView1.Columns(4).Visible = True
        DataGridView1.Columns(8).Visible = True
    End Sub
    Public Sub UpdateRecord()
        Dim ra As Integer
        Dim message As String = String.Empty
        BindingSource1.EndEdit()

        Dim sb As New System.Text.StringBuilder
        Dim ds2 = Dataset.GetChanges
        If Not IsNothing(ds2) Then
            If DbAdapter1.TBUsersSaveChanges(ds2, message, ra) Then
                sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
                Dataset.Merge(ds2)
                Dataset.AcceptChanges()
            End If
            If Dataset.HasErrors Then
                sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
                MessageBox.Show(sb.ToString)
                loadData()
            Else
                If sb.ToString <> "" Then
                    MessageBox.Show(sb.ToString)
                End If
                For Each deluser As String In deletedUser
                    DbAdapter1.DropUserDb(deluser)
                Next

            End If
        Else
            MessageBox.Show("Nothing to save")
        End If
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        UpdateRecord()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message & "Some Records has been deleted by other user.Refreshing record in progress")
    End Sub

    Private Sub RefreshToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripButton.Click
        If Dataset.HasChanges Then
            Dim datasetchanges As DataSet
            datasetchanges = Dataset.GetChanges()
            Dim response As Windows.Forms.DialogResult
            response = MessageBox.Show(datasetchanges.Tables(0).Rows.Count & " unsaved data. Do you want to store to the database?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case response
                Case Windows.Forms.DialogResult.Yes
                    UpdateRecord()
                    loadData()
                Case Windows.Forms.DialogResult.Cancel

                Case Windows.Forms.DialogResult.No
                    loadData()
            End Select
        End If
    End Sub

    Private Sub BindingNavigatorDeleteItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorDeleteItem.Click
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            deletedUser.Clear()
            Try
                If DataGridView1.SelectedRows.Count = 0 Or DataGridView1.SelectedRows.Count = 1 Then
                    deletedUser.Add(DataGridView1.Rows(cm.Position).Cells(0).Value)
                    BindingSource1.RemoveAt(cm.Position)
                Else
                    For Each a As DataGridViewRow In DataGridView1.SelectedRows
                        deletedUser.Add(DataGridView1.Rows(a.Index).Cells(0).Value)
                        BindingSource1.RemoveAt(a.Index)
                    Next
                End If
                UpdateRecord()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub FormUser_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Dataset.HasChanges Then
            Dim datasetchanges As DataSet
            datasetchanges = Dataset.GetChanges()
            Dim response As Windows.Forms.DialogResult
            response = MessageBox.Show(datasetchanges.Tables(0).Rows.Count & " unsaved data. Do you want to store to the database?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case response
                Case Windows.Forms.DialogResult.Yes
                    UpdateRecord()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
                Case Windows.Forms.DialogResult.No
            End Select
        End If
    End Sub


    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        Dim myform = New DialogCreateUser
        Dim result As System.Windows.Forms.DialogResult = myform.ShowDialog()
        loadData()
    End Sub

    Private Sub ResetPasswordToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetPasswordToolStripButton.Click
        If MessageBox.Show("Reset Passsword selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                Dim newpassword As String = String.Empty
                If DataGridView1.SelectedRows.Count <= 1 Then
                    Dim user = DataGridView1.Rows(cm.Position).Cells(0).Value
                    DbAdapter1.ChangePasswordDb(user, "password123")
                    newpassword = DJLib.AppConfig.MembershipService.ResetPassword(user, "password123")



                Else
                    For Each a As DataGridViewRow In DataGridView1.SelectedRows
                        Dim user = DataGridView1.Rows(a.Index).Cells(0).Value
                        newpassword = DJLib.AppConfig.MembershipService.ResetPassword(user, "password123")
                        DbAdapter1.ChangePasswordDb(user, "password123")
                    Next
                End If

                MsgBox("New Password: " & newpassword)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Try
            Dim dg As DataGridView = CType(sender, DataGridView)
            CurrentUser = DataGridView1.Rows(cm.Position).Cells(0).Value
            Label1.Text = String.Format("Add ""{0}"" to Roles", CurrentUser)
            'assign checkbox
            inRoles = DJLib.AppConfig.RoleAttribute.GetRolesForUser(CurrentUser)
            Dim chk As CheckBox
            For Each mycontrol In FlowLayoutPanel1.Controls
                If TypeOf (mycontrol) Is CheckBox Then
                    chk = CType(mycontrol, CheckBox)
                    chk.Checked = inRoles.Contains(chk.Text)
                End If
            Next

        Catch ex As Exception
        End Try
    End Sub

    Private Sub onCheckboxClicked(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = CType(sender, CheckBox)
        Dim users() As String = {CurrentUser}
        Dim roles() As String = {chk.Text}
        If chk.Checked Then
            DJLib.AppConfig.RoleAttribute.AddUsersToRoles(users, roles)
        Else
            DJLib.AppConfig.RoleAttribute.RemoveUsersFromRoles(users, roles)
        End If
    End Sub



End Class