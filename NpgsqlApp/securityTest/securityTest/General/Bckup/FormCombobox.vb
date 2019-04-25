Imports securityTest.HelperClass
Public Class FormCombobox
    Dim Dataset As DataSet
    Dim sqlstr As String
    Dim myCheck As New ArrayList
    Dim CM As CurrencyManager
    Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
    Private newRecord As Boolean = False

    Friend WithEvents AnimalId As DataGridViewComboBoxColumn
    Friend WithEvents LatestUpdate As DJLib.DataGridViewCalendarColumn
    Friend WithEvents Weight As DJLib.DataGridViewBarGraphColumn



    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        'BindingSource1 = New BindingSource
        loadData()

        'DataGridViewBindingColumn()
    End Sub

    Private Sub loadData()
        BindingSource1 = New BindingSource
        Dataset = New DataSet
        DataGridView1.DataSource = BindingSource1

        BindingNavigator1.BindingSource = BindingSource1
        sqlstr = "select * from expet order by expetid;select * from exanimal order by exanimaltype"

        DbAdapter1.TbgetDataSet(sqlstr, Dataset)
        Dataset.Tables(0).TableName = "TBExPet"
        Dataset.Tables(1).TableName = "TBExAnimal"
        BindingSource1.DataSource = Dataset.Tables("TBExPet")
        BindingSource2.DataSource = Dataset.Tables("TBExAnimal")

        BindingSource1.DataSource = Dataset.Tables(0)
        BindingSource2.DataSource = Dataset.Tables(1)
        DataGridViewBindingColumn()
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
    End Sub
    Private Sub DataGridViewBindingColumn()

        'hide original columns
        For i = 0 To 1
            DataGridView1.Columns(i).Visible = False
        Next
        'DataGridView1.Columns(5).Visible = False

        Me.AnimalId = New DataGridViewComboBoxColumn 'CreateComboboxColumn
        Me.AnimalId.DataPropertyName = "exanimalid"
        Me.AnimalId.HeaderText = "Animal Type"
        Me.AnimalId.DropDownWidth = 160
        Me.AnimalId.Width = 90
        Me.AnimalId.MaxDropDownItems = 6
        Me.AnimalId.FlatStyle = FlatStyle.Flat
        Me.AnimalId.DataPropertyName = "exanimalid"
        Me.AnimalId.DataSource = BindingSource2
        Me.AnimalId.ValueMember = "exanimalid"
        Me.AnimalId.DisplayMember = "exanimaltype"

        'Me.AnimalId.FilteringEnabled = False
        Me.AnimalId.HeaderText = "Animal Type"
        Me.AnimalId.Name = "exanimaltype"
        Me.AnimalId.Resizable = Windows.Forms.DataGridViewTriState.True
        Me.AnimalId.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.AnimalId.Width = 87

        Me.DataGridView1.Columns.Insert(4, AnimalId)


        Me.LatestUpdate = New DJLib.DataGridViewCalendarColumn
        Me.LatestUpdate.DataPropertyName = "latestupdate"
        Me.LatestUpdate.HeaderText = "Latest Update"
        Me.LatestUpdate.Name = "latestupdate"
        Me.LatestUpdate.Resizable = Windows.Forms.DataGridViewTriState.True
        Me.LatestUpdate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.LatestUpdate.Width = 87
        Me.DataGridView1.Columns.Insert(5, LatestUpdate)

        Me.Weight = New DJLib.DataGridViewBarGraphColumn
        Me.Weight.DataPropertyName = "weight"
        Me.Weight.HeaderText = "Weight"
        Me.Weight.Name = "Weight"
        Me.Weight.Resizable = Windows.Forms.DataGridViewTriState.True
        Me.Weight.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Weight.Width = 87
        Me.DataGridView1.Columns.Add(Weight)


    End Sub
    Private Function CreateComboboxColumn() As DataGridViewComboBoxColumn
        Dim column As New DataGridViewComboBoxColumn
        With column
            .DataPropertyName = "exanimalid"
            .HeaderText = "Animal Type"
            .DropDownWidth = 160
            .Width = 90
            .MaxDropDownItems = 6
            .FlatStyle = FlatStyle.Flat
        End With
        Return column
    End Function
    Public Sub UpdateRecord()
        Try
            Dim ra As Integer
            Dim message As String = String.Empty
            Dim sb As New System.Text.StringBuilder

            BindingSource1.EndEdit()
            Dim ds2 = Dataset.GetChanges
            If DbAdapter1.TBPetSaveChanges(ds2, message, ra) Then
                sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected. (Table Animal)")
                Dataset.Merge(ds2)
                Dataset.AcceptChanges()
            Else
            End If

            If Dataset.HasErrors Then
                If sb.Length > 0 Then
                    sb.Append(vbCrLf)
                End If
                sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
                MessageBox.Show(sb.ToString)
                loadData()
            Else
                MessageBox.Show(sb.ToString)
            End If

        Catch ex As Exception
        End Try
        'loadData() No need anymore
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        UpdateRecord()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        MessageBox.Show(e.Exception.Message)
    End Sub


    Private Sub RefreshToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripButton.Click
        If Dataset.HasChanges Then
            Dim response As Windows.Forms.DialogResult
            response = MessageBox.Show("You have unsaved data. Do you want to store to the database?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case response
                Case Windows.Forms.DialogResult.Yes
                    UpdateRecord()
                    loadData()
                Case Windows.Forms.DialogResult.Cancel

                Case Windows.Forms.DialogResult.No
                    loadData()
            End Select
        Else
            loadData()
        End If

    End Sub

    Private Sub BindingNavigatorDeleteItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorDeleteItem.Click
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                If DataGridView1.SelectedRows.Count = 0 Then
                    BindingSource1.RemoveAt(CM.Position)
                Else
                    For Each a As DataGridViewRow In DataGridView1.SelectedRows
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
            Dim response As Windows.Forms.DialogResult
            response = MessageBox.Show("You have unsaved data. Do you want to store to the database?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case response
                Case Windows.Forms.DialogResult.Yes
                    Me.UpdateRecord()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
                Case Windows.Forms.DialogResult.No
            End Select
        End If
    End Sub

    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        BindingSource1.AddNew()
    End Sub


    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim dg As DataGridView = CType(sender, DataGridView)
    End Sub



End Class