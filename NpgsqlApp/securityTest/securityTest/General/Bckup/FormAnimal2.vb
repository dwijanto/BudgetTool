Imports securityTest.HelperClass
Public Class FormAnimal2
    Dim Dataset As DataSet
    Dim sqlstr As String
    Dim myCheck As New ArrayList
    Dim CM As CurrencyManager
    Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
    Private newRecord As Boolean = False

    Friend WithEvents AnimalType As DJLib.DataGridViewAutoFilterTextBoxColumn
    Friend WithEvents LatestUpdate As DJLib.DataGridViewAutoFilterTextBoxColumn

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        BindingSource1 = New BindingSource
        loadData()

        CM = CType(BindingContext(BindingSource1), CurrencyManager)
        DataGridViewBindingColumn()
    End Sub

    Private Sub loadData()

        Dataset = New DataSet
        DataGridView1.DataSource = BindingSource1
        DataGridView2.DataSource = BindingSource2
        DataGridView3.DataSource = BindingSource3
        BindingNavigator1.BindingSource = BindingSource1

        If DbAdapter1.GetRefCursorAnimal(Dataset) Then
            Dataset.Tables(0).TableName = "TBExAnimal"
            Dataset.Tables(1).TableName = "TBExPet"
            Dataset.Tables(2).TableName = "TBExPetBelonging"
            BindingSource1.DataSource = Dataset.Tables("TBExAnimal")

            Dataset.Relations.Add("AniToPet", Dataset.Tables("TBExAnimal").Columns("exanimalid"), Dataset.Tables("TBExPet").Columns("exanimalid"))
            Dataset.Relations.Add("PetToBelonging", Dataset.Tables("TBExPet").Columns("expetid"), Dataset.Tables("TBExPetBelonging").Columns("expetid"))

            BindingSource2.DataSource = BindingSource1
            BindingSource2.DataMember = "AniToPet"
            BindingSource3.DataSource = BindingSource2
            BindingSource3.DataMember = "PetToBelonging"
        End If
        Dim highlightcellstyle As New DataGridViewCellStyle
        highlightcellstyle.BackColor = Color.Red
        With DataGridView1
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            '.GridColor = Color.BlueViolet
            '.BorderStyle = BorderStyle.Fixed3D
            '.CellBorderStyle = DataGridViewCellBorderStyle.None
            '.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            '.DefaultCellStyle.BackColor = Color.Pink
            '.DefaultCellStyle.Font = New Font("Tahoma", 12)
            '.Rows(3).DefaultCellStyle = highlightcellstyle
            '.Columns("latestupdate").DefaultCellStyle.Format = "d"
            '.DefaultCellStyle.SelectionForeColor = Color.Yellow
            '.DefaultCellStyle.SelectionBackColor = Color.Black
            '.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige


        End With
        'With DataGridView1.RowTemplate
        '    .DefaultCellStyle.BackColor = Color.Chocolate
        '    .Height = 15
        '    .MinimumHeight = 20
        'End With


        With DataGridView2
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End With
        With DataGridView3
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End With

    End Sub
    Private Sub DataGridViewBindingColumn()

        'hide original columns
        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).Visible = False
        Next
        SuspendLayout()

        Me.AnimalType = New DJLib.DataGridViewAutoFilterTextBoxColumn
        Me.AnimalType.DataPropertyName = "exanimaltype"
        Me.AnimalType.FilteringEnabled = False
        Me.AnimalType.HeaderText = "Animal Type"
        Me.AnimalType.Name = "exanimaltype"
        Me.AnimalType.Resizable = Windows.Forms.DataGridViewTriState.True
        Me.AnimalType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.AnimalType.Width = 87

        Me.DataGridView1.Columns.Add(AnimalType)


        Me.LatestUpdate = New DJLib.DataGridViewAutoFilterTextBoxColumn
        Me.LatestUpdate.DataPropertyName = "latestupdate"
        Me.LatestUpdate.FilteringEnabled = False
        Me.LatestUpdate.HeaderText = "Latest Update"
        Me.LatestUpdate.Name = "latestupdate"
        Me.LatestUpdate.Resizable = Windows.Forms.DataGridViewTriState.True
        Me.LatestUpdate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.LatestUpdate.Width = 87
        Me.DataGridView1.Columns.Add(LatestUpdate)
        ResumeLayout(False)
        'Me.DataGridView1.RowTemplate.HeaderCell = New ClassRowHeaderImage

    End Sub
    Public Sub UpdateRecord()
        Try
            Dim ra As Integer
            Dim message As String = String.Empty
            Dim sb As New System.Text.StringBuilder
            'BindingSource1.EndEdit()

            UpdateRecordDetail(1, BindingSource1, Dataset, sb, message, ra)
            UpdateRecordDetail(2, BindingSource2, Dataset, sb, message, ra)
            UpdateRecordDetail(3, BindingSource3, Dataset, sb, message, ra)

            If Dataset.HasErrors Then
                If sb.Length > 0 Then
                    sb.Append(vbCrLf)
                End If
                sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
                loadData()
            End If

            Dataset.AcceptChanges()
            MessageBox.Show(sb.ToString)
            'For i = 0 To DataGridView1.Rows.Count
            '    DataGridView1.Rows(i).HeaderCell = New ClassRowHeaderImage(ClassRowHeaderImage.CurrentRowState.Normal)
            'Next

        Catch ex As Exception
        End Try      
    End Sub

    Private Sub UpdateRecordDetail(ByVal TableNo As Integer, ByRef bs As BindingSource, ByRef ds As DataSet, ByRef sb As System.Text.StringBuilder, ByRef message As String, ByRef ra As Integer)
        bs.EndEdit()
        Dim ds2 = ds.GetChanges
        Try
            Dim dr2 As DataRow = CType(ds2.Tables(TableNo - 1).Rows(0), DataRow)
            If ds2 IsNot Nothing AndAlso dr2.RowState <> DataRowState.Unchanged Then
                Select Case TableNo
                    Case 1
                       DbAdapter1.TBAnimalSaveChanges(ds2, message, ra)
                    Case 2
                        If DbAdapter1.TBPetSaveChanges(ds2, message, ra) Then
                            If sb.Length > 0 Then
                                sb.Append(vbCrLf)
                            End If
                        End If
                    Case 3
                        If DbAdapter1.TBPetBelongingSaveChanges(ds2, message, ra) Then
                            If sb.Length > 0 Then
                                sb.Append(vbCrLf)
                            End If
                        End If
                End Select
                sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected. (" & ds2.Tables(TableNo - 1).ToString & ")")
                ds.Merge(ds2)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        UpdateRecord()
    End Sub
    Private Sub DataGridView1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        Dim headertext As String = DataGridView1.Columns(e.ColumnIndex).HeaderText
        If Not headertext.Equals("Animal Type") Then Return
        If (String.IsNullOrEmpty(e.FormattedValue.ToString())) Then
            DataGridView1.Rows(e.RowIndex).ErrorText = "Animal Type must not empty"
            e.Cancel = True
        End If
    End Sub
    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        DataGridView1.Rows(e.RowIndex).ErrorText = String.Empty
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub

    Private Sub DataGridView2_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub
    Private Sub DataGridView3_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView3.DataError
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
                    UpdateRecord()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
                Case Windows.Forms.DialogResult.No
            End Select
        End If
    End Sub

    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        BindingSource1.AddNew()
        'DataGridView1.Rows.Add()
        'DataGridView1.CurrentRow.HeaderCell = New ClassRowHeaderImage(ClassRowHeaderImage.CurrentRowState.New)
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim dg As DataGridView = CType(sender, DataGridView)
    End Sub

    Private Sub StoreProcedureToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StoreProcedureToolStripButton.Click
        MsgBox(DbAdapter1.InsertAnimal("Hi"))
        loadData()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If e.ColumnIndex <> -1 Then
            If DataGridView1.Columns(e.ColumnIndex).Name = "latestupdate" Then
                If e.Value IsNot Nothing Then
                    Dim thedate As Date = DateTime.Parse(e.Value.ToString)
                    e.Value = thedate.ToString("dd-MMM-yyyy")
                End If
            End If
        End If
    End Sub



    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        'DataGridView1.CurrentRow.HeaderCell = New ClassRowHeaderImage(ClassRowHeaderImage.CurrentRowState.Modified)
    End Sub
    Private Sub DataGridView2_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged
        DataGridView2.CurrentRow.HeaderCell = New ClassRowHeaderImage(ClassRowHeaderImage.CurrentRowState.Modified)
    End Sub

    Private Sub DataGridView1_DefaultValuesNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DataGridView1.DefaultValuesNeeded
        With e.Row
            .Cells(3).Value = "HI"
            .Cells(4).Value = DateTime.Now.ToString
        End With

        'Me.DataGridView1.HeaderCell()
        'e.Row.HeaderCell = New ClassRowHeaderImage(ClassRowHeaderImage.CurrentRowState.New)
    End Sub

End Class