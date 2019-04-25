Public Delegate Sub UpdateRecordDelegate()
Public Delegate Sub ReloadDelegate()
Public Delegate Function SaveChangesRecordDelegate(ByRef DataSet As DataSet, ByRef message As String, ByRef RecordAffected As Integer, ByVal continueupdateonerror As Boolean) As Boolean

Public Class HelperClass
    Public Shared DbAdapter1 As DbAdapter

    Public Shared Sub DeleteRecordHelper(ByVal datagridview1 As DataGridView, ByRef bindingsource1 As BindingSource, ByVal cm As CurrencyManager)
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                If datagridview1.SelectedRows.Count = 0 Then
                    bindingsource1.RemoveAt(cm.Position)
                Else
                    For Each a As DataGridViewRow In datagridview1.SelectedRows
                        bindingsource1.RemoveAt(a.Index)
                    Next
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub


    Public Shared Sub UpdateRecordHelper(ByRef dataset As DataSet, ByVal SaveChangesRecord As SaveChangesRecordDelegate, ByVal loaddata As ReloadDelegate, ByRef cm As CurrencyManager)
        Dim sb As New System.Text.StringBuilder
        Dim message As String = String.Empty
        Dim ra As Integer

        Dim DS = dataset.GetChanges
        If Not IsNothing(DS) Then
            If (SaveChangesRecord(DS, message, ra, True)) Then
                sb.Append(String.Format("Result: {0} Record{1} Affected.", ra, IIf(ra > 1, "s", "")) & vbCrLf & vbCrLf)
                'Two statement below need to show the errors
                dataset.Merge(DS)
                dataset.AcceptChanges()
            End If
            If DS.HasErrors Then
                ShowErrorHelper(sb, DS.Tables(0).Rows)
            End If
            MessageBox.Show(sb.ToString)
            'Move Cursor to Current Position
            'Those statement to solve combobox issue changing record
            Dim position = cm.Position
            loaddata()
            cm.Position = position

        End If
    End Sub

    Private Shared Sub ShowErrorHelper(ByRef sb As System.Text.StringBuilder, ByVal DataRowColl As DataRowCollection)
        Dim myquery = From row As DataRow In DataRowColl
                              Where row.RowError <> ""
                              Select row.RowError

        Dim i As Integer
        sb.Append(String.Format("Found {0} error(s)", myquery.Count) & vbCrLf)

        For Each myerro In myquery
            i += 1
            sb.Append(String.Format("Error #{0} {1} {2}", i, vbCrLf, myerro) & vbCrLf)
        Next
        sb.Append(String.Format(vbCrLf & "Data will refresh shortly."))
    End Sub

    Public Shared Sub CheckFormClosingHelper(ByRef dataset As DataSet, ByVal UpdateRecord As UpdateRecordDelegate, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        If dataset.HasChanges Then
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

    Public Shared Sub columnheadermouseclickHelper(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs, ByRef bindingsource1 As BindingSource)
        Dim dg = CType(sender, DataGridView)
        Dim headercell = CType(dg.Columns(e.ColumnIndex).HeaderCell, DataGridViewColumnHeaderCell)
        Dim sort As String = String.Empty
        If (headercell.SortGlyphDirection = SortOrder.Descending) Then
            headercell.SortGlyphDirection = SortOrder.Ascending
            sort = SortDirection.Asc.ToString
        Else
            headercell.SortGlyphDirection = SortOrder.Descending
            sort = SortDirection.Desc.ToString
        End If
        bindingsource1.Sort = String.Format("{0} {1}", dg.Columns(e.ColumnIndex).DataPropertyName, sort)
    End Sub

    Public Shared Sub LoadToolstripFilterSort(ByVal hidetoolbar As HideToolbarDelegate, ByVal DG As DataGridView, ByRef mypanel1 As UCSortTx, ByRef toolstrip As ToolStrip, ByRef mypanel As UCFilterTx)
        'Dim myaction As HideToolbarDelegate = AddressOf toolstripvisible
        'Dim myheader As New UCHeader(myaction)
        Dim myheader As New UCHeader(hidetoolbar)
        myheader.ToolStripLabel1.Text = "Advance Filter && Sort"

        mypanel = New UCFilterTx(DG)
        Dim myhost = New ToolStripControlHost(mypanel)
        mypanel1 = New UCSortTx(DG)
        Dim myhost2 = New ToolStripControlHost(mypanel1)
        Dim myhost3 = New ToolStripControlHost(myheader)

        toolstrip.Items.Add(myhost3)
        toolstrip.Items.Add(myhost)
        toolstrip.Items.Add(myhost2)
        toolstrip.Items(0).Margin = New Padding(0, 0, 0, 4)
        toolstrip.Items(1).Margin = New Padding(0)
        toolstrip.Items(2).Margin = New Padding(0)
    End Sub
End Class

Public Enum Direction
    MoveDown = 1
    MoveUp = -1
End Enum