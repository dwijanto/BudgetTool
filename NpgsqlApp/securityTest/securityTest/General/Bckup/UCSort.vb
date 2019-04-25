Public Class UCSort
    Dim bs As New BindingSource
    Dim DG As DataGridView
    Dim _sort As String
    Dim mylist As New List(Of String)

    Public Property Sort As String
        Get
            Return _sort
        End Get
        Set(ByVal value As String)
            _sort = value
        End Set
    End Property

    Public Sub New(ByVal DG As DataGridView)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        BS = CType(DG.DataSource, BindingSource)
        Me.DG = DG
        '_sort = Nothing
        InitDataLayout()
    End Sub

    Private Sub InitDataLayout()
        BindFilterFields()
    End Sub
    Private Sub BindFilterFields()
        Dim cols As List(Of FieldClass) = New List(Of FieldClass)
        If Dg.Columns.Count > 0 Then

            For i = 0 To Dg.Columns.Count - 1
                If Dg.Columns(i).Visible Then
                    cols.Add(New FieldClass With {.id = DG.Columns(i).DataPropertyName,
                                                  .name = DG.Columns(i).HeaderText,
                                                  .ColumnType = DG.Columns(i).GetType.Name,
                                                  .ColumnIndex = i})
                End If
            Next
            'cols.Sort()
            cols.Insert(0, New FieldClass With {.id = "None", .name = "None", .ColumnType = "None"})
        End If
        SortCombobox.DataSource = cols
        SortCombobox.DisplayMember = "Name"
        SortCombobox.SelectedItem = "id"
    End Sub

    Private Sub executeSort()
        bs = CType(DG.DataSource, BindingSource)
        Dim mysort As String = String.Empty
        mylist.Clear()
        If CheckBox1.Checked Then
            Dim abc() = Split(bs.Sort.ToString, ",")
            mylist.Clear()
            For i = 0 To abc.Count - 1
                mylist.Add(abc(i))
            Next
        Else
            Sort = ""
        End If

        If bs.List.Count <= 0 OrElse SortCombobox.Items.Count <= 0 OrElse
            SortCombobox.SelectedIndex <= 0 Then Return

        Dim myfieldclass = CType(SortCombobox.SelectedItem, FieldClass)

        mylist.Remove(myfieldclass.id & " Asc")
        mylist.Remove(myfieldclass.id & " Desc")
        mylist.Add(myfieldclass.id & " " & IIf(RadioButton1.Checked, SortDirection.Asc.ToString, SortDirection.Desc.ToString))
        For i = 0 To mylist.Count - 1
            mysort = mysort + IIf(mysort = "", "", ",") + mylist.Item(i).ToString

        Next
        Sort = mysort
        bs.Sort = Sort

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        executeSort()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.Sort = ""
        Sort = ""
        mylist.Clear()
    End Sub
End Class


