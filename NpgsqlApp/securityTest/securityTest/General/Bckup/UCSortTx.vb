Public Class UCSortTx
    Inherits UCToolstrip
    Dim dg As DataGridView
    Dim bs As New BindingSource
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
    Public Sub New(ByVal dg As DataGridView)

        ' This call is required by the designer.
        InitializeComponent()
        ToolStripLabel1.Text = "Sort"
        ' Add any initialization after the InitializeComponent() call.        
        bs = CType(DG.DataSource, BindingSource)
        Me.DG = DG
        InitDataLayout()
        MyBase.setPanelSize = Me.Panel1.Size
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub InitDataLayout()
        BindFieldCombobox()
    End Sub

    Private Sub BindFieldCombobox()
        UCHelper.BindFieldCombobox(FieldComboBox, dg)
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

        If bs.List.Count <= 0 OrElse FieldCombobox.Items.Count <= 0 OrElse
            FieldCombobox.SelectedIndex <= 0 Then Return

        Dim myfieldclass = CType(FieldCombobox.SelectedItem, FieldClass)

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
