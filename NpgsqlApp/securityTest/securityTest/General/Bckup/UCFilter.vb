Imports System.ComponentModel

Public Class UCFilter
    Inherits UserControl

    Private _hidetoolbar As HideToolbarDelegate
    Private BS As New BindingSource
    Private DG As DataGridView
    Private FilterOperatorHash As New Hashtable
    Private ComboList As New Dictionary(Of String, String)
    Private isDateTime As Boolean
    Dim _Collapsible As ExpandedState = ExpandedState.Expanded
    Dim dofade As Boolean = True
    Dim myLoc As Point
    Dim maxHeight As Integer

    Public Property Collapsible As ExpandedState
        Get
            Return _Collapsible
        End Get
        Set(ByVal value As ExpandedState)
            _Collapsible = value
        End Set
    End Property


    Dim _myfilter As String

    Dim _BS As New BindingSource

    Public ReadOnly Property myFilter As String
        Get
            Return _myfilter
        End Get
    End Property
    Public Property HideToolbar As HideToolbarDelegate
        Get
            Return _hidetoolbar
        End Get
        Set(ByVal value As HideToolbarDelegate)
            _hidetoolbar = value
        End Set
    End Property


    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        _hidetoolbar.Invoke(True)
    End Sub

    Public Sub New(ByVal HideToolbar As HideToolbarDelegate, ByVal DG As DataGridView)
        InitializeComponent()
        _hidetoolbar = HideToolbar
        BS = CType(DG.DataSource, BindingSource)
        Me.DG = DG
        InitDataLayout()
        myLoc = Panel1.Size
        maxHeight = myLoc.Y
    End Sub

    Private Sub InitDataLayout()
        BindFieldCombobox()
        BuildAutoCompleteString()
        OperatorComboBox.DataSource = System.Enum.GetNames(GetType(FilterOperator))
        InitFilterOperatorHash()
    End Sub

    Private Sub InitFilterOperatorHash()
        FilterOperatorHash.Add(0, "None")
        FilterOperatorHash.Add(1, "=")
        FilterOperatorHash.Add(2, "Like")
        FilterOperatorHash.Add(3, "<")
        FilterOperatorHash.Add(4, "<=")
        FilterOperatorHash.Add(5, ">")
        FilterOperatorHash.Add(6, ">=")
    End Sub

    Private Sub BindFieldCombobox()
        Dim cols As List(Of FieldClass) = New List(Of FieldClass)
        If DG.Columns.Count > 0 Then

            For i = 0 To DG.Columns.Count - 1
                If DG.Columns(i).Visible Then
                    cols.Add(New FieldClass With {.id = DG.Columns(i).DataPropertyName,
                                                  .name = DG.Columns(i).HeaderText,
                                                  .ColumnType = DG.Columns(i).GetType.Name,
                                                  .ColumnIndex = i})
                End If
            Next
            cols.Insert(0, New FieldClass With {.id = "None", .name = "None", .ColumnType = "None"})
        End If
        FieldComboBox.DataSource = cols
        FieldComboBox.DisplayMember = "Name"
        FieldComboBox.SelectedItem = "id"
    End Sub
    Private Sub BuildAutoCompleteString()
        Dim myfilter As String
        Dim myFieldClass = CType(FieldComboBox.SelectedItem, FieldClass)
        isDateTime = False
        'clear first
        FilterTextBox.AutoCompleteCustomSource.Clear()

        If BS.List.Count <= 0 OrElse FieldComboBox.Items.Count <= 0 OrElse
            FieldComboBox.SelectedIndex <= 0 Then Return

        'Get Column Name
        myfilter = BS.Filter
        If RadioButton2.Checked Then
            BS.Filter = ""
        End If

        Dim FilterField As String = myFieldClass.id 'CType(FieldComboBox.SelectedItem, FieldClass).id.ToString
        Dim filterVals As AutoCompleteStringCollection = New AutoCompleteStringCollection


        If myFieldClass.ColumnType = "DataGridViewComboBoxColumn" Then
            Dim bs2 As New BindingSource
            Dim dgcombo = CType(DG.Columns(myFieldClass.ColumnIndex), DataGridViewComboBoxColumn)

            bs2 = CType(dgcombo.DataSource, BindingSource)

            Try
                For Each dataitem As Object In bs2.List
                    Dim props As PropertyDescriptorCollection = TypeDescriptor.GetProperties(dataitem)
                    Dim propdesc As PropertyDescriptor = props.Find(dgcombo.DisplayMember, True)
                    Dim propdesc2 As PropertyDescriptor = props.Find(dgcombo.DataPropertyName, True)
                    Dim mykey As String = propdesc.GetValue(dataitem).ToString.ToLower
                    Dim myvalue As String = propdesc2.GetValue(dataitem).ToString
                    Try
                        ComboList.Add(mykey, myvalue)
                    Catch ex As Exception

                    End Try
                    filterVals.Add(mykey)
                Next
            Catch ex As Exception

            End Try
        Else
            For Each dataitem As Object In BS.List
                Dim props As PropertyDescriptorCollection = TypeDescriptor.GetProperties(dataitem)
                Dim propdesc As PropertyDescriptor = props.Find(FilterField, True)
                Try
                    Dim fieldval As String = propdesc.GetValue(dataitem).ToString
                    If propdesc.PropertyType.Name = "DateTime" Then
                        isDateTime = True
                    End If
                    filterVals.Add(fieldval)
                Catch ex As Exception

                End Try

            Next
        End If

        BS.Filter = myfilter
        FilterTextBox.AutoCompleteCustomSource = filterVals
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        FilterOperatorHash.Clear()
        ComboList.Clear()
    End Sub
    Private Sub FieldComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FieldComboBox.SelectedIndexChanged, RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        BuildAutoCompleteString()
    End Sub

    'Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    BuildAutoCompleteString()
    'End Sub


    Private Sub executefilter()
        If BS.List.Count <= 0 OrElse
            FieldComboBox.Items.Count <= 0 OrElse
            FieldComboBox.SelectedIndex <= 0 OrElse
            OperatorComboBox.SelectedIndex <= 0 Then
            Return
        End If

        If String.IsNullOrEmpty(FilterTextBox.Text) Then Return

        'inFilterMode = True

        '##getpropertyname##
        '1.get columnname from combo
        Dim myFieldClass = CType(FieldComboBox.SelectedItem, FieldClass)
        Dim filterMember As String = myFieldClass.id.ToString 'CType(FieldComboBox.SelectedItem, FieldClass).id.ToString

        '1.b Check for ComboboxColumn
        Dim filterValue As String = Nothing
        Dim SearchValue As String = Nothing
        SearchValue = FilterTextBox.Text
        If myFieldClass.ColumnType = "DataGridViewComboBoxColumn" Then
            Try
                SearchValue = ComboList(SearchValue.ToLower)
            Catch ex As Exception
            End Try
        End If
        '2.Get dataitem from bindinglist.list(0)
        Dim DataItem As Object = BS.List(0)
        '3.Get Propertiescollection from dataitem
        Dim props As PropertyDescriptorCollection = TypeDescriptor.GetProperties(DataItem)
        '4.Get Selected PropertyDescriptor based on filtermember
        Dim propDesc As PropertyDescriptor = props.Find(filterMember, True)

        'getoperator
        Dim stringoperator As String = FilterOperatorHash(OperatorComboBox.SelectedIndex).ToString
        'putbindingfilter
        'Check for different format

        Dim JoinFilter As String = "AND "
        If Not CheckBox1.Checked Then
            _myfilter = ""

        End If
        'If BS.Filter <> "" AndAlso BS.Filter IsNot Nothing Then
        If _myfilter <> "" AndAlso _myfilter IsNot Nothing Then
            If RadioButton2.Checked Then JoinFilter = "OR "
        Else
            JoinFilter = ""
        End If

        Select Case OperatorComboBox.SelectedIndex
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
                If isDateTime Then
                    filterValue = String.Format("{0}{1} {2} '#{3}#'", JoinFilter, propDesc.Name, stringoperator, SearchValue)
                Else
                    filterValue = String.Format("{0}{1} {2} '{3}'", JoinFilter, propDesc.Name, stringoperator, SearchValue)
                End If

        End Select

        Try
            _myfilter = _myfilter & filterValue
            BS.Filter = myFilter
        Catch ex As Exception
            _myfilter = ""
            BS.Filter = ""
        End Try

    End Sub


    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        RadioButton1.Enabled = CheckBox1.Checked
        RadioButton2.Enabled = CheckBox1.Checked
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If dofade Then
            If myLoc.Y > 0 Then
                Panel1.Size = myLoc
                myLoc.Offset(0, -10)
            Else
                Timer1.Stop()
                dofade = False
                Panel1.Visible = False
                ToolStripButton2.Image = My.Resources.go_bottom
            End If
        Else
            Timer1.Stop()
            ToolStripButton2.Image = My.Resources.go_top
            Panel1.Height = maxHeight
            Panel1.Visible = True
            myLoc = Panel1.Size
            dofade = True
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        executefilter()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        BS.Filter = ""
        _myfilter = ""
    End Sub
End Class
