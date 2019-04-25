Imports DJLib
Public Class Form1
    Friend list1 As New List(Of TestClass)
    Public DataGridView1 As New DataGridView

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        SetupDataGridView()

        Dim obj As New TestClass
        'obj.Column1 = dr1("Column1")
        'obj.Column5 = dr1("Column5")
        'obj.Column2 = dr1("Column2")
        'obj.Column3 = dr1("Column3")
        'obj.Column4 = dr1("Column4")
        'obj.Column6 = dt
        list1.Add(obj)


        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Sub SetupDataGridView()

        Me.Controls.Add(DataGridView1)

        With DataGridView1
            .AutoGenerateColumns = False
            .Name = "DataGridView1"
            .Location = New Point(25, 25)
            .Size = New Size(642, 475)
            .RowHeadersVisible = False
            .BackgroundColor = SystemColors.Control
            .BorderStyle = BorderStyle.None
            .AllowUserToOrderColumns = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .StandardTab = True
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ReadOnly = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
        End With

        Dim Col1 As New DataGridViewTextBoxColumn()
        With Col1
            '.DataPropertyName = ""
            .Name = "Column1"
            .Visible = False
        End With

        Dim Col2 As New DataGridViewTextBoxColumn()
        With Col2
            '.DataPropertyName = ""
            .Name = "Column2"
            .Visible = False
        End With

        Dim Col3 As New DataGridViewTextBoxColumn()
        With Col3
            '.DataPropertyName = ""
            .Name = "Column3"
            .HeaderText = "Column3"
            .Width = 45
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        End With

        Dim Col4 As New DataGridViewTextBoxColumn()
        With Col4
            '.DataPropertyName = ""
            .Name = "Column4"
            .HeaderText = "Column4"
            .Width = 100
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        End With

        Dim Col5 As New DataGridViewTextBoxColumn()
        With Col5
            '.DataPropertyName = ""
            .Name = "Column5"
            .HeaderText = "Column5"
            .Width = 100
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        End With

        Dim Col6 As New NestedDgvColumn()
        With Col6
            '.DataPropertyName = ""
            .Name = "Column6"
            .HeaderText = "Column6"
            .Width = 377
        End With

        With DataGridView1
            .Columns.Insert(0, Col1)
            .Columns.Insert(1, Col2)
            .Columns.Insert(2, Col3)
            .Columns.Insert(3, Col4)
            .Columns.Insert(4, Col5)
            .Columns.Insert(5, Col6)
        End With

    End Sub


End Class


Friend Class TestClass

    Private m_column1 As Integer

    Public Property Column1() As Integer
        Get
            Return m_column1
        End Get
        Set(ByVal value As Integer)
            m_column1 = value
        End Set
    End Property

    Private m_column2 As String

    Public Property Column2() As String
        Get
            Return m_column2
        End Get
        Set(ByVal value As String)
            m_column2 = value
        End Set
    End Property

    Private m_column3 As String

    Public Property Column3() As String
        Get
            Return m_column3
        End Get
        Set(ByVal value As String)
            m_column3 = value
        End Set
    End Property

    Private m_column4 As Date

    Public Property Column4() As Date
        Get
            Return m_column4
        End Get
        Set(ByVal value As Date)
            m_column4 = value
        End Set
    End Property

    Private m_column6 As DataTable

    Public Property Column6() As DataTable
        Get
            Return m_column6
        End Get
        Set(ByVal value As DataTable)
            m_column6 = value
        End Set
    End Property

    Private m_column5 As Integer

    Public Property Column5() As Integer
        Get
            Return m_column5
        End Get
        Set(ByVal value As Integer)
            m_column5 = value
        End Set
    End Property

End Class

