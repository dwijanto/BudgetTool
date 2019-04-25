Public Class Form2

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        loaddata()

    End Sub

    Private Sub loaddata()
        ' Create a new ListView control.
        ' Dim listView1 As New ListView()
        Me.ListView1.View = View.Details

        ' Add columns using the ColHeader class. The fourth    
        ' parameter specifies true for an ascending sort order.
        ListView1.Groups.Add(New ListViewGroup("hello 1", HorizontalAlignment.Left))
        'ListView1.Groups.Add(New ListViewGroup("hello 2", HorizontalAlignment.Right))
        listView1.Columns.Add(New ColHeader("Name", 110, HorizontalAlignment.Left, True))
        listView1.Columns.Add(New ColHeader("Region", 50, HorizontalAlignment.Left, True))
        listView1.Columns.Add(New ColHeader("Sales", 70, HorizontalAlignment.Left, True))


        ' Add the data.
        listView1.Items.Add(New ListViewItem(New String() {"Archer, Karen", "4", "0521.28"}))
        listView1.Items.Add(New ListViewItem(New String() {"Benson, Max", "8", "0828.54"}))
        listView1.Items.Add(New ListViewItem(New String() {"Bezio, Marin", "3", "0535.22"}))
        listView1.Items.Add(New ListViewItem(New String() {"Higa, Sidney", "2", "0987.50"}))
        listView1.Items.Add(New ListViewItem(New String() {"Martin, Linda", "6", "1122.12"}))
        listView1.Items.Add(New ListViewItem(New String() {"Nash, Mike", "7", "1030.11"}))
        listView1.Items.Add(New ListViewItem(New String() {"Sanchez, Ken", "1", "0958.78"}))
        listView1.Items.Add(New ListViewItem(New String() {"Smith, Ben", "5", "0763.25"}))
        For i = 0 To ListView1.Items.Count - 1
            ListView1.Items.Item(i).Group = ListView1.Groups(0)
        Next


        ' Connect the ListView.ColumnClick event to the ColumnClick event handler.
        AddHandler listView1.ColumnClick, AddressOf listView1_ColumnClick



    End Sub
    Private Sub listView1_ColumnClick(ByVal sender As Object, ByVal e As ColumnClickEventArgs)

        ' Create an instance of the ColHeader class. 
        Dim clickedCol As ColHeader = CType(Me.listView1.Columns(e.Column), ColHeader)

        ' Set the ascending property to sort in the opposite order.
        clickedCol.ascending = Not clickedCol.ascending

        ' Get the number of items in the list.
        Dim numItems As Integer = Me.listView1.Items.Count

        ' Turn off display while data is repoplulated.
        Me.listView1.BeginUpdate()

        ' Populate an ArrayList with a SortWrapper of each list item.
        Dim SortArray As New ArrayList
        Dim i As Integer
        For i = 0 To numItems - 1
            SortArray.Add(New SortWrapper(Me.listView1.Items(i), e.Column))
        Next i

        ' Sort the elements in the ArrayList using a new instance of the SortComparer
        ' class. The parameters are the starting index, the length of the range to sort,
        ' and the IComparer implementation to use for comparing elements. Note that
        ' the IComparer implementation (SortComparer) requires the sort  
        ' direction for its constructor; true if ascending, othwise false.
        SortArray.Sort(0, SortArray.Count, New SortWrapper.SortComparer(clickedCol.ascending))

        ' Clear the list, and repopulate with the sorted items.
        Me.listView1.Items.Clear()
        Dim z As Integer
        For z = 0 To numItems - 1
            Me.listView1.Items.Add(CType(SortArray(z), SortWrapper).sortItem)
        Next z
        ' Turn display back on.
        Me.listView1.EndUpdate()
    End Sub

End Class
' An instance of the SortWrapper class is created for 
' each item and added to the ArrayList for sorting.
Public Class SortWrapper
    Friend sortItem As ListViewItem
    Friend sortColumn As Integer

    ' A SortWrapper requires the item and the index of the clicked column.
    Public Sub New(ByVal Item As ListViewItem, ByVal iColumn As Integer)
        sortItem = Item
        sortColumn = iColumn
    End Sub

    ' Text property for getting the text of an item.
    Public ReadOnly Property [Text]() As String
        Get
            Return sortItem.SubItems(sortColumn).Text
        End Get
    End Property

    ' Implementation of the IComparer 
    ' interface for sorting ArrayList items.
    Public Class SortComparer
        Implements IComparer
        Private ascending As Boolean


        ' Constructor requires the sort order;
        ' true if ascending, otherwise descending.
        Public Sub New(ByVal asc As Boolean)
            Me.ascending = asc
        End Sub


        ' Implemnentation of the IComparer:Compare 
        ' method for comparing two objects.
        Public Function [Compare](ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Dim xItem As SortWrapper = CType(x, SortWrapper)
            Dim yItem As SortWrapper = CType(y, SortWrapper)

            Dim xText As String = xItem.sortItem.SubItems(xItem.sortColumn).Text
            Dim yText As String = yItem.sortItem.SubItems(yItem.sortColumn).Text
            Return xText.CompareTo(yText) * IIf(Me.ascending, 1, -1)
        End Function
    End Class
End Class
' The ColHeader class is a ColumnHeader object with an 
' added property for determining an ascending or descending sort.
' True specifies an ascending order, false specifies a descending order.
Public Class ColHeader
    Inherits ColumnHeader
    Public ascending As Boolean

    Public Sub New(ByVal [text] As String, ByVal width As Integer, ByVal align As HorizontalAlignment, ByVal asc As Boolean)
        Me.Text = [text]
        Me.Width = width
        Me.TextAlign = align
        Me.ascending = asc
    End Sub
End Class
