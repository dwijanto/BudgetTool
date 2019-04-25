﻿Public Class FormDataGridViewBand
    Friend WithEvents dataGridView1 As DataGridView
    Friend WithEvents Button1 As Button = New Button()
    Friend WithEvents Button2 As Button = New Button()
    Friend WithEvents Button3 As Button = New Button()
    Friend WithEvents Button4 As Button = New Button()
    Friend WithEvents Button5 As Button = New Button()
    Friend WithEvents Button6 As Button = New Button()
    Friend WithEvents Button7 As Button = New Button()
    Friend WithEvents Button8 As Button = New Button()
    Friend WithEvents Button9 As Button = New Button()
    Friend WithEvents Button10 As Button = New Button()
    'Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel = New FlowLayoutPanel

    Public Sub New()
        MyBase.New()

        InitializeComponent()

        Dim e As System.EventArgs = New System.EventArgs
        InitializeDataGridView(Me, e)
        AddButton(Button1, "Reset")
        AddButton(Button2, "Change Column 3 Header")
        AddButton(Button3, "Change Meatloaf Recipe")
        AddAdditionalButtons()
    End Sub
#Region "setup DataGridView"
    Private thirdColumnHeader As String = "Main Ingredients"
    Private boringMeatloaf As String = "ground beef"
    Private boringMeatloafRanking As String = "*"
    Private boringRecipe As Boolean
    Private shortMode As Boolean

    Private Sub InitializeDataGridView(ByVal ignored As Object, _
    ByVal ignoredToo As EventArgs)
        dataGridView1 = New System.Windows.Forms.DataGridView

        Controls.Add(DataGridView1)
        DataGridView1.Size = New Size(300, 200)

        ' Create an unbound DataGridView by declaring a
        ' column count.
        DataGridView1.ColumnCount = 4
        AdjustDataGridViewSizing()

        ' Set the column header style.
        Dim columnHeaderStyle As New DataGridViewCellStyle
        columnHeaderStyle.BackColor = Color.Aqua
        columnHeaderStyle.Font = _
            New Font("Verdana", 10, FontStyle.Bold)
        DataGridView1.ColumnHeadersDefaultCellStyle = _
            columnHeaderStyle

        ' Set the column header names.
        DataGridView1.Columns(0).Name = "Recipe"
        DataGridView1.Columns(1).Name = "Category"
        DataGridView1.Columns(2).Name = thirdColumnHeader
        DataGridView1.Columns(3).Name = "Rating"

        ' Populate the rows.
        Dim row1 As String() = New String() _
            {"Meatloaf", "Main Dish", boringMeatloaf, _
            boringMeatloafRanking}
        Dim row2 As String() = New String() _
            {"Key Lime Pie", "Dessert", _
            "lime juice, evaporated milk", _
            "****"}
        Dim row3 As String() = New String() _
            {"Orange-Salsa Pork Chops", "Main Dish", _
            "pork chops, salsa, orange juice", "****"}
        Dim row4 As String() = New String() _
            {"Black Bean and Rice Salad", "Salad", _
            "black beans, brown rice", _
            "****"}
        Dim row5 As String() = New String() _
            {"Chocolate Cheesecake", "Dessert", "cream cheese", _
            "***"}
        Dim row6 As String() = New String() _
            {"Black Bean Dip", "Appetizer", _
            "black beans, sour cream", _
                "***"}
        Dim rows As Object() = New Object() {row1, row2, _
            row3, row4, row5, row6}

        Dim rowArray As String()
        For Each rowArray In rows
            DataGridView1.Rows.Add(rowArray)
        Next

        PostRowCreation()

        shortMode = False
        boringRecipe = True
    End Sub

    Protected Sub AddButton(ByVal button As Button, _
        ByVal buttonLabel As String)

        FlowLayoutPanel1.Controls.Add(button)
        button.TabIndex = FlowLayoutPanel1.Controls.Count
        button.Text = buttonLabel
        button.AutoSize = True
    End Sub

    ' Reset columns to initial disorderly arrangement.
    Private Sub ResetToDisorder(ByVal sender As Object, _
        ByVal e As System.EventArgs) _
        Handles Button1.Click
        Controls.Remove(dataGridView1)
        dataGridView1.Dispose()
        'DataGridView1.DataSource = Nothing
        InitializeDataGridView(Nothing, Nothing)
    End Sub

    ' Change the header in column three.
    Private Sub Button2_Click(ByVal sender As Object, _
        ByVal e As System.EventArgs) _
        Handles Button2.Click

        Toggle(shortMode)
        If shortMode Then DataGridView1.Columns(2).HeaderText = _
            "S" _
            Else DataGridView1.Columns(2).HeaderText = _
                thirdColumnHeader
    End Sub

    Private Shared Sub Toggle(ByRef toggleThis As Boolean)
        toggleThis = Not toggleThis
    End Sub

    ' Change the meatloaf recipe.
    Private Sub Button3_Click(ByVal sender As Object, _
        ByVal e As System.EventArgs) _
        Handles Button3.Click

        Toggle(boringRecipe)
        If boringRecipe Then
            SetMeatloaf(boringMeatloaf, boringMeatloafRanking)
        Else
            Dim greatMeatloafRecipe As String = "1 lb. lean ground beef, " _
                & "1/2 cup bread crumbs, 1/4 cup ketchup," _
                & "1/3 tsp onion powder, " _
                & "1 clove of garlic, 1/2 pack onion soup mix " _
                & " dash of your favorite BBQ Sauce"
            SetMeatloaf(greatMeatloafRecipe, "***")
        End If
    End Sub

    Private Sub SetMeatloaf(ByVal recipe As String, _
        ByVal rating As String)

        DataGridView1.Rows(0).Cells(2).Value = recipe
        DataGridView1.Rows(0).Cells(3).Value = rating
    End Sub
#End Region

#Region "demonstration code"
    Private Sub AddAdditionalButtons()
        AddButton(Button4, "Freeze First Row")
        AddButton(Button5, "Freeze Second Column")
        AddButton(Button6, "Hide Salad Row")
        AddButton(Button7, "Disable First Column Resizing")
        AddButton(Button8, "Make ReadOnly")
        AddButton(Button9, "Style Using Tag")
    End Sub

    Private Sub AdjustDataGridViewSizing()
        DataGridView1.AutoSizeRowsMode = _
            DataGridViewAutoSizeRowsMode.AllCells
        DataGridView1.ColumnHeadersHeightSizeMode = _
            DataGridViewColumnHeadersHeightSizeMode.AutoSize
    End Sub

    ' Freeze the first row.
    Private Sub Button4_Click(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles Button4.Click

        FreezeBand(DataGridView1.Rows(0))
    End Sub

    Private Sub FreezeColumn(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles Button5.Click

        FreezeBand(DataGridView1.Columns(1))
    End Sub

    Private Shared Sub FreezeBand(ByVal band As DataGridViewBand)

        band.Frozen = True
        Dim style As DataGridViewCellStyle = New DataGridViewCellStyle()
        style.BackColor = Color.WhiteSmoke
        band.DefaultCellStyle = style

    End Sub

    ' Hide a band of cells.
    Private Sub Button6_Click(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles Button6.Click

        Dim band As DataGridViewBand = DataGridView1.Rows(3)
        band.Visible = False
    End Sub

    ' Turn off user's ability to resize a column.
    Private Sub Button7_Click(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles Button7.Click

        Dim band As DataGridViewBand = DataGridView1.Columns(0)
        band.Resizable = DataGridViewTriState.False
    End Sub

    ' Make the the entire DataGridView read only.
    Private Sub Button8_Click(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles Button8.Click

        For Each band As DataGridViewBand In DataGridView1.Columns
            band.ReadOnly = True
        Next
    End Sub

    Private Sub PostRowCreation()
        SetBandColor(DataGridView1.Columns(0), Color.CadetBlue)
        SetBandColor(DataGridView1.Rows(1), Color.Coral)
        SetBandColor(DataGridView1.Columns(2), Color.DodgerBlue)
    End Sub

    Private Shared Sub SetBandColor(ByVal band As DataGridViewBand, _
        ByVal color As Color)
        band.Tag = color
    End Sub

    ' Color the bands by the value stored in their tag.
    Private Sub Button9_Click(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles Button9.Click

        For Each band As DataGridViewBand In DataGridView1.Columns
            If band.Tag IsNot Nothing Then
                band.DefaultCellStyle.BackColor = _
                    CType(band.Tag, Color)
            End If
        Next

        'For Each band As DataGridViewBand In DataGridView1.Rows
        '    If band.Tag IsNot Nothing Then
        '        band.DefaultCellStyle.BackColor = _
        '            CType(band.Tag, Color)
        '    End If
        'Next
    End Sub




#End Region

    <STAThreadAttribute()> _
    Public Shared Sub Main()
        Application.Run(New ClassDataGridViewBand())
    End Sub

    Private Sub dataGridView_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs)
        If e.ColumnIndex = dataGridView1.Columns(3).Index AndAlso (e.Value IsNot Nothing) Then
            With Me.dataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)

                If e.Value.Equals("*") Then
                    .ToolTipText = "very bad"
                ElseIf e.Value.Equals("**") Then
                    .ToolTipText = "bad"
                ElseIf e.Value.Equals("***") Then
                    .ToolTipText = "good"
                ElseIf e.Value.Equals("****") Then
                    .ToolTipText = "very good"
                End If

            End With


        End If
    End Sub
End Class