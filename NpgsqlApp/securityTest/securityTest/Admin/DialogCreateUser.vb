Imports System.Windows.Forms
Imports DJLib.AppConfig
Imports DJLib.AccountValidation
Imports DJLib.Dbtools
Imports HR.HelperClass

Public Class DialogCreateUser

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MessageBox.Show("All Fields are mandatory!")
            Exit Sub
        End If
        Try
            Dim status = MembershipService.CreateUser(TextBox1.Text, TextBox2.Text, TextBox3.Text)
            If status = System.Web.Security.MembershipCreateStatus.Success Then
                'update region
                DbAdapter1.UpdateUserRegion(TextBox3.Text, ComboBox1.SelectedValue)
                'Create user database
                DbAdapter1.CreateUserDb(TextBox1.Text, TextBox2.Text)
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show(ErrorCodeToString(status))
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub DialogCreateUser_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        populatecombobox()
    End Sub

    Private Sub populatecombobox()
        dbTools.FillComboboxDataSource(ComboBox1, "select regionid,regionname from region order by regionname")
    End Sub

End Class
