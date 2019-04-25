Imports System.Windows.Forms

Public Class DialogChangePassword

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If DJLib.AppConfig.MembershipService.ChangePassword(DJLib.AppConfig.Identity.Name, TextBox1.Text, TextBox2.Text) Then
            MessageBox.Show("Change Password successful")
        Else
            MessageBox.Show("Change Password failed")
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
