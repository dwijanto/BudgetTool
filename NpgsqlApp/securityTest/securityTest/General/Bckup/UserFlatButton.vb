
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class UserFlatButton
    Inherits System.Windows.Forms.Form

    Private button1 As System.Windows.Forms.Button = New Button
    Private button2 As System.Windows.Forms.Button = New Button

    <System.STAThreadAttribute()> _
    Public Shared Sub Main()
        System.Windows.Forms.Application.Run(New UserFlatButton)
    End Sub

    Public Sub New()
        InitializeComponent()
        Me.button2.Location = New Point(0, button1.Height + 10)
        AddHandler Me.button2.Click, AddressOf Me.button2_Click
        Me.Controls.Add(Me.button1)
        Me.Controls.Add(Me.button2)
    End Sub

    Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ' Draws a flat button on button1.


    End Sub 'button2_Click
End Class
