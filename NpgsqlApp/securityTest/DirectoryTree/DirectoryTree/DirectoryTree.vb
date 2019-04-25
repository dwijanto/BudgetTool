Imports System.IO
Imports System.Windows.Forms
Public Class DirectoryTree
    Inherits TreeView

    Public Delegate Sub DirectorySelectedDelegate(ByVal sender As Object, ByVal e As DirectorySelectedEventArgs)
    Public Event DirectorySelected As DirectorySelectedDelegate

    Dim _Drive As Char

    Public Property Drive As Char
        Get
            Return _Drive
        End Get
        Set(ByVal value As Char)
            _Drive = value
            refreshdisplay()
        End Set
    End Property

    Private Sub refreshdisplay()
        Me.Nodes.Clear()
        Dim rootnode As TreeNode = New TreeNode(_Drive + ":\\")
        Me.Nodes.Add(rootnode)

        Fill(rootnode)
        Me.Nodes(0).Expand()
    End Sub

    Private Sub Fill(ByVal dirnode As TreeNode)
        Dim dir As DirectoryInfo = New DirectoryInfo(dirnode.FullPath)
        For Each diritem As DirectoryInfo In dir.GetDirectories
            Dim newnode As TreeNode = New TreeNode(diritem.Name)
            dirnode.Nodes.Add(newnode)
            newnode.Nodes.Add("*")
        Next
    End Sub

    Protected Overrides Sub OnBeforeExpand(ByVal e As System.Windows.Forms.TreeViewCancelEventArgs)
        MyBase.OnBeforeExpand(e)
        If (e.Node.Nodes(0).Text = "*") Then
            e.Node.Nodes.Clear()
            Fill(e.Node)
        End If
    End Sub
    Protected Overrides Sub OnAfterSelect(ByVal e As System.Windows.Forms.TreeViewEventArgs)
        MyBase.OnAfterSelect(e)
        RaiseEvent DirectorySelected(Me, New DirectorySelectedEventArgs(e.Node.FullPath))
    End Sub

End Class
Public Class DirectorySelectedEventArgs
    Inherits EventArgs
    Public Property DirectoryName As String

    Public Sub New(ByVal directoryname As String)
        Me.DirectoryName = directoryname
    End Sub
End Class