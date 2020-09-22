Imports DJLib.Dbtools
Imports System.Threading
Imports DJLib.AppConfig
Imports System.Reflection

Public Class FormMenu
    Private CancelFormClose As Boolean = True
    Private DynamicMenu1 As DJLib.DynamicMenu
    Private MenuStrip1 As MenuStrip
    Private bubbleperiod As Integer

    Private Sub FormMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.SuspendLayout()
        Me.Load_Menu()
        Me.ResumeLayout(False)
    End Sub

    Private Sub Load_Menu()
        Dim errMessage As String = vbNull
        Dim strbld As New System.Text.StringBuilder
        Me.Controls.Clear()
        For Each Str As String In DJLib.AppConfig.RoleAttribute.GetRolesForUser(DJLib.AppConfig.Identity.Name)
            If Str Is Nothing Then
                'Me.Close()
                Throw New Exception("User is not assign in any group")
            End If
            If strbld.Length <> 0 Then
                strbld.Append(" union ")
            End If
            If "localhost,172.22.13.139,SW58O951".Contains(ConnectionStringCollections.Item("HOST")) Then  'ConnectionStringCollections.Item("HOST") = "localhost" Or ConnectionStringCollections.Item("HOST") = "172.22.13.139" Then
                strbld.Append("(Select isactive,programid,parentid,myorder,description,programname,icon,iconindex,formname from tbprogram " & _
                          " where isactive and  members ~ '\m" & Str & "\y' order by parentid,myorder)")
            Else
                strbld.Append("(Select isactive,programid,parentid,myorder,description,programname,icon,iconindex,formname from tbprogram " & _
                          " where isactive and  members ~ '\\m" & Str & "\\y' order by parentid,myorder)")
            End If
            
            
        Next
        Dim DataTable1 As New DataTable
        If Not dbTools.getData(strbld.ToString, DataTable1, errMessage) Then
            MsgBox(errMessage)
        Else
            If DataTable1.Rows.Count > 0 Then
                DynamicMenu1 = New DJLib.DynamicMenu(Me, DataTable1, ImageList1)
                DynamicMenu1.LoadMenu(MenuStrip1)
                Me.MainMenuStrip = MenuStrip1
                Me.Controls.Add(MenuStrip1)                
            Else
                Me.Close()
                Throw New Exception("You don't have any access. Please contact admin! No menu found for this user")
            End If
        End If
        Me.Text = GetMenuDesc()
        Me.Location = New Point(300, 10)
    End Sub

    Private Function GetMenuDesc() As String
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & ConnectionStringCollections.Item("HOST") & ", Database: " & ConnectionStringCollections.Item("DATABASE") & ", Userid: " & DJLib.AppConfig.Identity.Name
    End Function

    Private Sub FormMenu_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        dbTools.Dispose()
        dbTools = Nothing
        If Not DynamicMenu1 Is Nothing Then
            DynamicMenu1.Dispose()
        End If
        DynamicMenu1 = Nothing
        If Not MenuStrip1 Is Nothing Then
            MenuStrip1.Dispose()
        End If
        MenuStrip1 = Nothing
        If Not StatusStrip1 Is Nothing Then
            StatusStrip1.Dispose()
        End If
        StatusStrip1 = Nothing

        ConnectionStringCollections = Nothing
        Me.Dispose()
    End Sub

    Private Sub MenuItemOnClick_mLoadMenu(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.Load_Menu()
    End Sub

    Private Sub MenuItemOnClick_mChangeUser(ByVal sender As Object, ByVal e As System.EventArgs)
        HR.HelperClass.ChangeUser = False
        FormLogon.ShowDialog()
        If HR.HelperClass.ChangeUser Then
            Me.CloseOpenForm()
            Me.Load_Menu()
        End If
        'close opened form
        
    End Sub



    Private Sub MenuItemOnClick_mMenuItemClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim inMemory As Boolean = False

        Dim frm As Form = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        For i = 0 To My.Application.OpenForms.Count - 1
            If My.Application.OpenForms.Item(i).Name = frm.Name Then
                ExecuteForm(My.Application.OpenForms.Item(i))
                inMemory = True
            End If
        Next
        If Not inMemory Then
            ExecuteForm(frm)
        End If
    End Sub

    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj

            .WindowState = FormWindowState.Normal
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .Focus()
        End With
    End Sub

    Private Sub MenuItemOnClick_mExit(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.ApplicationExitCall Then
            If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Me.CloseOpenForm()
                fadeout(Me)
            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub CloseOpenForm()
        For i = 1 To (My.Application.OpenForms.Count - 1)
            My.Application.OpenForms.Item(1).Close()
        Next
    End Sub
    Protected Friend Sub setBubbleMessage(ByVal title As String, ByVal message As String)
        NotifyIcon1.BalloonTipText = message
        NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
        NotifyIcon1.BalloonTipTitle = title
        NotifyIcon1.Visible = True
        NotifyIcon1.ShowBalloonTip(200)
        'ShowballonWindow(200)
    End Sub

    'Private Sub ShowballonWindow(ByVal timeout As Integer)
    '    Timer1.Interval = 1
    '    Timer1.Tag = timeout
    '    Timer1.Enabled = True
    '    'If timeout <= 0 Then
    '    '    Exit Sub
    '    'End If
    '    'Dim timeoutcount As Integer = 0
    '    'While (timeoutcount < timeout)
    '    '    Thread.Sleep(1)
    '    '    timeoutcount += 1
    '    'End While
    '    'NotifyIcon1.Visible = False
    'End Sub

    'Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '    If bubbleperiod > CInt(Timer1.Tag) Then
    '        bubbleperiod = 0
    '        Timer1.Enabled = False
    '        NotifyIcon1.Visible = False
    '    Else
    '        bubbleperiod += 1
    '    End If
    'End Sub
End Class