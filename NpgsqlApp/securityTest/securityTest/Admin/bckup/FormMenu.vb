Imports DJLib.Dbtools
Imports System.Threading
Imports DJLib.AppConfig

Public Class FormMenu
    Private CancelFormClose As Boolean = True
    Private DynamicMenu1 As DJLib.DynamicMenu
    Private MenuStrip1 As MenuStrip

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Me.Text = "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & ConnectionStringCollections.Item("HOST") & ", Database: " & ConnectionStringCollections.Item("DATABASE") & ", Userid: " & DJLib.AppConfig.Identity.Name 'dbTools.Userid
        Me.Location = New Point(300, 10)
    End Sub

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



    Private Sub FormMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SuspendLayout()
        Load_Menu()
        ResumeLayout(False)
    End Sub

    Private Sub Load_Menu()
        Dim errMessage As String = vbNull
        Dim strbld As New System.Text.StringBuilder
        Me.Controls.Clear()
        For Each Str As String In DJLib.AppConfig.RoleAttribute.GetRolesForUser(DJLib.AppConfig.Identity.Name)
            If Str Is Nothing Then
                Me.Close()
                Throw New Exception("User is not assign in any group")
            End If
            If strbld.Length <> 0 Then
                strbld.Append(" union ")
            End If
            strbld.Append("(Select isactive,programid,parentid,myorder,description,programname,icon,iconindex from tbprogram " & _
                          " where isactive and  members ~ '\\m" & Str & "\\y' order by parentid,myorder)")
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
                'Me.Text += " Number user Online:" & DJLib.AppConfig.MembershipService.GetNumberOFUsersOnline()
            Else
                Me.Close()
                Throw New Exception("You don't have any access. Please contact admin! No menu found for this user")
            End If
        End If
        Me.Text = "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & ConnectionStringCollections.Item("HOST") & ", Database: " & ConnectionStringCollections.Item("DATABASE") & ", Userid: " & DJLib.AppConfig.Identity.Name 'dbTools.Userid
    End Sub
    Private Sub MenuItemOnClick_mUsers(ByVal sender As Object, ByVal e As System.EventArgs)
        FormUser.Show()
    End Sub
    Private Sub MenuItemOnClick_mConcurrency(ByVal sender As Object, ByVal e As System.EventArgs)
        ConcurrencyHandling.Show()
    End Sub

    Private Sub MenuItemOnClick_mProgram(ByVal sender As Object, ByVal e As System.EventArgs)
        FormProgramCB.Show()
    End Sub
    Private Sub MenuItemOnClick_mRoles(ByVal sender As Object, ByVal e As System.EventArgs)
        FormRoles.Show()
    End Sub
    Private Sub MenuItemOnClick_mChangePassword(ByVal sender As Object, ByVal e As System.EventArgs)
        DialogChangePassword.Show()
    End Sub

    Private Sub MenuItemOnClick_mUserRoles(ByVal sender As Object, ByVal e As System.EventArgs)
        FormUserRoles.Show()
    End Sub

    Private Sub MenuItemOnClick_mAnimal(ByVal sender As Object, ByVal e As System.EventArgs)
        'refcursor
        FormAnimal.Show()
    End Sub

    Private Sub MenuItemOnClick_mAnimal2(ByVal sender As Object, ByVal e As System.EventArgs)
        FormAnimal2.Show()
    End Sub
    Private Sub MenuItemOnClick_mPet(ByVal sender As Object, ByVal e As System.EventArgs)
        FormGetChanges.Show()
    End Sub
    Private Sub MenuItemOnClick_mComboBox(ByVal sender As Object, ByVal e As System.EventArgs)
        FormCombobox.Show()
    End Sub
    Private Sub MenuItemOnClick_mComboBoxManual(ByVal sender As Object, ByVal e As System.EventArgs)
        FormComboBoxManual.Show()
    End Sub

    Private Sub MenuItemOnClick_mComboBoxManualCopy(ByVal sender As Object, ByVal e As System.EventArgs)
        FormComboBoxManualCopy.Show()
    End Sub
    Private Sub MenuItemOnClick_mMenuEditor(ByVal sender As Object, ByVal e As System.EventArgs)
        FormMenuEditor.Show()
    End Sub

    Private Sub MenuItemOnClick_mForm2(ByVal sender As Object, ByVal e As System.EventArgs)
        Form2.Show()
    End Sub

    Private Sub MenuItemOnClick_mMenuMember(ByVal sender As Object, ByVal e As System.EventArgs)
        FormMenuMember.Show()
    End Sub
    Private Sub MenuItemOnClick_mLoadMenu(ByVal sender As Object, ByVal e As System.EventArgs)
        Load_Menu()
    End Sub
    Private Sub MenuItemOnClick_mChangeUser(ByVal sender As Object, ByVal e As System.EventArgs)
        FormLogon.ShowDialog()
        Load_Menu()
    End Sub
    Private Sub MenuItemOnClick_mCustomCombo(ByVal sender As Object, ByVal e As System.EventArgs)
        ClassComboCustom.Show()
    End Sub
    Private Sub MenuItemOnClick_mUserFlatButton(ByVal sender As Object, ByVal e As System.EventArgs)
        UserFlatButton.Show()
    End Sub
    Private Sub MenuItemOnClick_mUnbounDataGridView(ByVal sender As Object, ByVal e As System.EventArgs)
        FormUnboundDataGridView.Show()
    End Sub
    Private Sub MenuItemOnClick_mTicTacToe(ByVal sender As Object, ByVal e As System.EventArgs)
        TicTacToe.Show()
    End Sub
    Private Sub MenuItemOnClick_mDataGridViewButton(ByVal sender As Object, ByVal e As System.EventArgs)
        'ClassDataGridViewButton.Show()
    End Sub

    Private Sub MenuItemOnClick_mExpetTx(ByVal sender As Object, ByVal e As System.EventArgs)
        ' FormExPetT.Show()
    End Sub

    Private Sub MenuItemOnClick_mDataGridViewBand(ByVal sender As Object, ByVal e As System.EventArgs)
        FormDataGridViewBand.Show()
    End Sub

    Private Sub MenuItemOnClick_mInherit2(ByVal sender As Object, ByVal e As System.EventArgs)
        FormInherit2.Show()
    End Sub
    Private Sub MenuItemOnClick_mInherit3(ByVal sender As Object, ByVal e As System.EventArgs)
        FormInherit3.Show()
        FormInherit3.Focus()
    End Sub
    Private Sub MenuItemOnClick_mAnimalInherit(ByVal sender As Object, ByVal e As System.EventArgs)
        FormAnimalInherit.Show()
    End Sub
    Private Sub MenuItemOnClick_mForm1(ByVal sender As Object, ByVal e As System.EventArgs)
        Form1.Show()

    End Sub
    Private Sub MenuItemOnClick_mFormTest(ByVal sender As Object, ByVal e As System.EventArgs)
        FormToolstripContainerTest.Show()
    End Sub

    Private Sub MenuItemOnClick_mThreadOne(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim f As Threading1 = New Threading1    
        f.Show()
    End Sub

    Private Sub MenuItemOnClick_mThreadTwo(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim f As Threading2 = New Threading2
        'f.Show()
        ExecuteForm(New Threading2)
    End Sub
    Private Sub MenuItemOnClick_mGDI(ByVal sender As Object, ByVal e As System.EventArgs)
        ExecuteForm(FormGDI)
    End Sub
    Private Sub MenuItemOnClick_mGDI2(ByVal sender As Object, ByVal e As System.EventArgs)
        ExecuteForm(FormGDI2)
    End Sub
    Private Sub MenuItemOnClick_mGDI3(ByVal sender As Object, ByVal e As System.EventArgs)
        ExecuteForm(FormGDI3)
    End Sub
    Private Sub MenuItemOnClick_mGDI4(ByVal sender As Object, ByVal e As System.EventArgs)
        ExecuteForm(ClassGDI4)
    End Sub
    Private Sub MenuItemOnClick_mGDI5(ByVal sender As Object, ByVal e As System.EventArgs)
        ExecuteForm(FormVisualStyleRenderer)
    End Sub
    Private Sub MenuItemOnClick_mGDI6(ByVal sender As Object, ByVal e As System.EventArgs)
        ExecuteForm(ClassGDI6)
    End Sub
    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj
            .Show()
            .WindowState = FormWindowState.Normal
            .Focus()
        End With
    End Sub
    Private Sub MenuItemOnClick_mVirtualMode(ByVal sender As Object, ByVal e As System.EventArgs)
        'using Default instance
        With FormVirtualMode
            .Show()
            .BringToFront()
            .WindowState = FormWindowState.Normal
        End With
        FormVirtualMode.Text = "hello"

        'Using Instance
        'Preparation: in form closing in formvirtualmode, set _instance = nothing
        'otherwise you'll get error of disposed object
        'Dim myform1 As FormVirtualMode = FormVirtualMode.Instance
        'With myform1
        '    .Show()
        '    .WindowState = FormWindowState.Normal
        '    .Focus()
        'End With
    End Sub

    Private Sub MenuItemOnClick_mExit(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
            For i = 1 To (My.Application.OpenForms.Count - 1)
                My.Application.OpenForms.Item(1).Close()
            Next
            fadeout(Me)                        
        Else
            e.Cancel = True
        End If
    End Sub

    Protected Friend Sub setBubbleMessage(ByVal title As String, ByVal message As String)
        NotifyIcon1.Visible = True
        NotifyIcon1.BalloonTipText = message
        NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
        NotifyIcon1.BalloonTipTitle = title
        NotifyIcon1.ShowBalloonTip(200)
        ShowballonWindow(200)
    End Sub

    Private Sub ShowballonWindow(ByVal timeout As Integer)
        If timeout <= 0 Then
            Exit Sub
        End If
        Dim timeoutcount As Integer = 0
        While (timeoutcount < timeout)
            Thread.Sleep(1)
            timeoutcount += 1
        End While
        NotifyIcon1.Visible = False
    End Sub


End Class