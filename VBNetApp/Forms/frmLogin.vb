Imports MySql.Data.MySqlClient

Public Class frmLogin

    Private Sub frmLogin_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        txtUserName.Focus()
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If Not CheckTextBox(txtUserName, "Please enter Username") Then txtUserName.Focus() : Exit Sub
        If Not CheckTextBox(txtPassword, "Please enter Password") Then txtPassword.Focus() : Exit Sub

        If LoginUser(txtUserName.Text, txtPassword.Text) Then
            isLogin = True
            'Dim iCount As Integer
            'For iCount = 90 To 10 Step -10
            '    Me.Opacity = iCount / 100
            '    Me.Refresh()
            '    Threading.Thread.Sleep(50)
            'Next
            Close()
        Else
            MessageBox.Show("Username or Password is incorrect", "COMELEC Module", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Function LoginUser(ByVal UName As String, ByVal UPass As String) As Boolean
        On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim cmd As New MySqlCommand("SELECT * FROM employees WHERE e_username='" & UName & "' AND e_password=md5('" & UPass & "') LIMIT 1", con)
        Dim dr As MySqlDataReader = cmd.ExecuteReader()
        dr.Read()
        If dr.HasRows Then
            CURRENT_USER.e_fname = dr("e_fname").ToString()
            CURRENT_USER.e_id = dr("e_id").ToString()
            CURRENT_USER.e_lname = dr("e_lname").ToString()
            CURRENT_USER.e_mname = dr("e_mname").ToString()
            CURRENT_USER.e_username = dr("e_username").ToString()
            CURRENT_USER.e_password = dr("e_password").ToString()
            LoginUser = True
            isLogin = LoginUser
        Else
            LoginUser = False
            isLogin = LoginUser
        End If
        dr.Close()
        con.Close()
        Exit Function
err:
        DisplayErrorMsg(Me.Name, "LoginUser", Err.Number, Err.Description)
    End Function

    Private Sub txtUserName_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.KeyDown
        If e.KeyCode = Keys.Enter Then
            OK_Click(Me, New System.EventArgs())
        End If
    End Sub

    Private Sub txtPassword_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            OK_Click(Me, New System.EventArgs())
        End If
    End Sub
End Class
