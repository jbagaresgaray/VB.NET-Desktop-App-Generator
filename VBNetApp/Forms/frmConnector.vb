Imports MySql.Data.MySqlClient
Imports System.Data.Common

Public Class frmConnector
    Const DefaultConnection As String = "server=MyServer;port=3306;database=SampleDatabase;user id=sysadmin;password=adminpwd"
    Private _csb As New DbConnectionStringBuilder
    Private Sub UpdateCSB()
        Dim CnnStr As String = ""
        With _csb
            CnnStr = "Server = " & ServerTextBox.Text & ";"
            CnnStr += "Port = " & PortTextBox.Text & ";"
            CnnStr += "Database = " & DatabaseTextBox.Text & ";"
            CnnStr += "User ID = " & UserTextBox.Text & ";"
            CnnStr += "Password = " & PasswordTextBox.Text & ""
            .ConnectionString = CnnStr
        End With
    End Sub
    Private Sub UpdateControls()
        With _csb
            .TryGetValue("Server", ServerTextBox.Text)
            .TryGetValue("Port", PortTextBox.Text)
            .TryGetValue("Database", DatabaseTextBox.Text)
            .TryGetValue("User ID", UserTextBox.Text)
            .TryGetValue("Password", PasswordTextBox.Text)
        End With
    End Sub
    Public Function TestConnection() As Boolean

        ' +===============================================+
        ' ¦ 'For MySQL you may write something like this: ¦
        ' +===============================================+
        UpdateCSB()
        Dim cnn As New MySql.Data.MySqlClient.MySqlConnection(_csb.ConnectionString)
        Cursor = Cursors.WaitCursor
        Try
            cnn.Open()
            cnn.Close()
            MessageBox.Show("Test connection succeeded.", "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True
        Catch ex As Exception
            MessageBox.Show("Test connection failed because of an error in initializing provider." & vbCrLf & ex.Message, "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            Cursor = Cursors.Default
        End Try

        MessageBox.Show("Test connection succeeded.", "Fake Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Return True
    End Function

    Public ReadOnly Property ConnectionProperties() As System.Data.Common.DbConnectionStringBuilder
        Get
            UpdateCSB()
            Return _csb
        End Get
    End Property

    Public Property ConnectionString() As String
        Get
            UpdateCSB()
            Return _csb.ConnectionString
        End Get
        Set(ByVal value As String)
            If value = "" Then
                'load a sample ConnectionString
                _csb.ConnectionString = DefaultConnection
            Else
                _csb.ConnectionString = value
            End If
            UpdateControls()
        End Set
    End Property
    Public Sub ShowAdvancedProperties()
        UpdateCSB()
        Dim tempCSB As New DbConnectionStringBuilder
        tempCSB.ConnectionString = _csb.ConnectionString
        If ShowAdvancedPropertiesDialog(tempCSB) = DialogResult.OK Then
            _csb.ConnectionString = tempCSB.ConnectionString
            UpdateControls()
        End If
    End Sub

    Private Sub TestConnectionButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TestConnectionButton.Click
        TestConnection()
    End Sub

    Private Sub PropertiesButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PropertiesButton.Click
        ShowAdvancedProperties()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        If MsgBox("Application will restart after this configuration?" & vbNewLine & vbNewLine & "Continue restart?", vbYesNo + vbQuestion) = vbYes Then
            SaveINI("MySQL_Database", "MySQLServer", ServerTextBox.Text, CONFIG_INI_FILE)
            SaveINI("MySQL_Database", "UID", UserTextBox.Text, CONFIG_INI_FILE)
            SaveINI("MySQL_Database", "DefaultDB", DatabaseTextBox.Text, CONFIG_INI_FILE)
            SaveINI("MySQL_Database", "Password", PasswordTextBox.Text, CONFIG_INI_FILE)
            SaveINI("MySQL_Database", "Port", PortTextBox.Text, CONFIG_INI_FILE)

            Application.Restart()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub frmConnector_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DatabaseTextBox.Text = ReadINI("MySQL_Database", "DefaultDB", CONFIG_INI_FILE) 'DEFAULT_DATABASE
        ServerTextBox.Text = ReadINI("MySQL_Database", "MySQLServer", CONFIG_INI_FILE) 'MySQL_SERVER
        PortTextBox.Text = ReadINI("MySQL_Database", "Port", CONFIG_INI_FILE) 'DB_PORT
        PasswordTextBox.Text = ReadINI("MySQL_Database", "Password", CONFIG_INI_FILE) 'DB_PASSWORD
        UserTextBox.Text = ReadINI("MySQL_Database", "UID", CONFIG_INI_FILE) 'DB_USERID
    End Sub


End Class