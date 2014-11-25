'======================================================================================
'Filename:  modConnection.vb
'Type:      Module
'Author:    Philip Cesar B.Garay
'Date:      12.Oct.2012
'Email:     philipgaray2@gmail.com
'
'Purpose:   This module contains all the functionalities that prepares and binds connection to the database
'           it contains also the procedure that will execute the Main Application
'
'
'=======================================================================================


Imports MySql.Data.MySqlClient
Imports System.Data.Common



Module modConnection
    Public CONFIG_INI_FILE As String        'the configuration file
    Public REPORT_FILE As String            'the Report file

    Public DEFAULT_DATABASE As String       'Database
    Public MySQL_SERVER As String           'MySQL Server
    Public DB_USERID As String              'DB User
    Public DB_CONNECTION_STRING As String   'DB Connection String
    Public DB_PORT As String                'DB PORT
    Public DB_PASSWORD As String            'DB Password


    Public isLogin As Boolean


    '======================================================
    ' This procedure will Prepare the connection to the database
    '======================================================
    Public Function prepareDBConnString() As Integer
        Dim con As New MySqlConnection
        Dim isOpen As Boolean
        Dim ANS As MsgBoxResult
        isOpen = False
        '----------------------------
        ' Reset
        '----------------------------
        Dim strServer As String
        Dim strDatabase As String
        Dim strUserID As String
        Dim strPassword As String
        Dim strPort As String

        '----------------------------
        ' Init DB Config
        '----------------------------

        MySQL_SERVER = ReadINI("MySQL_Database", "MySQLServer", CONFIG_INI_FILE)

        DEFAULT_DATABASE = ReadINI("MySQL_Database", "DefaultDB", CONFIG_INI_FILE)

        DB_USERID = (ReadINI("MySQL_Database", "UID", CONFIG_INI_FILE))

        DB_PORT = (ReadINI("MySQL_Database", "Port", CONFIG_INI_FILE))

        DB_PASSWORD = (ReadINI("MySQL_Database", "Password", CONFIG_INI_FILE))

        'DB_CONNECTION_STRING = (ReadINI("MySQL_Database", "ConnectionString", CONFIG_INI_FILE))

        'SQLite_Database = ReadINI("SQLite", "Path", CONFIG_INI_FILE)                                                 'fetch all public-database-related variables

        '----------------------------
        ' Connection String
        '----------------------------

        'If Check1.Value = vbChecked Then                                        'Use Windows NT authentication using the network login ID.
        On Error GoTo prepareDBConnString_Error

        strServer = "Server=" & MySQL_SERVER & ";"                          'the .NET Mysql Connector
        strDatabase = "Database=" & DEFAULT_DATABASE & ";"                  'default Database
        strPort = "Port=" & DB_PORT & ";"                                   'Database Port
        strUserID = "UID=" & DB_USERID & ";"                                'database user
        strPassword = "Password=" & DB_PASSWORD & ";"                       'database password

        Do Until isOpen = True
            'DB_CONNECTION_STRING = "Data Source=" & SQLite_Database & ";Version=3;"

            DB_CONNECTION_STRING = strServer & strDatabase & strPort & strUserID & strPassword

            con.ConnectionString = DB_CONNECTION_STRING
            con.Open()
            isOpen = True
            con.Close()
        Loop
        prepareDBConnString = isOpen
        Exit Function

prepareDBConnString_Error:
        ANS = MsgBox("Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & " in procedure prepareDBConnString of Module mod_Main", _
                       MsgBoxStyle.RetryCancel)
        If ANS = vbCancel Then
            prepareDBConnString = MsgBoxResult.Cancel
            End
        ElseIf ANS = vbRetry Then
            prepareDBConnString = MsgBoxResult.Retry
        End If

    End Function

    ' ===========================================================
    '   Use to Display MySQL Connector Properties
    ' ===========================================================

    Public Function ShowAdvancedPropertiesDialog(ByRef CSB As DbConnectionStringBuilder) As DialogResult
        Dim frmx As New frmProperties()
        frmx.ConnectionProperties = CSB
        If frmx.ShowDialog() = DialogResult.OK Then
            CSB.ConnectionString = frmx.ConnectionProperties.ConnectionString
            Return DialogResult.OK
        Else
            Return DialogResult.Cancel
        End If
    End Function



    '======================================================
    ' This procedure will execute the Main Application
    '======================================================
    Public Sub Main()

        Application.EnableVisualStyles()                        ' This is already default on Visual Basic Application 
        Application.SetCompatibleTextRenderingDefault(False)    ' This is already default on Visual Basic Application 

        isLogin = False

        CONFIG_INI_FILE = Application.StartupPath & "\Config\config.ini"           'the main configuration .ini file
        REPORT_FILE = ReadINI("Paths", "Reports", CONFIG_INI_FILE)          'the main report file location

        MySQL_SERVER = ReadINI("MySQL_Database", "MySQLServer", CONFIG_INI_FILE)

        DEFAULT_DATABASE = ReadINI("MySQL_Database", "DefaultDB", CONFIG_INI_FILE)

        DB_USERID = (ReadINI("MySQL_Database", "UID", CONFIG_INI_FILE))

        DB_PORT = (ReadINI("MySQL_Database", "Port", CONFIG_INI_FILE))

        DB_PASSWORD = (ReadINI("MySQL_Database", "Password", CONFIG_INI_FILE))


        If (Trim(MySQL_SERVER) = "") And (Trim(DEFAULT_DATABASE) = "") Then
JumpHere:
            frmConnector.ShowDialog()
        End If

        If prepareDBConnString() = MsgBoxResult.Retry Then GoTo JumpHere

        Application.Run(New frmLogin())  ' Method use to execute run command on what form to display first on the Application

        If isLogin Then

            Application.Run(New frmMain())
        End If
    End Sub

End Module
