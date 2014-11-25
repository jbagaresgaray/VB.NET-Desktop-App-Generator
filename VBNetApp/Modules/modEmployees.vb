
Option Explicit On

Imports MySql.Data.MySqlClient

Module modEmployees

    Public Structure employees
        Dim e_id As Integer
        Dim e_fname As String
        Dim e_lname As String
        Dim e_mname As String
        Dim e_username As String
        Dim e_password As String
        Dim mug_id As Integer
        Dim mug_name As String
        Dim is_deleted As Short
    End Structure

    Public Function GetUserByCredentials(ByVal UserName As String, _
                                            ByVal Password As String, Optional ByVal Show_Deleted As Boolean = True) As Boolean

        Dim sSQL As String = "SELECT e.*, mug.mug_id, mug.mug_name FROM employees e " & _
            "LEFT JOIN module_user_groups mug ON e.mug_id = mug.mug_id "

        If Len(UserName) > 0 And Len(Password) > 0 Then
            sSQL = sSQL & "WHERE e.e_username='" & UserName & "' AND e.e_password='" & Password & "'"
        End If

        If Show_Deleted = False Then
            sSQL = sSQL & " AND e.is_deleted='1'"
        End If

        If QueryHasRows(sSQL) Then
            GetUserByCredentials = True
        Else
            GetUserByCredentials = False
        End If
    End Function

    Public Function CheckEmployeeNameExisting(ByVal fname As String, ByVal lname As String, ByVal mname As String, _
                                        ByVal e_id As String) As Boolean

        Dim sSQL As String = "SELECT * FROM employees WHERE e_fname='" & fname & _
                        "' AND e_lname='" & fname & "' AND e_mname='" & mname & "' "

        If Len(e_id) > 0 Then
            sSQL = sSQL & "WHERE e_id <> '" & e_id & "' "
        End If
        sSQL = sSQL & "LIMIT 1"

        If QueryHasRows(sSQL) Then
            CheckEmployeeNameExisting = True
        Else
            CheckEmployeeNameExisting = False
        End If
    End Function

    Public Function CheckUserNameExisting(ByVal fname As String, ByVal lname As String, ByVal mname As String, _
                                    ByVal e_id As String) As Boolean

        Dim sSQL As String = "SELECT * FROM employees WHERE e_fname='" & fname & _
                        "' AND e_lname='" & fname & "' AND e_mname='" & mname & "' "

        If Len(e_id) > 0 Then
            sSQL = sSQL & "WHERE e_id <> '" & e_id & "' "
        End If
        sSQL = sSQL & "LIMIT 1"

        If QueryHasRows(sSQL) Then
            CheckUserNameExisting = True
        Else
            CheckUserNameExisting = False
        End If
    End Function

    Public Function AddEmployees(ByVal E As employees) As Boolean
        On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim sSQL As String = "INSERT INTO employees VALUE (null,@e_fname,@e_lname,@e_mname,@e_username,md5(@e_password),'" & E.mug_id & "','" & E.is_deleted & "')"
        Dim com As New MySqlCommand(sSQL, con)
        com.Parameters.AddWithValue("@e_fname", E.e_fname)
        com.Parameters.AddWithValue("@e_lname", E.e_lname)
        com.Parameters.AddWithValue("@e_mname", E.e_mname)
        com.Parameters.AddWithValue("@e_username", E.e_username)
        com.Parameters.AddWithValue("@e_password", E.e_password)
        com.ExecuteNonQuery()
        com.Parameters.Clear()
        con.Close()
        AddEmployees = True
        Exit Function
err:
        AddEmployees = False
        DisplayErrorMsg("modEmployees", "AddEmployees", Err.Number, Err.Description)
    End Function

    Public Function UpdateEmployees(ByVal E As employees) As Boolean
        On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim sSQL As String = "UPDATE employees SET e_fname=@e_fname,e_lname=@e_lname,e_mname=@e_mname,e_username=@e_username,e_password=md5(@e_password) WHERE e_id='" & E.e_id & "'"
        Dim com As New MySqlCommand(sSQL, con)
        com.Parameters.AddWithValue("@e_fname", E.e_fname)
        com.Parameters.AddWithValue("@e_lname", E.e_lname)
        com.Parameters.AddWithValue("@e_mname", E.e_mname)
        com.Parameters.AddWithValue("@e_username", E.e_username)
        com.Parameters.AddWithValue("@e_password", E.e_password)
        com.ExecuteNonQuery()
        com.Parameters.Clear()
        con.Close()
        UpdateEmployees = True
        Exit Function
err:
        UpdateEmployees = False
        DisplayErrorMsg("modEmployees", "UpdateEmployees", Err.Number, Err.Description)
    End Function

    Public Function Delete_Employees(ByVal e_id As String) As Boolean
        If ExecuteQry("DELETE FROM employees WHERE e_id='" & e_id & "'") Then
            Delete_Employees = True
        Else
            Delete_Employees = False
        End If
    End Function

    Public Function DeleteEmployee(ByVal emp As employees) As Boolean
        On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim sSQL As String = "UPDATE employees SET is_deleted=1 WHERE e_id='" & emp.e_id & "'"
        Dim com As New MySqlCommand(sSQL, con)
        com.ExecuteNonQuery()
        com.Parameters.Clear()
        con.Close()
        DeleteEmployee = True
        Exit Function
err:
        DeleteEmployee = False
        DisplayErrorMsg("modEmployees", "DeleteEmployee", Err.Number, Err.Description)
    End Function

    Public Function GetEmployeeByID(ByVal e_id As Integer, ByRef E As employees) As Boolean
        '  On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim com As New MySqlCommand("Select E.e_id,E.e_fname,E.e_lname,E.e_mname,E.e_username,M.mug_name,E.is_deleted,M.mug_id,E.e_password,E.e_username  from employees E JOIN module_user_groups M ON M.mug_id = E.mug_id where E.e_id = '" & e_id & "'", con)
        Dim vRS As MySqlDataReader = com.ExecuteReader()
        vRS.Read()
        If vRS.HasRows Then
            With E
                .e_fname = vRS("e_fname").ToString()
                .e_id = vRS("e_id").ToString()
                .e_lname = vRS("e_lname").ToString()
                .e_mname = vRS("e_mname").ToString()
                .e_password = vRS("e_password").ToString()
                .e_username = vRS("e_username").ToString()
                .is_deleted = BooleanToInt(vRS("is_deleted").ToString())
                .mug_id = vRS("mug_id").ToString()
                .mug_name = vRS("mug_name").ToString()
            End With
            GetEmployeeByID = True
        Else
            GetEmployeeByID = False
        End If
        con.Close()
        'Get_Employee_By_ID = True
        'Exit Function
        'err:
        'Get_Employee_By_ID = False
        'DisplayErrorMsg("modEmployees", "Get_Employee_By_ID", Err.Number, Err.Description)
    End Function

End Module
