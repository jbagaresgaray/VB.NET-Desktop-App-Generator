'======================================================================================
'Filename:  modFunctions.vb
'Type:      Module
'Author:    Philip Cesar B.Garay
'Date:      15.Jul.2013
'Email:     philipgaray2@gmail.com
'
'Purpose:   This module contains all the functions and procedures for the whole application.
'           Functions that can be usable and applicable anyware in the application
'           
'
'
'=======================================================================================

Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.Common
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module modFunctions

    Public CURRENT_USER As employees

    '===============================================
    ' PURPOSE: use to determine whether the Form Transaction State is Adding or Updating Record.
    '===============================================
    Public Enum FormState
        addFormState = 1
        updateFormState = 2
        deleteFormState = 3
        readFormState = 4
    End Enum


    ' ===========================================================
    '   Use to check if Texbox has input value
    ' ===========================================================
    Public Function CheckTextBox(ByRef txt As System.Windows.Forms.TextBox, Optional ByVal sMSG As String = "TextBox") As Boolean
        On Error Resume Next
        txt.Text = txt.Text.Trim()
        If txt.Text = "" Then
            MsgBox(sMSG, vbExclamation)

            txt.Text = ""
            txt.Focus()

            CheckTextBox = False
        Else
            CheckTextBox = True
        End If
    End Function

    ' ===========================================================
    '   Use to check if Combobox has input value
    ' ===========================================================
    Public Function CheckCombobox(ByRef cmb As System.Windows.Forms.ComboBox, Optional ByVal sMSG As String = "Combobox") As Boolean
        On Error Resume Next
        If cmb.SelectedIndex = -1 And cmb.Text.Trim() = "" Then
            MsgBox(sMSG, vbExclamation)

            cmb.Focus()

            CheckCombobox = False
        Else
            CheckCombobox = True
        End If
    End Function

    ' ===========================================================
    '   Use to populate Listview
    ' ===========================================================
    Public Sub FillListView(ByVal sSQL As String, ByRef ListView As ListView, Optional ByVal ImageIndex As Integer = 0)
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim c As Integer
        Dim com As New MySqlCommand(sSQL, con)
        Dim vRS As MySqlDataReader = com.ExecuteReader()

        ListView.Items.Clear()
        Do While vRS.Read()
            Dim lv As New ListViewItem(vRS(0).ToString())

            For c = 1 To vRS.FieldCount - 1
                lv.SubItems.Add(vRS.Item(c).ToString)
            Next
            lv.ImageIndex = ImageIndex


            ListView.Items.AddRange(New ListViewItem() {lv})
        Loop
        vRS.Close()
        con.Close()
    End Sub

    ' ===========================================================
    '   Use to populate Combobox, Remember to set the first field to 'Value' and the second to 'Display' in your query
    ' ===========================================================
    Public Sub FillCombobox(ByVal sSQL As String, ByRef ComboBox As ComboBox)
        Dim dv As New DataViewManager()
        Dim da As New MySqlDataAdapter()
        Dim ds As New DataSet()
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim com As New MySqlCommand(sSQL, con)
        da.SelectCommand = New MySqlCommand(sSQL, con)
        da.TableMappings.Add("Table", "table")
        da.Fill(ds)
        dv = ds.DefaultViewManager

        ComboBox.DataSource = dv
        ComboBox.ValueMember = "table.Value"
        ComboBox.DisplayMember = "table.Display"
        ComboBox.SelectedIndex = -1
        con.Close()
    End Sub

    ' ===========================================================
    '   Use to populate Combobox
    ' ===========================================================
    Public Sub fillCombo(ByVal cmb As ComboBox, ByVal sSQL As String)
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim com As New MySqlCommand(sSQL, con)
        Dim reader As MySqlDataReader = com.ExecuteReader()
        cmb.Items.Clear()
        While reader.Read()
            cmb.Items.AddRange(New Object() {reader(0).ToString()})
        End While
        reader.Close()
        con.Close()
    End Sub

    ' NOTE: FillDataGridView() modified on 9/13/2013 by Arnel. Added the OPTIONAL paramater [flushData]
    '       If [flushData] is set to false then the function will simply append the result of [sSQL] to the [DataGridView]
    Public Sub FillDataGridView(ByVal sSQL As String, ByRef DataGridView As DataGridView, Optional ByVal flushData As Boolean = True) ' Fill up DatagridView with data from tables
        Dim c As Integer

        If flushData Then DataGridView.Rows.Clear()


        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim com As New MySqlCommand(sSQL, con)
        Dim vRS As MySqlDataReader = com.ExecuteReader()

        While vRS.Read()
            With DataGridView
                Dim n As Integer = .Rows.Add
                For c = 0 To vRS.FieldCount - 1
                    .Rows(n).Cells(c).Value = vRS.Item(c).ToString
                Next
                '.FirstDisplayedScrollingRowIndex = n
                '.CurrentCell = .Rows(n).Cells(0)
                '.Rows(n).Selected = True
            End With
        End While
        vRS.Close()
        con.Close()
    End Sub

    ' ===========================================================
    '   GENERATE CUSTOM ERROR MESSAGE
    ' ===========================================================
    Public Function DisplayErrorMsg(ByVal pModule As String, ByVal pProcedure As String, ByVal ErrorNbr As Long, ByVal ErrorDesc As String)

        Dim strErrorMsg As String

        On Error Resume Next

        strErrorMsg = "------------------------------------------------------------------------------------------------------------------------------------------" & vbNewLine & _
                      "Error          : " & ErrorNbr & vbNewLine & _
                      "Description: " & ErrorDesc & vbNewLine & _
                      "------------------------------------------------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine & _
                      "Module       : " & pModule & vbNewLine & _
                      "Procedure: " & pProcedure & vbNewLine & vbNewLine & _
                      "If this is the first time you saw this message, " & _
                      "kindly inform your Database Administrator" & vbNewLine & _
                      "or notify the Technical Support of People Index Solutions." & vbNewLine & vbNewLine & _
                      "info@peopleindexsolutions.com " & vbNewLine & vbNewLine & _
                      "Indicate the information shown in this dialog and what you were doing when this error occured." & vbNewLine & vbNewLine & _
                      "Thank you."

        MsgBox(strErrorMsg, vbCritical + vbOKOnly, "Error No. " & ErrorNbr)

    End Function



    ' ===========================================================
    '   Use to execute SQL Query Commands ex. INSERT, UPDATE, DELETE
    ' ===========================================================
    Public Function ExecuteQry(ByVal sSQL As String) As Boolean
        On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim com As New MySqlCommand(sSQL, con)
        com.ExecuteNonQuery()
        con.Close()
        ExecuteQry = True
        Exit Function
err:
        ExecuteQry = False
        DisplayErrorMsg("modFunction", "ExecuteQry", Err.Number, Err.Description)
    End Function

    ' ===========================================================
    '   Use to execute SQL Query Commands with return values
    ' ===========================================================
    Public Function ExecuteQryReturn(ByVal sSQL As String) As String
        On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim com As New MySqlCommand(sSQL, con)
        ExecuteQryReturn = CStr(com.ExecuteScalar())
        con.Close()
        Exit Function
err:
        ExecuteQryReturn = ""
    End Function

    ' ===========================================================
    '   Use for record checking , duplicate data checking
    ' ===========================================================
    Public Function QueryHasRows(ByVal sSQL As String) As Boolean
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim com As New MySqlCommand(sSQL, con)
        Dim vRS As MySqlDataReader = com.ExecuteReader()
        vRS.Read()
        If vRS.HasRows Then
            QueryHasRows = True
        Else
            QueryHasRows = False
        End If
        vRS.Close()
        con.Close()
    End Function

    ' ===========================================================
    '   Use to retrieve Blob photo on Database
    ' ===========================================================
    Public Sub GetPhoto(ByVal sSQL As String, ByRef Pic As PictureBox)
        On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()

        Dim adapter As New MySqlDataAdapter
        adapter.SelectCommand = New MySqlCommand(sSQL, con)
        Dim Data As New System.Data.DataTable
        Dim commandbuild As New MySqlCommandBuilder(adapter)
        adapter.Fill(Data)

        Dim lb() As Byte = Data.Rows(0).Item(0)
        Dim lstr As New System.IO.MemoryStream(lb)
        Pic.Image = Image.FromStream(lstr)
        Pic.SizeMode = PictureBoxSizeMode.StretchImage
        lstr.Close()
        con.Close()
        Exit Sub

err:
        'Pic.Image = My.Resources.no_image
        Debug.Print(Err.Number & " - " & Err.Description)
        'DisplayErrorMsg("modFunction", "GetPhoto", Err.Number, Err.Description)
    End Sub

    ' ===========================================================
    '   Use to retrieve Blob photo on Database
    ' ===========================================================

    Public Function SavePhoto(ByVal sSQL As String, Photo As PictureBox) As Boolean
        On Error GoTo Err

        Dim FileSize As UInt32

        Dim mstream As New System.IO.MemoryStream()
        Photo.Image.Save(mstream, System.Drawing.Imaging.ImageFormat.Jpeg)
        Dim arrImage() As Byte = mstream.GetBuffer()

        FileSize = mstream.Length
        mstream.Close()

        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()

        Dim com As New MySqlCommand(sSQL, con)
        com.Parameters.AddWithValue("@Pic", arrImage)

        com.ExecuteNonQuery()
        com.Parameters.Clear()
        con.Close()

        SavePhoto = True

        Exit Function
Err:
        SavePhoto = False
        DisplayErrorMsg("modFunctions", "SavePhoto", Err.Number, Err.Description)
    End Function

    ' ===========================================================
    '   Use to Hash Inputs based on Database Hash standards
    ' ===========================================================

    Public Function fn_HashPass(ByVal pass As String) As String
        On Error GoTo err
        Dim con As New MySqlConnection(DB_CONNECTION_STRING)
        con.Open()
        Dim com As New MySqlCommand("SELECT md5('" & pass & "')", con)
        Dim vRS As MySqlDataReader = com.ExecuteReader
        vRS.Read()
        If vRS.HasRows Then
            fn_HashPass = vRS(0).ToString()
        Else
            fn_HashPass = ""
        End If
        vRS.Close()
        con.Close()
        Exit Function
err:
        fn_HashPass = ""
        DisplayErrorMsg("modCategories", "fn_HashPass", Err.Number, Err.Description)
    End Function

    ' ===========================================================
    '   Use to converts boolean to integer
    ' ===========================================================

    Public Function BooleanToInt(ByVal Value As Boolean) As Integer
        Select Case Value
            Case False
                Return 0
            Case Else
                Return 1
        End Select
    End Function

    ' ===========================================================
    '   Use to convert integer to boolean
    ' ===========================================================
    Public Function IntToBoolean(ByVal Value As Integer) As Boolean
        Select Case Value
            Case 0
                IntToBoolean = False
            Case Else
                IntToBoolean = True
        End Select
    End Function

    ' ===========================================================
    '   Use to Boolean to Word
    ' ===========================================================

    Public Function BooleanToWord(ByVal Value As Boolean) As String
        Select Case Value
            Case True
                BooleanToWord = "YES"
            Case Else
                BooleanToWord = "NO"
        End Select
    End Function

    ' ===========================================================
    '   Use to convert Integer to Word
    ' ===========================================================

    Public Function IntToWord(ByVal Value As Boolean) As String
        Select Case Value
            Case 0
                IntToWord = "NO"
            Case Else
                IntToWord = "YES"
        End Select
    End Function

    ' ===========================================================
    '   Use to Convert Short Gender to Long Gender
    ' ===========================================================
    Public Function CGender(ByVal Gender As String) As String
        Select Case Gender
            Case "M"
                CGender = "Male"
            Case "F"
                CGender = "Female"
        End Select
    End Function

    ' ===========================================================
    '   Use to convert Military Time to Standard Time
    ' ===========================================================
    Public Function MilitaryToStandard(ByVal pTimeValue As Long) As String

        MilitaryToStandard = CDate(Format(pTimeValue, "00:00:00"))

    End Function

    ' ===========================================================
    '   Use to convert Standard Time to Military Time
    ' ===========================================================
    Public Function toMilitaryTime(ByVal pTime As String) As String

        toMilitaryTime = Format(CDate(pTime), "HHmmss")

    End Function

    ' ===========================================================
    '   Use to Check Email Address Inputs
    ' ===========================================================

    Public Function EmailAddressCheck(ByVal emailAddress As String) As Boolean

        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If

    End Function

    ' ===========================================================
    '   Use to Get IPv4Address on the workstation
    ' ===========================================================
    Public Function GetIPv4Address() As String
        GetIPv4Address = String.Empty
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                GetIPv4Address = ipheal.ToString()
            End If
        Next

    End Function

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub

    ' ===========================================================
    '   Use To limit textbox input to numeric characters only
    ' ===========================================================

    Public Sub FilterNumericInput(ByRef e As System.Windows.Forms.KeyEventArgs)
        Dim keyData As Keys = e.KeyCode

        If (keyData >= Keys.D0 And keyData <= Keys.D9) Or (keyData >= Keys.NumPad0 And keyData <= Keys.NumPad9) Or (e.Modifiers = Keys.Control AndAlso keyData = Keys.A) Or keyData = Keys.Back Or keyData = Keys.LShiftKey Or keyData = Keys.RShiftKey Or keyData = Keys.RControlKey Or keyData = Keys.LControlKey Or keyData = Keys.Tab _
          Or keyData = Keys.Home Or keyData = Keys.Delete Or keyData = Keys.End Or keyData = Keys.Left Or keyData = Keys.Right Or keyData = Keys.OemPeriod Or keyData = Keys.Decimal Then

            Exit Sub
        Else
            e.SuppressKeyPress = True
            e.Handled = True
        End If
    End Sub

    ' ===========================================================
    '   Use To limit input textbox value
    ' ===========================================================

    Public Sub LimitTxtboxVal(ByRef txtVal As String, ByVal maxVal As Integer)
        Dim percentVal As Decimal = CDec(txtVal)

        If percentVal > maxVal Then
            txtVal = maxVal
        End If
    End Sub

    ' ===========================================================
    '   Use To check and get selected row index ID
    ' ===========================================================

    Public Function checkSelectedRow(ByVal listviewObject As ListView, ByVal IDindex As Integer) As Integer
        Dim returnVal As Integer = 0

        If listviewObject.Items.Count > 0 And listviewObject.SelectedItems.Count > 0 Then
            returnVal = CInt(listviewObject.FocusedItem.SubItems(IDindex).Text)
        End If

        Return returnVal
    End Function

    Public Function GetFieldValue(ByVal srcSQL As String, ByVal strField As String) As Decimal
        Try
            Dim returnVal As String
            Dim con As New MySqlConnection(DB_CONNECTION_STRING)
            con.Open()
            Dim cmd As New MySqlCommand(srcSQL, con)
            Dim rdr As MySqlDataReader = cmd.ExecuteReader
            rdr.Read()
            If rdr.HasRows Then
                returnVal = rdr(strField).ToString()
            Else
                returnVal = ""
            End If
            rdr.Close()
            con.Close()

            If Len(returnVal) > 0 Then
                GetFieldValue = returnVal
            Else
                GetFieldValue = 0
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Function Strip(ByVal pString As String) As String

        On Error Resume Next

        Dim strResult As String

        strResult = Replace(pString, "'", "''")         'replace all single-quotes with double-single-quote

        strResult = Replace(strResult, "&", "&&")       'replace all ampersand with double-ampersand

        Strip = Trim(strResult)

    End Function

    Public Function ConvertHoursIntoDecimal(ByVal D As Date) As Decimal
        Dim TB() As String

        TB = Split(D, ":")

        ConvertHoursIntoDecimal = TB(0) + ((TB(1) * 100) / 60) / 100
    End Function

    Public Function ConvertDecimalIntoHours(ByVal D As Decimal) As String
        Dim e As String
        'For example
        e = CStr(Math.Round((D - Int(D)) / 100 * 60, 2)) & "0"
        ConvertDecimalIntoHours = CStr(Int(D)) & ":" & Mid(e, 3, 2)
    End Function

    ' ===========================================================
    '   Use to convert Monitary to Decimal
    ' ===========================================================

    Public Function Words_Money(ByVal num As Decimal) As String
        Dim dollars As Decimal
        Dim cents As Integer
        Dim dollars_result As String
        Dim cents_result As String

        ' Dollars.
        dollars = Int(num)
        dollars_result = Words_1_all(dollars)
        If Len(dollars_result) = 0 Then dollars_result = "zero"

        If dollars_result = "one" Then
            dollars_result = dollars_result & " peso"
        Else
            dollars_result = dollars_result & " pesos"
        End If

        ' Cents.
        cents = CInt((num - dollars) * 100.0#)
        cents_result = Words_1_all(cents)
        If Len(cents_result) = 0 Then
            cents_result = ""
        ElseIf cents_result = "one" Then
            cents_result = cents_result & " centavo only"
        Else
            cents_result = cents_result & " centavos only"
        End If

        ' Combine the results.


        If Len(cents_result) > 0 Then
            Words_Money = dollars_result & _
            " and " & cents_result
        Else
            Words_Money = dollars_result & " only"
        End If

    End Function
    ' Return words for this value between 1 and 999.
    Private Function Words_1_999(ByVal num As Integer) As String
        Dim hundreds As Integer
        Dim remainder As Integer
        Dim result As String = ""

        hundreds = num \ 100
        remainder = num - hundreds * 100

        If hundreds > 0 Then
            result = Words_1_19(hundreds) & " hundred "
        End If

        If remainder > 0 Then
            result = result & Words_1_99(remainder)
        End If

        Words_1_999 = Trim$(result)
    End Function
    ' Return a word for this value between 1 and 99.
    Private Function Words_1_99(ByVal num As Integer) As String
        Dim result As String = ""
        Dim tens As Integer

        tens = num \ 10

        If tens <= 1 Then
            ' 1 <= num <= 19
            result = result & " " & Words_1_19(num)
        Else
            ' 20 <= num
            ' Get the tens digit word.
            Select Case tens
                Case 2
                    result = "twenty"
                Case 3
                    result = "thirty"
                Case 4
                    result = "forty"
                Case 5
                    result = "fifty"
                Case 6
                    result = "sixty"
                Case 7
                    result = "seventy"
                Case 8
                    result = "eighty"
                Case 9
                    result = "ninety"
            End Select

            ' Add the ones digit number.
            result = result & " " & Words_1_19(num - tens * 10)
        End If

        Words_1_99 = Trim$(result)
    End Function
    ' Return a word for this value between 1 and 19.
    Private Function Words_1_19(ByVal num As Integer) As String
        Words_1_19 = ""
        Select Case num
            Case 1
                Words_1_19 = "one"
            Case 2
                Words_1_19 = "two"
            Case 3
                Words_1_19 = "three"
            Case 4
                Words_1_19 = "four"
            Case 5
                Words_1_19 = "five"
            Case 6
                Words_1_19 = "six"
            Case 7
                Words_1_19 = "seven"
            Case 8
                Words_1_19 = "eight"
            Case 9
                Words_1_19 = "nine"
            Case 10
                Words_1_19 = "ten"
            Case 11
                Words_1_19 = "eleven"
            Case 12
                Words_1_19 = "twelve"
            Case 13
                Words_1_19 = "thirteen"
            Case 14
                Words_1_19 = "fourteen"
            Case 15
                Words_1_19 = "fifteen"
            Case 16
                Words_1_19 = "sixteen"
            Case 17
                Words_1_19 = "seventeen"
            Case 18
                Words_1_19 = "eightteen"
            Case 19
                Words_1_19 = "nineteen"
        End Select
    End Function

    Private Function Words_1_all(ByVal num As Decimal) As String
        Dim power_value(0 To 4) As Decimal
        Dim power_name(0 To 4) As String
        Dim digits As Integer
        Dim result As String
        Dim i As Integer

        ' Initialize the power names and values.
        power_name(0) = "trillion" : power_value(0) = 1000000000000.0#
        power_name(1) = "billion" : power_value(1) = 1000000000
        power_name(2) = "million" : power_value(2) = 1000000
        power_name(3) = "thousand" : power_value(3) = 1000
        power_name(4) = "" : power_value(4) = 1

        For i = 0 To 4
            ' See if we have digits in this range.
            If num >= power_value(i) Then
                ' Get the digits.
                digits = Int(num / power_value(i))

                ' Add the digits to the result.
                If Len(result) > 0 Then result = result & ", "
                result = result & _
                    Words_1_999(digits) & _
                    " " & power_name(i)

                ' Get the number without these digits.
                num = num - digits * power_value(i)
            End If
        Next i

        Words_1_all = Trim(result)
    End Function


    ' ===========================================================
    '   Use to Export Listview Data to Excel
    ' ===========================================================
    Public Sub ListViewExportToExcel(ByVal ListView1 As System.Windows.Forms.ListView)
        Try
            Dim objExcel As New Excel.Application
            Dim bkWorkBook As Workbook
            Dim shWorkSheet As Worksheet
            Dim i As Integer
            Dim j As Integer

            objExcel = New Excel.Application
            bkWorkBook = objExcel.Workbooks.Add
            shWorkSheet = CType(bkWorkBook.ActiveSheet, Worksheet)
            For i = 0 To ListView1.Columns.Count - 1
                shWorkSheet.Cells(1, i + 1) = ListView1.Columns(i).Text
            Next
            For i = 0 To ListView1.Items.Count - 1
                For j = 0 To ListView1.Items(i).SubItems.Count - 1
                    shWorkSheet.Cells(i + 2, j + 1) = ListView1.Items(i).SubItems(j).Text
                Next
            Next

            objExcel.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ' ===========================================================
    '   Use to export datagridview values to excel with headers
    ' ===========================================================

    Public Sub DatagridViewToExcel(ByVal dgView As System.Windows.Forms.DataGridView)
        Dim ExcelApp As Object, ExcelBook As Object
        Dim ExcelSheet As Object
        Dim i As Integer
        Dim j As Integer

        'create object of excel
        ExcelApp = CreateObject("Excel.Application")
        ExcelBook = ExcelApp.WorkBooks.Add
        ExcelSheet = ExcelBook.WorkSheets(1)

        With ExcelSheet
            ' HEADER
            For Each column As DataGridViewColumn In dgView.Columns
                .cells(1, column.Index + 1) = column.HeaderText
            Next
            ' DATA
            For i = 1 To dgView.RowCount
                .cells(i + 1, 1) = dgView.Rows(i - 1).Cells(0).Value
                For j = 1 To dgView.Columns.Count - 1
                    .cells(i + 1, j + 1) = dgView.Rows(i - 1).Cells(j).Value
                Next
            Next
        End With

        ExcelApp.Visible = True
        '
        ExcelSheet = Nothing
        ExcelBook = Nothing
        ExcelApp = Nothing
    End Sub

End Module
