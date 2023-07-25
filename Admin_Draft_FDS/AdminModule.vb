Imports System.Data.Common
Imports System.Reflection
Imports CrystalDecisions.CrystalReports.Engine
Imports Microsoft.VisualBasic.ApplicationServices
Imports MySql.Data.MySqlClient

Module AdminModule

    '===============================================================

    Public databasename As String = "mock_db"
    Public tablename As String = "product"

    'Data Source =localhost; Database= mock_db; User= root; Password=;

    Public connStr As String = "Data Source =localhost; Database=" + databasename + "; User= root; Password=;"
    Public conn As MySqlConnection = New MySqlConnection(connStr)

    Public sql As String
    Public dbcomm As MySqlCommand
    Public dbread As MySqlDataReader
    Public dataAdapter, dataAdptr As MySqlDataAdapter
    Public ds As DataSet = New DataSet
    Public dt As DataTable = New DataTable()

    'SQL Employee Join Statement
    Public employeeJoin As String =
    " INNER JOIN user_has_department ON user.id = user_has_department.user_id
    INNER JOIN department ON user_has_department.department_id = department.id
    INNER JOIN login ON user.id = login.user_id
    INNER JOIN role ON login.role_id = role.id "

    ' SQL Statement to Data Table
    Public Function dtablefill(statement As String) As DataTable
        dataAdapter = New MySqlDataAdapter(statement, conn)
        dt = New DataTable
        dataAdapter.Fill(dt)

        Return dt
    End Function

    ' DataTable to Arraylist
    Public Function dtable2array(dtable As DataTable) As ArrayList
        Dim columnCount As Integer = dtable.Columns.Count
        Dim rowCount As Integer = dtable.Rows.Count
        Dim arrRow As New ArrayList()

        For Each row As DataRow In dtable.Rows
            For col = 0 To columnCount - 1
                arrRow.Add((row(col)).ToString)
            Next
        Next

        Return arrRow
    End Function

    ' POPULATE
    Public Function populateFunction(sqlStatement As String, ByRef rowLen As Integer, ByRef colLen As Integer) As ArrayList
        ' DataAdapter to DataTable
        Dim dtTable As DataTable = dtablefill(sqlStatement)
        colLen = dtTable.Columns.Count
        rowLen = dtTable.Rows.Count

        ' Data Table to ArrayList
        Dim arrList As ArrayList = dtable2array(dtTable)

        Return arrList
    End Function

    ' TRY CRYSTAL REPORT WIP
    Public Sub populate2nd()
        'select id, item_name, category from product;

        dataAdapter = New MySqlDataAdapter("Select * from " + tablename + "", conn)
        dt = New DataTable

        dataAdapter.Fill(dt)
    End Sub

    ' =================== SEARCH FILTER ========================

    ' SQL Statement to DataAdapter to DataTable
    Public Function contentSearcher(slctColumns As String, table As String, StatementCondition As String) As DataTable
        Dim sqlstmt As String = $"SELECT {slctColumns} FROM {table} WHERE {StatementCondition}"
        Dim dtable As New DataTable
        Try
            conn.Open()
            dataAdapter = New MySqlDataAdapter(sqlstmt, conn)
            dataAdapter.Fill(dtable)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            conn.Close()
        End Try

        Return dtable
    End Function

    ' SQL Statement to DataAdapter to DataTable ORDER BY
    Public Function groupBySeacrh(slctColumns As String, table As String, StatementCondition As String) As DataTable
        Dim sqlstmt As String = $"SELECT {slctColumns} FROM {table} ORDER BY {StatementCondition}"
        Dim dtable As New DataTable
        Try
            conn.Open()
            dataAdapter = New MySqlDataAdapter(sqlstmt, conn)
            dataAdapter.Fill(dtable)


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            conn.Close()
        End Try

        Return dtable
    End Function

    'Data Table to FlowLayoutPanelRows (exclusive only for product search)
    Public Sub dispProdSearch(dtable As DataTable)
        AdminMain.ProductListFLP.Controls.Clear()

        For Each row As DataRow In dtable.Rows
            Dim iid As Integer = CInt(row("id"))
            Dim nname As String = CStr(row("item_name"))
            Dim ccategory As String = CStr(row("category"))

            AdminMain.createProdPanel(iid, nname, ccategory)
        Next
    End Sub

    'Data Table to FlowLayoutPanelRows (exclusive only for employee search)
    Public Sub dispEmpSearch(dtable As DataTable)
        AdminMain.EmployeeListFLP.Controls.Clear()

        For Each row As DataRow In dtable.Rows
            Dim iid As Integer = CInt(row("id"))
            Dim usnm As String = CStr(row("username"))
            Dim role As String = CStr(row("rolenm"))
            Dim dept As String = CStr(row("dptnm"))

            AdminMain.createEmployPanel(iid, usnm, role, dept)
        Next
    End Sub

    'Open AdminComponentInfo
    Public Sub compInfo(dtable As DataTable)
        For Each row As DataRow In dtable.Rows
            For Each col As DataColumn In dtable.Columns

                Dim colname As String = CStr(col.ColumnName)
                Dim colobject As Object = (row(col.ColumnName))
                Dim content As String = colobject.ToString

                AdmnComponentInfo.generataInfoPanel(colname, content)
            Next
        Next
        AdmnComponentInfo.Show()
    End Sub

    'SQL Statment Execute Non Query
    Public Function exquery(statement As String) As DataTable

        Dim dtable As New DataTable
        Try
            conn.Open()
            dataAdapter = New MySqlDataAdapter(statement, conn)
            dataAdapter.Fill(dtable)

            MessageBox.Show("Executed Successfully")

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            conn.Close()
        End Try
        Return dtable
    End Function






End Module
