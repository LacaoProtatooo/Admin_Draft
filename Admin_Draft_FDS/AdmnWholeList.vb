Imports System.Diagnostics.Eventing.Reader
Imports System.Runtime.Remoting
Imports System.Web.UI.WebControls
Imports CrystalDecisions.CrystalReports.Engine

Public Class AdmnWholeList
    Dim arrlist As New ArrayList
    Dim arrListContent As New ArrayList
    Dim dgridSelectedRow, optable, stm As String

    Private Sub AdmnWholeList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        productORemployee()

    End Sub

    Private Sub AdmnWholeList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        AdminMain.Show()

    End Sub

    Public Sub dgridPopulate(statement As String)
        Dim dtable As DataTable = dtablefill(statement)
        dgrid.DataSource = dtable.DefaultView

    End Sub

    Private Sub dgrid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgrid.CellClick
        ' Data Grid Cell Click
        arrlist.Clear()
        arrListContent.Clear()
        Button2.Enabled = True

        txtbxSelectedRow.Text = dgrid.Rows(e.RowIndex).Cells(0).Value
        Dim slctRow As String = txtbxSelectedRow.Text
        dgridSelectedRow = slctRow

        Select Case optable
            Case "product"
                flpProductLoad(slctRow)
            Case "user"
                flpEmployeeLoad(slctRow)
        End Select

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Add / Insert
        productORemployee()

        If (optable = "product") Then
            Try
                Dim tbl As String = optable
                AdmnInsert.prodSetup(tbl)
                AdmnInsert.Show()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
                resetform()
            End Try
        Else
            'Employee
            Try
                Dim tbl As String = optable
                AdmnInsert.empSetup(tbl)
                AdmnInsert.Show()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
                resetform()
            End Try
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Show Specific Columns


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Edit Row
        Try
            Dim upsatement As String = updateStatement(arrlist, arrListContent, dgridSelectedRow, optable)
            Dim dtable As DataTable = exquery(upsatement)
            dgrid.DataSource = dtable.DefaultView
            resetform()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            resetform()
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Delete Row
        Try
            Dim delst As String = deleteStatement()
            Dim dtable As DataTable = exquery(delst)
            dgrid.DataSource = dtable.DefaultView
            resetform()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try
    End Sub

    Public Sub flpProductLoad(selectedRow As String)
        viewFLP.Controls.Clear()
        Dim selectedColumns As String = "*"
        Dim table As String = "product"
        Dim condition As String = $"id = {selectedRow}"

        Try
            Dim dtable As DataTable = contentSearcher(selectedColumns, table, condition)

            For Each row As DataRow In dtable.Rows
                For Each col As DataColumn In dtable.Columns

                    Dim colname As String = CStr(col.ColumnName)
                    Dim colobject As Object = (row(col.ColumnName))
                    Dim content As String = colobject.ToString

                    slctRowFLP(colname, content)
                Next
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub flpEmployeeLoad(selectedRow As String)
        viewFLP.Controls.Clear()
        Dim selectedColumns As String = " user.*, role.name as rolenm, department.name as dptnm "
        Dim table As String = "user" + employeeJoin
        Dim condition As String = $"user.id = {selectedRow}"

        Try
            Dim dtable As DataTable = contentSearcher(selectedColumns, table, condition)

            For Each row As DataRow In dtable.Rows
                For Each col As DataColumn In dtable.Columns

                    Dim colname As String = CStr(col.ColumnName)
                    Dim colobject As Object = (row(col.ColumnName))
                    Dim content As String = colobject.ToString

                    slctRowFLP(colname, content)
                Next
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Sub slctRowFLP(col As String, coldata As String)
        Dim contentPanel = New System.Windows.Forms.Panel()
        Dim lbl_info = New System.Windows.Forms.Label()
        Dim dummytextbox = New System.Windows.Forms.TextBox()
        contentPanel.SuspendLayout()

        contentPanel.Size = New System.Drawing.Size(470, 35)
        lbl_info.Text = col
        dummytextbox.Text = coldata

        '
        'contentPanel
        '
        'contentPanel.Controls.Add(lbl_content)
        contentPanel.Controls.Add(lbl_info)
        contentPanel.Controls.Add(dummytextbox)
        contentPanel.Location = New System.Drawing.Point(3, 3)
        contentPanel.Name = "contentPanel"
        contentPanel.TabIndex = 0
        contentPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        '
        'lbl_info
        '
        lbl_info.AutoSize = True
        lbl_info.Font = New System.Drawing.Font("Lexend", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        lbl_info.Location = New System.Drawing.Point(7, 7)
        lbl_info.Name = "lbl_info"
        lbl_info.Size = New System.Drawing.Size(53, 21)
        lbl_info.TabIndex = 0
        '
        'dummytextbox
        '
        dummytextbox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        dummytextbox.Font = New System.Drawing.Font("Lexend", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        dummytextbox.Location = New System.Drawing.Point(134, 7)
        dummytextbox.Name = "dummytextbox"
        dummytextbox.Size = New System.Drawing.Size(330, 24)
        dummytextbox.TabIndex = 0

        contentPanel.ResumeLayout(False)
        contentPanel.PerformLayout()
        viewFLP.Controls.Add(contentPanel)

        ' Add Event Handler to TextBoxes
        AddHandler dummytextbox.TextChanged, AddressOf dummytextbox_TextChanged
        dummytextbox.Name = col

    End Sub

    Private Sub dummytextbox_TextChanged(sender As Object, e As EventArgs)
        Dim txtbx As Windows.Forms.TextBox
        txtbx = CType(sender, Object)
        Dim name As String = txtbx.Name
        Dim editedText As String = txtbx.Text
        Dim op As String = productORemployee()
        Button4.Enabled = True

        ' Selected Column
        If arrlist.Count = 0 Then
            arrlist.Add(name)
            arrListContent.Add(editedText)
        Else
            Dim index As Integer = arrlist.IndexOf(name)

            If index = -1 Then
                arrlist.Add(name)
                arrListContent.Add(editedText)
            Else
                arrListContent(index) = editedText
            End If
        End If


    End Sub

    ' Update Statement Builder
    Private Function updateStatement(arrColname As ArrayList, arrContent As ArrayList, selectedRow As String, table As String) As String
        Dim condition As String = String.Empty
        For item As Integer = 0 To arrColname.Count - 1
            Dim append As String = ""
            If (Not item = arrColname.Count - 1) Then
                append = $"{arrColname(item)} = '{arrContent(item)}',"
            Else
                append = $"{arrColname(item)} = '{arrContent(item)}'"
            End If

            condition = condition + append
        Next

        ' Update Builder
        Dim statement As String = $"UPDATE {table} SET {condition} WHERE ID = {selectedRow}"
        'MessageBox.Show(statement)

        Return statement
    End Function

    Private Function deleteStatement() As String
        Dim delst As String = ""
        Dim dtable As New DataTable
        productORemployee()

        Try
            Select Case optable
                Case "product"
                    delst = $"
                        DELETE FROM product_has_supplier WHERE product_id = {dgridSelectedRow};
                        DELETE FROM product WHERE id = {dgridSelectedRow};
                        "
                Case "user"
                    delst = $"
                        DELETE FROM user_has_department WHERE user_id = {dgridSelectedRow};
                        DELETE FROM user_has_salary WHERE user_id = {dgridSelectedRow};
                        DELETE FROM user_has_schedule WHERE user_id = {dgridSelectedRow};
                        DELETE FROM login WHERE id = {dgridSelectedRow};
                        DELETE FROM user WHERE id = {dgridSelectedRow};
                        "
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Return delst
    End Function

    Private Function productORemployee() As String
        Dim lblttl As String
        If lblTitle.Text = "PRODUCTS INVENTORY" Then
            lblttl = "PRODUCTS INVENTORY"
            optable = "product"
            stm = "SELECT * FROM product"
        Else
            lblttl = "EMPLOYEE MANAGEMENT"
            optable = "user"
            stm = "SELECT user.*, role.name as rolenm, department.name as dptnm FROM user" + employeeJoin
        End If

        Return lblttl
    End Function

    Public Sub resetform()
        dgrid.DataSource = Nothing
        dgrid.Rows.Clear()
        productORemployee()

        dgridPopulate(stm)
        viewFLP.Controls.Clear()
        Button4.Enabled = False
        Button2.Enabled = False
    End Sub





End Class