Imports System.Net.Http.Headers
Imports System.Security.AccessControl
Imports System.Security.Cryptography.X509Certificates
Imports System.Web.UI.WebControls
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.[Shared]
Imports Microsoft.VisualBasic.ApplicationServices

Public Class AdminMain

    Private Sub AdminMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Done
        Try
            productLoad()
            employeeLoad()
            dtpcurrent()
            CRload()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    ' Text Box Filter
    Private Sub txtbxprodlist_TextChanged(sender As Object, e As EventArgs) Handles txtbxprodlist.TextChanged
        Dim prdSearch As String = txtbxprodlist.Text

        'Configure This
        Dim selectedColumns As String = "id, item_name, category"
        Dim table As String = "product"
        Dim stmntCondition As String = "id LIKE '%" + prdSearch + "%' OR item_name LIKE '%" + prdSearch + "%' OR category LIKE '%" + prdSearch + "%'"

        'General Use
        Dim dtable As DataTable = contentSearcher(selectedColumns, table, stmntCondition)

        ' Exclusive only for Product
        dispProdSearch(dtable)
    End Sub

    Private Sub txtbxemploylist_TextChanged(sender As Object, e As EventArgs) Handles txtbxemploylist.TextChanged
        Dim empSearch As String = txtbxemploylist.Text

        Dim selectedColumns As String = "user.id, user.username, role.name AS rolenm, department.name AS dptnm"
        Dim table As String = "user " + employeeJoin
        Dim stmntCondition As String = "user.id LIKE '%" + empSearch + "%' OR user.username LIKE '%" + empSearch + "%' OR role.name LIKE '%" + empSearch + "%' OR department.name LIKE '%" + empSearch + "%'"

        Dim dtable As DataTable = contentSearcher(selectedColumns, table, stmntCondition)

        dispEmpSearch(dtable)
    End Sub

    ' Combo Boxes
    Private Sub cmbxprodlist_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbxprodlist.SelectedIndexChanged
        Dim slctItem As String = cmbxprodlist.SelectedItem.ToString

        Dim selectedColumns As String = "id, item_name, category"
        Dim table As String = "product"
        Dim stmntGrpBy As String = "" + slctItem + ""

        Dim dtable As DataTable = groupBySeacrh(selectedColumns, table, stmntGrpBy)

        dispProdSearch(dtable)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Open Whole Products List
        AdmnWholeList.Show()
        AdmnWholeList.lblTitle.Text = "PRODUCTS INVENTORY"
        Me.Hide()

        Dim statement As String = "Select * from product"
        AdmnWholeList.dgridPopulate(statement)


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Open Whole Employees List
        AdmnWholeList.Show()
        AdmnWholeList.lblTitle.Text = "EMPLOYEE MANAGEMENT"
        Me.Hide()

        Dim statement As String = "SELECT user.*, role.name as rolenm, department.name as dptnm FROM user " + employeeJoin
        AdmnWholeList.dgridPopulate(statement)

    End Sub

    Private Sub cmbxemploylist_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbxemploylist.SelectedIndexChanged
        Dim slctItem As String = cmbxemploylist.SelectedItem.ToString
        If (slctItem = "role") Then
            slctItem = "rolenm"
        End If
        If (slctItem = "department") Then
            slctItem = "dptnm"
        End If

        Dim selectedColumns As String = "user.id, user.username, role.name AS rolenm, department.name AS dptnm"
        Dim table As String = "user " + employeeJoin
        Dim stmntGrpBy As String = "" + slctItem + ""

        Dim dtable As DataTable = groupBySeacrh(selectedColumns, table, stmntGrpBy)

        dispEmpSearch(dtable)
    End Sub

    '====================== FLOWLAYOUTPANEL CONTROLS =======================

    Public Sub createProdPanel(prdID As Integer, prdName As String, prdCategory As String)

        ' Create Panel and its Contents
        Dim prodpanel As New System.Windows.Forms.Panel()
        Dim prodinfo As New System.Windows.Forms.Button()
        Dim prodCategorylabel As New System.Windows.Forms.Label()
        Dim prodNamelabel As New System.Windows.Forms.Label()
        Dim prodIDlabel As New System.Windows.Forms.Label()
        prodpanel.SuspendLayout()
        ProductListFLP.Controls.Add(prodpanel)

        prodpanel.Name = prdID.ToString

        prodCategorylabel.Name = "prodCategorylabel"
        prodCategorylabel.Text = prdCategory.ToString

        prodNamelabel.Name = "prodNamelabel"
        prodNamelabel.Text = prdName.ToString

        prodIDlabel.Name = "prodIDlabel"
        prodIDlabel.Text = prdID.ToString

        ' New Panel Properties
        prodpanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        prodpanel.Controls.Add(prodCategorylabel)
        prodpanel.Controls.Add(prodNamelabel)
        prodpanel.Controls.Add(prodIDlabel)
        prodpanel.Controls.Add(prodinfo)
        prodpanel.Location = New System.Drawing.Point(6, 6)

        prodpanel.Padding = New System.Windows.Forms.Padding(2)
        prodpanel.Size = New System.Drawing.Size(562, 32)
        prodpanel.TabIndex = 0
        '
        ' Product Panels Label
        '
        'prodCategorylabel
        '
        prodCategorylabel.AutoSize = True
        prodCategorylabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        prodCategorylabel.Location = New System.Drawing.Point(350, 6)
        prodCategorylabel.Size = New System.Drawing.Size(64, 19)
        prodCategorylabel.TabIndex = 3
        '
        'prodNamelabel
        '
        prodNamelabel.AutoSize = True
        prodNamelabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        prodNamelabel.Location = New System.Drawing.Point(101, 6)
        prodNamelabel.Size = New System.Drawing.Size(91, 19)
        prodNamelabel.TabIndex = 2
        '
        'prodIDlabel
        '
        '
        prodIDlabel.AutoSize = True
        prodIDlabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        prodIDlabel.Location = New System.Drawing.Point(5, 5)
        prodIDlabel.Size = New System.Drawing.Size(70, 19)
        prodIDlabel.TabIndex = 1
        '
        ' Prod Info Button
        '
        prodinfo.Location = New System.Drawing.Point(530, 5)
        prodinfo.Size = New System.Drawing.Size(25, 20)
        prodinfo.TabIndex = 0
        prodinfo.Text = "..."
        prodinfo.UseVisualStyleBackColor = True

        ' Button
        AddHandler prodinfo.Click, AddressOf prodinfo_Click
        prodinfo.Name = (prdID)

    End Sub

    Public Sub createEmployPanel(empID As Integer, empName As String, empRole As String, empDepartment As String)

        ' Create Panel and its Contents
        Dim emppanel As New System.Windows.Forms.Panel()
        Dim empinfo As New System.Windows.Forms.Button()
        Dim empDepartmentlabel As New System.Windows.Forms.Label()
        Dim empNamelabel As New System.Windows.Forms.Label()
        Dim empIDlabel As New System.Windows.Forms.Label()
        Dim empRolelabel As New System.Windows.Forms.Label()
        emppanel.SuspendLayout()
        EmployeeListFLP.Controls.Add(emppanel)

        emppanel.Name = "dummyname"

        empDepartmentlabel.Name = "empDepartmentlabel"
        empDepartmentlabel.Text = empDepartment

        empNamelabel.Name = "empNamelabel"
        empNamelabel.Text = empName

        empIDlabel.Name = "empIDlabel"
        empIDlabel.Text = empID.ToString

        empRolelabel.Name = "empRolelabel"
        empRolelabel.Text = empRole

        ' New Panel Properties
        emppanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        emppanel.Controls.Add(empDepartmentlabel)
        emppanel.Controls.Add(empNamelabel)
        emppanel.Controls.Add(empIDlabel)
        emppanel.Controls.Add(empRolelabel)
        emppanel.Controls.Add(empinfo)
        emppanel.Location = New System.Drawing.Point(6, 6)
        emppanel.Padding = New System.Windows.Forms.Padding(2)
        emppanel.Size = New System.Drawing.Size(562, 32)
        emppanel.TabIndex = 0
        '
        ' Product Panels Label
        '
        'empDepartmentlabel
        '
        empDepartmentlabel.AutoSize = True
        empDepartmentlabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        empDepartmentlabel.Location = New System.Drawing.Point(370, 6)
        empDepartmentlabel.Size = New System.Drawing.Size(64, 19)
        empDepartmentlabel.TabIndex = 3
        '
        'empNamelabel
        '
        empNamelabel.AutoSize = True
        empNamelabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        empNamelabel.Location = New System.Drawing.Point(101, 6)
        empNamelabel.Size = New System.Drawing.Size(91, 19)
        empNamelabel.TabIndex = 2
        '
        'empRole
        '
        empRolelabel.AutoSize = True
        empRolelabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        empRolelabel.Location = New System.Drawing.Point(242, 6)
        empRolelabel.Size = New System.Drawing.Size(91, 19)
        empRolelabel.TabIndex = 2
        '
        'empIDlabel
        '
        '
        empIDlabel.AutoSize = True
        empIDlabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        empIDlabel.Location = New System.Drawing.Point(5, 6)
        empIDlabel.Size = New System.Drawing.Size(70, 19)
        empIDlabel.TabIndex = 1
        '
        ' Prod Info Button
        '
        empinfo.Location = New System.Drawing.Point(530, 5)
        empinfo.Size = New System.Drawing.Size(25, 20)
        empinfo.TabIndex = 0
        empinfo.Text = "..."
        empinfo.UseVisualStyleBackColor = True
        '
        ' Button
        '
        AddHandler empinfo.Click, AddressOf empinfo_Click
        empinfo.Name = empID

    End Sub

    ' Button Event Handler inside Product List Flow Layout Panel
    Private Sub prodinfo_Click(sender As Object, e As EventArgs)
        Dim boton As Windows.Forms.Button
        boton = CType(sender, Object)

        Dim idname As String = boton.Name
        Dim selectedColumns As String = "*"
        Dim table As String = "product"
        Dim stmntCondition As String = $"id = {idname}"

        Dim dtable As DataTable = contentSearcher(selectedColumns, table, stmntCondition)

        AdmnComponentInfo.Close()
        compInfo(dtable)
    End Sub


    ' Button Event Handler inside Employee List Flow Layout Panel
    Private Sub empinfo_Click(sender As Object, e As EventArgs)
        Dim boton As Windows.Forms.Button
        boton = CType(sender, Object)

        Dim idname As String = boton.Name
        Dim selectedColumns As String = "*"
        Dim table As String = "user" + employeeJoin
        Dim stmntCondition As String = $"user.id = {idname}"

        Dim dtable As DataTable = contentSearcher(selectedColumns, table, stmntCondition)

        AdmnComponentInfo.Close()
        compInfo(dtable)
    End Sub


    '====================== FORM LOAD =======================

    ' Product List Load (Not For General use)
    Private Sub productLoad()
        ProductListFLP.Controls.Clear()
        Dim sqlstmnt As String = "Select id, item_name, category from product"

        Dim rowLength, colLength As Integer
        Dim arrList As ArrayList = populateFunction(sqlstmnt, rowLength, colLength)
        Dim jump As Integer = 0

        For row = 0 To rowLength - 1
            ' 3 Columns - id, name, category
            ' Configure this per table

            jump = row * colLength
            createProdPanel(arrList(0 + jump), arrList(1 + jump), arrList(2 + jump))
        Next
    End Sub

    Private Sub employeeLoad()
        EmployeeListFLP.Controls.Clear()
        Dim sqlstmnt As String =
        $"SELECT user.id, user.username, role.name AS rolenm, department.name AS dptnm FROM user {employeeJoin} ;"

        Dim rowLength, colLength As Integer
        Dim arrList As ArrayList = populateFunction(sqlstmnt, rowLength, colLength)
        Dim jump As Integer = 0

        For row = 0 To rowLength - 1
            ' 4 Columns - id, name, rolenm, dptnm
            ' Configure this per table

            jump = row * colLength
            createEmployPanel(arrList(0 + jump), arrList(1 + jump), arrList(2 + jump), arrList(3 + jump))
        Next

    End Sub

    ' Date Time Component
    Private Sub dtpcurrent()
        dtp1.Value = DateTime.Now
    End Sub

    '====================== FORM LOAD =======================

    '====================== OTHERS =======================

    ' Crystal Report - WIP
    Public strReportPath As String
    Private Sub verifyCr()
        ' This is Your Crystal Report Directory
        strReportPath = "C:\Users\vince\Source\Repos\Admin_Draft\Admin_Draft_FDS\ProductReport.rpt"

        If Not IO.File.Exists(strReportPath) Then
            MessageBox.Show("Unable to Locate File. " & vbCrLf & strReportPath)
        End If
    End Sub


    Private Sub CRload()
        Try
            connODBC.Open()
            verifyCr()
            populateCRYSTAL()

            ' Load the Crystal Report .rpt File and pass it onto Datatable
            Dim cr As New ReportDocument

            cr.Load(strReportPath)
            cr.SetDataSource(ds)

            ' Set the CrystalReportViewer's Appearance and Set the ReportSource:
            CrystalReportViewer1.ShowRefreshButton = False
            CrystalReportViewer1.ShowCloseButton = False
            CrystalReportViewer1.ShowGroupTreeButton = False

            CrystalReportViewer1.ReportSource = cr

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            connODBC.Close()
        End Try
    End Sub

    '====================== OTHERS =======================


End Class