Imports System.Diagnostics.Eventing.Reader
Imports System.Web.UI.WebControls
Imports CrystalDecisions.[Shared]
Imports CrystalDecisions.Windows.Forms

Public Class AdmnInsert
    Dim prodOrEmploy, cmbRole, cmbDepartment As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Save
        Select Case prodOrEmploy
            Case "product"
                productAdd(prodOrEmploy)

            Case "user"
                employeeAdd(prodOrEmploy)
        End Select

    End Sub

    Public Sub productAdd(table As String)
        Dim itemname, price, stock_quantity, description, brand, category, imagedir As String
        itemname = TextBox1.Text
        price = TextBox2.Text
        stock_quantity = TextBox3.Text
        description = TextBox4.Text
        brand = TextBox5.Text
        category = TextBox6.Text
        imagedir = TextBox7.Text

        Dim arrInsert As New ArrayList From {
            itemname,
            price,
            stock_quantity,
            description,
            brand,
            category,
            imagedir
        }

        Dim arrNamelist As New ArrayList From {
            Label1.Text.ToString(),
            Label2.Text.ToString(),
            Label3.Text.ToString(),
            Label4.Text.ToString(),
            Label5.Text.ToString(),
            Label6.Text.ToString(),
            Label7.Text.ToString()
        }

        Try
            Dim statement As String = insertBuild(table, arrInsert, arrNamelist)
            Dim dtable As DataTable = exquery(statement)
            AdmnWholeList.dgrid.DataSource = dtable
            AdmnWholeList.resetform()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub employeeAdd(table As String)
        Dim username, first_name, last_name, email, role, department As String
        username = TextBox1.Text
        first_name = TextBox2.Text
        last_name = TextBox3.Text
        email = TextBox4.Text
        role = TextBox5.Text
        department = TextBox6.Text

        Dim arrInsert As New ArrayList From {
            username,
            first_name,
            last_name,
            email,
            role,
            department
        }

        Dim arrNamelist As New ArrayList From {
            Label1.Text.ToString(),
            Label2.Text.ToString(),
            Label3.Text.ToString(),
            Label4.Text.ToString(),
            Label5.Text.ToString(),
            Label6.Text.ToString()
        }

        Try
            Dim statement As String = insertBuild(table, arrInsert, arrNamelist)
            Dim dtable As DataTable = exquery(statement)
            AdmnWholeList.dgrid.DataSource = dtable
            AdmnWholeList.resetform()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Function insertBuild(table As String, arrList As ArrayList, arrNamelist As ArrayList) As String
        Dim statement As String = ""
        Dim tblRole = "login", tblDeparment As String = "user_has_department" 'Table
        Dim userCol = "", userVal = "", prodCol = "", prodVal As String = ""
        Dim uscolbuild, prodcolbuild As New ArrayList
        Dim usbuild, prodbuild As New ArrayList
        Dim dpStm = "", userStm = "", usnm = "", rolStm As String = ""

        Select Case prodOrEmploy
            Case "product" 'Product
                For item As Integer = 0 To arrNamelist.Count - 1
                    prodcolbuild.Add(arrNamelist(item))
                    prodbuild.Add(arrList(item))
                Next
                prodCol = insertStmColBuild(prodcolbuild)
                prodVal = insertStmValBuild(prodbuild)

                statement = $"INSERT INTO product({prodCol}) VALUES({prodVal});"
                MessageBox.Show(statement)

            Case "user" ' Employee
                For item As Integer = 0 To arrNamelist.Count - 1
                    Select Case arrNamelist(item) 'Tables
                        Case "department"
                            dpStm = newInsertStatement(tblDeparment, usnm, cmbDepartment)
                        Case "role"
                            rolStm = newInsertStatement(tblRole, usnm, cmbRole)
                        Case Else
                            uscolbuild.Add(arrNamelist(item))
                            usbuild.Add(arrList(item))

                            If (arrNamelist(item) = "username") Then
                                usnm = arrList(item)
                            End If

                    End Select
                Next

                userCol = insertStmColBuild(uscolbuild)
                userVal = insertStmValBuild(usbuild)
                userStm = $"INSERT INTO user({userCol}) VALUES({userVal});"

                statement = userStm + rolStm + dpStm
                MessageBox.Show(statement)

        End Select
        Return statement
    End Function

    ' Insert ColNames Statment Build
    Public Function insertStmColBuild(collectedColName As ArrayList) As String
        Dim Statement As String = ""
        For item As Integer = 0 To collectedColName.Count - 1
            If Not (item = collectedColName.Count - 1) Then
                Statement = Statement + $"{collectedColName(item)},"
            Else
                Statement = Statement + $"{collectedColName(item)}"
            End If
        Next

        Return Statement
    End Function

    ' Insert Values Statement Build
    Public Function insertStmValBuild(collectedValues As ArrayList) As String
        Dim statment As String = ""
        For item As Integer = 0 To collectedValues.Count - 1
            If Not (item = collectedValues.Count - 1) Then
                If IsNumeric(collectedValues(item)) Then
                    statment = statment + $"{collectedValues(item)},"
                Else
                    statment = statment + $"'{collectedValues(item)}',"
                End If
            Else
                If IsNumeric(collectedValues(item)) Then
                    statment = statment + $"{collectedValues(item)}"
                Else
                    statment = statment + $"'{collectedValues(item)}'"
                End If
            End If
        Next

        Return statment
    End Function

    ' For Department Table and Role Table Build
    Public Function newInsertStatement(difftable As String, username As String, content As String) As String
        Dim statement As String = ""
        Dim pkey As Integer
        Dim index As String = ""

        Select Case content
            Case "Sales Admin" ' Roles
                index = 1
            Case "Maintenance Admin"
                index = 2
            Case "Service Admin"
                index = 3
            Case "Employee"
                index = 4
            Case "Sales" ' Employee Department
                index = 1
            Case "Service"
                index = 2
            Case "Maintenance"
                index = 3
        End Select

        Select Case difftable
            Case "login" ' ROLE
                pkey = getHighestPK() + 1
                statement = $"INSERT INTO login(id, username, password, role_id) VALUES({pkey.ToString},'{username}','mypassword',{index});"

            Case "user_has_department" ' DEPARTMENT
                pkey = getHighestPK() + 1
                statement = $"INSERT INTO user_has_department(user_id, department_id) VALUES({pkey.ToString},{index});"

        End Select

        Return statement
    End Function

    'Get Highest Primary Key
    Private Function getHighestPK() As Integer
        Dim pkStatement As String = $"SELECT MAX(id) FROM user;"
        Dim result As Integer = exqueyPK(pkStatement)
        Return result
    End Function

    Public Sub prodSetup(ByRef table As String)
        Label1.Text = "item_name"
        Label2.Text = "price"
        Label3.Text = "stock_quantity"
        Label4.Text = "description"
        Label5.Text = "brand"
        Label6.Text = "category"
        Label7.Text = "image_dir"
        Label8.Text = ""
        Label9.Text = ""
        Label10.Text = ""

        prodOrEmploy = table
    End Sub

    Private Sub cmbxRole_SelectIndexChanged(sender As Object, e As EventArgs)
        Dim cmbxRole As Windows.Forms.ComboBox
        cmbxRole = CType(sender, Object)
        cmbRole = cmbxRole.SelectedItem

    End Sub

    Private Sub cmbxDepartment_SelectIndexChanged(sender As Object, e As EventArgs)
        Dim cmbxDepartment As Windows.Forms.ComboBox
        cmbxDepartment = CType(sender, Object)
        cmbDepartment = cmbxDepartment.SelectedItem

    End Sub

    Public Sub empSetup(ByRef table As String)
        Label1.Text = "username"
        Label2.Text = "first_name"
        Label3.Text = "last_name"
        Label4.Text = "email"
        Label5.Text = "role"
        Label6.Text = "department"
        Label7.Text = ""
        Label8.Text = ""
        Label9.Text = ""
        Label10.Text = ""

        Dim cmbxRole As New Windows.Forms.ComboBox()
        Dim cmbxDepartment As New Windows.Forms.ComboBox()

        cmbxRole.Location = New System.Drawing.Point(129, 10)
        cmbxRole.Name = "TextBox5"
        cmbxRole.Size = New System.Drawing.Size(313, 20)
        cmbxRole.TabIndex = 1
        cmbxRole.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

        cmbxDepartment.Location = New System.Drawing.Point(129, 10)
        cmbxDepartment.Name = "TextBox6"
        cmbxDepartment.Size = New System.Drawing.Size(313, 20)
        cmbxDepartment.TabIndex = 1
        cmbxDepartment.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

        Panel7.Controls.Add(cmbxRole)
        Panel8.Controls.Add(cmbxDepartment)

        Panel7.Controls.Remove(TextBox5)
        Panel8.Controls.Remove(TextBox6)

        cmbxRole.DropDownStyle = ComboBoxStyle.DropDownList
        cmbxDepartment.DropDownStyle = ComboBoxStyle.DropDownList

        cmbxRole.Items.Add("Sales Admin")
        cmbxRole.Items.Add("Maintenance Admin")
        cmbxRole.Items.Add("Service Admin")
        cmbxRole.Items.Add("Employee")

        cmbxDepartment.Items.Add("Sales")
        cmbxDepartment.Items.Add("Service")
        cmbxDepartment.Items.Add("Maintenance")

        AddHandler cmbxRole.SelectedIndexChanged, AddressOf cmbxRole_SelectIndexChanged
        AddHandler cmbxDepartment.SelectedIndexChanged, AddressOf cmbxDepartment_SelectIndexChanged

        prodOrEmploy = table
    End Sub
End Class