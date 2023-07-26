Public Class AdmnInsert
    Dim prodOrEmploy As String
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
        Dim username, first_name, last_name, email, departmentname, rolename As String
        username = TextBox1.Text
        first_name = TextBox2.Text
        last_name = TextBox3.Text
        email = TextBox4.Text
        departmentname = TextBox5.Text
        rolename = TextBox6.Text

        Dim arrInsert As New ArrayList From {
            username,
            first_name,
            last_name,
            email
        }

        Dim arrNamelist As New ArrayList From {
            Label1.Text.ToString(),
            Label2.Text.ToString(),
            Label3.Text.ToString(),
            Label4.Text.ToString()
        }

        Try
            Dim statement As String = insertBuild(table, arrInsert, arrNamelist)
            'Dim dtable As DataTable = exquery(statement)
            'AdmnWholeList.dgrid.DataSource = dtable
            'AdmnWholeList.resetform()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Function insertBuild(table As String, arrList As ArrayList, arrNamelist As ArrayList) As String
        Dim columns = "", values As String = ""

        For item As Integer = 0 To arrNamelist.Count - 1
            If (Not item = arrNamelist.Count - 1) Then
                columns = columns + $"{arrNamelist(item)},"
                If IsNumeric(arrList(item)) Then
                    values = values + $"{arrList(item)},"
                Else
                    values = values + $"'{arrList(item)}',"
                End If

            Else
                Dim tblprefix As String
                Select Case arrNamelist(item)
                    Case "department"
                        tblprefix = "department.name"
                        columns = columns + $"{tblprefix}"
                    Case "role"
                        tblprefix = "role."
                        columns = columns + $"{tblprefix}"
                    Case Else
                        tblprefix = "user."
                        columns = columns + $"{tblprefix + arrNamelist(item)}"
                End Select

                If IsNumeric(arrList(item)) Then
                    values = values + $"{arrList(item)}"
                Else
                    values = values + $"'{arrList(item)}'"
                End If

            End If
        Next

        Dim statement As String = $"INSERT INTO {table}({columns}) VALUES({values})"
        MessageBox.Show(statement)
        Return statement
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

    Public Sub empSetup(ByRef table As String)
        Label1.Text = "username"
        Label2.Text = "first_name"
        Label3.Text = "last_name"
        Label4.Text = "email"
        Label5.Text = ""
        Label6.Text = ""
        Label7.Text = ""
        Label8.Text = ""
        Label9.Text = ""
        Label10.Text = ""

        prodOrEmploy = table
    End Sub
End Class