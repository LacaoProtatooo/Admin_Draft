
Imports System.Web.UI.WebControls

Public Class AdmnComponentInfo

    Public Sub generataInfoPanel(info As String, content As String)
        Dim Panel2 As New System.Windows.Forms.Panel()
        Dim Label2 As New System.Windows.Forms.Label()
        Dim Label1 As New System.Windows.Forms.Label()
        Panel2.SuspendLayout()
        InfoFLP.Controls.Add(Panel2)

        ' Info
        Label2.Name = "Label2"
        If (content.Length > 28 AndAlso content.Length < 60) Then
            Panel2.Size = New System.Drawing.Size(384, 78)

            Dim cutStr As String = content.Substring(0, 28)
            Dim cutStrEx As String = content.Substring(28)
            Label2.Text = cutStr

            Dim extendlabel As New System.Windows.Forms.Label()
            extendlabel.AutoSize = True
            extendlabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            extendlabel.Location = New System.Drawing.Point(130, 42)
            extendlabel.Size = New System.Drawing.Size(34, 19)
            extendlabel.TabIndex = 0

            extendlabel.Text = cutStrEx
            Panel2.Controls.Add(extendlabel)

        ElseIf (content.Length > 60) Then
            Panel2.Size = New System.Drawing.Size(384, 90)

            Dim cutStr As String = content.Substring(0, 28)
            Dim cutStrEx As String = content.Substring(28, 32)
            Dim cutStrExEx As String = content.Substring(60)
            Label2.Text = cutStr

            Dim extendlabel As New System.Windows.Forms.Label()
            extendlabel.AutoSize = True
            extendlabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            extendlabel.Location = New System.Drawing.Point(130, 39)
            extendlabel.Size = New System.Drawing.Size(34, 19)
            extendlabel.TabIndex = 0

            Dim eextendlabel As New System.Windows.Forms.Label()
            eextendlabel.AutoSize = True
            eextendlabel.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            eextendlabel.Location = New System.Drawing.Point(130, 62)
            eextendlabel.Size = New System.Drawing.Size(34, 19)
            eextendlabel.TabIndex = 0

            extendlabel.Text = cutStrEx
            eextendlabel.Text = cutStrExEx
            Panel2.Controls.Add(extendlabel)
            Panel2.Controls.Add(eextendlabel)
        Else

            Panel2.Size = New System.Drawing.Size(384, 39)
            Label2.Text = content
        End If

        ' Label
        Label1.Name = "Label1"
        Label1.Text = info
        '
        ' Panel
        '
        Panel2.Controls.Add(Label2)
        Panel2.Controls.Add(Label1)
        Panel2.Location = New System.Drawing.Point(25, 30)
        Panel2.Name = "Panel2"
        Panel2.TabIndex = 0
        Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        '
        'Label2 - Label Content
        '
        Label2.AutoSize = True
        Label2.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label2.Location = New System.Drawing.Point(130, 11)
        Label2.Size = New System.Drawing.Size(58, 19)
        Label2.TabIndex = 1
        '
        'Label1 - Label Info
        '
        Label1.AutoSize = True
        Label1.Font = New System.Drawing.Font("Lexend", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label1.Location = New System.Drawing.Point(12, 11)
        Label1.Size = New System.Drawing.Size(34, 19)
        Label1.TabIndex = 0

    End Sub


End Class