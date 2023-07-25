<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AdmnComponentInfo
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.InfoFLP = New System.Windows.Forms.FlowLayoutPanel()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.InfoFLP)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Padding = New System.Windows.Forms.Padding(10)
        Me.Panel1.Size = New System.Drawing.Size(461, 361)
        Me.Panel1.TabIndex = 0
        '
        'InfoFLP
        '
        Me.InfoFLP.AutoScroll = True
        Me.InfoFLP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.InfoFLP.Location = New System.Drawing.Point(10, 10)
        Me.InfoFLP.Name = "InfoFLP"
        Me.InfoFLP.Padding = New System.Windows.Forms.Padding(12)
        Me.InfoFLP.Size = New System.Drawing.Size(439, 339)
        Me.InfoFLP.TabIndex = 1
        '
        'AdmnComponentInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(461, 361)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "AdmnComponentInfo"
        Me.Text = "Info"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents InfoFLP As FlowLayoutPanel
End Class
