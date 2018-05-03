<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSwitchWarning
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSwitchWarning))
        Me.btnWarnOK = New System.Windows.Forms.Button()
        Me.lblWarnTtitle = New System.Windows.Forms.Label()
        Me.lblWarnAlert = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnWarnOK
        '
        Me.btnWarnOK.Location = New System.Drawing.Point(506, 292)
        Me.btnWarnOK.Name = "btnWarnOK"
        Me.btnWarnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnWarnOK.TabIndex = 0
        Me.btnWarnOK.Text = "OK"
        Me.btnWarnOK.UseVisualStyleBackColor = True
        '
        'lblWarnTtitle
        '
        Me.lblWarnTtitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarnTtitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblWarnTtitle.Location = New System.Drawing.Point(20, 25)
        Me.lblWarnTtitle.Name = "lblWarnTtitle"
        Me.lblWarnTtitle.Size = New System.Drawing.Size(555, 23)
        Me.lblWarnTtitle.TabIndex = 1
        Me.lblWarnTtitle.Text = "WARNING"
        Me.lblWarnTtitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblWarnAlert
        '
        Me.lblWarnAlert.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarnAlert.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblWarnAlert.Location = New System.Drawing.Point(13, 73)
        Me.lblWarnAlert.Name = "lblWarnAlert"
        Me.lblWarnAlert.Size = New System.Drawing.Size(562, 193)
        Me.lblWarnAlert.TabIndex = 3
        Me.lblWarnAlert.Text = resources.GetString("lblWarnAlert.Text")
        Me.lblWarnAlert.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmSwitchWarning
        '
        Me.AcceptButton = Me.btnWarnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(593, 327)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblWarnAlert)
        Me.Controls.Add(Me.lblWarnTtitle)
        Me.Controls.Add(Me.btnWarnOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSwitchWarning"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Conform Safety Warning"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnWarnOK As System.Windows.Forms.Button
    Friend WithEvents lblWarnTtitle As System.Windows.Forms.Label
    Friend WithEvents lblWarnAlert As System.Windows.Forms.Label
End Class
