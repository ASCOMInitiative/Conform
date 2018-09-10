<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWarning
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWarning))
        Me.btnWarnOK = New System.Windows.Forms.Button()
        Me.lblWarnTtitle = New System.Windows.Forms.Label()
        Me.lblWarnBody1 = New System.Windows.Forms.Label()
        Me.lblWarnAlert = New System.Windows.Forms.Label()
        Me.chkWarnAgree = New System.Windows.Forms.CheckBox()
        Me.btnWarnExit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnWarnOK
        '
        Me.btnWarnOK.Location = New System.Drawing.Point(506, 328)
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
        'lblWarnBody1
        '
        Me.lblWarnBody1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarnBody1.Location = New System.Drawing.Point(17, 236)
        Me.lblWarnBody1.Name = "lblWarnBody1"
        Me.lblWarnBody1.Size = New System.Drawing.Size(558, 60)
        Me.lblWarnBody1.TabIndex = 2
        Me.lblWarnBody1.Text = resources.GetString("lblWarnBody1.Text")
        Me.lblWarnBody1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblWarnAlert
        '
        Me.lblWarnAlert.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarnAlert.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblWarnAlert.Location = New System.Drawing.Point(13, 73)
        Me.lblWarnAlert.Name = "lblWarnAlert"
        Me.lblWarnAlert.Size = New System.Drawing.Size(562, 148)
        Me.lblWarnAlert.TabIndex = 3
        Me.lblWarnAlert.Text = resources.GetString("lblWarnAlert.Text")
        Me.lblWarnAlert.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'chkWarnAgree
        '
        Me.chkWarnAgree.AutoSize = True
        Me.chkWarnAgree.Location = New System.Drawing.Point(440, 332)
        Me.chkWarnAgree.Name = "chkWarnAgree"
        Me.chkWarnAgree.Size = New System.Drawing.Size(60, 17)
        Me.chkWarnAgree.TabIndex = 4
        Me.chkWarnAgree.Text = "I Agree"
        Me.chkWarnAgree.UseVisualStyleBackColor = True
        '
        'btnWarnExit
        '
        Me.btnWarnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnWarnExit.Location = New System.Drawing.Point(506, 299)
        Me.btnWarnExit.Name = "btnWarnExit"
        Me.btnWarnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnWarnExit.TabIndex = 5
        Me.btnWarnExit.Text = "Exit"
        Me.btnWarnExit.UseVisualStyleBackColor = True
        '
        'frmWarning
        '
        Me.AcceptButton = Me.btnWarnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnWarnExit
        Me.ClientSize = New System.Drawing.Size(593, 362)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnWarnExit)
        Me.Controls.Add(Me.chkWarnAgree)
        Me.Controls.Add(Me.lblWarnAlert)
        Me.Controls.Add(Me.lblWarnBody1)
        Me.Controls.Add(Me.lblWarnTtitle)
        Me.Controls.Add(Me.btnWarnOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmWarning"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Conform Safety Warning"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnWarnOK As System.Windows.Forms.Button
    Friend WithEvents lblWarnTtitle As System.Windows.Forms.Label
    Friend WithEvents lblWarnBody1 As System.Windows.Forms.Label
    Friend WithEvents lblWarnAlert As System.Windows.Forms.Label
    Friend WithEvents chkWarnAgree As System.Windows.Forms.CheckBox
    Friend WithEvents btnWarnExit As System.Windows.Forms.Button
End Class
