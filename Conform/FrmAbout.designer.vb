<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAbout
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAbout))
        Me.lblAbout1 = New System.Windows.Forms.Label
        Me.lblAbout2 = New System.Windows.Forms.Label
        Me.lblAbout3 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblAbout1
        '
        Me.lblAbout1.AutoSize = True
        Me.lblAbout1.Location = New System.Drawing.Point(51, 33)
        Me.lblAbout1.Name = "lblAbout1"
        Me.lblAbout1.Size = New System.Drawing.Size(154, 13)
        Me.lblAbout1.TabIndex = 0
        Me.lblAbout1.Text = "ASCOM Conformance Checker"
        '
        'lblAbout2
        '
        Me.lblAbout2.AutoSize = True
        Me.lblAbout2.Location = New System.Drawing.Point(51, 46)
        Me.lblAbout2.Name = "lblAbout2"
        Me.lblAbout2.Size = New System.Drawing.Size(102, 13)
        Me.lblAbout2.TabIndex = 1
        Me.lblAbout2.Text = "Peter Simpson 2007"
        '
        'lblAbout3
        '
        Me.lblAbout3.AutoSize = True
        Me.lblAbout3.Location = New System.Drawing.Point(54, 63)
        Me.lblAbout3.Name = "lblAbout3"
        Me.lblAbout3.Size = New System.Drawing.Size(0, 13)
        Me.lblAbout3.TabIndex = 2
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Conform5.My.Resources.Resources.ASCOM
        Me.PictureBox1.InitialImage = Global.Conform5.My.Resources.Resources.ASCOM
        Me.PictureBox1.Location = New System.Drawing.Point(101, 96)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(52, 59)
        Me.PictureBox1.TabIndex = 3
        Me.PictureBox1.TabStop = False
        '
        'frmAbout
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblAbout3)
        Me.Controls.Add(Me.lblAbout2)
        Me.Controls.Add(Me.lblAbout1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAbout"
        Me.Text = "frmAbout"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblAbout1 As System.Windows.Forms.Label
    Friend WithEvents lblAbout2 As System.Windows.Forms.Label
    Friend WithEvents lblAbout3 As System.Windows.Forms.Label

    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblAbout3.Text = Application.ProductVersion
    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
