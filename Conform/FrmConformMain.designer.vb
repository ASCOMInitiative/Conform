<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmConformMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmConformMain))
        Me.BtnCheckConformance = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckConformanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SelectDriverToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DriverSetupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConformanceOptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuTestTelescope = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestDome = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestCamera = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestVideoCamera = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestFocuser = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestFilterWheel = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestObservingConditions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestRotator = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestSwitch = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTestSafetyMonitor = New System.Windows.Forms.ToolStripMenuItem()
        Me.ValidateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuConfirmValidationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCreateValidationFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SpecialTestsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTelescopeSideOfPierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTelescopeDestinationSideOfPierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTelescopeFlipRangeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTelescopeCommandsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.SideOfPierBehaviourPDFToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckForUpdatesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.BtnExit = New System.Windows.Forms.Button()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.staTest = New System.Windows.Forms.ToolStripStatusLabel()
        Me.staAction = New System.Windows.Forms.ToolStripStatusLabel()
        Me.staStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.BtnStop = New System.Windows.Forms.Button()
        Me.BtnClearAndCheckConformance = New System.Windows.Forms.Button()
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
        Me.chkDomeShutter = New System.Windows.Forms.CheckBox()
        Me.dlgValidation = New System.Windows.Forms.OpenFileDialog()
        Me.chkSwitchSet = New System.Windows.Forms.CheckBox()
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnCheckConformance
        '
        Me.BtnCheckConformance.Location = New System.Drawing.Point(842, 402)
        Me.BtnCheckConformance.Name = "BtnCheckConformance"
        Me.BtnCheckConformance.Size = New System.Drawing.Size(94, 37)
        Me.BtnCheckConformance.TabIndex = 1
        Me.BtnCheckConformance.Text = "Check Conformance"
        Me.BtnCheckConformance.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.OptionsToolStripMenuItem, Me.ValidateToolStripMenuItem, Me.SpecialTestsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(948, 24)
        Me.MenuStrip1.TabIndex = 2
        Me.MenuStrip1.Text = "MenuStrip"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CheckConformanceToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'CheckConformanceToolStripMenuItem
        '
        Me.CheckConformanceToolStripMenuItem.Name = "CheckConformanceToolStripMenuItem"
        Me.CheckConformanceToolStripMenuItem.Size = New System.Drawing.Size(183, 22)
        Me.CheckConformanceToolStripMenuItem.Text = "Check Conformance"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(183, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectDriverToolStripMenuItem, Me.DriverSetupToolStripMenuItem, Me.ConformanceOptionsToolStripMenuItem, Me.ToolStripSeparator1, Me.mnuTestTelescope, Me.mnuTestDome, Me.mnuTestCamera, Me.mnuTestVideoCamera, Me.mnuTestFocuser, Me.mnuTestFilterWheel, Me.mnuTestObservingConditions, Me.mnuTestRotator, Me.mnuTestSwitch, Me.mnuTestSafetyMonitor})
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.OptionsToolStripMenuItem.Text = "Options"
        '
        'SelectDriverToolStripMenuItem
        '
        Me.SelectDriverToolStripMenuItem.Name = "SelectDriverToolStripMenuItem"
        Me.SelectDriverToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
        Me.SelectDriverToolStripMenuItem.Text = "Select Driver"
        '
        'DriverSetupToolStripMenuItem
        '
        Me.DriverSetupToolStripMenuItem.Name = "DriverSetupToolStripMenuItem"
        Me.DriverSetupToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
        Me.DriverSetupToolStripMenuItem.Text = "Driver Setup"
        '
        'ConformanceOptionsToolStripMenuItem
        '
        Me.ConformanceOptionsToolStripMenuItem.Name = "ConformanceOptionsToolStripMenuItem"
        Me.ConformanceOptionsToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
        Me.ConformanceOptionsToolStripMenuItem.Text = "Conformance Options"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(219, 6)
        '
        'mnuTestTelescope
        '
        Me.mnuTestTelescope.Name = "mnuTestTelescope"
        Me.mnuTestTelescope.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestTelescope.Text = "Check Telescope"
        '
        'mnuTestDome
        '
        Me.mnuTestDome.Name = "mnuTestDome"
        Me.mnuTestDome.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestDome.Text = "Check Dome"
        '
        'mnuTestCamera
        '
        Me.mnuTestCamera.Name = "mnuTestCamera"
        Me.mnuTestCamera.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestCamera.Text = "Check Camera"
        '
        'mnuTestVideoCamera
        '
        Me.mnuTestVideoCamera.Name = "mnuTestVideoCamera"
        Me.mnuTestVideoCamera.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestVideoCamera.Text = "Check Video Camera"
        '
        'mnuTestFocuser
        '
        Me.mnuTestFocuser.Name = "mnuTestFocuser"
        Me.mnuTestFocuser.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestFocuser.Text = "Check Focuser"
        '
        'mnuTestFilterWheel
        '
        Me.mnuTestFilterWheel.Name = "mnuTestFilterWheel"
        Me.mnuTestFilterWheel.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestFilterWheel.Text = "Check Filter Wheel"
        '
        'mnuTestObservingConditions
        '
        Me.mnuTestObservingConditions.Name = "mnuTestObservingConditions"
        Me.mnuTestObservingConditions.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestObservingConditions.Text = "Check ObservingConditions"
        '
        'mnuTestRotator
        '
        Me.mnuTestRotator.Name = "mnuTestRotator"
        Me.mnuTestRotator.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestRotator.Text = "Check Rotator"
        '
        'mnuTestSwitch
        '
        Me.mnuTestSwitch.Name = "mnuTestSwitch"
        Me.mnuTestSwitch.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestSwitch.Text = "Check Switch"
        '
        'mnuTestSafetyMonitor
        '
        Me.mnuTestSafetyMonitor.Name = "mnuTestSafetyMonitor"
        Me.mnuTestSafetyMonitor.Size = New System.Drawing.Size(222, 22)
        Me.mnuTestSafetyMonitor.Text = "Check Safety Monitor"
        '
        'ValidateToolStripMenuItem
        '
        Me.ValidateToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuConfirmValidationToolStripMenuItem, Me.mnuCreateValidationFileToolStripMenuItem})
        Me.ValidateToolStripMenuItem.Name = "ValidateToolStripMenuItem"
        Me.ValidateToolStripMenuItem.Size = New System.Drawing.Size(60, 20)
        Me.ValidateToolStripMenuItem.Text = "Validate"
        '
        'mnuConfirmValidationToolStripMenuItem
        '
        Me.mnuConfirmValidationToolStripMenuItem.Name = "mnuConfirmValidationToolStripMenuItem"
        Me.mnuConfirmValidationToolStripMenuItem.Size = New System.Drawing.Size(250, 22)
        Me.mnuConfirmValidationToolStripMenuItem.Text = "Confirm Current Driver Validation"
        '
        'mnuCreateValidationFileToolStripMenuItem
        '
        Me.mnuCreateValidationFileToolStripMenuItem.CheckOnClick = True
        Me.mnuCreateValidationFileToolStripMenuItem.Name = "mnuCreateValidationFileToolStripMenuItem"
        Me.mnuCreateValidationFileToolStripMenuItem.Size = New System.Drawing.Size(250, 22)
        Me.mnuCreateValidationFileToolStripMenuItem.Text = "Create Validation File"
        '
        'SpecialTestsToolStripMenuItem
        '
        Me.SpecialTestsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuTelescopeSideOfPierToolStripMenuItem, Me.mnuTelescopeDestinationSideOfPierToolStripMenuItem, Me.mnuTelescopeFlipRangeToolStripMenuItem, Me.mnuTelescopeCommandsToolStripMenuItem})
        Me.SpecialTestsToolStripMenuItem.Name = "SpecialTestsToolStripMenuItem"
        Me.SpecialTestsToolStripMenuItem.Size = New System.Drawing.Size(85, 20)
        Me.SpecialTestsToolStripMenuItem.Text = "Special Tests"
        '
        'mnuTelescopeSideOfPierToolStripMenuItem
        '
        Me.mnuTelescopeSideOfPierToolStripMenuItem.Name = "mnuTelescopeSideOfPierToolStripMenuItem"
        Me.mnuTelescopeSideOfPierToolStripMenuItem.Size = New System.Drawing.Size(244, 22)
        Me.mnuTelescopeSideOfPierToolStripMenuItem.Text = "Telescope SideOfPier"
        '
        'mnuTelescopeDestinationSideOfPierToolStripMenuItem
        '
        Me.mnuTelescopeDestinationSideOfPierToolStripMenuItem.Name = "mnuTelescopeDestinationSideOfPierToolStripMenuItem"
        Me.mnuTelescopeDestinationSideOfPierToolStripMenuItem.Size = New System.Drawing.Size(244, 22)
        Me.mnuTelescopeDestinationSideOfPierToolStripMenuItem.Text = "Telescope DestinationSideOfPier"
        '
        'mnuTelescopeFlipRangeToolStripMenuItem
        '
        Me.mnuTelescopeFlipRangeToolStripMenuItem.Name = "mnuTelescopeFlipRangeToolStripMenuItem"
        Me.mnuTelescopeFlipRangeToolStripMenuItem.Size = New System.Drawing.Size(244, 22)
        Me.mnuTelescopeFlipRangeToolStripMenuItem.Text = "Telescope SideOfPier Tests"
        '
        'mnuTelescopeCommandsToolStripMenuItem
        '
        Me.mnuTelescopeCommandsToolStripMenuItem.Name = "mnuTelescopeCommandsToolStripMenuItem"
        Me.mnuTelescopeCommandsToolStripMenuItem.Size = New System.Drawing.Size(244, 22)
        Me.mnuTelescopeCommandsToolStripMenuItem.Text = "TelescopeCommands"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HelpToolStripMenuItem1, Me.SideOfPierBehaviourPDFToolStripMenuItem, Me.CheckForUpdatesToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'HelpToolStripMenuItem1
        '
        Me.HelpToolStripMenuItem1.Name = "HelpToolStripMenuItem1"
        Me.HelpToolStripMenuItem1.Size = New System.Drawing.Size(209, 22)
        Me.HelpToolStripMenuItem1.Text = "Help"
        '
        'SideOfPierBehaviourPDFToolStripMenuItem
        '
        Me.SideOfPierBehaviourPDFToolStripMenuItem.Name = "SideOfPierBehaviourPDFToolStripMenuItem"
        Me.SideOfPierBehaviourPDFToolStripMenuItem.Size = New System.Drawing.Size(209, 22)
        Me.SideOfPierBehaviourPDFToolStripMenuItem.Text = "SideOfPier Behaviour PDF"
        '
        'CheckForUpdatesToolStripMenuItem
        '
        Me.CheckForUpdatesToolStripMenuItem.Name = "CheckForUpdatesToolStripMenuItem"
        Me.CheckForUpdatesToolStripMenuItem.Size = New System.Drawing.Size(209, 22)
        Me.CheckForUpdatesToolStripMenuItem.Text = "Check for Updates"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(209, 22)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'txtStatus
        '
        Me.txtStatus.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtStatus.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.Location = New System.Drawing.Point(13, 28)
        Me.txtStatus.Multiline = True
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStatus.Size = New System.Drawing.Size(814, 540)
        Me.txtStatus.TabIndex = 3
        '
        'BtnExit
        '
        Me.BtnExit.Location = New System.Drawing.Point(842, 531)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(94, 37)
        Me.BtnExit.TabIndex = 4
        Me.BtnExit.Text = "Exit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnClear
        '
        Me.BtnClear.Location = New System.Drawing.Point(842, 488)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(94, 37)
        Me.BtnClear.TabIndex = 5
        Me.BtnClear.Text = "Clear"
        Me.BtnClear.UseVisualStyleBackColor = True
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.staTest, Me.staAction, Me.staStatus})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 580)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(948, 22)
        Me.StatusStrip.TabIndex = 6
        Me.StatusStrip.Text = "StatusStrip1"
        '
        'staTest
        '
        Me.staTest.Name = "staTest"
        Me.staTest.Size = New System.Drawing.Size(0, 17)
        '
        'staAction
        '
        Me.staAction.Name = "staAction"
        Me.staAction.Size = New System.Drawing.Size(0, 17)
        '
        'staStatus
        '
        Me.staStatus.Name = "staStatus"
        Me.staStatus.Size = New System.Drawing.Size(0, 17)
        '
        'BtnStop
        '
        Me.BtnStop.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnStop.Location = New System.Drawing.Point(842, 445)
        Me.BtnStop.Name = "BtnStop"
        Me.BtnStop.Size = New System.Drawing.Size(94, 37)
        Me.BtnStop.TabIndex = 7
        Me.BtnStop.Text = "Stop Tests"
        Me.BtnStop.UseVisualStyleBackColor = True
        '
        'BtnClearAndCheckConformance
        '
        Me.BtnClearAndCheckConformance.Location = New System.Drawing.Point(842, 359)
        Me.BtnClearAndCheckConformance.Name = "BtnClearAndCheckConformance"
        Me.BtnClearAndCheckConformance.Size = New System.Drawing.Size(94, 37)
        Me.BtnClearAndCheckConformance.TabIndex = 8
        Me.BtnClearAndCheckConformance.Text = "Clear and Check Conformance"
        Me.BtnClearAndCheckConformance.UseVisualStyleBackColor = True
        '
        'HelpProvider1
        '
        Me.HelpProvider1.HelpNamespace = "Conform.chm"
        '
        'chkDomeShutter
        '
        Me.chkDomeShutter.AutoSize = True
        Me.chkDomeShutter.Location = New System.Drawing.Point(842, 196)
        Me.chkDomeShutter.Name = "chkDomeShutter"
        Me.chkDomeShutter.Size = New System.Drawing.Size(89, 17)
        Me.chkDomeShutter.TabIndex = 9
        Me.chkDomeShutter.Text = "Open Shutter"
        Me.chkDomeShutter.UseVisualStyleBackColor = True
        '
        'dlgValidation
        '
        Me.dlgValidation.FileName = "OpenFileDialog1"
        '
        'chkSwitchSet
        '
        Me.chkSwitchSet.AutoSize = True
        Me.chkSwitchSet.Location = New System.Drawing.Point(842, 196)
        Me.chkSwitchSet.Name = "chkSwitchSet"
        Me.chkSwitchSet.Size = New System.Drawing.Size(88, 17)
        Me.chkSwitchSet.TabIndex = 10
        Me.chkSwitchSet.Text = "Set Switches"
        Me.chkSwitchSet.UseVisualStyleBackColor = True
        '
        'frmConformMain
        '
        Me.AcceptButton = Me.BtnCheckConformance
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnStop
        Me.ClientSize = New System.Drawing.Size(948, 602)
        Me.Controls.Add(Me.chkSwitchSet)
        Me.Controls.Add(Me.chkDomeShutter)
        Me.Controls.Add(Me.BtnClearAndCheckConformance)
        Me.Controls.Add(Me.BtnStop)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.BtnClear)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.BtnCheckConformance)
        Me.Controls.Add(Me.MenuStrip1)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmConformMain"
        Me.Text = "ASCOM Conformance Checker"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnCheckConformance As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckConformanceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SelectDriverToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ConformanceOptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BtnClear As System.Windows.Forms.Button
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents staTest As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents staAction As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents staStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents BtnStop As System.Windows.Forms.Button
    Friend WithEvents mnuTestTelescope As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTestFocuser As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTestDome As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTestCamera As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTestFilterWheel As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTestRotator As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BtnClearAndCheckConformance As System.Windows.Forms.Button
    Friend WithEvents mnuTestSwitch As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Friend WithEvents chkDomeShutter As System.Windows.Forms.CheckBox
    Friend WithEvents ValidateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuConfirmValidationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCreateValidationFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckForUpdatesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents dlgValidation As System.Windows.Forms.OpenFileDialog
    Friend WithEvents mnuTestSafetyMonitor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DriverSetupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SpecialTestsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTelescopeSideOfPierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTelescopeDestinationSideOfPierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTelescopeFlipRangeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SideOfPierBehaviourPDFToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTelescopeCommandsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents chkSwitchSet As System.Windows.Forms.CheckBox
    Friend WithEvents mnuTestVideoCamera As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTestObservingConditions As ToolStripMenuItem
End Class
