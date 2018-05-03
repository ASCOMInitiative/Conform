<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetup
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSetup))
        Me.dirLogFile = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnSetLogFileDir = New System.Windows.Forms.Button()
        Me.txtLogFileDirectory = New System.Windows.Forms.TextBox()
        Me.btnSetupOK = New System.Windows.Forms.Button()
        Me.btnSetupCancel = New System.Windows.Forms.Button()
        Me.chkSetupTestProperties = New System.Windows.Forms.CheckBox()
        Me.chkSetupTestMethods = New System.Windows.Forms.CheckBox()
        Me.chkSetupTestPerformance = New System.Windows.Forms.CheckBox()
        Me.chkUpdates = New System.Windows.Forms.CheckBox()
        Me.txtLastUpdate = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkDisplayMethodCalls = New System.Windows.Forms.CheckBox()
        Me.chkUseDriverAccess = New System.Windows.Forms.CheckBox()
        Me.chkInferStateFromErrorMessages = New System.Windows.Forms.CheckBox()
        Me.chkDebug = New System.Windows.Forms.CheckBox()
        Me.chkRunAs32Bit = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkSetupTestSideOfPierWrite = New System.Windows.Forms.CheckBox()
        Me.chkSetupTestSideOfPierRead = New System.Windows.Forms.CheckBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.General = New System.Windows.Forms.TabPage()
        Me.TelescopeMethods = New System.Windows.Forms.TabPage()
        Me.btnSetTScopeMethods = New System.Windows.Forms.Button()
        Me.btnClearTscopeMethods = New System.Windows.Forms.Button()
        Me.TScopeMethodTests = New System.Windows.Forms.CheckedListBox()
        Me.CameraTests = New System.Windows.Forms.TabPage()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.numMaxBinY = New System.Windows.Forms.NumericUpDown()
        Me.numMaxBinX = New System.Windows.Forms.NumericUpDown()
        Me.DomeTests = New System.Windows.Forms.TabPage()
        Me.txtDomeStabilisationWait = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtDomeAltitudeTimeout = New System.Windows.Forms.TextBox()
        Me.txtDomeAzimuthTimeout = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtDomeShutterTimout = New System.Windows.Forms.TextBox()
        Me.ObservingConditionsTests = New System.Windows.Forms.TabPage()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.UpDownObservingConditionsWaitTime = New System.Windows.Forms.NumericUpDown()
        Me.UpDownObservingConditionsRetries = New System.Windows.Forms.NumericUpDown()
        Me.SwitchTests = New System.Windows.Forms.TabPage()
        Me.txtSwitchReadDelay = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtSwitchWriteDelay = New System.Windows.Forms.TextBox()
        Me.CommandStrings = New System.Windows.Forms.TabPage()
        Me.chkCommandBoolRaw = New System.Windows.Forms.CheckBox()
        Me.chkCommandBlindRaw = New System.Windows.Forms.CheckBox()
        Me.chkCommandStringRaw = New System.Windows.Forms.CheckBox()
        Me.chkCommandBool = New System.Windows.Forms.CheckBox()
        Me.chkCommandBlind = New System.Windows.Forms.CheckBox()
        Me.chkCommandString = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SpecialTests = New System.Windows.Forms.TabPage()
        Me.TabControl2 = New System.Windows.Forms.TabControl()
        Me.DestinationSideOfPier = New System.Windows.Forms.TabPage()
        Me.cmbDSOPSide = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtDECStepE = New System.Windows.Forms.TextBox()
        Me.txtEndDECE = New System.Windows.Forms.TextBox()
        Me.txtStartDECE = New System.Windows.Forms.TextBox()
        Me.txtEndHAE = New System.Windows.Forms.TextBox()
        Me.txtStartHAE = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtDecStepW = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtEndDECW = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtStartDECW = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtEndHAW = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtStartHAW = New System.Windows.Forms.TextBox()
        Me.SideOfPier = New System.Windows.Forms.TabPage()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.txtSwitchExtendedTestRange = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.General.SuspendLayout()
        Me.TelescopeMethods.SuspendLayout()
        Me.CameraTests.SuspendLayout()
        CType(Me.numMaxBinY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numMaxBinX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DomeTests.SuspendLayout()
        Me.ObservingConditionsTests.SuspendLayout()
        CType(Me.UpDownObservingConditionsWaitTime, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UpDownObservingConditionsRetries, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SwitchTests.SuspendLayout()
        Me.CommandStrings.SuspendLayout()
        Me.SpecialTests.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.DestinationSideOfPier.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dirLogFile
        '
        Me.dirLogFile.SelectedPath = "C:\ASCOM Log"
        '
        'btnSetLogFileDir
        '
        Me.btnSetLogFileDir.Location = New System.Drawing.Point(528, 30)
        Me.btnSetLogFileDir.Name = "btnSetLogFileDir"
        Me.btnSetLogFileDir.Size = New System.Drawing.Size(75, 23)
        Me.btnSetLogFileDir.TabIndex = 2
        Me.btnSetLogFileDir.Text = "Select"
        Me.btnSetLogFileDir.UseVisualStyleBackColor = True
        '
        'txtLogFileDirectory
        '
        Me.txtLogFileDirectory.Location = New System.Drawing.Point(15, 32)
        Me.txtLogFileDirectory.Name = "txtLogFileDirectory"
        Me.txtLogFileDirectory.ReadOnly = True
        Me.txtLogFileDirectory.Size = New System.Drawing.Size(496, 20)
        Me.txtLogFileDirectory.TabIndex = 2
        '
        'btnSetupOK
        '
        Me.btnSetupOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnSetupOK.Location = New System.Drawing.Point(522, 375)
        Me.btnSetupOK.Name = "btnSetupOK"
        Me.btnSetupOK.Size = New System.Drawing.Size(75, 23)
        Me.btnSetupOK.TabIndex = 0
        Me.btnSetupOK.Text = "OK"
        Me.btnSetupOK.UseVisualStyleBackColor = True
        '
        'btnSetupCancel
        '
        Me.btnSetupCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnSetupCancel.Location = New System.Drawing.Point(603, 375)
        Me.btnSetupCancel.Name = "btnSetupCancel"
        Me.btnSetupCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnSetupCancel.TabIndex = 1
        Me.btnSetupCancel.Text = "Cancel"
        Me.btnSetupCancel.UseVisualStyleBackColor = True
        '
        'chkSetupTestProperties
        '
        Me.chkSetupTestProperties.AutoSize = True
        Me.chkSetupTestProperties.Checked = True
        Me.chkSetupTestProperties.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSetupTestProperties.Location = New System.Drawing.Point(96, 35)
        Me.chkSetupTestProperties.Name = "chkSetupTestProperties"
        Me.chkSetupTestProperties.Size = New System.Drawing.Size(97, 17)
        Me.chkSetupTestProperties.TabIndex = 5
        Me.chkSetupTestProperties.Text = "Test Properties"
        Me.chkSetupTestProperties.UseVisualStyleBackColor = True
        '
        'chkSetupTestMethods
        '
        Me.chkSetupTestMethods.AutoSize = True
        Me.chkSetupTestMethods.Checked = True
        Me.chkSetupTestMethods.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSetupTestMethods.Location = New System.Drawing.Point(96, 58)
        Me.chkSetupTestMethods.Name = "chkSetupTestMethods"
        Me.chkSetupTestMethods.Size = New System.Drawing.Size(91, 17)
        Me.chkSetupTestMethods.TabIndex = 6
        Me.chkSetupTestMethods.Text = "Test Methods"
        Me.chkSetupTestMethods.UseVisualStyleBackColor = True
        '
        'chkSetupTestPerformance
        '
        Me.chkSetupTestPerformance.AutoSize = True
        Me.chkSetupTestPerformance.Location = New System.Drawing.Point(96, 127)
        Me.chkSetupTestPerformance.Name = "chkSetupTestPerformance"
        Me.chkSetupTestPerformance.Size = New System.Drawing.Size(110, 17)
        Me.chkSetupTestPerformance.TabIndex = 7
        Me.chkSetupTestPerformance.Text = "Test Performance"
        Me.chkSetupTestPerformance.UseVisualStyleBackColor = True
        '
        'chkUpdates
        '
        Me.chkUpdates.AutoSize = True
        Me.chkUpdates.Checked = True
        Me.chkUpdates.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUpdates.Location = New System.Drawing.Point(21, 35)
        Me.chkUpdates.Name = "chkUpdates"
        Me.chkUpdates.Size = New System.Drawing.Size(113, 17)
        Me.chkUpdates.TabIndex = 9
        Me.chkUpdates.Text = "Check for updates"
        Me.chkUpdates.UseVisualStyleBackColor = True
        '
        'txtLastUpdate
        '
        Me.txtLastUpdate.Enabled = False
        Me.txtLastUpdate.Location = New System.Drawing.Point(149, 32)
        Me.txtLastUpdate.Name = "txtLastUpdate"
        Me.txtLastUpdate.Size = New System.Drawing.Size(121, 20)
        Me.txtLastUpdate.TabIndex = 10
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(161, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Last update check"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.chkDisplayMethodCalls)
        Me.GroupBox1.Controls.Add(Me.chkUseDriverAccess)
        Me.GroupBox1.Controls.Add(Me.chkInferStateFromErrorMessages)
        Me.GroupBox1.Controls.Add(Me.chkDebug)
        Me.GroupBox1.Controls.Add(Me.chkRunAs32Bit)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.chkUpdates)
        Me.GroupBox1.Controls.Add(Me.txtLastUpdate)
        Me.GroupBox1.Location = New System.Drawing.Point(352, 127)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(287, 185)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Conform Settings"
        '
        'chkDisplayMethodCalls
        '
        Me.chkDisplayMethodCalls.AutoSize = True
        Me.chkDisplayMethodCalls.Location = New System.Drawing.Point(21, 127)
        Me.chkDisplayMethodCalls.Name = "chkDisplayMethodCalls"
        Me.chkDisplayMethodCalls.Size = New System.Drawing.Size(252, 17)
        Me.chkDisplayMethodCalls.TabIndex = 20
        Me.chkDisplayMethodCalls.Text = "DIsplay method calls (can help to debug drivers)"
        Me.chkDisplayMethodCalls.UseVisualStyleBackColor = True
        '
        'chkUseDriverAccess
        '
        Me.chkUseDriverAccess.AutoSize = True
        Me.chkUseDriverAccess.Location = New System.Drawing.Point(21, 81)
        Me.chkUseDriverAccess.Name = "chkUseDriverAccess"
        Me.chkUseDriverAccess.Size = New System.Drawing.Size(240, 17)
        Me.chkUseDriverAccess.TabIndex = 19
        Me.chkUseDriverAccess.Text = "Use DriverAccess to communicate with driver"
        Me.chkUseDriverAccess.UseVisualStyleBackColor = True
        '
        'chkInferStateFromErrorMessages
        '
        Me.chkInferStateFromErrorMessages.AutoSize = True
        Me.chkInferStateFromErrorMessages.Location = New System.Drawing.Point(21, 104)
        Me.chkInferStateFromErrorMessages.Name = "chkInferStateFromErrorMessages"
        Me.chkInferStateFromErrorMessages.Size = New System.Drawing.Size(185, 17)
        Me.chkInferStateFromErrorMessages.TabIndex = 18
        Me.chkInferStateFromErrorMessages.Text = "Infer state from error message text"
        Me.chkInferStateFromErrorMessages.UseVisualStyleBackColor = True
        '
        'chkDebug
        '
        Me.chkDebug.AutoSize = True
        Me.chkDebug.Location = New System.Drawing.Point(21, 150)
        Me.chkDebug.Name = "chkDebug"
        Me.chkDebug.Size = New System.Drawing.Size(75, 17)
        Me.chkDebug.TabIndex = 17
        Me.chkDebug.Text = "Debug On"
        Me.chkDebug.UseVisualStyleBackColor = True
        '
        'chkRunAs32Bit
        '
        Me.chkRunAs32Bit.AutoSize = True
        Me.chkRunAs32Bit.Location = New System.Drawing.Point(21, 58)
        Me.chkRunAs32Bit.Name = "chkRunAs32Bit"
        Me.chkRunAs32Bit.Size = New System.Drawing.Size(154, 17)
        Me.chkRunAs32Bit.TabIndex = 18
        Me.chkRunAs32Bit.Text = "Run as 32bit on a 64bit OS"
        Me.chkRunAs32Bit.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkSetupTestSideOfPierWrite)
        Me.GroupBox2.Controls.Add(Me.chkSetupTestSideOfPierRead)
        Me.GroupBox2.Controls.Add(Me.chkSetupTestProperties)
        Me.GroupBox2.Controls.Add(Me.chkSetupTestMethods)
        Me.GroupBox2.Controls.Add(Me.chkSetupTestPerformance)
        Me.GroupBox2.Location = New System.Drawing.Point(19, 127)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(287, 185)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Conformance test scope"
        '
        'chkSetupTestSideOfPierWrite
        '
        Me.chkSetupTestSideOfPierWrite.AutoSize = True
        Me.chkSetupTestSideOfPierWrite.Checked = True
        Me.chkSetupTestSideOfPierWrite.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSetupTestSideOfPierWrite.Location = New System.Drawing.Point(96, 104)
        Me.chkSetupTestSideOfPierWrite.Name = "chkSetupTestSideOfPierWrite"
        Me.chkSetupTestSideOfPierWrite.Size = New System.Drawing.Size(128, 17)
        Me.chkSetupTestSideOfPierWrite.TabIndex = 9
        Me.chkSetupTestSideOfPierWrite.Text = "Test SideOfPier Write"
        Me.chkSetupTestSideOfPierWrite.UseVisualStyleBackColor = True
        '
        'chkSetupTestSideOfPierRead
        '
        Me.chkSetupTestSideOfPierRead.AutoSize = True
        Me.chkSetupTestSideOfPierRead.Checked = True
        Me.chkSetupTestSideOfPierRead.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSetupTestSideOfPierRead.Location = New System.Drawing.Point(96, 81)
        Me.chkSetupTestSideOfPierRead.Name = "chkSetupTestSideOfPierRead"
        Me.chkSetupTestSideOfPierRead.Size = New System.Drawing.Size(129, 17)
        Me.chkSetupTestSideOfPierRead.TabIndex = 8
        Me.chkSetupTestSideOfPierRead.Text = "Test SideOfPier Read"
        Me.chkSetupTestSideOfPierRead.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnSetLogFileDir)
        Me.GroupBox4.Controls.Add(Me.txtLogFileDirectory)
        Me.GroupBox4.Location = New System.Drawing.Point(19, 15)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(620, 77)
        Me.GroupBox4.TabIndex = 16
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Log File Directory"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.General)
        Me.TabControl1.Controls.Add(Me.TelescopeMethods)
        Me.TabControl1.Controls.Add(Me.CameraTests)
        Me.TabControl1.Controls.Add(Me.DomeTests)
        Me.TabControl1.Controls.Add(Me.ObservingConditionsTests)
        Me.TabControl1.Controls.Add(Me.SwitchTests)
        Me.TabControl1.Controls.Add(Me.CommandStrings)
        Me.TabControl1.Controls.Add(Me.SpecialTests)
        Me.TabControl1.Location = New System.Drawing.Point(13, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(669, 354)
        Me.TabControl1.TabIndex = 17
        '
        'General
        '
        Me.General.Controls.Add(Me.GroupBox4)
        Me.General.Controls.Add(Me.GroupBox1)
        Me.General.Controls.Add(Me.GroupBox2)
        Me.General.Location = New System.Drawing.Point(4, 22)
        Me.General.Name = "General"
        Me.General.Padding = New System.Windows.Forms.Padding(3)
        Me.General.Size = New System.Drawing.Size(661, 328)
        Me.General.TabIndex = 0
        Me.General.Text = "General"
        Me.General.UseVisualStyleBackColor = True
        '
        'TelescopeMethods
        '
        Me.TelescopeMethods.Controls.Add(Me.btnSetTScopeMethods)
        Me.TelescopeMethods.Controls.Add(Me.btnClearTscopeMethods)
        Me.TelescopeMethods.Controls.Add(Me.TScopeMethodTests)
        Me.TelescopeMethods.Location = New System.Drawing.Point(4, 22)
        Me.TelescopeMethods.Name = "TelescopeMethods"
        Me.TelescopeMethods.Padding = New System.Windows.Forms.Padding(3)
        Me.TelescopeMethods.Size = New System.Drawing.Size(661, 328)
        Me.TelescopeMethods.TabIndex = 4
        Me.TelescopeMethods.Text = "Telescope Method Tests"
        Me.TelescopeMethods.UseVisualStyleBackColor = True
        '
        'btnSetTScopeMethods
        '
        Me.btnSetTScopeMethods.Location = New System.Drawing.Point(460, 132)
        Me.btnSetTScopeMethods.Name = "btnSetTScopeMethods"
        Me.btnSetTScopeMethods.Size = New System.Drawing.Size(75, 23)
        Me.btnSetTScopeMethods.TabIndex = 2
        Me.btnSetTScopeMethods.Text = "Set Values"
        Me.btnSetTScopeMethods.UseVisualStyleBackColor = True
        '
        'btnClearTscopeMethods
        '
        Me.btnClearTscopeMethods.Location = New System.Drawing.Point(460, 102)
        Me.btnClearTscopeMethods.Name = "btnClearTscopeMethods"
        Me.btnClearTscopeMethods.Size = New System.Drawing.Size(75, 23)
        Me.btnClearTscopeMethods.TabIndex = 1
        Me.btnClearTscopeMethods.Text = "Clear Values"
        Me.btnClearTscopeMethods.UseVisualStyleBackColor = True
        '
        'TScopeMethodTests
        '
        Me.TScopeMethodTests.FormattingEnabled = True
        Me.TScopeMethodTests.Location = New System.Drawing.Point(193, 18)
        Me.TScopeMethodTests.Name = "TScopeMethodTests"
        Me.TScopeMethodTests.Size = New System.Drawing.Size(239, 274)
        Me.TScopeMethodTests.TabIndex = 0
        '
        'CameraTests
        '
        Me.CameraTests.Controls.Add(Me.Label22)
        Me.CameraTests.Controls.Add(Me.Label21)
        Me.CameraTests.Controls.Add(Me.Label20)
        Me.CameraTests.Controls.Add(Me.Label19)
        Me.CameraTests.Controls.Add(Me.Label18)
        Me.CameraTests.Controls.Add(Me.numMaxBinY)
        Me.CameraTests.Controls.Add(Me.numMaxBinX)
        Me.CameraTests.Location = New System.Drawing.Point(4, 22)
        Me.CameraTests.Name = "CameraTests"
        Me.CameraTests.Size = New System.Drawing.Size(661, 328)
        Me.CameraTests.TabIndex = 8
        Me.CameraTests.Text = "Camera Tests"
        Me.CameraTests.UseVisualStyleBackColor = True
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(129, 171)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(410, 13)
        Me.Label22.TabIndex = 6
        Me.Label22.Text = "To use the maximum bin values supplied by the camera, set the values above to zer" &
    "o."
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(142, 63)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(370, 13)
        Me.Label21.TabIndex = 5
        Me.Label21.Text = "The settings below can be used to limit testing to specificed maximum values."
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(39, 50)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(592, 13)
        Me.Label20.TabIndex = 4
        Me.Label20.Text = "Conform will normally test camera imaging at all supported bin values that the ca" &
    "mera reports through MaxBinX and MaxBinY."
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(336, 129)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(104, 13)
        Me.Label19.TabIndex = 3
        Me.Label19.Text = "Max Bin Y for testing"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(336, 102)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(104, 13)
        Me.Label18.TabIndex = 2
        Me.Label18.Text = "Max Bin X for testing"
        '
        'numMaxBinY
        '
        Me.numMaxBinY.Location = New System.Drawing.Point(241, 127)
        Me.numMaxBinY.Name = "numMaxBinY"
        Me.numMaxBinY.Size = New System.Drawing.Size(89, 20)
        Me.numMaxBinY.TabIndex = 1
        Me.numMaxBinY.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'numMaxBinX
        '
        Me.numMaxBinX.Location = New System.Drawing.Point(241, 100)
        Me.numMaxBinX.Name = "numMaxBinX"
        Me.numMaxBinX.Size = New System.Drawing.Size(89, 20)
        Me.numMaxBinX.TabIndex = 0
        Me.numMaxBinX.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DomeTests
        '
        Me.DomeTests.Controls.Add(Me.txtDomeStabilisationWait)
        Me.DomeTests.Controls.Add(Me.Label13)
        Me.DomeTests.Controls.Add(Me.txtDomeAltitudeTimeout)
        Me.DomeTests.Controls.Add(Me.txtDomeAzimuthTimeout)
        Me.DomeTests.Controls.Add(Me.Label12)
        Me.DomeTests.Controls.Add(Me.Label11)
        Me.DomeTests.Controls.Add(Me.Label10)
        Me.DomeTests.Controls.Add(Me.txtDomeShutterTimout)
        Me.DomeTests.Location = New System.Drawing.Point(4, 22)
        Me.DomeTests.Name = "DomeTests"
        Me.DomeTests.Padding = New System.Windows.Forms.Padding(3)
        Me.DomeTests.Size = New System.Drawing.Size(661, 328)
        Me.DomeTests.TabIndex = 5
        Me.DomeTests.Text = "Dome Tests"
        Me.DomeTests.UseVisualStyleBackColor = True
        '
        'txtDomeStabilisationWait
        '
        Me.txtDomeStabilisationWait.Location = New System.Drawing.Point(347, 191)
        Me.txtDomeStabilisationWait.Name = "txtDomeStabilisationWait"
        Me.txtDomeStabilisationWait.Size = New System.Drawing.Size(100, 20)
        Me.txtDomeStabilisationWait.TabIndex = 7
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(171, 194)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(168, 13)
        Me.Label13.TabIndex = 6
        Me.Label13.Text = "Dome Stabilisation Wait (seconds)"
        '
        'txtDomeAltitudeTimeout
        '
        Me.txtDomeAltitudeTimeout.Location = New System.Drawing.Point(347, 164)
        Me.txtDomeAltitudeTimeout.Name = "txtDomeAltitudeTimeout"
        Me.txtDomeAltitudeTimeout.Size = New System.Drawing.Size(100, 20)
        Me.txtDomeAltitudeTimeout.TabIndex = 5
        '
        'txtDomeAzimuthTimeout
        '
        Me.txtDomeAzimuthTimeout.Location = New System.Drawing.Point(347, 135)
        Me.txtDomeAzimuthTimeout.Name = "txtDomeAzimuthTimeout"
        Me.txtDomeAzimuthTimeout.Size = New System.Drawing.Size(100, 20)
        Me.txtDomeAzimuthTimeout.TabIndex = 4
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(176, 167)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(163, 13)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "Dome Altitude Timeout (seconds)"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(174, 138)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(165, 13)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "Dome Azimuth Timeout (seconds)"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(177, 112)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(162, 13)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "Dome Shutter Timeout (seconds)"
        '
        'txtDomeShutterTimout
        '
        Me.txtDomeShutterTimout.Location = New System.Drawing.Point(347, 109)
        Me.txtDomeShutterTimout.Name = "txtDomeShutterTimout"
        Me.txtDomeShutterTimout.Size = New System.Drawing.Size(100, 20)
        Me.txtDomeShutterTimout.TabIndex = 0
        '
        'ObservingConditionsTests
        '
        Me.ObservingConditionsTests.Controls.Add(Me.Label17)
        Me.ObservingConditionsTests.Controls.Add(Me.Label16)
        Me.ObservingConditionsTests.Controls.Add(Me.UpDownObservingConditionsWaitTime)
        Me.ObservingConditionsTests.Controls.Add(Me.UpDownObservingConditionsRetries)
        Me.ObservingConditionsTests.Location = New System.Drawing.Point(4, 22)
        Me.ObservingConditionsTests.Name = "ObservingConditionsTests"
        Me.ObservingConditionsTests.Padding = New System.Windows.Forms.Padding(3)
        Me.ObservingConditionsTests.Size = New System.Drawing.Size(661, 328)
        Me.ObservingConditionsTests.TabIndex = 7
        Me.ObservingConditionsTests.Text = "Observing Conditions Tests"
        Me.ObservingConditionsTests.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(86, 134)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(385, 13)
        Me.Label17.TabIndex = 3
        Me.Label17.Text = "Number of attempts to read a property and InvalidOperationException is returned"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(183, 160)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(288, 13)
        Me.Label16.TabIndex = 2
        Me.Label16.Text = "Wait time between InvalidOperationException retry attempts"
        '
        'UpDownObservingConditionsWaitTime
        '
        Me.UpDownObservingConditionsWaitTime.Location = New System.Drawing.Point(477, 158)
        Me.UpDownObservingConditionsWaitTime.Name = "UpDownObservingConditionsWaitTime"
        Me.UpDownObservingConditionsWaitTime.Size = New System.Drawing.Size(51, 20)
        Me.UpDownObservingConditionsWaitTime.TabIndex = 1
        '
        'UpDownObservingConditionsRetries
        '
        Me.UpDownObservingConditionsRetries.Location = New System.Drawing.Point(477, 132)
        Me.UpDownObservingConditionsRetries.Name = "UpDownObservingConditionsRetries"
        Me.UpDownObservingConditionsRetries.Size = New System.Drawing.Size(51, 20)
        Me.UpDownObservingConditionsRetries.TabIndex = 0
        '
        'SwitchTests
        '
        Me.SwitchTests.Controls.Add(Me.Label24)
        Me.SwitchTests.Controls.Add(Me.Label23)
        Me.SwitchTests.Controls.Add(Me.txtSwitchExtendedTestRange)
        Me.SwitchTests.Controls.Add(Me.txtSwitchReadDelay)
        Me.SwitchTests.Controls.Add(Me.Label15)
        Me.SwitchTests.Controls.Add(Me.Label14)
        Me.SwitchTests.Controls.Add(Me.txtSwitchWriteDelay)
        Me.SwitchTests.Location = New System.Drawing.Point(4, 22)
        Me.SwitchTests.Name = "SwitchTests"
        Me.SwitchTests.Padding = New System.Windows.Forms.Padding(3)
        Me.SwitchTests.Size = New System.Drawing.Size(661, 328)
        Me.SwitchTests.TabIndex = 6
        Me.SwitchTests.Text = "SwitchTests"
        Me.SwitchTests.UseVisualStyleBackColor = True
        '
        'txtSwitchReadDelay
        '
        Me.txtSwitchReadDelay.Location = New System.Drawing.Point(338, 123)
        Me.txtSwitchReadDelay.Name = "txtSwitchReadDelay"
        Me.txtSwitchReadDelay.Size = New System.Drawing.Size(100, 20)
        Me.txtSwitchReadDelay.TabIndex = 0
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(213, 152)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(119, 13)
        Me.Label15.TabIndex = 3
        Me.Label15.Text = "Switch Write Delay (ms)"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(212, 126)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(120, 13)
        Me.Label14.TabIndex = 2
        Me.Label14.Text = "Switch Read Delay (ms)"
        '
        'txtSwitchWriteDelay
        '
        Me.txtSwitchWriteDelay.Location = New System.Drawing.Point(338, 149)
        Me.txtSwitchWriteDelay.Name = "txtSwitchWriteDelay"
        Me.txtSwitchWriteDelay.Size = New System.Drawing.Size(100, 20)
        Me.txtSwitchWriteDelay.TabIndex = 1
        '
        'CommandStrings
        '
        Me.CommandStrings.Controls.Add(Me.chkCommandBoolRaw)
        Me.CommandStrings.Controls.Add(Me.chkCommandBlindRaw)
        Me.CommandStrings.Controls.Add(Me.chkCommandStringRaw)
        Me.CommandStrings.Controls.Add(Me.chkCommandBool)
        Me.CommandStrings.Controls.Add(Me.chkCommandBlind)
        Me.CommandStrings.Controls.Add(Me.chkCommandString)
        Me.CommandStrings.Controls.Add(Me.Label1)
        Me.CommandStrings.Location = New System.Drawing.Point(4, 22)
        Me.CommandStrings.Name = "CommandStrings"
        Me.CommandStrings.Padding = New System.Windows.Forms.Padding(3)
        Me.CommandStrings.Size = New System.Drawing.Size(661, 328)
        Me.CommandStrings.TabIndex = 1
        Me.CommandStrings.Text = "Command Tests"
        Me.CommandStrings.UseVisualStyleBackColor = True
        '
        'chkCommandBoolRaw
        '
        Me.chkCommandBoolRaw.AutoSize = True
        Me.chkCommandBoolRaw.Location = New System.Drawing.Point(279, 168)
        Me.chkCommandBoolRaw.Name = "chkCommandBoolRaw"
        Me.chkCommandBoolRaw.Size = New System.Drawing.Size(119, 17)
        Me.chkCommandBoolRaw.TabIndex = 6
        Me.chkCommandBoolRaw.Text = "CommandBool Raw"
        Me.chkCommandBoolRaw.UseVisualStyleBackColor = True
        '
        'chkCommandBlindRaw
        '
        Me.chkCommandBlindRaw.AutoSize = True
        Me.chkCommandBlindRaw.Location = New System.Drawing.Point(279, 145)
        Me.chkCommandBlindRaw.Name = "chkCommandBlindRaw"
        Me.chkCommandBlindRaw.Size = New System.Drawing.Size(121, 17)
        Me.chkCommandBlindRaw.TabIndex = 5
        Me.chkCommandBlindRaw.Text = "CommandBlind Raw"
        Me.chkCommandBlindRaw.UseVisualStyleBackColor = True
        '
        'chkCommandStringRaw
        '
        Me.chkCommandStringRaw.AutoSize = True
        Me.chkCommandStringRaw.Location = New System.Drawing.Point(279, 122)
        Me.chkCommandStringRaw.Name = "chkCommandStringRaw"
        Me.chkCommandStringRaw.Size = New System.Drawing.Size(125, 17)
        Me.chkCommandStringRaw.TabIndex = 4
        Me.chkCommandStringRaw.Text = "CommandString Raw"
        Me.chkCommandStringRaw.UseVisualStyleBackColor = True
        '
        'chkCommandBool
        '
        Me.chkCommandBool.AutoSize = True
        Me.chkCommandBool.Checked = True
        Me.chkCommandBool.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCommandBool.Location = New System.Drawing.Point(279, 99)
        Me.chkCommandBool.Name = "chkCommandBool"
        Me.chkCommandBool.Size = New System.Drawing.Size(94, 17)
        Me.chkCommandBool.TabIndex = 3
        Me.chkCommandBool.Text = "CommandBool"
        Me.chkCommandBool.UseVisualStyleBackColor = True
        '
        'chkCommandBlind
        '
        Me.chkCommandBlind.AutoSize = True
        Me.chkCommandBlind.Checked = True
        Me.chkCommandBlind.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCommandBlind.Location = New System.Drawing.Point(279, 76)
        Me.chkCommandBlind.Name = "chkCommandBlind"
        Me.chkCommandBlind.Size = New System.Drawing.Size(96, 17)
        Me.chkCommandBlind.TabIndex = 2
        Me.chkCommandBlind.Text = "CommandBlind"
        Me.chkCommandBlind.UseVisualStyleBackColor = True
        '
        'chkCommandString
        '
        Me.chkCommandString.AutoSize = True
        Me.chkCommandString.Checked = True
        Me.chkCommandString.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCommandString.Location = New System.Drawing.Point(279, 53)
        Me.chkCommandString.Name = "chkCommandString"
        Me.chkCommandString.Size = New System.Drawing.Size(100, 17)
        Me.chkCommandString.TabIndex = 1
        Me.chkCommandString.Text = "CommandString"
        Me.chkCommandString.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(276, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(144, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Carry out test using null string"
        '
        'SpecialTests
        '
        Me.SpecialTests.Controls.Add(Me.TabControl2)
        Me.SpecialTests.Location = New System.Drawing.Point(4, 22)
        Me.SpecialTests.Name = "SpecialTests"
        Me.SpecialTests.Padding = New System.Windows.Forms.Padding(3)
        Me.SpecialTests.Size = New System.Drawing.Size(661, 328)
        Me.SpecialTests.TabIndex = 3
        Me.SpecialTests.Text = "Special Tests"
        Me.SpecialTests.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.DestinationSideOfPier)
        Me.TabControl2.Controls.Add(Me.SideOfPier)
        Me.TabControl2.Location = New System.Drawing.Point(-5, 0)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(670, 254)
        Me.TabControl2.TabIndex = 0
        '
        'DestinationSideOfPier
        '
        Me.DestinationSideOfPier.Controls.Add(Me.cmbDSOPSide)
        Me.DestinationSideOfPier.Controls.Add(Me.Label9)
        Me.DestinationSideOfPier.Controls.Add(Me.Label8)
        Me.DestinationSideOfPier.Controls.Add(Me.txtDECStepE)
        Me.DestinationSideOfPier.Controls.Add(Me.txtEndDECE)
        Me.DestinationSideOfPier.Controls.Add(Me.txtStartDECE)
        Me.DestinationSideOfPier.Controls.Add(Me.txtEndHAE)
        Me.DestinationSideOfPier.Controls.Add(Me.txtStartHAE)
        Me.DestinationSideOfPier.Controls.Add(Me.Label7)
        Me.DestinationSideOfPier.Controls.Add(Me.txtDecStepW)
        Me.DestinationSideOfPier.Controls.Add(Me.Label6)
        Me.DestinationSideOfPier.Controls.Add(Me.txtEndDECW)
        Me.DestinationSideOfPier.Controls.Add(Me.Label5)
        Me.DestinationSideOfPier.Controls.Add(Me.txtStartDECW)
        Me.DestinationSideOfPier.Controls.Add(Me.Label4)
        Me.DestinationSideOfPier.Controls.Add(Me.txtEndHAW)
        Me.DestinationSideOfPier.Controls.Add(Me.Label3)
        Me.DestinationSideOfPier.Controls.Add(Me.txtStartHAW)
        Me.DestinationSideOfPier.Location = New System.Drawing.Point(4, 22)
        Me.DestinationSideOfPier.Name = "DestinationSideOfPier"
        Me.DestinationSideOfPier.Padding = New System.Windows.Forms.Padding(3)
        Me.DestinationSideOfPier.Size = New System.Drawing.Size(662, 228)
        Me.DestinationSideOfPier.TabIndex = 0
        Me.DestinationSideOfPier.Text = "DestinationSideOfPier"
        Me.DestinationSideOfPier.UseVisualStyleBackColor = True
        '
        'cmbDSOPSide
        '
        Me.cmbDSOPSide.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbDSOPSide.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbDSOPSide.FormattingEnabled = True
        Me.cmbDSOPSide.Items.AddRange(New Object() {"pierEast only", "pierWest only", "pierEast and pierWest"})
        Me.cmbDSOPSide.Location = New System.Drawing.Point(477, 30)
        Me.cmbDSOPSide.Name = "cmbDSOPSide"
        Me.cmbDSOPSide.Size = New System.Drawing.Size(128, 21)
        Me.cmbDSOPSide.TabIndex = 17
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(326, 14)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(91, 13)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "Scope at pierEast"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(159, 14)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(95, 13)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "Scope at pierWest"
        '
        'txtDECStepE
        '
        Me.txtDECStepE.Location = New System.Drawing.Point(309, 134)
        Me.txtDECStepE.Name = "txtDECStepE"
        Me.txtDECStepE.Size = New System.Drawing.Size(128, 20)
        Me.txtDECStepE.TabIndex = 14
        '
        'txtEndDECE
        '
        Me.txtEndDECE.Location = New System.Drawing.Point(309, 108)
        Me.txtEndDECE.Name = "txtEndDECE"
        Me.txtEndDECE.Size = New System.Drawing.Size(128, 20)
        Me.txtEndDECE.TabIndex = 13
        '
        'txtStartDECE
        '
        Me.txtStartDECE.Location = New System.Drawing.Point(309, 82)
        Me.txtStartDECE.Name = "txtStartDECE"
        Me.txtStartDECE.Size = New System.Drawing.Size(128, 20)
        Me.txtStartDECE.TabIndex = 12
        '
        'txtEndHAE
        '
        Me.txtEndHAE.Location = New System.Drawing.Point(309, 56)
        Me.txtEndHAE.Name = "txtEndHAE"
        Me.txtEndHAE.Size = New System.Drawing.Size(128, 20)
        Me.txtEndHAE.TabIndex = 11
        '
        'txtStartHAE
        '
        Me.txtStartHAE.Location = New System.Drawing.Point(309, 30)
        Me.txtStartHAE.Name = "txtStartHAE"
        Me.txtStartHAE.Size = New System.Drawing.Size(128, 20)
        Me.txtStartHAE.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(17, 137)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(108, 13)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Declination Step Size"
        '
        'txtDecStepW
        '
        Me.txtDecStepW.Location = New System.Drawing.Point(142, 134)
        Me.txtDecStepW.Name = "txtDecStepW"
        Me.txtDecStepW.Size = New System.Drawing.Size(128, 20)
        Me.txtDecStepW.TabIndex = 8
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(43, 111)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(82, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "End Declination"
        '
        'txtEndDECW
        '
        Me.txtEndDECW.Location = New System.Drawing.Point(142, 108)
        Me.txtEndDECW.Name = "txtEndDECW"
        Me.txtEndDECW.Size = New System.Drawing.Size(128, 20)
        Me.txtEndDECW.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(40, 85)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(85, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Start Declination"
        '
        'txtStartDECW
        '
        Me.txtStartDECW.Location = New System.Drawing.Point(142, 82)
        Me.txtStartDECW.Name = "txtStartDECW"
        Me.txtStartDECW.Size = New System.Drawing.Size(128, 20)
        Me.txtStartDECW.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(43, 59)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "End Hour Angle"
        '
        'txtEndHAW
        '
        Me.txtEndHAW.Location = New System.Drawing.Point(142, 56)
        Me.txtEndHAW.Name = "txtEndHAW"
        Me.txtEndHAW.Size = New System.Drawing.Size(128, 20)
        Me.txtEndHAW.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(40, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(85, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Start Hour Angle"
        '
        'txtStartHAW
        '
        Me.txtStartHAW.Location = New System.Drawing.Point(142, 30)
        Me.txtStartHAW.Name = "txtStartHAW"
        Me.txtStartHAW.Size = New System.Drawing.Size(128, 20)
        Me.txtStartHAW.TabIndex = 0
        '
        'SideOfPier
        '
        Me.SideOfPier.Location = New System.Drawing.Point(4, 22)
        Me.SideOfPier.Name = "SideOfPier"
        Me.SideOfPier.Padding = New System.Windows.Forms.Padding(3)
        Me.SideOfPier.Size = New System.Drawing.Size(662, 228)
        Me.SideOfPier.TabIndex = 1
        Me.SideOfPier.Text = "SideOfPier"
        Me.SideOfPier.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'txtSwitchExtendedTestRange
        '
        Me.txtSwitchExtendedTestRange.Location = New System.Drawing.Point(338, 175)
        Me.txtSwitchExtendedTestRange.Name = "txtSwitchExtendedTestRange"
        Me.txtSwitchExtendedTestRange.Size = New System.Drawing.Size(100, 20)
        Me.txtSwitchExtendedTestRange.TabIndex = 4
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(159, 178)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(173, 13)
        Me.Label23.TabIndex = 6
        Me.Label23.Text = "Extended switch number test range"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(444, 178)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(137, 13)
        Me.Label24.TabIndex = 7
        Me.Label24.Text = "-XXX ... +MaxSwitch + XXX"
        '
        'frmSetup
        '
        Me.AcceptButton = Me.btnSetupOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnSetupCancel
        Me.ClientSize = New System.Drawing.Size(702, 410)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnSetupCancel)
        Me.Controls.Add(Me.btnSetupOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSetup"
        Me.Text = "ASCOM Conformance Checker Setup"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.General.ResumeLayout(False)
        Me.TelescopeMethods.ResumeLayout(False)
        Me.CameraTests.ResumeLayout(False)
        Me.CameraTests.PerformLayout()
        CType(Me.numMaxBinY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numMaxBinX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DomeTests.ResumeLayout(False)
        Me.DomeTests.PerformLayout()
        Me.ObservingConditionsTests.ResumeLayout(False)
        Me.ObservingConditionsTests.PerformLayout()
        CType(Me.UpDownObservingConditionsWaitTime, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UpDownObservingConditionsRetries, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SwitchTests.ResumeLayout(False)
        Me.SwitchTests.PerformLayout()
        Me.CommandStrings.ResumeLayout(False)
        Me.CommandStrings.PerformLayout()
        Me.SpecialTests.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.DestinationSideOfPier.ResumeLayout(False)
        Me.DestinationSideOfPier.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dirLogFile As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents btnSetLogFileDir As System.Windows.Forms.Button
    Friend WithEvents txtLogFileDirectory As System.Windows.Forms.TextBox
    Friend WithEvents btnSetupOK As System.Windows.Forms.Button
    Friend WithEvents btnSetupCancel As System.Windows.Forms.Button
    Friend WithEvents chkSetupTestProperties As System.Windows.Forms.CheckBox
    Friend WithEvents chkSetupTestMethods As System.Windows.Forms.CheckBox
    Friend WithEvents chkSetupTestPerformance As System.Windows.Forms.CheckBox
    Friend WithEvents chkUpdates As System.Windows.Forms.CheckBox
    Friend WithEvents txtLastUpdate As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents General As System.Windows.Forms.TabPage
    Friend WithEvents CommandStrings As System.Windows.Forms.TabPage
    Friend WithEvents chkCommandBoolRaw As System.Windows.Forms.CheckBox
    Friend WithEvents chkCommandBlindRaw As System.Windows.Forms.CheckBox
    Friend WithEvents chkCommandStringRaw As System.Windows.Forms.CheckBox
    Friend WithEvents chkCommandBool As System.Windows.Forms.CheckBox
    Friend WithEvents chkCommandBlind As System.Windows.Forms.CheckBox
    Friend WithEvents chkCommandString As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents DestinationSideOfPier As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtStartHAW As System.Windows.Forms.TextBox
    Friend WithEvents SideOfPier As System.Windows.Forms.TabPage
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtDecStepW As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtEndDECW As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtStartDECW As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtEndHAW As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtDECStepE As System.Windows.Forms.TextBox
    Friend WithEvents txtEndDECE As System.Windows.Forms.TextBox
    Friend WithEvents txtStartDECE As System.Windows.Forms.TextBox
    Friend WithEvents txtEndHAE As System.Windows.Forms.TextBox
    Friend WithEvents txtStartHAE As System.Windows.Forms.TextBox
    Friend WithEvents cmbDSOPSide As System.Windows.Forms.ComboBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents chkSetupTestSideOfPierRead As System.Windows.Forms.CheckBox
    Friend WithEvents SpecialTests As System.Windows.Forms.TabPage
    Friend WithEvents TelescopeMethods As System.Windows.Forms.TabPage
    Friend WithEvents TScopeMethodTests As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnSetTScopeMethods As System.Windows.Forms.Button
    Friend WithEvents btnClearTscopeMethods As System.Windows.Forms.Button
    Friend WithEvents DomeTests As System.Windows.Forms.TabPage
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtDomeShutterTimout As System.Windows.Forms.TextBox
    Friend WithEvents txtDomeAltitudeTimeout As System.Windows.Forms.TextBox
    Friend WithEvents txtDomeAzimuthTimeout As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtDomeStabilisationWait As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents chkRunAs32Bit As System.Windows.Forms.CheckBox
    Friend WithEvents chkDebug As System.Windows.Forms.CheckBox
    Friend WithEvents chkInferStateFromErrorMessages As System.Windows.Forms.CheckBox
    Friend WithEvents SwitchTests As System.Windows.Forms.TabPage
    Friend WithEvents txtSwitchReadDelay As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtSwitchWriteDelay As System.Windows.Forms.TextBox
    Friend WithEvents chkUseDriverAccess As System.Windows.Forms.CheckBox
    Friend WithEvents ObservingConditionsTests As TabPage
    Friend WithEvents Label17 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents UpDownObservingConditionsWaitTime As NumericUpDown
    Friend WithEvents UpDownObservingConditionsRetries As NumericUpDown
    Friend WithEvents chkSetupTestSideOfPierWrite As CheckBox
    Friend WithEvents chkDisplayMethodCalls As CheckBox
    Friend WithEvents CameraTests As TabPage
    Friend WithEvents Label22 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents Label20 As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents numMaxBinY As NumericUpDown
    Friend WithEvents numMaxBinX As NumericUpDown
    Friend WithEvents Label24 As Label
    Friend WithEvents Label23 As Label
    Friend WithEvents txtSwitchExtendedTestRange As TextBox
End Class
