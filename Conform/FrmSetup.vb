Imports ASCOM.DeviceInterface
Imports System.Windows.Forms
Imports ASCOM.Utilities

Public Class frmSetup
    Private Sub BtnSetLogFileDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetLogFileDir.Click
        Dim l_DlgResult As System.Windows.Forms.DialogResult
        l_DlgResult = dirLogFile.ShowDialog()
        If l_DlgResult = Windows.Forms.DialogResult.OK Then txtLogFileDirectory.Text = dirLogFile.SelectedPath
    End Sub

    Private Sub FrmSetup_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
#If Not Debug Then
        TabControl1.Controls.Remove(Me.TabControl1.Controls.Item("SpecialTests")) 'Hide special tests configuration in release version
#End If
        txtLogFileDirectory.Text = g_Settings.LogFileDirectory
        dirLogFile.SelectedPath = g_Settings.LogFileDirectory
        chkSetupTestProperties.Checked = g_Settings.TestProperties
        chkSetupTestMethods.Checked = g_Settings.TestMethods
        chkSetupTestSideOfPierRead.Checked = g_Settings.TestSideOfPierRead
        chkSetupTestSideOfPierWrite.Checked = g_Settings.TestSideOfPierWrite
        chkSetupTestPerformance.Checked = g_Settings.TestPerformance
        chkUpdates.Checked = g_Settings.UpdateCheck
        txtLastUpdate.Text = g_Settings.UpdateCheckDate.ToString
        chkDebug.Checked = g_Settings.Debug
        chkRunAs32Bit.Checked = g_Settings.RunAs32Bit
        If VersionCode.OSBits = Bitness.Bits32 Then chkRunAs32Bit.Enabled = False 'If we are on a 32bit OS doable this option
        chkInferStateFromErrorMessages.Checked = g_Settings.InterpretErrorMessages
        chkUseDriverAccess.Checked = g_Settings.UseDriverAccess
        UpDownObservingConditionsRetries.Value = Convert.ToDecimal(g_Settings.ObservingConditionsMaxRetries)
        UpDownObservingConditionsWaitTime.Value = Convert.ToDecimal(g_Settings.ObservingConditionsRetryTime)
        chkDisplayMethodCalls.Checked = g_Settings.DisplayMethodCalls
        numMaxBinX.Value = Convert.ToInt32(g_Settings.CameraMaxBinX)
        numMaxBinY.Value = Convert.ToInt32(g_Settings.CameraMaxBinY)
        txtSwitchExtendedTestRange.Text = g_Settings.ExtendedSwitchNumberTestRange

        If chkSetupTestSideOfPierWrite.Checked Then ' Test SideOfpier Write is enabled
            chkSetupTestSideOfPierRead.Checked = True
            chkSetupTestMethods.Checked = True
            chkSetupTestProperties.Checked = True
            chkSetupTestSideOfPierRead.Enabled = False
            chkSetupTestMethods.Enabled = False
            chkSetupTestProperties.Enabled = False
        Else ' SideofPier write is not enabled
            If chkSetupTestSideOfPierRead.Checked Then ' Test SideOfPier read is enabled
                chkSetupTestMethods.Checked = True
                chkSetupTestProperties.Checked = True
                chkSetupTestMethods.Enabled = False
                chkSetupTestProperties.Enabled = False
            Else ' Neither Test SideOfPier Read nor Test SideOfPier Write are enabled
                If chkSetupTestMethods.Checked Then ' Test Methods is enabled
                    chkSetupTestProperties.Checked = True
                    chkSetupTestProperties.Enabled = False
                Else
                    chkSetupTestProperties.Enabled = True ' Just Test Properties is enabled
                End If
            End If
        End If

        'Set up command test check values
        chkCommandString.Checked = g_Settings.CommandString
        chkCommandBlind.Checked = g_Settings.CommandBlind
        chkCommandBool.Checked = g_Settings.CommandBool
        chkCommandStringRaw.Checked = g_Settings.CommandStringRaw
        chkCommandBlindRaw.Checked = g_Settings.CommandBlindRaw
        chkCommandBoolRaw.Checked = g_Settings.CommandBoolRaw

        'Dome Tests
        txtDomeAltitudeTimeout.Text = g_Settings.DomeAltitudeTimeout.ToString
        txtDomeAzimuthTimeout.Text = g_Settings.DomeAzimuthTimeout.ToString
        txtDomeShutterTimout.Text = g_Settings.DomeShutterTimeout.ToString
        txtDomeStabilisationWait.Text = g_Settings.DomeStabilisationWait.ToString

        'Switch Tests
        txtSwitchReadDelay.Text = g_Settings.SwitchReadDelay.ToString
        txtSwitchWriteDelay.Text = g_Settings.SwitchWriteDelay.ToString

        'Special - DestinationSideOfPier
        txtStartHAW.Text = g_Settings.FlipTestHAStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest).ToString
        txtEndHAW.Text = g_Settings.FlipTestHAEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest).ToString
        txtStartDECW.Text = g_Settings.FlipTestDECStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest).ToString
        txtEndDECW.Text = g_Settings.FlipTestDECEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest).ToString
        txtDecStepW.Text = g_Settings.FlipTestDECStep(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest).ToString

        txtStartHAE.Text = (-g_Settings.FlipTestHAStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast)).ToString
        txtEndHAE.Text = (-g_Settings.FlipTestHAEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast)).ToString
        txtStartDECE.Text = g_Settings.FlipTestDECStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast).ToString
        txtEndDECE.Text = g_Settings.FlipTestDECEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast).ToString
        txtDECStepE.Text = g_Settings.FlipTestDECStep(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast).ToString

        cmbDSOPSide.Text = g_Settings.DSOPSide

        'Retrieve stored values for all the requested parameters
        g_TelescopeTests = g_Settings.TeleScopeTests()
        TScopeMethodTests.CheckOnClick = True
        TScopeMethodTests.Items.Clear()

        'Construct the collection from the contents of the checked list box
        For Each kvp As Generic.KeyValuePair(Of String, CheckState) In g_TelescopeTests
            TScopeMethodTests.Items.Add(kvp.Key, kvp.Value = CheckState.Checked)
        Next
        LogMsg("", MessageLevel.msgAlways, "")

        btnSetupOK.Focus()
    End Sub

    Private Sub FrmSetup_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles Me.FormClosed
        Dim NewBitness As String
        If VersionCode.OSBits = Utilities.Bitness.Bits64 Then 'Only test for a bitness switch if we are on a 64bt OS
            If Not ((g_Settings.RunAs32Bit = (IntPtr.Size = 4)) Or (Not g_Settings.RunAs32Bit = (IntPtr.Size = 8))) Then
                If g_Settings.RunAs32Bit Then
                    NewBitness = "32bit"
                Else
                    NewBitness = "64bit"
                End If

                If MsgBox("Do you wish to restart Conform as " & NewBitness, MsgBoxStyle.YesNo, "Restart Conform") = MsgBoxResult.Yes Then
                    ChangeBitnessIfNeeded()
                End If
            End If
        End If
    End Sub

    Private Sub BtnSetupOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetupOK.Click
        Try
            g_Settings.LogFileDirectory = txtLogFileDirectory.Text
            g_Settings.TestProperties = chkSetupTestProperties.Checked
            g_Settings.TestMethods = chkSetupTestMethods.Checked
            g_Settings.TestSideOfPierRead = chkSetupTestSideOfPierRead.Checked
            g_Settings.TestSideOfPierWrite = chkSetupTestSideOfPierWrite.Checked
            g_Settings.TestPerformance = chkSetupTestPerformance.Checked
            g_Settings.UpdateCheck = chkUpdates.Checked
            g_Settings.Debug = chkDebug.Checked
            g_Settings.RunAs32Bit = chkRunAs32Bit.Checked
            g_Settings.InterpretErrorMessages = chkInferStateFromErrorMessages.Checked
            g_Settings.UseDriverAccess = chkUseDriverAccess.Checked
            g_Settings.DisplayMethodCalls = chkDisplayMethodCalls.Checked

            g_Settings.CommandString = chkCommandString.Checked
            g_Settings.CommandBlind = chkCommandBlind.Checked
            g_Settings.CommandBool = chkCommandBool.Checked
            g_Settings.CommandStringRaw = chkCommandStringRaw.Checked
            g_Settings.CommandBlindRaw = chkCommandBlindRaw.Checked
            g_Settings.CommandBoolRaw = chkCommandBoolRaw.Checked

            'Dome tests
            g_Settings.DomeAltitudeTimeout = CDbl(txtDomeAltitudeTimeout.Text)
            g_Settings.DomeAzimuthTimeout = CDbl(txtDomeAzimuthTimeout.Text)
            g_Settings.DomeShutterTimeout = CDbl(txtDomeShutterTimout.Text)
            g_Settings.DomeStabilisationWait = CDbl(txtDomeStabilisationWait.Text)

            'Switch Tests
            g_Settings.SwitchReadDelay = CInt(txtSwitchReadDelay.Text)
            g_Settings.SwitchWriteDelay = CInt(txtSwitchWriteDelay.Text)
            g_Settings.ExtendedSwitchNumberTestRange = CInt(txtSwitchExtendedTestRange.Text)

            'ObservingConditions
            g_Settings.ObservingConditionsMaxRetries = CInt(UpDownObservingConditionsRetries.Value)
            g_Settings.ObservingConditionsRetryTime = CInt(UpDownObservingConditionsWaitTime.Value)

            'Camera
            g_Settings.CameraMaxBinX = CInt(numMaxBinX.Value)
            g_Settings.CameraMaxBinY = CInt(numMaxBinY.Value)

            'Special - DestinationSideOfPier
            g_Settings.FlipTestHAStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest) = CDbl(txtStartHAW.Text)
            g_Settings.FlipTestHAEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest) = txtEndHAW.Text
            g_Settings.FlipTestDECStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest) = txtStartDECW.Text
            g_Settings.FlipTestDECEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest) = txtEndDECW.Text
            g_Settings.FlipTestDECStep(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest) = txtDecStepW.Text

            g_Settings.FlipTestHAStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast) = txtStartHAE.Text
            g_Settings.FlipTestHAEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast) = txtEndHAE.Text
            g_Settings.FlipTestDECStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast) = txtStartDECE.Text
            g_Settings.FlipTestDECEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast) = txtEndDECE.Text
            g_Settings.FlipTestDECStep(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast) = txtDECStepE.Text
            g_Settings.DSOPSide = cmbDSOPSide.Text

            Dim l_TScopeValues As New Dictionary(Of String, CheckState)

            'Set Telescope method test check flags

            For Each kvp As KeyValuePair(Of String, CheckState) In g_TelescopeTestsMaster
                l_TScopeValues.Add(kvp.Key, TScopeMethodTests.GetItemCheckState(TScopeMethodTests.Items.IndexOf(kvp.Key)))
            Next

            'Update the global variable with current values
            For Each kvp As KeyValuePair(Of String, CheckState) In l_TScopeValues
                g_TelescopeTests.Item(kvp.Key) = kvp.Value
            Next

            g_Settings.TeleScopeTests() = l_TScopeValues 'persist the new values to registry

            Me.Close()
        Catch ex As Exception
            LogMsgError("BtnSetupOK_Click", $"Unexpected Conform exception: {ex}")
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BtnSetupCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetupCancel.Click
        Me.Close()
    End Sub

    Private Sub ChkSetupTestMethods_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSetupTestMethods.CheckedChanged
        If chkSetupTestMethods.Checked Then 'Has just been checked therefore set and disable the properties box
            chkSetupTestProperties.Checked = True
            chkSetupTestProperties.Enabled = False
        Else 'Has just been unset so enable the properties box
            chkSetupTestProperties.Enabled = True
        End If
    End Sub

    Sub TextChangedNumber(ByVal sender As Object)
        If IsNumeric(sender.Text) Then
            ErrorProvider1.SetError(sender, "")
        Else
            ErrorProvider1.SetError(sender, "Must be a number")
            ErrorProvider1.SetIconAlignment(sender, ErrorIconAlignment.MiddleRight)
        End If
    End Sub
    Sub TextChangedInteger(ByVal sender As Object)
        Dim int As Integer
        If Integer.TryParse(sender.Text, int) Then
            ErrorProvider1.SetError(sender, "")
        Else
            ErrorProvider1.SetError(sender, "Must be an integer")
            ErrorProvider1.SetIconAlignment(sender, ErrorIconAlignment.MiddleRight)
        End If
    End Sub
    Sub ValidateNumber(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If Not IsNumeric(sender.Text) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub TxtDomeShutterTimeout_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDomeShutterTimout.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtDomeShutterTimeout_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDomeShutterTimout.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtDomeAzimuthTimeout_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDomeAzimuthTimeout.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtDomeAzimuthTimeout_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDomeAzimuthTimeout.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtDomeAltitiudeTimeout_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDomeAltitudeTimeout.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtDomeAltitudeTimeout_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDomeAltitudeTimeout.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtDomeStabilisationWait_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDomeStabilisationWait.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtDomeStabilisationWait_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDomeStabilisationWait.Validating
        ValidateNumber(sender, e)
    End Sub

    Private Sub TxtStartHAE_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStartHAE.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtStartHAE_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtStartHAE.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtEndHAE_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEndHAE.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtEndHAE_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtEndHAE.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtStartDECE_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStartDECE.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtStartDECE_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtStartDECE.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtEndDECE_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEndDECE.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtEndDECE_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtEndDECE.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtDECStepE_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDECStepE.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtDECStepE_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDECStepE.Validating
        ValidateNumber(sender, e)
    End Sub

    Private Sub TxtStartHAW_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStartHAW.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtStartHAW_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtStartHAW.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtEndHAW_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEndHAW.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtEndHAW_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtEndHAW.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtStartDECW_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStartDECW.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtStartDECW_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtStartDECW.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtEndDECW_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEndDECW.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtEndDECW_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtEndDECW.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtDECStepW_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDecStepW.TextChanged
        TextChangedNumber(sender)
    End Sub
    Private Sub TxtDECStepW_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDecStepW.Validating
        ValidateNumber(sender, e)
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ChkSetupTestSideOfPierRead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSetupTestSideOfPierRead.CheckedChanged
        If chkSetupTestSideOfPierRead.Checked Then 'Has just been checked therefore set and disable the properties box
            chkSetupTestProperties.Checked = True
            chkSetupTestProperties.Enabled = False
            chkSetupTestMethods.Checked = True
            chkSetupTestMethods.Enabled = False
        Else 'Has just been unset so enable the methods box
            chkSetupTestMethods.Enabled = True
        End If

    End Sub
    Private Sub ChkSetupTestSideOfPierWrite_CheckedChanged(sender As Object, e As EventArgs) Handles chkSetupTestSideOfPierWrite.CheckedChanged
        If chkSetupTestSideOfPierWrite.Checked Then 'Has just been checked therefore set and disable the properties box
            chkSetupTestSideOfPierRead.Checked = True
            chkSetupTestMethods.Checked = True
            chkSetupTestProperties.Checked = True
            chkSetupTestSideOfPierRead.Enabled = False
            chkSetupTestMethods.Enabled = False
            chkSetupTestProperties.Enabled = False
        Else 'Has just been unset so enable the SideOfPier Read box
            chkSetupTestSideOfPierRead.Enabled = True
        End If

    End Sub

    Private Sub BtnClearTscopeMethods_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearTscopeMethods.Click
        For Each kvp As KeyValuePair(Of String, CheckState) In g_TelescopeTests
            TScopeMethodTests.SetItemCheckState(TScopeMethodTests.Items.IndexOf(kvp.Key), CheckState.Unchecked)
        Next
    End Sub

    Private Sub BtnSetTScopeMethods_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetTScopeMethods.Click
        For Each kvp As KeyValuePair(Of String, CheckState) In g_TelescopeTests
            TScopeMethodTests.SetItemCheckState(TScopeMethodTests.Items.IndexOf(kvp.Key), CheckState.Checked)
        Next
    End Sub

    Private Sub TxtSwitchReadDelay_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSwitchReadDelay.TextChanged
        TextChangedInteger(sender)
    End Sub
    Private Sub TxtSwitchReadDelay_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSwitchReadDelay.Validating
        ValidateNumber(sender, e)
    End Sub
    Private Sub TxtSwitchWriteDelay_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSwitchWriteDelay.TextChanged
        TextChangedInteger(sender)
    End Sub
    Private Sub TxtSwitchWriteDelay_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSwitchWriteDelay.Validating
        ValidateNumber(sender, e)
    End Sub

End Class