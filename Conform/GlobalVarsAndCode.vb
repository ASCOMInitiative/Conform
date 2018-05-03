'Option Strict On
Imports ASCOM.Utilities
Imports System.Threading

Module GlobalVarsAndCode

#Region "Enums"
    Friend Enum MessageLevel
        msgNone = 0
        msgDebug = 1
        msgComment = 2
        msgInfo = 3
        msgOK = 4
        msgWarning = 5
        msgIssue = 6
        msgError = 7
        msgAlways = 8
    End Enum
    Friend Enum StatusType
        staTest = 1
        staAction = 2
        staStatus = 3
    End Enum
    Friend Enum DeviceType
        Telescope = 0
        Camera = 1
        Dome = 2
        FilterWheel = 3
        Focuser = 4
        ObservingConditions = 5
        Rotator = 6
        Switch = 7
        SafetyMonitor = 8
        Video = 9
    End Enum
    Enum SpecialTests
        TelescopeSideOfPier
        TelescopeDestinationSideOfPier
        TelescopeSideOfPierAnalysis
        TelescopeCommands
    End Enum
    Friend Enum MandatoryMethod
        Connected = 0
        Description = 1
        DriverInfo = 2
        DriverVersion = 3
        Name = 4
        CommandXXX = 5
    End Enum
#End Region

#Region "Constants"
    Friend Const DRIVER_PROGID As String = "Driver ProgID: "
    Friend Const NOT_IMP_NET As String = ".NET - Feature not implemented"
    Friend Const NOT_IMP_COM As String = "COM - Feature not implemented"
    'Const INV_VAL_NET As String = ".NET - Invalid value: "
    Friend Const EX_DRV_NET As String = ".NET - Driver Exception: "
    Friend Const EX_NET As String = ".NET - Exception: "
    Friend Const EX_COM As String = "COM - Exception: "
    Friend Const PERF_LOOP_TIME As Integer = 5 'Performance loop run time in seconds
    Friend Const SLEEP_TIME As Integer = 200 'Loop time for testing wherther slewing has completed
    Friend Const CAMERA_SLEEP_TIME As Integer = 10 'Loop time for testing wherther camera events have completed
    Friend Const DEVICE_DESTROY_WAIT = 500 'Time to wait after destroying a device before continuing

    'TelescopeTest Constants
    Friend Const TELTEST_ABORT_SLEW As String = "AbortSlew"
    Friend Const TELTEST_AXIS_RATE As String = "AxisRate"
    Friend Const TELTEST_CAN_MOVE_AXIS As String = "CanMoveAxis"
    Friend Const TELTEST_COMMANDXXX As String = "CommandXXX"
    Friend Const TELTEST_DESTINATION_SIDE_OF_PIER As String = "DestinationSideOfPier"
    Friend Const TELTEST_FIND_HOME As String = "FindHome"
    Friend Const TELTEST_MOVE_AXIS As String = "MoveAxis"
    Friend Const TELTEST_PARK_UNPARK As String = "Park/Unpark"
    Friend Const TELTEST_PULSE_GUIDE As String = "PulseGuide"
    Friend Const TELTEST_SLEW_TO_ALTAZ As String = "SlewToAltAz"
    Friend Const TELTEST_SLEW_TO_ALTAZ_ASYNC As String = "SlewToAltAzAsync"
    Friend Const TELTEST_SLEW_TO_TARGET As String = "SlewToTarget"
    Friend Const TELTEST_SLEW_TO_TARGET_ASYNC As String = "SlewToTargetAsync"
    Friend Const TELTEST_SYNC_TO_ALTAZ As String = "SyncToAltAz"
    Friend Const TELTEST_SLEW_TO_COORDINATES As String = "SlewToCoordinates"
    Friend Const TELTEST_SLEW_TO_COORDINATES_ASYNC As String = "SlewToCoordinatesAsync"
    Friend Const TELTEST_SYNC_TO_COORDINATES As String = "SyncToCoordinates"
    Friend Const TELTEST_SYNC_TO_TARGET As String = "SyncToTarget"

    Friend Const FWTEST_IS_MOVING As Integer = -1
    Friend Const FWTEST_TIMEOUT As Integer = 30

#End Region

#Region "Global Variables"
    Friend g_LogLevel As MessageLevel
    Friend g_CountError, g_CountWarning, g_CountIssue As Integer
    Friend g_LogFile As System.IO.StreamWriter
    Friend g_ValidationTempFile As System.IO.StreamWriter
    Friend g_CmdStrings As ConformCommandStrings, g_CmdStringsRaw As ConformCommandStrings
    Friend g_DeviceObject As Object
    Friend g_Settings As ApplicationSettings
    Friend g_InterfaceVersion As Integer ' Variable to held interface version of the current device

    'Exception number variables
    Friend g_ExNotImplemented, g_ExNotSet1, g_ExNotSet2 As Integer
    Friend g_ExInvalidValue1, g_ExInvalidValue2, g_ExInvalidValue3, g_ExInvalidValue4, g_ExInvalidValue5, g_ExInvalidValue6 As Integer

    Friend g_Stop As Boolean = False

    'Helper variables
    Friend g_Util As ASCOM.Utilities.Util
    'Friend g_Util As ASCOM.Utilities.Util

    'Device ProgID variables
    Friend g_SafetyMonitorProgID As String
    Friend g_SwitchProgID As String
    Friend g_FilterWheelProgID As String
    Friend g_FocuserProgID As String
    Friend g_RotatorProgID As String
    Friend g_CameraProgID As String
    Friend g_VideoCameraProgID As String
    Friend g_DomeProgID As String
    Friend g_CurrentProgID As String
    Friend g_TelescopeProgID As String
    Friend g_ObservingConditionsProgID As String

    'Status update class
    Friend g_Status As Stat = New Stat

    Friend g_TelescopeTests As New Generic.Dictionary(Of String, CheckState)

    Friend Sub SetTelescopeTestOptions()
        'Set up list of Telescope Tests called from main and setup forms onload events
        g_TelescopeTests.Add(TELTEST_CAN_MOVE_AXIS, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_PARK_UNPARK, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_ABORT_SLEW, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_AXIS_RATE, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_COMMANDXXX, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_FIND_HOME, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_MOVE_AXIS, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_PULSE_GUIDE, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SLEW_TO_COORDINATES, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SLEW_TO_COORDINATES_ASYNC, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SLEW_TO_TARGET, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SLEW_TO_TARGET_ASYNC, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_DESTINATION_SIDE_OF_PIER, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SLEW_TO_ALTAZ, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SLEW_TO_ALTAZ_ASYNC, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SYNC_TO_COORDINATES, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SYNC_TO_TARGET, CheckState.Checked)
        g_TelescopeTests.Add(TELTEST_SYNC_TO_ALTAZ, CheckState.Checked)
        g_TelescopeTests = g_Settings.TeleScopeTests() 'Get actual values from registry
    End Sub

#End Region

#Region "Code"
    Friend Function TestStop() As Boolean
        Application.DoEvents()
        Thread.Sleep(10)
        Application.DoEvents()
        Thread.Sleep(10)
        Application.DoEvents()
        'Thread.Sleep(10)
        'Application.DoEvents()
        'Thread.Sleep(10)
        'Application.DoEvents()
        'Thread.Sleep(10)
        'Application.DoEvents()
        'Thread.Sleep(10)
        'Application.DoEvents()
        'Thread.Sleep(10)
        'Application.DoEvents()
        'Thread.Sleep(10)
        TestStop = g_Stop
    End Function

    ''' <summary>
    ''' Delays execution for the given time period in milliseconds
    ''' </summary>
    ''' <param name="p_Duration">Delay duration in milliseconds</param>
    ''' <remarks></remarks>
    Friend Sub WaitFor(ByVal p_Duration As Integer)
        Dim l_StartTime As Date, WaitDuration As Integer
        WaitDuration = CInt(p_Duration / 100)
        If WaitDuration > SLEEP_TIME Then WaitDuration = SLEEP_TIME
        If WaitDuration < 1 Then WaitDuration = 1
        'Wait for p_Duration milliseconds
        l_StartTime = Now 'Save starttime
        Do
            System.Threading.Thread.Sleep(WaitDuration)
            Application.DoEvents()
        Loop Until (Now.Subtract(l_StartTime).TotalMilliseconds > p_Duration) Or TestStop()
    End Sub

    Friend Sub WaitForAbsolute(ByVal p_Duration As Integer, ByVal p_Message As String)
        LogMsg("WaitForAbsolute", MessageLevel.msgDebug, p_Duration & " " & p_Message)
        For i As Integer = 0 To CInt(p_Duration / 100)
            System.Threading.Thread.Sleep(100)
            Application.DoEvents()
            SetStatus(p_Message, (((p_Duration / 100) - i) / 10).ToString, "")
        Next
        SetStatus("", "", "")

    End Sub

    Friend Sub LogMsgOK(ByVal p_Test As String, ByVal p_Msg As String)
        LogMsg(p_Test, MessageLevel.msgOK, p_Msg)
    End Sub

    Friend Sub LogMsgInfo(ByVal p_Test As String, ByVal p_Msg As String)
        LogMsg(p_Test, MessageLevel.msgInfo, p_Msg)
    End Sub

    Friend Sub LogMsgIssue(ByVal p_Test As String, ByVal p_Msg As String)
        LogMsg(p_Test, MessageLevel.msgIssue, p_Msg)
    End Sub
    Friend Sub LogMsgError(ByVal p_Test As String, ByVal p_Msg As String)
        LogMsg(p_Test, MessageLevel.msgError, p_Msg)
    End Sub
    Friend Sub LogMsgWarning(ByVal p_Test As String, ByVal p_Msg As String)
        LogMsg(p_Test, MessageLevel.msgWarning, p_Msg)
    End Sub

    Friend Sub LogMsg(ByVal p_Test As String, ByVal p_MsgLevel As MessageLevel, ByVal p_Msg As String)
        Const c_Spacing As Integer = 52, TEST_NAME_WIDTH As Integer = 34
        Dim l_MsgLevelFormatted, l_Msg, l_TestFormatted, l_MsgFormatted As String, i, j As Integer
        l_Msg = ""
        Try
            If p_MsgLevel >= g_LogLevel Then
                l_TestFormatted = p_Test.PadRight(TEST_NAME_WIDTH)
                l_TestFormatted = Left(l_TestFormatted, TEST_NAME_WIDTH)
                l_MsgLevelFormatted = "        "
                i = 1
                l_MsgFormatted = p_Msg

                ' Remove CRLF from the message text
                j = InStr(i, l_MsgFormatted, vbCrLf)
                Do While j > 0
                    l_MsgFormatted = Microsoft.VisualBasic.Left(l_MsgFormatted, j + 1) & StrDup(c_Spacing, " ") & Mid(l_MsgFormatted, j + 2)
                    i = j + c_Spacing + 2
                    j = InStr(i, l_MsgFormatted, vbCrLf)
                Loop

                Select Case p_MsgLevel
                    Case MessageLevel.msgNone
                        l_MsgLevelFormatted = "        "
                    Case MessageLevel.msgDebug
                        l_MsgLevelFormatted = "DEBUG   "
                    Case MessageLevel.msgComment
                        l_MsgLevelFormatted = "        "
                    Case MessageLevel.msgInfo
                        l_MsgLevelFormatted = "INFO    "
                    Case MessageLevel.msgOK
                        l_MsgLevelFormatted = "OK      "
                    Case MessageLevel.msgWarning
                        l_MsgLevelFormatted = "WARNING "
                        g_CountWarning += 1
                    Case MessageLevel.msgIssue
                        l_MsgLevelFormatted = "ISSUE   "
                        g_CountIssue += 1
                    Case MessageLevel.msgError
                        l_MsgLevelFormatted = "ERROR   "
                        g_CountError += 1
                    Case MessageLevel.msgAlways
                        l_MsgLevelFormatted = "        "
                    Case Else
                        MsgBox("Conform:LogMsg - UNEXPECTED LogMessageType: " & p_MsgLevel.ToString & " " & p_Msg)
                End Select

                If FrmConformMain.txtStatus.TextLength > 50000 Then ' Limit lines to a sensible number
                    FrmConformMain.txtStatus.Text = Microsoft.VisualBasic.Right(FrmConformMain.txtStatus.Text, 28000)
                End If
                Select Case p_MsgLevel
                    Case MessageLevel.msgAlways
                        l_Msg = p_Test & " " & p_Msg
                    Case Else
                        l_Msg = Format(Now(), "HH:mm:ss.fff") & " " & l_TestFormatted & l_MsgLevelFormatted & " " & l_MsgFormatted
                End Select
                FrmConformMain.txtStatus.AppendText(l_Msg & vbCrLf)
                FrmConformMain.txtStatus.SelectionStart = Len(FrmConformMain.txtStatus.Text)
                'Me.txtStatus.Refresh()
                FrmConformMain.Focus()
                If Not g_LogFile Is Nothing Then
                    g_LogFile.WriteLine(l_Msg) 'Update log file
                End If
                If Not (g_ValidationTempFile Is Nothing) Then g_ValidationTempFile.WriteLine(l_Msg) 'Update validation file
            End If
        Catch ex As Exception
            'MsgBox(Len(l_Msg) & " *" & l_Msg & "* " & ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    ''' <summary>
    ''' Update the status display
    ''' </summary>
    ''' <param name="Status">Type of message to set, Test, Action or Status</param>
    ''' <param name="Message">Message text</param>
    ''' <remarks></remarks>
    Friend Sub Status(ByVal Status As StatusType, ByVal Message As String)
        Select Case Status
            Case StatusType.staTest
                FrmConformMain.staTest.Text = Message
            Case StatusType.staAction
                FrmConformMain.staAction.Text = Message
            Case StatusType.staStatus
                FrmConformMain.staStatus.Text = Message
        End Select
        Application.DoEvents()
    End Sub

    ''' <summary>
    ''' Set the test, action and status in one call
    ''' </summary>
    ''' <param name="Test">Name of the test being conducted</param>
    ''' <param name="Action">Specific action within the test</param>
    ''' <param name="Status">Status of the action</param>
    ''' <remarks></remarks>
    Sub SetStatus(ByVal Test As String, ByVal Action As String, ByVal Status As String)
        FrmConformMain.staTest.Text = Test
        FrmConformMain.staAction.Text = Action
        FrmConformMain.staStatus.Text = Status
    End Sub
    Friend Sub ExTest(ByVal p_TestName As String, ByVal p_EXMessage As String, ByVal p_LogMessage As String)
        Dim l_Msg As String
        l_Msg = UCase(p_EXMessage)
        If ((l_Msg.Contains("NOT") Or l_Msg.Contains("DOESN'T")) And (l_Msg.Contains("SET") Or l_Msg.Contains("IMPLEMENTED") Or l_Msg.Contains("SUPPORTED") Or l_Msg.Contains("PRESENT"))) Or l_Msg.Contains("INVALID") Or l_Msg.Contains("SUPPORT") Then
            '3.0.0.12 - removed next two lines and added third to make this an OK message
            If g_Settings.InterpretErrorMessages Then ' We are interpreting error messages so report this as OK (This was Conform's behaviour prior to 6.0.0.37)
                LogMsg(p_TestName, MessageLevel.msgOK, p_LogMessage)
            Else ' As of v6.0.0.37 default behaviour is not to interpret error messages
                LogMsg(p_TestName, MessageLevel.msgInfo, "The following Issue can be changed to OK by setting ""Interpret error messages"" in Conform's setup dialogue")
                LogMsg(p_TestName, MessageLevel.msgIssue, p_LogMessage)
            End If
        Else
            LogMsg(p_TestName, MessageLevel.msgIssue, p_LogMessage)
        End If
    End Sub

    Friend Sub ChangeBitnessIfNeeded()
        Dim P As Process, LogicString As String

        If g_Settings Is Nothing Then g_Settings = New ApplicationSettings 'Populate the setting variable

        'By default we start the MSIL version of Conform so it is 32bit on a 32bit OS and 64bit on a 64bit OS
        'Here we check if the user actually wants to run the 32bit version of Conform on a 64bit OS and start this if required
        'If g_Settings.Debug Then MsgBox("Conform: New " & IntPtr.Size & " " & g_Settings.RunAs32Bit & " " & Application.StartupPath & "\Conform32.exe")

        If VersionCode.OSBits = Utilities.Bitness.Bits64 Then 'Only test for a bitness switch if we are on a 64bt OS
            LogicString = g_Settings.RunAs32Bit.ToString & (IntPtr.Size = 4).ToString
            Select Case LogicString
                Case "TrueFalse" 'Require 32bit but running as 64bit
                    If Application.StartupPath.Contains("Visual Studio") Then
                        P = Process.Start("..\x86\Release\Conform.exe") 'Start the 32bit version of Conform
                    Else
                        P = Process.Start(Application.StartupPath & "\Conform32.exe") 'Start the 32bit version of Conform
                    End If
                    'If g_Settings.Debug Then MsgBox("Ending 64bit Application! " & (P Is Nothing))
                    End 'Close this 64bit version before any of its UI appears!
                Case "FalseTrue" 'Require 64bit but running as 32bit
                    If Application.StartupPath.Contains("Visual Studio") Then
                        P = Process.Start("..\Release\Conform.exe") 'Start the 64bit version of Conform
                    Else
                        P = Process.Start(Application.StartupPath & "\Conform.exe") 'Start the 64bit version of Conform
                    End If
                    'If g_Settings.Debug Then MsgBox("Ending 64bit Application! " & (P Is Nothing))
                    End 'Close this 64bit version before any of its UI appears!
                Case Else 'No action because required and actual are the same

            End Select

        End If
    End Sub
#End Region

End Module