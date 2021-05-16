Option Strict On
Imports Microsoft.Win32
Imports ASCOM.DeviceInterface

'Class to manage state storage between Conform runs

'To add a new saved value:
'1) Decide on the variable name and its default value
'2) Create appropriately named constants similar to those below
'3) Create a property of the relevant type
'4) Create Get and Set code based on the patterns already implemented
'5) If the property is of a type not already handled,you will need to create a GetXXX function in the Utility code region

Friend Class ApplicationSettings

#Region "Constants"

    Private Const REGISTRY_CONFORM_FOLDER As String = "Software\ASCOM\Conform"
    Private Const REGISTRY_SIDEOFPIER_FOLDER As String = REGISTRY_CONFORM_FOLDER & "\Side Of Pier"
    Private Const REGISTRY_DESTINATIONSIDEOFPIER_FOLDER As String = REGISTRY_CONFORM_FOLDER & "\Destination Side Of Pier"
    Private Const REGISTRY_TELESCOPE_TESTS As String = REGISTRY_CONFORM_FOLDER & "\Telescope Tests"

    Private Const FLIP_TEST_HA_START As String = "Flip Test HA Start", FLIP_TEST_HA_START_DEFAULT As Double = 3.0
    Private Const FLIP_TEST_HA_END As String = "Flip Test HA End", FLIP_TEST_HA_END_DEFAULT As Double = -3.0
    Private Const FLIP_TEST_DEC_START As String = "Flip Test DEC Start", FLIP_TEST_DEC_START_DEFAULT As Double = 0.0
    Private Const FLIP_TEST_DEC_END As String = "Flip Test DEC End", FLIP_TEST_DEC_END_DEFAULT As Double = 0.0
    Private Const FLIP_TEST_DEC_STEP As String = "Flip Test DEC Step", FLIP_TEST_DEC_STEP_DEFAULT As Double = 10.0
    Private Const DSOP_SIDE As String = "Destination Side Of Pier Side", DSOP_SIDE_DEFAULT As String = "pierWest only"
    Private Const SOP_SIDE As String = "Side Of Pier Side", SOP_SIDE_DEFAULT As String = "pierWest only"

    Private Const COMMAND_BLIND As String = "Command Blind", COMMAND_BLIND_DEFAULT As Boolean = True
    Private Const COMMAND_BLIND_RAW As String = "Command Blind Raw", COMMAND_BLIND_RAW_DEFAULT As Boolean = False
    Private Const COMMAND_BOOL As String = "Command Bool", COMMAND_BOOL_DEFAULT As Boolean = True
    Private Const COMMAND_BOOL_RAW As String = "Command Bool Raw", COMMAND_BOOL_RAW_DEFAULT As Boolean = False
    Private Const COMMAND_STRING As String = "Command String", COMMAND_STRING_DEFAULT As Boolean = True
    Private Const COMMAND_STRING_RAW As String = "Command String Raw", COMMAND_STRING_RAW_DEFAULT As Boolean = False

    Private Const CREATE_VALIDATION_FILE As String = "Create Validation File", CREATE_VALIDATION_FILE_DEFAULT As Boolean = False
    Private Const DE_BUG As String = "Debug", DE_BUG_DEFAULT As Boolean = False

    Private Const DEVICE_CAMERA As String = "Device Camera", DEVICE_CAMERA_DEFAULT As String = "CCDSimulator.Camera"
    Private Const DEVICE_VIDEO As String = "Device Video Camera", DEVICE_VIDEO_DEFAULT As String = "ASCOM.Simulator.Video"
    Private Const DEVICE_COVER_CALIBRATOR As String = "Device Cover Calibrator", DEVICE_COVER_CALIBRATOR_DEFAULT As String = "ASCOM.Simulator.CoverCalibrator"
    Private Const DEVICE_DOME As String = "Device Dome", DEVICE_DOME_DEFAULT As String = "DomeSim.Dome"
    Private Const DEVICE_FILTER_WHEEL As String = "Device Filter Wheel", DEVICE_FILTER_WHEEL_DEFAULT As String = "FilterWheelSim.FilterWheel"
    Private Const DEVICE_FOCUSER As String = "Device Focuser", DEVICE_FOCUSER_DEFAULT As String = "FocusSim.Focuser"
    Private Const DEVICE_OBSERVINGCONDITIONS As String = "Device Observing Conditions", DEVICE_OBSERVINGCONDITIONS_DEFAULT As String = "ASCOM.OCH.ObservingConditions"
    Private Const DEVICE_ROTATOR As String = "Device Rotator", DEVICE_ROTATOR_DEFAULT As String = "ASCOM.Simulator.Rotator"
    Private Const DEVICE_SAFETY_MONITOR As String = "Device Safety Monitor", DEVICE_SAFETY_MONITOR_DEFAULT As String = "ASCOM.Simulator.SafetyMonitor"
    Private Const DEVICE_SWITCH As String = "Device Switch", DEVICE_SWITCH_DEFAULT As String = "SwitchSim.Switch"
    Private Const DEVICE_TELESCOPE As String = "Device Telescope", DEVICE_TELESCOPE_DEFAULT As String = "ScopeSim.Telescope"
    Private Const CURRENT_DEVICE_TYPE As String = "Current Device Type", CURRENT_DEVICE_TYPE_DEFAULT As DeviceType = GlobalVarsAndCode.DeviceType.Telescope

    Private Const LOG_FILES_DIRECTORY As String = "Log File Directory", LOG_FILES_DIRECTORY_DEFAULT As String = "\ASCOM"
    Private Const MESSAGE_LEVEL As String = "Message Level", MESSAGE_LEVEL_DEFAULT As MessageLevel = GlobalVarsAndCode.MessageLevel.msgInfo
    Private Const TEST_METHODS As String = "Test Methods", TEST_METHODS_DEFAULT As Boolean = True
    Private Const TEST_PERFORMANCE As String = "Test Performance", TEST_PERFORMANCE_DEFAULT As Boolean = False
    Private Const TEST_PROPERTIES As String = "Test Properties", TEST_PROPERTIES_DEFAULT As Boolean = True
    Private Const UPDATE_CHECK As String = "Update Check", UPDATE_CHECK_DEFAULT As Boolean = True
    Private Const UPDATE_CHECK_DATE As String = "Update Check Date", UPDATE_CHECK_DATE_DEFAULT As Date = #1/1/2008 1:00:00 AM#
    Private Const WARNING_MESSAGE As String = "Warning Message Platform 6", WARNING_MESSAGE_DEFAULT As Boolean = False ' Updated for Platform 6 to force a redisplay as the words have changed
    Private Const TEST_SIDEOFPIER_READ As String = "Test SideOfPier Read", TEST_SIDEOFPIER_READ_DEFAULT As Boolean = True
    Private Const TEST_SIDEOFPIER_WRITE As String = "Test SideOfPier Write", TEST_SIDEOFPIER_WRITE_DEFAULT As Boolean = True
    Private Const RUN_AS_THIRTYTWO_BITS As String = "Run As 64Bit", RUN_AS_THIRTYTWO_BITS_DEFAULT As Boolean = False
    Private Const INTERPRET_ERROR_MESSAGES As String = "Interpret Error Messages", INTERPRET_ERROR_MESSAGES_DEFAULT As Boolean = False
    Private Const SWITCH_WARNING_MESSAGE As String = "Switch Warning Message", SWITCH_WARNING_MESSAGE_DEFAULT As Boolean = False ' Updated for Platform 6 to force a redisplay as the words have changed
    Private Const USE_DRIVERACCESS As String = "Use DriverAccess", USE_DRIVERACCESS_DEFAULT As Boolean = True
    Private Const DISPLAY_METHOD_CALLS As String = "Display Method Calls", DISPLAY_METHOD_CALLS_DEFAULT As Boolean = False

    'Dome device variables
    Private Const DOME_SHUTTER As String = "Dome Shutter", DOME_SHUTTER_DEFAULT As Boolean = False
    Private Const DOME_SHUTTER_TIMEOUT As String = "Dome Shutter Timeout", DOME_SHUTTER_TMEOUT_DEFAULT As Double = 240.0 'Timeout for dome to open or close shutter
    Private Const DOME_AZIMUTH_TIMEOUT As String = "Dome Azimuth Timeout", DOME_AZIMUTH_TIMEOUT_DEFAULT As Double = 240.0 'Timeout for dome to move to azimuth
    Private Const DOME_ALTITUDE_TIMEOUT As String = "Dome Altitude Timeout", DOME_ALTITUDE_TIMEOUT_DEFAULT As Double = 240.0 'Timeout for dome to move to altitude
    Private Const DOME_STABILISATION_WAIT As String = "Dome Stabilisation Wait", DOME_STABILISATION_WAIT_DEFAULT As Double = 10.0 'Time to wait for Dome to stabilise after a command

    'Switch device variables
    Private Const SWITCH_SET As String = "Switch Set", SWITCH_SET_DEFAULT As Boolean = False
    Private Const SWITCH_READ_DELAY As String = "Switch Read Delay", SWITCH_READ_DELAY_DEFAULT As Integer = 500 ' Switch delay after a set before undertaking a read (ms)
    Private Const SWITCH_WRITE_DELAY As String = "Switch Write Delay", SWITCH_WRITE_DELAY_DEFAULT As Integer = 3000 ' Switch write delay between changing states (ms)
    Private Const EXTENDED_SWITCH_NUMBER_TEST_RANGE As String = "Extended Switch Number Test Range", EXTENDED_SWITCH_NUMBER_TEST_RANGE_DEFAULT As Integer = 100 ' Switch write delay between changing states (ms)

    'ObservingConditions device variables
    Private Const OBSERVINGCONDITIONS_RETRIES As String = "ObservingConditions Retries", OBSERVINGCONDITIONS_MAX_RETRIES_DEFAULT As Integer = 5
    Private Const OBSERVINGCONDITIONS_RETRY_TIME As String = "ObservingConditions Retry Time", OBSERVINGCONDITIONS_RETRY_TIME_DEFAULT As Integer = 1

    'Camera device variables
    Private Const CAMERA_MAX_BIN_X As String = "Camera Max Bin X", CAMERA_MAX_BIN_X_DEFAULT As Integer = 0
    Private Const CAMERA_MAX_BIN_Y As String = "Camera Max Bin Y", CAMERA_MAX_BIN_Y_DEFAULT As Integer = 0

#End Region

#Region "Variables"

    Private m_HKCU, m_SettingsKey, m_SideOfPierKey, m_DestinationSideOfPierKey, m_TelescopeTestsKey As RegistryKey

#End Region

#Region "New and Finalize"
    Sub New()
        m_HKCU = Registry.CurrentUser
        m_HKCU.CreateSubKey(REGISTRY_CONFORM_FOLDER)
        m_HKCU.CreateSubKey(REGISTRY_SIDEOFPIER_FOLDER)
        m_HKCU.CreateSubKey(REGISTRY_DESTINATIONSIDEOFPIER_FOLDER)
        m_HKCU.CreateSubKey(REGISTRY_TELESCOPE_TESTS)
        m_SettingsKey = m_HKCU.OpenSubKey(REGISTRY_CONFORM_FOLDER, True)
        m_SideOfPierKey = m_HKCU.OpenSubKey(REGISTRY_SIDEOFPIER_FOLDER, True)
        m_DestinationSideOfPierKey = m_HKCU.OpenSubKey(REGISTRY_DESTINATIONSIDEOFPIER_FOLDER, True)
        m_TelescopeTestsKey = m_HKCU.OpenSubKey(REGISTRY_TELESCOPE_TESTS, True)
    End Sub
    Protected Overrides Sub Finalize()
        m_SideOfPierKey.Flush()
        m_SideOfPierKey.Close()
        m_SideOfPierKey = Nothing
        m_DestinationSideOfPierKey.Flush()
        m_DestinationSideOfPierKey.Close()
        m_DestinationSideOfPierKey = Nothing
        m_SettingsKey.Flush()
        m_SettingsKey.Close()
        m_SettingsKey = Nothing
        m_HKCU.Flush()
        m_HKCU.Close()
        m_HKCU = Nothing
        MyBase.Finalize()
    End Sub
#End Region

#Region "Parameters"

    'Dome
    Property DomeShutter() As Boolean
        Get
            Return GetBool(DOME_SHUTTER, DOME_SHUTTER_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, DOME_SHUTTER, value.ToString)
        End Set
    End Property
    Property DomeStabilisationWait() As Double
        Get
            Return GetDouble(m_SettingsKey, DOME_STABILISATION_WAIT, DOME_STABILISATION_WAIT_DEFAULT)
        End Get
        Set(ByVal value As Double)
            SetName(m_SettingsKey, DOME_STABILISATION_WAIT, value.ToString)
        End Set
    End Property
    Property DomeShutterTimeout() As Double
        Get
            Return GetDouble(m_SettingsKey, DOME_SHUTTER_TIMEOUT, DOME_SHUTTER_TMEOUT_DEFAULT)
        End Get
        Set(ByVal value As Double)
            SetName(m_SettingsKey, DOME_SHUTTER_TIMEOUT, value.ToString)
        End Set
    End Property
    Property DomeAzimuthTimeout() As Double
        Get
            Return GetDouble(m_SettingsKey, DOME_AZIMUTH_TIMEOUT, DOME_AZIMUTH_TIMEOUT_DEFAULT)
        End Get
        Set(ByVal value As Double)
            SetName(m_SettingsKey, DOME_AZIMUTH_TIMEOUT, value.ToString)
        End Set
    End Property
    Property DomeAltitudeTimeout() As Double
        Get
            Return GetDouble(m_SettingsKey, DOME_ALTITUDE_TIMEOUT, DOME_ALTITUDE_TIMEOUT_DEFAULT)
        End Get
        Set(ByVal value As Double)
            SetName(m_SettingsKey, DOME_ALTITUDE_TIMEOUT, value.ToString)
        End Set
    End Property

    'Telescope
    Property TeleScopeTests() As Dictionary(Of String, CheckState)
        Get
            Dim RetVal As New Dictionary(Of String, CheckState)
            Dim testState As CheckState
            Dim registryValue As String

            For Each kvp As KeyValuePair(Of String, CheckState) In g_TelescopeTestsMaster
                Try
                    LogMsgDebug("TeleScopeTests", $"Retrieving key: {kvp.Key}")
                    registryValue = m_TelescopeTestsKey.GetValue(kvp.Key).ToString
                    LogMsgDebug("TeleScopeTests", $"Retrieved registry value: {registryValue} for {kvp.Key}")
                    testState = CType(System.Enum.Parse(GetType(CheckState), registryValue, True), CheckState)
                    LogMsgDebug("TeleScopeTests", $"Retrieved checked state: {testState} for {kvp.Key}")
                    RetVal.Add(kvp.Key, testState)
                Catch ex As IO.IOException 'Value doesn't exist so create it
                    LogMsgDebug("TeleScopeTests", $"IOException for key {kvp.Key}: {ex}")
                    SetName(m_TelescopeTestsKey, kvp.Key, CheckState.Checked.ToString)
                    RetVal.Add(kvp.Key, CheckState.Checked)
                Catch ex As Exception
                    LogMsgDebug("TeleScopeTests", $"Unexpected exception for key {kvp.Key}: {ex}")
                    SetName(m_TelescopeTestsKey, kvp.Key, CheckState.Checked.ToString)
                    RetVal.Add(kvp.Key, CheckState.Checked)
                End Try
            Next
            LogMsgDebug("TeleScopeTests", $"Returning {RetVal.Count} values.")
            Return RetVal
        End Get
        Set(ByVal value As Dictionary(Of String, CheckState))
            LogMsgDebug("TeleScopeTests", $"Setting {value.Count} values.")
            For Each kvp As KeyValuePair(Of String, CheckState) In value
                SetName(m_TelescopeTestsKey, kvp.Key, kvp.Value.ToString)
            Next
        End Set
    End Property
    Property TestSideOfPierRead() As Boolean
        Get
            Return GetBool(TEST_SIDEOFPIER_READ, TEST_SIDEOFPIER_READ_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, TEST_SIDEOFPIER_READ, value.ToString)
        End Set
    End Property
    Property TestSideOfPierWrite() As Boolean
        Get
            Return GetBool(TEST_SIDEOFPIER_WRITE, TEST_SIDEOFPIER_WRITE_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, TEST_SIDEOFPIER_WRITE, value.ToString)
        End Set
    End Property
    Property DSOPSide() As String
        Get
            Return GetString(DSOP_SIDE, DSOP_SIDE_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DSOP_SIDE, value)
        End Set
    End Property
    Property SOPSide() As String
        Get
            Return GetString(SOP_SIDE, SOP_SIDE_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, SOP_SIDE, value)
        End Set
    End Property

    Property FlipTestHAStart(ByVal p_TestType As SpecialTests, ByVal p_PierSide As PierSide) As Double
        Get
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast
                            Return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_START, FLIP_TEST_HA_START_DEFAULT)
                        Case PierSide.pierWest
                            Return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_START, -FLIP_TEST_HA_START_DEFAULT)
                        Case Else
                            MsgBox("FlipTest HAStart Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast
                            Return GetDouble(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_START, FLIP_TEST_HA_START_DEFAULT)
                        Case PierSide.pierWest
                            Return GetDouble(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_START, -FLIP_TEST_HA_START_DEFAULT)
                        Case Else
                            MsgBox("FlipTest HAStart Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest HAStart - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Get
        Set(ByVal value As Double)
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_START, value.ToString)
                        Case Else
                            MsgBox("FlipTest HAStart Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_START, value.ToString)
                        Case Else
                            MsgBox("FlipTest HAStart Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest HAStart - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Set
    End Property
    Property FlipTestHAEnd(ByVal p_TestType As SpecialTests, ByVal p_PierSide As PierSide) As Double
        Get
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast
                            Return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_END, FLIP_TEST_HA_END_DEFAULT)
                        Case PierSide.pierWest
                            Return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_END, -FLIP_TEST_HA_END_DEFAULT)
                        Case Else
                            MsgBox("FlipTest HAEnd Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast
                            Return GetDouble(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_END, FLIP_TEST_HA_END_DEFAULT)
                        Case PierSide.pierWest
                            Return GetDouble(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_END, -FLIP_TEST_HA_END_DEFAULT)
                        Case Else
                            MsgBox("FlipTest HAEnd Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest HAEnd - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Get
        Set(ByVal value As Double)
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_END, value.ToString)
                        Case Else
                            MsgBox("FlipTest HAEnd Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_HA_END, value.ToString)
                        Case Else
                            MsgBox("FlipTest HAEnd Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest HAEnd - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Set
    End Property
    Property FlipTestDECStart(ByVal p_TestType As SpecialTests, ByVal p_PierSide As PierSide) As Double
        Get
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            Return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_START, FLIP_TEST_DEC_START_DEFAULT)
                        Case Else
                            MsgBox("FlipTest DECStart Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            Return GetDouble(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_START, FLIP_TEST_DEC_START_DEFAULT)
                        Case Else
                            MsgBox("FlipTest DECStart Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest DECStart - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Get
        Set(ByVal value As Double)
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_START, value.ToString)
                        Case Else
                            MsgBox("FlipTest DECStart Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_START, value.ToString)
                        Case Else
                            MsgBox("FlipTest DECStart Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest DECStart - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Set
    End Property
    Property FlipTestDECEnd(ByVal p_TestType As SpecialTests, ByVal p_PierSide As PierSide) As Double
        Get
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            Return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_END, FLIP_TEST_DEC_END_DEFAULT)
                        Case Else
                            MsgBox("FlipTest DECEnd Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            Return GetDouble(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_END, FLIP_TEST_DEC_END_DEFAULT)
                        Case Else
                            MsgBox("FlipTest DECEnd Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest DECStart - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Get
        Set(ByVal value As Double)
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_END, value.ToString)
                        Case Else
                            MsgBox("FlipTest DECEnd Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_END, value.ToString)
                        Case Else
                            MsgBox("FlipTest DECEnd Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest DECEnd - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Set
    End Property
    Property FlipTestDECStep(ByVal p_TestType As SpecialTests, ByVal p_PierSide As PierSide) As Double
        Get
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            Return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_STEP, FLIP_TEST_DEC_STEP_DEFAULT)
                        Case Else
                            MsgBox("FlipTest DECStep Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            Return GetDouble(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_STEP, FLIP_TEST_DEC_STEP_DEFAULT)
                        Case Else
                            MsgBox("FlipTest DECStep Get - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest DECStep - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Get
        Set(ByVal value As Double)
            Select Case p_TestType
                Case SpecialTests.TelescopeDestinationSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_DestinationSideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_STEP, value.ToString)
                        Case Else
                            MsgBox("FlipTest DECStep Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case SpecialTests.TelescopeSideOfPier
                    Select Case p_PierSide
                        Case PierSide.pierEast, PierSide.pierWest
                            SetName(m_SideOfPierKey, p_PierSide.ToString & " " & FLIP_TEST_DEC_STEP, value.ToString)
                        Case Else
                            MsgBox("FlipTest DECStep Set - Unexpected pier side: " & p_PierSide.ToString, MsgBoxStyle.Critical)
                    End Select
                Case Else
                    MsgBox("FlipTest DECStep - Unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
        End Set
    End Property

    'Common
    Property CommandBlind() As Boolean
        Get
            Return GetBool(COMMAND_BLIND, COMMAND_BLIND_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, COMMAND_BLIND, value.ToString)
        End Set
    End Property
    Property CommandBlindRaw() As Boolean
        Get
            Return GetBool(COMMAND_BLIND_RAW, COMMAND_BLIND_RAW_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, COMMAND_BLIND_RAW, value.ToString)
        End Set
    End Property
    Property CommandBool() As Boolean
        Get
            Return GetBool(COMMAND_BOOL, COMMAND_BOOL_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, COMMAND_BOOL, value.ToString)
        End Set
    End Property
    Property CommandBoolRaw() As Boolean
        Get
            Return GetBool(COMMAND_BOOL_RAW, COMMAND_BOOL_RAW_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, COMMAND_BOOL_RAW, value.ToString)
        End Set
    End Property
    Property CommandString() As Boolean
        Get
            Return GetBool(COMMAND_STRING, COMMAND_STRING_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, COMMAND_STRING, value.ToString)
        End Set
    End Property
    Property CommandStringRaw() As Boolean
        Get
            Return GetBool(COMMAND_STRING_RAW, COMMAND_STRING_RAW_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, COMMAND_STRING_RAW, value.ToString)
        End Set
    End Property
    Property CreateValidationFile() As Boolean
        Get
            Return GetBool(CREATE_VALIDATION_FILE, CREATE_VALIDATION_FILE_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, CREATE_VALIDATION_FILE, value.ToString)
        End Set
    End Property

    'Internal
    Property CurrentDeviceType() As DeviceType
        Get
            Return GetDeviceType(CURRENT_DEVICE_TYPE, CURRENT_DEVICE_TYPE_DEFAULT)
        End Get
        Set(ByVal value As DeviceType)
            SetName(m_SettingsKey, CURRENT_DEVICE_TYPE, value.ToString)
        End Set
    End Property
    Property Debug() As Boolean
        Get
            Return GetBool(DE_BUG, DE_BUG_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, DE_BUG, value.ToString)
        End Set
    End Property
    Property DeviceCamera() As String
        Get
            Return GetString(DEVICE_CAMERA, DEVICE_CAMERA_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_CAMERA, value.ToString)
        End Set
    End Property
    Property DeviceVideo() As String
        Get
            Return GetString(DEVICE_VIDEO, DEVICE_VIDEO_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_VIDEO, value.ToString)
        End Set
    End Property
    Property DeviceCoverCalibrator() As String
        Get
            Return GetString(DEVICE_COVER_CALIBRATOR, DEVICE_COVER_CALIBRATOR_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_COVER_CALIBRATOR, value.ToString)
        End Set
    End Property
    Property DeviceDome() As String
        Get
            Return GetString(DEVICE_DOME, DEVICE_DOME_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_DOME, value.ToString)
        End Set
    End Property
    Property DeviceFilterWheel() As String
        Get
            Return GetString(DEVICE_FILTER_WHEEL, DEVICE_FILTER_WHEEL_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_FILTER_WHEEL, value.ToString)
        End Set
    End Property
    Property DeviceFocuser() As String
        Get
            Return GetString(DEVICE_FOCUSER, DEVICE_FOCUSER_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_FOCUSER, value.ToString)
        End Set
    End Property
    Property DeviceObservingConditions() As String
        Get
            Return GetString(DEVICE_OBSERVINGCONDITIONS, DEVICE_OBSERVINGCONDITIONS_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_OBSERVINGCONDITIONS, value.ToString)
        End Set
    End Property
    Property DeviceRotator() As String
        Get
            Return GetString(DEVICE_ROTATOR, DEVICE_ROTATOR_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_ROTATOR, value.ToString)
        End Set
    End Property
    Property DeviceSafetyMonitor() As String
        Get
            Return GetString(DEVICE_SAFETY_MONITOR, DEVICE_SAFETY_MONITOR_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_SAFETY_MONITOR, value.ToString)
        End Set
    End Property
    Property DeviceSwitch() As String
        Get
            Return GetString(DEVICE_SWITCH, DEVICE_SWITCH_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_SWITCH, value.ToString)
        End Set
    End Property
    Property DeviceTelescope() As String
        Get
            Return GetString(DEVICE_TELESCOPE, DEVICE_TELESCOPE_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, DEVICE_TELESCOPE, value.ToString)
        End Set
    End Property
    Property LogFileDirectory() As String
        Get
            Return GetString(LOG_FILES_DIRECTORY, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & LOG_FILES_DIRECTORY_DEFAULT)
        End Get
        Set(ByVal value As String)
            SetName(m_SettingsKey, LOG_FILES_DIRECTORY, value.ToString)
        End Set
    End Property
    Property MessageLevel() As MessageLevel
        Get
            Return GetMessageLevel(MESSAGE_LEVEL, MESSAGE_LEVEL_DEFAULT)
        End Get
        Set(ByVal value As MessageLevel)
            SetName(m_SettingsKey, MESSAGE_LEVEL, value.ToString)
        End Set
    End Property
    Property TestMethods() As Boolean
        Get
            Return GetBool(TEST_METHODS, TEST_METHODS_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, TEST_METHODS, value.ToString)
        End Set
    End Property
    Property TestPerformance() As Boolean
        Get
            Return GetBool(TEST_PERFORMANCE, TEST_PERFORMANCE_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, TEST_PERFORMANCE, value.ToString)
        End Set
    End Property
    Property TestProperties() As Boolean
        Get
            Return GetBool(TEST_PROPERTIES, TEST_PROPERTIES_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, TEST_PROPERTIES, value.ToString)
        End Set
    End Property
    Property UpdateCheck() As Boolean
        Get
            Return GetBool(UPDATE_CHECK, UPDATE_CHECK_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, UPDATE_CHECK, value.ToString)
        End Set
    End Property
    Property UpdateCheckDate() As Date
        Get
            Return GetDate(UPDATE_CHECK_DATE, UPDATE_CHECK_DATE_DEFAULT)
        End Get
        Set(ByVal value As Date)
            SetName(m_SettingsKey, UPDATE_CHECK_DATE, value.ToString)
        End Set
    End Property
    Property WarningMessage() As Boolean
        Get
            Return GetBool(WARNING_MESSAGE, WARNING_MESSAGE_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, WARNING_MESSAGE, value.ToString)
        End Set
    End Property
    Property RunAs32Bit() As Boolean
        Get
            Return GetBool(RUN_AS_THIRTYTWO_BITS, RUN_AS_THIRTYTWO_BITS_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, RUN_AS_THIRTYTWO_BITS, value.ToString)
        End Set
    End Property
    Property InterpretErrorMessages() As Boolean
        Get
            Return GetBool(INTERPRET_ERROR_MESSAGES, INTERPRET_ERROR_MESSAGES_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, INTERPRET_ERROR_MESSAGES, value.ToString)
        End Set
    End Property
    Property SwitchWarningMessage() As Boolean
        Get
            Return GetBool(SWITCH_WARNING_MESSAGE, SWITCH_WARNING_MESSAGE_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, SWITCH_WARNING_MESSAGE, value.ToString)
        End Set
    End Property
    Property UseDriverAccess() As Boolean
        Get
            Return GetBool(USE_DRIVERACCESS, USE_DRIVERACCESS_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, USE_DRIVERACCESS, value.ToString)
        End Set
    End Property
    Property DisplayMethodCalls() As Boolean
        Get
            Return GetBool(DISPLAY_METHOD_CALLS, DISPLAY_METHOD_CALLS_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, DISPLAY_METHOD_CALLS, value.ToString)
        End Set
    End Property

    ' Switch
    Property SwitchSet() As Boolean
        Get
            Return GetBool(SWITCH_SET, SWITCH_SET_DEFAULT)
        End Get
        Set(ByVal value As Boolean)
            SetName(m_SettingsKey, SWITCH_SET, value.ToString)
        End Set
    End Property
    Property SwitchReadDelay() As Integer
        Get
            Return GetInteger(SWITCH_READ_DELAY, SWITCH_READ_DELAY_DEFAULT)
        End Get
        Set(ByVal value As Integer)
            SetName(m_SettingsKey, SWITCH_READ_DELAY, value.ToString)
        End Set
    End Property
    Property SwitchWriteDelay() As Integer
        Get
            Return GetInteger(SWITCH_WRITE_DELAY, SWITCH_WRITE_DELAY_DEFAULT)
        End Get
        Set(ByVal value As Integer)
            SetName(m_SettingsKey, SWITCH_WRITE_DELAY, value.ToString)
        End Set
    End Property

    Property ExtendedSwitchNumberTestRange() As Integer
        Get
            Return GetInteger(EXTENDED_SWITCH_NUMBER_TEST_RANGE, EXTENDED_SWITCH_NUMBER_TEST_RANGE_DEFAULT)
        End Get
        Set(ByVal value As Integer)
            SetName(m_SettingsKey, EXTENDED_SWITCH_NUMBER_TEST_RANGE, value.ToString)
        End Set
    End Property

    'ObservingConditions
    Property ObservingConditionsRetryTime() As Integer
        Get
            Return GetInteger(OBSERVINGCONDITIONS_RETRY_TIME, OBSERVINGCONDITIONS_RETRY_TIME_DEFAULT)
        End Get
        Set(ByVal value As Integer)
            SetName(m_SettingsKey, OBSERVINGCONDITIONS_RETRY_TIME, value.ToString)
        End Set
    End Property
    Property ObservingConditionsMaxRetries() As Integer
        Get
            Return GetInteger(OBSERVINGCONDITIONS_RETRIES, OBSERVINGCONDITIONS_MAX_RETRIES_DEFAULT)
        End Get
        Set(ByVal value As Integer)
            SetName(m_SettingsKey, OBSERVINGCONDITIONS_RETRIES, value.ToString)
        End Set
    End Property

    'Camera
    Property CameraMaxBinX() As Integer
        Get
            Return GetInteger(CAMERA_MAX_BIN_X, CAMERA_MAX_BIN_X_DEFAULT)
        End Get
        Set(ByVal value As Integer)
            SetName(m_SettingsKey, CAMERA_MAX_BIN_X, value.ToString)
        End Set
    End Property
    Property CameraMaxBinY() As Integer
        Get
            Return GetInteger(CAMERA_MAX_BIN_Y, CAMERA_MAX_BIN_Y_DEFAULT)
        End Get
        Set(ByVal value As Integer)
            SetName(m_SettingsKey, CAMERA_MAX_BIN_Y, value.ToString)
        End Set
    End Property


#End Region

#Region "Utility Code"

    Private Overloads Function GetBool(ByVal p_Key As RegistryKey, ByVal p_Name As String, ByVal p_DefaultValue As Boolean) As Boolean
        Dim l_Value As Boolean
        Try
            If p_Key.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = CBool(p_Key.GetValue(p_Name))
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(p_Key, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetBool", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Overloads Function GetBool(ByVal p_Name As String, ByVal p_DefaultValue As Boolean) As Boolean
        Dim l_Value As Boolean
        Try
            If m_SettingsKey.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = CBool(m_SettingsKey.GetValue(p_Name))
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetBool", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Function GetString(ByVal p_Name As String, ByVal p_DefaultValue As String) As String
        Dim l_Value As String
        l_Value = ""
        Try
            If m_SettingsKey.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = m_SettingsKey.GetValue(p_Name).ToString
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetString", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Overloads Function GetInteger(ByVal p_Name As String, ByVal p_DefaultValue As Integer) As Integer
        Dim l_Value As Integer
        Try
            If m_SettingsKey.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = CInt(m_SettingsKey.GetValue(p_Name))
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetInteger", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Overloads Function GetInteger(ByVal p_Key As RegistryKey, ByVal p_Name As String, ByVal p_DefaultValue As Integer) As Integer
        Dim l_Value As Integer
        LogMsg("GetInteger", GlobalVarsAndCode.MessageLevel.msgDebug, p_Name.ToString & " " & p_DefaultValue.ToString)
        Try
            If p_Key.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = CInt(p_Key.GetValue(p_Name))
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(p_Key, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetInteger", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Function GetDouble(ByVal p_Key As RegistryKey, ByVal p_Name As String, ByVal p_DefaultValue As Double) As Double
        Dim l_Value As Double
        LogMsg("GetDouble", GlobalVarsAndCode.MessageLevel.msgDebug, p_Name.ToString & " " & p_DefaultValue.ToString)
        Try
            If p_Key.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = CDbl(p_Key.GetValue(p_Name))
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(p_Key, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetDouble", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Function GetDate(ByVal p_Name As String, ByVal p_DefaultValue As Date) As Date
        Dim l_Value As Date
        Try
            If m_SettingsKey.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = CDate(m_SettingsKey.GetValue(p_Name))
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetDate", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Function GetDeviceType(ByVal p_Name As String, ByVal p_DefaultValue As DeviceType) As DeviceType
        Dim l_Value As DeviceType
        Try
            If m_SettingsKey.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = CType(System.Enum.Parse(GetType(DeviceType), m_SettingsKey.GetValue(p_Name).ToString, True), DeviceType)
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetDeviceType", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Function GetMessageLevel(ByVal p_Name As String, ByVal p_DefaultValue As MessageLevel) As MessageLevel
        Dim l_Value As MessageLevel
        Try
            If m_SettingsKey.GetValueKind(p_Name) = RegistryValueKind.String Then ' Value does exist
                l_Value = CType(System.Enum.Parse(GetType(MessageLevel), m_SettingsKey.GetValue(p_Name).ToString, True), MessageLevel)
            End If
        Catch ex As System.IO.IOException 'Value doesn't exist so create it
            SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString)
            l_Value = p_DefaultValue
        Catch ex As Exception
            LogMsg("GetMessageLevel", GlobalVarsAndCode.MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
            l_Value = p_DefaultValue
        End Try
        Return l_Value
    End Function
    Private Sub SetName(ByVal p_Key As RegistryKey, ByVal p_Name As String, ByVal p_Value As String)
        p_Key.SetValue(p_Name, p_Value.ToString, RegistryValueKind.String)
        p_Key.Flush()
    End Sub
#End Region

End Class
