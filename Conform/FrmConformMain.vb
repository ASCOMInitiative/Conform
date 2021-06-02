'Conformance level rules:
'ERRORS are unexpected behaviours from the application
'ISSUES are deviations from the written ASCOM specification
'WARNING is used for unreal values such as DEC > 90 degrees or RA > 24 hours
'INFO is used for reporting values that are legal but deviate from a tight definition of what is expected
'OK is used when values are within a tight definition of expected behaviour

#Region "Platform 5 Changes"
'Version 1.0.3
'Added a STOP button and altered the sidereal time check to stop the app and not to close it
'Added STOP button detection to all internal loops and inside long running tests
'Added Setup dialogue with log file directory and selection of tests properties, methods and performance
'Tidied up RA and DEC precision checks so the rules are located in one place for easy maintenance
'Improved error trapping for optional properties so that these are flagged as OK when not implemented
'Replaced About box with tidier version

'Version 1.0.4
'Corrected sync fail messages
'Tidied code reduce repetition and increase consistency
'Revised messages to reduce impact on customers while retaining usefulness for developers

'Version 1.05
'Made error messages more consistent with conformance level rules
'Corrected CanMoveAxis messages
'Changed TargetDeclination test value from 60.0 to 1.0 so it will be reachable from the whole planet
'Changed TargetRightAscension test value so that it is based on current sidereal time to ensure a reachable target is used

'Version 1.0.6
'Added support for reporting .NET DriverException errors
'Added support to ensure correct reporting where authors have chosen different error numbers to report the same error
'Fixed bug that would sometimes create an invalid RA to which to slew the scope
'Added test for FindHome parking the scope as well as finding home. Parking is not stated as an outcome of finding home.
'Added extra error checking around Chooser.DeviceTypeV in select driver code 
'Switched over to late binding for Telescope object to improve compatibility with platform 4 drivers

'Version 1.0.7
'Added support for Celestron driver error numbers
'Added checks to ensure performance tests are only carried out when appropriate
'Conform version number now written to log file
'Added object clean up for Chooser object
'Added reporting of mount UTCDate

'Version 1.0.8
'Added support for AltAz methods
'Added reporting of secondary and tertiary axis rates
'Added a new message if an empty tracking rate is returned for any axis
'Reverted TrackingRates and AxisRates variables from objects to AxisRates and ITrackingRates types.

'Version 1.0.9
'Added support for TemmaLite error numbers
'Fixed bugs in reporting and testing of syncXX commands introduced in 1.0.8
'Extended error reporting in SlewToAltAz
'Fixed SlewToAltAzAsync test was actually calling SlewToAltAz
'Partial implementation of testing for conditions by checking text of error messages

'Version 1.0.10
'Fixed formatting of start-up PC UTCdate to match scope UTCdate
'Added a specific test for whether the Telescope object exactly matches ASCOM.DeviceInterface.ITelescope
'Added reporting of whether the Telescope driver is COM or .NET based
'Added reporting of supported interface types for .NET drivers
'Fixed reporting of Altitude differences that were being reported as HH:MM:SS
'Fixed logging of PierSide Write - east and west were reversed in the test name
'Added support for POTH error numbers 
'Added test to confirm that MoveAxis generates an exception when parked
'Added MoveAxis tests per Telescope specification
'Added One time warning about responsibility for mount movement
'Increased checks for stop button to make the tests stop more quickly
'Added framework to support ASCOM device types other than Telescope
'Added Focuser conformance tests
'Added trap to ensure that only one test can run at once
'Added Filter Wheel conformance tests

'Version 1.0.11.0
'Fixed Unparking of mount before telescope test
'Added Automatic periodic update check by Internet (can be turned off)
'Added Manual update check option to menu
'Added Tracking turned off at end of telescope test for mount safety
'Added ASCOM icon to setup form
'Redesigned options setup screen
'Conform is now installed to C:\Program Files\ASCOM rather than C:\Program Files\Common Files\ASCOM 
'Cleans up old install directory and uninstall shortcut

'Version 1.0.11.8
'Added camera conformance tests
'Added AxisRate checks for overlaps, negative values, duplicate rates and rate maximum < minimum
'Added missing movement stop commands in two places in MoveAxis tests. Issue only showed when CansetTracking was false
'Added MountSafety message about enabling scope tracking
'Added test to confirm that UnPark generates a not implemented exception when CanUnPark is false
'Added missing test for UnPark when CanPark is false
'Added MoveAxis test to confirm that a zero rate is valid even if it isn't presented in AxisRates
'Fixed Vista issue by setting CPU flag from Any to X86 to force 32 bit operation on Vista
'Fixed the version update check so that it will work through an authenticating proxy server
'Fixed Slew and Sync tests so that they do not try and change the tracking state if CanSetTracking is false
'Fixed UTCDate Write so that it is permissible to throw a property not implemented exception
'Fixed IsPulseGuiding where PropertyNotFound exception was incorrectly accepted as valid when CanPulseGuide was true
'Fixed RightAscensionRate and DeclinationRate where PropertyNotFound exceptions were incorrectly accepted as valid - property should always return a number
'Fixed Azimuth test that didn't accept PropertyNotFound as valid
'Fixed CanReadSideOfPier test to ensure a value of false is used internally if an exception is generated on read
'Fixed SideOfPier to suppress performance test if SideOfPier cannot be read
'Fixed value reporting issue in MoveAxis tests when attempting to move at greater than the maximum rate
'Fixed SideOfPier write test to ensure that slewing is false before continuing with other tests
'Changed ScopeSafety references to MountSafety for consistency

'Version 1.0.11.9
'Fixed StopExposure and AbortExposure so that exceptions are OK if relevant Can properties are false

'Version 1.0.11.10
'Fixed error in synchronous pulse guide test when IsPulseGuiding property is not supported
'Changed Bin exception tests to check for errors on StartExposure

'Version 1.0.11.11
'Added Switch conformance checks
'Added Focuser performance tests
'Added Rotator conformance checks
'Added Dome conformance checks
'Restored camera BinX and BinY bad values tests removed in 1.0.11.10
'Fixed camera negative duration test to use negative duration not negative BinX!
'Fixed camera description to properly display description

'Version 1.0.11.12
'SideOfPier behaviour updated to match revised wording in latest telescope programmer's reference

'Version 1.0.11.13
'IConform interface added plus checks in Conform to use values returned

'Version 1.0.11.14
'Signed Conform.exe and IConform.dll

'Version 1.0.11.15
'Added ability in Conform setup UI to skip command tests
'For Camera, added tests for accessing driver through late binding, early binding and client toolkit
'Moved to VS2008 SP1

'Version 1.0.11.16
'Fixed platform version test to disable action buttons if not running on ASCOM Platform 5
'Refactored conformance checks to create consistency between device types and remove duplicate code
'Added performance tests for filter wheel

'Version 1.0.11.17
'Added check to warn user that the required level of client access toolkit is 1.0.5

'Version 1.0.11.18
'Created versions bound to client toolkit 1.0.4 to test policy redirection

'Version 1.0.11.19
'Created versions bound to client toolkit 1.0.5 to test policy redirection

'Version 2.0.0.0
'Added validation code beta 
'FilterWheel - added check for number of filter offsets and filter names being the same

'Version 2.0.2.0
'Fixed issue with identification of current driver ProgID

'Version 2.1.0.0
'Added Safety monitor device tests
'Changed IConform checks to only happen for Telescope and Dome drivers

'Version 2.1.2.0
'Refactored device tests into OO style to aid adding new device types
'Fully globalised search directory name for DriverAccess.dll
'Migrated configuration from My.Settings to HKCU registry hive
'Increased resilience to driver creation failures

'Version 2.1.4.0
'Moved telescope SideOfPier write test to methods section
'Altered destination side of pier test to check at HA +-3 hours rather than +-1 hour to avoid issues on mounts with large flip zone
'Hid debug tests menu items
'Test version for Chris Rowland to review

'Version 2.1.5
'Fixed issues in Dome shutter open and close tests when driver flags no support for these commands
'
'Version 2.1.6
'Added extra debug to DriverAccess version checking and changed l_V3 algorithm to fix internationalisation issue
'Added Marshal.Release for Util and Util2 objects

'Version 2.1.7
'Made DriverAccess testing more robust through use of reflection to in assembly version number checking
'Made Form load error checking more granular to aid debugging

'Version 2.1.8 - ***** Last Platform 5 Release *****
'Set safety monitor default to ASCOM.Simulator.SafetyMonitor
'Moved to new master interface definition in preparation for Platform 6

'Version 3.0.0 - *** First release for new platform *****
'Removed references to COM VB6 helpers and replaced with ASCOM.Utilities reference

'Version 3.0.0.1
'Added extra check to ensure that dispose is called if an error occurs in the top level test controller
'Moved log files to My Documents\ASCOM folder
'Added support for new ASCOM.InvalidValueException
'Added support for new ASCOM.ValueNotSetException
'Removed an extraneous info message from RightAscensionRate Read when we receive a not implemented exception
'Removed an extraneous info message from DeclinationRate Read when we receive a not implemented exception

'Version 3.0.0.2
'Removed built in IConform interface and moved this to ASCOM.Utilities
'Copied command strings to a local object instead of relying on an object that was discarded
'Added support for new not implemented error codes
'Improved error reporting in command string tests

'Version 3.0.0.3
'Provided more move test status information
'Made telescope method tests individually selectable through setup screen

'Version 3.0.0.4
'Further improved granularity of move axis status messages
'Moved to IConform from ASCOM.Utilities to ASCOM.IConform

'Version 3.0.0.5
'Enhanced SideOfPier tests to allow for optional DestinationSideOfPier being not implemented
'Restored performance tests that were not running when requested

'Version 3.0.0.6
'Reverted to MasterInterfaces v1.0.0.0
'Removed SafetyMonitor support
'Changed Camera CheckAccessibility to use Camera rather than ICamera to allow access to Dispose

'Version 3.0.0.7
'Moved to ASCOM exceptions v1.1

'Version 3.0.0.8 - Released
'Moved back to MSIL

'Version 3.0.0.9
'Made messages on starting device configuration debug level rather than info level
'Converted to Utilities 5.5.0.0

'Version 3.0.0.10 - Public Release on ASCOM Web Site

'Version 3.0.0.11 - Released
'FilterWheel - Removed parameter tests from Names() and Offsets() as the interface does not support these
'Updated SideOfPier test document to v1.1 to correct caption's swapped on figures 4 and 5
'Removed Switch tests
'Removed duplicate Utilities component creation, left over from VB6 migration to new .NET Utilities
'Up-levelled platform version check from 5.0 to 5.5 to ensure use only on platform 5.5 or above

'Version 3.0.0.12
'Exceptions - Exceptions that use words like "not supported" are accepted as "not implemented" exceptions
'Focuser - Now restores the original TempComp value
'Focuser - Added delay between successive create attempts to allow previous instances to clean up properly
'Focuser - Added missing test that Move should throw an exception when TempComp is set to True
'Telescope - Added test for ability to set tracking before setting it in three places
'Start-up - Fixed internationalisation issue when testing for platform version. Decimal separator of "," caused an exception

'Version 3.0.0.14
'Telescope - Added test to skip SideOfPier tests if SideOfPier Read is not supported

'Version 3.0.0.15
'Focuser - Will now accept .NET InvalidValueException thrown by Move when TempComp is true

'Version 3.0.0.16
'Camera - Fixed bug that flagged all attempts at early binding as errors
'Focuser - Added ASCOM.InvalidOperationException as OK if attempt to move focuser when temperature compensation is active
'Installer - migrated to Inno 5.3.7 Unicode
'Update - Corrected download location URL in "A new version is available" message

'Version 3.0.0.17
'Dome - Fixed exception when making dome safe after tests and CanPark or CanSetShutter were False
'Installer - Moved to 5.3.7 Unicode

'Version 3.0.0.18
'Platform 6 Changed to using Client Toolkit and new device interface versions
'Added tests for IASCOMDriver and IDeviceControl

#End Region

#Region "Platform 6 Changes"
'Version 6.0.0.38
'All         Added tests for Platform 5 and Platform 6 interfaces
'All         Now uses DriverAccess to access devices
'All         Can now configure Conform to run as 32bit on a 64bit OS 
'All         Moved the Debug option to the "General" options tab
'All         Will now report non-string types returned in the SupportedActions ArrayList as errors
'All         Places Conform version number in the Profile store Platform key
'All         The "first time" safety warning message has been updated to encompass all device types not just telescope
'All         Conform no longer attempts to interpret state from the returned error message. The original behaviour can be enabled through a configuration option
'Camera      Added test for description > 68 characters to ensure compatibility with FITS headers
'Camera      Fixed bug that could fail LastExposure tests for in-process camera drivers
'Camera      Widened tolerance on LastExposure check; within 2% of expected is now OK
'Camera      The ASCOM Camera specification was revised to state that StopExposure should not throw an exception if there is no exposure is in progress. Conform now tests for this revised behaviour.
'Camera      SetCCDTemperature now reports the ability to set temperature outside the range -280C to +100C as Issues rather than Errors
'FilterWheel Added pre-run check to ensure the filter wheel is stationary before tests start
'Telescope   Added InvalidOperation exception as OK when target declination and target right ascension are read without first having been set
'Telescope   Improved diagnostics if TrackingRates is not correctly implemented.
'Telescope   Fixed incorrect message when returned RA is out of valid range
'Telescope   Added extra test to check whether commands can unpark the scope when they should not.
'Telescope   Added test for IsPulseGuiding stuck on
'Telescope   Fixed issue in SiderealTime test when PC and scope clocks are close to 00:00:00
'Telescope   Moved time check to pre-run checks section
'Telescope   Added support for testing Telescope Interface V1 drivers
'Telescope   Fixed SyncToAltAz so that it won't try to use Azimuth and Altitude properties if these are not implemented
'Telescope   Added better reporting if a driver returns a null object for AxisRates(Axis)
'Telescope   Failure to retain tracking state after MoveAxis was shown as OK, now shows as an Issue

'Version 6.0.0.39
'All         Added test for System.InvalidOperationException, now provides message advising to use ASCOM.InvalidOperationException
'All         Removed infinite loop bug in releasing .NET in-process drivers

'Version 6.0.0.40
'Telescope   Added support for ServoCAT.Telescope invalid value COM exception number
'Telescope   Added extra debug information for SyncToCoordinates and SyncToTarget

'Version 6.0.0.42
'Telescope   AxisRates: Improved testing of enumerator methods and associated debug diagnostics
'Telescope   Corrected MoveAxis test names in Conform display
'Telescope   TrackingRates: Improved debug level diagnostic messages

'Version 6.0.0.43
'Telescope   Improved clean-up of axis rates objects in MoveAxis test

'Version 6.0.0.44
'Telescope   Allow for possibility that TargetRA and TargetDEC may throw not implemented exceptions when testing SlewToTarget
'Telescope   Allow for possibility that TargetRA and TargetDEC may throw not implemented exceptions when testing SyncToTarget

'Version 6.0.0.45
'Rotator     Rotator.Halt was incorrectly expecting a PropertyNotImplementedException, it now checks for MethodNotImplementedException
'Camera      Allow for BinX and BinY to throw InvalidValueExceptions within the range 1 to MaxBin if these bin values are not supported.

'Version 6.0.0.46
'Camera      Added tests for the new ICameraV2 properties

'Version 6.0.0.47
'Camera      Fixed crash when ReadoutModes was not present

'Version 6.0.0.48
'Camera      Revised testing of ICameraV2.BayerOffset to match specification
'Camera      Improve error reporting for PercentCompleted
'Camera      Moved SensorType test ahead of BayerOffset tests so that expected BayerOfset responses are known in advance
'Telescope   Revised DestinationSideOfPier test to just check that returned values are different for both sides of the meridian

'Version 6.0.0.49
'Telescope   Added test for whether interface V1 drivers can unpark the scope, fixes issue in MountSafety when mounts do not support unparking
'Telescope   Made SideOfPier Write test more realistic by tracking through the Meridian before attempting to flip the scope
'Focuser     Added better status messaging during asynchronous focuser moves

'Version 6.0.0.50
'Rotator     Increased Position test tolerance. A Move is now flagged as OK if the reported position is within 1 step of the required position

'Version 6.1.0.52
'All         Fix to suppress Windows "Application has stopped working" dialogues due to errors occurring in driver unmanaged code DLLs
'SafetyMonitor Added safety monitors as testable device types
'Switch      Added tests for ISwitchV2

' New comment section started!

#End Region

#Region "Platform 6.1 Changes"
' Always put new entries at the top to make it easier to transfer to the release notes text file!
'Version 6.1.0.55
'Telescope   Improved messaging for SideOfPier test results

'Version 6.1.0.54
'All         Consistent handling for not implemented and invalid value exceptions throughout the tester. 
'All         More guidance for authors when inappropriate exceptions are thrown
'All         Support for parsing error messages to determine intent has  been removed - use of inappropriate exceptions now count as issues
'All         Now reports when PropertyNotImplementedException is used instead of MethodNotImplementedException and vice versa
'Telescope   Added tests to confirm that slew and sync methods reject bad coordinates
'Telescope   Formatting consistency improvements for RA and Dec, Altitude and Azimuth values
'Telescope   Sync information made easier to read
'Telescope   TrackingRate Read is now required to be implemented per the specification, TrackingRate Write remains optional
'Rotator     Now properly handles rotators where returned positions are slightly less than the requested position

'Version 6.1.0.53
'All         New configuration option: Conform can now test use either late bound COM or DriverAccess to test drivers
'Camera      Added tests for the new ICameraV2 properties
'Video       Improved tests for Video image array compliance with the published interface standard

#End Region

#Region "Platform 6.2 Changes"
'Version 6.2.59
'All          Improved error handling for common methods
'All          Extended function name width to 34 characters to accommodate longer names
'Focuser      Removed incorrect error message in StepSize test that was introduced when refactoring exception handling.
'Focuser      Corrected reported test name when reading TempComp if temperature compensation write is not implemented.
'Focuser      Corrected reported test name when moving the focuser during the TempComp write test
'Rotator      Added extra debug logging to facilitate issue resolution
'Telescope    Sync - Added separate test enable flags for SyncToCoordinates and SyncToTarget
'Telescope    Sync - Improved, comprehensive test sequence to confirm that sync fully works
'Telescope    Sync - Telescope now slews to a known test position before commencing tests

'Version 6.2.58
'All          Up levelled to platform 6.2
'All          Changed version number format to PlatformMajor.PlatformMinor.ConformVersion.BetaVersion. BetaVersion will be 0 for production versions
'ObsCond      Added ObservingConditions interface tests
'Telescope    Increased resolution of UTCDate information to include 3 decimal places.
'Telescope    Fixed PulseGuide short duration bug introduced in release 6.1.0.55
'Telescope    Fixed SideOfPier tests to work reliably in both hemispheres
'Telescope    Added test for ASCOM convention that German Equatorial mounts should return pierWest when observing at hour angles from -6.0 to 0.0
#End Region

#Region "Platform 6.3 Changes"

'Version 6.3.60 - test version only
#End Region

' Always put new entries at the top to make it easier to transfer to the release notes text file!
#Region "Platform 6.4 Changes"
'Version 6.4.63 - Requires Platform 6.4
'Telescope           Fixed the DestinationSideOfPier test which used a "below the horizon" declination in the southern hemisphere

'Version 6.4.62 - Requires Platform 6.4
'All                 Change all HTTP references to HTTPS
'All                 Spelling corrections applied to application and comments
'Focuser             Added tests to make sure that absolute focusers can reach positions 0 and MaxStep
'Focuser             Added tests to make sure that absolute focusers gracefully fail when commanded outside their supported position range
'Telescope           Confirm that Slewing is False when RightAScensionRate and DeclinationRate are non-zero 
'Telescope           Confirm that SlewToCoordinates, SlewToCoordinatesAsync and SyncToCoordinates set the TargetRightAScension and TargetDeclination properties
'ObservingConditions Add extended SupportedActions / Action tests specifically for aggregated drivers presented by the Observing Conditions Hub

'Version 6.4.61
'Camera              New Conform configuration options to override the camera supplied MaxBinX and MaxBinY values
'Camera              Camera.PercentCompleted is now tested as having an optional implementation, in line with the published specification
'Camera              Revised exception handling for some properties and improved messaging
'ObservingConditions Added test to confirm that Wind direction is reported as 0.0 when wind speed is reported as 0.0
'Telescope           New configuration setting that will log a message before every driver call to aid debugging
'Telescope           TargetRightAscension - Messages now report the target RA rather than the RA that the telescope actually reported
'Telescope           Add check that no exception is thrown when calling AxisRates.Dispose
'Telescope           Add check that no exception is thrown when calling Rate.Dispose
'Telescope           Add check that no exception is thrown when calling TrackingRates.Dispose
'Switch              The range of switch numbers that will be searched for valid switches is now configurable through the Options/Conformance Options menu
'Switch              New configuration setting that will log a message before every driver call to aid debugging
#End Region


#Region "Imports"
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports Microsoft.Win32
#End Region

Public Class FrmConformMain

#Region "Variables and constants"
    ' Constants
    Const MINIMUM_PLATFORM_VERSION As Double = 6.4 ' Must be running under at least this version of the Platform. Currently 6.4 because of move to IFocuserV3
    Const FORCE_DEBUG As Boolean = False
    Const UPDATE_CHECK_FILE_DEBUG As String = "https://download.ascom-standards.org/devtools/conform.txt"
    Const UPDATE_CHECK_FILE_RELEASE As String = "https://download.ascom-standards.org/devtools/conform.txt"
    Const UPDATE_CHECK_TIMEOUT As Integer = 10000 'Version check timeout time in milliseconds
    Const UPDATE_FREQUENCY As Integer = 7 'Number of days between checks for version updates

    Const OLD_DIRECTORY As String = "C:\Program Files\Common Files\ASCOM\Conform"
    Const CONFORM_HELP_FILE As String = "Conform.chm"
    Const CONFORM_TOOLKIT_VERSION As Integer = 1000005 'Version of client access toolkit against which Conform is compiled

    ' Constants for text that the validation code searches for to find particular items of information
    Const VALIDATION_HASH_DRIVER As String = "Driver Hash Value:"
    Const VALIDATION_HASH_REPORTV1 As String = "Conform Report Hash (V1):"

    ' General variables
    Private m_TestRunning As Boolean
    Private m_CreateValidationFile As Boolean
    Private m_CurrentDeviceType As DeviceType
    Private m_LastLogFileDirectory As String

    ' Private security key for hash code initialisation
    Private m_SecKey As Byte() = New Byte() {29, 127, 3, 254, 69, 34, 78, 79, 43, 55, 80, 175, 35, 4, 251, 190, 43, 87, 113, 155, 217, 233, 2, 55}

#End Region

    ''' <summary>
    ''' Top level Conformance checking routine
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConformanceCheck()
        Dim l_FName, l_HashHex, l_Message As String, l_ValidationFileName, l_ValidationTempFileName As String
        Dim l_InStream, l_OutStream As FileStream
        Dim l_BytesRead As Integer, l_Buffer(1023), l_ByteArr() As Byte
        Dim l_Encoding As New System.Text.ASCIIEncoding()

        Dim l_TestDevice As DeviceTesterBaseClass 'Variable to hold the device being tested

        'Master routine to start tests
        If FORCE_DEBUG Then g_LogLevel = MessageLevel.msgDebug
#If DEBUG Then
        LogMsg("ConformanceCheck", MessageLevel.msgAlways, "THIS IS A DEBUG VERSION OF CONFORM AND NOT A RELEASE VERSION")
        LogMsg("", MessageLevel.msgAlways, "")
#End If
        If m_TestRunning Then
            MsgBox("A Conformance check is already running!")
            Exit Sub
        End If

        If g_Settings.Debug Then SupportInfo()

        m_TestRunning = True
        'Initialise error counters
        g_CountError = 0
        g_CountWarning = 0
        g_CountIssue = 0

        'Open file for validation
        My.Computer.FileSystem.CreateDirectory(g_Settings.LogFileDirectory & "\Logs " & Format(Now, "yyyy-MM-dd"))
        l_ValidationTempFileName = g_Settings.LogFileDirectory & "\Logs " & Format(Now, "yyyy-MM-dd") & "\" & Format(Now, "sffffhm")
        l_ValidationFileName = ""
        If g_Settings.CreateValidationFile Then
            g_ValidationTempFile = New StreamWriter(l_ValidationTempFileName, False) With {
                .AutoFlush = True
            }
        End If
        l_TestDevice = Nothing 'Initialise variables
        g_CurrentProgID = ""
        Select Case m_CurrentDeviceType 'Set current progID and device test class
            Case DeviceType.Telescope
                g_CurrentProgID = g_TelescopeProgID
                l_TestDevice = New TelescopeTester
            Case DeviceType.Dome
                g_CurrentProgID = g_DomeProgID
                l_TestDevice = New DomeTester
            Case DeviceType.Camera
                g_CurrentProgID = g_CameraProgID
                l_TestDevice = New CameraTester
            Case DeviceType.Video
                g_CurrentProgID = g_VideoCameraProgID
                l_TestDevice = New VideoTester
            Case DeviceType.Rotator
                g_CurrentProgID = g_RotatorProgID
                l_TestDevice = New RotatorTester
            Case DeviceType.Focuser
                g_CurrentProgID = g_FocuserProgID
                l_TestDevice = New FocuserTester
            Case DeviceType.ObservingConditions
                g_CurrentProgID = g_ObservingConditionsProgID
                l_TestDevice = New ObservingConditionsTester
            Case DeviceType.FilterWheel
                g_CurrentProgID = g_FilterWheelProgID
                l_TestDevice = New FilterWheelTester
            Case DeviceType.Switch
                g_CurrentProgID = g_SwitchProgID
                l_TestDevice = New SwitchTester
            Case DeviceType.SafetyMonitor
                g_CurrentProgID = g_SafetyMonitorProgID
                l_TestDevice = New SafetyMonitorTester
            Case DeviceType.CoverCalibrator
                g_CurrentProgID = g_CoverCalibratorProgID
                l_TestDevice = New CoverCalibratorTester
            Case Else
                LogMsg("Conform:ConformanceCheck", MessageLevel.msgError, "Unknown device type: " & m_CurrentDeviceType.ToString & ". You need to add it to the ConformanceCheck subroutine")
        End Select

        If g_CurrentProgID <> "" Then
            l_ValidationFileName = g_Settings.LogFileDirectory & "\Logs " & Format(Now, "yyyy-MM-dd") & "\" & g_CurrentProgID & ".Validation.txt"
            Try
                l_TestDevice.CheckInitialise()
                l_TestDevice.CheckAccessibility()
                LogMsg("", MessageLevel.msgAlways, "")
                Try
                    g_Stop = True 'Reset stop flag in case connect fails
                    l_TestDevice.CreateDevice()
                    g_Stop = False 'It worked so allow late steps to run
                    LogMsg("ConformanceCheck", MessageLevel.msgOK, "Driver instance created successfully")
                Catch ex As COMException
                    LogMsg("Initialise", MessageLevel.msgError, EX_COM & ex.Message) : LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot create the driver")
                Catch ex As PropertyNotImplementedException
                    LogMsg("Initialise", MessageLevel.msgError, NOT_IMP_NET & ex.Message) : LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot create the driver")
                Catch ex As DriverException
                    LogMsg("Initialise", MessageLevel.msgError, EX_DRV_NET & ex.Message) : LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot create the driver")
                Catch ex As Exception
                    LogMsg("Initialise", MessageLevel.msgError, EX_NET & ex.Message) : LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot create the driver")
                End Try

                If Not TestStop() And l_TestDevice.HasPreConnectCheck Then
                    LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("Pre-connect checks", MessageLevel.msgAlways, "")
                    l_TestDevice.PreConnectChecks()
                    LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("Connect", MessageLevel.msgAlways, "")
                End If

                If Not TestStop() Then 'Only connect if we successfully created the device
                    Try
                        g_Stop = True 'Reset stop flag in case connect fails
                        If g_Settings.DisplayMethodCalls Then LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Connected property")
                        l_TestDevice.Connected = True
                        g_Stop = False 'It worked so allow late steps to run
                        LogMsg("ConformanceCheck", MessageLevel.msgOK, "Connected OK")
                        LogMsg("", MessageLevel.msgAlways, "")
                    Catch ex As COMException
                        LogMsg("Connected", MessageLevel.msgError, EX_COM & ex.Message) : LogMsg("", MessageLevel.msgAlways, "")
                        LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot connect to the driver")
                    Catch ex As PropertyNotImplementedException
                        LogMsg("Connected", MessageLevel.msgError, NOT_IMP_NET & ex.Message) : LogMsg("", MessageLevel.msgAlways, "")
                        LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot connect to the driver")
                    Catch ex As DriverException
                        LogMsg("Connected", MessageLevel.msgError, EX_DRV_NET & ex.Message) : LogMsg("", MessageLevel.msgAlways, "")
                        LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot connect to the driver")
                    Catch ex As Exception
                        LogMsg("Connected", MessageLevel.msgError, EX_NET & ex.Message) : LogMsg("", MessageLevel.msgAlways, "")
                        LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot connect to the driver")
                    End Try
                End If

                'Run tests
                If (Not TestStop()) And g_Settings.TestProperties Then
                    l_TestDevice.CheckCommonMethods()
                End If

                If (Not TestStop()) And g_Settings.TestProperties And l_TestDevice.HasCanProperties Then
                    LogMsg("Can Properties", MessageLevel.msgAlways, "")
                    l_TestDevice.ReadCanProperties()
                    LogMsg("", MessageLevel.msgAlways, "")
                End If

                If (Not TestStop()) And l_TestDevice.HasPreRunCheck Then
                    LogMsg("Pre-run Checks", MessageLevel.msgAlways, "")
                    l_TestDevice.PreRunCheck()
                    LogMsg("", MessageLevel.msgAlways, "")
                End If

                If (Not TestStop()) And g_Settings.TestProperties And l_TestDevice.HasProperties Then
                    LogMsg("Properties", MessageLevel.msgAlways, "")
                    l_TestDevice.CheckProperties()
                    LogMsg("", MessageLevel.msgAlways, "")
                End If

                If (Not TestStop()) And g_Settings.TestMethods And l_TestDevice.HasMethods Then
                    LogMsg("Methods", MessageLevel.msgAlways, "")
                    l_TestDevice.CheckMethods()
                    LogMsg("", MessageLevel.msgAlways, "") 'Blank line
                End If
                If (Not TestStop()) And g_Settings.TestPerformance And l_TestDevice.HasPerformanceCheck Then
                    LogMsg("Performance", MessageLevel.msgAlways, "")
                    l_TestDevice.CheckPerformance()
                    LogMsg("", MessageLevel.msgAlways, "")
                End If

                If (Not TestStop()) And l_TestDevice.HasPostRunCheck Then
                    LogMsg("Post-run Checks", MessageLevel.msgAlways, "")
                    l_TestDevice.PostRunCheck()
                    LogMsg("", MessageLevel.msgAlways, "") 'Blank line
                End If

                If Not TestStop() Then
                    LogMsg("Conformance test complete", MessageLevel.msgAlways, "")
                Else
                    LogMsg("Conformance test interrupted by STOP button or to protect the device.", MessageLevel.msgAlways, "")
                End If

                l_TestDevice.Dispose()
                l_TestDevice = Nothing
                GC.Collect()
            Catch ex As Exception
                LogMsg("Conform:ConformanceCheck Exception: ", MessageLevel.msgError, ex.ToString)
            End Try

            'Extra check just to make sure the device has been disposed if there was an unexpected exception above!
            If Not l_TestDevice Is Nothing Then
                Try : l_TestDevice.Dispose() : Catch : End Try
                l_TestDevice = Nothing
            End If

        Else
            MsgBox("No device driver selected, please use the options menu to choose one.", MsgBoxStyle.Critical)
        End If
        m_TestRunning = False

        'Report the success or failure of conformance checking
        LogMsg("", MessageLevel.msgAlways, "")
        If g_CountError = 0 And g_CountWarning = 0 And g_CountIssue = 0 And Not g_Stop Then
            LogMsg("No errors, warnings or issues found: your driver passes ASCOM validation!!", MessageLevel.msgAlways, "")
        Else
            l_Message = "Your driver had " & g_CountError & " error"
            If g_CountError <> 1 Then l_Message = l_Message & "s"
            l_Message = l_Message & ", " & g_CountWarning & " warning"
            If g_CountWarning <> 1 Then l_Message = l_Message & "s"
            l_Message = l_Message & " and " & g_CountIssue & " issue"
            If g_CountWarning <> 1 Then l_Message = l_Message & "s"
            LogMsg(l_Message, MessageLevel.msgAlways, "")
        End If

        If m_CreateValidationFile Then ' User has requested that a validation file is created
            l_FName = FindDriverExecutable()

            If (l_FName Is Nothing) Or (l_FName = "") Then
                LogMsg("", MessageLevel.msgAlways, "")
                LogMsg("CONFORM ERROR:", MessageLevel.msgAlways, "During validation, Conform was unable to find the path to your driver's executable.")
                LogMsg("Please turn on Conform debug, run again and post the log file to the ASCOM Talk Yahoo group. Thanks, Peter", MessageLevel.msgAlways, "")
            Else
                LogMsg("CLSID FILE NAME", MessageLevel.msgDebug, l_FName)

                If Not My.Computer.FileSystem.FileExists(l_FName) Then
                    LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("CONFORM ERROR:", MessageLevel.msgAlways, "During validation, Conform could not find your executable at the location given in the registry.")
                    LogMsg("Please turn on Conform debug, run again and post the log file to the ASCOM Talk Yahoo group. Thanks, Peter", MessageLevel.msgAlways, "")

                    ' Trying other ways of opening the file
                    Try
                        LogMsg("File.Exists", MessageLevel.msgDebug, "Trying File.Exists")
                        If File.Exists(l_FName) Then
                            LogMsg("File.Exists", MessageLevel.msgDebug, "OK found the file through File.Exists")
                        Else
                            LogMsg("File.Exists", MessageLevel.msgDebug, "File does not exist for File.Exists!")
                        End If
                    Catch ex As Exception
                        LogMsg("File.Exists", MessageLevel.msgDebug, ex.ToString)
                    End Try

                    ' Trying other ways of opening the file
                    Try
                        LogMsg("File.OpenRead", MessageLevel.msgDebug, "Trying File.OpenRead")
                        Dim fs As FileStream
                        fs = File.OpenRead(l_FName)
                        LogMsg("File.OpenRead", MessageLevel.msgDebug, "OK found the file through File.OpenRead")
                        fs.Close()
                    Catch ex As Exception
                        LogMsg("File.OpenRead", MessageLevel.msgDebug, ex.ToString)
                    End Try

                    Try
                        LogMsg("New FileStream", MessageLevel.msgDebug, "Trying New FileStream")
                        l_InStream = New FileStream(l_FName, FileMode.Open, FileAccess.Read, FileShare.Read)
                        LogMsg("New FileStream", MessageLevel.msgDebug, "OK found the file through New FileStream")
                        l_InStream.Close()
                    Catch ex As Exception
                        LogMsg("New FileStream", MessageLevel.msgDebug, ex.ToString)
                    End Try

                Else
                    LogMsg("CLSID FILE EXISTS", MessageLevel.msgDebug, My.Computer.FileSystem.FileExists(l_FName).ToString) 'Confirm file exists

                    'Create hash code for file
                    l_InStream = New FileStream(l_FName, FileMode.Open, FileAccess.Read, FileShare.Read)
                    l_HashHex = CreateHashCode(m_SecKey, l_InStream)

                    LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg(VALIDATION_HASH_DRIVER, MessageLevel.msgAlways, l_HashHex)

                    'Clean up validation temporary file
                    g_ValidationTempFile.Flush() : g_ValidationTempFile.Close()
                    g_ValidationTempFile.Dispose() : g_ValidationTempFile = Nothing

                    'Create hash code for report file
                    l_InStream = New FileStream(l_ValidationTempFileName, FileMode.Open)
                    l_OutStream = New FileStream(l_ValidationFileName, FileMode.Create)

                    'Insert report hash code header
                    l_ByteArr = l_Encoding.GetBytes(VALIDATION_HASH_REPORTV1 & " ")
                    l_OutStream.Write(l_ByteArr, 0, l_ByteArr.Length)
                    l_HashHex = CreateHashCode(m_SecKey, l_InStream)
                    l_ByteArr = l_Encoding.GetBytes(l_HashHex)
                    l_OutStream.Write(l_ByteArr, 0, l_ByteArr.Length)
                    l_OutStream.WriteByte(13) : l_OutStream.WriteByte(10) 'Append CR LF
                    l_OutStream.WriteByte(13) : l_OutStream.WriteByte(10)

                    LogMsg("Report Hash Value:", MessageLevel.msgAlways, l_HashHex)
                    'Now append rest of the report file
                    l_InStream.Position = 0
                    Do
                        l_BytesRead = l_InStream.Read(l_Buffer, 0, 1024)
                        l_OutStream.Write(l_Buffer, 0, l_BytesRead)
                    Loop While l_BytesRead > 0

                    'Close the streams
                    l_InStream.Close()
                    l_OutStream.Close()

                    My.Computer.FileSystem.DeleteFile(l_ValidationTempFileName) 'Remove the temporary file
                    'Report the validation file name
                    LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("The validation file is: " & l_ValidationFileName, MessageLevel.msgAlways, "")

                    If g_Settings.Debug Then
                        LogMsg("", MessageLevel.msgAlways, "")
                        LogMsg("WARNING", MessageLevel.msgAlways, "DEBUG is turned on, this is not advisable when creating a validation file")
                    End If

                End If
            End If
        End If
        ' Clear down the test device and release memory
        l_TestDevice = Nothing
        GC.Collect()
#If DEBUG Then
        LogMsg("", MessageLevel.msgAlways, "")
        LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Conformance check subroutine finished")
#End If
    End Sub

#Region "Button, text box and checkbox event handlers"
    ' Check box events
    Private Sub ChkDomeShutter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDomeShutter.CheckedChanged
        'Save the new state when the open shutter check box changes state
        g_Settings.DomeShutter = chkDomeShutter.Checked
    End Sub
    Private Sub ChkSwitchSet_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSwitchSet.CheckedChanged
        'Save the new state when the set switch check box changes state
        g_Settings.SwitchSet = chkSwitchSet.Checked
    End Sub

    ' Button events
    Private Sub BtnClearAndCheckConformance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClearAndCheckConformance.Click
        txtStatus.Clear()
        ConformanceCheck()
    End Sub
    Private Sub BtnCheckConformance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCheckConformance.Click
        ConformanceCheck()
    End Sub
    Private Sub BtnCheckConformance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BtnCheckConformance.KeyUp
        If e.KeyCode = 116 Then
            ConformanceCheck()
            e.Handled = True
        End If
    End Sub
    Private Sub BtnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClear.Click
        txtStatus.Clear()
    End Sub
    Private Sub BtnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnStop.Click
        g_Stop = True
    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        End
    End Sub

    ' Text box events
    Private Sub TxtStatus_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStatus.KeyUp
        If e.KeyCode = 116 Then
            ConformanceCheck()
            e.Handled = True
        End If
    End Sub
#End Region

#Region "Menu Item event handlers"
    ' File menu
    Private Sub CheckConformanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckConformanceToolStripMenuItem.Click
        ConformanceCheck()
    End Sub
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    ' Options menu
    Private Sub SelectDriverToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectDriverToolStripMenuItem.Click
        Dim l_Chooser As ASCOM.Utilities.Chooser = Nothing, l_NewProgID As String = ""

        Try
            l_Chooser = New ASCOM.Utilities.Chooser
            l_Chooser.DeviceType = m_CurrentDeviceType.ToString 'Select the correct device type

            Select Case m_CurrentDeviceType
                Case DeviceType.Camera
                    l_NewProgID = l_Chooser.Choose(g_CameraProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceCamera = l_NewProgID 'Only update settings if OK pressed
                        g_CameraProgID = l_NewProgID
                    End If
                Case DeviceType.Video
                    l_NewProgID = l_Chooser.Choose(g_VideoCameraProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceVideo = l_NewProgID 'Only update settings if OK pressed
                        g_VideoCameraProgID = l_NewProgID
                    End If
                Case DeviceType.Dome
                    l_NewProgID = l_Chooser.Choose(g_DomeProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceDome = l_NewProgID 'Only update settings if OK pressed
                        g_DomeProgID = l_NewProgID
                    End If
                Case DeviceType.FilterWheel
                    l_NewProgID = l_Chooser.Choose(g_FilterWheelProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceFilterWheel = l_NewProgID 'Only update settings if OK pressed
                        g_FilterWheelProgID = l_NewProgID
                    End If
                Case DeviceType.Focuser
                    l_NewProgID = l_Chooser.Choose(g_FocuserProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceFocuser = l_NewProgID 'Only update settings if OK pressed
                        g_FocuserProgID = l_NewProgID
                    End If
                Case DeviceType.ObservingConditions
                    l_NewProgID = l_Chooser.Choose(g_ObservingConditionsProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceObservingConditions = l_NewProgID 'Only update settings if OK pressed
                        g_ObservingConditionsProgID = l_NewProgID
                    End If
                Case DeviceType.Rotator
                    l_NewProgID = l_Chooser.Choose(g_RotatorProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceRotator = l_NewProgID 'Only update settings if OK pressed
                        g_RotatorProgID = l_NewProgID
                    End If
                Case DeviceType.Telescope
                    l_NewProgID = l_Chooser.Choose(g_TelescopeProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceTelescope = l_NewProgID 'Only update settings if OK pressed
                        g_TelescopeProgID = l_NewProgID
                    End If
                Case DeviceType.Switch
                    l_NewProgID = l_Chooser.Choose(g_SwitchProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceSwitch = l_NewProgID 'Only update settings if OK pressed
                        g_SwitchProgID = l_NewProgID
                    End If
                Case DeviceType.SafetyMonitor
                    l_NewProgID = l_Chooser.Choose(g_SafetyMonitorProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceSafetyMonitor = l_NewProgID 'Only update settings if OK pressed
                        g_SafetyMonitorProgID = l_NewProgID
                    End If
                Case DeviceType.CoverCalibrator
                    l_NewProgID = l_Chooser.Choose(g_CoverCalibratorProgID)
                    If l_NewProgID <> "" Then
                        g_Settings.DeviceCoverCalibrator = l_NewProgID 'Only update settings if OK pressed
                        g_CoverCalibratorProgID = l_NewProgID
                    End If
                Case Else
                    LogMsg("SelectDriverToolStripMenuItem_Click", MessageLevel.msgError, "Invalid device type: " & m_CurrentDeviceType.ToString & " You need to add the device to the SelectDriverToolStripMenuItem_Click subroutine")
            End Select
            LogMsg("SelectDriver", MessageLevel.msgDebug, "Selected ProgID is: " & l_NewProgID & " " & m_CurrentDeviceType.ToString)
        Catch ex As Exception
            LogMsg("SelectDriver Choose", MessageLevel.msgError, ex.ToString)
        End Try

        'Clean up
        Try : l_Chooser.Dispose() : Catch : End Try
        Try : Marshal.ReleaseComObject(l_Chooser) : Catch ex As Exception : End Try
        l_Chooser = Nothing
        Me.BringToFront()
    End Sub
    Private Sub DriverSetupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriverSetupToolStripMenuItem.Click
        Dim l_Device As Object
        SetDeviceType()
        'MsgBox(g_CurrentProgID)
        LogMsg("DriverSetup", MessageLevel.msgDebug, "Before create object")

        l_Device = CreateObject(g_CurrentProgID)
        LogMsg("DriverSetup", MessageLevel.msgDebug, "Before setup")
        l_Device.SetupDialog()
        Try : l_Device.Dispose() : Catch : End Try
        Try : Marshal.ReleaseComObject(l_Device) : Catch : End Try
        LogMsg("DriverSetup", MessageLevel.msgDebug, "After try statements")
        l_Device = Nothing
        GC.Collect()
    End Sub
    Private Sub ConformanceOptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConformanceOptionsToolStripMenuItem.Click
        Dim SetupForm As frmSetup

        SetupForm = New frmSetup
        SetupForm.ShowDialog()
        If m_LastLogFileDirectory <> g_Settings.LogFileDirectory Then
            LogFileCreate(g_Settings.LogFileDirectory) 'Create new log file directory if needed
        End If
        If g_Settings.Debug Then
            g_LogLevel = MessageLevel.msgDebug
        Else
            g_LogLevel = MessageLevel.msgComment
        End If
        SetupForm.Close()
    End Sub
    Private Sub MnuTestCamera_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTestCamera.Click
        m_CurrentDeviceType = DeviceType.Camera
        SetDeviceType()
    End Sub
    Private Sub MnuTestCoverCalibrator_Click(sender As Object, e As EventArgs) Handles mnuTestCoverCalibrator.Click
        m_CurrentDeviceType = DeviceType.CoverCalibrator
        SetDeviceType()
    End Sub
    Private Sub MnuTestTelescope_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTestTelescope.Click
        m_CurrentDeviceType = DeviceType.Telescope
        SetDeviceType()
    End Sub
    Private Sub MnuTestDome_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTestDome.Click
        m_CurrentDeviceType = DeviceType.Dome
        SetDeviceType()
    End Sub
    Private Sub MnuTestFilterWheel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTestFilterWheel.Click
        m_CurrentDeviceType = DeviceType.FilterWheel
        SetDeviceType()
    End Sub
    Private Sub MnuTestFocuser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTestFocuser.Click
        m_CurrentDeviceType = DeviceType.Focuser
        SetDeviceType()
    End Sub
    Private Sub MnuTestRotator_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTestRotator.Click
        m_CurrentDeviceType = DeviceType.Rotator
        SetDeviceType()
    End Sub
    Private Sub MnuTestSwitch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTestSwitch.Click
        m_CurrentDeviceType = DeviceType.Switch
        SetDeviceType()
    End Sub
    Private Sub MnuTestSafetyMonitor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTestSafetyMonitor.Click
        m_CurrentDeviceType = DeviceType.SafetyMonitor
        SetDeviceType()
    End Sub
    Private Sub MnuTestVideoCamera_Click(sender As System.Object, e As System.EventArgs) Handles mnuTestVideoCamera.Click
        m_CurrentDeviceType = DeviceType.Video
        SetDeviceType()
    End Sub
    Private Sub CheckObservingConditionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuTestObservingConditions.Click
        m_CurrentDeviceType = DeviceType.ObservingConditions
        SetDeviceType()
    End Sub

    ' Validate menu
    Private Sub MnuCreateValidationFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCreateValidationFileToolStripMenuItem.Click
        'Create validation file flag has changed so save it
        g_Settings.CreateValidationFile = mnuCreateValidationFileToolStripMenuItem.Checked
        m_CreateValidationFile = mnuCreateValidationFileToolStripMenuItem.Checked
        g_Settings.TestProperties = True 'Turn on all tests if validation is requested
        g_Settings.TestMethods = True
        g_Settings.TestSideOfPierRead = True
    End Sub
    Private Sub MnuConfirmValidationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuConfirmValidationToolStripMenuItem.Click
        Try
            dlgValidation.DefaultExt = "txt"
            dlgValidation.Filter = "Validation files (*.txt)|*.txt"
            dlgValidation.InitialDirectory = g_Settings.LogFileDirectory & "\" & Format(Now, "yyyy-MM-dd")
            dlgValidation.FileName = g_CurrentProgID & ".Validation"
            dlgValidation.AutoUpgradeEnabled = True
            dlgValidation.Title = "Find Conform Validation File"
            dlgValidation.ShowDialog()
        Catch ex As Exception
            LogMsg("ERROR", MessageLevel.msgError, ex.ToString)
        End Try

    End Sub

    ' Special tests menu
    Private Sub MnuTelescopeSideOfPierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTelescopeSideOfPierToolStripMenuItem.Click
        Dim l_Device As TelescopeTester = New TelescopeTester
        Try
            If m_CurrentDeviceType = DeviceType.Telescope Then
                l_Device.SpecialTests(SpecialTests.TelescopeSideOfPier)
            Else
                LogMsg("This test only applies to Telescope drivers, please select Telescope as your device type", MessageLevel.msgAlways, "")
            End If
            l_Device.Dispose()
            l_Device = Nothing
        Catch ex As Exception
            LogMsg("Test", MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
        End Try
    End Sub
    Private Sub MnuTelescopeDestinationSideOfPierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTelescopeDestinationSideOfPierToolStripMenuItem.Click
        Dim l_Device As TelescopeTester = New TelescopeTester
        Try
            If m_CurrentDeviceType = DeviceType.Telescope Then
                l_Device.SpecialTests(SpecialTests.TelescopeDestinationSideOfPier)
            Else
                LogMsg("This test only applies to Telescope drivers, please select Telescope as your device type", MessageLevel.msgAlways, "")
            End If
            l_Device.Dispose()
            l_Device = Nothing
        Catch ex As Exception
            LogMsg("Test", MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
        End Try

    End Sub
    Private Sub MnuTelescopeFlipRangeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTelescopeFlipRangeToolStripMenuItem.Click
        Dim l_Device As TelescopeTester = New TelescopeTester
        Try
            If m_CurrentDeviceType = DeviceType.Telescope Then
                l_Device.SpecialTests(SpecialTests.TelescopeSideOfPierAnalysis)
            Else
                LogMsg("This test only applies to Telescope drivers, please select Telescope as your device type", MessageLevel.msgAlways, "")
            End If
            l_Device.Dispose()
            l_Device = Nothing
        Catch ex As Exception
            LogMsg("Test", MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
        End Try

    End Sub
    Private Sub MnuTelescopeCommandsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTelescopeCommandsToolStripMenuItem.Click
        Dim l_Device As TelescopeTester = New TelescopeTester
        Try
            If m_CurrentDeviceType = DeviceType.Telescope Then
                l_Device.SpecialTests(SpecialTests.TelescopeCommands)
            Else
                LogMsg("This test only applies to Telescope drivers, please select Telescope as your device type", MessageLevel.msgAlways, "")
            End If
            l_Device.Dispose()
            l_Device = Nothing
        Catch ex As Exception
            LogMsg("Test", MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
        End Try

    End Sub

    ' Help menu
    Private Sub HelpToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem1.Click
        Dim l_HelpFileName As String
        l_HelpFileName = Application.StartupPath & "\" & CONFORM_HELP_FILE
        LogMsg("HelpToolStripMenuItem1_Click", MessageLevel.msgDebug, "Help path: " & l_HelpFileName)
        If System.IO.File.Exists(l_HelpFileName) Then
            Diagnostics.Process.Start(l_HelpFileName)
        Else
            MsgBox("Sorry, no help yet!")
        End If
    End Sub
    Private Sub SideOfPierBehaviourPDFToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SideOfPierBehaviourPDFToolStripMenuItem.Click
        Dim file As String = ""
        Try
            file = Application.StartupPath & "\ASCOM SideOfPier Behaviour 1.2.pdf"
            LogMsg("Help", MessageLevel.msgDebug, file)
            Process.Start(file)
        Catch ex As Exception
            LogMsg("Help", MessageLevel.msgError, "Error searching for file: """ & file & """" & vbCrLf & ex.ToString)
        End Try
    End Sub
    Private Sub CheckForUpdatesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckForUpdatesToolStripMenuItem.Click
        CheckForUpdates()
    End Sub
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frmAboutBox.Show()
    End Sub

#End Region

#Region "New, Finalize, Form Load and Form Event handlers"
    Public Sub New()
        Try
            ' Add handlers for unhandled application and thread exceptions.
            AddHandler currentDomain.UnhandledException, AddressOf UnhandledExceptionHandler
            AddHandler Application.ThreadException, AddressOf UnhandledThreadExceptionHandler

            ' Change to the required bitness executable if necessary
            ChangeBitnessIfNeeded()

            ' This call is required by the Windows Form Designer.
            InitializeComponent()

            ' Create components
            g_Util = New ASCOM.Utilities.Util
        Catch ex As Exception
            MsgBox("Conform:New Exception " & ex.ToString)
        End Try
    End Sub

    Private Sub FrmConformMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try : Marshal.ReleaseComObject(g_Util) : Catch : End Try
        g_Util = Nothing
    End Sub
    Private Sub ConformMainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim l_DriverAccessDLL As String = ""
        Dim LoadedAssembly As Assembly = Nothing
        Dim FileVersion As System.Version = Nothing
        Dim OK As Boolean = False
        Dim Assemblies As Assembly(), AppDom As AppDomain

        Try
#If Not DEBUG Then
            SpecialTestsToolStripMenuItem.Visible = False 'Hide the special tests menu in release version
#End If
            m_CreateValidationFile = g_Settings.CreateValidationFile
            mnuCreateValidationFileToolStripMenuItem.Checked = m_CreateValidationFile
            g_TelescopeProgID = g_Settings.DeviceTelescope
            'g_LogLevel = CType(g_Settings.MessageLevel, MessageLevel)
            g_LogLevel = MessageLevel.msgComment
            If g_Settings.Debug Then g_LogLevel = MessageLevel.msgDebug
            If FORCE_DEBUG Then g_LogLevel = MessageLevel.msgDebug
            LogFileCreate(g_Settings.LogFileDirectory) 'Set up the log file
            Me.WindowState = FormWindowState.Maximized

            'Set up list of Telescope Test options
            Call SetTelescopeTestOptions()
            g_TelescopeTests = g_Settings.TeleScopeTests() 'Get actual values from registry

            If System.IntPtr.Size = 4 Then
                LogMsg("Start-up", MessageLevel.msgAlways, "ASCOM Device Conformance Checker - 32bit mode")
            Else
                LogMsg("Start-up", MessageLevel.msgAlways, "ASCOM Device Conformance Checker - 64bit mode")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Try
            LogMsg("Start-up", MessageLevel.msgDebug, "Initial application error level: " & NativeMethods.GetErrorMode.ToString("F"))
            ' Set the error level to ensure that all exceptions are returned to the application
            NativeMethods.SetErrorMode(NativeMethods.GetErrorMode Or NativeMethods.ErrorModes.ERRORMODE_SEM_FAILCRITICALERRORS Or NativeMethods.ErrorModes.ERRORMODE_SEM_NOGPFAULTERRORBOX Or NativeMethods.ErrorModes.ERRORMODE_SEM_NOOPENFILEERRORBOX)
            LogMsg("Start-up", MessageLevel.msgDebug, "New application error level:     " & NativeMethods.GetErrorMode.ToString("F") & vbCrLf)

        Catch ex As Exception
            LogMsg("Form.Load.EX2", MessageLevel.msgDebug, ex.ToString)
        End Try

        Try
            'Force Ensure operation only on supported Platform versions
            If Double.Parse(g_Util.PlatformVersion, CultureInfo.InvariantCulture) < MINIMUM_PLATFORM_VERSION Then
                LogMsg("", MessageLevel.msgAlways, "")
                LogMsg("This", MessageLevel.msgAlways, String.Format("version of Conform requires at least ASCOM Platform {0}, you appear to have platform version {1}", MINIMUM_PLATFORM_VERSION, g_Util.PlatformVersion))
                BtnClearAndCheckConformance.Enabled = False
                BtnCheckConformance.Enabled = False
                BtnClear.Enabled = False
                CheckConformanceToolStripMenuItem.Enabled = False
                OptionsToolStripMenuItem.Enabled = False
                ValidateToolStripMenuItem.Enabled = False

            Else ' Get the platform version details
                Using Reg As New Utilities.RegistryAccess
                    Dim RegKey As RegistryKey
                    RegKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\ASCOM\Platform")
                    LogMsg("Start-up", MessageLevel.msgAlways, RegKey.GetValue("Platform Name", "Platform Name not found!").ToString() & " " & RegKey.GetValue("Platform Version", "Platform Version not found!").ToString)
                    RegKey.Close()
                    RegKey.Dispose()
                End Using
            End If

            If Not g_Settings.WarningMessage Then frmWarning.ShowDialog() ' Show warning dialogue if this is first time through
            m_CurrentDeviceType = CType(g_Settings.CurrentDeviceType, DeviceType)
            SetDeviceType()

#If DEBUG Then
            LogMsg("", MessageLevel.msgAlways, "")
            LogMsg("Start-up", MessageLevel.msgAlways, "OPERATING IN DEBUG MODE!")
            LogMsg("", MessageLevel.msgAlways, "")
#End If
        Catch ex As Exception
            LogMsg("Form.Load.EX1", MessageLevel.msgError, ex.ToString)
        End Try

        Try 'List assemblies loaded in memory
            AppDom = AppDomain.CurrentDomain
            Assemblies = AppDom.GetAssemblies 'Get a list of loaded assemblies
            LogMsg("", MessageLevel.msgAlways, "")
            For Each FoundAssembly As Assembly In Assemblies 'Parse each to find if DriverAccess DLL is one of them
                LogMsg("Form.Load", MessageLevel.msgDebug, "Found loaded assembly: " & FoundAssembly.GetName.Name & " " & FoundAssembly.GetName.Version.ToString)
            Next
        Catch ex As Exception
            LogMsg("Form.Load.EX3", MessageLevel.msgDebug, ex.ToString)
        End Try
        LogMsg("", MessageLevel.msgDebug, "")

        'Check for updates if we are older than the last update check interval
        Try
            If g_Settings.UpdateCheck Then 'Only check if we are allowed to do so
                'Only check if we have gone longer than the maximum update test time
                If Now.CompareTo(g_Settings.UpdateCheckDate.AddDays(UPDATE_FREQUENCY)) > 0 Then
                    CheckForUpdates()
                End If
            End If
        Catch ex As Exception
            LogMsg("Form.Load.EX7", MessageLevel.msgError, "Exception checking for updates: " & ex.ToString)
        End Try

        Try 'Remove the old directory if it exists
            My.Computer.FileSystem.DeleteDirectory(OLD_DIRECTORY, FileIO.DeleteDirectoryOption.DeleteAllContents)
        Catch
        End Try
        Try 'Remove the old shortcut file
            My.Computer.FileSystem.DeleteFile("C:\Documents and Settings\All Users\Start Menu\Programs\ASCOM Platform\Tools\Uninstall Conform.lnk")
        Catch ex As Exception
        End Try
    End Sub
    Private Sub ConformMainForm_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Const l_ButtonIndent As Integer = 115
        Const l_ButtonTop As Integer = 225
        Const l_ButtonSpace As Integer = 42
        BtnCheckConformance.Left = Me.Width - l_ButtonIndent
        BtnCheckConformance.Top = Me.Height - l_ButtonTop
        BtnClear.Left = BtnCheckConformance.Left
        BtnClear.Top = BtnCheckConformance.Top + l_ButtonSpace
        BtnStop.Left = BtnCheckConformance.Left
        BtnStop.Top = BtnCheckConformance.Top + 2 * l_ButtonSpace
        BtnExit.Left = BtnCheckConformance.Left
        BtnExit.Top = BtnCheckConformance.Top + 3 * l_ButtonSpace
        BtnClearAndCheckConformance.Left = BtnCheckConformance.Left
        BtnClearAndCheckConformance.Top = BtnCheckConformance.Top - l_ButtonSpace
        txtStatus.Width = Me.Width - 143
        txtStatus.Height = Me.Height - 90
        chkDomeShutter.Left = BtnCheckConformance.Left
        chkDomeShutter.Top = BtnCheckConformance.Top - 3 * l_ButtonSpace
        chkSwitchSet.Left = BtnCheckConformance.Left
        chkSwitchSet.Top = BtnCheckConformance.Top - 3 * l_ButtonSpace
    End Sub

    ' Get the your application's application domain.
    Dim currentDomain As AppDomain = AppDomain.CurrentDomain

    Private Sub UnhandledExceptionHandler(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        LogMsg("UnhandledException", MessageLevel.msgWarning, e.ExceptionObject.ToString())
    End Sub

    Private Sub UnhandledThreadExceptionHandler(ByVal sender As Object, ByVal e As Threading.ThreadExceptionEventArgs)
        LogMsg("UnhandledThreadException", MessageLevel.msgWarning, e.Exception.ToString())
    End Sub
#End Region

#Region "Common Utility Functions"
    Private Sub SetDeviceType()
        'Clear any existing checked flags
        mnuTestCamera.Checked = False
        mnuTestDome.Checked = False
        mnuTestFilterWheel.Checked = False
        mnuTestTelescope.Checked = False
        mnuTestRotator.Checked = False
        mnuTestFocuser.Checked = False
        mnuTestSwitch.Checked = False
        mnuTestSafetyMonitor.Checked = False
        mnuTestVideoCamera.Checked = False
        mnuTestObservingConditions.Checked = False
        mnuTestCoverCalibrator.Checked = False

        'Set the appropriate new flag
        Select Case m_CurrentDeviceType
            Case DeviceType.Camera
                mnuTestCamera.Checked = True
                g_CurrentProgID = g_Settings.DeviceCamera
                g_CameraProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
            Case DeviceType.Dome
                mnuTestDome.Checked = True
                g_CurrentProgID = g_Settings.DeviceDome
                g_DomeProgID = g_CurrentProgID
                chkDomeShutter.Checked = g_Settings.DomeShutter
                chkDomeShutter.Visible = True 'Make the dome shutter control visible
                chkSwitchSet.Visible = False
            Case DeviceType.FilterWheel
                mnuTestFilterWheel.Checked = True
                g_CurrentProgID = g_Settings.DeviceFilterWheel
                g_FilterWheelProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
            Case DeviceType.Focuser
                mnuTestFocuser.Checked = True
                g_CurrentProgID = g_Settings.DeviceFocuser
                g_FocuserProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
            Case DeviceType.ObservingConditions
                mnuTestObservingConditions.Checked = True
                g_CurrentProgID = g_Settings.DeviceObservingConditions
                g_ObservingConditionsProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
            Case DeviceType.Rotator
                mnuTestRotator.Checked = True
                g_CurrentProgID = g_Settings.DeviceRotator
                g_RotatorProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
            Case DeviceType.Telescope
                mnuTestTelescope.Checked = True
                g_CurrentProgID = g_Settings.DeviceTelescope
                g_TelescopeProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
            Case DeviceType.Switch
                mnuTestSwitch.Checked = True
                g_CurrentProgID = g_Settings.DeviceSwitch
                g_SwitchProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Checked = g_Settings.SwitchSet
                chkSwitchSet.Visible = True

                If Not g_Settings.SwitchWarningMessage Then ' We have not yet shown the switch warning message, so do it now
                    Dim frmSwitchWarn As New frmSwitchWarning ' Create the warning message form
                    frmSwitchWarn.ShowDialog() ' Display the form as a dialogue and wait for the user to click OK
                    frmSwitchWarn.Dispose() ' Dispose of the form
                End If

                LogMsg("", MessageLevel.msgAlways, "")
                LogMsg("Switch read and write delays can be adjusted through the Options / Conformance Options menu", MessageLevel.msgAlways, "")
            Case DeviceType.SafetyMonitor
                mnuTestSafetyMonitor.Checked = True
                g_CurrentProgID = g_Settings.DeviceSafetyMonitor
                g_SafetyMonitorProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
            Case DeviceType.Video
                mnuTestVideoCamera.Checked = True
                g_CurrentProgID = g_Settings.DeviceVideo
                g_VideoCameraProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
            Case DeviceType.CoverCalibrator
                mnuTestCoverCalibrator.Checked = True
                g_CurrentProgID = g_Settings.DeviceCoverCalibrator
                g_CoverCalibratorProgID = g_CurrentProgID
                chkDomeShutter.Visible = False 'Make the dome shutter control invisible
                chkSwitchSet.Visible = False
        End Select
        g_Settings.CurrentDeviceType = m_CurrentDeviceType 'Save new device type for next run
    End Sub

    Private Sub LogFileCreate(ByVal p_LogFileDirectory As String)
        'Close existing log file if open and create log file in log file directory
        If m_LastLogFileDirectory <> "" Then
            If Not (g_LogFile Is Nothing) Then
                LogMsg("Conform:LogFileCreate", MessageLevel.msgAlways, "Closing current log file and creating a new one in " & p_LogFileDirectory)
                g_LogFile.Close()
                g_LogFile.Dispose()
                g_LogFile = Nothing
            End If
        End If
        My.Computer.FileSystem.CreateDirectory(p_LogFileDirectory & "\Logs " & Format(Now, "yyyy-MM-dd"))
        g_LogFile = New StreamWriter(p_LogFileDirectory & "\Logs " & Format(Now, "yyyy-MM-dd") & "\Conform " & Format(Now, "HHmm.ssfff") & ".txt", False) With {
            .AutoFlush = True
        }
        m_LastLogFileDirectory = g_Settings.LogFileDirectory
    End Sub
    Private Sub CheckForUpdates()
        Dim l_Version As String, l_MyVerNo, l_WebVerNo As Long
        Dim l_HttpWebResponse As HttpWebResponse, dataStream As Stream, reader As StreamReader
        Dim responseFromServer As String, l_WebRequest As WebRequest
        Dim l_UpdateFileURI As Uri

        LogMsg("", MessageLevel.msgAlways, "")
        LogMsg("CheckForUpdates", MessageLevel.msgAlways, "Checking for Conform updates with " & UPDATE_CHECK_TIMEOUT / 1000 & " second timeout")

        l_Version = My.Application.Info.Version.ToString
        LogMsg("CheckForUpdates", MessageLevel.msgInfo, "This version of Conform is:           " & l_Version)

        Try

            ServicePointManager.SecurityProtocol = SecurityProtocolType.SystemDefault ' Specify use of the latest secure transport mechanics

#If DEBUG Then
            l_WebRequest = WebRequest.Create(UPDATE_CHECK_FILE_DEBUG) 'URI of the test update version file
            l_UpdateFileURI = New Uri(UPDATE_CHECK_FILE_DEBUG)
#Else
            l_WebRequest = WebRequest.Create(UPDATE_CHECK_FILE_RELEASE) 'URI of the production update version file
            l_UpdateFileURI = New Uri(UPDATE_CHECK_FILE_RELEASE)
#End If
            LogMsg("CheckForUpdates", MessageLevel.msgDebug, "Target URI: " & l_UpdateFileURI.ToString)

            If WebRequest.DefaultWebProxy.GetProxy(l_UpdateFileURI) = l_UpdateFileURI Then ' No proxy is in use so go straight out
                LogMsg("CheckForUpdates", MessageLevel.msgDebug, "No proxy server detected, going directly to Internet")
            Else 'Proxy is in use so set it and apply credentials
                l_WebRequest.Proxy = WebRequest.DefaultWebProxy
                l_WebRequest.Proxy.Credentials = CredentialCache.DefaultCredentials
                LogMsg("CheckForUpdates", MessageLevel.msgDebug, "Using proxy server: " & WebRequest.DefaultWebProxy.GetProxy(l_UpdateFileURI).ToString)
            End If

            l_WebRequest.Timeout = UPDATE_CHECK_TIMEOUT 'Wait for up to 10 seconds
            LogMsg("CheckForUpdates", MessageLevel.msgDebug, "Sending request to web server with " & UPDATE_CHECK_TIMEOUT / 1000 & " second timeout...")

            l_HttpWebResponse = CType(l_WebRequest.GetResponse(), HttpWebResponse) ' Get the response.
            LogMsg("CheckForUpdates", MessageLevel.msgDebug, "Web status is: " & l_HttpWebResponse.StatusDescription)

            dataStream = l_HttpWebResponse.GetResponseStream() ' Get the stream containing content returned by the server.
            LogMsg("CheckForUpdates", MessageLevel.msgDebug, "Completed dataStream = response.GetResponseStream()")

            reader = New StreamReader(dataStream) ' Open the stream using a StreamReader for easy access.
            LogMsg("CheckForUpdates", MessageLevel.msgDebug, "Completed reader = New StreamReader(dataStream)")
            responseFromServer = reader.ReadLine ' Read the content.

            LogMsg("CheckForUpdates", MessageLevel.msgInfo, "The production version of Conform is: " & responseFromServer)

            'Horrible version compare - I can do it much more elegantly now!
            l_MyVerNo = VersionPart(l_Version, 1) * 10000L * 10000L * 10000L + VersionPart(l_Version, 2) * 10000L * 10000L + VersionPart(l_Version, 3) * 10000L + VersionPart(l_Version, 4)
            l_WebVerNo = VersionPart(responseFromServer, 1) * 10000L * 10000L * 10000L + VersionPart(responseFromServer, 2) * 10000L * 10000L + VersionPart(responseFromServer, 3) * 10000L + VersionPart(responseFromServer, 4)

            'Act depending on what the difference is between the retrieved version number and the installed version number
            Select Case l_WebVerNo - l_MyVerNo
                Case Is > 0
                    LogMsg("", MessageLevel.msgAlways, "")
                    LogMsg("CheckForUpdates", MessageLevel.msgInfo, "***** A NEWER VERSION OF CONFORM IS AVAILABLE: " & responseFromServer & " *****")
                    LogMsg("CheckForUpdates", MessageLevel.msgInfo, "***** It can be found at https://ascom-standards.org/Downloads/DevTools.htm *****")
                    LogMsg("", MessageLevel.msgAlways, "")

                    Do 'Get the rest of the update description and list it for the user
                        LogMsg("CheckForUpdates", MessageLevel.msgInfo, reader.ReadLine())
                    Loop Until reader.EndOfStream
                Case 0
                    LogMsg("CheckForUpdates", MessageLevel.msgInfo, "You already have the most up to date version of Conform")
                Case Is < 0
                    LogMsg("CheckForUpdates", MessageLevel.msgInfo, "You are using a development BETA version of Conform")
            End Select

            g_Settings.UpdateCheckDate = Now ' Save the date and time of this check so we know when to make the next check.

            ' Clean-up the streams and the response.            
            Try : reader.Close() : Catch : End Try
            Try : l_HttpWebResponse.Close() : Catch : End Try
        Catch ex As WebException
            LogMsg("CheckForUpdates", MessageLevel.msgError, "WebException: " & ex.Message)
        Catch ex As Exception
            LogMsg("CheckForUpdates", MessageLevel.msgError, "Exception: " & ex.Message)
            LogMsg("CheckForUpdates", MessageLevel.msgDebug, ex.StackTrace)
        End Try
        LogMsg("", MessageLevel.msgAlways, "")
    End Sub
    Private Function VersionPart(ByVal Version As String, ByVal Element As Integer) As Long
        Dim i, j As Integer, VerNo As String
        VerNo = ""
        Try
            Select Case Element
                Case 1
                    VerNo = Microsoft.VisualBasic.Left(Version, InStr(Version, ".") - 1)
                Case 2
                    i = InStr(Version, ".")
                    j = InStr(i + 1, Version, ".")
                    VerNo = Mid(Version, i + 1, j - i - 1)
                Case 3
                    i = InStr(Version, ".") 'Find first .
                    i = InStr(i + 1, Version, ".") 'Find second .
                    j = InStr(i + 1, Version, ".") 'Find third .
                    VerNo = Mid(Version, i + 1, j - i - 1)
                Case 4
                    i = InStr(Version, ".") 'Find first .
                    i = InStr(i + 1, Version, ".") 'Find second .
                    i = InStr(i + 1, Version, ".") 'Find third .
                    VerNo = Mid(Version, i + 1)
                Case Else
                    MsgBox("Invalid Element number: " & Element & "in function VersionPart")
            End Select
            'MsgBox(VerNo)
            VersionPart = CLng(VerNo)
        Catch ex As Exception
            VersionPart = 0
        End Try
    End Function
    Private Sub DlgValidation_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dlgValidation.FileOk
        Dim l_FName, l_ProgIDFile, l_ReportHashFile, l_ReportHashCalc, l_ExeHashFile, l_ExeHashCalc, l_Line As String
        Dim l_InStream As FileStream
        Dim l_ReportFile As StreamReader

        LogMsg("", MessageLevel.msgAlways, "")
        LogMsg("Validation", MessageLevel.msgInfo, "Validating driver: " & g_CurrentProgID)
        LogMsg("Validation", MessageLevel.msgInfo, "Validation file:   " & dlgValidation.FileName)

        l_FName = FindDriverExecutable() 'Find relevant driver executable for current driver

        If (l_FName Is Nothing) Or (l_FName = "") Then
            LogMsg("", MessageLevel.msgAlways, "")
            LogMsg("CONFORM ERROR:", MessageLevel.msgAlways, "During driver validation, Conform was unable to find the path to your driver's executable.")
            LogMsg("Please turn on Conform debug, run again and post the log file to the ASCOM Talk Yahoo group. Thanks, Peter", MessageLevel.msgAlways, "")
        Else
            LogMsg("CLSID FILE NAME", MessageLevel.msgDebug, l_FName)
            If Not My.Computer.FileSystem.FileExists(l_FName) Then
                LogMsg("", MessageLevel.msgAlways, "")
                LogMsg("CONFORM ERROR:", MessageLevel.msgAlways, "During driver validation, Conform could not find your executable at the location given in the registry.")
                LogMsg("Please turn on Conform debug, run again and post the log file to the ASCOM Talk Yahoo group. Thanks, Peter", MessageLevel.msgAlways, "")
            Else
                LogMsg("CLSID FILE EXISTS", MessageLevel.msgDebug, My.Computer.FileSystem.FileExists(l_FName).ToString) 'Confirm file exists

                'Create hash code for file
                l_InStream = New FileStream(l_FName, FileMode.Open, FileAccess.Read, FileShare.Read)
                l_ExeHashCalc = CreateHashCode(m_SecKey, l_InStream)
                l_InStream.Close()
                LogMsg("", MessageLevel.msgDebug, "")
                LogMsg("Driver Hash Value", MessageLevel.msgDebug, l_ExeHashCalc)

                If Not My.Computer.FileSystem.FileExists(dlgValidation.FileName) Then
                    LogMsg("Validation", MessageLevel.msgError, "Unable to find report file: " & dlgValidation.FileName)
                Else
                    l_ReportFile = My.Computer.FileSystem.OpenTextFileReader(dlgValidation.FileName)
                    l_ProgIDFile = ""
                    l_ExeHashFile = ""
                    l_ReportHashFile = ""
                    Do
                        l_Line = l_ReportFile.ReadLine()
                        If InStr(UCase(l_Line), UCase(DRIVER_PROGID)) > 0 Then l_ProgIDFile = l_Line
                        If InStr(UCase(l_Line), UCase(VALIDATION_HASH_DRIVER)) > 0 Then l_ExeHashFile = l_Line
                        If InStr(UCase(l_Line), UCase(VALIDATION_HASH_REPORTV1)) > 0 Then l_ReportHashFile = l_Line
                        'LogMsg("REPORTLINE", MessageLevel.msgDebug, l_Line)
                    Loop Until l_ReportFile.EndOfStream
                    l_ReportFile.Close() 'Close the stream reader
                    l_ReportFile = Nothing

                    If l_ProgIDFile = "" Then
                        LogMsg("Validation", MessageLevel.msgError, "Unable to find driver ProgID in conformance report")
                    Else
                        LogMsg("", MessageLevel.msgAlways, "")
                        l_ProgIDFile = Mid(l_ProgIDFile, InStr(UCase(l_ProgIDFile), UCase(DRIVER_PROGID)) + Len(DRIVER_PROGID)) 'Remove extraneous leading characters
                        If l_ProgIDFile = g_CurrentProgID Then
                            LogMsg("Validation", MessageLevel.msgOK, "Driver being validated is the same as the one on which the report is based!")
                        Else
                            LogMsg("Validation", MessageLevel.msgError, "Current driver is different to that stated in the conformance report: " & l_ProgIDFile)
                        End If
                        If l_ExeHashFile = "" Then 'Didn't find hash string
                            LogMsg("Validation", MessageLevel.msgError, "Unable to find driver hash code in a line starting with: " & VALIDATION_HASH_DRIVER)
                        Else
                            l_ExeHashFile = Mid(l_ExeHashFile, Len(VALIDATION_HASH_DRIVER) + 2)
                            LogMsg("ReportHashValue", MessageLevel.msgDebug, l_ExeHashFile)
                            If l_ExeHashCalc = l_ExeHashFile Then 'Matches OK
                                LogMsg("Validation", MessageLevel.msgOK, "Executable code integrity is confirmed!")
                            Else
                                LogMsg("Validation", MessageLevel.msgError, "The executable code is different from the code on which the conformance report is based.")
                            End If
                            If l_ReportHashFile = "" Then
                                LogMsg("Validation", MessageLevel.msgError, "Unable to find report hash code in a line starting with: " & VALIDATION_HASH_REPORTV1)
                            Else
                                l_ReportHashFile = Mid(l_ReportHashFile, Len(VALIDATION_HASH_REPORTV1) + 2)
                                LogMsg("Report Hash Value:", MessageLevel.msgDebug, l_ReportHashFile)
                                'Create hash code for report file
                                l_InStream = New FileStream(dlgValidation.FileName, FileMode.Open) With {
                                    .Position = Len(VALIDATION_HASH_REPORTV1) + 1 + Len(l_ReportHashFile) + 4
                                }
                                l_ReportHashCalc = CreateHashCode(m_SecKey, l_InStream)
                                LogMsg("ReportHashCalc", MessageLevel.msgDebug, l_ReportHashCalc)
                                If l_ReportHashCalc = l_ReportHashFile Then 'Matches OK
                                    LogMsg("Validation", MessageLevel.msgOK, "Conformance report integrity is confirmed!")
                                Else
                                    LogMsg("Validation", MessageLevel.msgError, "The conformance report has been altered after it was created")
                                End If
                                l_InStream.Close() : l_InStream = Nothing
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Function FindDriverExecutable() As String
        FindDriverExecutable = FindDriverInRegistry(Registry.ClassesRoot)
        If ASCOM.Utilities.ApplicationBits = Utilities.Bitness.Bits64 Then ' Check the 32bit registry on a 64bit system if we haven't found an executable yet
            If String.IsNullOrEmpty(FindDriverExecutable) Then ' We are on a 64bit OS and haven't yet found an executable
                LogMsg("FindDriverExecutable", MessageLevel.msgDebug, "We are a 64bit application and an executable has not yet been found, looking in 32bit registry")
                'Using Reg As ASCOM.Utilities.RegistryAccess = New ASCOM.Utilities.RegistryAccess
                Dim Rkey As RegistryKey
                Rkey = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32)
                FindDriverExecutable = FindDriverInRegistry(Rkey)
                Rkey.Close()
                Rkey.Dispose()
                'End Using
            End If
        End If

        LogMsg("FindDriverExecutable", MessageLevel.msgDebug, "Returned file name: " & FindDriverExecutable)
    End Function

    Private Function FindDriverInRegistry(ByVal Key As RegistryKey) As String
        Dim l_key, l_keyname As String
        Dim MyUri As Uri = Nothing
        Dim Retval As String

        'Find executable corresponding to tested COM object
        Retval = Nothing
        Try
            ' This section looks in the 64bit registry on a 64bit OS and in the 32bit registry on 1 32bit OS
            l_keyname = g_CurrentProgID & "\CLSID" 'Create class ID folder name based on ProgID
            LogMsg("PROGID KEY NAME", MessageLevel.msgDebug, l_keyname)
            l_key = Key.OpenSubKey(l_keyname, False).GetValue("", Nothing).ToString 'Retrieve CLSID from folder
            LogMsg("PROGID KEY VALUE", MessageLevel.msgDebug, l_key)
            l_keyname = "CLSID\" & l_key & "\LocalServer32" 'Create file folder name based on CLSID
            LogMsg("CLSID KEY NAME", MessageLevel.msgDebug, l_keyname)
            Try
                Retval = Key.OpenSubKey(l_keyname, False).GetValue("", Nothing).ToString 'Retrieve program file name
            Catch ex As Exception
            End Try

            If Retval Is Nothing Then
                LogMsg("CLSID KEY NAME", MessageLevel.msgDebug, "Could not find key: " & l_keyname)
                l_keyname = "CLSID\" & l_key & "\InprocServer32" 'Create file folder name based on CLSID
                LogMsg("CLSID KEY NAME", MessageLevel.msgDebug, l_keyname)
                Try
                    Retval = Key.OpenSubKey(l_keyname, False).GetValue("", Nothing).ToString
                Catch ex As Exception
                End Try

                If InStr(UCase(Retval), "MSCOREE.DLL") > 0 Then ' This is an assembly so look for codebase
                    LogMsg("CLSID KEY NAME", MessageLevel.msgDebug, "Found MSCOREE.DLL")
                    Retval = Key.OpenSubKey(l_keyname, False).GetValue("CodeBase", Nothing).ToString 'Retrieve program file name
                    LogMsg("CLSID KEY NAME", MessageLevel.msgDebug, "Found CodeBase: " & Retval)
                    If Uri.TryCreate(Retval, UriKind.Absolute, MyUri) Then
                        LogMsg("CLSID FILE NAME CLEAN", MessageLevel.msgDebug, "Codebase is a well formed URI path")
                        Retval = MyUri.LocalPath
                        LogMsg("CLSID FILE NAME CLEAN", MessageLevel.msgDebug, "Cleaned path: " & Retval)
                    Else
                        LogMsg("CLSID FILE NAME CLEAN", MessageLevel.msgDebug, "Codebase is not a well formed URI path")
                    End If
                Else
                    If String.IsNullOrEmpty(Retval) Then
                        LogMsg("CLSID KEY NAME", MessageLevel.msgDebug, "Could not find key: " & l_keyname)
                    Else
                        LogMsg("CLSID KEY NAME", MessageLevel.msgDebug, "Found driver: " & Retval)
                    End If
                End If
            End If

            If Not (Retval Is Nothing) Then Retval = Retval.Trim("""".ToCharArray)

            LogMsg("FindDriverInRegistry", MessageLevel.msgDebug, "Returned file name: " & Retval)
        Catch ex As Exception
            LogMsg("FindDriverInRegistry", MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
        End Try
        Return Retval
    End Function

    Private Function CreateHashCode(ByVal p_SecCode() As Byte, ByVal p_File As FileStream) As String
        Dim l_HMACSHA512 As HMACSHA512, l_HashValue() As Byte, l_i As Integer
        l_HMACSHA512 = New HMACSHA512(m_SecKey)
        'Set old style in case customer is running on pre .NET 2 SP1 code
        Try : l_HMACSHA512.ProduceLegacyHmacValues = True : Catch : End Try

        'Create hash code for report file
        l_HashValue = l_HMACSHA512.ComputeHash(p_File)
        CreateHashCode = ""
        For l_i = 0 To l_HashValue.Length - 1
            CreateHashCode = CreateHashCode & Microsoft.VisualBasic.Right("00" & Hex(l_HashValue(l_i)), 2)
        Next
        l_HMACSHA512.Clear()
        l_HMACSHA512 = Nothing
    End Function

    Private Sub SupportInfo()
        LogMsg("", MessageLevel.msgAlways, "")
        LogMsg("Support information", MessageLevel.msgAlways, "")
        Try : LogMsg("Support", MessageLevel.msgInfo, "OS version: " & Environment.OSVersion.ToString) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, ".NET version: " & Environment.Version.ToString) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Processor count: " & Environment.ProcessorCount.ToString) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "System directory: " & Environment.SystemDirectory) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Program files directory: " & Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Common program files directory: " & Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles)) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Common application data directory: " & Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Current directory: " & Environment.CurrentDirectory) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Current culture: " & My.Computer.Info.InstalledUICulture.ToString) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Total physical memory: " & My.Computer.Info.TotalPhysicalMemory.ToString("###,###,###,##0")) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Total virtual memory: " & My.Computer.Info.TotalVirtualMemory.ToString("###,###,###,##0")) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "OS full name: " & My.Computer.Info.OSFullName) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Available physical memory: " & My.Computer.Info.AvailablePhysicalMemory.ToString("###,###,###,##0")) : Catch : End Try
        Try : LogMsg("Support", MessageLevel.msgInfo, "Available virtual memory: " & My.Computer.Info.AvailableVirtualMemory.ToString("###,###,###,##0")) : Catch : End Try
        'Try : LogMsg("Support", MessageLevel.msgInfo, "Screen: " & My.Computer.Screen.ToString) : Catch : End Try
        LogMsg("", MessageLevel.msgAlways, "")
        LogMsg("", MessageLevel.msgAlways, "")
    End Sub

#End Region

End Class