'Option Strict On
Imports ASCOM.DeviceInterface

Friend Class TelescopeTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"
    Private Const TRACKING_COMMAND_DELAY As Integer = 1000 'Time to wait between changing Tracking state
    Private Const PERF_LOOP_TIME As Integer = 5 'Performance loop run time in seconds
    Private Const MOVE_AXIS_TIME As Integer = 2000 ' Number of milliseconds for each move axis command
    Private Const NUM_AXIS_RATES As Integer = 1000
    Private Const AXIS_RATE_MINIMUM As Integer = 0 ' Mnemonics for the axis rate array second dimension
    Private Const AXIS_RATE_MAXIMUM As Integer = 1
    Private Const WAIT_FOR_SLEW_MINIMUM_DURATION As Integer = 5 ' Minimum number of seconds to wait before declaring an asynchronous slew is finished (allows for mounts that don't set IsSlewing straight away!)
    Private Const SIDEOFPIER_MERIDIAN_TRACKING_PERIOD As Integer = 7 * 60 * 1000 ' 7 minutes in milliseconds
    Private Const DISPLAY_DECIMAL_DIGITS As Integer = 2

    Private Const PULSEGUIDE_MOVEMENT_TIME As Integer = 2 ' Initialise a pulse guide movement for this number of seconds
    Private Const PULSEGUIDE_TIMEOUT_TIME As Integer = 6 ' Wait up to this number of seconds before timing out a pulse guide command

    Private Const BAD_RA_LOW As Double = -1.0 ' Good range is 0.0 to 23.99999
    Private Const BAD_RA_HIGH As Double = 25.0
    Private Const BAD_DEC_LOW As Double = -100.0 ' Good range is -90.0 to +90.0
    Private Const BAD_DEC_HIGH As Double = 100.0
    Private Const BAD_AZIMUTH_LOW As Double = -10.0 ' Good range is 0.0 to 360.0
    Private Const BAD_AZIMUTH_HIGH As Double = 370.0
    Private Const BAD_ALTITUDE_LOW As Double = -100.0 ' Good range is -90.0 to +90.0 (-90.0 to allow the scope tube to be parked pointing vertically downwards)
    Private Const BAD_ALTITUDE_HIGH As Double = 100.0

    Private Const SIDE_OF_PIER_INVALID_LATITUDE As Double = 10.0 ' +- this value is the range of latitudes where side of pier tests will not be conducted
    Private Const SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR As Double = 0.5 ' Must be in the range 0..+1.0. Target declinations will be calculated as this fraction of the altitude of the highest pole from the horizon

    Private Const SLEW_SYNC_OK_TOLERANCE As Double = 10.0 ' (Arc seconds) Upper limit of Slew or Sync error that will generate an OK output, otherwise generates an Info message detailing distance
    Private Const SYNC_SIMULATED_ERROR As Double = 60.0 ' (Arc minutes) Distance that the telescope will be told that it is in error  when the sync command is issued. The error is applied to both RA and DEC axes

    Private canFindHome, canPark, canPulseGuide, canSetDeclinationRate, canSetGuideRates, canSetPark, canSetPierside, canSetRightAscensionRate As Boolean
    Private canSetTracking, canSlew, canSlewAltAz, canSlewAltAzAsync, canSlewAsync, canSync, canSyncAltAz, canUnpark As Boolean
    Private m_AlignmentMode As AlignmentModes, m_Altitude As Double, m_ApertureArea As Double, m_ApertureDiameter As Double
    Private m_AtHome As Boolean, m_AtPark As Boolean, m_Azimuth As Double, m_Declination As Double, m_DeclinationRate As Double
    Private m_DoesRefraction As Boolean, m_EquatorialSystem As EquatorialCoordinateType
    Private m_FocalLength As Double, m_GuideRateDeclination As Double, m_GuideRateRightAscension As Double
    Private m_IsPulseGuiding As Boolean, m_RightAscension As Double, m_RightAscensionRate As Double
    Private m_SideOfPier As PierSide, m_SiderealTimeScope As Double, m_SiteElevation As Double, m_SiteLatitude As Double, m_SiteLongitude As Double
    Private m_Slewing As Boolean, m_SlewSettleTime As Short, m_TargetDeclination As Double, m_TargetRightAscension As Double, m_Tracking As Boolean
    Private m_UTCDate As Date
    Private m_CanMoveAxisPrimary, m_CanMoveAxisSecondary, m_CanMoveAxisTertiary As Boolean
    Private m_DestinationSideOfPier, m_DestinationSideOfPierEast, m_DestinationSideOfPierWest As PierSide
    Private m_SiderealTimeASCOM As Double
    Private m_StartTime, m_EndTime, m_StartTimeUTC As Date
    Private m_LastLogFileDirectory As String, m_CanReadSideOfPier As Boolean
    Private m_TargetAltitude, m_TargetAzimuth As Double
    Private canReadAltitide, canReadAzimuth, canReadSiderealTime As Boolean
    Private trackingCanActuallybeChanged As Boolean

    Private telescopeDeviceType As DeviceType

#If DEBUG Then
    Private telescopeDevice As ASCOM.DriverAccess.Telescope
#Else
    Private telescopeDevice As Object
#End If

    Private DriverAsObject As Object
    'Axis rate checks
    Private m_AxisRatePrimaryOverlap, m_AxisRateSecondaryOverlap, m_AxisRateTertiaryOverlap As Boolean
    Private m_AxisRatesPrimaryArray(NUM_AXIS_RATES, 1), m_AxisRatesArray(NUM_AXIS_RATES, 1) As Double
#End Region

#Region "Enums"
    Private Enum CanType
        CanFindHome = 1
        CanPark = 2
        CanPulseGuide = 3
        CanSetDeclinationRate = 4
        CanSetGuideRates = 5
        CanSetPark = 6
        CanSetPierSide = 7
        CanSetRightAscensionRate = 8
        CanSetTracking = 9
        CanSlew = 10
        CanSlewAsync = 11
        CanSlewAltAz = 12
        CanSlewAltAzAsync = 13
        CanSync = 14
        CanSyncAltAz = 15
        CanUnPark = 16
        CanAbortExposure = 17
        CanAsymmetricBin = 18
        CanGetCoolerPower = 19
        CanSetCCDTemperature = 20
        CanStopExposure = 21
    End Enum
    Private Enum OptionalMethodType
        AbortSlew = 1
        DestinationSideOfPier = 2
        FindHome = 3
        MoveAxisPrimary = 4
        MoveAxisSecondary = 5
        MoveAxisTertiary = 6
        PulseGuide = 7
        SideOfPierWrite = 8
    End Enum
    Private Enum RequiredMethodType
        tstAxisrates = 1
        tstCanMoveAxisPrimary = 2
        tstCanMoveAxisSecondary = 3
        tstCanMoveAxisTertiary = 4
    End Enum
    Private Enum ParkedExceptionType
        tstPExcepAbortSlew = 1
        tstPExcepFindHome = 2
        tstPExcepMoveAxisPrimary = 3
        tstPExcepMoveAxisSecondary = 4
        tstPExcepMoveAxisTertiary = 5
        tstPExcepSlewToCoordinates = 6
        tstPExcepSlewToCoordinatesAsync = 7
        tstPExcepSlewToTarget = 8
        tstPExcepSlewToTargetAsync = 9
        tstPExcepSyncToCoordinates = 10
        tstPExcepSyncToTarget = 11
        tstPExcepPulseGuide = 12
    End Enum
    '   Private Enum SyncType
    '   End Enum
    Private Enum SlewSyncType
        SlewToCoordinates = 1
        SlewToCoordinatesAsync = 2
        SlewToTarget = 3
        SlewToTargetAsync = 4
        SlewToAltAz = 5
        SlewToAltAzAsync = 6
        SyncToCoordinates = 7
        SyncToTarget = 8
        SyncToAltAz = 9
    End Enum
    Private Enum PerformanceType
        tstPerfAltitude = 1
        tstPerfAtHome = 1
        tstPerfAtPark = 2
        tstPerfAzimuth = 3
        tstPerfDeclination = 4
        tstPerfIsPulseGuiding = 5
        tstPerfRightAscension = 6
        tstPerfSideOfPier = 7
        tstPerfSiderealTime = 8
        tstPerfSlewing = 9
        tstPerfUTCDate = 10
    End Enum
    Enum FlipTestType
        DestinationSideOfPier
        SideOfPier
    End Enum
    Private Enum InterfaceType
        ITelescope
        ITelescopeV2
        ITelescopeV3
    End Enum
#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(True, True, True, True, False, True, True) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        LogMsg("Dispose", MessageLevel.msgDebug, "Disposing of Telescope driver: " & disposing.ToString & " " & disposedValue.ToString)
        If Not Me.disposedValue Then
            If disposing Then
                If True Then 'Should be True but make False to stop Conform from cleanly dropping the telescope object (useful for retaining driver in memory to change flags)

                    Try : DisposeAndReleaseObject("Telescope Device", telescopeDevice) : Catch : End Try

                    telescopeDevice = Nothing
                    g_DeviceObject = Nothing
                    GC.Collect()
                End If
            End If

        End If

        MyBase.Dispose(disposing)
        disposedValue = True
    End Sub

#End Region

#Region "Code"
    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(telescopeDevice, DeviceType.Telescope)
    End Sub
    Overrides Sub CheckInitialise()
        'Set the error type numbers according to the standards adopted by individual authors.
        'Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
        'messages to driver authors!
        Select Case g_TelescopeProgID
            Case "Hub.Telescope"
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040401
                g_ExInvalidValue2 = &H80040402
                g_ExInvalidValue3 = &H80040405
                g_ExInvalidValue4 = &H80040402
                g_ExInvalidValue5 = &H80040402
                g_ExInvalidValue6 = &H80040402
                g_ExNotSet1 = &H80040403

            Case "ASCOM.MI250SA.Telescope", "Celestron.Telescope", "ASCOM.MI250.Telescope"
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040401
                g_ExInvalidValue2 = &H80040402
                g_ExInvalidValue3 = &H80040402
                g_ExInvalidValue4 = &H80040402
                g_ExInvalidValue5 = &H80040402
                g_ExInvalidValue6 = &H80040402
                g_ExNotSet1 = &H80040403
            Case "TemmaLite.Telescope"
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040410
                g_ExInvalidValue2 = &H80040418
                g_ExInvalidValue3 = &H80040418
                g_ExInvalidValue4 = &H80040418
                g_ExInvalidValue5 = &H80040418
                g_ExInvalidValue6 = &H80040418
                g_ExNotSet1 = &H80040417
            Case "Gemini.Telescope"
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040410
                g_ExInvalidValue2 = &H80040418
                g_ExInvalidValue3 = &H80040419
                g_ExInvalidValue4 = &H80040420
                g_ExInvalidValue5 = &H80040420
                g_ExInvalidValue6 = &H80040420
                g_ExNotSet1 = &H80040417
            Case "POTH.Telescope"
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040405
                g_ExInvalidValue2 = &H80040406
                g_ExInvalidValue3 = &H80040406
                g_ExInvalidValue4 = &H80040406
                g_ExInvalidValue5 = &H80040406
                g_ExInvalidValue6 = &H80040406
                g_ExNotSet1 = &H80040403
            Case "ServoCAT.Telescope"
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = ErrorCodes.InvalidValue
                g_ExInvalidValue2 = &H80040405
                g_ExInvalidValue3 = &H80040405
                g_ExInvalidValue4 = &H80040405
                g_ExInvalidValue5 = &H80040405
                g_ExInvalidValue6 = &H80040405
                g_ExNotSet1 = &H80040403
                g_ExNotSet2 = &H80040404
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = ErrorCodes.InvalidValue
                g_ExInvalidValue2 = &H80040405
                g_ExInvalidValue3 = &H80040405
                g_ExInvalidValue4 = &H80040405
                g_ExInvalidValue5 = &H80040405
                g_ExInvalidValue6 = &H80040405
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_TelescopeProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Try
            Dim l_DriverAccessTelescope As ASCOM.DriverAccess.Telescope = Nothing
            Dim l_ErrMsg As String = ""
            Dim l_TryCount As Integer = 0
            Try
                LogMsg("AccessChecks", MessageLevel.msgDebug, "Before MyBase.CheckAccessibility")
                MyBase.CheckAccessibility(g_TelescopeProgID, DeviceType.Telescope)
                LogMsg("AccessChecks", MessageLevel.msgDebug, "After MyBase.CheckAccessibility")
                Try
                    TestEarlyBinding(InterfaceType.ITelescopeV2)
                    TestEarlyBinding(InterfaceType.ITelescopeV3)

                    'Try client access toolkit
                    l_DriverAccessTelescope = Nothing
                    l_TryCount = 0
                    Do
                        l_TryCount += 1
                        Try
                            If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to create DriverAccess instance")
                            l_DriverAccessTelescope = New ASCOM.DriverAccess.Telescope(g_TelescopeProgID)
                            WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise")
                            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
                            Try
                                If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property true")
                                l_DriverAccessTelescope.Connected = True
                                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                                If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property false")
                                l_DriverAccessTelescope.Connected = False
                                LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully disconnected using driver access toolkit")
                            Catch ex As Exception
                                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.ToString)
                                LogMsg("", MessageLevel.msgAlways, "")
                            End Try
                        Catch ex As Exception
                            l_ErrMsg = ex.ToString
                        End Try
                        If l_DriverAccessTelescope Is Nothing Then WaitFor(200)
                    Loop Until ((l_TryCount = 3) Or (Not l_DriverAccessTelescope Is Nothing)) 'Exit if created OK

                    If l_DriverAccessTelescope Is Nothing Then
                        LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & l_ErrMsg)
                        LogMsg("", MessageLevel.msgAlways, "")
                    Else
                        LogMsg("Telescope:CreateDevice", MessageLevel.msgDebug, "Created telescope on attempt: " & l_TryCount.ToString)
                    End If
                Catch ex As Exception
                    LogMsg("Telescope:CheckAcc.EX3", MessageLevel.msgError, ex.ToString)
                End Try
            Catch ex As Exception
                LogMsg("Telescope:CheckAcc.EX2", MessageLevel.msgError, ex.ToString)
            End Try

            'Clean up
            Try : l_DriverAccessTelescope.Dispose() : Catch : End Try
            'Try : Marshal.ReleaseComObject(l_DriverAccessTelescope) : Catch : End Try
            l_DriverAccessTelescope = Nothing
            GC.Collect()
            WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for Accessibility Telescope Object to Dispose")
        Catch ex As Exception
            LogMsg("Telescope:CheckAcc.EX1", MessageLevel.msgError, ex.ToString)
        End Try
    End Sub
    Overrides Sub CreateDevice()
        Dim l_TryCount As Integer = 0
        Do
            l_TryCount += 1
            Try
                LogMsg("Telescope:CreateDevice", MessageLevel.msgDebug, "Creating ProgID: " & g_TelescopeProgID)
#If DEBUG Then
                LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Telescope to get a Telescope object")
                If g_Settings.DisplayMethodCalls Then LogMsg("CreateDevice", MessageLevel.msgComment, "About to create driver using DriverAccess")
                telescopeDevice = New DriverAccess.Telescope(g_TelescopeProgID)
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
                If g_Settings.UseDriverAccess Then
                    LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Telescope to get a Telescope object")
                    If g_Settings.DisplayMethodCalls Then LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to create driver using DriverAccess")
                    telescopeDevice = New DriverAccess.Telescope(g_TelescopeProgID)
                    LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
                Else
                    LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Telescope object")
                    If g_Settings.DisplayMethodCalls Then LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to create driver using CreateObject")
                    telescopeDevice = CreateObject(g_TelescopeProgID)
                    LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
                End If
#End If
                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise")
                g_Stop = False

            Catch ex As Exception
                LogMsg("", MessageLevel.msgDebug, "Attempt " & l_TryCount & " - exception thrown: " & ex.Message)
                If l_TryCount = 3 Then Throw 'Re throw exception if on our third attempt
            End Try
            If g_Stop Then WaitFor(200)
        Loop Until (Not g_Stop) 'Exit if created OK

        LogMsg("CreateDevice", MessageLevel.msgDebug, "Created telescope on attempt: " & l_TryCount.ToString)

        ' Create a pointer to the raw COM object that represents the Telescope (Only used for rate object Dispose() tests)
        Try
            If g_Settings.UseDriverAccess Then ' Use an internal DriverAccess field to get a pointer to the underlying COM driver
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Using DriverAccess to get underlying driver as an object")
                DriverAsObject = CType(telescopeDevice, DriverAccess.Telescope).memberFactory.GetLateBoundObject ' Have to convert the device from object to DriverAccess.Telescope in order to be able to access the internal GetLateBoundObject field
            Else
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Driver is already an object so using it ""as is"" for driver as an object")
                DriverAsObject = telescopeDevice
            End If
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Got driver object OK")

        Catch ex As Exception
            LogMsg("CreateDevice", MessageLevel.msgError, "Exception: " & ex.ToString())
        End Try

        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver as an object")

    End Sub
    Overrides Property Connected() As Boolean
        Get
            If g_Settings.DisplayMethodCalls Then LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Connected property")
            Connected = telescopeDevice.Connected
        End Get
        Set(ByVal value As Boolean)
            If g_Settings.DisplayMethodCalls Then LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Connected property " & value.ToString())
            telescopeDevice.Connected = value
            g_Stop = False
        End Set
    End Property
    Public Overrides Sub PreRunCheck()
        'Get into a consistent state
        If g_InterfaceVersion > 1 Then
            If g_Settings.DisplayMethodCalls Then LogMsg("Mount Safety", MessageLevel.msgComment, "About to get AtPark property")
            If telescopeDevice.AtPark Then
                If canUnpark Then
                    If g_Settings.DisplayMethodCalls Then LogMsg("Mount Safety", MessageLevel.msgComment, "About to call Unpark method")
                    telescopeDevice.Unpark()
                    LogMsg("Mount Safety", MessageLevel.msgInfo, "Scope is parked, so it has been unparked for testing")
                Else
                    LogMsg("Mount Safety", MessageLevel.msgError, "Scope reports that it is parked but CanUnPark is false - please manually unpark the scope")
                    g_Stop = True
                End If
            Else
                LogMsg("Mount Safety", MessageLevel.msgInfo, "Scope is not parked, continuing testing")
            End If
        Else
            LogMsg("Mount Safety", MessageLevel.msgInfo, "Skipping AtPark test as this method is not supported in interface V" & g_InterfaceVersion)
            Try
                If canUnpark Then
                    If g_Settings.DisplayMethodCalls Then LogMsg("Mount Safety", MessageLevel.msgComment, "About to call Unpark method")
                    telescopeDevice.Unpark()
                    LogMsg("Mount Safety", MessageLevel.msgOK, "Scope has been unparked for testing")
                Else
                    LogMsg("Mount Safety", MessageLevel.msgOK, "Scope reports that it cannot unpark, unparking skipped")
                End If
            Catch ex As Exception
                LogMsg("Mount Safety", MessageLevel.msgError, "Driver threw an exception while unparking: " & ex.Message)
            End Try
        End If

        ' Determine whether the tracking state can actually be changed
        Try
            canSetTracking = False
            If g_Settings.DisplayMethodCalls Then LogMsg("Mount Safety", MessageLevel.msgComment, "About to call CanSetTracking method")
            canSetTracking = telescopeDevice.CanSetTracking
            LogMsg("Mount Safety", MessageLevel.msgInfo, $"Can read CanSetTracking property: {canSetTracking}")
        Catch ex As Exception
            ' Ignore errors in getting the CanSetTracking value, at this early point we just assume that the tracking state cannot be set
            LogMsg("Mount Safety", MessageLevel.msgInfo, $"Cannot read CanSetTracking property... {ex.GetType().Name} - {ex.Message}")
        End Try

        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("Mount Safety", MessageLevel.msgComment, "About to set Tracking property true")
            telescopeDevice.Tracking = True
            LogMsg("Mount Safety", MessageLevel.msgInfo, "Can set Tracking property - Scope tracking has been enabled")
        Catch ex As Exception
            ' Ignore errors in setting the Tracking state, at this early point we just assume that the tracking state cannot be set
            LogMsg("Mount Safety", MessageLevel.msgInfo, $"Cannot set Tracking property... {ex.GetType().Name} - {ex.Message}")
        End Try

        If Not TestStop() Then
            LogMsg("TimeCheck", MessageLevel.msgInfo, "PC Time Zone:  " & g_Util.TimeZoneName & ", offset " & g_Util.TimeZoneOffset.ToString & " hours.")
            LogMsg("TimeCheck", MessageLevel.msgInfo, "PC UTCDate:    " & Format(g_Util.UTCDate, "dd-MMM-yyyy HH:mm:ss.fff"))
            'v1.0.12.0 Added catch logic for any UTCDate issues
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("TimeCheck", MessageLevel.msgComment, "About to get UTCDate property")
                LogMsg("TimeCheck", MessageLevel.msgDebug, "Mount UTCDate Unformatted: " & telescopeDevice.UTCDate)
                LogMsg("TimeCheck", MessageLevel.msgInfo, "Mount UTCDate: " & Format(telescopeDevice.UTCDate, "dd-MMM-yyyy HH:mm:ss.fff"))
            Catch ex As COMException
                If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
                    LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: COM exception - UTCDate not implemented in this driver")
                Else
                    LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: COM Exception - " & ex.Message)
                End If
            Catch ex As PropertyNotImplementedException
                LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: .NET exception - UTCDate not implemented in this driver")
            Catch ex As Exception
                LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: .NET Exception - " & ex.Message)
            End Try

        End If
    End Sub
    Public Overrides Sub ReadCanProperties()

        TelescopeCanTest(CanType.CanFindHome, "CanFindHome")
        TelescopeCanTest(CanType.CanPark, "CanPark")
        TelescopeCanTest(CanType.CanPulseGuide, "CanPulseGuide")
        TelescopeCanTest(CanType.CanSetDeclinationRate, "CanSetDeclinationRate")
        TelescopeCanTest(CanType.CanSetGuideRates, "CanSetGuideRates")
        TelescopeCanTest(CanType.CanSetPark, "CanSetPark")
        TelescopeCanTest(CanType.CanSetPierSide, "CanSetPierSide")
        Try
            If (telescopeDevice.AlignmentMode <> AlignmentModes.algGermanPolar) And canSetPierside Then LogMsg("CanSetPierSide", MessageLevel.msgIssue, "AlignmentMode is not GermanPolar but CanSetPierSide is true - contrary to ASCOM specification")
        Catch ex As Exception
        End Try
        TelescopeCanTest(CanType.CanSetRightAscensionRate, "CanSetRightAscensionRate")
        TelescopeCanTest(CanType.CanSetTracking, "CanSetTracking")
        TelescopeCanTest(CanType.CanSlew, "CanSlew")
        TelescopeCanTest(CanType.CanSlewAltAz, "CanSlewltAz")
        TelescopeCanTest(CanType.CanSlewAltAzAsync, "CanSlewAltAzAsync")
        TelescopeCanTest(CanType.CanSlewAsync, "CanSlewAsync")
        TelescopeCanTest(CanType.CanSync, "CanSync")
        TelescopeCanTest(CanType.CanSyncAltAz, "CanSyncAltAz")
        TelescopeCanTest(CanType.CanUnPark, "CanUnPark")
        If canUnpark And (Not canPark) Then LogMsg("CanUnPark", MessageLevel.msgIssue, "CanUnPark is true but CanPark is false - this does not comply with ASCOM specification")

    End Sub
    Overrides Sub CheckProperties()
        Dim l_OriginalTrackingState As Boolean, l_DriveRate As DriveRates, l_TimeDifference As Double
#If DEBUG Then
        Dim l_TrackingRates As ITrackingRates = Nothing
        Dim l_TrackingRate As DriveRates
#Else
        Dim l_TrackingRates As Object = Nothing
        Dim l_TrackingRate As Object
#End If

        'AlignmentMode - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("AlignmentMode", MessageLevel.msgComment, "About to get AlignmentMode property")
            m_AlignmentMode = telescopeDevice.AlignmentMode
            LogMsg("AlignmentMode", MessageLevel.msgOK, m_AlignmentMode.ToString)
        Catch ex As Exception
            HandleException("AlignmentMode", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'Altitude - Optional
        Try
            canReadAltitide = False
            If g_Settings.DisplayMethodCalls Then LogMsg("Altitude", MessageLevel.msgComment, "About to get Altitude property")
            m_Altitude = telescopeDevice.Altitude
            canReadAltitide = True 'Read successfully
            Select Case m_Altitude
                Case Is < 0.0
                    LogMsg("Altitude", MessageLevel.msgWarning, "Altitude is <0.0 degrees: " & Format(m_Altitude, "0.00000000"))
                Case Is > 90.0000001
                    LogMsg("Altitude", MessageLevel.msgWarning, "Altitude is >90.0 degrees: " & Format(m_Altitude, "0.00000000"))
                Case Else
                    LogMsg("Altitude", MessageLevel.msgOK, Format(m_Altitude, "0.00"))
            End Select
        Catch ex As Exception
            HandleException("Altitude", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'ApertureArea - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("ApertureArea", MessageLevel.msgComment, "About to get ApertureArea property")
            m_ApertureArea = telescopeDevice.ApertureArea
            Select Case m_ApertureArea
                Case Is < 0
                    LogMsg("ApertureArea", MessageLevel.msgWarning, "ApertureArea is < 0.0 : " & m_ApertureArea.ToString)
                    Exit Select
                Case 0.0
                    LogMsg("ApertureArea", MessageLevel.msgInfo, "ApertureArea is 0.0")
                Case Else
                    LogMsg("ApertureArea", MessageLevel.msgOK, m_ApertureArea.ToString)
            End Select
        Catch ex As Exception
            HandleException("ApertureArea", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'ApertureDiameter - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("ApertureDiameter", MessageLevel.msgComment, "About to get ApertureDiameter property")
            m_ApertureDiameter = telescopeDevice.ApertureDiameter
            Select Case m_ApertureDiameter
                Case Is < 0.0
                    LogMsg("ApertureDiameter", MessageLevel.msgWarning, "ApertureDiameter is < 0.0 : " & m_ApertureDiameter.ToString)
                    Exit Select
                Case 0.0
                    LogMsg("ApertureDiameter", MessageLevel.msgInfo, "ApertureDiameter is 0.0")
                Case Else
                    LogMsg("ApertureDiameter", MessageLevel.msgOK, m_ApertureDiameter.ToString)
            End Select
        Catch ex As Exception
            HandleException("ApertureDiameter", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'AtHome - Required
        If g_InterfaceVersion > 1 Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("AtHome", MessageLevel.msgComment, "About to get AtHome property")
                m_AtHome = telescopeDevice.AtHome
                LogMsg("AtHome", MessageLevel.msgOK, m_AtHome.ToString)
            Catch ex As Exception
                HandleException("AtHome", MemberType.Property, Required.Mandatory, ex, "")
            End Try
        Else
            LogMsg("AtHome", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        If TestStop() Then Exit Sub

        'AtPark - Required
        If g_InterfaceVersion > 1 Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("AtPark", MessageLevel.msgComment, "About to get AtPark property")
                m_AtPark = telescopeDevice.AtPark
                LogMsg("AtPark", MessageLevel.msgOK, m_AtPark.ToString)
            Catch ex As Exception
                HandleException("AtPark", MemberType.Property, Required.Mandatory, ex, "")
            End Try
        Else
            LogMsg("AtPark", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        If TestStop() Then Exit Sub

        'Azimuth - Optional
        Try
            canReadAzimuth = False
            If g_Settings.DisplayMethodCalls Then LogMsg("Azimuth", MessageLevel.msgComment, "About to get Azimuth property")
            m_Azimuth = telescopeDevice.Azimuth
            canReadAzimuth = True ' Read successfully
            Select Case m_Azimuth
                Case Is < 0.0
                    LogMsg("Azimuth", MessageLevel.msgWarning, "Azimuth is <0.0 degrees: " & Format(m_Azimuth, "0.00"))
                Case Is > 360.0000000001
                    LogMsg("Azimuth", MessageLevel.msgWarning, "Azimuth is >360.0 degrees: " & Format(m_Azimuth, "0.00"))
                Case Else
                    LogMsg("Azimuth", MessageLevel.msgOK, Format(m_Azimuth, "0.00"))
            End Select
        Catch ex As Exception
            HandleException("Azimuth", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'Declination - Required
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("Declination", MessageLevel.msgComment, "About to get Declination property")
            m_Declination = telescopeDevice.Declination
            Select Case m_Declination
                Case Is < -90.0, Is > 90.0
                    LogMsg("Declination", MessageLevel.msgWarning, "Declination is <-90 or >90 degrees: " & FormatDec(m_Declination))
                Case Else
                    LogMsg("Declination", MessageLevel.msgOK, FormatDec(m_Declination))
            End Select
        Catch ex As Exception
            HandleException("Declination", MemberType.Property, Required.Mandatory, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'DeclinationRate Read - Mandatory - must return a number even when CanSetDeclinationRate is False
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("DeclinationRate Read", MessageLevel.msgComment, "About to get DeclinationRate property")
            m_DeclinationRate = telescopeDevice.DeclinationRate
            'Read has been successful
            If canSetDeclinationRate Then 'Any value is acceptable
                Select Case m_DeclinationRate
                    Case Is >= 0.0
                        LogMsg("DeclinationRate Read", MessageLevel.msgOK, Format(m_DeclinationRate, "0.00"))
                    Case Else
                        LogMsg("DeclinationRate Read", MessageLevel.msgWarning, "Negative DeclinatioRate: " & Format(m_DeclinationRate, "0.00"))
                End Select
            Else 'Only zero is acceptable
                Select Case m_DeclinationRate
                    Case 0.0
                        LogMsg("DeclinationRate Read", MessageLevel.msgOK, Format(m_DeclinationRate, "0.00"))
                    Case Else
                        LogMsg("DeclinationRate Read", MessageLevel.msgIssue, "DeclinationRate is non zero when CanSetDeclinationRate is False " & Format(m_DeclinationRate, "0.00"))
                End Select
            End If
        Catch ex As Exception
            If Not canSetDeclinationRate Then LogMsg("DeclinationRate Read", MessageLevel.msgIssue, "DeclinationRate must return 0 even when CanSetDeclinationRate is false.")
            HandleException("DeclinationRate Read", MemberType.Property, Required.Mandatory, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'DeclinationRate Write - Optional
        If g_InterfaceVersion > 1 Then
            If canSetDeclinationRate Then 'Any value is acceptable
                If TestRADecRate("DeclinationRate Write", "Set rate to 0.0", Axis.Dec, 0.0, False) Then
                    TestRADecRate("DeclinationRate Write", "Set rate to 0.1", Axis.Dec, 0.1, False)
                    TestRADecRate("DeclinationRate Write", "Set rate to 1.0", Axis.Dec, 1.0, False)
                    TestRADecRate("DeclinationRate Write", "Reset rate to 0.0", Axis.Dec, 0.0, True) ' Reset the rate to zero, skipping the slewing test
                End If

            Else 'Should generate an error
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("DeclinationRate Write", MessageLevel.msgComment, "About to set DeclinationRate property to 0.0")
                    telescopeDevice.DeclinationRate = 0.0 'Set to a harmless value
                    LogMsg("DeclinationRate", MessageLevel.msgIssue, "CanSetDeclinationRate is False but setting DeclinationRate did not generate an error")
                Catch ex As Exception
                    HandleException("DeclinationRate Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetDeclinationRate is False")
                End Try
            End If
        Else
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("DeclinationRate Write", MessageLevel.msgComment, "About to set DeclinationRate property to 0.0")
                telescopeDevice.DeclinationRate = 0.0 'Set to a harmless value
                LogMsg("DeclinationRate Write", MessageLevel.msgOK, Format(m_DeclinationRate, "0.00"))
            Catch ex As Exception
                HandleException("DeclinationRate Write", MemberType.Property, Required.Optional, ex, "")
            End Try
        End If

        If TestStop() Then Exit Sub

        'DoesRefraction Read - Optional
        If g_InterfaceVersion > 1 Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("DoesRefraction Read", MessageLevel.msgComment, "About to DoesRefraction get property")
                m_DoesRefraction = telescopeDevice.DoesRefraction
                LogMsg("DoesRefraction Read", MessageLevel.msgOK, m_DoesRefraction.ToString)
            Catch ex As Exception
                HandleException("DoesRefraction Read", MemberType.Property, Required.Optional, ex, "")
            End Try
            If TestStop() Then Exit Sub
        Else
            LogMsg("DoesRefraction Read", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'DoesRefraction Write - Optional
        If g_InterfaceVersion > 1 Then
            If m_DoesRefraction Then 'Try opposite value
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("DoesRefraction Write", MessageLevel.msgComment, "About to set DoesRefraction property false")
                    telescopeDevice.DoesRefraction = False
                    LogMsg("DoesRefraction Write", MessageLevel.msgOK, "Can set DoesRefraction to False")
                Catch ex As Exception
                    HandleException("DoesRefraction Write", MemberType.Property, Required.Optional, ex, "")
                End Try
            Else 'Try other opposite value
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("DoesRefraction Write", MessageLevel.msgComment, "About to set DoesRefraction property true")
                    telescopeDevice.DoesRefraction = True
                    LogMsg("DoesRefraction Write", MessageLevel.msgOK, "Can set DoesRefraction to True")
                Catch ex As Exception
                    HandleException("DoesRefraction Write", MemberType.Property, Required.Optional, ex, "")
                End Try
            End If
        Else
            LogMsg("DoesRefraction Write", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        If TestStop() Then Exit Sub

        'EquatorialSystem - Required
        If g_InterfaceVersion > 1 Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("EquatorialSystem", MessageLevel.msgComment, "About to get EquatorialSystem property")
                m_EquatorialSystem = telescopeDevice.EquatorialSystem
                LogMsg("EquatorialSystem", MessageLevel.msgOK, m_EquatorialSystem.ToString)
            Catch ex As Exception
                HandleException("EquatorialSystem", MemberType.Property, Required.Mandatory, ex, "")
            End Try
        Else
            LogMsg("EquatorialSystem", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        If TestStop() Then Exit Sub

        'FocalLength - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("FocalLength", MessageLevel.msgComment, "About to get FocalLength property")
            m_FocalLength = telescopeDevice.FocalLength
            Select Case m_FocalLength
                Case Is < 0.0
                    LogMsg("FocalLength", MessageLevel.msgWarning, "FocalLength is <0.0 : " & m_FocalLength.ToString)
                    Exit Select
                Case 0.0
                    LogMsg("FocalLength", MessageLevel.msgInfo, "FocalLength is 0.0")
                Case Else
                    LogMsg("FocalLength", MessageLevel.msgOK, m_FocalLength.ToString)
            End Select
        Catch ex As Exception
            HandleException("FocalLength", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'GuideRateDeclination - Optional
        If g_InterfaceVersion > 1 Then
            If canSetGuideRates Then ' Can set guide rates so read and write are mandatory
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("GuideRateDeclination Read", MessageLevel.msgComment, "About to get GuideRateDeclination property")
                    m_GuideRateDeclination = telescopeDevice.GuideRateDeclination 'Read guiderateDEC
                    Select Case m_GuideRateDeclination
                        Case Is < 0.0
                            LogMsg("GuideRateDeclination Read", MessageLevel.msgWarning, "GuideRateDeclination is < 0.0 " & Format(m_GuideRateDeclination, "0.00"))
                        Case Else
                            LogMsg("GuideRateDeclination Read", MessageLevel.msgOK, Format(m_GuideRateDeclination, "0.00"))
                    End Select
                Catch ex As Exception 'Read failed
                    HandleException("GuideRateDeclination Read", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True")
                End Try
                Try 'Read OK so now try to write
                    If g_Settings.DisplayMethodCalls Then LogMsg("GuideRateDeclination Read", MessageLevel.msgComment, "About to set GuideRateDeclination property to " & m_GuideRateDeclination)
                    telescopeDevice.GuideRateDeclination = m_GuideRateDeclination
                    LogMsg("GuideRateDeclination Write", MessageLevel.msgOK, "Can write Declination Guide Rate OK")
                Catch ex As Exception 'Write failed
                    HandleException("GuideRateDeclination Write", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True")
                End Try
            Else
                Try 'Cannot set guide rates so Read is Optional and may generate an error
                    If g_Settings.DisplayMethodCalls Then LogMsg("GuideRateDeclination Read", MessageLevel.msgComment, "About to get GuideRateDeclination property")
                    m_GuideRateDeclination = telescopeDevice.GuideRateDeclination
                    Select Case m_GuideRateDeclination
                        Case Is < 0.0
                            LogMsg("GuideRateDeclination Read", MessageLevel.msgWarning, "GuideRateDeclination is < 0.0 " & Format(m_GuideRateDeclination, "0.00"))
                        Case Else
                            LogMsg("GuideRateDeclination Read", MessageLevel.msgOK, Format(m_GuideRateDeclination, "0.00"))
                    End Select
                Catch ex As Exception 'Some other error
                    HandleException("GuideRateDeclination Read", MemberType.Property, Required.Optional, ex, "")
                End Try
                Try 'Write should definitely raise an error
                    If g_Settings.DisplayMethodCalls Then LogMsg("GuideRateDeclination Write", MessageLevel.msgComment, "About to set GuideRateDeclination property to " & m_GuideRateDeclination)
                    telescopeDevice.GuideRateDeclination = m_GuideRateDeclination
                    LogMsg("GuideRateDeclination Write", MessageLevel.msgIssue, "CanSetGuideRates is false but no exception generated; value returned: " & Format(m_GuideRateDeclination, "0.00"))
                Catch ex As Exception 'Some other error so OK
                    HandleException("GuideRateDeclination Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetGuideRates is False")
                End Try
            End If
        Else
            LogMsg("GuideRateDeclination", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        If TestStop() Then Exit Sub

        'GuideRateRightAscension - Optional
        If g_InterfaceVersion > 1 Then
            If canSetGuideRates Then
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("GuideRateRightAscension Read", MessageLevel.msgComment, "About to get GuideRateRightAscension property")
                    m_GuideRateRightAscension = telescopeDevice.GuideRateRightAscension 'Read guiderateRA
                    Select Case m_GuideRateDeclination
                        Case Is < 0.0
                            LogMsg("GuideRateRightAscension Read", MessageLevel.msgWarning, "GuideRateRightAscension is < 0.0 " & Format(m_GuideRateRightAscension, "0.00"))
                        Case Else
                            LogMsg("GuideRateRightAscension Read", MessageLevel.msgOK, Format(m_GuideRateRightAscension, "0.00"))
                    End Select
                Catch ex As Exception 'Read failed
                    HandleException("GuideRateRightAscension Read", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True")
                End Try
                Try 'Read OK so now try to write
                    If g_Settings.DisplayMethodCalls Then LogMsg("GuideRateRightAscension Read", MessageLevel.msgComment, "About to set GuideRateRightAscension property to " & m_GuideRateRightAscension)
                    telescopeDevice.GuideRateRightAscension = m_GuideRateRightAscension
                    LogMsg("GuideRateRightAscension Write", MessageLevel.msgOK, "Can set RightAscension Guide OK")
                Catch ex As Exception 'Write failed
                    HandleException("GuideRateRightAscension Write", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True")
                End Try
            Else
                Try 'Cannot set guide rates so read may generate an error
                    If g_Settings.DisplayMethodCalls Then LogMsg("GuideRateRightAscension Read", MessageLevel.msgComment, "About to get GuideRateRightAscension property")
                    m_GuideRateRightAscension = telescopeDevice.GuideRateRightAscension 'Read guiderateRA
                    Select Case m_GuideRateDeclination
                        Case Is < 0.0
                            LogMsg("GuideRateRightAscension Read", MessageLevel.msgWarning, "GuideRateRightAscension is < 0.0 " & Format(m_GuideRateRightAscension, "0.00"))
                        Case Else
                            LogMsg("GuideRateRightAscension Read", MessageLevel.msgOK, Format(m_GuideRateRightAscension, "0.00"))
                    End Select
                Catch ex As Exception 'Some other error
                    HandleException("GuideRateRightAscension Read", MemberType.Property, Required.Optional, ex, "")
                End Try
                Try 'Write should definitely raise an error
                    If g_Settings.DisplayMethodCalls Then LogMsg("GuideRateRightAscension Write", MessageLevel.msgComment, "About to set GuideRateRightAscension property to " & m_GuideRateRightAscension)
                    telescopeDevice.GuideRateRightAscension = m_GuideRateRightAscension
                    LogMsg("GuideRateRightAscension Write", MessageLevel.msgIssue, "CanSetGuideRates is false but no exception generated; value returned: " & Format(m_GuideRateRightAscension, "0.00"))
                Catch ex As Exception 'Some other error so OK
                    HandleException("GuideRateRightAscension Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetGuideRates is False")
                End Try
            End If
        Else
            LogMsg("GuideRateRightAscension", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        If TestStop() Then Exit Sub

        'IsPulseGuiding - Optional
        If g_InterfaceVersion > 1 Then
            If canPulseGuide Then 'Can pulse guide so test if we can successfully read IsPulseGuiding
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("IsPulseGuiding", MessageLevel.msgComment, "About to get IsPulseGuiding property")
                    m_IsPulseGuiding = telescopeDevice.IsPulseGuiding
                    LogMsg("IsPulseGuiding", MessageLevel.msgOK, m_IsPulseGuiding.ToString)
                Catch ex As Exception 'Read failed
                    HandleException("IsPulseGuiding", MemberType.Property, Required.MustBeImplemented, ex, "CanPulseGuide is True")
                End Try
            Else 'Can't pulse guide so confirm that an error is raised
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("IsPulseGuiding", MessageLevel.msgComment, "About to get IsPulseGuiding property")
                    m_IsPulseGuiding = telescopeDevice.IsPulseGuiding
                    LogMsg("IsPulseGuiding", MessageLevel.msgIssue, "CanPulseGuide is False but no error was raised on calling IsPulseGuiding")
                Catch ex As Exception
                    HandleException("IsPulseGuiding", MemberType.Property, Required.MustNotBeImplemented, ex, "CanPulseGuide is False")
                End Try
            End If
        Else
            LogMsg("IsPulseGuiding", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        If TestStop() Then Exit Sub

        'RightAscension - Required
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("RightAscension", MessageLevel.msgComment, "About to get RightAscension property")
            m_RightAscension = telescopeDevice.RightAscension
            Select Case m_RightAscension
                Case Is < 0.0, Is >= 24.0
                    LogMsg("RightAscension", MessageLevel.msgWarning, "RightAscension is <0 or >=24 hours: " & m_RightAscension & " " & FormatRA(m_RightAscension))
                Case Else
                    LogMsg("RightAscension", MessageLevel.msgOK, FormatRA(m_RightAscension))
            End Select
        Catch ex As Exception
            HandleException("RightAscension", MemberType.Property, Required.Mandatory, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'RightAscensionRate Read - Mandatory because read must always return 0.0
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("RightAscensionRate Read", MessageLevel.msgComment, "About to get RightAscensionRate property")
            m_RightAscensionRate = telescopeDevice.RightAscensionRate
            'Read has been successful
            If canSetRightAscensionRate Then 'Any value is acceptable
                Select Case m_DeclinationRate
                    Case Is >= 0.0
                        LogMsg("RightAscensionRate Read", MessageLevel.msgOK, Format(m_RightAscensionRate, "0.00"))
                    Case Else
                        LogMsg("RightAscensionRate Read", MessageLevel.msgWarning, "Negative RightAscensionRate: " & Format(m_RightAscensionRate, "0.00"))
                End Select
            Else 'Only zero is acceptable
                Select Case m_RightAscensionRate
                    Case 0.0
                        LogMsg("RightAscensionRate Read", MessageLevel.msgOK, Format(m_RightAscensionRate, "0.00"))
                    Case Else
                        LogMsg("RightAscensionRate Read", MessageLevel.msgIssue, "RightAscensionRate is non zero when CanSetRightAscensionRate is False " & Format(m_DeclinationRate, "0.00"))
                End Select
            End If
        Catch ex As Exception
            If Not canSetRightAscensionRate Then LogMsg("RightAscensionRate Read", MessageLevel.msgInfo, "RightAscensionRate must return 0 if CanSetRightAscensionRate is false.")
            HandleException("RightAscensionRate Read", MemberType.Property, Required.Mandatory, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'RightAscensionRate Write - Optional
        If g_InterfaceVersion > 1 Then
            If canSetRightAscensionRate Then 'Perform several tests starting with proving we can set a rate of 0.0
                If TestRADecRate("RightAscensionRate Write", "Set rate to 0.0", Axis.RA, 0.0, False) Then
                    TestRADecRate("RightAscensionRate Write", "Set rate to 0.1", Axis.RA, 0.1, False)
                    TestRADecRate("RightAscensionRate Write", "Set rate to 1.0", Axis.RA, 1.0, False)
                    TestRADecRate("RightAscensionRate Write", "Reset rate to 0.0", Axis.RA, 0.0, True) ' Reset the rate to zero, skipping the slewing test
                End If

            Else 'Should generate an error
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("RightAscensionRate Write", MessageLevel.msgComment, "About to set RightAscensionRate property to 0.00")
                    telescopeDevice.RightAscensionRate = 0.0 'Set to a harmless value
                    LogMsg("RightAscensionRate Write", MessageLevel.msgIssue, "CanSetRightAscensionRate is False but setting RightAscensionRate did not generate an error")
                Catch ex As Exception
                    HandleException("RightAscensionRate Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetRightAscensionRate is False")
                End Try
            End If
        Else 'V1 has no Can property so just test anyway, it is optional
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("RightAscensionRate Write", MessageLevel.msgComment, "About to set RightAscensionRate property to 0.00")
                telescopeDevice.RightAscensionRate = 0.0 'Set to a harmless value
                LogMsg("RightAscensionRate Write", MessageLevel.msgOK, Format(m_RightAscensionRate, "0.00"))
            Catch ex As Exception
                HandleException("RightAscensionRate Write", MemberType.Property, Required.Optional, ex, "")
            End Try
        End If

        If TestStop() Then Exit Sub

        'SiteElevation Read - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteElevation Read", MessageLevel.msgComment, "About to get SiteElevation property")
            m_SiteElevation = telescopeDevice.SiteElevation
            Select Case m_SiteElevation
                Case Is < -300.0
                    LogMsg("SiteElevation Read", MessageLevel.msgIssue, "SiteElevation is <-300m")
                Case Is > 10000.0
                    LogMsg("SiteElevation Read", MessageLevel.msgIssue, "SiteElevation is >10,000m")
                Case Else
                    LogMsg("SiteElevation Read", MessageLevel.msgOK, m_SiteElevation.ToString)
            End Select
        Catch ex As Exception
            HandleException("SiteElevation Read", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'SiteElevation Write - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteElevation Write", MessageLevel.msgComment, "About to set SiteElevation property to -301.0")
            telescopeDevice.SiteElevation = -301.0
            LogMsg("SiteElevation Write", MessageLevel.msgIssue, "No error generated on set site elevation < -300m")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("SiteElevation Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site elevation < -300m")
        End Try

        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteElevation Write", MessageLevel.msgComment, "About to set SiteElevation property to 100001.0")
            telescopeDevice.SiteElevation = 10001.0
            LogMsg("SiteElevation Write", MessageLevel.msgIssue, "No error generated on set site elevation > 10,000m")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("SiteElevation Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site elevation > 10,000m")
        End Try

        Try
            If (m_SiteElevation < -300.0) Or (m_SiteElevation > 10000.0) Then m_SiteElevation = 1000
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteElevation Write", MessageLevel.msgComment, "About to set SiteElevation property to " & m_SiteElevation)
            telescopeDevice.SiteElevation = m_SiteElevation 'Restore original value
            LogMsg("SiteElevation Write", MessageLevel.msgOK, "Legal value " & m_SiteElevation.ToString & "m written successfully")
        Catch ex As Exception
            HandleException("SiteElevation Write", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'SiteLatitude Read - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteLatitude Read", MessageLevel.msgComment, "About to get SiteLatitude property")
            m_SiteLatitude = telescopeDevice.SiteLatitude
            Select Case m_SiteLatitude
                Case Is < -90.0
                    LogMsg("SiteLatitude Read", MessageLevel.msgWarning, "SiteLatitude is < -90 degrees")
                Case Is > 90.0
                    LogMsg("SiteLatitude Read", MessageLevel.msgWarning, "SiteLatitude is > 90 degrees")
                Case Else
                    LogMsg("SiteLatitude Read", MessageLevel.msgOK, FormatDec(m_SiteLatitude))
            End Select
        Catch ex As Exception
            HandleException("SiteLatitude Read", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'SiteLatitude Write - Optional
        Try 'Invalid low value
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteLatitude Write", MessageLevel.msgComment, "About to set SiteLatitude property to -91.0")
            telescopeDevice.SiteLatitude = -91.0
            LogMsg("SiteLatitude Write", MessageLevel.msgIssue, "No error generated on set site latitude < -90 degrees")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site latitude < -90 degrees")
        End Try
        Try 'Invalid high value
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteLatitude Write", MessageLevel.msgComment, "About to set SiteLatitude property to 91.0")
            telescopeDevice.SiteLatitude = 91.0
            LogMsg("SiteLatitude Write", MessageLevel.msgIssue, "No error generated on set site latitude > 90 degrees")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site latitude > 90 degrees")
        End Try
        Try 'Valid value
            If (m_SiteLatitude < -90.0) Or (m_SiteLatitude > 90.0) Then m_SiteLatitude = 45.0
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteLatitude Write", MessageLevel.msgComment, "About to set SiteLatitude property to " & m_SiteLatitude)
            telescopeDevice.SiteLatitude = m_SiteLatitude 'Restore original value
            LogMsg("SiteLatitude Write", MessageLevel.msgOK, "Legal value " & FormatDec(m_SiteLatitude) & " degrees written successfully")
        Catch ex As COMException
            If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
                LogMsg("SiteLatitude Write", MessageLevel.msgOK, NOT_IMP_COM)
            Else
                ExTest("SiteLatitude Write", ex.Message, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
            End If
        Catch ex As Exception
            HandleException("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'SiteLongitude Read - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteLongitude Read", MessageLevel.msgComment, "About to get SiteLongitude property")
            m_SiteLongitude = telescopeDevice.SiteLongitude
            Select Case m_SiteLongitude
                Case Is < -180.0
                    LogMsg("SiteLongitude Read", MessageLevel.msgWarning, "SiteLongitude is < -180 degrees")
                Case Is > 180.0
                    LogMsg("SiteLongitude Read", MessageLevel.msgWarning, "SiteLongitude is > 180 degrees")
                Case Else
                    LogMsg("SiteLongitude Read", MessageLevel.msgOK, FormatDec(m_SiteLongitude))
            End Select
        Catch ex As Exception
            HandleException("SiteLongitude Read", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'SiteLongitude Write - Optional
        Try 'Invalid low value
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteLongitude Write", MessageLevel.msgComment, "About to set SiteLongitude property to -181.0")
            telescopeDevice.SiteLongitude = -181.0
            LogMsg("SiteLongitude Write", MessageLevel.msgIssue, "No error generated on set site longitude < -180 degrees")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site longitude < -180 degrees")
        End Try

        Try 'Invalid high value
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteLongitude Write", MessageLevel.msgComment, "About to set SiteLongitude property to 181.0")
            telescopeDevice.SiteLongitude = 181.0
            LogMsg("SiteLongitude Write", MessageLevel.msgIssue, "No error generated on set site longitude > 180 degrees")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site longitude > 180 degrees")
        End Try

        Try 'Valid value
            If (m_SiteLongitude < -180.0) Or (m_SiteLongitude > 180.0) Then m_SiteLongitude = 60.0
            If g_Settings.DisplayMethodCalls Then LogMsg("SiteLongitude Write", MessageLevel.msgComment, "About to set SiteLongitude property to " & m_SiteLongitude)
            telescopeDevice.SiteLongitude = m_SiteLongitude 'Restore original value
            LogMsg("SiteLongitude Write", MessageLevel.msgOK, "Legal value " & FormatDec(m_SiteLongitude) & " degrees written successfully")
        Catch ex As Exception
            HandleException("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'Slewing - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("Slewing", MessageLevel.msgComment, "About to get Slewing property")
            m_Slewing = telescopeDevice.Slewing
            Select Case m_Slewing
                Case False
                    LogMsg("Slewing", MessageLevel.msgOK, m_Slewing.ToString)
                Case True
                    LogMsg("Slewing", MessageLevel.msgIssue, "Slewing should be false and it reads as " & m_Slewing.ToString)
            End Select
        Catch ex As Exception
            HandleException("Slewing", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'SlewSettleTime Read - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("SlewSettleTime Read", MessageLevel.msgComment, "About to get SlewSettleTime property")
            m_SlewSettleTime = telescopeDevice.SlewSettleTime
            Select Case m_SlewSettleTime
                Case Is < CShort(0)
                    LogMsg("SlewSettleTime Read", MessageLevel.msgWarning, "SlewSettleTime is < 0 seconds")
                Case Is > CShort(30.0)
                    LogMsg("SlewSettleTime Read", MessageLevel.msgInfo, "SlewSettleTime is > 30 seconds")
                Case Else
                    LogMsg("SlewSettleTime Read", MessageLevel.msgOK, m_SlewSettleTime.ToString)
            End Select
        Catch ex As Exception
            HandleException("SlewSettleTime Read", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'SlewSettleTime Write - Optional
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("SlewSettleTime Write", MessageLevel.msgComment, "About to set SlewSettleTime property to -1")
            telescopeDevice.SlewSettleTime = CShort(-1)
            LogMsg("SlewSettleTime Write", MessageLevel.msgIssue, "No error generated on set SlewSettleTime < 0 seconds")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("SlewSettleTime Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set slew settle time < 0")
        End Try
        Try
            If (m_SlewSettleTime < 0) Then m_SlewSettleTime = 0
            If g_Settings.DisplayMethodCalls Then LogMsg("SlewSettleTime Write", MessageLevel.msgComment, "About to set SlewSettleTime property to " & m_SlewSettleTime)
            telescopeDevice.SlewSettleTime = m_SlewSettleTime 'Restore original value
            LogMsg("SlewSettleTime Write", MessageLevel.msgOK, "Legal value " & m_SlewSettleTime.ToString & " seconds written successfully")
        Catch ex As Exception
            HandleException("SlewSettleTime Write", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'SideOfPier Read - Optional
        m_CanReadSideOfPier = False 'Start out assuming that we actually can't read side of pier so the performance test can be omitted
        If g_InterfaceVersion > 1 Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("SideOfPier Read", MessageLevel.msgComment, "About to get SideOfPier property")
                m_SideOfPier = telescopeDevice.SideOfPier
                LogMsg("SideOfPier Read", MessageLevel.msgOK, m_SideOfPier.ToString)
                m_CanReadSideOfPier = True 'Flag that it is OK to read SideOfPier
            Catch ex As Exception
                HandleException("SideOfPier Read", MemberType.Property, Required.Optional, ex, "")
            End Try
        Else
            LogMsg("SideOfPier Read", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'SideOfPier Write - Optional
        'Moved to methods section as this really is a method rather than a property

        'SiderealTime - Required
        Try
            canReadSiderealTime = False
            If g_Settings.DisplayMethodCalls Then LogMsg("SiderealTime", MessageLevel.msgComment, "About to get SiderealTime property")
            m_SiderealTimeScope = telescopeDevice.SiderealTime
            canReadSiderealTime = True

            m_SiderealTimeASCOM = (18.697374558 + 24.065709824419081 * (g_Util.DateLocalToJulian(Now) - 2451545.0) + (m_SiteLongitude / 15.0)) Mod 24.0
            Select Case m_SiderealTimeScope
                Case Is < 0.0, Is >= 24.0
                    LogMsg("SiderealTime", MessageLevel.msgWarning, "SiderealTime is <0 or >=24 hours: " & FormatRA(m_SiderealTimeScope))
                Case Else 'Valid time returned
                    'Now do a sense check on the received value
                    LogMsg("SiderealTime", MessageLevel.msgOK, FormatRA(m_SiderealTimeScope))
                    l_TimeDifference = System.Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM) 'Get time difference between scope and PC
                    'Process edge cases where the two clocks are on either side of 0:0:0/24:0:0
                    If (m_SiderealTimeASCOM > 23.0 And m_SiderealTimeASCOM < 23.999) And (m_SiderealTimeScope > 0.0 And m_SiderealTimeScope < 1.0) Then
                        l_TimeDifference = System.Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM + 24.0)
                    End If
                    If (m_SiderealTimeScope > 23.0 And m_SiderealTimeScope < 23.999) And (m_SiderealTimeASCOM > 0.0 And m_SiderealTimeASCOM < 1.0) Then
                        l_TimeDifference = System.Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM - 24.0)
                    End If

                    Select Case l_TimeDifference
                        Case Is <= (1.0 / 3600.0) '1 seconds
                            LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 1 second, Scope: " & FormatRA(m_SiderealTimeScope) & ", ASCOM: " & FormatRA(m_SiderealTimeASCOM))
                            Exit Select
                        Case Is <= (2.0 / 3600.0) '2 seconds
                            LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 2 seconds, Scope: " & FormatRA(m_SiderealTimeScope) & ", ASCOM: " & FormatRA(m_SiderealTimeASCOM))
                            Exit Select
                        Case Is <= (5.0 / 3600.0) '5 seconds
                            LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 5 seconds, Scope: " & FormatRA(m_SiderealTimeScope) & ", ASCOM: " & FormatRA(m_SiderealTimeASCOM))
                            Exit Select
                        Case Is <= (1.0 / 60.0) '1 minute
                            LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 1 minute, Scope: " & FormatRA(m_SiderealTimeScope) & ", ASCOM: " & FormatRA(m_SiderealTimeASCOM))
                            Exit Select
                        Case Is <= (5.0 / 60.0) '5 minutes
                            LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 5 minutes, Scope: " & FormatRA(m_SiderealTimeScope) & ", ASCOM: " & FormatRA(m_SiderealTimeASCOM))
                            Exit Select
                        Case Is <= 0.5 '0.5 an hour
                            LogMsg("SiderealTime", MessageLevel.msgInfo, "Scope and ASCOM sidereal times are up to 0.5 hour different, Scope: " & FormatRA(m_SiderealTimeScope) & ", ASCOM: " & FormatRA(m_SiderealTimeASCOM))
                            Exit Select
                        Case Is <= 1.0 '1.0 an hour
                            LogMsg("SiderealTime", MessageLevel.msgInfo, "Scope and ASCOM sidereal times are up to 1.0 hour different, Scope: " & FormatRA(m_SiderealTimeScope) & ", ASCOM: " & FormatRA(m_SiderealTimeASCOM))
                            Exit Select
                        Case Else
                            LogMsg("SiderealTime", MessageLevel.msgError, "Scope and ASCOM sidereal times are more than 1 hour apart, Scope: " & FormatRA(m_SiderealTimeScope) & ", ASCOM: " & FormatRA(m_SiderealTimeASCOM))
                            MsgBox("Following tests rely on correct sidereal time to calculate target RAs. The sidereal time returned by this driver is more than 1 hour from the expected value based on your computer clock and site longitude, so this program will end now to protect your mount from potential harm caused by slewing to an inappropriate location." & vbCrLf & vbCrLf & "Please check the longitude set by the driver and your PC clock (time, time zone and summer time) before checking the sidereal time code in your driver or your mount. Thanks, Peter", MsgBoxStyle.Critical, "CONFORM - MOUNT PROTECTION")
                            g_Stop = True
                            Exit Sub
                    End Select
            End Select
        Catch ex As Exception
            HandleException("SiderealTime", MemberType.Property, Required.Mandatory, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'TargetDeclination Read - Optional
        Try 'First read should fail!
            If g_Settings.DisplayMethodCalls Then LogMsg("TargetDeclination Read", MessageLevel.msgComment, "About to get TargetDeclination property")
            m_TargetDeclination = telescopeDevice.TargetDeclination
            LogMsg("TargetDeclination Read", MessageLevel.msgIssue, "Read before write should generate an error and didn't")
        Catch ex As COMException When (ex.ErrorCode = ErrorCodes.ValueNotSet) Or (ex.ErrorCode = g_ExNotSet1) Or (ex.ErrorCode = g_ExNotSet2)
            LogMsg("TargetDeclination Read", MessageLevel.msgOK, "COM Not Set exception generated on read before write")
        Catch ex As COMException When (ex.ErrorCode = ErrorCodes.InvalidOperationException)
            LogMsg("TargetDeclination Read", MessageLevel.msgOK, "COM InvalidOperationException generated on read before write")
        Catch ex As ASCOM.InvalidOperationException
            LogMsg("TargetDeclination Read", MessageLevel.msgOK, ".NET InvalidOperationException generated on read before write")
        Catch ex As DriverException When (ex.Number = ErrorCodes.ValueNotSet) Or (ex.Number = g_ExNotSet1) Or (ex.Number = g_ExNotSet2)
            LogMsg("TargetDeclination Read", MessageLevel.msgOK, ".NET Not Set exception generated on read before write")
        Catch ex As System.InvalidOperationException
            LogMsg("TargetDeclination Read", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException")
        Catch ex As Exception
            HandleException("TargetDeclination Read", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'TargetDeclination Write - Optional
        LogMsg("TargetDeclination Write", MessageLevel.msgInfo, "Tests moved after the SlewToCoordinates tests so that Conform can check they properly set target coordinates.")

        'TargetRightAscension Read - Optional
        Try 'First read should fail!
            If g_Settings.DisplayMethodCalls Then LogMsg("TargetRightAscension Read", MessageLevel.msgComment, "About to get TargetRightAscension property")
            m_TargetRightAscension = telescopeDevice.TargetRightAscension
            LogMsg("TargetRightAscension Read", MessageLevel.msgIssue, "Read before write should generate an error and didn't")
        Catch ex As COMException When (ex.ErrorCode = ErrorCodes.ValueNotSet) Or (ex.ErrorCode = g_ExNotSet1) Or (ex.ErrorCode = g_ExNotSet2)
            LogMsg("TargetRightAscension Read", MessageLevel.msgOK, "COM Not Set exception generated on read before write")
        Catch ex As COMException When (ex.ErrorCode = ErrorCodes.InvalidOperationException)
            LogMsg("TargetDeclination Read", MessageLevel.msgOK, "COM InvalidOperationException generated on read before write")
        Catch ex As ASCOM.InvalidOperationException
            LogMsg("TargetRightAscension Read", MessageLevel.msgOK, ".NET InvalidOperationException generated on read before write")
        Catch ex As DriverException When (ex.Number = ErrorCodes.ValueNotSet) Or (ex.Number = g_ExNotSet1) Or (ex.Number = g_ExNotSet2)
            LogMsg("TargetRightAscension Read", MessageLevel.msgOK, ".NET Not Set exception generated on read before write")
        Catch ex As System.InvalidOperationException
            LogMsg("TargetRightAscension Read", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException")
        Catch ex As Exception
            HandleException("TargetRightAscension Read", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'TargetRightAscension Write - Optional
        LogMsg("TargetRightAscension Write", MessageLevel.msgInfo, "Tests moved after the SlewToCoordinates tests so that Conform can check they properly set target coordinates.")

        'Tracking Read - Required
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("Tracking Read", MessageLevel.msgComment, "About to get Tracking property")
            m_Tracking = telescopeDevice.Tracking 'Read of tracking state is mandatory
            LogMsg("Tracking Read", MessageLevel.msgOK, m_Tracking.ToString)
        Catch ex As Exception
            HandleException("Tracking Read", MemberType.Property, Required.Mandatory, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'Tracking Write - Optional
        l_OriginalTrackingState = m_Tracking
        trackingCanActuallybeChanged = False ' Start by assuming that tracking cannot actually be changed
        If canSetTracking Then ' Set should work OK
            Try
                If m_Tracking Then 'OK try turning tracking off
                    If g_Settings.DisplayMethodCalls Then LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property false")
                    telescopeDevice.Tracking = False
                Else 'OK try turning tracking on
                    If g_Settings.DisplayMethodCalls Then LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property true")
                    telescopeDevice.Tracking = True
                End If
                WaitFor(TRACKING_COMMAND_DELAY) 'Wait for a short time to allow mounts to implement the tracking state change
                If g_Settings.DisplayMethodCalls Then LogMsg("Tracking Write", MessageLevel.msgComment, "About to get Tracking property")
                m_Tracking = telescopeDevice.Tracking
                If m_Tracking <> l_OriginalTrackingState Then
                    LogMsg("Tracking Write", MessageLevel.msgOK, m_Tracking.ToString)
                Else
                    LogMsg("Tracking Write", MessageLevel.msgIssue, "Tracking didn't change state on write: " & m_Tracking.ToString)
                End If
                If g_Settings.DisplayMethodCalls Then LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property " & l_OriginalTrackingState)
                telescopeDevice.Tracking = l_OriginalTrackingState 'Restore original state
                WaitFor(TRACKING_COMMAND_DELAY) 'Wait for a short time to allow mounts to implement the tracking state change
                trackingCanActuallybeChanged = True ' Log that Tracking can be successfully changed
            Catch ex As Exception
                HandleException("Tracking Write", MemberType.Property, Required.MustBeImplemented, ex, "CanSetTracking is True")
            End Try
        Else 'Can read OK but Set tracking should fail
            Try
                If m_Tracking Then 'OK try turning tracking off
                    If g_Settings.DisplayMethodCalls Then LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property false")
                    telescopeDevice.Tracking = False
                Else 'OK try turning tracking on
                    If g_Settings.DisplayMethodCalls Then LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property true")
                    telescopeDevice.Tracking = True
                End If
                If g_Settings.DisplayMethodCalls Then LogMsg("Tracking Write", MessageLevel.msgComment, "About to get Tracking property")
                m_Tracking = telescopeDevice.Tracking
                trackingCanActuallybeChanged = True ' Log that Tracking can be successfully changed

                LogMsg("Tracking Write", MessageLevel.msgIssue, "CanSetTracking is false but no error generated when value is set")
            Catch ex As Exception
                HandleException("Tracking Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetTracking is False")
            End Try
        End If
        If TestStop() Then Exit Sub

        'TrackingRates - Required
        If g_InterfaceVersion > 1 Then
            Dim l_Count As Integer = 0

            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRates", MessageLevel.msgComment, "About to get TrackingRates property")
                l_TrackingRates = telescopeDevice.TrackingRates
                If l_TrackingRates Is Nothing Then
                    LogMsg("TrackingRates", MessageLevel.msgDebug, "ERROR: The driver did NOT return an TrackingRates object!")
                Else
                    LogMsg("TrackingRates", MessageLevel.msgDebug, "OK - the driver returned an TrackingRates object")
                End If

                l_Count = l_TrackingRates.Count 'Save count for use later if no members are returned in the for each loop test
                LogMsg("TrackingRates Count", MessageLevel.msgDebug, l_Count)

                Dim i As Integer
                For i = 1 To l_TrackingRates.Count
                    LogMsg("TrackingRates Count", MessageLevel.msgDebug, "Found drive rate: " & [Enum].GetName(GetType(DriveRates), l_TrackingRates.Item(i)))
                Next
            Catch ex As Exception
                HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "")
            End Try

            Try
                Dim l_Enum As System.Collections.IEnumerator, l_Obj As Object, l_Drv As DriveRates
                l_Enum = l_TrackingRates.GetEnumerator()
                If l_Enum Is Nothing Then
                    LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "ERROR: The driver did NOT return an Enumerator object!")
                Else
                    LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "OK - the driver returned an Enumerator object")
                End If

                l_Enum.Reset()
                LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Reset Enumerator")
                Do While l_Enum.MoveNext()
                    LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Reading Current")
                    l_Obj = l_Enum.Current
                    LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Read Current OK, Type: " & l_Obj.GetType.Name)
                    l_Drv = CType(l_Obj, DriveRates)
                    LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Found drive rate: " & [Enum].GetName(GetType(DriveRates), l_Drv))
                Loop

                l_Enum.Reset()
                l_Enum = Nothing

                ' Clean up TrackingRates object
                If Not (l_TrackingRates Is Nothing) Then
                    Try : l_TrackingRates.Dispose() : Catch : End Try
                    Try : Marshal.ReleaseComObject(l_TrackingRates) : Catch : End Try
                    l_TrackingRates = Nothing
                End If

            Catch ex As Exception
                HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "")
            End Try

            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRates", MessageLevel.msgComment, "About to get TrackingRates property")
                l_TrackingRates = telescopeDevice.TrackingRates
                LogMsg("TrackingRates", MessageLevel.msgDebug, "Read TrackingRates OK, Count: " & l_TrackingRates.Count)

                Dim l_RateCount As Integer = 0
                For Each l_DriveRate In l_TrackingRates
                    LogMsg("TrackingRates", MessageLevel.msgComment, "Found drive rate: " & l_DriveRate.ToString())
                    l_RateCount += 1
                Next
                If l_RateCount > 0 Then
                    LogMsg("TrackingRates", MessageLevel.msgOK, "Drive rates read OK")
                Else
                    If l_Count > 0 Then 'We did get some members on the first call, but now they have disappeared!
                        'This can be due to the driver returning the same TrackingRates object on every TrackingRates call but not resetting the iterator pointer
                        LogMsg("TrackingRates", MessageLevel.msgError, "Multiple calls to TrackingRates returned different answers!")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "The first call to TrackingRates returned " & l_Count & " drive rates; the next call appeared to return no rates.")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "This can arise when the SAME TrackingRates object is returned on every TrackingRates call.")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "The root cause is usually that the enumeration pointer in the object is set to the end of the")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "collection through the application's use of the first object; subsequent uses see the pointer at the end")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "of the collection, which indicates no more members and is interpreted as meaning the collection is empty.")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "It is recommended to return a new TrackingRates object on each call. Alternatively, you could reset the")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "object's enumeration pointer every time the GetEnumerator method is called.")
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "")
                    Else
                        LogMsg("TrackingRates", MessageLevel.msgIssue, "No drive rates returned")
                    End If
                End If
            Catch ex As Exception
                HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "")
            End Try

            ' Clean up TrackingRates object
            If Not (l_TrackingRates Is Nothing) Then
                Try : l_TrackingRates.Dispose() : Catch : End Try
                Try : Marshal.ReleaseComObject(l_TrackingRates) : Catch : End Try
                l_TrackingRates = Nothing
            End If

            ' Test the TrackingRates.Dispose() method
            LogMsg("TrackingRates", MessageLevel.msgDebug, "Getting tracking rates")
            l_TrackingRates = DriverAsObject.TrackingRates
            Try
                LogMsg("TrackingRates", MessageLevel.msgDebug, "Disposing tracking rates")
                l_TrackingRates.Dispose()
                LogMsg("TrackingRates", MessageLevel.msgOK, "Disposed tracking rates OK")
            Catch ex As MissingMemberException
                LogMsg("TrackingRates", MessageLevel.msgOK, "Dispose member not present")
            Catch ex As Exception
                LogMsgWarning("TrackingRates", "TrackingRates.Dispose() threw an exception but it is poor practice to throw exceptions in Dispose() methods: " & ex.Message())
                LogMsg("TrackingRates.Dispose", MessageLevel.msgDebug, "Exception: " & ex.ToString())
            End Try
            Try : Marshal.ReleaseComObject(l_TrackingRates) : Catch : End Try
            l_TrackingRates = Nothing
            If TestStop() Then Exit Sub

            'TrackingRate - Test after TrackingRates so we know what the valid values are
            'TrackingRate Read - Required
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRates", MessageLevel.msgComment, "About to get TrackingRates property")
                l_TrackingRates = telescopeDevice.TrackingRates
                If Not (l_TrackingRates Is Nothing) Then ' Make sure that we have received a TrackingRates object after the Dispose() method was called
                    LogMsgOK("TrackingRates", "Successfully obtained a TrackingRates object after the previous TrackingRates object was disposed")

                    If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRate Read", MessageLevel.msgComment, "About to get TrackingRate property")
                    l_TrackingRate = telescopeDevice.TrackingRate
                    LogMsg("TrackingRate Read", MessageLevel.msgOK, l_TrackingRate.ToString)

                    'TrackingRate Write - Optional
                    'We can read TrackingRate so now test trying to set each tracking rate in turn
                    Try
                        LogMsgDebug("TrackingRate Write", "About to enumerate tracking rates object")
                        For Each l_DriveRate In l_TrackingRates
                            Application.DoEvents()
                            If TestStop() Then Exit Sub
                            Try
                                If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to set TrackingRate property to " & l_DriveRate.ToString())
                                telescopeDevice.TrackingRate = l_DriveRate
                                If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to get TrackingRate property")
                                If telescopeDevice.TrackingRate = l_DriveRate Then
                                    LogMsg("TrackingRate Write", MessageLevel.msgOK, "Successfully set drive rate: " & l_DriveRate.ToString)
                                Else
                                    LogMsg("TrackingRate Write", MessageLevel.msgIssue, "Unable to set drive rate: " & l_DriveRate.ToString)
                                End If
                            Catch ex As Exception
                                HandleException("TrackingRate Write", MemberType.Property, Required.Optional, ex, "")
                            End Try
                        Next
                    Catch ex As System.NullReferenceException ' Catch issues in iterating over a new TrackingRates object after a previous TrackingRates object was disposed.
                        LogMsgError("TrackingRate Write 1", "A NullReferenceException was thrown while iterating a new TrackingRates instance after a previous TrackingRates instance was disposed. TrackingRate.Write testing skipped")
                        LogMsgInfo("TrackingRate Write 1", "This may indicate that the TrackingRates.Dispose method cleared a global variable shared by all TrackingRates instances.")
                    Catch ex As Exception
                        HandleException("TrackingRate Write 1", MemberType.Property, Required.Mandatory, ex, "")
                    End Try

                    ' Attempt to write an invalid high tracking rate
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to set TrackingRate property to invalid value (5)")
                        telescopeDevice.TrackingRate = 5
                        LogMsg("TrackingRate Write", MessageLevel.msgIssue, "No error generated when TrackingRate is set to an invalid value (5)")
                    Catch ex As Exception
                        HandleInvalidValueExceptionAsOK("TrackingRate Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected when TrackingRate is set to an invalid value (5)")
                    End Try

                    ' Attempt to write an invalid low tracking rate
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to set TrackingRate property to invalid value (-1)")
                        telescopeDevice.TrackingRate = -1
                        LogMsg("TrackingRate Write", MessageLevel.msgIssue, "No error generated when TrackingRate is set to an invalid value (-1)")
                    Catch ex As Exception
                        HandleInvalidValueExceptionAsOK("TrackingRate Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected when TrackingRate is set to an invalid value (-1)")
                    End Try

                    'Finally restore original TrackingRate
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to set TrackingRate property to " & l_TrackingRate.ToString())
                        telescopeDevice.TrackingRate = l_TrackingRate
                    Catch ex As Exception
                        HandleException("TrackingRate Write", MemberType.Property, Required.Optional, ex, "Unable to restore original tracking rate")
                    End Try
                Else ' No TrackingRates object received after disposing of a previous instance
                    LogMsgError("TrackingRate Write", "TrackingRates did not return an object after calling Disposed() on a previous instance, TrackingRate.Write testing skipped")
                End If
            Catch ex As Exception
                HandleException("TrackingRate Read", MemberType.Property, Required.Mandatory, ex, "")
            End Try
        Else
            LogMsg("TrackingRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        If TestStop() Then Exit Sub

        'UTCDate Read - Required
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("UTCDate Read", MessageLevel.msgComment, "About to get UTCDate property")
            m_UTCDate = telescopeDevice.UTCDate 'Save starting value
            LogMsg("UTCDate Read", MessageLevel.msgOK, m_UTCDate.ToString("dd-MMM-yyyy HH:mm:ss.fff"))
            Try ' UTCDate Write is optional since if you are using the PC time as UTCTime then you should not write to the PC clock!
                If g_Settings.DisplayMethodCalls Then LogMsg("UTCDate Write", MessageLevel.msgComment, "About to set UTCDate property to " & m_UTCDate.AddHours(1.0).ToString())
                telescopeDevice.UTCDate = m_UTCDate.AddHours(1.0) 'Try and write a new UTCDate in the future
                LogMsg("UTCDate Write", MessageLevel.msgOK, "New UTCDate written successfully: " & m_UTCDate.AddHours(1.0).ToString)
                If g_Settings.DisplayMethodCalls Then LogMsg("UTCDate Write", MessageLevel.msgComment, "About to set UTCDate property to " & m_UTCDate.ToString())
                telescopeDevice.UTCDate = m_UTCDate 'Restore original value
            Catch ex As Exception
                HandleException("UTCDate Write", MemberType.Property, Required.Optional, ex, "")
            End Try
        Catch ex As Exception
            HandleException("UTCDate Read", MemberType.Property, Required.Mandatory, ex, "")
        End Try
        If TestStop() Then Exit Sub
    End Sub
    Overrides Sub CheckMethods()

        'CanMoveAxis - Required - This must be first test as Parked tests use its results
        If g_InterfaceVersion > 1 Then
            If (g_TelescopeTests.Item(TELTEST_CAN_MOVE_AXIS) = CheckState.Checked) Or
                (g_TelescopeTests.Item(TELTEST_MOVE_AXIS) = CheckState.Checked) Or
                (g_TelescopeTests.Item(TELTEST_PARK_UNPARK) = CheckState.Checked) Then
                TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisPrimary, "CanMoveAxis:Primary") : If TestStop() Then Exit Sub
                TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisSecondary, "CanMoveAxis:Secondary") : If TestStop() Then Exit Sub
                TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisTertiary, "CanMoveAxis:Tertiary") : If TestStop() Then Exit Sub
            Else
                LogMsg(TELTEST_CAN_MOVE_AXIS, MessageLevel.msgInfo, "Tests skipped")
            End If
        Else
            LogMsg("CanMoveAxis", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'Test Park, Unpark - Optional
        If g_InterfaceVersion > 1 Then
            If g_TelescopeTests.Item(TELTEST_PARK_UNPARK) = CheckState.Checked Then
                If canPark Then 'Can Park
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg("Park", MessageLevel.msgComment, "About to get AtPark property")
                        If Not telescopeDevice.AtPark Then ' OK We are unparked so check that no error is generated
                            Status(StatusType.staTest, "Park")
                            Try
                                Status(StatusType.staAction, "Park scope")
                                If g_Settings.DisplayMethodCalls Then LogMsg("Park", MessageLevel.msgComment, "About to call Park method")
                                telescopeDevice.Park()
                                Status(StatusType.staStatus, "Waiting for scope to park")
                                Do
                                    WaitFor(SLEEP_TIME)
                                    Application.DoEvents()
                                    If g_Settings.DisplayMethodCalls Then LogMsg("Park", MessageLevel.msgComment, "About to get AtPark property")
                                Loop Until telescopeDevice.AtPark Or TestStop()
                                If TestStop() Then Exit Sub
                                Status(StatusType.staStatus, "Scope parked")
                                LogMsg("Park", MessageLevel.msgOK, "Success")

                                'Scope Parked OK
                                Try 'Confirm second park is harmless
                                    If g_Settings.DisplayMethodCalls Then LogMsg("Park", MessageLevel.msgComment, "About to Park call method")
                                    telescopeDevice.Park()
                                    LogMsg("Park", MessageLevel.msgOK, "Success if already parked")
                                Catch ex As COMException
                                    LogMsg("Park", MessageLevel.msgIssue, "Exception when calling Park two times in succession: " & ex.Message & " " & Hex(ex.ErrorCode))
                                Catch ex As Exception
                                    LogMsg("Park", MessageLevel.msgIssue, "Exception when calling Park two times in succession: " & ex.Message)
                                End Try

                                'Confirm that methods do raise exceptions when scope is parked
                                If canSlew Or canSlewAsync Or canSlewAltAz Or canSlewAltAzAsync Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepAbortSlew, "AbortSlew") : If TestStop() Then Exit Sub
                                If canFindHome Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepFindHome, "FindHome") : If TestStop() Then Exit Sub
                                If m_CanMoveAxisPrimary Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisPrimary, "MoveAxis Primary") : If TestStop() Then Exit Sub
                                If m_CanMoveAxisSecondary Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisSecondary, "MoveAxis Secondary") : If TestStop() Then Exit Sub
                                If m_CanMoveAxisTertiary Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisTertiary, "MoveAxis Tertiary") : If TestStop() Then Exit Sub
                                If canPulseGuide Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepPulseGuide, "PulseGuide") : If TestStop() Then Exit Sub
                                If canSlew Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToCoordinates, "SlewToCoordinates") : If TestStop() Then Exit Sub
                                If canSlewAsync Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToCoordinatesAsync, "SlewToCoordinatesAsync") : If TestStop() Then Exit Sub
                                If canSlew Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToTarget, "SlewToTarget") : If TestStop() Then Exit Sub
                                If canSlewAsync Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToTargetAsync, "SlewToTargetAsync") : If TestStop() Then Exit Sub
                                If canSync Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSyncToCoordinates, "SyncToCoordinates") : If TestStop() Then Exit Sub
                                If canSync Then TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSyncToTarget, "SyncToTarget") : If TestStop() Then Exit Sub

                                'Test unpark after park
                                If canUnpark Then
                                    Try
                                        Status(StatusType.staAction, "UnPark scope after park")
                                        If g_Settings.DisplayMethodCalls Then LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method")
                                        telescopeDevice.Unpark()
                                        Do
                                            WaitFor(SLEEP_TIME)
                                            Application.DoEvents()
                                            If g_Settings.DisplayMethodCalls Then LogMsg("UnPark", MessageLevel.msgComment, "About to get AtPark property")
                                        Loop Until Not telescopeDevice.AtPark Or TestStop()
                                        If TestStop() Then Exit Sub
                                        Try 'Make sure tracking doesn't generate an error if it is not implemented
                                            If g_Settings.DisplayMethodCalls Then LogMsg("UnPark", MessageLevel.msgComment, "About to set Tracking property true")
                                            telescopeDevice.Tracking = True
                                        Catch ex As Exception
                                        End Try
                                        Status(StatusType.staStatus, "Scope UnParked")
                                        LogMsg("UnPark", MessageLevel.msgOK, "Success")

                                        'Scope unparked
                                        Try 'Confirm UnPark is harmless if already unparked
                                            If g_Settings.DisplayMethodCalls Then LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method")
                                            telescopeDevice.Unpark()
                                            LogMsg("UnPark", MessageLevel.msgOK, "Success if already unparked")
                                        Catch ex As COMException
                                            LogMsg("UnPark", MessageLevel.msgIssue, "Exception when calling UnPark two times in succession: " & ex.Message & " " & Hex(ex.ErrorCode))
                                        Catch ex As Exception
                                            LogMsg("UnPark", MessageLevel.msgIssue, "Exception when calling UnPark two times in succession: " & ex.Message)
                                        End Try
                                    Catch ex As COMException
                                        LogMsg("UnPark", MessageLevel.msgError, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
                                    Catch ex As Exception
                                        LogMsg("UnPark", MessageLevel.msgError, EX_NET & ex.Message)
                                    End Try
                                Else ' Can't UnPark
                                    'Confirm that UnPark generates an error
                                    Try
                                        If g_Settings.DisplayMethodCalls Then LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method")
                                        telescopeDevice.Unpark()
                                        LogMsg("UnPark", MessageLevel.msgIssue, "No exception thrown by UnPark when CanUnPark is false")
                                    Catch ex As COMException
                                        If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
                                            LogMsg("UnPark", MessageLevel.msgOK, NOT_IMP_COM)
                                        Else
                                            ExTest("UnPark", ex.Message, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
                                        End If
                                    Catch ex As MethodNotImplementedException
                                        LogMsg("UnPark", MessageLevel.msgOK, NOT_IMP_NET)
                                    Catch ex As Exception
                                        ExTest("UnPark", ex.Message, EX_NET & ex.Message)
                                    End Try
                                    'Create user interface message asking for manual scope UnPark
                                    LogMsg("UnPark", MessageLevel.msgComment, "CanUnPark is false so you need to unpark manually")
                                    MsgBox("This scope cannot be unparked automatically, please unpark it now", MsgBoxStyle.Critical, "UnPark")
                                End If
                            Catch ex As COMException
                                LogMsg("Park", MessageLevel.msgError, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
                            Catch ex As Exception
                                LogMsg("Park", MessageLevel.msgError, EX_NET & ex.Message)
                            End Try
                        Else ' We are still in parked status despite a successful UnPark
                            LogMsg("Park", MessageLevel.msgError, "AtPark still true despite an earlier successful unpark")
                        End If
                    Catch ex As Exception
                        HandleException("Park", MemberType.Method, Required.MustBeImplemented, ex, "CanPark is True")
                    End Try
                Else 'Can't park
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg("UnPark", MessageLevel.msgComment, "About to call Park method")
                        telescopeDevice.Park()
                        LogMsg("Park", MessageLevel.msgError, "CanPark is false but no exception was generated on use")
                    Catch ex As Exception
                        HandleException("Park", MemberType.Method, Required.MustNotBeImplemented, ex, "CanPark is False")
                    End Try
                    'v1.0.12.0 Added test for unpark if CanPark is false
                    'Now test unpark
                    If canUnpark Then ' We should already be unparked so confirm that unpark works fine
                        Try
                            If g_Settings.DisplayMethodCalls Then LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method")
                            telescopeDevice.Unpark()
                            LogMsg("UnPark", MessageLevel.msgOK, "CanPark is false and CanUnPark is true; no exception generated as expected")
                        Catch ex As Exception
                            HandleException("UnPark", MemberType.Method, Required.MustBeImplemented, ex, "CanUnPark is True")
                        End Try
                    Else 'Can't unpark so confirm an exception is raised
                        Try
                            If g_Settings.DisplayMethodCalls Then LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method")
                            telescopeDevice.Unpark()
                            LogMsg("UnPark", MessageLevel.msgError, "CanPark and CanUnPark are false but no exception was generated on use")
                        Catch ex As Exception
                            HandleException("UnPark", MemberType.Method, Required.MustNotBeImplemented, ex, "CanUnPark is False")
                        End Try
                    End If
                End If
                g_Status.Clear() 'Clear status messages
                If TestStop() Then Exit Sub
            Else
                LogMsg(TELTEST_PARK_UNPARK, MessageLevel.msgInfo, "Tests skipped")
            End If
        Else
            LogMsg("Park", MessageLevel.msgInfo, "Skipping tests since behaviour of this method is not well defined in interface V" & g_InterfaceVersion)
        End If

        'AbortSlew - Optional
        If g_TelescopeTests.Item(TELTEST_ABORT_SLEW) = CheckState.Checked Then
            TelescopeOptionalMethodsTest(OptionalMethodType.AbortSlew, "AbortSlew", True) : If TestStop() Then Exit Sub
        Else
            LogMsg(TELTEST_ABORT_SLEW, MessageLevel.msgInfo, "Tests skipped")
        End If

        'AxisRates - Required
        If g_InterfaceVersion > 1 Then
            If (g_TelescopeTests.Item(TELTEST_AXIS_RATE) = CheckState.Checked) Or (g_TelescopeTests.Item(TELTEST_MOVE_AXIS) = CheckState.Checked) Then
                TelescopeAxisRateTest("AxisRate:Primary", TelescopeAxes.axisPrimary)
                TelescopeAxisRateTest("AxisRate:Secondary", TelescopeAxes.axisSecondary)
                TelescopeAxisRateTest("AxisRate:Tertiary", TelescopeAxes.axisTertiary)
            Else
                LogMsg(TELTEST_AXIS_RATE, MessageLevel.msgInfo, "Tests skipped")
            End If
        Else
            LogMsg("AxisRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'FindHome - Optional
        If g_TelescopeTests.Item(TELTEST_FIND_HOME) = CheckState.Checked Then
            TelescopeOptionalMethodsTest(OptionalMethodType.FindHome, "FindHome", canFindHome) : If TestStop() Then Exit Sub
        Else
            LogMsg(TELTEST_FIND_HOME, MessageLevel.msgInfo, "Tests skipped")
        End If

        'MoveAxis - Optional
        If g_InterfaceVersion > 1 Then
            If g_TelescopeTests.Item(TELTEST_MOVE_AXIS) = CheckState.Checked Then
                TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisPrimary, "MoveAxis Primary", m_CanMoveAxisPrimary) : If TestStop() Then Exit Sub
                TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisSecondary, "MoveAxis Secondary", m_CanMoveAxisSecondary) : If TestStop() Then Exit Sub
                TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisTertiary, "MoveAxis Tertiary", m_CanMoveAxisTertiary) : If TestStop() Then Exit Sub
            Else
                LogMsg(TELTEST_MOVE_AXIS, MessageLevel.msgInfo, "Tests skipped")
            End If
        Else
            LogMsg("MoveAxis", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'PulseGuide - Optional
        If g_TelescopeTests.Item(TELTEST_PULSE_GUIDE) = CheckState.Checked Then
            TelescopeOptionalMethodsTest(OptionalMethodType.PulseGuide, "PulseGuide", canPulseGuide) : If TestStop() Then Exit Sub
        Else
            LogMsg(TELTEST_PULSE_GUIDE, MessageLevel.msgInfo, "Tests skipped")
        End If

        'Test Equatorial slewing to coordinates - Optional
        If g_TelescopeTests.Item(TELTEST_SLEW_TO_COORDINATES) = CheckState.Checked Then
            TelescopeSlewTest(SlewSyncType.SlewToCoordinates, "SlewToCoordinates", canSlew, "CanSlew") : If TestStop() Then Exit Sub
            If canSlew Then ' Test slewing to bad co-ordinates
                TelescopeBadCoordinateTest("SlewToCoordinates (Bad L)", SlewSyncType.SlewToCoordinates, BAD_RA_LOW, BAD_DEC_LOW) : If TestStop() Then Exit Sub
                TelescopeBadCoordinateTest("SlewToCoordinates (Bad H)", SlewSyncType.SlewToCoordinates, BAD_RA_HIGH, BAD_DEC_HIGH) : If TestStop() Then Exit Sub
            End If

        Else
            LogMsg(TELTEST_SLEW_TO_COORDINATES, MessageLevel.msgInfo, "Tests skipped")
        End If

        'Test Equatorial slewing to coordinates asynchronous - Optional
        If g_TelescopeTests.Item(TELTEST_SLEW_TO_COORDINATES_ASYNC) = CheckState.Checked Then
            TelescopeSlewTest(SlewSyncType.SlewToCoordinatesAsync, "SlewToCoordinatesAsync", canSlewAsync, "CanSlewAsync") : If TestStop() Then Exit Sub
            If canSlewAsync Then ' Test slewing to bad co-ordinates
                TelescopeBadCoordinateTest("SlewToCoordinatesAsync (Bad L)", SlewSyncType.SlewToCoordinatesAsync, BAD_RA_LOW, BAD_DEC_LOW) : If TestStop() Then Exit Sub
                TelescopeBadCoordinateTest("SlewToCoordinatesAsync (Bad H)", SlewSyncType.SlewToCoordinatesAsync, BAD_RA_HIGH, BAD_DEC_HIGH) : If TestStop() Then Exit Sub
            End If
        Else
            LogMsg(TELTEST_SLEW_TO_COORDINATES_ASYNC, MessageLevel.msgInfo, "Tests skipped")
        End If

        'Equatorial Sync to Coordinates - Optional - Moved here so that it can be tested before any target coordinates are set - Peter 4th August 2018
        If g_TelescopeTests.Item(TELTEST_SYNC_TO_COORDINATES) = CheckState.Checked Then
            TelescopeSyncTest(SlewSyncType.SyncToCoordinates, "SyncToCoordinates", canSync, "CanSync") : If TestStop() Then Exit Sub
            If canSync Then ' Test syncing to bad co-ordinates
                TelescopeBadCoordinateTest("SyncToCoordinates (Bad L)", SlewSyncType.SyncToCoordinates, BAD_RA_LOW, BAD_DEC_LOW) : If TestStop() Then Exit Sub
                TelescopeBadCoordinateTest("SyncToCoordinates (Bad H)", SlewSyncType.SyncToCoordinates, BAD_RA_HIGH, BAD_DEC_HIGH) : If TestStop() Then Exit Sub
            End If
        Else
            LogMsg(TELTEST_SYNC_TO_COORDINATES, MessageLevel.msgInfo, "Tests skipped")
        End If

        'TargetRightAscension Write - Optional - Test moved here so that Conform can check that the SlewTo... methods properly set target coordinates.")
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("TargetRightAscension Write", MessageLevel.msgComment, "About to set TargetRightAscension property to -1.0")
            telescopeDevice.TargetRightAscension = -1.0
            LogMsg("TargetRightAscension Write", MessageLevel.msgIssue, "No error generated on set TargetRightAscension < 0 hours")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetRightAscension < 0 hours")
        End Try

        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("TargetRightAscension Write", MessageLevel.msgComment, "About to set TargetRightAscension property to 25.0")
            telescopeDevice.TargetRightAscension = 25.0
            LogMsg("TargetRightAscension Write", MessageLevel.msgIssue, "No error generated on set TargetRightAscension > 24 hours")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetRightAscension > 24 hours")
        End Try

        Try
            m_TargetRightAscension = TelescopeRAFromSiderealTime("TargetRightAscension Write", -4.0)
            If g_Settings.DisplayMethodCalls Then LogMsg("TargetRightAscension Write", MessageLevel.msgComment, "About to set TargetRightAscension property to " & m_TargetRightAscension)
            telescopeDevice.TargetRightAscension = m_TargetRightAscension 'Set a valid value
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("TargetRightAscension Write", MessageLevel.msgComment, "About to get TargetRightAscension property")
                Select Case System.Math.Abs(telescopeDevice.TargetRightAscension - m_TargetRightAscension)
                    Case 0.0
                        LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Legal value " & FormatRA(m_TargetRightAscension) & " HH:MM:SS written successfully")
                    Case Is <= (1.0 / 3600.0) '1 seconds
                        LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 1 second of the value set: " & FormatRA(m_TargetRightAscension))
                        Exit Select
                    Case Is <= (2.0 / 3600.0) '2 seconds
                        LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 2 seconds of the value set: " & FormatRA(m_TargetRightAscension))
                        Exit Select
                    Case Is <= (5.0 / 3600.0) '5 seconds
                        LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 5 seconds of the value set: " & FormatRA(m_TargetRightAscension))
                        Exit Select
                    Case Else
                        LogMsg("TargetRightAscension Write", MessageLevel.msgInfo, "Target RightAscension: " & FormatRA(telescopeDevice.TargetRightAscension))
                End Select
            Catch ex As Exception
                HandleException("TargetRightAscension Write", MemberType.Property, Required.MustBeImplemented, ex, "Unable to read TargetRightAscension before writing to it.")
            End Try
        Catch ex As Exception
            HandleException("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'TargetDeclination Write - Optional - Test moved here so that Conform can check that the SlewTo... methods properly set target coordinates.")
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("TargetDeclination Write", MessageLevel.msgComment, "About to set TargetDeclination property to -91.0")
            telescopeDevice.TargetDeclination = -91.0
            LogMsg("TargetDeclination Write", MessageLevel.msgIssue, "No error generated on set TargetDeclination < -90 degrees")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetDeclination < -90 degrees")
        End Try

        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("TargetDeclination Write", MessageLevel.msgComment, "About to set TargetDeclination property to 91.0")
            telescopeDevice.TargetDeclination = 91.0
            LogMsg("TargetDeclination Write", MessageLevel.msgIssue, "No error generated on set TargetDeclination > 90 degrees")
        Catch ex As Exception
            HandleInvalidValueExceptionAsOK("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetDeclination < -90 degrees")
        End Try

        Try
            m_TargetDeclination = 1.0
            If g_Settings.DisplayMethodCalls Then LogMsg("TargetDeclination Write", MessageLevel.msgComment, "About to set TargetDeclination property to " & m_TargetDeclination)
            telescopeDevice.TargetDeclination = m_TargetDeclination 'Set a valid value
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("TargetDeclination Write", MessageLevel.msgComment, "About to get TargetDeclination property")
                Select Case System.Math.Abs(telescopeDevice.TargetDeclination - m_TargetDeclination)
                    Case 0.0
                        LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Legal value " & FormatDec(m_TargetDeclination) & " DD:MM:SS written successfully")
                    Case Is <= (1.0 / 3600.0) '1 seconds
                        LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 1 second of the value set: " & FormatDec(m_TargetDeclination))
                        Exit Select
                    Case Is <= (2.0 / 3600.0) '2 seconds
                        LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 2 seconds of the value set: " & FormatDec(m_TargetDeclination))
                        Exit Select
                    Case Is <= (5.0 / 3600.0) '5 seconds
                        LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 5 seconds of the value set: " & FormatDec(m_TargetDeclination))
                        Exit Select
                    Case Else
                        LogMsg("TargetDeclination Write", MessageLevel.msgInfo, "Target Declination: " & FormatDec(m_TargetDeclination))
                End Select
            Catch ex As Exception
                HandleException("TargetDeclination Write", MemberType.Property, Required.MustBeImplemented, ex, "Unable to read TargetDeclination before writing to it.")
            End Try
        Catch ex As Exception
            HandleException("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'Test Equatorial target slewing - Optional
        If g_TelescopeTests.Item(TELTEST_SLEW_TO_TARGET) = CheckState.Checked Then
            TelescopeSlewTest(SlewSyncType.SlewToTarget, "SlewToTarget", canSlew, "CanSlew") : If TestStop() Then Exit Sub
            If canSlew Then ' Test slewing to bad co-ordinates
                TelescopeBadCoordinateTest("SlewToTarget (Bad L)", SlewSyncType.SlewToTarget, BAD_RA_LOW, BAD_DEC_LOW) : If TestStop() Then Exit Sub
                TelescopeBadCoordinateTest("SlewToTarget (Bad H)", SlewSyncType.SlewToTarget, BAD_RA_HIGH, BAD_DEC_HIGH) : If TestStop() Then Exit Sub
            End If
        Else
            LogMsg(TELTEST_SLEW_TO_TARGET, MessageLevel.msgInfo, "Tests skipped")
        End If

        'Test Equatorial target slewing asynchronous - Optional
        If g_TelescopeTests.Item(TELTEST_SLEW_TO_TARGET_ASYNC) = CheckState.Checked Then
            TelescopeSlewTest(SlewSyncType.SlewToTargetAsync, "SlewToTargetAsync", canSlewAsync, "CanSlewAsync") : If TestStop() Then Exit Sub
            If canSlewAsync Then ' Test slewing to bad co-ordinates
                TelescopeBadCoordinateTest("SlewToTargetAsync (Bad L)", SlewSyncType.SlewToTargetAsync, BAD_RA_LOW, BAD_DEC_LOW) : If TestStop() Then Exit Sub
                TelescopeBadCoordinateTest("SlewToTargetAsync (Bad H)", SlewSyncType.SlewToTargetAsync, BAD_RA_HIGH, BAD_DEC_HIGH) : If TestStop() Then Exit Sub
            End If
        Else
            LogMsg(TELTEST_SLEW_TO_TARGET_ASYNC, MessageLevel.msgInfo, "Tests skipped")
        End If

        'DestinationSideOfPier - Optional
        If g_InterfaceVersion > 1 Then
            If g_TelescopeTests.Item(TELTEST_DESTINATION_SIDE_OF_PIER) = CheckState.Checked Then
                If m_AlignmentMode = AlignmentModes.algGermanPolar Then
                    TelescopeOptionalMethodsTest(OptionalMethodType.DestinationSideOfPier, "DestinationSideOfPier", True) : If TestStop() Then Exit Sub
                Else
                    LogMsg("DestinationSideOfPier", MessageLevel.msgComment, "Test skipped as AligmentMode is not German Polar")
                End If
            Else
                LogMsg(TELTEST_DESTINATION_SIDE_OF_PIER, MessageLevel.msgInfo, "Tests skipped")
            End If
        Else
            LogMsg("DestinationSideOfPier", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'Test AltAz Slewing - Optional
        If g_InterfaceVersion > 1 Then
            If g_TelescopeTests.Item(TELTEST_SLEW_TO_ALTAZ) = CheckState.Checked Then
                TelescopeSlewTest(SlewSyncType.SlewToAltAz, "SlewToAltAz", canSlewAltAz, "CanSlewAltAz") : If TestStop() Then Exit Sub
                If canSlewAltAz Then ' Test slewing to bad co-ordinates
                    TelescopeBadCoordinateTest("SlewToAltAz (Bad L)", SlewSyncType.SlewToAltAz, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW) : If TestStop() Then Exit Sub ' -100 is used for the Altitude limit to enable -90 to be used for parking the scope
                    TelescopeBadCoordinateTest("SlewToAltAz (Bad H)", SlewSyncType.SlewToAltAz, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH) : If TestStop() Then Exit Sub
                End If
            Else
                LogMsg(TELTEST_SLEW_TO_ALTAZ, MessageLevel.msgInfo, "Tests skipped")
            End If
        Else
            LogMsg("SlewToAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'Test AltAz Slewing asynchronous - Optional
        If g_InterfaceVersion > 1 Then
            If g_TelescopeTests.Item(TELTEST_SLEW_TO_ALTAZ_ASYNC) = CheckState.Checked Then
                TelescopeSlewTest(SlewSyncType.SlewToAltAzAsync, "SlewToAltAzAsync", canSlewAltAzAsync, "CanSlewAltAzAsync") : If TestStop() Then Exit Sub
                If canSlewAltAzAsync Then ' Test slewing to bad co-ordinates
                    TelescopeBadCoordinateTest("SlewToAltAzAsync (Bad L)", SlewSyncType.SlewToAltAzAsync, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW) : If TestStop() Then Exit Sub
                    TelescopeBadCoordinateTest("SlewToAltAzAsync (Bad H)", SlewSyncType.SlewToAltAzAsync, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH) : If TestStop() Then Exit Sub
                End If
            Else
                LogMsg(TELTEST_SLEW_TO_ALTAZ_ASYNC, MessageLevel.msgInfo, "Tests skipped")
            End If
        Else
            LogMsg("SlewToAltAzAsync", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'Equatorial Sync to Target - Optional
        If g_TelescopeTests.Item(TELTEST_SYNC_TO_TARGET) = CheckState.Checked Then

            TelescopeSyncTest(SlewSyncType.SyncToTarget, "SyncToTarget", canSync, "CanSync") : If TestStop() Then Exit Sub
            If canSync Then ' Test syncing to bad co-ordinates
                TelescopeBadCoordinateTest("SyncToTarget (Bad L)", SlewSyncType.SyncToTarget, BAD_RA_LOW, BAD_DEC_LOW) : If TestStop() Then Exit Sub
                TelescopeBadCoordinateTest("SyncToTarget (Bad H)", SlewSyncType.SyncToTarget, BAD_RA_HIGH, BAD_DEC_HIGH) : If TestStop() Then Exit Sub
            End If
        Else
            LogMsg(TELTEST_SYNC_TO_TARGET, MessageLevel.msgInfo, "Tests skipped")
        End If

        'AltAz Sync - Optional
        If g_InterfaceVersion > 1 Then
            If g_TelescopeTests.Item(TELTEST_SYNC_TO_ALTAZ) = CheckState.Checked Then
                TelescopeSyncTest(SlewSyncType.SyncToAltAz, "SyncToAltAz", canSyncAltAz, "CanSyncAltAz") : If TestStop() Then Exit Sub
                If canSyncAltAz Then ' Test syncing to bad co-ordinates
                    TelescopeBadCoordinateTest("SyncToAltAz (Bad L)", SlewSyncType.SyncToAltAz, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW) : If TestStop() Then Exit Sub
                    TelescopeBadCoordinateTest("SyncToAltAz (Bad H)", SlewSyncType.SyncToAltAz, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH) : If TestStop() Then Exit Sub
                End If
            Else
                LogMsg(TELTEST_SYNC_TO_ALTAZ, MessageLevel.msgInfo, "Tests skipped")
            End If
        Else
            LogMsg("SyncToAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If



        If g_Settings.TestSideOfPierRead Then
            LogMsg("", MessageLevel.msgAlways, "")
            LogMsg("SideOfPier Model Tests", MessageLevel.msgAlways, "")
            LogMsg("SideOfPier Model Tests", MessageLevel.msgDebug, "Starting tests")

            If g_InterfaceVersion > 1 Then
                '3.0.0.14 - Skip these tests if unable to read SideOfPier
                If m_CanReadSideOfPier Then

                    'Further side of pier tests
                    If g_Settings.DisplayMethodCalls Then LogMsg("SideOfPier Model Tests", MessageLevel.msgComment, "About to get AlignmentMode property")
                    If telescopeDevice.AlignmentMode = AlignmentModes.algGermanPolar Then
                        LogMsg("SideOfPier Model Tests", MessageLevel.msgDebug, "Calling SideOfPierTests()")
                        Select Case m_SiteLatitude
                            Case -SIDE_OF_PIER_INVALID_LATITUDE To SIDE_OF_PIER_INVALID_LATITUDE ' Refuse to handle this value because the Conform targeting logic or the mount's SideofPier flip logic may fail when the poles are this close to the horizon
                                LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Tests skipped because the site latitude is reported as " & g_Util.DegreesToDMS(m_SiteLatitude, ":", ":", "", 3))
                                LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "This places the celestial poles close to the horizon and the mount's flip logic may override Conform's expected behaviour.")
                                LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Please set the site latitude to a value within the ranges " & SIDE_OF_PIER_INVALID_LATITUDE.ToString("+0.0;-0.0") & " to +90.0 or " & (-SIDE_OF_PIER_INVALID_LATITUDE).ToString("+0.0;-0.0") & " to -90.0 to obtain a reliable result.")
                                Exit Select
                            Case -90.0 To 90.0 ' Normal case, just run the tests barbecue latitude is outside the invalid range but within -90.0 to +90.0
                                'SideOfPier write property test - Optional
                                If g_Settings.TestSideOfPierWrite Then
                                    LogMsg("SideOfPier Model Tests", MessageLevel.msgDebug, "Testing SideOfPier write...")
                                    TelescopeOptionalMethodsTest(OptionalMethodType.SideOfPierWrite, "SideOfPier Write", canSetPierside)
                                    If TestStop() Then Exit Sub
                                End If
                                SideOfPierTests() 'Only run these for German mounts
                                Exit Select
                            Case Else ' Values outside the range -90.0 to +90.0 are invalid
                                LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Test skipped because the site latitude Is outside the range -90.0 to +90.0")
                        End Select
                    Else
                        LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Test skipped because this Is Not a German equatorial mount")
                    End If
                Else
                    LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Tests skipped because this driver does Not support SideOfPier Read")
                End If
            Else
                LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Skipping test as this method Is Not supported in interface V" & g_InterfaceVersion)
            End If
        End If
        g_Status.Clear()  'Clear status messages

    End Sub
    Public Overrides Sub CheckPerformance()
        Status(StatusType.staTest, "Performance") 'Clear status messages
        TelescopePerformanceTest(PerformanceType.tstPerfAltitude, "Altitude") : If TestStop() Then Exit Sub
        If g_InterfaceVersion > 1 Then
            TelescopePerformanceTest(PerformanceType.tstPerfAtHome, "AtHome") : If TestStop() Then Exit Sub
        Else
            LogMsg("Performance: AtHome", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        If g_InterfaceVersion > 1 Then
            TelescopePerformanceTest(PerformanceType.tstPerfAtPark, "AtPark") : If TestStop() Then Exit Sub
        Else
            LogMsg("Performance: AtPark", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If
        TelescopePerformanceTest(PerformanceType.tstPerfAzimuth, "Azimuth") : If TestStop() Then Exit Sub
        TelescopePerformanceTest(PerformanceType.tstPerfDeclination, "Declination") : If TestStop() Then Exit Sub
        If g_InterfaceVersion > 1 Then
            If canPulseGuide Then
                TelescopePerformanceTest(PerformanceType.tstPerfIsPulseGuiding, "IsPulseGuiding") : If TestStop() Then Exit Sub
            Else
                LogMsg("Performance: IsPulseGuiding", MessageLevel.msgInfo, "Test omitted since IsPulseGuiding is not implemented")
            End If
        Else
            LogMsg("Performance: IsPulseGuiding", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface v1" & g_InterfaceVersion)
        End If
        TelescopePerformanceTest(PerformanceType.tstPerfRightAscension, "RightAscension") : If TestStop() Then Exit Sub
        If g_InterfaceVersion > 1 Then
            If m_AlignmentMode = AlignmentModes.algGermanPolar Then
                If m_CanReadSideOfPier Then
                    TelescopePerformanceTest(PerformanceType.tstPerfSideOfPier, "SideOfPier") : If TestStop() Then Exit Sub
                Else
                    LogMsg("Performance: SideOfPier", MessageLevel.msgInfo, "Test omitted since SideOfPier is not implemented")
                End If
            Else
                LogMsg("Performance: SideOfPier", MessageLevel.msgInfo, "Test omitted since alignment mode is not German Polar")
            End If
        Else
            LogMsg("Performance: SideOfPier", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface v1" & g_InterfaceVersion)
        End If
        If canReadSiderealTime Then
            TelescopePerformanceTest(PerformanceType.tstPerfSiderealTime, "SiderealTime") : If TestStop() Then Exit Sub
        Else
            LogMsgInfo("Performance: SiderealTime", "Skipping test because the SiderealTime property throws an exception.")
        End If
        TelescopePerformanceTest(PerformanceType.tstPerfSlewing, "Slewing") : If TestStop() Then Exit Sub
        TelescopePerformanceTest(PerformanceType.tstPerfUTCDate, "UTCDate") : If TestStop() Then Exit Sub
        g_Status.Clear()
    End Sub
    Public Overrides Sub PostRunCheck()
        'Make things safe
        ' LogMsg("", MessageLevel.msgAlways, "") 'Blank line
        Try
            If telescopeDevice.CanSetTracking Then
                telescopeDevice.Tracking = False
                LogMsg("Mount Safety", MessageLevel.msgOK, "Tracking stopped to protect your mount.")
            Else
                LogMsg("Mount Safety", MessageLevel.msgInfo, "Tracking can't be turned off for this mount, please switch off manually.")
            End If
        Catch ex As Exception
            LogMsg("Mount Safety", MessageLevel.msgError, "Exception when disabling tracking to protect mount: " & ex.ToString)
        End Try
    End Sub

    Private Sub TelescopeSyncTest(ByVal testType As SlewSyncType, ByVal testName As String, ByVal driverSupportsMethod As Boolean, ByVal canDoItName As String)
        Dim showOutcome As Boolean = False
        Dim difference, syncRA, syncDEC, syncAlt, syncAz, newAlt, newAz, currentAz, currentAlt, startRA, startDec, currentRA, currentDec As Double
        Dim trackingState As Boolean

        ' Basic test to make sure the method is either implemented OK or fails as expected if it is not supported in this driver.
        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get RightAscension property")
        syncRA = telescopeDevice.RightAscension
        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Declination property")
        syncDEC = telescopeDevice.Declination

        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, $"About to get Tracking property")
        trackingState = telescopeDevice.Tracking

        'Enable tracking where appropriate
        Select Case testName
            Case SlewSyncType.SyncToCoordinates, SlewSyncType.SyncToTarget
                If canSetTracking And trackingCanActuallybeChanged Then
                    If canSetTracking And (Not trackingState) Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to true")
                        telescopeDevice.Tracking = True
                        trackingState = True
                    End If
                Else
                    LogMsgDebug(testName, $"Not setting tracking - CanSetTracking:  {canSetTracking}, TrackingCanActuallyBeChanged: {trackingCanActuallybeChanged}")
                End If
            Case SlewSyncType.SyncToAltAz
                If canSetTracking And trackingCanActuallybeChanged Then
                    If canSetTracking And (Not trackingState) Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to false")
                        telescopeDevice.Tracking = False
                        trackingState = False
                    End If
                Else
                    LogMsgDebug(testName, $"Not setting tracking - CanSetTracking:  {canSetTracking}, TrackingCanActuallyBeChanged: {trackingCanActuallybeChanged}")
                End If

            Case Else
                ' Does not need to be enabled for remaining tests so no action
        End Select

        If Not driverSupportsMethod Then ' Call should fail
            Try
                Select Case testType
                    Case SlewSyncType.SyncToCoordinates 'SyncToCoordinates
                        LogMsg(testName, MessageLevel.msgDebug, "SyncToCoordinates: " & FormatRA(syncRA) & " " & FormatDec(syncDEC))
                        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to call SyncToCoordinates method, RA: " & FormatRA(syncRA) & ", Declination: " & FormatDec(syncDEC))
                        telescopeDevice.SyncToCoordinates(syncRA, syncDEC)
                        LogMsg(testName, MessageLevel.msgError, "CanSyncToCoordinates is False but call to SyncToCoordinates did not throw an exception.")
                    Case SlewSyncType.SyncToTarget 'SyncToTarget
                        Try
                            LogMsg(testName, MessageLevel.msgDebug, "Setting TargetRightAscension: " & FormatRA(syncRA))
                            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to set TargetRightAscension property to " & FormatRA(syncRA))
                            telescopeDevice.TargetRightAscension = syncRA
                            LogMsg(testName, MessageLevel.msgDebug, "Completed Set TargetRightAscension")
                        Catch ex As Exception
                            ' Ignore errors at this point as we aren't trying to test Telescope.TargetRightAscension
                        End Try
                        Try
                            LogMsg(testName, MessageLevel.msgDebug, "Setting TargetDeclination: " & FormatDec(syncDEC))
                            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to set TargetDeclination property to " & FormatDec(syncDEC))
                            telescopeDevice.TargetDeclination = syncDEC
                            LogMsg(testName, MessageLevel.msgDebug, "Completed Set TargetDeclination")
                        Catch ex As Exception
                            ' Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                        End Try

                        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to call SyncToTarget method")
                        telescopeDevice.SyncToTarget() 'Sync to target coordinates
                        LogMsg(testName, MessageLevel.msgError, "CanSyncToTarget is False but call to SyncToTarget did not throw an exception.")
                    Case SlewSyncType.SyncToAltAz
                        If canReadAltitide Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Altitude property")
                            syncAlt = telescopeDevice.Altitude
                        End If
                        If canReadAzimuth Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Azimuth property")
                            syncAz = telescopeDevice.Azimuth
                        End If
                        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to call SyncToAltAz method, Altitude: " & FormatDec(syncAlt) & ", Azimuth: " & FormatDec(syncAz))
                        telescopeDevice.SyncToAltAz(syncAz, syncAlt) 'Sync to new Alt Az
                        LogMsg(testName, MessageLevel.msgError, "CanSyncToAltAz is False but call to SyncToAltAz did not throw an exception.")
                    Case Else
                        LogMsg(testName, MessageLevel.msgError, "Conform:SyncTest: Unknown test type " & testType.ToString)
                End Select

            Catch ex As Exception
                HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, canDoItName & " is False")
            End Try
        Else ' Call should work
            Try
                Select Case testType
                    Case SlewSyncType.SyncToCoordinates, SlewSyncType.SyncToTarget 'Only do this for equatorial syncs

                        ' Calculate the Sync test RA position
                        startRA = TelescopeRAFromHourAngle(testName, +3.0)
                        LogMsg(testName, MessageLevel.msgDebug, String.Format("RA for sync tests: {0}", FormatRA(startRA)))

                        ' Calculate the Sync test DEC position
                        If m_SiteLatitude > 0.0 Then ' We are in the northern hemisphere
                            startDec = 90.0 - ((180.0 - m_SiteLatitude) * 0.5) ' Calculate for northern hemisphere
                        Else ' We are in the southern hemisphere
                            startDec = -90.0 + ((180.0 + m_SiteLatitude) * 0.5) ' Calculate for southern hemisphere
                        End If
                        LogMsg(testName, MessageLevel.msgDebug, String.Format("Declination for sync tests: {0}", FormatDec(startDec)))

                        SlewScope(startRA, startDec, String.Format("Start position - RA: {0}, Dec: {1}", FormatRA(startRA), FormatDec(startDec)))
                        If TestStop() Then Exit Sub

                        ' Now test that we have actually arrived
                        CheckScopePosition(testName, "Slewed to start position", startRA, startDec)

                        ' Calculate the sync test RA coordinate as a variation from the current RA coordinate
                        syncRA = startRA - SYNC_SIMULATED_ERROR / (15.0 * 60.0) ' Convert sync error in arc minutes to RA hours
                        If syncRA < 0.0 Then syncRA = syncRA + 24.0 'Ensure legal RA

                        ' Calculate the sync test DEC coordinate as a variation from the current DEC coordinate
                        syncDEC = startDec - SYNC_SIMULATED_ERROR / 60.0 ' Convert sync error in arc minutes to degrees

                        ' Sync the scope to the offset RA and DEC coordinates
                        SyncScope(testName, canDoItName, testType, syncRA, syncDEC)

                        ' Check that the scope's synchronised position is as expected
                        CheckScopePosition(testName, "Synced to sync position", syncRA, syncDEC)

                        ' Check that the TargetRA and TargetDec were 
                        If testType = SlewSyncType.SyncToCoordinates Then
                            ' Check that target coordinates are present and set correctly per the ASCOM Telescope specification
                            Try
                                currentRA = telescopeDevice.TargetRightAscension
                                LogMsg(testName, MessageLevel.msgDebug, String.Format("Current TargetRightAscension: {0}, Set TargetRightAscension: {1}", currentRA, syncRA))

                                Dim raDifference As Double
                                raDifference = RaDifferenceInSeconds(syncRA, currentRA)

                                Select Case raDifference
                                    Case Is <= SLEW_SYNC_OK_TOLERANCE  ' Within specified tolerance
                                        LogMsg(testName, MessageLevel.msgOK, String.Format("The TargetRightAscension property {0} matches the expected RA OK. ", FormatRA(syncRA)))
                                    Case Else ' Outside specified tolerance
                                        LogMsg(testName, MessageLevel.msgError, String.Format("The TargetRightAscension property {0} does not match the expected RA {1}", FormatRA(currentRA), FormatRA(syncRA)))
                                End Select

                            Catch ex As COMException When (ex.ErrorCode = ErrorCodes.ValueNotSet) Or (ex.ErrorCode = g_ExNotSet1) Or (ex.ErrorCode = g_ExNotSet2)
                                LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet COM exception was thrown instead.")
                            Catch ex As ASCOM.InvalidOperationException
                                LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, An InvalidOperationException was thrown instead.")
                            Catch ex As DriverException When (ex.Number = ErrorCodes.ValueNotSet) Or (ex.Number = g_ExNotSet1) Or (ex.Number = g_ExNotSet2)
                                LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.")
                            Catch ex As Exception
                                HandleException(testName, MemberType.Property, Required.Mandatory, ex, "")
                            End Try

                            Try
                                currentDec = telescopeDevice.TargetDeclination
                                LogMsg(testName, MessageLevel.msgDebug, String.Format("Current TargetDeclination: {0}, Set TargetDeclination: {1}", currentDec, syncDEC))

                                Dim decDifference As Double
                                decDifference = Math.Round(Math.Abs(currentDec - syncDEC) * 60.0 * 60.0, 1, MidpointRounding.AwayFromZero) ' Dec difference is in arc seconds from degrees of Declination

                                Select Case decDifference
                                    Case Is <= SLEW_SYNC_OK_TOLERANCE ' Within specified tolerance
                                        LogMsg(testName, MessageLevel.msgOK, String.Format("The TargetDeclination property {0} matches the expected Declination OK. ", FormatDec(syncDEC)))
                                    Case Else ' Outside specified tolerance
                                        LogMsg(testName, MessageLevel.msgError, String.Format("The TargetDeclination property {0} does not match the expected Declination {1}", FormatDec(currentDec), FormatDec(syncDEC)))
                                End Select
                            Catch ex As COMException When (ex.ErrorCode = ErrorCodes.ValueNotSet) Or (ex.ErrorCode = g_ExNotSet1) Or (ex.ErrorCode = g_ExNotSet2)
                                LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet COM exception was thrown instead.")
                            Catch ex As ASCOM.InvalidOperationException
                                LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetDeclination property as required by the Telescope specification, An InvalidOperationException was thrown instead.")
                            Catch ex As DriverException When (ex.Number = ErrorCodes.ValueNotSet) Or (ex.Number = g_ExNotSet1) Or (ex.Number = g_ExNotSet2)
                                LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.")
                            Catch ex As Exception
                                HandleException(testName, MemberType.Property, Required.Mandatory, ex, "")
                            End Try
                        End If

                        ' Now slew to the scope's original position
                        SlewScope(startRA, startDec, String.Format("Slewing back to start position - RA: {0}, Dec: {1}", FormatRA(startRA), FormatDec(startDec)))

                        ' Check that the scope's position is the original position
                        CheckScopePosition(testName, "Slewed back to start position", startRA, startDec)

                        ' Now "undo" the sync by reversing syncing in the opposition sense than originally made

                        ' Calculate the sync test RA coordinate as a variation from the current RA coordinate
                        syncRA = startRA + SYNC_SIMULATED_ERROR / (15.0 * 60.0) ' Convert sync error in arc minutes to RA hours
                        If syncRA >= 24.0 Then syncRA = syncRA - 24.0 'Ensure legal RA

                        ' Calculate the sync test DEC coordinate as a variation from the current DEC coordinate
                        syncDEC = startDec + SYNC_SIMULATED_ERROR / 60.0 ' Convert sync error in arc minutes to degrees

                        ' Sync back to the original coordinates
                        SyncScope(testName, canDoItName, testType, syncRA, syncDEC)

                        ' Check that the scope's synchronised position is as expected
                        CheckScopePosition(testName, "Synced to reversed sync position", syncRA, syncDEC)

                        ' Now slew to the scope's original position
                        SlewScope(startRA, startDec, String.Format("Slewing back to start position - RA: {0}, Dec: {1}", FormatRA(startRA), FormatDec(startDec)))

                        ' Check that the scope's position is the original position
                        CheckScopePosition(testName, "Slewed back to start position", startRA, startDec)

                    Case SlewSyncType.SyncToAltAz
                        If canReadAltitide Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Altitude property")
                            currentAlt = telescopeDevice.Altitude
                        End If
                        If canReadAzimuth Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Azimuth property")
                            currentAz = telescopeDevice.Azimuth
                        End If
                        syncAlt = currentAlt - 1.0
                        syncAz = currentAz + 1.0
                        If syncAlt < 0.0 Then syncAlt = 1.0 'Ensure legal Alt
                        If syncAz > 359.0 Then syncAz = 358.0 'Ensure legal Az
                        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to call SyncToAltAz method, Altitude: " & FormatDec(syncAlt) & ", Azimuth: " & FormatDec(syncAz))
                        telescopeDevice.SyncToAltAz(syncAz, syncAlt) 'Sync to new Alt Az
                        If canReadAltitide And canReadAzimuth Then 'Can check effects of a sync
                            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Altitude property")
                            newAlt = telescopeDevice.Altitude
                            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Azimuth property")
                            newAz = telescopeDevice.Azimuth

                            'Compare old and new values
                            difference = Math.Abs(syncAlt - newAlt)
                            Select Case difference
                                Case Is <= (1.0 / (60 * 60)) 'Within 1 seconds
                                    LogMsg(testName, MessageLevel.msgOK, "Synced Altitude OK")
                                    Exit Select
                                Case Is <= (2.0 / (60 * 60)) 'Within 2 seconds
                                    LogMsg(testName, MessageLevel.msgOK, "Synced within 2 seconds of Altitude")
                                    showOutcome = True
                                    Exit Select
                                Case Else
                                    LogMsg(testName, MessageLevel.msgInfo, "Synced to within " & FormatAltitude(difference) & " DD:MM:SS of expected Altitude: " & FormatAltitude(syncAlt))
                                    showOutcome = True
                            End Select
                            difference = Math.Abs(syncAz - newAz)
                            Select Case difference
                                Case Is <= (1.0 / (60 * 60)) 'Within 1 seconds
                                    LogMsg(testName, MessageLevel.msgOK, "Synced Azimuth OK")
                                    Exit Select
                                Case Is <= (2.0 / (60 * 60)) 'Within 2 seconds
                                    LogMsg(testName, MessageLevel.msgOK, "Synced within 2 seconds of Azimuth")
                                    showOutcome = True
                                    Exit Select
                                Case Else
                                    LogMsg(testName, MessageLevel.msgInfo, "Synced to within " & FormatAzimuth(difference) & " DD:MM:SS of expected Azimuth: " & FormatAzimuth(syncAz))
                                    showOutcome = True
                            End Select
                            If showOutcome Then
                                LogMsg(testName, MessageLevel.msgComment, "           Altitude    Azimuth")
                                LogMsg(testName, MessageLevel.msgComment, "Original:  " & FormatAltitude(currentAlt) & "   " & FormatAzimuth(currentAz))
                                LogMsg(testName, MessageLevel.msgComment, "Sync to:   " & FormatAltitude(syncAlt) & "   " & FormatAzimuth(syncAz))
                                LogMsg(testName, MessageLevel.msgComment, "New:       " & FormatAltitude(newAlt) & "   " & FormatAzimuth(newAz))
                            End If
                        Else 'Can't test effects of a sync
                            LogMsg(testName, MessageLevel.msgInfo, "Can't test SyncToAltAz because Altitude or Azimuth values are not implemented")
                        End If
                    Case Else 'Do nothing
                End Select
            Catch ex As Exception
                LogMsgDebug(testName, $"Test type: {testType}, Tracking state: {trackingState}, Driver Supports Method: {driverSupportsMethod}")

                Select Case testType
                    Case SlewSyncType.SyncToCoordinates, SlewSyncType.SyncToTarget
                        If trackingState Then ' The mount is tracking so test for normal operation
                            If driverSupportsMethod Then ' Mount says that it does support this capability
                                HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, canDoItName & " is True")
                            Else ' Mount says that it does not support this capability
                                HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, canDoItName & " is False")
                            End If
                        Else ' The mount is not tracking so expect an InValidOperationException or MethodNotImplementedException
                            If driverSupportsMethod Then 'Mount says that it does support this capability
                                HandleInvalidOperationExceptionAsOK(testName, MemberType.Method, Required.MustNotBeImplemented, ex, "An ASCOM.InvalidOperationException should be thrown when Tracking is false", $"{testName} correctly rejected the operation because Tracking is not enabled")
                            Else ' Mount says that it does not support this capability
                                HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, canDoItName & " is False")
                            End If
                        End If
                    Case SlewSyncType.SyncToAltAz
                        If Not trackingState Then ' The mount is stationary so test for normal operation
                            If driverSupportsMethod Then ' Mount says that it does support this capability
                                HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, canDoItName & " is True")
                            Else ' Mount says that it does not support this capability
                                HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, canDoItName & " is False")
                            End If
                        Else ' The mount is not tracking so expect an InValidOperationException or MethodNotImplementedException
                            If driverSupportsMethod Then 'Mount says that it does support this capability
                                HandleInvalidOperationExceptionAsOK(testName, MemberType.Method, Required.MustBeImplemented, ex, "An ASCOM.InvalidOperationException should be thrown when Tracking is false", $"{testName} correctly rejected the operation because Tracking is enabled")
                            Else ' Mount says that it does not support this capability
                                HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, canDoItName & " is False")
                            End If
                        End If
                    Case Else
                        ' No other cases
                End Select

            End Try

        End If
    End Sub

    Private Sub TelescopeSlewTest(ByVal p_Test As SlewSyncType, ByVal p_Name As String, ByVal p_CanDoIt As Boolean, ByVal p_CanDoItName As String)
        Dim l_Difference, l_ActualAltitude, l_ActualAzimuth, actualRA, actualDec As Double
        Dim trackingState As Boolean

        Status(StatusType.staTest, p_Name)

        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, $"About to get Tracking property")
        trackingState = telescopeDevice.Tracking
        LogMsgDebug(p_Name, $"Tracking: {trackingState}")

        'Enable tracking where appropriate
        Select Case p_Test
            Case SlewSyncType.SlewToCoordinates, SlewSyncType.SlewToCoordinatesAsync, SlewSyncType.SlewToTarget, SlewSyncType.SlewToTargetAsync, SlewSyncType.SyncToCoordinates, SlewSyncType.SyncToTarget
                If canSetTracking And trackingCanActuallybeChanged Then
                    If canSetTracking And (Not trackingState) Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true")
                        telescopeDevice.Tracking = True
                        trackingState = True
                    End If
                Else
                    LogMsgDebug(p_Name, $"Not setting tracking - Tracking: {trackingState}, CanSetTracking:  {canSetTracking}, TrackingCanActuallyBeChanged: {trackingCanActuallybeChanged}")
                End If
            Case SlewSyncType.SlewToAltAz, SlewSyncType.SlewToAltAzAsync
                If canSetTracking And trackingCanActuallybeChanged Then
                    If canSetTracking And (Not trackingState) Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to false")
                        telescopeDevice.Tracking = False
                        trackingState = False
                    End If
                Else
                    LogMsgDebug(p_Name, $"Not setting tracking - Tracking: {trackingState}, CanSetTracking:  {canSetTracking}, TrackingCanActuallyBeChanged: {trackingCanActuallybeChanged}")
                End If

            Case Else
                ' Does not need to be enabled for remaining tests so no action
        End Select

        Try
            Select Case p_Test
                Case SlewSyncType.SlewToCoordinates
                    m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -1.0)
                    m_TargetDeclination = 1.0
                    Status(StatusType.staAction, "Slewing")
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinates method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(m_TargetDeclination))
                    telescopeDevice.SlewToCoordinates(m_TargetRightAscension, m_TargetDeclination)
                Case SlewSyncType.SlewToCoordinatesAsync
                    m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0)
                    m_TargetDeclination = 2.0
                    Status(StatusType.staAction, "Slewing")
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(m_TargetDeclination))
                    telescopeDevice.SlewToCoordinatesAsync(m_TargetRightAscension, m_TargetDeclination)
                    WaitForSlew(p_Name)
                Case SlewSyncType.SlewToTarget
                    m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0)
                    m_TargetDeclination = 3.0

                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " & FormatRA(m_TargetRightAscension))
                        telescopeDevice.TargetRightAscension = m_TargetRightAscension
                    Catch ex As Exception
                        If p_CanDoIt Then
                            HandleException(p_Name, MemberType.Property, Required.MustBeImplemented, ex, p_CanDoItName & " is True but can't set TargetRightAscension")
                        Else
                            ' Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                        End If
                    End Try

                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " & FormatDec(m_TargetDeclination))
                        telescopeDevice.TargetDeclination = m_TargetDeclination
                    Catch ex As Exception
                        If p_CanDoIt Then
                            HandleException(p_Name, MemberType.Property, Required.MustBeImplemented, ex, p_CanDoItName & " is True but can't set TargetDeclination")
                        Else
                            ' Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                        End If
                    End Try
                    Status(StatusType.staAction, "Slewing")
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTarget method")
                    telescopeDevice.SlewToTarget()

                Case SlewSyncType.SlewToTargetAsync 'SlewToTargetAsync
                    m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -4.0)
                    m_TargetDeclination = 4.0
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " & FormatRA(m_TargetRightAscension))
                        telescopeDevice.TargetRightAscension = m_TargetRightAscension
                    Catch ex As Exception
                        If p_CanDoIt Then
                            HandleException(p_Name, MemberType.Property, Required.MustBeImplemented, ex, p_CanDoItName & " is True but can't set TargetRightAscension")
                        Else
                            ' Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                        End If
                    End Try
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " & FormatDec(m_TargetDeclination))
                        telescopeDevice.TargetDeclination = m_TargetDeclination
                    Catch ex As Exception
                        If p_CanDoIt Then
                            HandleException(p_Name, MemberType.Property, Required.MustBeImplemented, ex, p_CanDoItName & " is True but can't set TargetDeclination")
                        Else
                            ' Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                        End If
                    End Try
                    Status(StatusType.staAction, "Slewing")
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTargetAsync method")
                    telescopeDevice.SlewToTargetAsync()
                    WaitForSlew(p_Name)
                Case SlewSyncType.SlewToAltAz
                    LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 1: " & telescopeDevice.Tracking)
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                    If canSetTracking And telescopeDevice.Tracking Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set property Tracking to false")
                        telescopeDevice.Tracking = False
                        LogMsg(p_Name, MessageLevel.msgDebug, "Tracking turned off")
                    End If
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                    LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 2: " & telescopeDevice.Tracking)
                    m_TargetAltitude = 50.0
                    m_TargetAzimuth = 150.0
                    Status(StatusType.staAction, "Slewing to Alt/Az: " & FormatDec(m_TargetAltitude) & " " & FormatDec(m_TargetAzimuth))
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAz method, Altitude: " & FormatDec(m_TargetAltitude) & ", Azimuth: " & FormatDec(m_TargetAzimuth))
                    telescopeDevice.SlewToAltAz(m_TargetAzimuth, m_TargetAltitude)
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                    LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 3: " & telescopeDevice.Tracking)
                Case SlewSyncType.SlewToAltAzAsync
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                    LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 1: " & telescopeDevice.Tracking)
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                    If canSetTracking And telescopeDevice.Tracking Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property false")
                        telescopeDevice.Tracking = False
                        LogMsg(p_Name, MessageLevel.msgDebug, "Tracking turned off")
                    End If
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                    LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 2: " & telescopeDevice.Tracking)
                    m_TargetAltitude = 55.0
                    m_TargetAzimuth = 155.0
                    Status(StatusType.staAction, "Slewing to Alt/Az: " & FormatDec(m_TargetAltitude) & " " & FormatDec(m_TargetAzimuth))
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAzAsync method, Altitude: " & FormatDec(m_TargetAltitude) & ", Azimuth: " & FormatDec(m_TargetAzimuth))
                    telescopeDevice.SlewToAltAzAsync(m_TargetAzimuth, m_TargetAltitude)
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                    LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 3: " & telescopeDevice.Tracking)
                    WaitForSlew(p_Name)
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                    LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 4: " & telescopeDevice.Tracking)
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "Conform:SlewTest: Unknown test type " & p_Test.ToString)
            End Select
            If TestStop() Then Exit Sub

            If p_CanDoIt Then 'Should be able to do this so report what happened
                Select Case p_Test
                    Case SlewSyncType.SlewToCoordinates, SlewSyncType.SlewToCoordinatesAsync, SlewSyncType.SlewToTarget, SlewSyncType.SlewToTargetAsync
                        Status(StatusType.staAction, "Slew completed")
                        'Test how close the slew was to the required coordinates
                        CheckScopePosition(p_Name, "Slewed", m_TargetRightAscension, m_TargetDeclination)

                        ' Check that the slews and syncs set the target coordinates correctly per the ASCOM Telescope specification
                        Try
                            actualRA = telescopeDevice.TargetRightAscension
                            LogMsg(p_Name, MessageLevel.msgDebug, String.Format("Current TargetRightAscension: {0}, Set TargetRightAscension: {1}", actualRA, m_TargetRightAscension))
                            Dim raDifference As Double
                            raDifference = RaDifferenceInSeconds(actualRA, m_TargetRightAscension)
                            Select Case raDifference
                                Case Is <= SLEW_SYNC_OK_TOLERANCE  ' Within specified tolerance
                                    LogMsg(p_Name, MessageLevel.msgOK, String.Format("The TargetRightAscension property {0} matches the expected RA OK. ", FormatRA(m_TargetRightAscension)))
                                Case Else ' Outside specified tolerance
                                    LogMsg(p_Name, MessageLevel.msgError, String.Format("The TargetRightAscension property {0} does not match the expected RA {1}", FormatRA(actualRA), FormatRA(m_TargetRightAscension)))
                            End Select

                        Catch ex As COMException When (ex.ErrorCode = ErrorCodes.ValueNotSet) Or (ex.ErrorCode = g_ExNotSet1) Or (ex.ErrorCode = g_ExNotSet2)
                            LogMsg(p_Name, MessageLevel.msgError, "The Driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet COM exception was thrown instead.")
                        Catch ex As ASCOM.InvalidOperationException
                            LogMsg(p_Name, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, An InvalidOperationException was thrown instead.")
                        Catch ex As DriverException When (ex.Number = ErrorCodes.ValueNotSet) Or (ex.Number = g_ExNotSet1) Or (ex.Number = g_ExNotSet2)
                            LogMsg(p_Name, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.")
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "")
                        End Try

                        Try
                            actualDec = telescopeDevice.TargetDeclination
                            LogMsg(p_Name, MessageLevel.msgDebug, String.Format("Current TargetDeclination: {0}, Set TargetDeclination: {1}", actualDec, m_TargetDeclination))
                            Dim decDifference As Double
                            decDifference = Math.Round(Math.Abs(actualDec - m_TargetDeclination) * 60.0 * 60.0, 1, MidpointRounding.AwayFromZero) ' Dec difference is in arc seconds from degrees of Declination
                            Select Case decDifference
                                Case Is <= SLEW_SYNC_OK_TOLERANCE ' Within specified tolerance
                                    LogMsg(p_Name, MessageLevel.msgOK, String.Format("The TargetDeclination property {0} matches the expected Declination OK. ", FormatDec(m_TargetDeclination)))
                                Case Else ' Outside specified tolerance
                                    LogMsg(p_Name, MessageLevel.msgError, String.Format("The TargetDeclination property {0} does not match the expected Declination {1}", FormatDec(actualDec), FormatDec(m_TargetDeclination)))
                            End Select
                        Catch ex As COMException When (ex.ErrorCode = ErrorCodes.ValueNotSet) Or (ex.ErrorCode = g_ExNotSet1) Or (ex.ErrorCode = g_ExNotSet2)
                            LogMsg(p_Name, MessageLevel.msgError, "The Driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet COM exception was thrown instead.")
                        Catch ex As ASCOM.InvalidOperationException
                            LogMsg(p_Name, MessageLevel.msgError, "The Driver did not set the TargetDeclination property as required by the Telescope specification, An InvalidOperationException was thrown instead.")
                        Catch ex As DriverException When (ex.Number = ErrorCodes.ValueNotSet) Or (ex.Number = g_ExNotSet1) Or (ex.Number = g_ExNotSet2)
                            LogMsg(p_Name, MessageLevel.msgError, "The Driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.")
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "")
                        End Try

                    Case SlewSyncType.SlewToAltAz, SlewSyncType.SlewToAltAzAsync
                        Status(StatusType.staAction, "Slew completed")
                        'Test how close the slew was to the required coordinates
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Azimuth property")
                        l_ActualAzimuth = telescopeDevice.Azimuth
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Altitude property")
                        l_ActualAltitude = telescopeDevice.Altitude

                        l_Difference = Math.Abs(l_ActualAzimuth - m_TargetAzimuth)
                        If l_Difference > 350.0 Then l_Difference = 360.0 - l_Difference ' Deal with the case where the two elements are on different sides of 360 degrees
                        Select Case l_Difference
                            Case Is <= (1.0 / 3600.0) ' seconds
                                LogMsg(p_Name, MessageLevel.msgOK, "Slewed to target Azimuth OK: " & FormatAzimuth(m_TargetAzimuth))
                                Exit Select
                            Case Is <= (2.0 / 3600.0) '2 seconds
                                LogMsg(p_Name, MessageLevel.msgOK, "Slewed to within 2 seconds of Azimuth target: " & FormatAzimuth(m_TargetAzimuth) & " Actual Azimuth " & FormatAzimuth(l_ActualAzimuth))
                                Exit Select
                            Case Else
                                LogMsg(p_Name, MessageLevel.msgInfo, "Slewed to within " & FormatAzimuth(l_Difference) & " DD:MM:SS of expected Azimuth: " & FormatAzimuth(m_TargetAzimuth))
                        End Select

                        l_Difference = Math.Abs(l_ActualAltitude - m_TargetAltitude)
                        Select Case l_Difference
                            Case Is <= (1.0 / 3600.0) '<1 seconds
                                LogMsg(p_Name, MessageLevel.msgOK, "Slewed to target Altitude OK: " & FormatAltitude(m_TargetAltitude))
                                Exit Select
                            Case Is <= (2.0 / 3600.0) '2 seconds
                                LogMsg(p_Name, MessageLevel.msgOK, "Slewed to within 2 seconds of Altitude target: " & FormatAltitude(m_TargetAltitude) & " Actual Altitude " & FormatAltitude(l_ActualAltitude))
                                Exit Select
                            Case Else
                                LogMsg(p_Name, MessageLevel.msgInfo, "Slewed to within " & FormatAltitude(l_Difference) & " DD:MM:SS of expected Altitude: " & FormatAltitude(m_TargetAltitude))
                        End Select
                    Case Else 'Do nothing
                End Select
            Else 'Not supposed to be able to do this but no error generated so report an error
                LogMsg(p_Name, MessageLevel.msgIssue, p_CanDoItName & " is false but no exception was generated on use")
            End If
        Catch ex As Exception
            Select Case p_Test
                Case SlewSyncType.SlewToCoordinates, SlewSyncType.SlewToCoordinatesAsync, SlewSyncType.SlewToTarget, SlewSyncType.SlewToTargetAsync, SlewSyncType.SyncToCoordinates, SlewSyncType.SyncToTarget
                    If trackingState Then ' The mount is tracking so test for normal operation
                        If p_CanDoIt Then ' Mount says that it does support this capability
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, p_CanDoItName & " is True")
                        Else ' Mount says that it does not support this capability
                            HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, p_CanDoItName & " is False")
                        End If
                    Else ' The mount is not tracking so expect an InValidOperationException or MethodNotImplementedException
                        If p_CanDoIt Then 'Mount says that it does support this capability
                            HandleInvalidOperationExceptionAsOK(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "An ASCOM.InvalidOperationException should be thrown when Tracking is false", $"{p_Name} correctly rejected the operation because Tracking is not enabled")
                        Else ' Mount says that it does not support this capability
                            HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, p_CanDoItName & " is False")
                        End If
                    End If
                Case SlewSyncType.SlewToAltAz, SlewSyncType.SlewToAltAzAsync
                    If Not trackingState Then ' The mount is tracking so test for normal operation
                        If p_CanDoIt Then ' Mount says that it does support this capability
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, p_CanDoItName & " is True")
                        Else ' Mount says that it does not support this capability
                            HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, p_CanDoItName & " is False")
                        End If
                    Else ' The mount is not tracking so expect an InValidOperationException or MethodNotImplementedException
                        If p_CanDoIt Then 'Mount says that it does support this capability
                            HandleInvalidOperationExceptionAsOK(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "An ASCOM.InvalidOperationException should be thrown when Tracking is false", $"{p_Name} correctly rejected the operation because Tracking is enabled")
                        Else ' Mount says that it does not support this capability
                            HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, p_CanDoItName & " is False")
                        End If
                    End If
                Case Else
                    ' No other cases
            End Select

        End Try
        g_Status.Clear()  'Clear status messages
    End Sub

    ''' <summary>
    ''' Confirm that InValidValueExceptions are thrown for invalid values
    ''' </summary>
    ''' <param name="p_Test">The method to test</param>
    ''' <param name="BadCoordinate1">RA or Altitude</param>
    ''' <param name="BadCoordinate2">Dec or Azimuth</param>
    ''' <remarks></remarks>
    Private Sub TelescopeBadCoordinateTest(p_Name As String, p_Test As SlewSyncType, BadCoordinate1 As Double, BadCoordinate2 As Double)

        Dim trackingState As Boolean

        Status(StatusType.staTest, p_Name)

        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, $"About to get Tracking property")
        trackingState = telescopeDevice.Tracking

        'Enable tracking where appropriate
        Select Case p_Test
            Case SlewSyncType.SlewToCoordinates, SlewSyncType.SlewToCoordinatesAsync, SlewSyncType.SlewToTarget, SlewSyncType.SlewToTargetAsync, SlewSyncType.SyncToCoordinates, SlewSyncType.SyncToTarget
                If canSetTracking And trackingCanActuallybeChanged Then
                    If canSetTracking And (Not trackingState) Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true")
                        telescopeDevice.Tracking = True
                        trackingState = True
                    End If
                Else
                    LogMsgDebug(p_Name, $"Not setting tracking - CanSetTracking:  {canSetTracking}, TrackingCanActuallyBeChanged: {trackingCanActuallybeChanged}")
                End If
            Case SlewSyncType.SlewToAltAz, SlewSyncType.SlewToAltAzAsync, SlewSyncType.SyncToAltAz
                If canSetTracking And trackingCanActuallybeChanged Then
                    If canSetTracking And (Not trackingState) Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to false")
                        telescopeDevice.Tracking = False
                        trackingState = False
                    End If
                Else
                    LogMsgDebug(p_Name, $"Not setting tracking - CanSetTracking:  {canSetTracking}, TrackingCanActuallyBeChanged: {trackingCanActuallybeChanged}")
                End If

            Case Else
                ' Does not need to be enabled for remaining tests so no action
        End Select

        Select Case p_Test
            Case SlewSyncType.SlewToCoordinates, SlewSyncType.SlewToCoordinatesAsync
                Try
                    Status(StatusType.staAction, "Slew underway")
                    m_TargetRightAscension = BadCoordinate1
                    m_TargetDeclination = 0.0
                    If p_Test = SlewSyncType.SlewToCoordinates Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinates method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(m_TargetDeclination))
                        telescopeDevice.SlewToCoordinates(m_TargetRightAscension, m_TargetDeclination)
                    Else
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(m_TargetDeclination))
                        telescopeDevice.SlewToCoordinatesAsync(m_TargetRightAscension, m_TargetDeclination)
                    End If
                    Status(StatusType.staAction, "Attempting to abort slew")
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                        telescopeDevice.AbortSlew()
                    Catch : End Try ' Attempt to stop any motion that has actually started
                    LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                Catch ex As Exception
                    Status(StatusType.staAction, "Slew rejected")
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad RA coordinate", "Correctly rejected bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                End Try
                Try
                    Status(StatusType.staAction, "Slew underway")
                    m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0)
                    m_TargetDeclination = BadCoordinate2
                    If p_Test = SlewSyncType.SlewToCoordinates Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinates method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(m_TargetDeclination))
                        telescopeDevice.SlewToCoordinates(m_TargetRightAscension, m_TargetDeclination)
                    Else
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(m_TargetDeclination))
                        telescopeDevice.SlewToCoordinatesAsync(m_TargetRightAscension, m_TargetDeclination)
                    End If
                    Status(StatusType.staAction, "Attempting to abort slew")
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                        telescopeDevice.AbortSlew()
                    Catch : End Try ' Attempt to stop any motion that has actually started
                    LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                Catch ex As Exception
                    Status(StatusType.staAction, "Slew rejected")
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                End Try

            Case SlewSyncType.SyncToCoordinates
                Try
                    Status(StatusType.staAction, "Sync underway")
                    m_TargetRightAscension = BadCoordinate1
                    m_TargetDeclination = 0.0
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToCoordinates method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(m_TargetDeclination))
                    telescopeDevice.SyncToCoordinates(m_TargetRightAscension, m_TargetDeclination)
                    LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                Catch ex As Exception
                    Status(StatusType.staAction, "Sync rejected")
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad RA coordinate", "Correctly rejected bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                End Try
                Try
                    Status(StatusType.staAction, "Sync underway")
                    m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0)
                    m_TargetDeclination = BadCoordinate2
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToCoordinates method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(m_TargetDeclination))
                    telescopeDevice.SyncToCoordinates(m_TargetRightAscension, m_TargetDeclination)
                    LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                Catch ex As Exception
                    Status(StatusType.staAction, "Sync rejected")
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                End Try

            Case SlewSyncType.SlewToTarget, SlewSyncType.SlewToTargetAsync
                Try
                    Status(StatusType.staAction, "Slew underway")
                    m_TargetRightAscension = BadCoordinate1
                    m_TargetDeclination = 0.0
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " & FormatRA(m_TargetRightAscension))
                    telescopeDevice.TargetRightAscension = m_TargetRightAscension
                    ' Successfully set bad RA coordinate so now set the good Dec coordinate and see whether the move fails when the slew is attempted
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " & FormatDec(m_TargetDeclination))
                        telescopeDevice.TargetDeclination = m_TargetDeclination
                    Catch : End Try
                    Try
                        If p_Test = SlewSyncType.SlewToTarget Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTarget method")
                            telescopeDevice.SlewToTarget()
                        Else
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTargetAsync method")
                            telescopeDevice.SlewToTargetAsync()
                        End If
                        Status(StatusType.staAction, "Attempting to abort slew")
                        Try
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                            telescopeDevice.AbortSlew()
                        Catch : End Try ' Attempt to stop any motion that has actually started
                        LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                    Catch ex As Exception ' Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                        Status(StatusType.staAction, "Slew rejected")
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad RA coordinate", "Correctly rejected bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                    End Try
                Catch ex As Exception
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad RA coordinate", "Telescope.TargetRA correctly rejected bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                End Try
                Try
                    Status(StatusType.staAction, "Slew underway")
                    m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0)
                    m_TargetDeclination = BadCoordinate2
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " & FormatDec(m_TargetDeclination))
                    telescopeDevice.TargetDeclination = m_TargetDeclination
                    ' Successfully set bad Dec coordinate so now set the good RA coordinate and see whether the move fails when the slew is attempted
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " & FormatRA(m_TargetRightAscension))
                        telescopeDevice.TargetRightAscension = m_TargetRightAscension
                    Catch : End Try
                    Try
                        If p_Test = SlewSyncType.SlewToTarget Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTarget method")
                            telescopeDevice.SlewToTarget()
                        Else
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTargetAsync method")
                            telescopeDevice.SlewToTargetAsync()
                        End If
                        Status(StatusType.staAction, "Attempting to abort slew")
                        Try
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                            telescopeDevice.AbortSlew()
                        Catch : End Try ' Attempt to stop any motion that has actually started
                        LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                    Catch ex As Exception ' Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                        Status(StatusType.staAction, "Slew rejected")
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                    End Try
                Catch ex As Exception
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad Dec coordinate", "Telescope.TargetDeclination correctly rejected bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                End Try

            Case SlewSyncType.SyncToTarget
                Try
                    Status(StatusType.staAction, "Sync underway")
                    m_TargetRightAscension = BadCoordinate1
                    m_TargetDeclination = 0.0
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " & FormatRA(m_TargetRightAscension))
                    telescopeDevice.TargetRightAscension = m_TargetRightAscension
                    ' Successfully set bad RA coordinate so now set the good Dec coordinate and see whether the move fails when the slew is attempted
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " & FormatDec(m_TargetDeclination))
                        telescopeDevice.TargetDeclination = m_TargetDeclination
                    Catch : End Try
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToTarget method")
                        telescopeDevice.SyncToTarget()
                        LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                    Catch ex As Exception ' Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                        Status(StatusType.staAction, "Sync rejected")
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad RA coordinate", "Correctly rejected bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                    End Try
                Catch ex As Exception
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad RA coordinate", "Telescope.TargetRA correctly rejected bad RA coordinate: " & FormatRA(m_TargetRightAscension))
                End Try
                Try
                    Status(StatusType.staAction, "Sync underway")
                    m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0)
                    m_TargetDeclination = BadCoordinate2
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " & FormatDec(m_TargetDeclination))
                    telescopeDevice.TargetDeclination = m_TargetDeclination
                    ' Successfully set bad Dec coordinate so now set the good RA coordinate and see whether the move fails when the slew is attempted
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " & FormatRA(m_TargetRightAscension))
                        telescopeDevice.TargetRightAscension = m_TargetRightAscension
                    Catch : End Try
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToTarget method")
                        telescopeDevice.SyncToTarget()
                        LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                    Catch ex As Exception ' Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                        Status(StatusType.staAction, "Sync rejected")
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                    End Try
                Catch ex As Exception
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad Dec coordinate", "Telescope.TargetDeclination correctly rejected bad Dec coordinate: " & FormatDec(m_TargetDeclination))
                End Try

            Case SlewSyncType.SlewToAltAz, SlewSyncType.SlewToAltAzAsync
                Try
                    Status(StatusType.staAction, "Slew underway")
                    m_TargetAltitude = BadCoordinate1
                    m_TargetAzimuth = 45.0
                    If p_Test = SlewSyncType.SlewToAltAz Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAz method, Altitude: " & FormatDec(m_TargetAltitude) & ", Azimuth: " & FormatDec(m_TargetAzimuth))
                        telescopeDevice.SlewToAltAz(m_TargetAzimuth, m_TargetAltitude)
                    Else
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About To call SlewToAltAzAsync method, Altitude:  " & FormatDec(m_TargetAltitude) & ", Azimuth: " & FormatDec(m_TargetAzimuth))
                        telescopeDevice.SlewToAltAzAsync(m_TargetAzimuth, m_TargetAltitude)
                    End If
                    Status(StatusType.staAction, "Attempting to abort slew")
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                        telescopeDevice.AbortSlew()
                    Catch : End Try ' Attempt to stop any motion that has actually started
                    LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Altitude coordinate: " & FormatAltitude(m_TargetAltitude))
                Catch ex As Exception
                    Status(StatusType.staAction, "Slew rejected")
                    If IsInvalidValueException(p_Name, ex) Then
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Altitude coordinate", "Correctly rejected bad Altitude coordinate: " & FormatAltitude(m_TargetAltitude))
                    Else ' Handle valid InvalidoperationException or some other exception type
                        HandleInvalidOperationExceptionAsOK(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "An ASCOM.InvalidOperationException should be thrown when Tracking is false", $"{p_Name} correctly rejected the operation because Tracking is enabled")
                    End If
                End Try
                Try
                    Status(StatusType.staAction, "Slew underway")
                    m_TargetAltitude = 45.0
                    m_TargetAzimuth = BadCoordinate2
                    If p_Test = SlewSyncType.SlewToAltAz Then
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAz method, Altitude: " & FormatDec(m_TargetAltitude) & ", Azimuth: " & FormatDec(m_TargetAzimuth))
                        telescopeDevice.SlewToAltAz(m_TargetAzimuth, m_TargetAltitude)
                    Else
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAzAsync method, Altitude: " & FormatDec(m_TargetAltitude) & ", Azimuth: " & FormatDec(m_TargetAzimuth))
                        telescopeDevice.SlewToAltAzAsync(m_TargetAzimuth, m_TargetAltitude)
                    End If
                    Status(StatusType.staAction, "Attempting to abort slew")
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                        telescopeDevice.AbortSlew()
                    Catch : End Try ' Attempt to stop any motion that has actually started
                    LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Azimuth coordinate: " & FormatAzimuth(m_TargetAzimuth))
                Catch ex As Exception
                    Status(StatusType.staAction, "Slew rejected")
                    If IsInvalidValueException(p_Name, ex) Then
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Altitude coordinate", "Correctly rejected bad Altitude coordinate: " & FormatAltitude(m_TargetAltitude))
                    Else ' Handle valid InvalidoperationException or some other exception type
                        HandleInvalidOperationExceptionAsOK(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "An ASCOM.InvalidOperationException should be thrown when Tracking is false", $"{p_Name} correctly rejected the operation because Tracking is enabled")
                    End If
                End Try

            Case SlewSyncType.SyncToAltAz
                Try
                    Status(StatusType.staAction, "Sync underway")
                    m_TargetAltitude = BadCoordinate1
                    m_TargetAzimuth = 45.0
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToAltAz method, Altitude: " & FormatDec(m_TargetAltitude) & ", Azimuth: " & FormatDec(m_TargetAzimuth))
                    telescopeDevice.SyncToAltAz(m_TargetAzimuth, m_TargetAltitude)
                    LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Altitude coordinate: " & FormatAltitude(m_TargetAltitude))
                Catch ex As Exception
                    Status(StatusType.staAction, "Sync rejected")
                    If IsInvalidValueException(p_Name, ex) Then
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Altitude coordinate", "Correctly rejected bad Altitude coordinate: " & FormatAltitude(m_TargetAltitude))
                    Else ' Handle valid InvalidoperationException or some other exception type
                        HandleInvalidOperationExceptionAsOK(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "An ASCOM.InvalidOperationException should be thrown when Tracking is false", $"{p_Name} correctly rejected the operation because Tracking is enabled")
                    End If
                End Try
                Try
                    Status(StatusType.staAction, "Sync underway")
                    m_TargetAltitude = 45.0
                    m_TargetAzimuth = BadCoordinate2
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToAltAz method, Altitude: " & FormatDec(m_TargetAltitude) & ", Azimuth: " & FormatDec(m_TargetAzimuth))
                    telescopeDevice.SyncToAltAz(m_TargetAzimuth, m_TargetAltitude)
                    LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Azimuth coordinate: " & FormatAzimuth(m_TargetAzimuth))
                Catch ex As Exception
                    Status(StatusType.staAction, "Sync rejected")
                    If IsInvalidValueException(p_Name, ex) Then
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Altitude coordinate", "Correctly rejected bad Altitude coordinate: " & FormatAltitude(m_TargetAltitude))
                    Else ' Handle valid InvalidoperationException or some other exception type
                        HandleInvalidOperationExceptionAsOK(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "An ASCOM.InvalidOperationException should be thrown when Tracking is false", $"{p_Name} correctly rejected the operation because Tracking is enabled")
                    End If
                End Try

            Case Else
                LogMsg(p_Name, MessageLevel.msgError, "Conform:SlewTest: Unknown test type " & p_Test.ToString)
        End Select
        If TestStop() Then Exit Sub
    End Sub

    Private Sub TelescopePerformanceTest(ByVal p_Type As PerformanceType, ByVal p_Name As String)
        Dim l_StartTime As Date, l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate As Double
        Status(StatusType.staAction, p_Name)
        Try
            l_StartTime = Now
            l_Count = 0.0
            l_LastElapsedTime = 0.0
            Do
                l_Count += 1.0
                Select Case p_Type
                    Case PerformanceType.tstPerfAltitude
                        m_Altitude = telescopeDevice.Altitude
                    Case PerformanceType.tstPerfAtHome
                        m_AtHome = telescopeDevice.AtHome
                    Case PerformanceType.tstPerfAtPark
                        m_AtPark = telescopeDevice.AtPark
                    Case PerformanceType.tstPerfAzimuth
                        m_Azimuth = telescopeDevice.Azimuth
                    Case PerformanceType.tstPerfDeclination
                        m_Declination = telescopeDevice.Declination
                    Case PerformanceType.tstPerfIsPulseGuiding
                        m_IsPulseGuiding = telescopeDevice.IsPulseGuiding
                    Case PerformanceType.tstPerfRightAscension
                        m_RightAscension = telescopeDevice.RightAscension
                    Case PerformanceType.tstPerfSideOfPier
                        m_SideOfPier = telescopeDevice.SideOfPier
                    Case PerformanceType.tstPerfSiderealTime
                        m_SiderealTimeScope = telescopeDevice.SiderealTime
                    Case PerformanceType.tstPerfSlewing
                        m_Slewing = telescopeDevice.Slewing
                    Case PerformanceType.tstPerfUTCDate
                        m_UTCDate = telescopeDevice.UTCDate
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "Conform:PerformanceTest: Unknown test type " & p_Type.ToString)
                End Select
                l_ElapsedTime = Now.Subtract(l_StartTime).TotalSeconds
                If l_ElapsedTime > l_LastElapsedTime + 1.0 Then
                    Status(StatusType.staStatus, l_Count & " transactions in " & Format(l_ElapsedTime, "0") & " seconds")
                    l_LastElapsedTime = l_ElapsedTime
                    Application.DoEvents()
                    If TestStop() Then Exit Sub
                End If
            Loop Until l_ElapsedTime > PERF_LOOP_TIME
            l_Rate = l_Count / l_ElapsedTime
            Select Case l_Rate
                Case Is > 10.0
                    LogMsg("Performance: " & p_Name, MessageLevel.msgInfo, "Transaction rate: " & Format(l_Rate, "0.0") & " per second")
                    Exit Select
                Case 2.0 To 10.0
                    LogMsg("Performance: " & p_Name, MessageLevel.msgInfo, "Transaction rate: " & Format(l_Rate, "0.0") & " per second")
                    Exit Select
                Case 1.0 To 2.0
                    LogMsg("Performance: " & p_Name, MessageLevel.msgInfo, "Transaction rate: " & Format(l_Rate, "0.0") & " per second")
                    Exit Select
                Case Else
                    LogMsg("Performance: " & p_Name, MessageLevel.msgInfo, "Transaction rate: " & Format(l_Rate, "0.0") & " per second")
            End Select
        Catch ex As Exception
            LogMsg("Performance: " & p_Name, MessageLevel.msgError, EX_NET & ex.Message)
        End Try
    End Sub

    Private Sub TelescopeParkedExceptionTest(ByVal p_Type As ParkedExceptionType, ByVal p_Name As String)
        Dim l_TargetRA As Double
        g_Status.Action = p_Name
        g_Status.Test = p_Type.ToString
        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to get AtPark property")
        If telescopeDevice.AtPark Then ' We are still parked so test AbortSlew
            Try
                Select Case p_Type
                    Case ParkedExceptionType.tstPExcepAbortSlew
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                        telescopeDevice.AbortSlew()
                    Case ParkedExceptionType.tstPExcepFindHome
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call FindHome method")
                        telescopeDevice.FindHome()
                    Case ParkedExceptionType.tstPExcepMoveAxisPrimary
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call MoveAxis(Primary, 0.0) method")
                        telescopeDevice.MoveAxis(TelescopeAxes.axisPrimary, 0.0)
                    Case ParkedExceptionType.tstPExcepMoveAxisSecondary
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call MoveAxis(Secondary, 0.0) method")
                        telescopeDevice.MoveAxis(TelescopeAxes.axisSecondary, 0.0)
                    Case ParkedExceptionType.tstPExcepMoveAxisTertiary
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call MoveAxis(Tertiary, 0.0) method")
                        telescopeDevice.MoveAxis(TelescopeAxes.axisTertiary, 0.0)
                    Case ParkedExceptionType.tstPExcepPulseGuide
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call PulseGuide(East, 0.0) method")
                        telescopeDevice.PulseGuide(GuideDirections.guideEast, 0.0)
                    Case ParkedExceptionType.tstPExcepSlewToCoordinates
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call SlewToCoordinates method")
                        telescopeDevice.SlewToCoordinates(TelescopeRAFromSiderealTime("Parked:" & p_Name, 1.0), 0.0)
                    Case ParkedExceptionType.tstPExcepSlewToCoordinatesAsync
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method")
                        telescopeDevice.SlewToCoordinatesAsync(TelescopeRAFromSiderealTime("Parked:" & p_Name, 1.0), 0.0)
                        WaitForSlew("Parked:" & p_Name)
                    Case ParkedExceptionType.tstPExcepSlewToTarget
                        l_TargetRA = TelescopeRAFromSiderealTime("Parked:" & p_Name, 1.0)
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to set property TargetRightAscension to " & FormatRA(l_TargetRA))
                        telescopeDevice.TargetRightAscension = l_TargetRA
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to set property TargetDeclination to 0.0")
                        telescopeDevice.TargetDeclination = 0.0
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call SlewToTarget method")
                        telescopeDevice.SlewToTarget()
                    Case ParkedExceptionType.tstPExcepSlewToTargetAsync
                        l_TargetRA = TelescopeRAFromSiderealTime("Parked:" & p_Name, 1.0)
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to set property to " & FormatRA(l_TargetRA))
                        telescopeDevice.TargetRightAscension = l_TargetRA
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to set property to 0.0")
                        telescopeDevice.TargetDeclination = 0.0
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call method")
                        telescopeDevice.SlewToTargetAsync()
                        WaitForSlew("Parked:" & p_Name)
                    Case ParkedExceptionType.tstPExcepSyncToCoordinates
                        l_TargetRA = TelescopeRAFromSiderealTime("Parked:" & p_Name, 1.0)
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call method, RA: " & FormatRA(l_TargetRA) & ", Declination: 0.0")
                        telescopeDevice.SyncToCoordinates(l_TargetRA, 0.0)
                    Case ParkedExceptionType.tstPExcepSyncToTarget
                        l_TargetRA = TelescopeRAFromSiderealTime("Parked:" & p_Name, 1.0)
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to set property to " & FormatRA(l_TargetRA))
                        telescopeDevice.TargetRightAscension = l_TargetRA
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to set property to 0.0")
                        telescopeDevice.TargetDeclination = 0.0
                        If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to call SyncToTarget method")
                        telescopeDevice.SyncToTarget()
                    Case Else
                        LogMsg("Parked:" & p_Name, MessageLevel.msgError, "Conform:ParkedExceptionTest: Unknown test type " & p_Type.ToString)
                End Select
                LogMsg("Parked:" & p_Name, MessageLevel.msgIssue, p_Name & " didn't raise an error when Parked as required")
            Catch ex As Exception
                LogMsg("Parked:" & p_Name, MessageLevel.msgOK, p_Name & " did raise an exception when Parked as required")
            End Try
            'Check that Telescope is still parked after issuing the command!
            If g_Settings.DisplayMethodCalls Then LogMsg("Parked:" & p_Name, MessageLevel.msgComment, "About to get AtPark property")
            If Not telescopeDevice.AtPark Then LogMsg("Parked:" & p_Name, MessageLevel.msgIssue, "Telescope was unparked by the " & p_Name & " command. This should not happen!")

        Else
            LogMsg("Parked:" & p_Name, MessageLevel.msgIssue, "Not parked after Telescope.Park command, " & p_Name & " when parked test skipped")
        End If
        g_Status.Clear()
    End Sub

    Private Sub TelescopeAxisRateTest(ByVal p_Name As String, ByVal p_Axis As TelescopeAxes)
        Dim l_NAxisRates, l_i, l_j As Integer
        Dim l_AxisRateOverlap, l_AxisRateDuplicate, l_CanGetAxisRates, l_HasRates As Boolean
        Dim l_Count As Integer = 0

#If DEBUG Then
        Dim l_AxisRates As IAxisRates = Nothing
        Dim l_Rate As IRate = Nothing
#Else
        Dim l_AxisRatesIRates As IAxisRates
        Dim l_AxisRates As Object = Nothing
        Dim l_Rate As Object = Nothing
#End If

        Try
            l_NAxisRates = 0
            l_AxisRates = Nothing
            Select Case p_Axis
                Case TelescopeAxes.axisPrimary
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AxisRates method, Axis: " & TelescopeAxes.axisPrimary)
                    l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisPrimary) 'Get primary axis rates
                    m_AxisRatePrimaryOverlap = False
                Case TelescopeAxes.axisSecondary
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AxisRates method, Axis: " & TelescopeAxes.axisSecondary)
                    l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisSecondary) 'Get secondary axis rates
                    m_AxisRateSecondaryOverlap = False
                Case TelescopeAxes.axisTertiary
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AxisRates method, Axis: " & TelescopeAxes.axisTertiary)
                    l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisTertiary) 'Get tertiary axis rates
                    m_AxisRateTertiaryOverlap = False
                Case Else
                    LogMsg("TelescopeAxisRateTest", MessageLevel.msgError, "Unknown telescope axis: " & p_Axis.ToString)
            End Select

            Try

                If l_AxisRates Is Nothing Then
                    LogMsg(p_Name, MessageLevel.msgDebug, "ERROR: The driver did NOT return an AxisRates object!")
                Else
                    LogMsg(p_Name, MessageLevel.msgDebug, "OK - the driver returned an AxisRates object")
                End If

                l_Count = l_AxisRates.Count 'Save count for use later if no members are returned in the for each loop test
                LogMsg(p_Name & " Count", MessageLevel.msgDebug, "The driver returned " & l_Count & " rates")

                Dim i As Integer
                For i = 1 To l_Count
#If DEBUG Then
                    Dim AxisRateItem As IRate
#Else
                    Dim AxisRateItem As Object
#End If
                    AxisRateItem = l_AxisRates.Item(i)
                    LogMsg(p_Name & " Count", MessageLevel.msgDebug, "Rate " & i & " - Minimum: " & AxisRateItem.Minimum.ToString & ", Maximum: " & AxisRateItem.Maximum.ToString)
                Next
            Catch ex As COMException
                LogMsg(p_Name & " Count", MessageLevel.msgError, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
            Catch ex As Exception
                LogMsg(p_Name & " Count", MessageLevel.msgError, EX_NET & ex.ToString)
            End Try

            Try
                Dim l_Enum As System.Collections.IEnumerator, l_Obj As Object
#If DEBUG Then
                Dim AxisRateItem As IRate = Nothing
#Else
                Dim AxisRateItem As Object = Nothing
#End If

                l_Enum = l_AxisRates.GetEnumerator()
                If l_Enum Is Nothing Then
                    LogMsg(p_Name & " Enum", MessageLevel.msgDebug, "ERROR: The driver did NOT return an Enumerator object!")
                Else
                    LogMsg(p_Name & " Enum", MessageLevel.msgDebug, "OK - the driver returned an Enumerator object")
                End If

                l_Enum.Reset()
                LogMsg(p_Name & " Enum", MessageLevel.msgDebug, "Reset Enumerator")

                Do While l_Enum.MoveNext()
                    LogMsg(p_Name & " Enum", MessageLevel.msgDebug, "Reading Current")
                    l_Obj = l_Enum.Current
                    LogMsg(p_Name & " Enum", MessageLevel.msgDebug, "Read Current OK, Type: " & l_Obj.GetType.Name)
#If DEBUG Then
                    AxisRateItem = CType(l_Obj, IRate)
#Else
                    AxisRateItem = l_Obj
#End If
                    LogMsg(p_Name & " Enum", MessageLevel.msgDebug, "Found axis rate - Minimum: " & AxisRateItem.Minimum.ToString & ", Maximum: " & AxisRateItem.Maximum.ToString)
                Loop

                l_Enum.Reset()
                l_Enum = Nothing
                AxisRateItem = Nothing
            Catch ex As COMException
                LogMsg(p_Name & " Enum", MessageLevel.msgError, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
            Catch ex As Exception
                LogMsg(p_Name & " Enum", MessageLevel.msgError, EX_NET & ex.ToString)
            End Try

            If l_AxisRates.Count > 0 Then
                Try
#If DEBUG Then
                    For Each l_Rate In l_AxisRates
                        If (l_Rate.Minimum < 0) Or (l_Rate.Maximum < 0) Then ' Error because negative values are not allowed
                            LogMsg(p_Name, MessageLevel.msgError, "Minimum or maximum rate is negative: " & l_Rate.Minimum.ToString & ", " & l_Rate.Maximum.ToString)
                        Else ' All positive values so continue tests
                            If l_Rate.Minimum <= l_Rate.Maximum Then 'Minimum <= Maximum so OK
                                LogMsg(p_Name, MessageLevel.msgOK, "Axis rate minimum: " & l_Rate.Minimum.ToString & " Axis rate maximum: " & l_Rate.Maximum.ToString)
                            Else 'Minimum > Maximum so error!
                                LogMsg(p_Name, MessageLevel.msgError, "Maximum rate is less than minimum rate - minimum: " & l_Rate.Minimum.ToString & " maximum: " & l_Rate.Maximum.ToString)
                            End If
                        End If

                        'Save rates for overlap testing
                        l_NAxisRates += 1
                        m_AxisRatesArray(l_NAxisRates, AXIS_RATE_MINIMUM) = l_Rate.Minimum
                        m_AxisRatesArray(l_NAxisRates, AXIS_RATE_MAXIMUM) = l_Rate.Maximum

                        ' Set a flag to indicate that there are available rate objects in this axis that can be accessed through ForEach and then tested for correct Rate.Dispose() action
                        l_HasRates = True
                    Next
#Else
                    If g_Settings.UseDriverAccess Then
                        l_AxisRatesIRates = CType(l_AxisRates, IAxisRates)
                        For Each l_Rate In l_AxisRatesIRates
                            If (l_Rate.Minimum < 0) Or (l_Rate.Maximum < 0) Then ' Error because negative values are not allowed
                                LogMsg(p_Name, MessageLevel.msgError, "Minimum or maximum rate is negative: " & l_Rate.Minimum.ToString & ", " & l_Rate.Maximum.ToString)
                            Else ' All positive values so continue tests
                                If l_Rate.Minimum <= l_Rate.Maximum Then 'Minimum <= Maximum so OK
                                    LogMsg(p_Name, MessageLevel.msgOK, "Axis rate minimum: " & l_Rate.Minimum.ToString & " Axis rate maximum: " & l_Rate.Maximum.ToString)
                                Else 'Minimum > Maximum so error!
                                    LogMsg(p_Name, MessageLevel.msgError, "Maximum rate is less than minimum rate - minimum: " & l_Rate.Minimum.ToString & " maximum: " & l_Rate.Maximum.ToString)
                                End If
                            End If

                            'Save rates for overlap testing
                            l_NAxisRates += 1
                            m_AxisRatesArray(l_NAxisRates, AXIS_RATE_MINIMUM) = l_Rate.Minimum
                            m_AxisRatesArray(l_NAxisRates, AXIS_RATE_MAXIMUM) = l_Rate.Maximum
                            l_HasRates = True
                        Next
                    Else
                        For Each l_Rate In l_AxisRates
                            If (l_Rate.Minimum < 0) Or (l_Rate.Maximum < 0) Then ' Error because negative values are not allowed
                                LogMsg(p_Name, MessageLevel.msgError, "Minimum or maximum rate is negative: " & l_Rate.Minimum.ToString & ", " & l_Rate.Maximum.ToString)
                            Else ' All positive values so continue tests
                                If l_Rate.Minimum <= l_Rate.Maximum Then 'Minimum <= Maximum so OK
                                    LogMsg(p_Name, MessageLevel.msgOK, "Axis rate minimum: " & l_Rate.Minimum.ToString & " Axis rate maximum: " & l_Rate.Maximum.ToString)
                                Else 'Minimum > Maximum so error!
                                    LogMsg(p_Name, MessageLevel.msgError, "Maximum rate is less than minimum rate - minimum: " & l_Rate.Minimum.ToString & " maximum: " & l_Rate.Maximum.ToString)
                                End If
                                l_HasRates = True
                            End If

                            'Save rates for overlap testing
                            l_NAxisRates += 1
                            m_AxisRatesArray(l_NAxisRates, AXIS_RATE_MINIMUM) = l_Rate.Minimum
                            m_AxisRatesArray(l_NAxisRates, AXIS_RATE_MAXIMUM) = l_Rate.Maximum
                        Next
                    End If
#End If
                Catch ex As COMException
                    LogMsg(p_Name, MessageLevel.msgError, "COM Unable to read AxisRates object - Exception: " & ex.Message & " " & Hex(ex.ErrorCode))
                    LogMsg(p_Name, MessageLevel.msgDebug, "COM Unable to read AxisRates object - Exception: " & ex.ToString)
                Catch ex As DriverException
                    LogMsg(p_Name, MessageLevel.msgError, ".NET Unable to read AxisRates object - Exception: " & ex.Message & " " & Hex(ex.Number))
                    LogMsg(p_Name, MessageLevel.msgDebug, ".NET Unable to read AxisRates object - Exception: " & ex.ToString)
                Catch ex As Exception
                    LogMsg(p_Name, MessageLevel.msgError, "Unable to read AxisRates object - Exception: " & ex.Message)
                    LogMsg(p_Name, MessageLevel.msgDebug, "Unable to read AxisRates object - Exception: " & ex.ToString)
                End Try

                'Overlap testing
                If l_NAxisRates > 1 Then 'Confirm whether there are overlaps if number of axis rate pairs exceeds 1
                    For l_i = 1 To l_NAxisRates
                        For l_j = 1 To l_NAxisRates
                            If l_i <> l_j Then 'Only test different lines, shouldn't compare same lines!
                                If (m_AxisRatesArray(l_i, AXIS_RATE_MINIMUM) >= m_AxisRatesArray(l_j, AXIS_RATE_MINIMUM)) And
                                   (m_AxisRatesArray(l_i, AXIS_RATE_MINIMUM) <= m_AxisRatesArray(l_j, AXIS_RATE_MAXIMUM)) Then l_AxisRateOverlap = True
                            End If
                        Next
                    Next
                End If
                If l_AxisRateOverlap Then
                    LogMsg(p_Name, MessageLevel.msgIssue, "Overlapping axis rates found, suggest these be rationalised to remove overlaps")
                Else
                    LogMsg(p_Name, MessageLevel.msgOK, "No overlapping axis rates found")
                End If

                'Duplicate testing
                l_AxisRateDuplicate = False
                If l_NAxisRates > 1 Then 'Confirm whether there are overlaps if number of axis rate pairs exceeds 1
                    For l_i = 1 To l_NAxisRates
                        For l_j = 1 To l_NAxisRates
                            If l_i <> l_j Then 'Only test different lines, shouldn't compare same lines!
                                If (m_AxisRatesArray(l_i, AXIS_RATE_MINIMUM) = m_AxisRatesArray(l_j, AXIS_RATE_MINIMUM)) And
                                   (m_AxisRatesArray(l_i, AXIS_RATE_MAXIMUM) = m_AxisRatesArray(l_j, AXIS_RATE_MAXIMUM)) Then l_AxisRateDuplicate = True
                            End If
                        Next
                    Next
                End If
                If l_AxisRateDuplicate Then
                    LogMsg(p_Name, MessageLevel.msgIssue, "Duplicate axis rates found, suggest these be removed")
                Else
                    LogMsg(p_Name, MessageLevel.msgOK, "No duplicate axis rates found")
                End If

            Else
                LogMsg(p_Name, MessageLevel.msgOK, "Empty axis rate returned")
            End If

            l_CanGetAxisRates = True ' Record that this driver can deliver a viable AxisRates object that can be tested for AxisRates.Dispose() later

        Catch ex As COMException
            LogMsg(p_Name, MessageLevel.msgError, "COM Unable to get an AxisRates object - Exception: " & ex.Message & " " & Hex(ex.ErrorCode))
        Catch ex As DriverException
            LogMsg(p_Name, MessageLevel.msgError, ".NET Unable to get an AxisRates object - Exception: " & ex.Message & " " & Hex(ex.Number))
        Catch ex As NullReferenceException 'Report null objects returned by the driver that are caught by DriverAccess.
            LogMsg(p_Name, MessageLevel.msgError, ex.Message)
            LogMsg(p_Name, MessageLevel.msgDebug, ex.ToString()) 'If debug then give full information
        Catch ex As Exception
            LogMsg(p_Name, MessageLevel.msgError, "Unable to get or unable to use an AxisRates object - Exception: " & ex.ToString)
        End Try

        ' Clean up AxisRate object if used
        If Not (l_AxisRates Is Nothing) Then
            Try : l_AxisRates.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_AxisRates) : Catch : End Try
            l_AxisRates = Nothing
        End If

        'Clean up and release rate object if used
        If Not (l_Rate Is Nothing) Then
            Try : l_Rate.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_Rate) : Catch : End Try
            l_Rate = Nothing
        End If

        If l_CanGetAxisRates Then ' The driver does return a viable AxisRates object that can be tested for correct AxisRates.Dispose() and Rate.Dispose() operation
            Try
                ' Test Rate.Dispose()
                Select Case p_Axis ' Get the relevant axis rates object for this axis
                    Case TelescopeAxes.axisPrimary
                        l_AxisRates = DriverAsObject.AxisRates(TelescopeAxes.axisPrimary)
                    Case TelescopeAxes.axisSecondary
                        l_AxisRates = DriverAsObject.AxisRates(TelescopeAxes.axisSecondary)
                    Case TelescopeAxes.axisTertiary
                        l_AxisRates = DriverAsObject.AxisRates(TelescopeAxes.axisTertiary)
                    Case Else
                        LogMsgError(p_Name, "AxisRate.Dispose() - Unknown axis: " & p_Axis.ToString())
                End Select

                If l_HasRates Then ' This axis does have one or more rates that can be accessed through ForEach so test these for correct Rate.Dispose() action
                    For Each l_Rate In l_AxisRates
                        Try
                            l_Rate.Dispose()
                            LogMsgOK(p_Name, String.Format("Successfully disposed of rate {0} - {1}", l_Rate.Minimum, l_Rate.Maximum))
                        Catch ex1 As MissingMemberException
                            LogMsg(p_Name, MessageLevel.msgOK, String.Format("Rate.Dispose() member not present for rate {0} - {1}", l_Rate.Minimum, l_Rate.Maximum))
                        Catch ex1 As Exception
                            LogMsgWarning(p_Name, String.Format("Rate.Dispose() for rate {0} - {1} threw an exception but it is poor practice to throw exceptions in Dispose methods: {2}", l_Rate.Minimum, l_Rate.Maximum, ex1.Message()))
                            LogMsg("TrackingRates.Dispose", MessageLevel.msgDebug, "Exception: " & ex1.ToString())
                        End Try
                    Next
                End If

                ' Test AxisRates.Dispose()
                Try
                    LogMsg(p_Name, MessageLevel.msgDebug, "Disposing axis rates")
                    l_AxisRates.Dispose()
                    LogMsg(p_Name, MessageLevel.msgOK, "Disposed axis rates OK")
                Catch ex1 As MissingMemberException
                    LogMsg(p_Name, MessageLevel.msgOK, "AxisRates.Dispose() member not present for axis " & p_Axis.ToString())
                Catch ex1 As Exception
                    LogMsgWarning(p_Name, "AxisRates.Dispose() threw an exception but it is poor practice to throw exceptions in Dispose() methods: " & ex1.Message())
                    LogMsg("AxisRates.Dispose", MessageLevel.msgDebug, "Exception: " & ex1.ToString())
                End Try

            Catch ex As Exception
                LogMsg(p_Name, MessageLevel.msgError, "AxisRate.Dispose() - Unable to get or unable to use an AxisRates object - Exception: " & ex.ToString)
            End Try
        Else
            LogMsgInfo(p_Name, "AxisRates.Dispose() testing skipped because of earlier issues in obtaining a viable AxisRates object.")
        End If

        g_Status.Clear()  'Clear status messages

    End Sub

    Private Sub TelescopeRequiredMethodsTest(ByVal p_Type As RequiredMethodType, ByVal p_Name As String)
        Try
            g_Status.Test = p_Name
            Select Case p_Type
                Case RequiredMethodType.tstAxisrates
                    'This is now done by TelescopeAxisRateTest subroutine 
                Case RequiredMethodType.tstCanMoveAxisPrimary
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call CanMoveAxis method " & TelescopeAxes.axisPrimary)
                    m_CanMoveAxisPrimary = telescopeDevice.CanMoveAxis(TelescopeAxes.axisPrimary)
                    LogMsg(p_Name, MessageLevel.msgOK, p_Name & " " & m_CanMoveAxisPrimary.ToString)
                Case RequiredMethodType.tstCanMoveAxisSecondary
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call CanMoveAxis method " & TelescopeAxes.axisSecondary)
                    m_CanMoveAxisSecondary = telescopeDevice.CanMoveAxis(TelescopeAxes.axisSecondary)
                    LogMsg(p_Name, MessageLevel.msgOK, p_Name & " " & m_CanMoveAxisSecondary.ToString)
                Case RequiredMethodType.tstCanMoveAxisTertiary
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call CanMoveAxis method " & TelescopeAxes.axisTertiary)
                    m_CanMoveAxisTertiary = telescopeDevice.CanMoveAxis(TelescopeAxes.axisTertiary)
                    LogMsg(p_Name, MessageLevel.msgOK, p_Name & " " & m_CanMoveAxisTertiary.ToString)
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "Conform:RequiredMethodsTest: Unknown test type " & p_Type.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Method, Required.Mandatory, ex, "")
        End Try

        'Clean up and release each object after use
        'If Not (m_Rate Is Nothing) Then Try : Marshal.ReleaseComObject(m_Rate) : Catch : End Try
        'm_Rate = Nothing
        g_Status.Clear() 'Clear status messages
    End Sub

    Private Sub TelescopeOptionalMethodsTest(ByVal p_Type As OptionalMethodType, ByVal p_Name As String, ByVal p_CanTest As Boolean)
        Dim l_ct As Integer, l_TestDec, l_TestRAOffset As Double
#If DEBUG Then
        Dim l_AxisRates As IAxisRates = Nothing
#Else
        Dim l_AxisRates As Object = Nothing
#End If
        Status(StatusType.staTest, p_Name)

        LogMsg("TelescopeOptionalMethodsTest", MessageLevel.msgDebug, p_Type.ToString & " " & p_Name & " " & p_CanTest.ToString)
        If p_CanTest Then ' Confirm that an error is raised if the optional command is not implemented
            Try
                ' Set the test declination value depending on whether the scope is in the northern or southern hemisphere
                If m_SiteLatitude > 0.0 Then
                    l_TestDec = 45.0 ' Positive for the northern hemisphere
                Else
                    l_TestDec = -45.0 ' Negative for the southern hemisphere
                End If

                l_TestRAOffset = 3.0 ' Set the test RA offset as 3 hours from local sider5eal time
                LogMsg(p_Name, MessageLevel.msgDebug, String.Format("Test RA offset: {0}, Test declination: {1}", l_TestRAOffset, l_TestDec))

                Select Case p_Type
                    Case OptionalMethodType.AbortSlew
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                        telescopeDevice.AbortSlew()
                        LogMsg("AbortSlew", MessageLevel.msgOK, "AbortSlew OK when not slewing")

                    Case OptionalMethodType.DestinationSideOfPier
                        ' Get the DestinationSideOfPier for a target in the West i.e. for a German mount when the tube is on the East side of the pier
                        m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -l_TestRAOffset)
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call DestinationSideOfPier method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(l_TestDec))
                        m_DestinationSideOfPierEast = telescopeDevice.DestinationSideOfPier(m_TargetRightAscension, l_TestDec)
                        LogMsg(p_Name, MessageLevel.msgDebug, "German mount - scope on the pier's East side, target in the West : " & FormatRA(m_TargetRightAscension) & " " & FormatDec(l_TestDec) & " " & m_DestinationSideOfPierEast.ToString)

                        ' Get the DestinationSideOfPier for a target in the East i.e. for a German mount when the tube is on the West side of the pier
                        m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, l_TestRAOffset)
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call DestinationSideOfPier method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(l_TestDec))
                        m_DestinationSideOfPierWest = telescopeDevice.DestinationSideOfPier(m_TargetRightAscension, l_TestDec)
                        LogMsg(p_Name, MessageLevel.msgDebug, "German mount - scope on the pier's West side, target in the East: " & FormatRA(m_TargetRightAscension) & " " & FormatDec(l_TestDec) & " " & m_DestinationSideOfPierWest.ToString)

                        ' Make sure that we received two valid values i.e. that neither side returned PierSide.Unknown and that the two valid returned values are not the same i.e. we got one PierSide.PierEast and one PierSide.PierWest
                        If (m_DestinationSideOfPierEast = PierSide.pierUnknown) Or (m_DestinationSideOfPierWest = PierSide.pierUnknown) Then
                            LogMsg(p_Name, MessageLevel.msgError, "Invalid SideOfPier value received, Target in West: " & m_DestinationSideOfPierEast.ToString & ", Target in East: " & m_DestinationSideOfPierWest.ToString)
                        Else
                            If m_DestinationSideOfPierEast = m_DestinationSideOfPierWest Then
                                LogMsg(p_Name, MessageLevel.msgIssue, "Same value for DestinationSideOfPier received on both sides of the meridian: " & m_DestinationSideOfPierEast)
                            Else
                                LogMsg(p_Name, MessageLevel.msgOK, "DestinationSideOfPier is different on either side of the meridian")
                            End If
                        End If

                    Case OptionalMethodType.FindHome
                        If g_InterfaceVersion > 1 Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call FindHome method")
                            telescopeDevice.FindHome()
                            m_StartTime = Now
                            Status(StatusType.staAction, "Waiting for mount to home")
                            l_ct = 0
                            Do
                                WaitFor(SLEEP_TIME)
                                l_ct += 1
                                g_Status.Status = l_ct.ToString
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get AtHome property")
                            Loop Until telescopeDevice.AtHome Or TestStop() Or (Now.Subtract(m_StartTime).TotalMilliseconds > 60000) ' Wait up to a minute to find home
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get AtHome property")
                            If telescopeDevice.AtHome Then
                                LogMsg(p_Name, MessageLevel.msgOK, "Found home OK.")
                            Else
                                LogMsg(p_Name, MessageLevel.msgInfo, "Failed to Find home within 1 minute")
                            End If
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get AtPark property")
                            If telescopeDevice.AtPark Then
                                LogMsg(p_Name, MessageLevel.msgIssue, "FindHome has parked the scope as well as finding home")
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call UnPark method")
                                telescopeDevice.Unpark() 'Unpark it ready for further tests
                            End If
                        Else
                            Status(StatusType.staAction, "Waiting for mount to home")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call FindHome method")
                            telescopeDevice.FindHome()
                            g_Status.Clear()
                            LogMsg(p_Name, MessageLevel.msgOK, "Found home OK.")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call Unpark method")
                            telescopeDevice.Unpark() ' Make sure we are still  unparked!
                        End If

                    Case OptionalMethodType.MoveAxisPrimary
                        ' Get axis rates for primary axis
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AxisRates method for axis " & TelescopeAxes.axisPrimary)
                        l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisPrimary)
                        TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisPrimary, l_AxisRates)

                    Case OptionalMethodType.MoveAxisSecondary
                        ' Get axis rates for secondary axis
                        l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisSecondary)
                        TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisSecondary, l_AxisRates)

                    Case OptionalMethodType.MoveAxisTertiary
                        ' Get axis rates for tertiary axis
                        l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisTertiary)
                        TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisTertiary, l_AxisRates)

                    Case OptionalMethodType.PulseGuide
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property")
                        If telescopeDevice.IsPulseGuiding Then ' IsPulseGuiding is true before we've started so this is an error and voids a real test
                            LogMsg(p_Name, MessageLevel.msgError, "IsPulseGuiding is True when not pulse guiding - PulseGuide test omitted")
                        Else 'OK to test pulse guiding
                            Status(StatusType.staAction, "Start PulseGuide")
                            m_StartTime = Now
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call PulseGuide method, Direction: " & GuideDirections.guideEast & ", Duration: " & PULSEGUIDE_MOVEMENT_TIME * 1000 & "ms")
                            telescopeDevice.PulseGuide(GuideDirections.guideEast, PULSEGUIDE_MOVEMENT_TIME * 1000) 'Start a 2 second pulse
                            m_EndTime = Now
                            LogMsg(p_Name, MessageLevel.msgDebug, "PulseGuide command time: " & PULSEGUIDE_MOVEMENT_TIME * 1000 & " milliseconds, PulseGuide call duration: " & m_EndTime.Subtract(m_StartTime).TotalMilliseconds & " milliseconds")

                            If m_EndTime.Subtract(m_StartTime).TotalMilliseconds < PULSEGUIDE_MOVEMENT_TIME * 0.75 * 1000 Then 'If less than three quarters of the expected duration then assume we have returned early
                                l_ct = 0
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property")
                                If telescopeDevice.IsPulseGuiding Then
                                    Do
                                        WaitFor(SLEEP_TIME)
                                        l_ct += 1
                                        g_Status.Status = l_ct.ToString
                                        If TestStop() Then Exit Sub
                                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property")
                                    Loop Until (Not telescopeDevice.IsPulseGuiding) Or (DateTime.Now.Subtract(m_StartTime).TotalMilliseconds > PULSEGUIDE_TIMEOUT_TIME * 1000) 'Wait for success or timeout

                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property")
                                    If Not telescopeDevice.IsPulseGuiding Then
                                        LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous pulse guide found OK")
                                        LogMsg(p_Name, MessageLevel.msgDebug, "IsPulseGuiding = True duration: " & DateTime.Now.Subtract(m_StartTime).TotalMilliseconds & " milliseconds")
                                    Else
                                        LogMsg(p_Name, MessageLevel.msgIssue, "Asynchronous pulse guide expected but IsPulseGuiding is still TRUE " & PULSEGUIDE_TIMEOUT_TIME & " seconds beyond expected time")
                                    End If
                                Else
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Asynchronous pulse guide expected but IsPulseGuiding has returned FALSE")
                                End If
                            Else 'Assume synchronous pulse guide and that IsPulseGuiding is false
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property")
                                If Not telescopeDevice.IsPulseGuiding Then
                                    LogMsg(p_Name, MessageLevel.msgOK, "Synchronous pulse guide found OK")
                                Else
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Synchronous pulse guide expected but IsPulseGuiding has returned TRUE")
                                End If
                            End If
                        End If

                    Case OptionalMethodType.SideOfPierWrite
                        'SideOfPier Write

                        ' Tracking must be enabled for the SideOfPier tests so enable it if possible, otherwise skip the test
                        If g_Settings.DisplayMethodCalls Then LogMsg("SideOfPier Write", MessageLevel.msgComment, $"About to get Tracking property")
                        If Not telescopeDevice.Tracking Then
                            If trackingCanActuallybeChanged Then
                                If g_Settings.DisplayMethodCalls Then LogMsg("SideOfPier Write", MessageLevel.msgComment, $"About to set Tracking property True")
                                telescopeDevice.Tracking = True
                            Else
                                LogMsgInfo("SideOfPier Write", "Tests skipped because Tracking cannot be enabled.")
                                Exit Select
                            End If
                        End If

                        If canSetPierside Then 'Can set pier side so test if we can
                            SlewScope(TelescopeRAFromHourAngle(p_Name, -3.0), 0.0, "Slewing to far start point")
                            If TestStop() Then Exit Sub

                            SlewScope(TelescopeRAFromHourAngle(p_Name, -0.03), 0.0, "Slewing to near start point") '2 minutes from zenith
                            If TestStop() Then Exit Sub

                            ' We are now 2 minutes from the meridian looking east so allow the mount to track for 7 minutes 
                            ' so it passes through the meridian and ends up 5 minutes past the meridian
                            LogMsg(p_Name, MessageLevel.msgInfo, "This test will now wait for 7 minutes while the mount tracks through the Meridian")

                            'Wait for mount to move
                            m_StartTime = Now
                            Do
                                System.Threading.Thread.Sleep(SLEEP_TIME)
                                Application.DoEvents()
                                SetStatus(p_Name, "Waiting for transit through Meridian", Convert.ToInt32(Now.Subtract(m_StartTime).TotalSeconds) & "/" & SIDEOFPIER_MERIDIAN_TRACKING_PERIOD / 1000 & " seconds")
                            Loop Until (Now.Subtract(m_StartTime).TotalMilliseconds > SIDEOFPIER_MERIDIAN_TRACKING_PERIOD) Or TestStop()

                            'SlewScope(TelescopeRAFromHourAngle(+0.0833333), 0.0, "Slewing to flip point") '5 minutes past zenith
                            If TestStop() Then Exit Sub

                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get SideOfPier property")
                            Select Case telescopeDevice.SideOfPier
                                Case PierSide.pierEast 'We are on pierEast so try pierWest
                                    Try
                                        LogMsg(p_Name, MessageLevel.msgDebug, "Scope is pierEast so flipping West")
                                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set SideOfPier property to " & PierSide.pierWest)
                                        telescopeDevice.SideOfPier = PierSide.pierWest
                                        WaitForSlew(p_Name)
                                        If TestStop() Then Exit Sub
                                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get SideOfPier property")
                                        m_SideOfPier = telescopeDevice.SideOfPier
                                        If m_SideOfPier = PierSide.pierWest Then
                                            LogMsg(p_Name, MessageLevel.msgOK, "Successfully flipped pierEast to pierWest")
                                        Else
                                            LogMsg(p_Name, MessageLevel.msgIssue, "Failed to set SideOfPier to pierWest, got: " & m_SideOfPier.ToString)
                                        End If
                                    Catch ex As Exception
                                        HandleException("SideOfPier Write pierWest", MemberType.Method, Required.MustBeImplemented, ex, "CanSetPierSide is True")
                                    End Try
                                Case PierSide.pierWest 'We are on pierWest so try pierEast
                                    Try
                                        LogMsg(p_Name, MessageLevel.msgDebug, "Scope is pierWest so flipping East")
                                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set SideOfPier property to " & PierSide.pierEast)
                                        telescopeDevice.SideOfPier = PierSide.pierEast
                                        WaitForSlew(p_Name)
                                        If TestStop() Then Exit Sub
                                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get SideOfPier property")
                                        m_SideOfPier = telescopeDevice.SideOfPier
                                        If m_SideOfPier = PierSide.pierEast Then
                                            LogMsg(p_Name, MessageLevel.msgOK, "Successfully flipped pierWest to pierEast")
                                        Else
                                            LogMsg(p_Name, MessageLevel.msgIssue, "Failed to set SideOfPier to pierEast, got: " & m_SideOfPier.ToString)
                                        End If
                                    Catch ex As Exception
                                        HandleException("SideOfPier Write pierEast", MemberType.Method, Required.MustBeImplemented, ex, "CanSetPierSide is True")
                                    End Try
                                Case Else 'Unknown pier side
                                    LogMsg(p_Name, MessageLevel.msgError, "Unknown PierSide: " & m_SideOfPier.ToString)
                            End Select
                        Else 'Can't set pier side so it should generate an error
                            Try
                                LogMsg(p_Name, MessageLevel.msgDebug, "Attempting to set SideOfPier")
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set SideOfPier property to " & PierSide.pierEast)
                                telescopeDevice.SideOfPier = PierSide.pierEast
                                LogMsg(p_Name, MessageLevel.msgDebug, "SideOfPier set OK to pierEast but should have thrown an error")
                                WaitForSlew(p_Name)
                                LogMsg(p_Name, MessageLevel.msgIssue, "CanSetPierSide is false but no exception was generated when set was attempted")
                            Catch ex As Exception
                                HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "CanSetPierSide is False")
                            Finally
                                WaitForSlew(p_Name) 'Make sure slewing is stopped if an exception was thrown
                            End Try
                        End If

                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to false")
                        telescopeDevice.Tracking = False
                        If TestStop() Then Exit Sub

                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "Conform:OptionalMethodsTest: Unknown test type " & p_Type.ToString)
                End Select

                ' Clean up AxisRate object, if used
                If Not (l_AxisRates Is Nothing) Then
                    Try
#If DEBUG Then
                        LogMsg(p_Name, MessageLevel.msgComment, "About to dispose of AxisRates object")
                        l_AxisRates.Dispose()
#Else
                        If g_Settings.UseDriverAccess Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to dispose of AxisRates object")
                            l_AxisRates.Dispose()
                        End If
#End If
                        LogMsg(p_Name, MessageLevel.msgOK, "AxisRates object successfully disposed")
                    Catch ex As Exception
                        LogMsgError(p_Name, "AxisRates.Dispose threw an exception but must not: " & ex.Message)
                        LogMsg(p_Name, MessageLevel.msgDebug, "Exception: " & ex.ToString())
                    End Try
                    Try : Marshal.ReleaseComObject(l_AxisRates) : Catch : End Try
                    l_AxisRates = Nothing
                End If

                If TestStop() Then Exit Sub
            Catch ex As Exception
                HandleException(p_Name, MemberType.Method, Required.Optional, ex, "")
            End Try
        Else 'Can property is false so confirm that an error is generated
            Try
                Select Case p_Type
                    Case OptionalMethodType.AbortSlew
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method")
                        telescopeDevice.AbortSlew()
                    Case OptionalMethodType.DestinationSideOfPier
                        m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -1.0)
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call DestinationSideOfPier method, RA: " & FormatRA(m_TargetRightAscension) & ", Declination: " & FormatDec(0.0))
                        m_DestinationSideOfPier = telescopeDevice.DestinationSideOfPier(m_TargetRightAscension, 0.0)
                    Case OptionalMethodType.FindHome
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call FindHome method")
                        telescopeDevice.FindHome()
                    Case OptionalMethodType.MoveAxisPrimary
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & TelescopeAxes.axisPrimary & " at speed 0")
                        telescopeDevice.MoveAxis(TelescopeAxes.axisPrimary, 0.0)
                    Case OptionalMethodType.MoveAxisSecondary
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & TelescopeAxes.axisSecondary & " at speed 0")
                        telescopeDevice.MoveAxis(TelescopeAxes.axisSecondary, 0.0)
                    Case OptionalMethodType.MoveAxisTertiary
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & TelescopeAxes.axisTertiary & " at speed 0")
                        telescopeDevice.MoveAxis(TelescopeAxes.axisTertiary, 0.0)
                    Case OptionalMethodType.PulseGuide
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call PulseGuide method, Direction: " & GuideDirections.guideEast & ", Duration: 0ms")
                        telescopeDevice.PulseGuide(GuideDirections.guideEast, 0)
                    Case OptionalMethodType.SideOfPierWrite
                        If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set SideOfPier property to " & PierSide.pierEast)
                        telescopeDevice.SideOfPier = PierSide.pierEast
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "Conform:OptionalMethodsTest: Unknown test type " & p_Type.ToString)
                End Select
                LogMsg(p_Name, MessageLevel.msgIssue, "Can" & p_Name & " is false but no exception was generated on use")
            Catch ex As Exception
                If IsInvalidValueException(p_Name, ex) Then
                    LogMsg(p_Name, MessageLevel.msgOK, "Received an invalid value exception")
                Else
                    If p_Type = OptionalMethodType.SideOfPierWrite Then ' PierSide is actually a property even though I have it in the methods section!!
                        HandleException(p_Name, MemberType.Property, Required.MustNotBeImplemented, ex, "Can" & p_Name & " is False")
                    Else
                        HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "Can" & p_Name & " is False")
                    End If
                End If
            End Try
        End If
        g_Status.Clear()  'Clear status messages
    End Sub

    Private Sub TelescopeCanTest(ByVal p_Type As CanType, ByVal p_Name As String)
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, String.Format("About to get {0} property", p_Type.ToString()))
            Select Case p_Type
                Case CanType.CanFindHome
                    canFindHome = telescopeDevice.CanFindHome
                    LogMsg(p_Name, MessageLevel.msgOK, canFindHome.ToString)
                Case CanType.CanPark
                    canPark = telescopeDevice.CanPark
                    LogMsg(p_Name, MessageLevel.msgOK, canPark.ToString)
                Case CanType.CanPulseGuide
                    canPulseGuide = telescopeDevice.CanPulseGuide
                    LogMsg(p_Name, MessageLevel.msgOK, canPulseGuide.ToString)
                Case CanType.CanSetDeclinationRate
                    If g_InterfaceVersion > 1 Then
                        canSetDeclinationRate = telescopeDevice.CanSetDeclinationRate
                        LogMsg(p_Name, MessageLevel.msgOK, canSetDeclinationRate.ToString)
                    Else
                        LogMsg("CanSetDeclinationRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
                    End If
                Case CanType.CanSetGuideRates
                    If g_InterfaceVersion > 1 Then
                        canSetGuideRates = telescopeDevice.CanSetGuideRates
                        LogMsg(p_Name, MessageLevel.msgOK, canSetGuideRates.ToString)
                    Else
                        LogMsg("CanSetGuideRates", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
                    End If
                Case CanType.CanSetPark
                    canSetPark = telescopeDevice.CanSetPark
                    LogMsg(p_Name, MessageLevel.msgOK, canSetPark.ToString)
                Case CanType.CanSetPierSide
                    If g_InterfaceVersion > 1 Then
                        canSetPierside = telescopeDevice.CanSetPierSide
                        LogMsg(p_Name, MessageLevel.msgOK, canSetPierside.ToString)
                    Else
                        LogMsg("CanSetPierSide", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
                    End If
                Case CanType.CanSetRightAscensionRate
                    If g_InterfaceVersion > 1 Then
                        canSetRightAscensionRate = telescopeDevice.CanSetRightAscensionRate
                        LogMsg(p_Name, MessageLevel.msgOK, canSetRightAscensionRate.ToString)
                    Else
                        LogMsg("CanSetRightAscensionRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
                    End If
                Case CanType.CanSetTracking
                    canSetTracking = telescopeDevice.CanSetTracking
                    LogMsg(p_Name, MessageLevel.msgOK, canSetTracking.ToString)
                Case CanType.CanSlew
                    canSlew = telescopeDevice.CanSlew
                    LogMsg(p_Name, MessageLevel.msgOK, canSlew.ToString)
                Case CanType.CanSlewAltAz
                    If g_InterfaceVersion > 1 Then
                        canSlewAltAz = telescopeDevice.CanSlewAltAz
                        LogMsg(p_Name, MessageLevel.msgOK, canSlewAltAz.ToString)
                    Else
                        LogMsg("CanSlewAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
                    End If
                Case CanType.CanSlewAltAzAsync
                    If g_InterfaceVersion > 1 Then
                        canSlewAltAzAsync = telescopeDevice.CanSlewAltAzAsync
                        LogMsg(p_Name, MessageLevel.msgOK, canSlewAltAzAsync.ToString)
                    Else
                        LogMsg("CanSlewAltAzAsync", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
                    End If
                Case CanType.CanSlewAsync
                    canSlewAsync = telescopeDevice.CanSlewAsync
                    LogMsg(p_Name, MessageLevel.msgOK, canSlewAsync.ToString)
                Case CanType.CanSync
                    canSync = telescopeDevice.CanSync
                    LogMsg(p_Name, MessageLevel.msgOK, canSync.ToString)
                Case CanType.CanSyncAltAz
                    If g_InterfaceVersion > 1 Then
                        canSyncAltAz = telescopeDevice.CanSyncAltAz
                        LogMsg(p_Name, MessageLevel.msgOK, canSyncAltAz.ToString)
                    Else
                        LogMsg("CanSyncAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
                    End If
                Case CanType.CanUnPark
                    canUnpark = telescopeDevice.CanUnpark
                    LogMsg(p_Name, MessageLevel.msgOK, canUnpark.ToString)
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "Conform:CanTest: Unknown test type " & p_Type.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "")
        End Try
    End Sub

#If RELEASE Then

    Private Sub TelescopeMoveAxisTest(ByVal p_Name As String, ByVal p_Axis As TelescopeAxes, ByVal p_AxisRates As Object)
        Dim l_Rate As Object = Nothing
#Else
    Private Sub TelescopeMoveAxisTest(ByVal p_Name As String, ByVal p_Axis As TelescopeAxes, ByVal p_AxisRates As IAxisRates)
        Dim l_Rate As IRate = Nothing
#End If
        Dim l_MoveRate, l_RateMinimum, l_RateMaximum As Double, l_TrackingStart, l_TrackingEnd, l_CanSetZero As Boolean
        Dim l_RateCount As Integer

        'Determine lowest and highest tracking rates
        l_RateMinimum = Double.PositiveInfinity 'Set to invalid values
        l_RateMaximum = Double.NegativeInfinity
        LogMsg(p_Name, MessageLevel.msgDebug, "Number of rates found: " & p_AxisRates.Count)

        If p_AxisRates.Count > 0 Then
#If DEBUG Then
            For Each l_Rate In p_AxisRates
                If l_Rate.Minimum < l_RateMinimum Then l_RateMinimum = l_Rate.Minimum
                If l_Rate.Maximum > l_RateMaximum Then l_RateMaximum = l_Rate.Maximum
                LogMsg(p_Name, MessageLevel.msgDebug, "Checking rates: " & l_Rate.Minimum & " " & l_Rate.Maximum & ", Current rates: " & l_RateMinimum & " " & l_RateMaximum)
                l_RateCount += 1
            Next
#Else
            Dim l_AxisRatesIRates As IAxisRates
            l_RateCount = 0
            If g_Settings.UseDriverAccess Then
                l_AxisRatesIRates = CType(p_AxisRates, IAxisRates)
                For Each l_Rate In l_AxisRatesIRates
                    If l_Rate.Minimum < l_RateMinimum Then l_RateMinimum = l_Rate.Minimum
                    If l_Rate.Maximum > l_RateMaximum Then l_RateMaximum = l_Rate.Maximum
                    LogMsg(p_Name, MessageLevel.msgDebug, "Checking rates: " & l_Rate.Minimum & " " & l_Rate.Maximum & ", Current rates: " & l_RateMinimum & " " & l_RateMaximum)
                    l_RateCount += 1
                Next
            Else
                For Each l_Rate In p_AxisRates
                    If l_Rate.Minimum < l_RateMinimum Then l_RateMinimum = l_Rate.Minimum
                    If l_Rate.Maximum > l_RateMaximum Then l_RateMaximum = l_Rate.Maximum
                    LogMsg(p_Name, MessageLevel.msgDebug, "Checking rates: " & l_Rate.Minimum & " " & l_Rate.Maximum & ", Current rates: " & l_RateMinimum & " " & l_RateMaximum)
                    l_RateCount += 1
                Next
            End If
#End If

            If (l_RateMinimum <> Double.PositiveInfinity) And (l_RateMaximum <> Double.NegativeInfinity) Then 'Found valid rates
                LogMsg(p_Name, MessageLevel.msgDebug, "Found minimum rate: " & l_RateMinimum & " found maximum rate: " & l_RateMaximum)

                'Confirm setting a zero rate works
                Status(StatusType.staAction, "Set zero rate")
                l_CanSetZero = False
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                    telescopeDevice.MoveAxis(p_Axis, 0.0) ' Set a value of zero
                    LogMsg(p_Name, MessageLevel.msgOK, "Can successfully set a movement rate of zero")
                    l_CanSetZero = True
                Catch ex As COMException
                    LogMsg(p_Name, MessageLevel.msgError, "Unable to set a movement rate of zero - " & ex.Message & " " & Hex(ex.ErrorCode))
                Catch ex As DriverException
                    LogMsg(p_Name, MessageLevel.msgError, "Unable to set a movement rate of zero - " & ex.Message & " " & Hex(ex.Number))
                Catch ex As Exception
                    LogMsg(p_Name, MessageLevel.msgError, "Unable to set a movement rate of zero - " & ex.Message)
                End Try

                Status(StatusType.staAction, "Set lower rate")
                'Test that error is generated on attempt to set rate lower than minimum
                Try
                    If l_RateMinimum > 0 Then ' choose a value between the minimum and zero
                        l_MoveRate = l_RateMinimum / 2.0
                    Else ' Choose a large negative value
                        l_MoveRate = -l_RateMaximum - 1.0
                    End If
                    LogMsg(p_Name, MessageLevel.msgDebug, "Using minimum rate: " & l_MoveRate)
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & l_MoveRate)
                    telescopeDevice.MoveAxis(p_Axis, l_MoveRate) ' Set a value lower than the minimum
                    LogMsg(p_Name, MessageLevel.msgIssue, "No exception raised when move axis value < minimum rate: " & l_MoveRate)
                    'Clean up and release each object after use
                    Try : Marshal.ReleaseComObject(l_Rate) : Catch : End Try
                    l_Rate = Nothing
                Catch ex As Exception
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "when move axis is set below lowest rate (" & l_MoveRate & ")", "Exception correctly generated when move axis is set below lowest rate (" & l_MoveRate & ")")
                End Try
                'Clean up and release each object after use
                Try : Marshal.ReleaseComObject(l_Rate) : Catch : End Try
                l_Rate = Nothing
                If TestStop() Then Exit Sub

                'test that error is generated when rate is above maximum set
                Status(StatusType.staAction, "Set upper rate")
                Try
                    l_MoveRate = l_RateMaximum + 1.0
                    LogMsg(p_Name, MessageLevel.msgDebug, "Using maximum rate: " & l_MoveRate)
                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & l_MoveRate)
                    telescopeDevice.MoveAxis(p_Axis, l_MoveRate) ' Set a value higher than the maximum
                    LogMsg(p_Name, MessageLevel.msgIssue, "No exception raised when move axis value > maximum rate: " & l_MoveRate)
                    'Clean up and release each object after use
                    Try : Marshal.ReleaseComObject(l_Rate) : Catch : End Try
                    l_Rate = Nothing
                Catch ex As Exception
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "when move axis is set above highest rate (" & l_MoveRate & ")", "Exception correctly generated when move axis is set above highest rate (" & l_MoveRate & ")")
                End Try
                'Clean up and release each object after use
                Try : Marshal.ReleaseComObject(l_Rate) : Catch : End Try
                l_Rate = Nothing
                If TestStop() Then Exit Sub

                If l_CanSetZero Then ' Can set a rate of zero so undertake these tests
                    'Confirm that lowest tracking rate can be set
                    Status(StatusType.staAction, "Move at minimum rate")
                    If l_RateMinimum <> Double.PositiveInfinity Then ' Valid value found so try and set it
                        Try
                            Status(StatusType.staStatus, "Moving forward")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & l_RateMinimum)
                            telescopeDevice.MoveAxis(p_Axis, l_RateMinimum) ' Set the minimum rate
                            WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                            telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                            Status(StatusType.staStatus, "Moving back")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & -l_RateMinimum)
                            telescopeDevice.MoveAxis(p_Axis, -l_RateMinimum) ' Set the minimum rate
                            WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                            'v1.0.12 Next line added because movement wasn't stopped
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                            telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                            LogMsg(p_Name, MessageLevel.msgOK, "Successfully moved axis at minimum rate: " & l_RateMinimum)
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "when setting rate: " & l_RateMinimum)
                        End Try
                        Status(StatusType.staStatus, "") 'Clear status flag
                    Else ' No valid rate was found so print an error
                        LogMsg(p_Name, MessageLevel.msgError, "Minimum rate test - unable to find lowest axis rate")
                    End If
                    If TestStop() Then Exit Sub

                    'Confirm that highest tracking rate can be set
                    Status(StatusType.staAction, "Move at maximum rate")
                    If l_RateMaximum <> Double.NegativeInfinity Then ' Valid value found so try and set it
                        Try
                            'Confirm not slewing first
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Slewing property")
                            If telescopeDevice.Slewing Then
                                LogMsg(p_Name, MessageLevel.msgIssue, "Slewing was true before start of MoveAxis but should have been false, remaining tests skipped")
                                Exit Sub
                            End If
                            Status(StatusType.staStatus, "Moving forward")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & l_RateMaximum)
                            telescopeDevice.MoveAxis(p_Axis, l_RateMaximum) ' Set the minimum rate
                            'Confirm that slewing is active when the move is underway
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Slewing property")
                            If Not telescopeDevice.Slewing Then LogMsg(p_Name, MessageLevel.msgIssue, "Slewing is not true immediately after axis starts moving in positive direction")
                            WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Slewing property")
                            If Not telescopeDevice.Slewing Then LogMsg(p_Name, MessageLevel.msgIssue, "Slewing is not true after " & MOVE_AXIS_TIME / 1000 & " seconds moving in positive direction")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                            telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                            'Confirm that slewing is false when movement is stopped
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get property")
                            If telescopeDevice.Slewing Then
                                LogMsg(p_Name, MessageLevel.msgIssue, "Slewing incorrectly remains true after stopping positive axis movement, remaining test skipped")
                                Exit Sub
                            End If

                            Status(StatusType.staStatus, "Moving back")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & -l_RateMaximum)
                            telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum) ' Set the minimum rate
                            'Confirm that slewing is active when the move is underway
                            If Not telescopeDevice.Slewing Then LogMsg(p_Name, MessageLevel.msgIssue, "Slewing is not true immediately after axis starts moving in negative direction")
                            WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                            If Not telescopeDevice.Slewing Then LogMsg(p_Name, MessageLevel.msgIssue, "Slewing is not true after " & MOVE_AXIS_TIME / 1000 & " seconds moving in negative direction")
                            'Confirm that slewing is false when movement is stopped
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                            telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                            'Confirm that slewing is false when movement is stopped
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Slewing property")
                            If telescopeDevice.Slewing Then
                                LogMsg(p_Name, MessageLevel.msgIssue, "Slewing incorrectly remains true after stopping negative axis movement, remaining test skipped")
                                Exit Sub
                            End If
                            LogMsg(p_Name, MessageLevel.msgOK, "Successfully moved axis at maximum rate: " & l_RateMaximum)
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "when setting rate: " & l_RateMaximum)
                        End Try
                        Status(StatusType.staStatus, "") 'Clear status flag
                    Else ' No valid rate was found so print an error
                        LogMsg(p_Name, MessageLevel.msgError, "Maximum rate test - unable to find lowest axis rate")
                    End If
                    If TestStop() Then Exit Sub

                    'Confirm that tracking state is correctly restored after a move axis command
                    Try
                        Status(StatusType.staAction, "Tracking state restore")
                        If canSetTracking Then
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                            l_TrackingStart = telescopeDevice.Tracking 'Save the start tracking state
                            Status(StatusType.staStatus, "Moving forward")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & l_RateMaximum)
                            telescopeDevice.MoveAxis(p_Axis, l_RateMaximum) ' Set the maximum rate
                            WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                            Status(StatusType.staStatus, "Stop movement")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                            telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                            l_TrackingEnd = telescopeDevice.Tracking 'Save the final tracking state
                            If l_TrackingStart = l_TrackingEnd Then 'Successfully retained tracking state
                                If l_TrackingStart Then 'Tracking is true so switch to false for return movement
                                    Status(StatusType.staStatus, "Set tracking off")
                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property false")
                                    telescopeDevice.Tracking = False
                                    Status(StatusType.staStatus, "Move back")
                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & -l_RateMaximum)
                                    telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum) ' Set the maximum rate
                                    WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                                    telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                                    Status(StatusType.staStatus, "")
                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property false")
                                    If telescopeDevice.Tracking = False Then 'tracking correctly retained in both states
                                        LogMsg(p_Name, MessageLevel.msgOK, "Tracking state correctly retained for both tracking states")
                                    Else
                                        LogMsg(p_Name, MessageLevel.msgIssue, "Tracking state correctly retained when tracking is " & l_TrackingStart.ToString & ", but not when tracking is false")
                                    End If
                                Else 'Tracking false so switch to true for return movement
                                    Status(StatusType.staStatus, "Set tracking on")
                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property true")
                                    telescopeDevice.Tracking = True
                                    Status(StatusType.staStatus, "Move back")
                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & -l_RateMaximum)
                                    telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum) ' Set the maximum rate
                                    WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                                    telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                                    Status(StatusType.staStatus, "")
                                    If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                                    If telescopeDevice.Tracking = True Then 'tracking correctly retained in both states
                                        LogMsg(p_Name, MessageLevel.msgOK, "Tracking state correctly retained for both tracking states")
                                    Else
                                        LogMsg(p_Name, MessageLevel.msgIssue, "Tracking state correctly retained when tracking is " & l_TrackingStart.ToString & ", but not when tracking is true")
                                    End If
                                End If
                                Status(StatusType.staStatus, "") 'Clear status flag
                            Else 'Tracking state not correctly restored
                                Status(StatusType.staStatus, "Move back")
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & -l_RateMaximum)
                                telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum) ' Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                                Status(StatusType.staStatus, "")
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                                telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property " & l_TrackingStart)
                                telescopeDevice.Tracking = l_TrackingStart 'Restore original value
                                LogMsg(p_Name, MessageLevel.msgIssue, "Tracking state not correctly restored after MoveAxis when CanSetTracking is true")
                            End If
                        Else 'Can't set tracking so just test the current state
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                            l_TrackingStart = telescopeDevice.Tracking
                            Status(StatusType.staStatus, "Moving forward")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed " & l_RateMaximum)
                            telescopeDevice.MoveAxis(p_Axis, l_RateMaximum) ' Set the maximum rate
                            WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                            Status(StatusType.staStatus, "Stop movement")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                            telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property")
                            l_TrackingEnd = telescopeDevice.Tracking 'Save tracking state
                            Status(StatusType.staStatus, "Move back")
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call method MoveAxis for axis " & p_Axis & " at speed " & -l_RateMaximum)
                            telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum) ' Set the maximum rate
                            WaitFor(MOVE_AXIS_TIME) : If TestStop() Then Exit Sub
                            'v1.0.12 next line added because movement wasn't stopped
                            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " & p_Axis & " at speed 0")
                            telescopeDevice.MoveAxis(p_Axis, 0.0) 'Stop the movement on this axis
                            If l_TrackingStart = l_TrackingEnd Then
                                LogMsg(p_Name, MessageLevel.msgOK, "Tracking state correctly restored after MoveAxis when CanSetTracking is false")
                            Else
                                If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to " & l_TrackingStart)
                                telescopeDevice.Tracking = l_TrackingStart 'Restore correct value
                                LogMsg(p_Name, MessageLevel.msgIssue, "Tracking state not correctly restored after MoveAxis when CanSetTracking is false")
                            End If
                            Status(StatusType.staStatus, "") 'Clear status flag
                        End If
                    Catch ex As Exception
                        HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "")
                    End Try
                Else 'Cant set zero so tests skipped 
                    LogMsg(p_Name, MessageLevel.msgInfo, "Remaining MoveAxis tests skipped because unable to set a movement rate of zero")
                End If
                Status(StatusType.staStatus, "") 'Clear status flag
                Status(StatusType.staAction, "") 'Clear action flag
            Else 'Some problem in finding rates inside the AxisRates object
                LogMsg(p_Name, MessageLevel.msgInfo, "Found minimum rate: " & l_RateMinimum & " found maximum rate: " & l_RateMaximum)
                LogMsg(p_Name, MessageLevel.msgError, "Unable to determine lowest or highest rates, expected " & p_AxisRates.Count & " rates, found " & l_RateCount)
            End If
        Else
            LogMsg(p_Name, MessageLevel.msgWarning, "MoveAxis tests skipped because there are no AxisRate values")
        End If
    End Sub

    Private Sub SideOfPierTests()
        Dim l_PierSideMinus3, l_PierSideMinus9, l_PierSidePlus3, l_PierSidePlus9 As SideOfPierResults
        Dim l_Declination3, l_Declination9, l_StartRA As Double

        ' Tracking must be enabled for the SideOfPier tests so enable it if possible, otherwise skip the test
        If g_Settings.DisplayMethodCalls Then LogMsg("SideOfPierTests", MessageLevel.msgComment, $"About to get Tracking property")
        If Not telescopeDevice.Tracking Then
            If trackingCanActuallybeChanged Then
                If g_Settings.DisplayMethodCalls Then LogMsg("SideOfPierTests", MessageLevel.msgComment, $"About to set Tracking property True")
                telescopeDevice.Tracking = True
            Else
                LogMsgInfo("SideOfPier", "Tests skipped because Tracking cannot be enabled.")
                Return
            End If
        End If

        'Slew to starting position
        LogMsg("SideofPier", MessageLevel.msgDebug, "Starting Side of Pier tests")

        Status(StatusType.staTest, "Side of pier tests")
        l_StartRA = TelescopeRAFromHourAngle("SideofPier", -3.0)
        If m_SiteLatitude > 0.0 Then ' We are in the northern hemisphere
            l_Declination3 = 90.0 - ((180.0 - m_SiteLatitude) * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR) ' Calculate for northern hemisphere
            l_Declination9 = 90.0 - m_SiteLatitude * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR
        Else ' We are in the southern hemisphere
            l_Declination3 = -90.0 + ((180.0 + m_SiteLatitude) * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR) ' Calculate for southern hemisphere
            l_Declination9 = -90.0 - m_SiteLatitude * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR
        End If
        LogMsg("SideofPier", MessageLevel.msgDebug, "Declination for hour angle = +-3.0 tests: " & FormatDec(l_Declination3) & ", Declination for hour angle = +-9.0 tests: " & FormatDec(l_Declination9))

        SlewScope(l_StartRA, 0.0, "Move to starting position " & FormatRA(l_StartRA) & " " & FormatDec(0.0))
        If TestStop() Then Exit Sub

        'Run tests
        Status(StatusType.staAction, "Test hour angle -3.0 at declination: " & FormatDec(l_Declination3))
        l_PierSideMinus3 = SOPPierTest(l_StartRA, l_Declination3, "hour angle -3.0") : If TestStop() Then Exit Sub
        Status(StatusType.staAction, "Test hour angle +3.0 at declination: " & FormatDec(l_Declination3))
        l_PierSidePlus3 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +3.0), l_Declination3, "hour angle +3.0") : If TestStop() Then Exit Sub
        Status(StatusType.staAction, "Test hour angle -9.0 at declination: " & FormatDec(l_Declination9))
        l_PierSideMinus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", -9.0), l_Declination9, "hour angle -9.0") : If TestStop() Then Exit Sub
        Status(StatusType.staAction, "Test hour angle +9.0 at declination: " & FormatDec(l_Declination9))
        l_PierSidePlus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +9.0), l_Declination9, "hour angle +9.0") : If TestStop() Then Exit Sub

        If (l_PierSideMinus3.SideOfPier = l_PierSidePlus9.SideOfPier) And (l_PierSidePlus3.SideOfPier = l_PierSideMinus9.SideOfPier) Then 'Reporting physical pier side
            LogMsg("SideofPier", MessageLevel.msgIssue, "SideofPier reports physical pier side rather than pointing state")
        Else ' Make other tests
            If (l_PierSideMinus3.SideOfPier = l_PierSideMinus9.SideOfPier) And (l_PierSidePlus3.SideOfPier = l_PierSidePlus9.SideOfPier) Then
                LogMsg("SideofPier", MessageLevel.msgOK, "Reports the pointing state of the mount as expected")
            Else 'Don't know what this means!
                LogMsg("SideofPier", MessageLevel.msgInfo, "Unknown SideofPier reporting model: HA-3: " & l_PierSideMinus3.SideOfPier.ToString & " HA-9: " & l_PierSideMinus9.SideOfPier.ToString & " HA+3: " & l_PierSidePlus3.SideOfPier.ToString & " HA+9: " & l_PierSidePlus9.SideOfPier.ToString)
            End If
        End If
        LogMsg("SideofPier", MessageLevel.msgInfo, "Reported SideofPier at HA -9, +9: " & TranslatePierSide(l_PierSideMinus9.SideOfPier, False) & TranslatePierSide(l_PierSidePlus9.SideOfPier, False))
        LogMsg("SideofPier", MessageLevel.msgInfo, "Reported SideofPier at HA -3, +3: " & TranslatePierSide(l_PierSideMinus3.SideOfPier, False) & TranslatePierSide(l_PierSidePlus3.SideOfPier, False))

        ' Now test the ASCOM convention that pierWest is returned when the mount is on the west side of the pier facing east at hour angle -3
        If l_PierSideMinus3.SideOfPier = PierSide.pierWest Then
            LogMsg("SideofPier", MessageLevel.msgOK, "pierWest is returned when the mount is observing at an hour angle between -6.0 and 0.0")
        Else
            LogMsg("SideofPier", MessageLevel.msgIssue, "pierEast is returned when the mount is observing at an hour angle between -6.0 and 0.0")
            LogMsg("SideofPier", MessageLevel.msgInfo, "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when observing at hour angles from -6.0 to -0.0 and that pierEast must be returned at hour angles from 0.0 to +6.0.")
        End If
        If l_PierSidePlus3.SideOfPier = PierSide.pierEast Then
            LogMsg("SideofPier", MessageLevel.msgOK, "pierEast is returned when the mount is observing at an hour angle between 0.0 and +6.0")
        Else
            LogMsg("SideofPier", MessageLevel.msgIssue, "pierWest is returned when the mount is observing at an hour angle between 0.0 and +6.0")
            LogMsg("SideofPier", MessageLevel.msgInfo, "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when observing at hour angles from -6.0 to -0.0 and that pierEast must be returned at hour angles from 0.0 to +6.0.")
        End If

        'Test whether DestinationSideOfPier is implemented
        If l_PierSideMinus3.DestinationSideOfPier = PierSide.pierUnknown And
                l_PierSideMinus9.DestinationSideOfPier = PierSide.pierUnknown And
                l_PierSidePlus3.DestinationSideOfPier = PierSide.pierUnknown And
                l_PierSidePlus9.DestinationSideOfPier = PierSide.pierUnknown Then
            LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Analysis skipped as this method is not implemented") 'Not implemented
        Else 'It is implemented so assess the results

            If (l_PierSideMinus3.DestinationSideOfPier = l_PierSidePlus9.DestinationSideOfPier) And (l_PierSidePlus3.DestinationSideOfPier = l_PierSideMinus9.DestinationSideOfPier) Then 'Reporting physical pier side
                LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "DestinationSideofPier reports physical pier side rather than pointing state")
            Else ' Make other tests
                If (l_PierSideMinus3.DestinationSideOfPier = l_PierSideMinus9.DestinationSideOfPier) And (l_PierSidePlus3.DestinationSideOfPier = l_PierSidePlus9.DestinationSideOfPier) Then
                    LogMsg("DestinationSideofPier", MessageLevel.msgOK, "Reports the pointing state of the mount as expected")
                Else 'Don't know what this means!
                    LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Unknown DestinationSideofPier reporting model: HA-3: " & l_PierSideMinus3.SideOfPier.ToString & " HA-9: " & l_PierSideMinus9.SideOfPier.ToString & " HA+3: " & l_PierSidePlus3.SideOfPier.ToString & " HA+9: " & l_PierSidePlus9.SideOfPier.ToString)
                End If
            End If

            ' Now test the ASCOM convention that pierWest is returned when the mount is on the west side of the pier facing east at hour angle -3
            If l_PierSideMinus3.DestinationSideOfPier = PierSide.pierWest Then
                LogMsg("DestinationSideofPier", MessageLevel.msgOK, "pierWest is returned when the mount will observe at an hour angle between -6.0 and 0.0")
            Else
                LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "pierEast is returned when the mount will observe at an hour angle between -6.0 and 0.0")
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when the mount will observe at hour angles from -6.0 to -0.0 and that pierEast must be returned for hour angles from 0.0 to +6.0.")
            End If
            If l_PierSidePlus3.DestinationSideOfPier = PierSide.pierEast Then
                LogMsg("DestinationSideofPier", MessageLevel.msgOK, "pierEast is returned when the mount will observe at an hour angle between 0.0 and +6.0")
            Else
                LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "pierWest is returned when the mount will observe at an hour angle between 0.0 and +6.0")
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when the mount will observe at hour angles from -6.0 to -0.0 and that pierEast must be returned for hour angles from 0.0 to +6.0.")
            End If

        End If
        LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Reported DesintationSideofPier at HA -9, +9: " & TranslatePierSide(l_PierSideMinus9.DestinationSideOfPier, False) & TranslatePierSide(l_PierSidePlus9.DestinationSideOfPier, False))
        LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Reported DesintationSideofPier at HA -3, +3: " & TranslatePierSide(l_PierSideMinus3.DestinationSideOfPier, False) & TranslatePierSide(l_PierSidePlus3.DestinationSideOfPier, False))

        'Clean up
        '3.0.0.12 added conditional test to next line
        If canSetTracking Then telescopeDevice.Tracking = False
        Status(StatusType.staStatus, "")
        Status(StatusType.staAction, "")
        Status(StatusType.staTest, "")
    End Sub

    Function SOPPierTest(ByVal p_RA As Double, ByVal p_DEC As Double, ByVal p_Msg As String) As SideOfPierResults
        'Determine side of pier and destination side of pier results for a particular RA and DEC
        Dim l_Results As SideOfPierResults = New SideOfPierResults ' Create result set object
        Dim l_StartRA, l_StartDEC As Double
        Try
            'Prepare for tests
            l_StartRA = telescopeDevice.RightAscension
            l_StartDEC = telescopeDevice.Declination

            'Do destination side of pier test to see what side of pier we should end up on
            LogMsg("", MessageLevel.msgDebug, "")
            LogMsg("SOPPierTest", MessageLevel.msgDebug, "Testing RA DEC: " & FormatRA(p_RA) & " " & FormatDec(p_DEC) & " Current pierSide: " & TranslatePierSide(telescopeDevice.SideOfPier, True))
            Try
                l_Results.DestinationSideOfPier = telescopeDevice.DestinationSideOfPier(p_RA, p_DEC)
                LogMsg("SOPPierTest", MessageLevel.msgDebug, "Target DestinationSideOfPier: " & l_Results.DestinationSideOfPier.ToString)
            Catch ex As COMException
                Select Case ex.ErrorCode
                    Case ErrorCodes.NotImplemented
                        l_Results.DestinationSideOfPier = PierSide.pierUnknown
                        LogMsg("SOPPierTest", MessageLevel.msgDebug, "COM DestinationSideOfPier is not implemented setting result to: " & l_Results.DestinationSideOfPier.ToString)
                    Case Else
                        LogMsg("SOPPierTest", MessageLevel.msgError, "COM DestinationSideOfPier Exception: " & ex.ToString)
                End Select

            Catch ex As MethodNotImplementedException ' DestinationSideOfPier not available so mark as unknown
                l_Results.DestinationSideOfPier = PierSide.pierUnknown
                LogMsg("SOPPierTest", MessageLevel.msgDebug, ".NET DestinationSideOfPier is not implemented setting result to: " & l_Results.DestinationSideOfPier.ToString)
            Catch ex As Exception
                LogMsg("SOPPierTest", MessageLevel.msgError, ".NET DestinationSideOfPier Exception: " & ex.ToString)
            End Try
            'Now do an actual slew and record side of pier we actually get
            SlewScope(p_RA, p_DEC, "Testing " & p_Msg & ", co-ordinates: " & FormatRA(p_RA) & " " & FormatDec(p_DEC))
            l_Results.SideOfPier = telescopeDevice.SideOfPier
            LogMsg("SOPPierTest", MessageLevel.msgDebug, "Actual SideOfPier: " & l_Results.SideOfPier.ToString)

            'Return to original RA
            SlewScope(l_StartRA, l_StartDEC, "Returning to start point")
            LogMsg("SOPPierTest", MessageLevel.msgDebug, "Returned to: " & FormatRA(l_StartRA) & " " & FormatDec(l_StartDEC))
        Catch ex As Exception
            LogMsg("SOPPierTest", MessageLevel.msgError, "SideofPierException: " & ex.ToString)
        End Try
        Return l_Results
    End Function

    Private Sub DestinationSideOfPierTests()
        Dim l_PierSideMinus3, l_PierSideMinus9, l_PierSidePlus3, l_PierSidePlus9 As PierSide
        Dim l_Declination As Double = 70.0

        'Slew to one position, then call destination side of pier 4 times and report the pattern
        SlewScope(TelescopeRAFromHourAngle("DestinationSideofPier", -3.0), 0.0, "Slew to start position")
        l_PierSideMinus3 = telescopeDevice.DestinationSideOfPier(-3.0, 0.0)
        l_PierSidePlus3 = telescopeDevice.DestinationSideOfPier(3.0, 0.0)
        l_PierSideMinus9 = telescopeDevice.DestinationSideOfPier(-9.0, 90.0 - m_SiteLatitude)
        l_PierSidePlus9 = telescopeDevice.DestinationSideOfPier(9.0, 90.0 - m_SiteLatitude)
        If (l_PierSideMinus3 = l_PierSidePlus9) And (l_PierSidePlus3 = l_PierSideMinus9) Then 'Reporting physical pier side
            LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "The driver appears to be reporting physical pier side rather than pointing state")
        Else ' Make other tests
            If (l_PierSideMinus3 = l_PierSideMinus9) And (l_PierSidePlus3 = l_PierSidePlus9) Then
                LogMsg("DestinationSideofPier", MessageLevel.msgOK, "The driver reports the pointing state of the mount")
            Else 'Don't know what this means!
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Unknown pier side reporting model: HA-3: " & l_PierSideMinus3.ToString & " HA-9: " & l_PierSideMinus9.ToString & " HA+3: " & l_PierSidePlus3.ToString & " HA+9: " & l_PierSidePlus9.ToString)
            End If
        End If
        telescopeDevice.Tracking = False

        LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus9, False) & TranslatePierSide(l_PierSidePlus9, False))
        LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus3, False) & TranslatePierSide(l_PierSidePlus3, False))

    End Sub

#End Region

#Region "Support Code"

    Private Sub CheckScopePosition(testName As String, functionName As String, expectedRA As Double, expectedDec As Double)
        Dim actualRA, actualDec, difference As Double

        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get RightAscension property")
        actualRA = telescopeDevice.RightAscension
        LogMsg(testName, MessageLevel.msgDebug, "Read RightAscension: " & FormatRA(actualRA))

        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Declination property")
        actualDec = telescopeDevice.Declination
        LogMsg(testName, MessageLevel.msgDebug, "Read Declination: " & FormatDec(actualDec))

        ' Check that we have actually arrived where we are expected to be
        difference = RaDifferenceInSeconds(actualRA, expectedRA)

        Select Case difference
            Case Is <= SLEW_SYNC_OK_TOLERANCE  ' Convert arc seconds to hours of RA
                LogMsg(testName, MessageLevel.msgOK, String.Format("{0} OK. RA:   {1}", functionName, FormatRA(expectedRA)))
                Exit Select
            Case Else
                LogMsg(testName, MessageLevel.msgInfo, String.Format("{0} within {1} arc seconds of expected RA: {2}, actual RA: {3}", functionName, difference.ToString("0.0"), FormatRA(expectedRA), FormatRA(actualRA)))
        End Select

        difference = Math.Round(Math.Abs(actualDec - expectedDec) * 60.0 * 60.0, 1, MidpointRounding.AwayFromZero) ' Dec difference is in arc seconds from degrees of Declination
        Select Case difference
            Case Is <= SLEW_SYNC_OK_TOLERANCE
                LogMsg(testName, MessageLevel.msgOK, String.Format("{0} OK. DEC: {1}", functionName, FormatDec(expectedDec)))
                Exit Select
            Case Else
                LogMsg(testName, MessageLevel.msgInfo, String.Format("{0} within {1} arc seconds of expected DEC: {2}, actual DEC: {3}", functionName, difference.ToString("0.0"), FormatDec(expectedDec), FormatDec(actualDec)))
        End Select

    End Sub

    ''' <summary>
    ''' Return the difference between two RAs (in hours) as seconds
    ''' </summary>
    ''' <param name="FirstRA">First RA (hours)</param>
    ''' <param name="SecondRA">Second RA (hours)</param>
    ''' <returns>Difference (seconds) between the supplied RAs</returns>
    Private Function RaDifferenceInSeconds(FirstRA As Double, SecondRA As Double) As Double
        RaDifferenceInSeconds = Math.Abs(FirstRA - SecondRA) ' Calculate the difference allowing for negative outcomes

        If RaDifferenceInSeconds > 12.0 Then RaDifferenceInSeconds = 24.0 - RaDifferenceInSeconds ' Deal with the cases where the two elements are more than 12 hours apart going in the initial direction
        RaDifferenceInSeconds = Math.Round(RaDifferenceInSeconds * 15.0 * 60.0 * 60.0, 1, MidpointRounding.AwayFromZero) ' RA difference is in arc seconds from hours of RA

        Return RaDifferenceInSeconds
    End Function

    Private Sub SyncScope(testName As String, canDoItName As String, testType As SlewSyncType, syncRA As Double, syncDec As Double)
        Select Case testType
            Case SlewSyncType.SyncToCoordinates 'SyncToCoordinates
                If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Tracking property")
                If canSetTracking And (Not telescopeDevice.Tracking) Then
                    If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to true")
                    telescopeDevice.Tracking = True
                End If
                If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to call SyncToCoordinates method, RA: " & FormatRA(syncRA) & ", Declination: " & FormatDec(syncDec))
                telescopeDevice.SyncToCoordinates(syncRA, syncDec) 'Sync to slightly different coordinates
                LogMsg(testName, MessageLevel.msgDebug, "Completed SyncToCoordinates")
            Case SlewSyncType.SyncToTarget 'SyncToTarget
                If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Tracking property")
                If canSetTracking And (Not telescopeDevice.Tracking) Then
                    If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to true")
                    telescopeDevice.Tracking = True
                End If
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to set TargetRightAscension property to " & FormatRA(syncRA))
                    telescopeDevice.TargetRightAscension = syncRA
                    LogMsg(testName, MessageLevel.msgDebug, "Completed Set TargetRightAscension")
                Catch ex As Exception
                    HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, canDoItName & " is True but can't set TargetRightAscension")
                End Try
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to set TargetDeclination property to " & FormatDec(syncDec))
                    telescopeDevice.TargetDeclination = syncDec
                    LogMsg(testName, MessageLevel.msgDebug, "Completed Set TargetDeclination")
                Catch ex As Exception
                    HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, canDoItName & " is True but can't set TargetDeclination")
                End Try

                If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to call SyncToTarget method")
                telescopeDevice.SyncToTarget() 'Sync to slightly different coordinates
                LogMsg(testName, MessageLevel.msgDebug, "Completed SyncToTarget")
            Case Else
                LogMsg(testName, MessageLevel.msgError, "Conform:SyncTest: Unknown test type " & testType.ToString)
        End Select

    End Sub

    Sub SlewScope(ByVal p_RA As Double, ByVal p_DEC As Double, ByVal p_Msg As String)

        If canSetTracking Then
            If g_Settings.DisplayMethodCalls Then LogMsg("SlewScope", MessageLevel.msgComment, "About to set Tracking property to true")
            telescopeDevice.Tracking = True
        End If
        Status(StatusType.staAction, p_Msg)

        If canSlew Then
            If canSlewAsync Then
                LogMsg("SlewScope", MessageLevel.msgDebug, "Slewing asynchronously to " + p_Msg + " " + FormatRA(p_RA) + " " + FormatDec(p_DEC))
                If g_Settings.DisplayMethodCalls Then LogMsg("SlewScope", MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method, RA: " & FormatRA(p_RA) & ", Declination: " & FormatDec(p_DEC))
                telescopeDevice.SlewToCoordinatesAsync(p_RA, p_DEC)
                WaitForSlew("SlewScope")
            Else
                LogMsg("SlewScope", MessageLevel.msgDebug, "Slewing synchronously to " + p_Msg + " " + FormatRA(p_RA) + " " + FormatDec(p_DEC))
                If g_Settings.DisplayMethodCalls Then LogMsg("SlewScope", MessageLevel.msgComment, "About to call SlewToCoordinates method, RA: " & FormatRA(p_RA) & ", Declination: " & FormatDec(p_DEC))
                telescopeDevice.SlewToCoordinates(p_RA, p_DEC)
            End If

            If m_CanReadSideOfPier Then
                If g_Settings.DisplayMethodCalls Then LogMsg("SlewScope", MessageLevel.msgComment, "About to get SideOfPier property")
                LogMsg("SlewScope", MessageLevel.msgDebug, "SideOfPier: " & telescopeDevice.SideOfPier.ToString())
            End If
        Else
            LogMsg("SlewScope", MessageLevel.msgInfo, "Unable to slew this scope as CanSlew is false, slew omitted")
        End If
        Status(StatusType.staAction, "")
    End Sub

    Private Sub WaitForSlew(testName As String)
        Dim WaitStartTime As DateTime
        WaitStartTime = Now
        Do
            WaitFor(SLEEP_TIME)
            FrmConformMain.staStatus.Text = "Slewing"
            Application.DoEvents()
            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get Slewing property")
        Loop Until ((Not telescopeDevice.Slewing) And (Now.Subtract(WaitStartTime).TotalSeconds > WAIT_FOR_SLEW_MINIMUM_DURATION)) Or TestStop()
        FrmConformMain.staStatus.Text = "Slew completed"
    End Sub

    Private Function TelescopeRAFromHourAngle(testName As String, ByVal p_Offset As Double) As Double

        ' Handle the possibility that the mandatory SideealTime property has not been implemented
        If canReadSiderealTime Then
            'Create a legal RA based on an offset from Sidereal time
            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get SiderealTime property")
            TelescopeRAFromHourAngle = telescopeDevice.SiderealTime - p_Offset
            Select Case TelescopeRAFromHourAngle
                Case Is < 0.0 'Illegal if < 0 hours
                    TelescopeRAFromHourAngle += 24.0
                Case Is >= 24.0 'Illegal if > 24 hours
                    TelescopeRAFromHourAngle -= 24.0
            End Select
        Else
            TelescopeRAFromHourAngle = 0.0 - p_Offset
        End If
    End Function

    Private Function TelescopeRAFromSiderealTime(testName As String, ByVal p_Offset As Double) As Double
        Dim CurrentSiderealTime As Double

        ' Handle the possibility that the mandatory SideealTime property has not been implemented
        If canReadSiderealTime Then
            'Create a legal RA based on an offset from Sidereal time
            If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to get SiderealTime property")
            CurrentSiderealTime = telescopeDevice.SiderealTime
            Select Case CurrentSiderealTime 'Deal with possibility that sidereal time from the driver is bad
                Case Is < 0.0 'Illegal if < 0 hours
                    CurrentSiderealTime = 0
                Case Is >= 24.0 'Illegal if > 24 hours
                    CurrentSiderealTime = 0
            End Select
            TelescopeRAFromSiderealTime = CurrentSiderealTime + p_Offset
            Select Case TelescopeRAFromSiderealTime
                Case Is < 0.0 'Illegal if < 0 hours
                    TelescopeRAFromSiderealTime = TelescopeRAFromSiderealTime + 24.0
                Case Is >= 24.0 'Illegal if > 24 hours
                    TelescopeRAFromSiderealTime = TelescopeRAFromSiderealTime - 24.0
            End Select
        Else
            TelescopeRAFromSiderealTime = 0.0 + p_Offset
        End If
    End Function

    Private Sub TestEarlyBinding(ByVal TestType As InterfaceType)
        Dim l_ITelescope As Object
        Dim l_DriverAccessTelescope As ASCOM.DriverAccess.Telescope = Nothing
        Dim l_DeviceObject As Object = Nothing, l_ErrMsg As String = ""
        Dim l_TryCount As Integer = 0
        Try
            'Try early binding
            l_ITelescope = Nothing
            Do
                l_TryCount += 1
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to create driver object with CreateObject")
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Creating late bound object for interface test")
                    l_DeviceObject = CreateObject(g_TelescopeProgID)
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Created late bound object OK")
                    Select Case TestType
                        Case InterfaceType.ITelescopeV2
                            l_ITelescope = CType(l_DeviceObject, ASCOM.Interface.ITelescope)
                        Case InterfaceType.ITelescopeV3
                            l_ITelescope = CType(l_DeviceObject, ITelescopeV3)
                        Case Else
                            LogMsg("TestEarlyBinding", MessageLevel.msgError, "Unknown interface type: " & TestType.ToString)
                    End Select
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with interface " & TestType.ToString)
                    Try
                        If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property true")
                        l_ITelescope.Connected = True
                        LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface " & TestType.ToString)
                        If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property false")
                        l_ITelescope.Connected = False
                    Catch ex As Exception
                        LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface " & TestType.ToString)
                        LogMsg("", MessageLevel.msgAlways, "")
                    End Try
                Catch ex As Exception
                    l_ErrMsg = ex.ToString()
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Exception: " & ex.Message)
                End Try
                If l_DeviceObject Is Nothing Then WaitFor(200)
            Loop Until ((l_TryCount = 3) Or (Not l_ITelescope Is Nothing)) 'Exit if created OK
            If l_ITelescope Is Nothing Then
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface " & TestType.ToString)
            Else
                LogMsg("AccessChecks", MessageLevel.msgDebug, "Created telescope on attempt: " & l_TryCount.ToString)
            End If

            'Clean up
            Try : DisposeAndReleaseObject("Telescope V1", l_ITelescope) : Catch : End Try
            Try : DisposeAndReleaseObject("Telescope V3", l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_ITelescope = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()

            WaitForAbsolute(DEVICE_DESTROY_WAIT, "TestEarlyBinding waiting for Telescope Object to Dispose")
        Catch ex As Exception
            LogMsg("Telescope:TestEarlyBinding.EX1", MessageLevel.msgError, ex.ToString)
        End Try
    End Sub

    Private Function FormatRA(ra As Double) As String
        Return g_Util.HoursToHMS(ra, ":", ":", "", DISPLAY_DECIMAL_DIGITS)
    End Function

    Private Function FormatDec(Dec As Double) As String
        Return g_Util.DegreesToDMS(Dec, ":", ":", "", DISPLAY_DECIMAL_DIGITS).PadLeft(9 + IIf(DISPLAY_DECIMAL_DIGITS > 0, DISPLAY_DECIMAL_DIGITS + 1, 0))
    End Function

    Private Function FormatAltitude(Alt As Double)
        Return g_Util.DegreesToDMS(Alt, ":", ":", "", DISPLAY_DECIMAL_DIGITS)
    End Function

    Private Function FormatAzimuth(Az As Double) As String
        Return g_Util.DegreesToDMS(Az, ":", ":", "", DISPLAY_DECIMAL_DIGITS).PadLeft(9 + IIf(DISPLAY_DECIMAL_DIGITS > 0, DISPLAY_DECIMAL_DIGITS + 1, 0))
    End Function

    Function TranslatePierSide(ByVal p_PierSide As PierSide, ByVal p_Long As Boolean) As String
        Dim l_PierSide As String
        Select Case p_PierSide
            Case PierSide.pierEast
                If p_Long Then
                    l_PierSide = "pierEast"
                Else
                    l_PierSide = "E"
                End If

            Case PierSide.pierWest
                If p_Long Then
                    l_PierSide = "pierWest"
                Else
                    l_PierSide = "W"
                End If
            Case Else
                If p_Long Then
                    l_PierSide = "pierUnknown"
                Else
                    l_PierSide = "U"
                End If
        End Select
        Return l_PierSide
    End Function

    Private Enum Axis
        RA
        Dec
    End Enum

    Private Function TestRADecRate(TestName As String, Description As String, Axis As Axis, Rate As Double, SkipSlewiingTest As Boolean) As Boolean
        Dim success As Boolean = False

        Try
            If g_Settings.DisplayMethodCalls Then LogMsg(TestName, MessageLevel.msgComment, String.Format("{0} - About to get Slewing property", Description))
            m_Slewing = telescopeDevice.Slewing
            If (Not m_Slewing) Or SkipSlewiingTest Then ' Slewing should be false at this point or we are ignoring the test!
                'Check that we can set the rate to a non-zero value
                Try
                    Select Case Axis
                        Case Axis.RA
                            If g_Settings.DisplayMethodCalls Then LogMsg(TestName, MessageLevel.msgComment, String.Format("{0} - About to set RightAscensionRate property to {1}", Description, Rate))
                            telescopeDevice.RightAscensionRate = Rate
                            SetStatus(String.Format("Watling for mount to settle after setting RightAcensionRate to {0}", Rate), "", "")
                            WaitFor(2000) ' Give a short wait to allow the mount to settle

                            ' Value set OK, now check that the new rate is returned by RightAscensionRate Get and that Slewing is false
                            If g_Settings.DisplayMethodCalls Then LogMsg(TestName, MessageLevel.msgComment, String.Format("{0} - About to get RightAscensionRate property", Description))
                            m_RightAscensionRate = telescopeDevice.RightAscensionRate
                            If g_Settings.DisplayMethodCalls Then LogMsg(TestName, MessageLevel.msgComment, String.Format("{0} - About to get Slewing property", Description))
                            m_Slewing = telescopeDevice.Slewing

                            If (m_RightAscensionRate = Rate) And Not m_Slewing Then
                                LogMsg(TestName, MessageLevel.msgOK, String.Format("{0} - successfully set rate to {1}", Description, m_RightAscensionRate))
                                success = True
                            Else
                                If m_Slewing And (m_RightAscensionRate = Rate) Then LogMsg(TestName, MessageLevel.msgError, String.Format("RightAscensionRate was successfully set to {0} but Slewing is returning True, it should return False.", Rate, m_RightAscensionRate))
                                If m_Slewing And (m_RightAscensionRate <> Rate) Then LogMsg(TestName, MessageLevel.msgError, String.Format("RightAscensionRate Read does not return {0} as set, instead it returns {1}. Slewing is also returning True, it should return False.", Rate, m_RightAscensionRate))
                                If Not m_Slewing And (m_RightAscensionRate <> Rate) Then LogMsg(TestName, MessageLevel.msgError, String.Format("RightAscensionRate Read does not return {0} as set, instead it returns {1}.", Rate, m_RightAscensionRate))
                            End If

                        Case Axis.Dec
                            If g_Settings.DisplayMethodCalls Then LogMsg(TestName, MessageLevel.msgComment, String.Format("{0} - About to set DeclinationRate property to {1}", Description, Rate))
                            telescopeDevice.DeclinationRate = Rate
                            SetStatus(String.Format("Watling for mount to settle after setting DeclinationRate to {0}", Rate), "", "")
                            WaitFor(2000) ' Give a short wait to allow the mount to settle

                            ' Value set OK, now check that the new rate is returned by DeclinationRate Get and that Slewing is false
                            If g_Settings.DisplayMethodCalls Then LogMsg(TestName, MessageLevel.msgComment, String.Format("{0} - About to get DeclinationRate property", Description))
                            m_DeclinationRate = telescopeDevice.DeclinationRate
                            If g_Settings.DisplayMethodCalls Then LogMsg(TestName, MessageLevel.msgComment, String.Format("{0} - About to get Slewing property", Description))
                            m_Slewing = telescopeDevice.Slewing

                            If (m_DeclinationRate = Rate) And Not m_Slewing Then
                                LogMsg(TestName, MessageLevel.msgOK, String.Format("{0} - successfully set rate to {1}", Description, m_DeclinationRate))
                                success = True
                            Else
                                If m_Slewing And (m_DeclinationRate = Rate) Then LogMsg(TestName, MessageLevel.msgError, String.Format("DeclinationRate was successfully set to {0} but Slewing is returning True, it should return False.", Rate, m_DeclinationRate))
                                If m_Slewing And (m_DeclinationRate <> Rate) Then LogMsg(TestName, MessageLevel.msgError, String.Format("DeclinationRate Read does not return {0} as set, instead it returns {1}. Slewing is also returning True, it should return False.", Rate, m_DeclinationRate))
                                If Not m_Slewing And (m_DeclinationRate <> Rate) Then LogMsg(TestName, MessageLevel.msgError, String.Format("DeclinationRate Read does not return {0} as set, instead it returns {1}.", Rate, m_DeclinationRate))
                            End If

                        Case Else
                            MessageBox.Show(String.Format("Conform internal error - Unknown Axis value: {0}", Axis.ToString()))
                    End Select

                Catch ex As Exception
                    If IsInvalidOperationException(TestName, ex) Then ' We can't know what the valid range for this telescope is in advance so its possible that our test value will be rejected, if so just report this.
                        LogMsg(TestName, MessageLevel.msgInfo, String.Format("Unable to set test rate {0}, it was rejected as an invalid value.", Rate))
                    Else
                        HandleException(TestName, MemberType.Property, Required.MustBeImplemented, ex, "CanSetRightAscensionRate is True")
                    End If
                End Try
            Else
                LogMsg(TestName, MessageLevel.msgError, String.Format("{0} - Telescope.Slewing should be False at the start of this test but is returning True, test abandoned", Description))
            End If
        Catch ex As Exception
            HandleException(TestName, MemberType.Property, Required.MustBeImplemented, ex, "Tried to read Slewing property")
        End Try
        SetStatus("", "", "")

        Return success
    End Function

#End Region

#Region "Private Test Code"

    Protected Overrides Sub SpecialTelescopeDestinationSideOfPier()
        Dim l_FlipPointEast, l_FlipPointWest As Double
        CheckInitialise() ' Set up error codes
        CreateDevice() 'Create the telescope device
        Connected = True 'Connect
        PreRunCheck() 'Make sure we are ready to move the scope
        telescopeDevice.Tracking = True
        If (g_Settings.DSOPSide = "pierEast only") Or (g_Settings.DSOPSide = "pierEast and pierWest") Then
            l_FlipPointEast = FlipCheck(FlipTestType.DestinationSideOfPier,
                              g_Settings.FlipTestDECStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast),
                              g_Settings.FlipTestDECEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast),
                              g_Settings.FlipTestDECStep(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast),
                              g_Settings.FlipTestHAStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast),
                              g_Settings.FlipTestHAEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierEast))
        End If

        If (g_Settings.DSOPSide = "pierWest only") Or (g_Settings.DSOPSide = "pierEast and pierWest") Then
            l_FlipPointWest = FlipCheck(FlipTestType.DestinationSideOfPier,
                              g_Settings.FlipTestDECStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest),
                              g_Settings.FlipTestDECEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest),
                              g_Settings.FlipTestDECStep(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest),
                              g_Settings.FlipTestHAStart(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest),
                              g_Settings.FlipTestHAEnd(GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier, PierSide.pierWest))
        End If
        PostRunCheck() 'Make the mount safe
        Connected = False 'Disconnect

    End Sub

    Protected Overrides Sub SpecialTelescopeSideOfPier()
        Dim l_FlipPointEast, l_FlipPointWest As Double
        CheckInitialise() ' Set up error codes
        CreateDevice() 'Create the telescope device
        Connected = True 'Connect
        PreRunCheck() 'Make sure we are ready to move the scope
        telescopeDevice.Tracking = True
        If (g_Settings.DSOPSide = "pierEast only") Or (g_Settings.DSOPSide = "pierEast and pierWest") Then
            l_FlipPointEast = FlipCheck(FlipTestType.SideOfPier,
                              g_Settings.FlipTestDECStart(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierEast),
                              g_Settings.FlipTestDECEnd(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierEast),
                              g_Settings.FlipTestDECStep(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierEast),
                              g_Settings.FlipTestHAStart(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierEast),
                              g_Settings.FlipTestHAEnd(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierEast))
        End If
        If (g_Settings.DSOPSide = "pierWest only") Or (g_Settings.DSOPSide = "pierEast and pierWest") Then
            l_FlipPointWest = FlipCheck(FlipTestType.SideOfPier,
                              g_Settings.FlipTestDECStart(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierWest),
                              g_Settings.FlipTestDECEnd(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierWest),
                              g_Settings.FlipTestDECStep(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierWest),
                              g_Settings.FlipTestHAStart(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierWest),
                              g_Settings.FlipTestHAEnd(GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier, PierSide.pierWest))
        End If
        PostRunCheck() 'Make the mount safe
        Connected = False 'Disconnect
    End Sub

    Protected Overrides Sub SpecialTelescopeSideOfPierAnalysis()
        CheckInitialise() ' Set up error codes
        CreateDevice() 'Create the telescope device
        Connected = True 'Connect
        PreRunCheck() 'Make sure we are ready to move the scope
        telescopeDevice.Tracking = True
        canSlew = True

        'SideOfPierTests()
        Dim l_PierSideMinus3, l_PierSideMinus9, l_PierSidePlus3, l_PierSidePlus9 As SideOfPierResults
        Dim l_Declination, l_StartRA As Double

        'Slew to starting position
        Status(StatusType.staTest, "Side of pier tests")
        l_StartRA = TelescopeRAFromHourAngle("SideofPier", -3.0)
        l_Declination = 60.0

        SlewScope(l_StartRA, 0.0, "Move to starting position " & FormatRA(l_StartRA) & " " & FormatDec(0.0))
        If TestStop() Then Exit Sub

        'Run tests
        Status(StatusType.staAction, "Test hour angle -3.0")
        l_PierSideMinus3 = SOPPierTest(l_StartRA, 10.0, "hour angle -3.0 +") : If TestStop() Then Exit Sub
        l_PierSideMinus3 = SOPPierTest(l_StartRA, -10.0, "hour angle -3.0 -") : If TestStop() Then Exit Sub
        Status(StatusType.staAction, "Test hour angle +3.0")
        l_PierSidePlus3 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +3.0), 10.0, "hour angle +3.0 +") : If TestStop() Then Exit Sub
        l_PierSidePlus3 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +3.0), -10.0, "hour angle +3.0 -") : If TestStop() Then Exit Sub
        Status(StatusType.staAction, "Test hour angle -9.0")
        l_PierSideMinus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", -9.0), l_Declination, "hour angle -9.0") : If TestStop() Then Exit Sub
        Status(StatusType.staAction, "Test hour angle +9.0")
        l_PierSidePlus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +9.0), l_Declination, "hour angle +9.0") : If TestStop() Then Exit Sub

        If (l_PierSideMinus3.SideOfPier = l_PierSidePlus9.SideOfPier) And (l_PierSidePlus3.SideOfPier = l_PierSideMinus9.SideOfPier) Then 'Reporting physical pier side
            LogMsg("SideofPier", MessageLevel.msgIssue, "SideofPier reports physical pier side rather than pointing state")
            LogMsg("SideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus9.SideOfPier, False) & TranslatePierSide(l_PierSidePlus9.SideOfPier, False))
            LogMsg("SideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus3.SideOfPier, False) & TranslatePierSide(l_PierSidePlus3.SideOfPier, False))
        Else ' Make other tests
            If (l_PierSideMinus3.SideOfPier = l_PierSideMinus9.SideOfPier) And (l_PierSidePlus3.SideOfPier = l_PierSidePlus9.SideOfPier) Then
                LogMsg("SideofPier", MessageLevel.msgOK, "Reports the pointing state of the mount as expected")
            Else 'Don't know what this means!
                LogMsg("SideofPier", MessageLevel.msgInfo, "Unknown SideofPier reporting model: HA-3: " & l_PierSideMinus3.SideOfPier.ToString & " HA-9: " & l_PierSideMinus9.SideOfPier.ToString & " HA+3: " & l_PierSidePlus3.SideOfPier.ToString & " HA+9: " & l_PierSidePlus9.SideOfPier.ToString)
                LogMsg("SideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus9.SideOfPier, False) & TranslatePierSide(l_PierSidePlus9.SideOfPier, False))
                LogMsg("SideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus3.SideOfPier, False) & TranslatePierSide(l_PierSidePlus3.SideOfPier, False))
            End If
        End If

        'Test whether DestinationSideOfPier is implemented
        If l_PierSideMinus3.DestinationSideOfPier = PierSide.pierUnknown And
                l_PierSideMinus9.DestinationSideOfPier = PierSide.pierUnknown And
                l_PierSidePlus3.DestinationSideOfPier = PierSide.pierUnknown And
                l_PierSidePlus9.DestinationSideOfPier = PierSide.pierUnknown Then
            LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Analysis skipped as this method is not implemented") 'Not implemented
        Else 'It is implemented so assess the results

            If (l_PierSideMinus3.DestinationSideOfPier = l_PierSidePlus9.DestinationSideOfPier) And (l_PierSidePlus3.DestinationSideOfPier = l_PierSideMinus9.DestinationSideOfPier) Then 'Reporting physical pier side
                LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "DestinationSideofPier reports physical pier side rather than pointing state")
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus9.DestinationSideOfPier, False) & TranslatePierSide(l_PierSidePlus9.DestinationSideOfPier, False))
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus3.DestinationSideOfPier, False) & TranslatePierSide(l_PierSidePlus3.DestinationSideOfPier, False))
            Else ' Make other tests
                If (l_PierSideMinus3.DestinationSideOfPier = l_PierSideMinus9.DestinationSideOfPier) And (l_PierSidePlus3.DestinationSideOfPier = l_PierSidePlus9.DestinationSideOfPier) Then
                    LogMsg("DestinationSideofPier", MessageLevel.msgOK, "Reports the pointing state of the mount as expected")
                Else 'Don't know what this means!
                    LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Unknown DestinationSideofPier reporting model: HA-3: " & l_PierSideMinus3.SideOfPier.ToString & " HA-9: " & l_PierSideMinus9.SideOfPier.ToString & " HA+3: " & l_PierSidePlus3.SideOfPier.ToString & " HA+9: " & l_PierSidePlus9.SideOfPier.ToString)
                    LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus9.DestinationSideOfPier, False) & TranslatePierSide(l_PierSidePlus9.DestinationSideOfPier, False))
                    LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus3.DestinationSideOfPier, False) & TranslatePierSide(l_PierSidePlus3.DestinationSideOfPier, False))
                End If
            End If
        End If

        'Clean up
        '3.0.0.12 added conditional test to next line
        If canSetTracking Then telescopeDevice.Tracking = False
        Status(StatusType.staStatus, "")
        Status(StatusType.staAction, "")
        Status(StatusType.staTest, "")

        PostRunCheck() 'Make the mount safe
        Connected = False 'Disconnected
    End Sub

    Protected Overrides Sub SpecialTelescopeCommands()
        Dim Response As String
        CheckInitialise() ' Set up error codes
        CreateDevice() 'Create the telescope device
        Connected = True 'Connect
        PreRunCheck() 'Make sure we are ready to move the scope
        telescopeDevice.Tracking = True

        'Command tests
        Response = telescopeDevice.CommandString("Gc", False)
        LogMsg("TelescopeCommandTest", MessageLevel.msgInfo, "Sent Gc False - Received: " & Response)
        Response = telescopeDevice.CommandString(":Gc#", True)
        LogMsg("TelescopeCommandTest", MessageLevel.msgInfo, "Sent :Gc# True - Received: " & Response)

        PostRunCheck() 'Make the mount safe
        Connected = False 'Disconnected

    End Sub

    Function FlipCheck(ByVal p_FlipTestType As FlipTestType, ByVal p_DECStart As Double, ByVal p_DECEnd As Double, ByVal p_DECStep As Double, ByVal p_HAStart As Double, ByVal p_HAEnd As Double) As Double
        Dim l_FlipPoint, l_TargetDEC As Double


        LogMsg("FlipCheck", MessageLevel.msgOK, "Test parameters - Start DEC: " & p_DECStart.ToString & " End DEC: " & p_DECEnd.ToString & " DEC Step size: " & p_DECStep.ToString & " Start HA: " & p_HAStart & " End HA: " & p_HAEnd)
        LogMsg("", MessageLevel.msgAlways, "")

        For l_TargetDEC = p_DECStart To p_DECEnd Step p_DECStep
            If Not TestStop() Then l_FlipPoint = TelescopeFindFlipPoint(p_FlipTestType, p_HAStart, p_HAEnd, l_TargetDEC)
            If Not TestStop() Then LogMsg(p_FlipTestType.ToString, MessageLevel.msgOK, "Hour angle flip point when scope is east of pier (HH:MM:SS):" & FormatRA(l_FlipPoint) & " at DEC: " & FormatDec(l_TargetDEC))
        Next
        Return l_FlipPoint
    End Function

    Function TelescopeFindFlipPoint(ByVal p_TestType As FlipTestType, ByVal p_HARangeStart As Double, ByVal p_HARangeEnd As Double, ByVal p_TargetDEC As Double) As Double
        'Returns the flip point returned either by SideofPier or DestinationSideOfPier
        Dim l_SideOfPier, l_TargetSideOfPier As PierSide, l_TargetHA As Double
        Dim l_Iteration As Integer, l_StartHARange, l_EndHARange, l_Temp As Double

        l_StartHARange = p_HARangeStart
        l_TargetHA = p_HARangeEnd
        l_EndHARange = p_HARangeEnd

        'Use a halving method to determine where the flip point is
        For l_Iteration = 1 To 30
            'Move scope to requested starting position
            SlewScope(TelescopeRAFromHourAngle("FlipRange: " & l_SideOfPier.ToString, p_HARangeStart), p_TargetDEC, "Moving to start position")
            l_SideOfPier = telescopeDevice.SideOfPier 'Save current side of pier as reference to check when a flip happens
            Select Case p_TestType ' Carry out relevant test
                Case FlipTestType.DestinationSideOfPier
                    l_TargetSideOfPier = telescopeDevice.DestinationSideOfPier(TelescopeRAFromHourAngle("FlipRange: " & l_SideOfPier.ToString, l_TargetHA), p_TargetDEC)
                Case FlipTestType.SideOfPier
                    'Move scope to requested position
                    SlewScope(TelescopeRAFromHourAngle("FlipRange: " & l_SideOfPier.ToString, l_TargetHA), p_TargetDEC, "Moving to requested position")
                    l_TargetSideOfPier = telescopeDevice.SideOfPier 'Record pier side of target RA, DEC
                Case Else
                    MsgBox("TelescopeFindFlipPoint - unexpected test type: " & p_TestType.ToString, MsgBoxStyle.Critical)
            End Select
            Select Case l_TargetSideOfPier 'Check outcome in terms of pier side and then adjust parameters to home in on correct value
                Case l_SideOfPier 'We have come back too far so go further out
                    l_Temp = l_TargetHA
                    l_TargetHA = l_TargetHA + (l_EndHARange - l_StartHARange) / 2.0
                    l_StartHARange = l_Temp
                    LogMsg("FlipRange: " & l_SideOfPier.ToString, MessageLevel.msgDebug, "No Flip - Start, Target, End HA: " & FormatRA(l_StartHARange) & " " & FormatRA(l_TargetHA) & " " & FormatRA(l_EndHARange) & " " & FormatDec(p_TargetDEC))
                Case PierSide.pierUnknown
                    LogMsg("TelescopeFlipRange", MessageLevel.msgError, "Unexpected pier side: " & l_TargetSideOfPier.ToString)
                Case Else 'Side of pier has changed so come back a bit
                    l_Temp = l_TargetHA
                    l_TargetHA = l_TargetHA - (l_EndHARange - l_StartHARange) / 2.0
                    l_EndHARange = l_Temp
                    LogMsg("FlipRange: " & l_SideOfPier.ToString, MessageLevel.msgDebug, "Flipped - Start, Target, End HA: " & FormatRA(l_StartHARange) & " " & FormatRA(l_TargetHA) & " " & FormatRA(l_EndHARange) & " " & FormatDec(p_TargetDEC))
            End Select
            If TestStop() Then Exit For
        Next

        Return l_TargetHA
    End Function

#End Region

End Class