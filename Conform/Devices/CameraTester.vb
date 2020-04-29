Imports ASCOM.DeviceInterface
Imports ASCOM.DriverAccess
Friend Class CameraTester
    Inherits DeviceTesterBaseClass
#Region "Variables and Constants"
    Const CAMERA_PULSE_DURATION As Integer = 2000 'Duration of camera pulse guide test (ms)
    Const CAMERA_PULSE_TOLERANCE As Integer = 300 'Tolerance for acceptable performance (ms)

    'Camera variables
    Private m_CanAbortExposure, m_CanAsymmetricBin, m_CanGetCoolerPower, m_CanSetCCDTemperature, m_CanStopExposure, m_CanFastReadout As Boolean
    Private m_CoolerOn, m_HasShutter, m_ImageReady As Boolean
    Private m_CameraXSize, m_CameraYSize, m_MaxADU, m_NumX, m_NumY, m_StartX, m_StartY As Integer
    Private m_MaxBinX, m_MaxBinY, m_BinX, m_BinY As Short
    Private m_CCDTemperature, m_CoolerPower, m_ElectronsPerADU, m_FullWellCapacity, m_HeatSinkTemperature, m_LastExposureDuration As Double
    Private m_PixelSizeX, m_PixelSizeY, m_SetCCDTemperature As Double
    Private m_LastError, m_LastExposureStartTime As String
    Private m_CameraState As ASCOM.DeviceInterface.CameraStates
    Private m_ImageArray, m_ImageArrayVariant As Array
    Private m_IsPulseGuidingSupported As Boolean 'Confirm that IsPulseGuiding command will work
    Private m_CanPulseGuide As Boolean
    Private m_Description As String
    Private m_IsPulseGuiding As Boolean
    'ICameraV2 properties
    Private m_BayerOffsetX, m_BayerOffsetY, m_Gain, m_GainMax, m_GainMin, m_PercentCompleted, m_ReadoutMode As Short
    Private m_ExposureMax, m_ExposureMin, m_ExposureResolution As Double
    Private m_FastReadout, m_CanReadGain, m_CanReadGainMax, m_CanReadGainMin, m_CanReadGains, m_CanReadReadoutModes As Boolean
    Private m_Gains, m_ReadoutModes As ArrayList
    Private m_SensorName As String
    Private m_SensorType As SensorType, m_CanReadSensorType As Boolean = False
    Private sw As Stopwatch = New Stopwatch()

    'ICameraV3 properties
    Private m_Offset, m_OffsetMax, m_OffsetMin As Integer
    Private m_CanReadOffset, m_CanReadOffsetMax, m_CanReadOffsetMin, m_CanReadOffsets, m_CanReadPulseGuideStatus As Boolean
    Private m_Offsets As ArrayList
    Private m_SubExposureDuration As Double
    'Private m_PulseGuideStatus As PulseGuideState

#If DEBUG Then
    Private m_Camera As ASCOM.DriverAccess.Camera
#Else
    Private m_Camera As Object
#End If
#End Region

#Region "Enums"
    Private Enum CanType
        tstCanFindHome = 1
        tstCanPark = 2
        tstCanPulseGuide = 3
        tstCanSetDeclinationRate = 4
        tstCanSetGuideRates = 5
        tstCanSetPark = 6
        tstCanSetPierSide = 7
        tstCanSetRightAscensionRate = 8
        tstCanSetTracking = 9
        tstCanSlew = 10
        tstCanSlewAsync = 11
        tstCanSlewAltAz = 12
        tstCanSlewAltAzAsync = 13
        tstCanSync = 14
        tstCanSyncAltAz = 15
        tstCanUnPark = 16
        tstCanAbortExposure = 17
        tstCanAsymmetricBin = 18
        tstCanGetCoolerPower = 19
        tstCanSetCCDTemperature = 20
        tstCanStopExposure = 21
        ' ICameraV2 property
        tstCanFastReadout = 22
    End Enum
    Private Enum CameraPerformance As Integer
        CameraState
        CCDTemperature
        CoolerPower
        HeatSinkTemperature
        ImageReady
        IsPulseGuiding
        ImageArray
        ImageArrayVariant
    End Enum
    Private Enum CamPropertyType
        ' ICameraV1 Properties
        BinX = 1
        BinY = 2
        CameraState = 3
        CameraXSize = 4
        CameraYSize = 5
        CCDTemperature = 6
        CoolerOn = 7
        CoolerPower = 8
        Description = 9
        DriverInfo = 10
        ElectronsPerADU = 11
        FullWellCapacity = 12
        HasShutter = 13
        HeatSinkTemperature = 14
        ImageReady = 15
        IsPulseGuiding = 16
        MaxADU = 17
        MaxBinX = 18
        MaxBinY = 19
        NumX = 20
        NumY = 21
        PixelSizeX = 22
        PixelSizeY = 23
        SetCCDTemperature = 24
        StartX = 25
        StartY = 26

        ' ICameraV2 Properties
        BayerOffsetX = 27
        BayerOffsetY = 28
        ExposureMax = 29
        ExposureMin = 30
        ExposureResolution = 31
        FastReadout = 32
        Gain = 33
        GainMax = 34
        GainMin = 35
        Gains = 36
        PercentCompleted = 37
        ReadoutMode = 38
        ReadoutModes = 39
        SensorName = 40
        SensorType = 41

        ' ICameraV3 Properties
        SubExposureDuration = 42

    End Enum

#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(True, True, True, True, False, True, True) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then

                Try : DisposeAndReleaseObject("Camera", m_Camera) : Catch : End Try

                m_Camera = Nothing
                m_ImageArray = Nothing
                m_ImageArrayVariant = Nothing
                GC.Collect()
            End If
        End If

        MyBase.Dispose(disposing)
        disposedValue = True
    End Sub
#End Region

#Region "Code"
    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(m_Camera, DeviceType.Camera)
    End Sub

    Overrides Sub CheckInitialise()
        'Set the error type numbers according to the standards adopted by individual authors.
        'Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
        'messages to driver authors!
        Select Case g_CameraProgID.ToUpper
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040405
                g_ExInvalidValue2 = &H80040405
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_CameraProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_DeviceObject As Object
        Dim l_ICamera As Object, l_DriverAccessCamera As ASCOM.DriverAccess.Camera
        MyBase.CheckAccessibility(g_CameraProgID, DeviceType.Camera)

        'Try early binding to V1 interface
        l_ICamera = Nothing
        l_DeviceObject = Nothing
        Try
            l_DeviceObject = CreateObject(g_CameraProgID)
            l_ICamera = CType(l_DeviceObject, ASCOM.Interface.ICamera)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface ICamera")
            Try
                l_ICamera.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ICamera")
                l_ICamera.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ICamera")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface ICamera")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_ICamera) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_ICamera = Nothing
            GC.Collect()
        End Try

        'Try early binding to V2 interface
        Try
            l_DeviceObject = CreateObject(g_CameraProgID)
            l_ICamera = CType(l_DeviceObject, ASCOM.DeviceInterface.ICameraV2)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface ICameraV2")
            Try
                l_ICamera.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ICameraV2")
                l_ICamera.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ICameraV2")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface ICameraV2")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_ICamera) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_ICamera = Nothing
            GC.Collect()
        End Try

        'Try client access toolkit 
        l_DriverAccessCamera = Nothing
        Try
            l_DriverAccessCamera = New ASCOM.DriverAccess.Camera(g_CameraProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                l_DriverAccessCamera.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                l_DriverAccessCamera.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Error creating driver using driver access toolkit: " & ex.ToString())
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : l_DriverAccessCamera.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DriverAccessCamera) : Catch : End Try
            l_DriverAccessCamera = Nothing
            GC.Collect()
        End Try

    End Sub
    Overrides Sub CreateDevice()
#If DEBUG Then
        LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Camera to get a Camera object")
        m_Camera = New ASCOM.DriverAccess.Camera(g_CameraProgID)
        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Camera to get a Camera object")
            m_Camera = New ASCOM.DriverAccess.Camera(g_CameraProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Camera object")
            m_Camera = CreateObject(g_CameraProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        End If
#End If
        g_Stop = False
    End Sub
    Overrides Property Connected() As Boolean
        Get
            Connected = m_Camera.Connected
        End Get
        Set(ByVal value As Boolean)
            m_Camera.Connected = value
            g_Stop = False
        End Set
    End Property
    Public Overrides Sub ReadCanProperties()
        ' ICameraV1 properties
        CameraCanTest(CanType.tstCanAbortExposure, "CanAbortExposure")
        CameraCanTest(CanType.tstCanAsymmetricBin, "CanAsymmetricBin")
        CameraCanTest(CanType.tstCanGetCoolerPower, "CanGetCoolerPower")
        CameraCanTest(CanType.tstCanPulseGuide, "CanPulseGuide")
        CameraCanTest(CanType.tstCanSetCCDTemperature, "CanSetCCDTemperature")
        CameraCanTest(CanType.tstCanStopExposure, "CanStopExposure")

        ' ICameraV2 properties
        CameraCanTest(CanType.tstCanFastReadout, "CanFastReadout")
    End Sub
    Private Sub CameraCanTest(ByVal p_Type As CanType, ByVal p_Name As String)
        Try
            Select Case p_Type
                Case CanType.tstCanAbortExposure
                    m_CanAbortExposure = m_Camera.CanAbortExposure
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanAbortExposure.ToString)
                Case CanType.tstCanAsymmetricBin
                    m_CanAsymmetricBin = m_Camera.CanAsymmetricBin
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanAsymmetricBin.ToString)
                Case CanType.tstCanGetCoolerPower
                    m_CanGetCoolerPower = m_Camera.CanGetCoolerPower
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanGetCoolerPower.ToString)
                Case CanType.tstCanPulseGuide
                    m_CanPulseGuide = m_Camera.CanPulseGuide
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanPulseGuide.ToString)
                Case CanType.tstCanSetCCDTemperature
                    m_CanSetCCDTemperature = m_Camera.CanSetCCDTemperature
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanSetCCDTemperature.ToString)
                Case CanType.tstCanStopExposure
                    m_CanStopExposure = m_Camera.CanStopExposure
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanStopExposure.ToString)
                Case CanType.tstCanFastReadout
                    m_CanFastReadout = m_Camera.CanFastReadout
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanFastReadout.ToString)
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "Conform:CanTest: Unknown test type " & p_Type.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "")
        End Try
    End Sub
    Public Overrides Sub PreRunCheck()
        Dim l_VStringPtr, l_V1, l_V2, l_V3 As Integer, l_VString As String
        'Add a test for a back level version of the Camera simulator - just abandon this process if any errors occur
        If g_CameraProgID.ToUpper = "CCDSIMULATOR.CAMERA" Then
            l_VString = ""
            Try
                l_VStringPtr = InStr(m_Camera.Description.ToUpper, "VERSION ") 'Point at the start of the version string
                If l_VStringPtr > 0 Then 'There is a version string so read what it is
                    l_VString = Mid(m_Camera.Description.ToUpper, l_VStringPtr + 8) 'Get the version string
                    l_VStringPtr = InStr(l_VString, ".")
                    If l_VStringPtr > 1 Then 'there is a first version number part
                        l_V1 = CInt(Mid(l_VString, 1, l_VStringPtr - 1)) 'Extract the number
                        l_VString = Mid(l_VString, l_VStringPtr + 1) 'Get the second version number part
                        l_VStringPtr = InStr(l_VString, ".")
                        If l_VStringPtr > 1 Then 'there is a second version number part
                            l_V2 = CInt(Mid(l_VString, 1, l_VStringPtr - 1)) 'Extract the number
                            l_VString = Mid(l_VString, l_VStringPtr + 1) 'Get the third version number part
                            l_V3 = CInt(l_VString) 'Extract the number
                            'Turn the version parts into a whole number
                            l_V1 = l_V1 * 1000000 + l_V2 * 1000 + l_V3
                            If l_V1 < 5000008 Then 'We are using a version earlier that 5.0.6 so issue a message
                                LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the camera simulator has known conformance issues, ***")
                                LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***")
                                LogMsg("", MessageLevel.msgAlways, "")
                            End If
                        End If
                    End If
                Else 'Must be an early version 5.0.1 or earlier
                    LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the camera simulator has known conformance issues, ***")
                    LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***")
                    LogMsg("", MessageLevel.msgAlways, "")
                End If
            Catch ex As Exception
                LogMsg("ConformanceCheck", MessageLevel.msgError, ex.ToString)
            End Try
        End If

        'Run camera tests
        If Not TestStop() Then
            LogMsg("", MessageLevel.msgAlways, "")
            'Check LastError throws an exception
            LogMsg("Last Tests", MessageLevel.msgAlways, "")
            Try
                m_LastExposureDuration = m_Camera.LastExposureDuration
                LogMsg("LastExposureDuration", MessageLevel.msgError, "LastExposureDuration did not generate an exception when called before an exposure was made")
            Catch ex As COMException
                LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration correctly generated a COM exception before an exposure was made")
            Catch ex As Exception
                LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration correctly generated a .NET exception before an exposure was made")
            End Try

            Try
                m_LastExposureStartTime = m_Camera.LastExposureStartTime
                LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime did not generate an exception when called before an exposure was made")
            Catch ex As COMException
                LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime correctly generated a COM exception before an exposure was made")
            Catch ex As Exception
                LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime correctly generated a .NET exception before an exposure was made")
            End Try
        End If
    End Sub

    Overrides Sub CheckProperties()
        Dim l_BinX, l_BinY, l_MaxBinX, l_MaxBinY As Integer

#Region "ICameraV1 Properties"

        'Basic read tests
        m_MaxBinX = CShort(CameraPropertyTestInteger(CamPropertyType.MaxBinX, "MaxBinX", 1, 10)) : If TestStop() Then Exit Sub
        m_MaxBinY = CShort(CameraPropertyTestInteger(CamPropertyType.MaxBinY, "MaxBinY", 1, 10)) : If TestStop() Then Exit Sub

        If Not m_CanAsymmetricBin Then ' Only symmetric binning is supported so confirm MaxBinX and Y match
            If m_MaxBinX <> m_MaxBinY Then LogMsg("CanAsymmetricBin", MessageLevel.msgError, "CanAsymmetricBin is false but MaxBinX and MaxBinY are not equal!")
        End If

        m_BinX = CShort(CameraPropertyTestInteger(CamPropertyType.BinX, "BinX Read", 1, 1)) : If TestStop() Then Exit Sub 'Must default to 1 on start-up
        m_BinY = CShort(CameraPropertyTestInteger(CamPropertyType.BinY, "BinY Read", 1, 1)) : If TestStop() Then Exit Sub 'Must default to 1 on start-up
        If Not m_CanAsymmetricBin Then ' Only symmetric binning is supported so confirm MaxBinX and Y match
            If m_BinX <> m_BinY Then LogMsg("CanAsymmetricBin", MessageLevel.msgError, "CanAsymmetricBin is false but BinX and BinY are not equal!")
        End If

        'Test writing low and high Bin values outside maximum range
        Try 'Invalid low value
            m_Camera.BinX = 0
            LogMsg("BinX Write", MessageLevel.msgError, "Invalid value 0 written but no exception generated")
        Catch ex As Exception
            LogMsg("BinX Write", MessageLevel.msgOK, "Exception correctly generated on setting BinX to 0")
        End Try
        Try 'Invalid high value
            m_Camera.BinX = CShort(m_MaxBinX + 1)
            LogMsg("BinX Write", MessageLevel.msgError, "Invalid value " & m_MaxBinX + 1 & " written but no exception generated")
        Catch ex As Exception
            LogMsg("BinX Write", MessageLevel.msgOK, "Exception correctly generated on setting BinX to " & m_MaxBinX + 1)
        End Try
        Try 'Invalid low value
            m_Camera.BinY = 0
            LogMsg("BinY Write", MessageLevel.msgError, "Invalid value 0 written but no exception generated")
        Catch ex As Exception
            LogMsg("BinY Write", MessageLevel.msgOK, "Exception correctly generated on setting BinY to 0")
        End Try
        Try 'Invalid high value
            m_Camera.BinY = CShort(m_MaxBinY + 1)
            LogMsg("BinY Write", MessageLevel.msgError, "Invalid value " & m_MaxBinY + 1 & " written but no exception generated")
        Catch ex As Exception
            LogMsg("BinY Write", MessageLevel.msgOK, "Exception correctly generated on setting BinY to " & m_MaxBinY + 1)
        End Try

        'Use either the maximum values supplied by the camera driver or the maximum values defined in Conform's setup, if these have been set
        If (g_Settings.CameraMaxBinX > 0) Or (g_Settings.CameraMaxBinY > 0) Then LogMsg("", MessageLevel.msgAlways, "") ' Insert a blank line if required
        If g_Settings.CameraMaxBinX > 0 Then ' Use Conform's settings
            l_MaxBinX = g_Settings.CameraMaxBinX
            LogMsg("BinXY Write", MessageLevel.msgComment, String.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", l_MaxBinX, m_MaxBinX))
        Else ' Use the camera's settings
            l_MaxBinX = m_MaxBinX
        End If
        If g_Settings.CameraMaxBinY > 0 Then ' Use Conform's settings
            l_MaxBinY = g_Settings.CameraMaxBinY
            LogMsg("BinXY Write", MessageLevel.msgComment, String.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", l_MaxBinY, m_MaxBinY))
        Else ' Use the camera's settings
            l_MaxBinY = m_MaxBinY
        End If

        If (g_Settings.CameraMaxBinX > m_MaxBinX) Or (g_Settings.CameraMaxBinY > m_MaxBinY) Then LogMsg("", MessageLevel.msgAlways, "") ' Insert a blank line if required
        If g_Settings.CameraMaxBinX > m_MaxBinX Then LogMsg("BinXY Write", MessageLevel.msgComment, String.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", l_MaxBinX, m_MaxBinX))
        If g_Settings.CameraMaxBinY > m_MaxBinY Then LogMsg("BinXY Write", MessageLevel.msgComment, String.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", l_MaxBinY, m_MaxBinY))

        If (g_Settings.CameraMaxBinX > 0) Or (g_Settings.CameraMaxBinY > 0) Then LogMsg("", MessageLevel.msgAlways, "") ' Insert a blank line if required

        'Write BinX and BinY
        If m_CanAsymmetricBin Then 'Test all combinations
            For l_BinY = 1 To l_MaxBinY
                For l_BinX = 1 To l_MaxBinX
                    Try
                        m_Camera.BinY = CShort(l_BinY)
                        Try
                            m_Camera.BinX = CShort(l_BinX)
                            LogMsg("BinXY Write", MessageLevel.msgOK, "Successfully set asymmetric xy binning: " & l_BinX & " x " & l_BinY)
                        Catch ex1 As Exception
                            HandleInvalidValueExceptionAsInfo("BinX Write", MemberType.Property, Required.Mandatory, ex1, "when setting BinX to " & l_BinX, "InvalidValueException received when setting BinX to " & l_BinX)
                        End Try
                    Catch ex As Exception
                        HandleInvalidValueExceptionAsInfo("BinY Write", MemberType.Property, Required.Mandatory, ex, "when setting BinY to " & l_BinY, "InvalidValueException received when setting BinY to " & l_BinY)
                    End Try
                Next
            Next
        Else 'Test same bin combinations only
            For l_BinX = 1 To l_MaxBinX
                Try
                    m_Camera.BinY = CShort(l_BinX)
                    Try
                        m_Camera.BinX = CShort(l_BinX)
                        LogMsg("BinXY Write", MessageLevel.msgOK, "Successfully set symmetric xy binning: " & l_BinX & " x " & l_BinX)
                    Catch ex1 As Exception
                        HandleInvalidValueExceptionAsInfo("BinX Write", MemberType.Property, Required.Mandatory, ex1, "when setting BinX to " & l_BinX, "InvalidValueException received when setting BinX to " & l_BinX)
                    End Try
                Catch ex As Exception
                    HandleInvalidValueExceptionAsInfo("BinY Write", MemberType.Property, Required.Mandatory, ex, "when setting BinY to " & l_BinY, "InvalidValueException received when setting BinY to " & l_BinY)
                End Try
            Next

        End If

        'Reset binning to 1x1 state
        Try : m_Camera.BinX = 1 : Catch ex As Exception : End Try
        Try : m_Camera.BinY = 1 : Catch ex As Exception : End Try

        m_CameraState = CameraPropertyTestCameraState(CamPropertyType.CameraState, "CameraState") : If TestStop() Then Exit Sub
        m_CameraXSize = CameraPropertyTestInteger(CamPropertyType.CameraXSize, "CameraXSize", 1, Integer.MaxValue) : If TestStop() Then Exit Sub
        m_CameraYSize = CameraPropertyTestInteger(CamPropertyType.CameraYSize, "CameraYSize", 1, Integer.MaxValue) : If TestStop() Then Exit Sub

        m_CCDTemperature = CameraPropertyTestDouble(CamPropertyType.CCDTemperature, "CCDTemperature", -273.0, 100.0, False) : If TestStop() Then Exit Sub
        m_CoolerOn = CameraPropertyTestBoolean(CamPropertyType.CoolerOn, "CoolerOn Read", False) : If TestStop() Then Exit Sub

        'Write CoolerOn
        Dim l_OriginalCoolerState As Boolean
        Dim l_TargetCoolerState As String
        Try
            l_OriginalCoolerState = m_Camera.CoolerOn
            If l_OriginalCoolerState Then 'Set up error message element in case needed
                l_TargetCoolerState = "off"
            Else
                l_TargetCoolerState = "on"
            End If
            Try
                If l_OriginalCoolerState Then 'Cooler on so switch off
                    m_Camera.CoolerOn = False
                Else 'Cooler off so switch on
                    m_Camera.CoolerOn = True
                End If
                LogMsg("CoolerOn Write", MessageLevel.msgOK, "Successfully changed CoolerOn state")
            Catch ex As Exception
                HandleException("CoolerOn Write", MemberType.Property, Required.Optional, ex, "turning Cooler " & l_TargetCoolerState)
            End Try
            'Restore Cooler state
            Try : m_Camera.CoolerOn = l_OriginalCoolerState : Catch : End Try

        Catch ex As Exception
            'Skip tests as we can't even read the cooler state
            HandleException("CoolerOn Read", MemberType.Property, Required.Optional, ex, "")
        End Try

        m_CoolerPower = CameraPropertyTestDouble(CamPropertyType.CoolerPower, "CoolerPower", 0.0, 100.0, False) : If TestStop() Then Exit Sub
        m_ElectronsPerADU = CameraPropertyTestDouble(CamPropertyType.ElectronsPerADU, "ElectronsPerADU", 0.00001, Double.PositiveInfinity, False) : If TestStop() Then Exit Sub
        m_FullWellCapacity = CameraPropertyTestDouble(CamPropertyType.FullWellCapacity, "FullWellCapacity", 0.0, Double.PositiveInfinity, False) : If TestStop() Then Exit Sub
        m_HasShutter = CameraPropertyTestBoolean(CamPropertyType.HasShutter, "HasShutter", False) : If TestStop() Then Exit Sub
        m_HeatSinkTemperature = CameraPropertyTestDouble(CamPropertyType.HeatSinkTemperature, "HeatSinkTemperature", -273.0, 100, False) : If TestStop() Then Exit Sub

        m_ImageReady = CameraPropertyTestBoolean(CamPropertyType.ImageReady, "ImageReady", False) : If TestStop() Then Exit Sub
        If m_ImageReady Then LogMsg("ImageReady", MessageLevel.msgError, "Image is flagged as ready but no exposure has been started!")
        If m_ImageReady Then 'ImageReady is incorrectly flagged so don't know what to expect!
            Try
                m_ImageArray = CType(m_Camera.ImageArray, Integer(,))
                LogMsg("ImageArray", MessageLevel.msgError, "No image has been taken but ImageArray has not generated an exception")
            Catch ex As Exception
                LogMsg("ImageArray", MessageLevel.msgOK, "Exception correctly generated before an image has been taken")
            End Try
        Else 'ImageReady is false so should throw an exception
            Try
                m_ImageArray = CType(m_Camera.ImageArray, Integer(,))
                LogMsg("ImageArray", MessageLevel.msgError, "ImageReady is false and no image has been taken but ImageArray has not generated an exception")
            Catch ex As Exception
                LogMsg("ImageArray", MessageLevel.msgOK, "Exception correctly generated when ImageReady is false")
            End Try
        End If

        m_ImageArray = Nothing
        m_ImageArrayVariant = Nothing
        GC.Collect()

        If m_ImageReady Then 'ImageReady is incorrectly flagged so don't know what to expect!
            Try
                Dim ImageArrayVariantObject As Object
                ImageArrayVariantObject = m_Camera.ImageArrayVariant
                m_ImageArrayVariant = CType(ImageArrayVariantObject, Integer(,))
                LogMsg("ImageArrayVariant", MessageLevel.msgError, "No image has been taken but ImageArray has not generated an exception")
            Catch ex As Exception
                LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Exception correctly generated before an image has been taken")
            End Try
        Else 'ImageReady is false so should throw an exception
            Try
                m_ImageArrayVariant = CType(m_Camera.ImageArrayVariant, Integer(,))
                LogMsg("ImageArrayVariant", MessageLevel.msgError, "ImageReady is false and no image has been taken but ImageArray has not generated an exception")
            Catch ex As Exception
                LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Exception correctly generated when ImageReady is false")
            End Try
        End If
        Try : Marshal.ReleaseComObject(m_ImageArray) : Catch : End Try
        Try : Marshal.ReleaseComObject(m_ImageArrayVariant) : Catch : End Try
        m_ImageArray = Nothing
        m_ImageArrayVariant = Nothing
        GC.Collect()

        m_IsPulseGuiding = CameraPropertyTestBoolean(CamPropertyType.IsPulseGuiding, "IsPulseGuiding", False) : If TestStop() Then Exit Sub
        If m_IsPulseGuiding Then LogMsg("IsPulseGuiding", MessageLevel.msgError, "Camera is showing pulse guiding underway although no PulseGuide command has been issued!")

        m_MaxADU = CameraPropertyTestInteger(CamPropertyType.MaxADU, "MaxADU", 1, Integer.MaxValue) : If TestStop() Then Exit Sub

        m_NumX = CameraPropertyTestInteger(CamPropertyType.NumX, "NumX Read", 1, m_CameraXSize) : If TestStop() Then Exit Sub
        CameraPropertyWriteTest(CamPropertyType.NumX, "NumX", CInt(m_CameraXSize / 2))

        m_NumY = CameraPropertyTestInteger(CamPropertyType.NumY, "NumY Read", 1, m_CameraYSize) : If TestStop() Then Exit Sub
        CameraPropertyWriteTest(CamPropertyType.NumY, "NumY", CInt(m_CameraYSize / 2))

        m_PixelSizeX = CameraPropertyTestDouble(CamPropertyType.PixelSizeX, "PixelSizeX", 1.0, Double.PositiveInfinity, False) : If TestStop() Then Exit Sub
        m_PixelSizeY = CameraPropertyTestDouble(CamPropertyType.PixelSizeY, "PixelSizeY", 1.0, Double.PositiveInfinity, False) : If TestStop() Then Exit Sub

        m_SetCCDTemperature = CameraPropertyTestDouble(CamPropertyType.SetCCDTemperature, "SetCCDTemperature Read", -273.0, 100.0, False) : If TestStop() Then Exit Sub
        If m_CanSetCCDTemperature Then 'Should work OK
            Try
                m_Camera.SetCCDTemperature = 0 'Try an acceptable value
                LogMsg("SetCCDTemperature Write", MessageLevel.msgOK, "Successfully wrote 0")
                'Find lowest acceptable temperature
                Dim l_OK As Boolean
                Dim l_SetPoint As Double

                'Find low setpoint at which an exception is generated, stop at -280 as this is unphysical
                l_OK = False
                l_SetPoint = 0.0
                Do
                    l_SetPoint = l_SetPoint - 5
                    Try
                        m_Camera.SetCCDTemperature = l_SetPoint
                    Catch ex As Exception
                        l_OK = True 'Reached lower limit so exit loop
                    End Try
                Loop Until l_OK Or (l_SetPoint = -280.0)
                If l_SetPoint = -280.0 Then ' No lower limit so generate an issue
                    LogMsg("SetCCDTemperature Write", MessageLevel.msgIssue, "Setpoint could be set below absolute zero!")
                Else 'Have found a limit so log what it is
                    LogMsg("SetCCDTemperature Write", MessageLevel.msgInfo, "Setpoint lower limit (to nearest 5 degrees) found: " & l_SetPoint + 5.0)
                End If

                'Find high setpoint at which an exception is generated, stop at 100 as this is a rather high value
                l_OK = False
                l_SetPoint = 0.0
                Do
                    l_SetPoint = l_SetPoint + 5
                    Try
                        m_Camera.SetCCDTemperature = l_SetPoint
                    Catch ex As Exception
                        l_OK = True 'Reached lower limit so exit loop
                    End Try
                Loop Until l_OK Or (l_SetPoint = 100.0)
                If l_SetPoint = 100.0 Then ' No higher limit so generate an issue
                    LogMsg("SetCCDTemperature Write", MessageLevel.msgIssue, "Setpoint could be set to 100 degrees!")
                Else 'Have found a limit so log what it is
                    LogMsg("SetCCDTemperature Write", MessageLevel.msgInfo, "Setpoint high limit (to nearest 5 degrees) found: " & l_SetPoint - 5.0)
                End If
            Catch ex As Exception
                HandleException("SetCCDTemperature Write", MemberType.Property, Required.MustBeImplemented, ex, "Setting a legal value 0")
            End Try
            'Restore original value
            Try : m_Camera.SetCCDTemperature = m_SetCCDTemperature : Catch : End Try
        Else 'CanSetCCDTemperature is false, should generate an exception
            Try
                m_Camera.SetCCDTemperature = 0
                LogMsg("SetCCDTemperature Write", MessageLevel.msgError, "CanSetCCDTemperature is false but no exception generated on write")
            Catch ex As Exception
                HandleException("SetCCDTemperature Write", MemberType.Property, Required.Optional, ex, "")
            End Try
        End If

        m_StartX = CameraPropertyTestInteger(CamPropertyType.StartX, "StartX Read", 0, m_CameraXSize - 1) : If TestStop() Then Exit Sub
        CameraPropertyWriteTest(CamPropertyType.StartX, "StartX", CInt(m_CameraXSize / 2))
        m_StartY = CameraPropertyTestInteger(CamPropertyType.StartY, "StartY Read", 0, m_CameraYSize - 1) : If TestStop() Then Exit Sub
        CameraPropertyWriteTest(CamPropertyType.StartY, "StartY", CInt(m_CameraYSize / 2))

#End Region

#Region "ICameraV2 Properties"

        If m_Camera.InterfaceVersion > 1 Then ' Only for ICameraV2 and later
            ' SensorType - Mandatory
            ' This must be tested before BayerOffset because BayerOffset is mandatory for colour and optional for monochrome cameras
            Try
                m_SensorType = m_Camera.SensorType
                m_CanReadSensorType = True ' Set a flag to indicate that we have got a valid SensorType value
                'Successfully retrieved a value
                LogMsg("SensorType Read", MessageLevel.msgOK, m_SensorType.ToString)
            Catch ex As Exception
                HandleException("SensorType Read", MemberType.Property, Required.Mandatory, ex, "")
            End Try

            ' BayerOffset Read
            If m_CanReadSensorType Then ' SensorType value is available
                If m_SensorType = SensorType.Monochrome Then ' Monochrome chip
                    ' Monochrome so both BayerOffset properties should throw not implemented exceptions
                    CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetX, "BayerOffsetX Read")
                    CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetY, "BayerOffsetY Read")
                Else ' Bayer colour chip
                    m_BayerOffsetX = CameraPropertyTestShort(CamPropertyType.BayerOffsetX, "BayerOffsetX Read", 0, 10000, True)
                    m_BayerOffsetY = CameraPropertyTestShort(CamPropertyType.BayerOffsetY, "BayerOffsetY Read", 0, 10000, True)
                End If
            Else ' Unable to read SensorType so assume that the sensor is Monochrome
                LogMsg("BayerOffset Read", MessageLevel.msgInfo, "Unable to read SensorType value; assuming that the sensor is Monochrome")
                ' Monochrome so both BayerOffset properties should throw not implemented exceptions
                CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetX, "BayerOffsetX Read")
                CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetY, "BayerOffsetY Read")
            End If

            ' ExposureMin and ExpoosureMax Read
            m_ExposureMax = CameraPropertyTestDouble(CamPropertyType.ExposureMax, "ExposureMax Read", 0.0001, Double.MaxValue, True)
            m_ExposureMin = CameraPropertyTestDouble(CamPropertyType.ExposureMin, "ExposureMin Read", 0.0, Double.MaxValue, True)
            If m_ExposureMin <= m_ExposureMax Then
                LogMsg("ExposureMin", MessageLevel.msgOK, "ExposureMin is less than or equal to ExposureMax")
            Else
                LogMsg("ExposureMin", MessageLevel.msgError, "ExposureMin is greater than ExposureMax")
            End If

            ' ExposureResolution Read
            m_ExposureResolution = CameraPropertyTestDouble(CamPropertyType.ExposureResolution, "ExposureResolution Read", 0.0, Double.MaxValue, True)
            If m_ExposureResolution <= m_ExposureMax Then
                LogMsg("ExposureResolution", MessageLevel.msgOK, "ExposureResolution is less than or equal to ExposureMax")
            Else
                LogMsg("ExposureResolution", MessageLevel.msgError, "ExposureResolution is greater than ExposureMax")
            End If

            'FastReadout Read Optional
            If m_CanFastReadout Then ' Should be implemented and not throw a not implemented exception
                m_FastReadout = CameraPropertyTestBoolean(CamPropertyType.FastReadout, "FastReadout Read", True)
            Else ' Should throw a not implemented exception
                Try
                    m_FastReadout = m_Camera.FastReadout
                    LogMsg("FastReadout Read", MessageLevel.msgError, "CanFastReadout is False but a PropertyNotImplementedException was not thrown")
                Catch ex As Exception
                    HandleException("FastReadout Read", MemberType.Property, Required.Optional, ex, "")
                End Try
            End If

            ' FastReadout Write Optional
            If m_CanFastReadout Then ' Should be implemented and not throw a not implemented exception
                Try
                    m_Camera.FastReadout = Not m_FastReadout
                    m_Camera.FastReadout = m_FastReadout
                    LogMsg("FastReadout Write", MessageLevel.msgOK, "Able to change the FastReadout state OK")
                Catch ex As Exception
                    HandleException("FastReadout Write", MemberType.Property, Required.Mandatory, ex, "")
                End Try

            Else ' Should throw a not implemented exception
                Try
                    m_Camera.FastReadout = True
                    LogMsg("FastReadout Write", MessageLevel.msgError, "CanFastReadout is False but a PropertyNotImplementedException was not thrown")
                Catch ex As Exception
                    HandleException("FastReadout Write", MemberType.Property, Required.Optional, ex, "")
                End Try
            End If

            ' GainMin Read - Optional
            Try
                m_CanReadGainMin = False
                m_GainMin = m_Camera.GainMin
                'Successfully retrieved a value
                m_CanReadGainMin = True
                LogMsg("GainMin Read", MessageLevel.msgOK, m_GainMin.ToString)
            Catch ex As Exception
                HandleInvalidOperationExceptionAsOK("GainMin Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown")
            End Try

            ' GainMax Read - Optional
            Try
                m_CanReadGainMax = False
                m_GainMax = m_Camera.GainMax
                'Successfully retrieved a value
                m_CanReadGainMax = True
                LogMsg("GainMax Read", MessageLevel.msgOK, m_GainMax.ToString)
            Catch ex As Exception
                HandleInvalidOperationExceptionAsOK("GainMax Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown")
            End Try

            ' Gains Read - Optional
            Try
                m_CanReadGains = False
                m_Gains = m_Camera.Gains
                'Successfully retrieved a value
                m_CanReadGains = True
                For Each Gain As String In m_Gains
                    LogMsg("Gains Read", MessageLevel.msgOK, Gain.ToString)
                Next
            Catch ex As Exception
                HandleInvalidOperationExceptionAsOK("Gains Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown")
            End Try

            If m_CanReadGainMax And m_CanReadGainMin And m_CanReadGains Then ' Both GainMin/Max and Gains are enabled but only one mechanic should be used
                LogMsg("Gains", MessageLevel.msgError, "GainMin, GainMax and Gains are all readable. Only one of GainMin/Max as a pair or Gains should be used, the other should throw a PropertyNotImplementedException")
            Else ' Only one mechanic is active or no mechanic is active so no action

            End If

            ' Gain Read - Optional 
            Try
                m_CanReadGain = False ' Set default value to indicate can't read gain
                m_Gain = m_Camera.Gain
                m_CanReadGain = True ' Flag that we can read Gain OK
                If m_CanReadGains Then
                    LogMsg("Gain Read", MessageLevel.msgOK, m_Gain & " " & m_Gains(0).ToString)
                Else
                    LogMsg("Gain Read", MessageLevel.msgOK, m_Gain.ToString)
                End If

            Catch ex As Exception
                HandleException("Gain Read", MemberType.Property, Required.Optional, ex, "")
            End Try

            ' Now check that gain property groups are implemented to handle the three gain modes: NotImplemented, Gain Index (Gain + Gains) and Gain Value (Gain + GainMin + GainMax)
            If Not m_CanReadGain And Not m_CanReadGains And Not m_CanReadGainMin And Not m_CanReadGainMax Then ' We are in Not Implemented mode so all is OK
                LogMsgOK("Gain Read", "All four gain properties throw exceptions - the driver is in ""Gain Not Implemented"" mode.")
            Else ' Test for Gain Index and Gain Value modes
                If m_CanReadGain Then ' Can read Gain so the driver could be in either Gain Index or Gain Value mode
                    ' Test for Gain Index mode
                    If (m_CanReadGain And m_CanReadGains And Not m_CanReadGainMin And Not m_CanReadGainMax) Then ' We are in Gain Index mode so all is OK
                        LogMsgOK("Gain Read", "Gain and Gains can be read while GainMin and GainMax throw exceptions - the driver is in ""Gain Index"" mode.")
                    Else ' Test for Gain Value mode
                        If (m_CanReadGain And Not m_CanReadGains And m_CanReadGainMin And m_CanReadGainMax) Then ' We are in Gain Value mode so all is OK
                            LogMsgOK("Gain Read", "Gain, GainMin and GainMax can be read OK while Gains throws an exception - the driver is in ""Gain Value"" mode.")
                        Else ' Bad combination of properties - this is not a valid mode
                            LogMsgError("Gain Read", $"Unable to determine whether the driver is in ""Gain Not Implemented"", ""Gain Index"" or ""Gain Value"" mode. Please check the interface specification.")
                            LogMsgInfo("Gain Read", $"Gain threw an exception: {m_CanReadGain}, Gains threw an exception: {m_CanReadGains}, GainMin threw an exception: {m_CanReadGainMin}, GainMax threw an exception: {m_CanReadGainMax}.")
                            LogMsgInfo("Gain Read", $"""Gain Not Implemented"" mode: Gain, Gains, GainMin and GainMax must all throw exceptions.")
                            LogMsgInfo("Gain Read", $"""Gain Index"" mode: Gain and Gains must work while GainMin and GainMax must throw exceptions.")
                            LogMsgInfo("Gain Read", $"""Gain Value"" mode: Gain, GainMin and GainMax must work while Gains must throw an exception.")
                        End If
                    End If
                Else ' Can not read gain but can read at least one of the other gain properties, this is a misconfiguration
                    LogMsgError("Gain Read", $"Gain Read threw an exception but at least one of Gains, GainMin Or GainMax did not throw an exception. If Gain throws an exception, all the other gain properties should do likewise.")
                    LogMsgInfo("Gain Read", $"Gains threw an exception: {m_CanReadGains}, GainMin threw an exception: {m_CanReadGainMin}, GainMax threw an exception: {m_CanReadGainMax}.")
                End If
            End If

            ' PercentCompleted Read - Optional - corrected to match the specification
            Try
                m_PercentCompleted = m_Camera.PercentCompleted
                Select Case m_PercentCompleted
                    Case Is < 0 'Lower than minimum value
                        LogMsg("PercentCompleted Read", MessageLevel.msgError, "Invalid value: " & m_PercentCompleted.ToString)
                    Case Is > 100 'Higher than maximum value
                        LogMsg("PercentCompleted Read", MessageLevel.msgError, "Invalid value: " & m_PercentCompleted.ToString)
                    Case Else 'OK value
                        LogMsg("PercentCompleted Read", MessageLevel.msgOK, m_PercentCompleted.ToString)
                End Select
            Catch ex As Exception
                HandleInvalidValueExceptionAsOK("PercentCompleted Read", MemberType.Property, Required.Optional, ex, "", "Operation is invalid when camera is not imaging or downloading")
            End Try

            ' ReadoutModes - Mandatory
            Try
                m_CanReadReadoutModes = False
                m_ReadoutModes = m_Camera.ReadoutModes
                'Successfully retrieved a value
                m_CanReadReadoutModes = True
                For Each ReadoutMode As String In m_ReadoutModes
                    LogMsg("ReadoutModes Read", MessageLevel.msgOK, ReadoutMode.ToString)
                Next
            Catch ex As Exception
                HandleException("Gain Read", MemberType.Property, Required.Mandatory, ex, "")
            End Try

            ' ReadoutMode Read - Mandatory
            m_ReadoutMode = CameraPropertyTestShort(CamPropertyType.ReadoutMode, "ReadoutMode Read", 0, Short.MaxValue, True)
            If m_CanReadReadoutModes Then
                Try
                    If m_ReadoutMode < m_ReadoutModes.Count Then
                        LogMsg("ReadoutMode Index", MessageLevel.msgOK, "ReadReadoutMode is within the bounds of the ReadoutModes ArrayList")
                        LogMsg("ReadoutMode Index", MessageLevel.msgInfo, "Current value: " & m_ReadoutModes(m_ReadoutMode).ToString)
                    Else
                        LogMsg("ReadoutMode Index", MessageLevel.msgError, "ReadReadoutMode is outside the bounds of the ReadoutModes ArrayList")
                    End If
                Catch ex As Exception
                    LogMsg("ReadoutMode Index", MessageLevel.msgError, "Exception: " & ex.Message)
                    LogMsg("ReadoutMode Index", MessageLevel.msgDebug, ex.ToString)
                End Try
            Else
                LogMsg("ReadoutMode Index", MessageLevel.msgInfo, "Skipping ReadReadoutMode index test because ReadoutModes is unavailable")
            End If

            ' SensorName
            m_SensorName = CameraPropertyTestString(CamPropertyType.SensorName, "SensorName Read", 250, True)

        End If

#End Region

#Region "ICameraV3 Properties"

        If m_Camera.InterfaceVersion > 2 Then ' Only for ICameraV3 and later
            ' OffsetMin Read - Optional
            Try
                m_CanReadOffsetMin = False
                m_OffsetMin = m_Camera.OffsetMin
                'Successfully retrieved a value
                m_CanReadOffsetMin = True
                LogMsg("OffsetMin Read", MessageLevel.msgOK, m_OffsetMin.ToString)
            Catch ex As Exception
                HandleInvalidOperationExceptionAsOK("OffsetMin Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown")
            End Try

            ' OffsetMax Read - Optional
            Try
                m_CanReadOffsetMax = False
                m_OffsetMax = m_Camera.OffsetMax
                'Successfully retrieved a value
                m_CanReadOffsetMax = True
                LogMsg("OffsetMax Read", MessageLevel.msgOK, m_OffsetMax.ToString)
            Catch ex As Exception
                HandleInvalidOperationExceptionAsOK("OffsetMax Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown")
            End Try

            ' Offsets Read - Optional
            Try
                m_CanReadOffsets = False
                m_Offsets = m_Camera.Offsets
                'Successfully retrieved a value
                m_CanReadOffsets = True
                For Each Offset As String In m_Offsets
                    LogMsg("Offsets Read", MessageLevel.msgOK, Offset.ToString)
                Next
            Catch ex As Exception
                HandleInvalidOperationExceptionAsOK("Offsets Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown")
            End Try

            If m_CanReadOffsetMax And m_CanReadOffsetMin And m_CanReadOffsets Then ' Both OffsetMin/Max and Offsets are enabled but only one mechanic should be used
                LogMsg("Offsets", MessageLevel.msgError, "OffsetMin, OffsetMax and Offsets are all readable. Only one of OffsetMin/Max as a pair or Offsets should be used, the other should throw a PropertyNotImplementedException")
            Else ' Only one mechanic is active or no mechanic is active so no action

            End If

            ' Offset Read - Optional 
            Try
                m_CanReadOffset = False ' Set default value to indicate can't read offset
                m_Offset = m_Camera.Offset
                m_CanReadOffset = True ' Flag that we can read Offset OK
                If m_CanReadOffsets Then
                    LogMsg("Offset Read", MessageLevel.msgOK, m_Offset & " " & m_Offsets(0).ToString)
                Else
                    LogMsg("Offset Read", MessageLevel.msgOK, m_Offset.ToString)
                End If

            Catch ex As Exception
                HandleException("Offset Read", MemberType.Property, Required.Optional, ex, "")
            End Try

            ' Now check that offset property groups are implemented to handle the three offset modes: NotImplemented, Offset Index (Offset + Offsets) and Offset Value (Offset + OffsetMin + OffsetMax)
            If Not m_CanReadOffset And Not m_CanReadOffsets And Not m_CanReadOffsetMin And Not m_CanReadOffsetMax Then ' We are in Not Implemented mode so all is OK
                LogMsgOK("Offset Read", "All four offset properties throw exceptions - the driver is in ""Offset Not Implemented"" mode.")
            Else ' Test for Offset Index and Offset Value modes
                If m_CanReadOffset Then ' Can read Offset so the driver could be in either Offset Index or Offset Value mode
                    ' Test for Offset Index mode
                    If (m_CanReadOffset And m_CanReadOffsets And Not m_CanReadOffsetMin And Not m_CanReadOffsetMax) Then ' We are in Offset Index mode so all is OK
                        LogMsgOK("Offset Read", "Offset and Offsets can be read while OffsetMin and OffsetMax throw exceptions - the driver is in ""Offset Index"" mode.")
                    Else ' Test for Offset Value mode
                        If (m_CanReadOffset And Not m_CanReadOffsets And m_CanReadOffsetMin And m_CanReadOffsetMax) Then ' We are in Offset Value mode so all is OK
                            LogMsgOK("Offset Read", "Offset, OffsetMin and OffsetMax can be read OK while Offsets throws an exception - the driver is in ""Offset Value"" mode.")
                        Else ' Bad combination of properties - this is not a valid mode
                            LogMsgError("Offset Read", $"Unable to determine whether the driver is in ""Offset Not Implemented"", ""Offset Index"" or ""Offset Value"" mode. Please check the interface specification.")
                            LogMsgInfo("Offset Read", $"Offset threw an exception: {m_CanReadOffset}, Offsets threw an exception: {m_CanReadOffsets}, OffsetMin threw an exception: {m_CanReadOffsetMin}, OffsetMax threw an exception: {m_CanReadOffsetMax}.")
                            LogMsgInfo("Offset Read", $"""Offset Not Implemented"" mode: Offset, Offsets, OffsetMin and OffsetMax must all throw exceptions.")
                            LogMsgInfo("Offset Read", $"""Offset Index"" mode: Offset and Offsets must work while OffsetMin and OffsetMax must throw exceptions.")
                            LogMsgInfo("Offset Read", $"""Offset Value"" mode: Offset, OffsetMin and OffsetMax must work while Offsets must throw an exception.")
                        End If
                    End If
                Else ' Can not read offset but can read at least one of the other offset properties, this is a misconfiguration
                    LogMsgError("Offset Read", $"Offset Read threw an exception but at least one of Offsets, OffsetMin Or OffsetMax did not throw an exception. If Offset throws an exception, all the other offset properties should do likewise.")
                    LogMsgInfo("Offset Read", $"Offsets threw an exception: {m_CanReadOffsets}, OffsetMin threw an exception: {m_CanReadOffsetMin}, OffsetMax threw an exception: {m_CanReadOffsetMax}.")
                End If
            End If

            ' SubExposureDuration Read - Optional 
            m_SubExposureDuration = CameraPropertyTestDouble(CamPropertyType.SubExposureDuration, "SubExposureDuration", Double.Epsilon, Double.MaxValue, False) : If TestStop() Then Exit Sub

            ' SubExposureDuration Write - Optional 
            Try
                m_Camera.SubExposureDuration = m_SubExposureDuration
                LogMsg("SubExposureDuration write", MessageLevel.msgOK, $"Successfully wrote {m_SubExposureDuration}")
            Catch ex As Exception
                HandleException("SubExposureDuration write", MemberType.Property, Required.Optional, ex, "")
            End Try

            ' PulseGuideStatus - Mandatory
            'Try
            '    m_CanReadPulseGuideStatus = False ' Initialise to failure
            '    m_PulseGuideStatus = m_Camera.PulseGuideStatus
            '    LogMsg("PulseGuideStatus", MessageLevel.msgOK, m_PulseGuideStatus.ToString())
            '    m_CanReadPulseGuideStatus = True ' Record a successful PulseGuideStatus read

            '    Select Case m_PulseGuideStatus
            '        Case PulseGuideState.NotImplemented
            '            If Not m_CanPulseGuide Then
            '                LogMsgOK("PulseGuideStatus", "CanPulseGuide is False and PulseGuideStatus returned ""NotImplemented""")
            '            Else
            '                LogMsgError("PulseGuideStatus", "CanPulseGuide is True but PulseGuideStatus returned ""NotImplemented""")
            '            End If

            '        Case PulseGuideState.Inactive
            '            If m_CanPulseGuide Then ' Can pulse guide
            '                If Not m_IsPulseGuiding Then ' Can pulse guide but not pulse guiding right now
            '                    LogMsgOK("PulseGuideStatus", "CanPulseGuide is True and PulseGuideStatus returned ""Inactive""")
            '                Else ' IsPulseGuiding shows a guide as in progress even though it has not been initiated by Conform by PulseGuideStatus returns Inactive - The two methods are giving different answers
            '                    LogMsgError("PulseGuideStatus", $"IsPulseGuiding is True but PulseGuideStatus returned ""Inactive""")
            '                End If
            '            Else ' Can not pulse guide but state is reported as inactive
            '                LogMsgError("PulseGuideStatus", $"CanPulseGuide is False but PulseGuideStatus returned ""Inactive"" rather than ""NotImplemented""")
            '            End If

            '        Case PulseGuideState.Active
            '            LogMsgError("PulseGuideStatus", $"Conform has not initiated a PulseGuide but PulseGuideStatus returned ""Active"".")

            '        Case PulseGuideState.Error
            '            LogMsgError("PulseGuideStatus", $"PulseGuideStatus returned an ""Error"" status.")

            '    End Select

            'Catch ex As Exception
            '    HandleException("PulseGuideStatus", MemberType.Property, Required.Mandatory, ex, "")
            'End Try

        End If
#End Region

    End Sub
    Private Function CameraPropertyTestCameraState(ByVal p_Type As CamPropertyType, ByVal p_Name As String) As CameraStates
        Try
            CameraPropertyTestCameraState = CameraStates.cameraIdle
            Select Case p_Type
                Case CamPropertyType.CameraState
                    CameraPropertyTestCameraState = m_Camera.CameraState
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestCameraState: Unknown test type - " & p_Type.ToString)
            End Select
            LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestCameraState.ToString)
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Optional, ex, "")
        End Try
    End Function
    Private Function CameraPropertyTestShort(ByVal p_Type As CamPropertyType, ByVal p_Name As String, ByVal p_Min As Short, ByVal p_Max As Short, p_Mandatory As Boolean) As Short
        Try
            CameraPropertyTestShort = 0
            Select Case p_Type
                Case CamPropertyType.BayerOffsetX
                    CameraPropertyTestShort = m_Camera.BayerOffsetX
                Case CamPropertyType.BayerOffsetY
                    CameraPropertyTestShort = m_Camera.BayerOffsetY
                Case CamPropertyType.PercentCompleted
                    CameraPropertyTestShort = m_Camera.PercentCompleted
                Case CamPropertyType.ReadoutMode
                    CameraPropertyTestShort = m_Camera.ReadoutMode
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestShort: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case CameraPropertyTestShort
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & CameraPropertyTestShort.ToString)
                Case Is > p_Max 'Higher than maximum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & CameraPropertyTestShort.ToString)
                Case Else 'OK value
                    LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestShort.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    Private Function CameraPropertyMustNotImplemented(ByVal p_Type As CamPropertyType, ByVal p_Name As String) As Boolean
        Dim TestShort As Short ' Dummy variable to hold value that should never be returned

        CameraPropertyMustNotImplemented = True
        Try
            Select Case p_Type
                Case CamPropertyType.BayerOffsetX
                    TestShort = m_Camera.BayerOffsetX
                    CameraPropertyMustNotImplemented = False ' Property should throw an exception but did not so record that fact
                    LogMsg(p_Name, MessageLevel.msgError, "Sensor type is Monochrome so this property must throw a PropertyNotImplementedException; it must not return a value")
                Case CamPropertyType.BayerOffsetY
                    TestShort = m_Camera.BayerOffsetY
                    CameraPropertyMustNotImplemented = False ' Property should throw an exception but did not so record that fact
                    LogMsg(p_Name, MessageLevel.msgError, "Sensor type is Monochrome so this property must throw a PropertyNotImplementedException; it must not return a value")
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyMustNotImplemented: Unknown test type - " & p_Type.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.MustNotBeImplemented, ex, "Sensor type is Monochrome")
        End Try
        Return CameraPropertyMustNotImplemented ' Return success indicator, True means property did thrown the exception, False means that it did not
    End Function
    Private Function CameraPropertyTestInteger(ByVal p_Type As CamPropertyType, ByVal p_Name As String, ByVal p_Min As Integer, ByVal p_Max As Integer) As Integer
        Try
            CameraPropertyTestInteger = 0
            Select Case p_Type
                Case CamPropertyType.BinX
                    CameraPropertyTestInteger = m_Camera.BinX
                Case CamPropertyType.BinY
                    CameraPropertyTestInteger = m_Camera.BinY
                Case CamPropertyType.CameraState
                    CameraPropertyTestInteger = m_Camera.CameraState
                Case CamPropertyType.CameraXSize
                    CameraPropertyTestInteger = m_Camera.CameraXSize
                Case CamPropertyType.CameraYSize
                    CameraPropertyTestInteger = m_Camera.CameraYSize
                Case CamPropertyType.CameraYSize
                    CameraPropertyTestInteger = m_Camera.CameraYSize
                Case CamPropertyType.MaxADU
                    CameraPropertyTestInteger = m_Camera.MaxADU
                Case CamPropertyType.MaxBinX
                    CameraPropertyTestInteger = m_Camera.MaxBinX
                Case CamPropertyType.MaxBinY
                    CameraPropertyTestInteger = m_Camera.MaxBinY
                Case CamPropertyType.NumX
                    CameraPropertyTestInteger = m_Camera.NumX
                Case CamPropertyType.NumY
                    CameraPropertyTestInteger = m_Camera.NumY
                Case CamPropertyType.StartX
                    CameraPropertyTestInteger = m_Camera.StartX
                Case CamPropertyType.StartY
                    CameraPropertyTestInteger = m_Camera.StartY
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestInteger: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case CameraPropertyTestInteger
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & CameraPropertyTestInteger.ToString)
                Case Is > p_Max 'Higher than maximum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & CameraPropertyTestInteger.ToString)
                Case Else 'OK value
                    LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestInteger.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Optional, ex, "")
        End Try
    End Function
    Private Function CameraPropertyTestDouble(ByVal p_Type As CamPropertyType, ByVal p_Name As String, ByVal p_Min As Double, ByVal p_Max As Double, p_Mandatory As Boolean) As Double
        Try
            CameraPropertyTestDouble = 0.0
            Select Case p_Type
                Case CamPropertyType.CCDTemperature
                    CameraPropertyTestDouble = m_Camera.CCDTemperature
                Case CamPropertyType.CoolerPower
                    CameraPropertyTestDouble = m_Camera.CoolerPower
                Case CamPropertyType.ElectronsPerADU
                    CameraPropertyTestDouble = m_Camera.ElectronsPerADU
                Case CamPropertyType.FullWellCapacity
                    CameraPropertyTestDouble = m_Camera.FullWellCapacity
                Case CamPropertyType.HeatSinkTemperature
                    CameraPropertyTestDouble = m_Camera.HeatSinkTemperature
                Case CamPropertyType.PixelSizeX
                    CameraPropertyTestDouble = m_Camera.PixelSizeX
                Case CamPropertyType.PixelSizeY
                    CameraPropertyTestDouble = m_Camera.PixelSizeY
                Case CamPropertyType.SetCCDTemperature
                    CameraPropertyTestDouble = m_Camera.SetCCDTemperature
                Case CamPropertyType.ExposureMax
                    CameraPropertyTestDouble = m_Camera.ExposureMax
                Case CamPropertyType.ExposureMin
                    CameraPropertyTestDouble = m_Camera.ExposureMin
                Case CamPropertyType.ExposureResolution
                    CameraPropertyTestDouble = m_Camera.ExposureResolution
                Case CamPropertyType.SubExposureDuration
                    CameraPropertyTestDouble = m_Camera.SubExposureDuration
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestDouble: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case CameraPropertyTestDouble
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & CameraPropertyTestDouble.ToString)
                Case Is > p_Max 'Higher than maximum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & CameraPropertyTestDouble.ToString)
                Case Else 'OK value
                    LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestDouble.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    Private Function CameraPropertyTestBoolean(ByVal p_Type As CamPropertyType, ByVal p_Name As String, p_Mandatory As Boolean) As Boolean
        Try
            CameraPropertyTestBoolean = False
            Select Case p_Type
                Case CamPropertyType.CoolerOn
                    CameraPropertyTestBoolean = m_Camera.CoolerOn
                Case CamPropertyType.HasShutter
                    CameraPropertyTestBoolean = m_Camera.HasShutter
                Case CamPropertyType.ImageReady
                    CameraPropertyTestBoolean = m_Camera.ImageReady
                Case CamPropertyType.IsPulseGuiding
                    m_IsPulseGuidingSupported = False
                    CameraPropertyTestBoolean = m_Camera.IsPulseGuiding
                    m_IsPulseGuidingSupported = True 'Command works properly and doesn't cause a not implemented exception
                Case CamPropertyType.FastReadout
                    CameraPropertyTestBoolean = m_Camera.FastReadout
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestBoolean: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestBoolean.ToString)
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    Private Function CameraPropertyTestString(ByVal p_Type As CamPropertyType, ByVal p_Name As String, ByVal p_MaxLength As Integer, p_Mandatory As Boolean) As String
        CameraPropertyTestString = ""
        Try
            Select Case p_Type
                Case CamPropertyType.Description
                    CameraPropertyTestString = m_Camera.Description
                Case CamPropertyType.SensorName
                    CameraPropertyTestString = m_Camera.SensorName
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestString: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case CameraPropertyTestString
                Case Is = ""
                    LogMsg(p_Name, MessageLevel.msgOK, "The driver returned an empty string")
                Case Else
                    If Len(CameraPropertyTestString) <= p_MaxLength Then
                        LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestString)
                    Else
                        LogMsg(p_Name, MessageLevel.msgError, "String exceeds " & p_MaxLength & " characters maximum length - " & CameraPropertyTestString)
                    End If
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    Private Sub CameraPropertyWriteTest(ByVal p_Type As CamPropertyType, ByVal p_Property As String, ByVal p_TestOK As Integer)

        ' NOTE: Out of range values should not be tested here but later when the exposure is actually taken

        Try 'OK value first
            Select Case p_Type
                Case CamPropertyType.NumX
                    m_Camera.NumX = p_TestOK
                Case CamPropertyType.NumY
                    m_Camera.NumY = p_TestOK
                Case CamPropertyType.StartX
                    m_Camera.StartX = p_TestOK
                Case CamPropertyType.StartY
                    m_Camera.StartY = p_TestOK
            End Select
            LogMsg(p_Property & " write", MessageLevel.msgOK, "Successfully wrote " & p_TestOK)
        Catch ex As Exception
            HandleException(p_Property & " write", MemberType.Property, Required.MustBeImplemented, ex, "Can't set legal value: " & p_TestOK.ToString)
        End Try

    End Sub

    Overrides Sub CheckMethods()
        Dim l_i, l_j, l_MaxBinX, l_MaxBinY As Integer
        'AbortExposure - Mandatory
        Try
            m_CameraState = m_Camera.CameraState
            Select Case m_CameraState
                Case CameraStates.cameraIdle 'All is OK so test a call to AbortExposure
                    Try
                        m_Camera.AbortExposure()
                        If m_CanAbortExposure Then 'Can abort and no exception so OK
                            LogMsg("AbortExposure", MessageLevel.msgOK, "No exception generated when camera is already idle")
                        Else 'Can not abort and no exception - so violates specification
                            LogMsg("AbortExposure", MessageLevel.msgError, "CanAbortExposure is false but no exception is generated when AbortExposure is called")
                        End If
                    Catch ex As COMException
                        If m_CanAbortExposure Then 'Can abort so should not have generated an error
                            LogMsg("AbortExposure", MessageLevel.msgError, EX_COM & "exception incorrectly generated when camera is idle")
                        Else 'Can not abort so should have generated an error
                            LogMsg("AbortExposure", MessageLevel.msgOK, "CanAbortExposure is false and COM exception correctly generated")
                        End If
                    Catch ex As Exception
                        If m_CanAbortExposure Then 'Can abort so should not have generated an error
                            LogMsg("AbortExposure", MessageLevel.msgError, EX_NET & "exception incorrectly generated when camera is idle")
                        Else 'Can not abort so should have generated an error
                            LogMsg("AbortExposure", MessageLevel.msgOK, "CanAbortExposure is false and .NET exception correctly generated")
                        End If
                    End Try
                Case Else
                    LogMsg("AbortExposure", MessageLevel.msgError, EX_COM & "camera is not idle, further AbortExposure tests skipped: " & m_CameraState.ToString)
            End Select
        Catch ex As COMException
            LogMsg("AbortExposure", MessageLevel.msgError, EX_COM & "exception generated when reading camera state, further AbortExposure tests skipped")
        Catch ex As Exception
            LogMsg("AbortExposure", MessageLevel.msgError, EX_NET & "exception generated when reading camera state, further AbortExposure tests skipped")
        End Try
        'PulseGuide
        If m_CanPulseGuide Then 'Should work OK
            Try
                CameraPulseGuideTest(GuideDirections.guideNorth) : If TestStop() Then Exit Sub
                CameraPulseGuideTest(GuideDirections.guideSouth) : If TestStop() Then Exit Sub
                CameraPulseGuideTest(GuideDirections.guideEast) : If TestStop() Then Exit Sub
                CameraPulseGuideTest(GuideDirections.guideWest) : If TestStop() Then Exit Sub
            Catch ex As COMException
                LogMsg("PulseGuide", MessageLevel.msgError, EX_COM & "CanPulseGuide is true but exception generated when calling method - " & ex.ToString)
            Catch ex As Exception
                LogMsg("PulseGuide", MessageLevel.msgError, EX_NET & "CanPulseGuide is true but exception generated when calling method - " & ex.ToString)
            End Try
        Else 'CanPulseGuide is false, should generate an exception
            Try
                m_Camera.PulseGuide(GuideDirections.guideNorth, 0)
                LogMsg("PulseGuide", MessageLevel.msgError, "CanPulseGuide is false but no exception generated when calling method")
            Catch ex As COMException
                LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and exception correctly generated when calling method")
            Catch ex As MethodNotImplementedException
                LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and PulseGuide is not implemented in this driver")
            Catch ex As Exception
                LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and exception correctly generated when calling method")
            End Try
        End If

        'StopExposure
        Try
            m_CameraState = m_Camera.CameraState
            Select Case m_CameraState
                Case CameraStates.cameraIdle 'all is OK so test that a call to StopExposure doesn't generate an exception
                    Try
                        m_Camera.StopExposure()
                        If m_CanStopExposure Then 'Can abort and no exception so OK
                            LogMsg("StopExposure", MessageLevel.msgOK, "No exception generated when camera is already idle")
                        Else 'Can not abort and no exception - so violates specification
                            LogMsg("StopExposure", MessageLevel.msgError, "CanStopExposure is false but no exception is generated when AbortExposure is called")
                        End If
                    Catch ex As COMException
                        If m_CanStopExposure Then 'Can abort so should not have generated an error
                            LogMsg("StopExposure", MessageLevel.msgError, EX_COM & "exception incorrectly generated when camera is idle")
                        Else 'Can not abort so should have generated an error
                            LogMsg("StopExposure", MessageLevel.msgOK, "CanStopExposure is false and COM exception correctly generated")
                        End If
                    Catch ex As Exception
                        If m_CanStopExposure Then 'Can abort so should not have generated an error
                            LogMsg("StopExposure", MessageLevel.msgError, EX_NET & "exception incorrectly generated when camera is idle")
                        Else 'Can not abort so should have generated an error
                            LogMsg("StopExposure", MessageLevel.msgOK, "CanStopExposure is false and .NET exception correctly generated")
                        End If
                    End Try
                Case Else
                    LogMsg("StopExposure", MessageLevel.msgError, EX_COM & "camera is not idle, further StopExposure tests skipped: " & m_CameraState.ToString)
            End Select
        Catch ex As COMException
            LogMsg("StopExposure", MessageLevel.msgError, EX_COM & "exception generated when reading camera state, further StopExposure tests skipped")
        Catch ex As Exception
            LogMsg("StopExposure", MessageLevel.msgError, EX_NET & "exception generated when reading camera state, further StopExposure tests skipped")
        End Try

        'Use either the maximum values supplied by the camera driver or the maximum values defined in Conform's setup, if these have been set
        If (g_Settings.CameraMaxBinX > 0) Or (g_Settings.CameraMaxBinY > 0) Then LogMsg("", MessageLevel.msgAlways, "") ' Insert a blank line if required
        If g_Settings.CameraMaxBinX > 0 Then ' Use Conform's settings
            l_MaxBinX = g_Settings.CameraMaxBinX
            LogMsg("StartExposure", MessageLevel.msgComment, String.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", l_MaxBinX, m_MaxBinX))
        Else ' Use the camera's settings
            l_MaxBinX = m_MaxBinX
        End If
        If g_Settings.CameraMaxBinY > 0 Then ' Use Conform's settings
            l_MaxBinY = g_Settings.CameraMaxBinY
            LogMsg("StartExposure", MessageLevel.msgComment, String.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", l_MaxBinY, m_MaxBinY))
        Else ' Use the camera's settings
            l_MaxBinY = m_MaxBinY
        End If

        If (g_Settings.CameraMaxBinX > m_MaxBinX) Or (g_Settings.CameraMaxBinY > m_MaxBinY) Then LogMsg("", MessageLevel.msgAlways, "") ' Insert a blank line if required
        If g_Settings.CameraMaxBinX > m_MaxBinX Then LogMsg("StartExposure", MessageLevel.msgComment, String.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", l_MaxBinX, m_MaxBinX))
        If g_Settings.CameraMaxBinY > m_MaxBinY Then LogMsg("StartExposure", MessageLevel.msgComment, String.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", l_MaxBinY, m_MaxBinY))

        'StartExposure - Confirm that correct operation occurs
        Dim l_BinX, l_BinY As Integer
        If m_CanAsymmetricBin Then 'Test all combinations
            For l_BinY = 1 To l_MaxBinY
                For l_BinX = 1 To l_MaxBinX
                    CameraExposure("Take image full frame " & l_BinX & " x " & l_BinY & " bin", l_BinX, l_BinY, 0, 0, m_CameraXSize \ l_BinX, m_CameraYSize \ l_BinY, 2.0, "")
                    If TestStop() Then Exit Sub
                Next
            Next

        Else 'Test same bin combinations only
            For l_BinX = 1 To l_MaxBinX
                CameraExposure("Take image full frame " & l_BinX & " x " & l_BinX & " bin", l_BinX, l_BinX, 0, 0, m_CameraXSize \ l_BinX, m_CameraYSize \ l_BinX, 2.0, "")
                If TestStop() Then Exit Sub
            Next
        End If

        'StartExposure - Confirm error cases
        LogMsg("", MessageLevel.msgAlways, "")
        LogMsg("StartExposure error cases", MessageLevel.msgAlways, "")

        'StartExposure - Negative time
        CameraExposure("", 1, 1, 0, 0, m_CameraXSize, m_CameraYSize, -1.0, "negative duration") : If TestStop() Then Exit Sub 'Test that negative duration generates an error

        'StartExposure - Invalid Bin values
        For l_i = 1 To l_MaxBinX
            For l_j = 1 To l_MaxBinY
                If m_CanAsymmetricBin Then 'Carry out for all X and Y bin values
                    CameraExposure("", l_i, l_j, 0, 0, CInt((m_CameraXSize / l_i) + 1), CInt(m_CameraYSize / l_j), 0.1, "X size larger than binned chip size, Bin " & l_i & "x" & l_j) : If TestStop() Then Exit Sub 'X size too large for binned size
                    CameraExposure("", l_i, l_j, 0, 0, CInt(m_CameraXSize / l_i), CInt((m_CameraYSize / l_j) + 1), 0.1, "Y size larger than binned chip size, Bin " & l_i & "x" & l_j) : If TestStop() Then Exit Sub 'Y size too large for binned size
                    CameraExposure("", l_i, l_j, CInt((m_CameraXSize / l_i) + 1), 0, CInt(m_CameraXSize / l_i), CInt(m_CameraYSize / l_j), 0.1, "X start outside binned chip size, Bin " & l_i & "x" & l_j) : If TestStop() Then Exit Sub 'X start outside binned chip dimensions
                    CameraExposure("", l_i, l_j, 0, CInt((m_CameraYSize / l_j) + 1), CInt(m_CameraXSize / l_i), CInt(m_CameraYSize / l_j), 0.1, "Y start outside binned chip size, Bin " & l_i & "x" & l_j) : If TestStop() Then Exit Sub 'Y start outside binned chip dimensions
                Else 'Only carry out where X and Y bin are equal
                    If l_i = l_j Then 'X and Y bin are same
                        CameraExposure("", l_i, l_j, 0, 0, CInt((m_CameraXSize / l_i) + 1), CInt(m_CameraYSize / l_j), 0.1, "X size larger than binned chip size, Bin " & l_i & "x" & l_j) : If TestStop() Then Exit Sub 'X size too large for binned size
                        CameraExposure("", l_i, l_j, 0, 0, CInt(m_CameraXSize / l_i), CInt((m_CameraYSize / l_j) + 1), 0.1, "Y size larger than binned chip size, Bin " & l_i & "x" & l_j) : If TestStop() Then Exit Sub 'Y size too large for binned size
                        CameraExposure("", l_i, l_j, CInt((m_CameraXSize / l_i) + 1), 0, CInt(m_CameraXSize / l_i), CInt(m_CameraYSize / l_j), 0.1, "X start outside binned chip size, Bin " & l_i & "x" & l_j) : If TestStop() Then Exit Sub 'X start outside binned chip dimensions
                        CameraExposure("", l_i, l_j, 0, CInt((m_CameraYSize / l_j) + 1), CInt(m_CameraXSize / l_i), CInt(m_CameraYSize / l_j), 0.1, "Y start outside binned chip size, Bin " & l_i & "x" & l_j) : If TestStop() Then Exit Sub 'Y start outside binned chip dimensions
                    End If
                End If
            Next
        Next
    End Sub
    Private Sub CameraExposure(ByVal p_Description As String, ByVal p_BinX As Integer, ByVal p_BinY As Integer, ByVal p_StartX As Integer, ByVal p_StartY As Integer,
                               ByVal p_NumX As Integer, ByVal p_NumY As Integer, ByVal p_Duration As Double, ByVal p_ExpectedErrorMessage As String)
        Dim l_NumPlanes, l_VariantType, l_PercentCompletedMessage As String
        Dim l_ExposeOK, imageReadyTooEarly As Boolean 'Flag to determine whether we were successful or something failed
        Dim l_StartTime, l_StartTimeUTC, l_EndTime As Date
        Dim l_PercentCompleted As Short

        If p_Description <> "" Then
            LogMsg("", MessageLevel.msgAlways, "") 'Blank Line
            LogMsg(p_Description, MessageLevel.msgAlways, "")

        End If
        Try
            m_Camera.BinX = CShort(p_BinX)
            m_Camera.BinY = CShort(p_BinY)
            Try
                l_ExposeOK = False 'Start off by assuming the worst
                m_Camera.StartX = p_StartX
                m_Camera.StartY = p_StartY
                m_Camera.NumX = p_NumX
                m_Camera.NumY = p_NumY
                Try
                    Status(StatusType.staAction, "Start " & p_Duration.ToString & " second synchronous exposure")

                    ' Initiate exposure
                    l_StartTime = Now
                    l_StartTimeUTC = Date.UtcNow
                    m_Camera.StartExposure(p_Duration, True)

                    If p_ExpectedErrorMessage = "" Then 'Not expecting an error and didn't get one
                        l_EndTime = Now

                        ' Test whether we have a synchronous or asynchronous camera
                        If m_Camera.ImageReady And (m_Camera.CameraState = CameraStates.cameraIdle) Then 'Probably a synchronous camera
                            If l_EndTime.Subtract(l_StartTime).TotalSeconds >= p_Duration Then 'Is a synchronous camera
                                LogMsg("StartExposure", MessageLevel.msgOK, "Synchronous exposure found OK: " & p_Duration & " seconds")
                                CameraTestLast(p_Duration, l_StartTimeUTC)
                            Else ' Is synchronous but returned early!
                                LogMsg("StartExposure", MessageLevel.msgError, "Synchronous exposure found but image was returned before exposure time was complete")
                            End If
                        Else 'Assume asynchronous
                            Status(StatusType.staAction, "Waiting for exposure to start")

                            ' Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                            imageReadyTooEarly = CType(m_Camera.ImageReady, Boolean)

                            'Wait for exposing state
                            Do
                                WaitFor(CAMERA_SLEEP_TIME)
                                If TestStop() Then Exit Sub
                            Loop Until (m_Camera.CameraState = CameraStates.cameraExposing) Or (m_Camera.CameraState = CameraStates.cameraError)

                            ' Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                            imageReadyTooEarly = imageReadyTooEarly Or CType(m_Camera.ImageReady, Boolean)

                            'Wait for the exposing state to finish
                            l_StartTime = Now
                            l_StartTimeUTC = Date.UtcNow
                            Do
                                l_PercentCompletedMessage = "Not present in a V1 driver" ' Initialise PercentCompleted message
                                If m_Camera.InterfaceVersion > 1 Then
                                    Try
                                        l_PercentCompleted = m_Camera.PercentCompleted
                                        l_PercentCompletedMessage = l_PercentCompleted.ToString ' Operation completed OK
                                    Catch ex As COMException
                                        Select Case ex.ErrorCode
                                            Case ErrorCodes.NotImplemented
                                                l_PercentCompletedMessage = "COM - Not implemented" ' Not implemented
                                            Case ErrorCodes.InvalidOperationException
                                                l_PercentCompletedMessage = "COM - Invalid operation" ' Not valid at this time
                                            Case Else
                                                l_PercentCompletedMessage = "COM - Exception: " & ex.Message & " 0x" & ex.ErrorCode.ToString("X8") ' Something bad happened!
                                        End Select
                                    Catch ex As ASCOM.InvalidOperationException
                                        l_PercentCompletedMessage = ".NET - Invalid operation" ' Not valid at this time
                                    Catch ex As PropertyNotImplementedException
                                        l_PercentCompletedMessage = "Not implemented" ' Not implemented
                                    Catch ex As Exception
                                        l_PercentCompletedMessage = "Exception: " & ex.Message ' Something bad happened!
                                    End Try
                                End If
                                Status(StatusType.staAction, "Waiting for " & p_Duration.ToString & " second exposure to complete: " & Int(Now.Subtract(l_StartTime).TotalSeconds) & ",   PercentComplete: " & l_PercentCompletedMessage)
                                WaitFor(CAMERA_SLEEP_TIME)
                                If TestStop() Then Exit Sub
                            Loop Until (m_Camera.CameraState <> CameraStates.cameraExposing)

                            'Wait for camera to become idle
                            l_EndTime = Now
                            Status(StatusType.staAction, "Waiting for camera idle state, reading/downloading image")
                            Do
                                WaitFor(CAMERA_SLEEP_TIME)
                                If TestStop() Then Exit Sub
                            Loop Until (m_Camera.CameraState = CameraStates.cameraIdle) Or (m_Camera.CameraState = CameraStates.cameraError)

                            'Wait for image to become ready
                            Status(StatusType.staAction, "Waiting for image ready")
                            Do
                                WaitFor(CAMERA_SLEEP_TIME)
                                If TestStop() Then Exit Sub
                            Loop Until (m_Camera.ImageReady) Or (m_Camera.CameraState = CameraStates.cameraError)

                            If m_Camera.ImageReady Then 'Worked OK
                                LogMsg("StartExposure", MessageLevel.msgOK, "Asynchronous exposure found OK: " & p_Duration & " seconds")
                                CameraTestLast(p_Duration, l_StartTimeUTC)
                            Else 'Failed somehow
                                LogMsg("StartExposure", MessageLevel.msgError, "Camera state is CameraError")
                            End If
                        End If

                        ' Display a warning if ImageReady was set too early
                        If imageReadyTooEarly Then LogMsg("StartExposure", MessageLevel.msgIssue, "ImageReady was set True before the camera completed its exposure.")

                        'Camera exposed OK and didn't generate an exception
                        l_ExposeOK = True

                        'Check image array dimensions
                        Try
                            ' Retrieve the image array
                            If g_Settings.DisplayMethodCalls Then LogMsg("StartExposure", MessageLevel.msgComment, "About to call ImageArray method")
                            sw.Restart()
                            m_ImageArray = CType(m_Camera.ImageArray, Array)
                            sw.Stop()
                            If g_Settings.DisplayMethodCalls Then LogMsg("StartExposure", MessageLevel.msgComment, "Call completed in " & sw.ElapsedMilliseconds & "ms")

                            ' Examine the returned array
                            If (m_ImageArray.GetLength(0) = p_NumX) And (m_ImageArray.GetLength(1) = p_NumY) Then ' Image array dimensions match expected dimensions
                                If m_ImageArray.GetType.ToString = "System.Int32[,]" Or m_ImageArray.GetType.ToString = "System.Int32[,,]" Then ' Element types match the expected int32 type
                                    If m_ImageArray.Rank = 2 Then 'Single plane image be definition
                                        l_NumPlanes = "1 plane"
                                    Else 'Read the number of image planes from the maximum value of the third array index
                                        l_NumPlanes = "1 plane"
                                        If m_ImageArray.GetUpperBound(2) > 0 Then 'More than 1 plane
                                            l_NumPlanes = CStr(m_ImageArray.GetUpperBound(2) + 1) & " planes"
                                        End If
                                    End If
                                    LogMsg("ImageArray", MessageLevel.msgOK, "Successfully read 32 bit integer array (" & l_NumPlanes & ") " & m_ImageArray.GetLength(0) & " x " & m_ImageArray.GetLength(1) & " pixels")
                                Else ' Element types DO NOT match the expected int32 type
                                    LogMsg("ImageArray", MessageLevel.msgError, "Expected 32 bit integer array, actually got: " & m_ImageArray.GetType.ToString)
                                End If
                            Else ' Image array dimensions DO NOT match expected dimensions
                                If (m_ImageArray.GetLength(0) = p_NumY) And (m_ImageArray.GetLength(1) = p_NumX) Then ' Dimension values are swapped
                                    LogMsg("ImageArray", MessageLevel.msgError, "Camera image dimensions swapped, expected values: " & p_NumX & " x " & p_NumY & " - actual values: " & m_ImageArray.GetLength(0) & " x " & m_ImageArray.GetLength(1))
                                Else ' Dimension values are just wrong
                                    LogMsg("ImageArray", MessageLevel.msgError, "Camera image does not have the expected dimensions of: " & p_NumX & " x " & p_NumY & " - actual values: " & m_ImageArray.GetLength(0) & " x " & m_ImageArray.GetLength(1))
                                End If
                            End If
                        Catch ex As COMException
                            LogMsg("StartExposure", MessageLevel.msgError, EX_COM & "exception when reading ImageArray" & ex.ToString)
                        Catch ex As Exception
                            LogMsg("StartExposure", MessageLevel.msgError, EX_NET & "exception when reading ImageArray" & ex.ToString)
                        End Try

                        m_ImageArray = Nothing
                        m_ImageArrayVariant = Nothing
                        GC.Collect()

                        'Check image array variant dimensions
                        Dim imageArrayObject As Array
                        Try
                            If g_Settings.DisplayMethodCalls Then LogMsg("StartExposure", MessageLevel.msgComment, "About to call ImageArrayVariant method")
                            sw.Restart()
                            imageArrayObject = m_Camera.ImageArrayVariant
                            sw.Stop()
                            If g_Settings.DisplayMethodCalls Then LogMsg("StartExposure", MessageLevel.msgComment, "Call completed in " & sw.ElapsedMilliseconds & "ms")
                            sw.Restart()
                            m_ImageArrayVariant = CType(imageArrayObject, Array)
                            sw.Stop()
                            If g_Settings.DisplayMethodCalls Then LogMsg("StartExposure", MessageLevel.msgComment, "Conversion to Array completed in " & sw.ElapsedMilliseconds & "ms")
                            If (m_ImageArrayVariant.GetLength(0) = p_NumX) And (m_ImageArrayVariant.GetLength(1) = p_NumY) Then
                                If m_ImageArrayVariant.GetType.ToString = "System.Object[,]" Or m_ImageArrayVariant.GetType.ToString = "System.Object[,,]" Then
                                    If m_ImageArrayVariant.Rank = 2 Then 'Single plane image be definition
                                        l_NumPlanes = "1 plane"
                                        l_VariantType = m_ImageArrayVariant(0, 0).GetType.ToString()
                                    Else 'Read the number of image planes from the maximum value of the third array index
                                        l_NumPlanes = "1 plane"
                                        If m_ImageArrayVariant.GetUpperBound(2) > 0 Then 'More than 1 plane
                                            l_NumPlanes = CStr(m_ImageArrayVariant.GetUpperBound(2) + 1) & " planes"
                                            l_VariantType = m_ImageArrayVariant(0, 0, 0).GetType.ToString()
                                        Else 'Just one plane
                                            l_VariantType = m_ImageArrayVariant(0, 0).GetType.ToString()
                                        End If
                                    End If
                                    LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Successfully read variant array (" & l_NumPlanes & ") with " & l_VariantType & " elements " & m_ImageArrayVariant.GetLength(0) & " x " & m_ImageArrayVariant.GetLength(1) & " pixels")
                                Else
                                    LogMsg("ImageArrayVariant", MessageLevel.msgError, "Expected variant array, actually got: " & m_ImageArrayVariant.GetType.ToString)
                                End If
                            Else
                                If (m_ImageArrayVariant.GetLength(0) = p_NumY) And (m_ImageArrayVariant.GetLength(1) = p_NumX) Then
                                    LogMsg("ImageArrayVariant", MessageLevel.msgError, "Camera image dimensions swapped, expected values: " & p_NumX & " x " & p_NumY & " - actual values: " & m_ImageArrayVariant.GetLength(0) & " x " & m_ImageArrayVariant.GetLength(1))
                                Else
                                    LogMsg("ImageArrayVariant", MessageLevel.msgError, "Camera image does not have the expected dimensions of: " & p_NumX & " x " & p_NumY & " - actual values: " & m_ImageArrayVariant.GetLength(0) & " x " & m_ImageArrayVariant.GetLength(1))
                                End If
                            End If
                        Catch ex As COMException
                            LogMsg("ImageArrayVariant", MessageLevel.msgError, EX_COM & "exception when reading ImageArrayVariant" & ex.ToString)
                        Catch ex As Exception
                            LogMsg("ImageArrayVariant", MessageLevel.msgError, EX_NET & "exception when reading ImageArrayVariant" & ex.ToString)
                        End Try

                        ' Release large image objects from memory
                        m_ImageArrayVariant = Nothing
                        imageArrayObject = Nothing
                        GC.Collect()
                    Else 'Expecting an error and didn't get one!
                        LogMsg("StartExposure", MessageLevel.msgComment, "Test: " & p_ExpectedErrorMessage)
                        LogMsg("StartExposure", MessageLevel.msgError, "Expected an exception and didn't get one - BinX:" & p_BinX &
                                " BinY:" & p_BinY & " StartX:" & p_StartX & " StartY:" & p_StartY & " NumX:" & p_NumX &
                                " NumY:" & p_NumY)
                        Try : m_Camera.StopExposure() : Catch ex As Exception : End Try ' Try and do some clean up
                        Try : m_Camera.AbortExposure() : Catch ex As Exception : End Try
                    End If

                Catch ex As COMException
                    If p_ExpectedErrorMessage <> "" Then
                        LogMsg("StartExposure", MessageLevel.msgOK, "Exception correctly generated for " & p_ExpectedErrorMessage)
                    Else
                        LogMsg("", MessageLevel.msgInfo, p_ExpectedErrorMessage)
                        LogMsg("StartExposure", MessageLevel.msgError, EX_COM & "exception generated when exposing, further StartExposure tests skipped - " & ex.ToString)
                    End If
                Catch ex As Exception
                    If p_ExpectedErrorMessage <> "" Then
                        LogMsg("StartExposure", MessageLevel.msgOK, "Exception correctly generated for " & p_ExpectedErrorMessage)
                    Else
                        LogMsg("StartExposure", MessageLevel.msgError, EX_NET & "exception generated when exposing, further StartExposure tests skipped - " & ex.ToString)
                    End If
                End Try

            Catch ex As COMException
                LogMsg("StartExposure", MessageLevel.msgInfo, "COM Exception: " & ex.Message)
                LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped")
            Catch ex As Exception
                LogMsg("StartExposure", MessageLevel.msgInfo, ".NET Exception: " & ex.Message)
                LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped")
            End Try
        Catch ex As COMException
            Select Case ex.ErrorCode
                Case ErrorCodes.InvalidValue, g_ExInvalidValue1, g_ExInvalidValue2, g_ExInvalidValue3, g_ExInvalidValue4, g_ExInvalidValue5, g_ExInvalidValue6
                    LogMsg("StartExposure", MessageLevel.msgInfo, ex.Message)
                Case Else
                    LogMsg("StartExposure", MessageLevel.msgInfo, "COM Exception: " & ex.Message)
                    LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped")
            End Select
        Catch ex As InvalidValueException
            LogMsg("BinXY Write", MessageLevel.msgInfo, "Exposure skipped because BinX or BinY cannot be set. Exception message:")
            LogMsg("BinXY Write", MessageLevel.msgComment, ex.Message)
        Catch ex As Exception
            LogMsg("StartExposure", MessageLevel.msgInfo, ".NET Exception: " & ex.Message)
            LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped")

        End Try

        Status(StatusType.staAction, "")
    End Sub
    Private Sub CameraTestLast(ByVal p_Duration As Double, ByVal p_Start As Date)
        Dim l_StartTime As Date

        'LastExposureDuration
        Try
            m_LastExposureDuration = m_Camera.LastExposureDuration
            If (Math.Abs(m_LastExposureDuration - p_Duration) / p_Duration) < 0.02 Then ' Measured is within 2% of requested
                LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration is: " & m_LastExposureDuration & " seconds")
            Else
                LogMsg("LastExposureDuration", MessageLevel.msgIssue, "LastExposureDuration is not the same as image duration: " & m_LastExposureDuration & " " & p_Duration)
            End If
        Catch ex As Exception
            HandleException("LastExposureDuration", MemberType.Property, Required.Optional, ex, "")
        End Try

        'LastExposurestartTime
        Try 'Confirm that it can be read
            m_LastExposureStartTime = m_Camera.LastExposureStartTime
            'Confirm that the format is as expected
            Dim l_i As Integer, l_FormatOK As Boolean
            l_FormatOK = True
            If Len(m_LastExposureStartTime) >= 19 Then 'String is at least of correct minimum length
                For l_i = 1 To 19
                    Select Case l_i
                        Case 5, 8 '"-"
                            If Mid(m_LastExposureStartTime, l_i, 1) <> "-" Then l_FormatOK = False
                        Case 11 ' "T"
                            If Mid(m_LastExposureStartTime, l_i, 1) <> "T" Then l_FormatOK = False
                        Case 14, 17 '":"
                            If Mid(m_LastExposureStartTime, l_i, 1) <> ":" Then l_FormatOK = False
                        Case Else ' Numeric
                            If Not IsNumeric(Mid(m_LastExposureStartTime, l_i, 1)) Then l_FormatOK = False
                    End Select
                Next
                If l_FormatOK Then 'Passed format check so now try and read as a date-time to compare with expected value
                    Try 'Confirm that it parses as a valid date and check for correct value
                        l_StartTime = Date.Parse(m_LastExposureStartTime)
                        If p_Start.Subtract(l_StartTime).TotalSeconds < 2.0 Then
                            LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime is correct to within 2 seconds: " & m_LastExposureStartTime & " UTC")
                        Else
                            LogMsg("LastExposureStartTime", MessageLevel.msgIssue, "LastExposureStartTime is more than 2 seconds inaccurate : " & m_LastExposureStartTime & ", expected: " & Format(p_Start, "yyyy-MM-ddTHH:mm:ss") & " UTC")
                        End If
                    Catch ex As COMException
                        LogMsg("LastExposureStartTime", MessageLevel.msgError, EX_COM & "parsing LastExposureStartTime - " & ex.Message & " " & Hex(ex.ErrorCode))
                    Catch ex As Exception
                        LogMsg("LastExposureStartTime", MessageLevel.msgError, EX_NET & "parsing LastExposureStartTime - " & ex.Message & " " & m_LastExposureStartTime)
                    End Try
                Else 'Give format error message
                    LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime not in the expected format yyyy-mm-ddThh:mm:ss - " & m_LastExposureStartTime)
                End If
            Else 'Not of required length so give an error message
                If m_LastExposureStartTime = "" Then 'Empty string
                    LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime has returned an empty string - expected yyyy-mm-ddThh:mm:ss")
                Else 'Some characters but less than 19
                    LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime is less than 19 characters - expected yyyy-mm-ddThh:mm:ss - " & m_LastExposureStartTime)
                End If
            End If

        Catch ex As Exception
            HandleException("LastExposureStartTime", MemberType.Property, Required.Optional, ex, "")
        End Try

    End Sub
    Private Sub CameraPulseGuideTest(ByVal p_Direction As GuideDirections)
        Dim l_StartTime, l_EndTime As Date
        'Dim pulseGuideStatus As PulseGuideState

        l_StartTime = Now
        Status(StatusType.staAction, "Start " & CAMERA_PULSE_DURATION / 1000 & " second pulse guide " & p_Direction.ToString)
        m_Camera.PulseGuide(p_Direction, CAMERA_PULSE_DURATION) 'Start a 2 second pulse
        l_EndTime = Now

        If m_IsPulseGuidingSupported Then ' IsPulseGuiding is supported so go ahead and use it
            If l_EndTime.Subtract(l_StartTime).TotalMilliseconds < (CAMERA_PULSE_DURATION - 500) Then 'If less than 1.5 seconds then assume we have returned early
                If m_Camera.IsPulseGuiding Then ' IsPulseGuiding is true

                    '' Confirm correct operation of PulseGuideStatus
                    'pulseGuideStatus = m_Camera.PulseGuideStatus
                    'If pulseGuideStatus = PulseGuideState.Active Then
                    '    LogMsgOK("PulseGuide " & p_Direction.ToString, "PulseGuideStatus returns ""Active"" when IsPulseGuiding returns True.")
                    'Else
                    '    LogMsgError("PulseGuide " & p_Direction.ToString, $"PulseGuideStatus returns ""{pulseGuideStatus}"" rather than ""Active""when IsPulseGuiding returns True.")
                    'End If

                    Do
                        WaitFor(SLEEP_TIME)
                        Application.DoEvents()
                        If TestStop() Then Exit Sub
                    Loop Until (Not m_Camera.IsPulseGuiding) Or (Now.Subtract(l_StartTime).TotalMilliseconds > 3000) 'Wait for up to 3 seconds

                    If Not m_Camera.IsPulseGuiding Then
                        LogMsg("PulseGuide " & p_Direction.ToString, MessageLevel.msgOK, "Asynchronous pulse guide found OK")

                        '' Confirm correct operation of PulseGuideStatus
                        'pulseGuideStatus = m_Camera.PulseGuideStatus
                        'If pulseGuideStatus = PulseGuideState.Inactive Then
                        '    LogMsgOK("PulseGuide " & p_Direction.ToString, "PulseGuideStatus returns ""Inactive"" when IsPulseGuiding returns False.")
                        'Else
                        '    LogMsgError("PulseGuide " & p_Direction.ToString, $"PulseGuideStatus returns ""{pulseGuideStatus}"" rather than ""InActive""when IsPulseGuiding returns False.")
                        'End If
                    Else
                        LogMsg("PulseGuide " & p_Direction.ToString, MessageLevel.msgIssue, "Asynchronous pulse guide expected but IsPulseGuiding is TRUE beyond expected time of 2 seconds")
                    End If
                Else
                    LogMsg("PulseGuide " & p_Direction.ToString, MessageLevel.msgIssue, "Asynchronous pulse guide expected but IsPulseGuiding has returned FALSE")
                End If
            Else 'Assume synchronous pulse guide and that IsPulseGuiding is false
                If Not m_Camera.IsPulseGuiding Then
                    LogMsg("PulseGuide " & p_Direction.ToString, MessageLevel.msgOK, "Synchronous pulse guide found OK")
                Else
                    LogMsg("PulseGuide " & p_Direction.ToString, MessageLevel.msgIssue, "Synchronous pulse guide expected but IsPulseGuiding has returned TRUE")
                End If
            End If
        Else 'IsPulseGuiding is not supported so test for synchronous move
            Select Case l_EndTime.Subtract(l_StartTime).TotalMilliseconds - CAMERA_PULSE_DURATION
                Case Is > CAMERA_PULSE_TOLERANCE 'Duration was more than 0.5 seconds longer than expected
                    LogMsg("PulseGuide " & p_Direction.ToString, MessageLevel.msgIssue, "Synchronous pulse guide longer than expected " & (CAMERA_PULSE_DURATION) / 1000 & " seconds: " & l_EndTime.Subtract(l_StartTime).TotalSeconds & " seconds")
                Case Is < 20 'Duration was more than 20ms shorter than expected
                    LogMsg("PulseGuide " & p_Direction.ToString, MessageLevel.msgIssue, "Synchronous pulse guide shorter than expected " & (CAMERA_PULSE_DURATION) / 1000 & " seconds: " & l_EndTime.Subtract(l_StartTime).TotalSeconds & " seconds")
                Case Else 'Within acceptable tolerance
                    LogMsg("PulseGuide " & p_Direction.ToString, MessageLevel.msgOK, "Synchronous pulse guide found OK: " & l_EndTime.Subtract(l_StartTime).TotalSeconds & " seconds")
            End Select
        End If
    End Sub

    Public Overrides Sub CheckPerformance()
        CameraPerformanceTest(CameraPerformance.CameraState, "CameraState")
        CameraPerformanceTest(CameraPerformance.CCDTemperature, "CCDTemperature")
        CameraPerformanceTest(CameraPerformance.CoolerPower, "CoolerPower")
        If m_CanSetCCDTemperature Then CameraPerformanceTest(CameraPerformance.HeatSinkTemperature, "HeatSinkTemperature")
        CameraPerformanceTest(CameraPerformance.ImageReady, "ImageReady")
        If m_CanPulseGuide Then CameraPerformanceTest(CameraPerformance.IsPulseGuiding, "IsPulseGuiding")
        Status(StatusType.staAction, "Exposure for ImageArray Test")
        Status(StatusType.staStatus, "Start")
        m_Camera.BinX = 1
        m_Camera.BinY = 1
        m_Camera.StartX = 0
        m_Camera.StartY = 0
        m_Camera.NumX = m_Camera.MaxBinX
        m_Camera.NumY = m_Camera.MaxBinY
        m_Camera.StartExposure(1, True) '1 second exposure
        Do
            Status(StatusType.staStatus, "Waiting for ImageReady")
        Loop Until m_Camera.ImageReady
        Status(StatusType.staStatus, "Finished")
        CameraPerformanceTest(CameraPerformance.ImageArray, "ImageArray")
        CameraPerformanceTest(CameraPerformance.ImageArrayVariant, "ImageArrayVariant")
    End Sub
    Private Sub CameraPerformanceTest(ByVal p_Type As CameraPerformance, ByVal p_Name As String)
        Dim l_StartTime As Date, l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate As Double
        Status(StatusType.staAction, p_Name)
        Try
            l_StartTime = Now
            l_Count = 0.0
            l_LastElapsedTime = 0.0
            Do
                l_Count += 1.0
                Select Case p_Type
                    Case CameraPerformance.CameraState
                        m_CameraState = m_Camera.CameraState
                    Case CameraPerformance.CCDTemperature
                        m_CCDTemperature = m_Camera.CCDTemperature
                    Case CameraPerformance.CoolerPower
                        m_CoolerPower = m_Camera.CoolerPower
                    Case CameraPerformance.HeatSinkTemperature
                        m_HeatSinkTemperature = m_Camera.HeatSinkTemperature
                    Case CameraPerformance.ImageReady
                        m_ImageReady = m_Camera.ImageReady
                    Case CameraPerformance.IsPulseGuiding
                        m_IsPulseGuiding = m_Camera.IsPulseGuiding
                    Case CameraPerformance.ImageArray
                        m_ImageArray = CType(m_Camera.ImageArray, Array)
                    Case CameraPerformance.ImageArrayVariant
                        m_ImageArrayVariant = CType(m_Camera.ImageArrayVariant, Array)
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
                    LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " & Format(l_Rate, "0.0") & " per second")
                    Exit Select
                Case 2.0 To 10.0
                    LogMsg(p_Name, MessageLevel.msgOK, "Transaction rate: " & Format(l_Rate, "0.0") & " per second")
                    Exit Select
                Case 1.0 To 2.0
                    LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " & Format(l_Rate, "0.0") & " per second")
                    Exit Select
                Case Else
                    LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " & Format(l_Rate, "0.0") & " per second")
            End Select
        Catch ex As Exception
            LogMsg(p_Name, MessageLevel.msgInfo, "Unable to complete test: " & ex.ToString)
        End Try
    End Sub

    Public Overrides Sub PostRunCheck()
        If m_CanAbortExposure Then Try : m_Camera.AbortExposure() : Catch : End Try
        If m_CanStopExposure Then Try : m_Camera.StopExposure() : Catch : End Try
        If m_CanSetCCDTemperature Then Try : m_Camera.SetCCDTemperature = m_SetCCDTemperature : Catch : End Try
        Try : m_Camera.CoolerOn = m_CoolerOn : Catch : End Try
        LogMsg("PostRunCheck", MessageLevel.msgOK, "Camera returned to initial cooler temperature")
    End Sub

#End Region
End Class
