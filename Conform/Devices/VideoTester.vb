Imports ASCOM.DeviceInterface
Imports ASCOM.DriverAccess

Friend Class VideoTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"
    Const CAMERA_PULSE_DURATION As Integer = 2000 'Duration of camera pulse guide test (ms)
    Const CAMERA_PULSE_TOLERANCE As Integer = 300 'Tolerance for acceptable;e performance (ms)

    'Camera variables
    Private CanConfigureDeviceProperties, CanReadCameraState, CanReadSensorType, CanReadGain, CanReadGainMax, CanReadGainMin, CanReadGains, CanReadFrameRate As Boolean
    Private CanReadGammaMin, CanReadGammaMax, CanReadGamma, CanReadGammas, CanReadIntegrationRate, CanReadSupportedIntegrationRates, CanReadLastVideoFrame As Boolean
    Private CanReadVideoFrame As Boolean
    Private PixelSizeX, PixelSizeY, ExposureMax, ExposureMin As Double
    Private BitDepth, Height, Width, IntegrationRate, VideoFramesBufferSize As Integer
    Private Gain, GainMax, GainMin, Gamma, GammaMin, GammaMax As Short
    Private Gains, Gammas, SupportedIntegrationRates As ArrayList
    Private SensorName As String
    Private SensorType As SensorType
    Private CameraState As ASCOM.DeviceInterface.VideoCameraState
    Private FrameRate As VideoCameraFrameRate
    Private LastVideoFrame As IVideoFrame
    Private ExposureStartTime, VideoCaptureDeviceName, VideoCodec, VideoFileFormat As String

    ' VideoFrame properties
    Private ExposureDuration As Double
    Private FrameNumber As Long
    Private ImageArray As Object
    Private ImageArrayAsArray As Array
    Private ImageMetadata As ArrayList
    Private PreviewBitmap As Byte()

#If DEBUG Then
    Private Video As ASCOM.DriverAccess.Video
#Else
    Private Video As Object
#End If
#End Region

#Region "Enums"
    Private Enum CanProperty
        CanConfigureDeviceProperties = 1
    End Enum
    Private Enum CameraPerformance As Integer
        CameraState
    End Enum
    Private Enum VideoProperty
        ' IVideo properties
        BitDepth
        CameraState
        CanConfigureDeviceProperties
        ExposureMax
        ExposureMin
        CCDTemperature
        FrameRate
        Gain
        GainMax
        GainMin
        Gains
        Gamma
        GammaMax
        GammaMin
        Gammas
        Height
        IntegrationRate
        LastVideoFrame
        PixelSizeX
        PixelSizeY
        SensorName
        SensorType
        SupportedIntegrationRates
        VideoCaptureDeviceName
        VideoCodec
        VideoFileFormat
        VideoFramesBufferSize
        Width

        'IVideoFrame Properties
        ExposureDuration
        ExposureStartTime
        FrameNumber
        ImageArray
        ImageMetadata
        PreviewBitmap
    End Enum

#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(True, True, True, False, False, True, True) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub
    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If True Then 'Should be True but make False to stop Conform from cleanly dropping the video object (useful for retaining driver in memory to change flags
                    Try : DisposeAndReleaseObject("Video", Video) : Catch : End Try

                    Video = Nothing
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
        MyBase.CheckCommonMethods(Video, DeviceType.Camera)
    End Sub

    Overrides Sub CheckInitialise()
        'Set the error type numbers according to the standards adopted by individual authors.
        'Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
        'messages to driver authors!
        Select Case g_VideoCameraProgID.ToUpper
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040405
                g_ExInvalidValue2 = &H80040405
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_VideoCameraProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_DeviceObject As Object
        Dim l_ICamera As Object, l_DriverAccessCamera As ASCOM.DriverAccess.Camera
        MyBase.CheckAccessibility(g_VideoCameraProgID, DeviceType.Camera)

        'Try early binding to V1 interface
        l_ICamera = Nothing
        l_DeviceObject = Nothing
        Try
            l_DeviceObject = CreateObject(g_VideoCameraProgID)
            l_ICamera = CType(l_DeviceObject, ASCOM.Interface.ICamera)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface ICamera")
            Try
                LogCallToDriver("AccessChecks", "About to set Connected property")
                l_ICamera.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ICamera")
                LogCallToDriver("AccessChecks", "About to set Connected property")
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
            l_DeviceObject = CreateObject(g_VideoCameraProgID)
            l_ICamera = CType(l_DeviceObject, ASCOM.DeviceInterface.ICameraV2)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface ICameraV2")
            Try
                LogCallToDriver("AccessChecks", "About to set Connected property")
                l_ICamera.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ICameraV2")
                LogCallToDriver("AccessChecks", "About to set Connected property")
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
            l_DriverAccessCamera = New ASCOM.DriverAccess.Camera(g_VideoCameraProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                LogCallToDriver("AccessChecks", "About to set Connected property")
                l_DriverAccessCamera.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                LogCallToDriver("AccessChecks", "About to set Connected property")
                l_DriverAccessCamera.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
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
        LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Video to get a Video object")
        Video = New ASCOM.DriverAccess.Video(g_VideoCameraProgID)
        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Video to get a Video object")
            Video = New ASCOM.DriverAccess.Video(g_VideoCameraProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Video object")
            Video = CreateObject(g_VideoCameraProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        End If
#End If
            g_Stop = False
    End Sub
    Overrides Property Connected() As Boolean
        Get
            LogCallToDriver("Connected", "About to get Connected property")
            Connected = Video.Connected
        End Get
        Set(ByVal value As Boolean)
            LogCallToDriver("Connected", "About to set Connected property")
            Video.Connected = value
            g_Stop = False
        End Set
    End Property
    Public Overrides Sub ReadCanProperties()
        ' IVideoV1 properties
        CameraCanTest(CanProperty.CanConfigureDeviceProperties, "CanConfigureDeviceProperties")
    End Sub
    Private Sub CameraCanTest(ByVal p_Type As CanProperty, ByVal p_Name As String)
        Try
            Select Case p_Type
                Case CanProperty.CanConfigureDeviceProperties
                    LogCallToDriver(p_Type.ToString(), "About to get CanConfigureDeviceProperties property")
                    CanConfigureDeviceProperties = Video.CanConfigureDeviceProperties
                    LogMsg(p_Name, MessageLevel.msgOK, CanConfigureDeviceProperties.ToString)
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "Conform:CanTest: Unknown test type " & p_Type.ToString)
            End Select
        Catch ex As COMException
            LogMsg(p_Name, MessageLevel.msgError, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
        Catch ex As Exception
            LogMsg(p_Name, MessageLevel.msgError, EX_NET & ex.Message)
        End Try
    End Sub

    Overrides Sub CheckProperties()

        ' BitDepth - Mandatory
        BitDepth = TestInteger(VideoProperty.BitDepth, 1, Integer.MaxValue, True) : If TestStop() Then Exit Sub

        ' CameraState - Mandatory
        Try
            CanReadCameraState = False
            LogCallToDriver("CameraState", "About to get VideoCameraRunning property")
            CameraState = VideoCameraState.videoCameraRunning
            CameraState = Video.CameraState
            CanReadCameraState = True
            LogMsg("CameraState Read", MessageLevel.msgOK, CameraState.ToString)
        Catch ex As Exception
            HandleException("CameraState Read", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        ' ExposureMin and ExpoosureMax Read - Mandatory
        ExposureMax = TestDouble(VideoProperty.ExposureMax, 0.0001, Double.MaxValue, True)
        ExposureMin = TestDouble(VideoProperty.ExposureMin, 0.0, Double.MaxValue, True)

        'Apply tests to resultant exposure values
        If ExposureMin <= ExposureMax Then
            LogMsg("ExposureMin", MessageLevel.msgOK, "ExposureMin is less than or equal to ExposureMax")
        Else
            LogMsg("ExposureMin", MessageLevel.msgError, "ExposureMin is greater than ExposureMax")
        End If

        ' FrameRate - Mandatory
        Try
            CanReadFrameRate = False
            FrameRate = VideoCameraFrameRate.PAL
            LogCallToDriver("FrameRate", "About to get FrameRate property")
            FrameRate = Video.FrameRate
            CanReadFrameRate = True
            LogMsg("FrameRate Read", MessageLevel.msgOK, FrameRate.ToString)
        Catch ex As Exception
            HandleException("FrameRate Read", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        ' Read the Gain properties - Optional
        GainMax = TestShort(VideoProperty.GainMax, 0, Short.MaxValue, False)
        GainMin = TestShort(VideoProperty.GainMin, 0, Short.MaxValue, False)
        Gains = TestArrayList(VideoProperty.Gains, False, Type.GetType("System.String"))
        Gain = TestShort(VideoProperty.Gain, 0, Short.MaxValue, False)

        ' Now apply tests to the resultant Gain values
        If CanReadGainMin Xor CanReadGainMax Then
            If CanReadGainMin Then
                LogMsg("GainMinMax", MessageLevel.msgError, "Can read GainMin but GainMax threw an exception")
            Else
                LogMsg("GainMinMax", MessageLevel.msgError, "Can read GainMax but GainMin threw an exception")
            End If
        Else
            LogMsg("GainMinMax", MessageLevel.msgOK, "Both GainMin and GainMax are readable or both throw exceptions")
        End If

        ' Read the Gamma properties - Optional
        GammaMax = TestShort(VideoProperty.GammaMax, 0, Short.MaxValue, False)
        GammaMin = TestShort(VideoProperty.GammaMin, 0, Short.MaxValue, False)
        Gammas = TestArrayList(VideoProperty.Gammas, False, Type.GetType("System.String"))
        Gamma = TestShort(VideoProperty.Gamma, 0, Short.MaxValue, False)

        ' Now apply tests to the resultant Gamma values
        If CanReadGammaMin Xor CanReadGammaMax Then
            If CanReadGammaMin Then
                LogMsg("GammaMinMax", MessageLevel.msgError, "Can read GammaMin but GammaMax threw an exception")
            Else
                LogMsg("GammaMinMax", MessageLevel.msgError, "Can read GammaMax but GammaMin threw an exception")
            End If
        Else
            LogMsg("GammaMinMax", MessageLevel.msgOK, "Both GammaMin and GammaMax are readable or both throw exceptions")
        End If

        ' Height and width - Mandatory
        Height = TestInteger(VideoProperty.Height, 1, Integer.MaxValue, True)
        Width = TestInteger(VideoProperty.Width, 1, Integer.MaxValue, True)

        ' Integration rates - Optional
        IntegrationRate = TestInteger(VideoProperty.IntegrationRate, 0, Integer.MaxValue, False)
        SupportedIntegrationRates = TestArrayList(VideoProperty.SupportedIntegrationRates, False, Type.GetType("System.Double"))

        ' Now apply tests to the resultant integration rate values
        If CanReadIntegrationRate Xor CanReadSupportedIntegrationRates Then
            If CanReadIntegrationRate Then
                LogMsg("IntegrationRates", MessageLevel.msgError, "Can read IntegrationRate but SupportedIntegrationRates threw an exception")
            Else
                LogMsg("IntegrationRates", MessageLevel.msgError, "Can read SupportedIntegrationRates but IntegrationRate threw an exception")
            End If
        Else
            LogMsg("IntegrationRates", MessageLevel.msgOK, "Both IntegrationRate and SupportedIntegrationRates are readable or both throw exceptions")
        End If

        ' Pixel size - Mandatory
        PixelSizeX = TestDouble(VideoProperty.PixelSizeX, 1.0, Double.PositiveInfinity, False) : If TestStop() Then Exit Sub
        PixelSizeY = TestDouble(VideoProperty.PixelSizeY, 1.0, Double.PositiveInfinity, False) : If TestStop() Then Exit Sub

        ' SensorName - Mandatory
        SensorName = TestString(VideoProperty.SensorName, 250, True)

        ' SensorType - Mandatory
        Try
            CanReadSensorType = False
            LogCallToDriver("SensorType", "About to get SensorType property")
            SensorType = Video.SensorType
            CanReadSensorType = True ' Set a flag to indicate that we have got a valid SensorType value
            'Successfully retrieved a value
            LogMsg("SensorType", MessageLevel.msgOK, SensorType.ToString)
        Catch ex As Exception
            HandleException("SensorType", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'VideoCaptureDeviceName
        VideoCaptureDeviceName = TestString(VideoProperty.VideoCaptureDeviceName, 1000, False)

        ' VideoCodec 
        VideoCodec = TestString(VideoProperty.VideoCodec, 1000, False)

        'VideoFileFormat 
        VideoFileFormat = TestString(VideoProperty.VideoFileFormat, 1000, True)

        'VideoFramesBufferSize 
        VideoFramesBufferSize = TestInteger(VideoProperty.VideoFramesBufferSize, 0, Integer.MaxValue, True)

        'LastVideoFrame
        LastVideoFrame = TestVideoFrame(VideoProperty.LastVideoFrame, 0, 0, True)

        ' Check contents of received frame
        If CanReadVideoFrame Then
            ExposureDuration = TestDouble(VideoProperty.ExposureDuration, 0.0, Double.MaxValue, False)
            ExposureStartTime = TestString(VideoProperty.ExposureStartTime, Integer.MaxValue, False)
            FrameNumber = TestLong(VideoProperty.FrameNumber, 0, Long.MaxValue, True)
            ImageMetadata = TestArrayList(VideoProperty.ImageMetadata, True, GetType(ASCOM.Utilities.KeyValuePair))

            Try
                ImageArray = LastVideoFrame.ImageArray
                Try
                    LogMsg("ImageArray", MessageLevel.msgOK, "Received an image object from the driver of type: " & ImageArray.GetType.Name)
                Catch ex As Exception
                    LogMsg("ImageArray", MessageLevel.msgInfo, "Received an image object from the driver of indeterminate type")
                End Try

                'Check image array dimensions
                Try
                    ImageArrayAsArray = CType(ImageArray, Array)
                    LogMsg("ImageArray", MessageLevel.msgOK, "  Received an array of rank: " & ImageArrayAsArray.Rank & ", length: " & ImageArrayAsArray.LongLength.ToString("#,0") & " and type: " & ImageArrayAsArray.GetType.Name)

                    Select Case ImageArrayAsArray.Rank
                        Case 1 ' Rank 1
                            If ImageArrayAsArray.GetType.Equals(GetType(Integer())) Then ' Received expected type of array
                                LogMsg("ImageArray", MessageLevel.msgOK, "  Received a 1 dimension Integer array as expected.")
                            Else ' Got the wrong array type
                                LogMsg("ImageArray", MessageLevel.msgError, "  Did not receive a 1 dimension Integer array as expected. Received: " & ImageArrayAsArray.GetType.Name)
                            End If
                        Case 2 ' Rank 2
                            If ImageArrayAsArray.GetType.Equals(GetType(Integer(,))) Then ' Received expected type of array
                                LogMsg("ImageArray", MessageLevel.msgOK, "  Received a 2 dimension Integer array as expected.")
                            Else ' Got the wrong array type
                                LogMsg("ImageArray", MessageLevel.msgError, "  Did not receive a 2 dimension Integer array as expected. Received: " & ImageArrayAsArray.GetType.Name)
                            End If
                        Case 3 ' Rank 3
                            If ImageArrayAsArray.GetType.Equals(GetType(Integer(,,))) Then ' Received expected type of array
                                LogMsg("ImageArray", MessageLevel.msgOK, "  Received a 3 dimension Integer array as expected.")
                            Else ' Got the wrong array type
                                LogMsg("ImageArray", MessageLevel.msgError, "  Did not receive a 3 dimension Integer array as expected. Received: " & ImageArrayAsArray.GetType.Name)
                            End If
                        Case Else
                            LogMsg("ImageArray", MessageLevel.msgError, "  Array rank is 0 or exceeds 3: " & ImageArrayAsArray.GetType.Name)
                    End Select

                    If CanReadSensorType Then ' Carry out tests
                        Select Case SensorType
                            Case DeviceInterface.SensorType.Color ' This camera returns multiple image planes of colour information
                                Select Case ImageArrayAsArray.Rank
                                    Case 1 ' Invalid configuration
                                        LogMsg("ImageArray", MessageLevel.msgError, "  The SensorType is Colour and the zero based array rank is 0. For a colour sensor the array rank must be 1 or 2.")
                                        LogMsg("ImageArray", MessageLevel.msgInfo, "  Please see the IVideoFrame.ImageArray entry in the Platform Help file for allowed combinations of SensorType and ImageArray format.")
                                    Case 2 ' NumPlanes x (Height * Width). NumPlanes should be 3
                                        CheckImage(ImageArrayAsArray, 3, Height * Width, 0)
                                    Case 3 ' NumPlanes x Height x Width. NumPlanes should be 3
                                        CheckImage(ImageArrayAsArray, 3, Height, Width)
                                    Case Else
                                        ' This is an unsupported rank 0 or >3 so create an error
                                        LogMsg("ImageArray", MessageLevel.msgError, "  The zero based array rank must be 1, 2 or 3 . The returned array had rank: " & ImageArrayAsArray.Rank)
                                End Select
                            Case Else
                                ' This camera returns just one plane that may be literally monochrome or may appear monochrome 
                                ' but contain encoded colour information e.g. Bayer RGGB format
                                Select Case ImageArrayAsArray.Rank
                                    Case 1 ' (Height * Width)
                                        CheckImage(ImageArrayAsArray, Height * Width, 0, 0)
                                    Case 2 ' Height x Width.
                                        CheckImage(ImageArrayAsArray, Height, Width, 0)
                                    Case 3 ' Invalid configuration
                                        LogMsg("ImageArray", MessageLevel.msgError, "  The SensorType is not Colour and the array rank is 3. For non-colour sensors the array rank must be 1 or 2.")
                                        LogMsg("ImageArray", MessageLevel.msgInfo, "  Please see the IVideoFrame.ImageArray entry in the Platform Help file for allowed combinations of SensorType and ImageArray format.")
                                    Case Else
                                        ' This is an unsupported rank 0 or >3 so create an error
                                        LogMsg("ImageArray", MessageLevel.msgError, "  The ImageArray rank must be 1, 2 or 3. The returned array had rank: " & ImageArrayAsArray.Rank)
                                End Select
                        End Select
                    Else ' Can't read SensorType so don't do further tests
                        LogMsg("ImageArray", MessageLevel.msgInfo, "SensorType could not be determined so ImageArray quality tests have been skipped")
                    End If

                Catch ex As COMException
                    LogMsg("ImageArray", MessageLevel.msgError, "Unexpected COM exception when testing ImageArray: " & ex.ToString)
                Catch ex As Exception
                    LogMsg("ImageArray", MessageLevel.msgError, "Unexpected .NET exception when testing ImageArray: " & ex.ToString)
                End Try

            Catch ex As Exception
                HandleException("PreviewBitmap", MemberType.Property, Required.Mandatory, ex, "")
            End Try

            Try
                PreviewBitmap = LastVideoFrame.PreviewBitmap
                LogMsg("PreviewBitmap", MessageLevel.msgOK, "Received an array with " & PreviewBitmap.Count.ToString("#,#,#") & " entries")
            Catch ex As Exception
                HandleException("PreviewBitmap", MemberType.Property, Required.Mandatory, ex, "")
            End Try
        Else
            LogMsg("", MessageLevel.msgInfo, "Skipping VideoFrame contents check because of issue reading LastVideoFrame")
        End If

    End Sub

    ''' <summary>
    ''' Reports whether the overall array size matches the expected size
    ''' </summary>
    ''' <param name="TestArray">The array to test</param>
    ''' <param name="Dimension1">Size of the first array dimension</param>
    ''' <param name="Dimension2">Size of the second dimension or 0 for not present</param>
    ''' <param name="Dimension3">Size of the third dimension or 0 for not present</param>
    ''' <remarks></remarks>
    Private Sub CheckImage(TestArray As Array, Dimension1 As Long, Dimension2 As Long, Dimension3 As Long)
        Dim Length As Long
        Const CommaFormat As String = "#,0"

        Length = Dimension1 * IIf(Dimension2 > 0, Dimension2, 1) * IIf(Dimension3 > 0, Dimension3, 1) ' Calculate the overall expected size

        If TestArray.LongLength = Length Then ' Total pixels is correct
            LogMsg("CheckImage", MessageLevel.msgOK, "  ImageArray has the expected total number of pixels: " & Length.ToString(CommaFormat))
        Else ' Mismatch in pixel lengths
            LogMsg("CheckImage", MessageLevel.msgError, "  ImageArray returned a total of " & TestArray.Length.ToString(CommaFormat) & " pixels instead of the expected number: " & Length.ToString(CommaFormat))
        End If

        If Dimension1 >= 1 Then ' We have at least a 1 dimension array
            If Dimension2 > 0 Then ' At least a 2 dimension array
                If Dimension3 >= 1 Then 'We have a 3 dimension array
                    If TestArray.GetLongLength(0) = Dimension1 Then ' OK: Dimension 1 lengths match
                        LogMsg("CheckImage", MessageLevel.msgOK, "  ImageArray dimension 1 has the expected length:: " & Dimension1.ToString(CommaFormat))
                    Else ' Error: Dimension 1 lengths do not match
                        LogMsg("CheckImage", MessageLevel.msgError, "  ImageArray dimension 1 does not has the expected length:: " & Dimension1.ToString(CommaFormat) & ", received: " & TestArray.GetLongLength(0).ToString(CommaFormat))
                    End If
                    If TestArray.GetLongLength(1) = Dimension2 Then ' OK: Dimension 2 lengths match
                        LogMsg("CheckImage", MessageLevel.msgOK, "  ImageArray dimension 2 has the expected length:: " & Dimension2.ToString(CommaFormat))
                    Else ' Error: Dimension 2 lengths do not match
                        LogMsg("CheckImage", MessageLevel.msgError, "  ImageArray dimension 2 does not has the expected length:: " & Dimension2.ToString(CommaFormat) & ", received: " & TestArray.GetLongLength(1).ToString(CommaFormat))
                    End If
                    If TestArray.GetLongLength(2) = Dimension3 Then ' OK: Dimension 3 lengths match
                        LogMsg("CheckImage", MessageLevel.msgOK, "  ImageArray dimension 3 has the expected length:: " & Dimension3.ToString(CommaFormat))
                    Else ' Error: Dimension 3 lengths do not match
                        LogMsg("CheckImage", MessageLevel.msgError, "  ImageArray dimension 3 does not has the expected length:: " & Dimension3.ToString(CommaFormat) & ", received: " & TestArray.GetLongLength(2).ToString(CommaFormat))
                    End If
                Else ' We have a two dimension array
                    If TestArray.GetLongLength(0) = Dimension1 Then ' OK: Dimension 1 lengths match
                        LogMsg("CheckImage", MessageLevel.msgOK, "  ImageArray dimension 1 has the expected length:: " & Dimension1.ToString(CommaFormat))
                    Else ' Error: Dimension 1 lengths do not match
                        LogMsg("CheckImage", MessageLevel.msgError, "  ImageArray dimension 1 does not has the expected length:: " & Dimension1.ToString(CommaFormat) & ", received: " & TestArray.GetLongLength(0).ToString(CommaFormat))
                    End If
                    If TestArray.GetLongLength(1) = Dimension2 Then ' OK: Dimension 2 lengths match
                        LogMsg("CheckImage", MessageLevel.msgOK, "  ImageArray dimension 2 has the expected length:: " & Dimension2.ToString(CommaFormat))
                    Else ' Error: Dimension 2 lengths do not match
                        LogMsg("CheckImage", MessageLevel.msgError, "  ImageArray dimension 2 does not has the expected length:: " & Dimension2.ToString(CommaFormat) & ", received: " & TestArray.GetLongLength(1).ToString(CommaFormat))
                    End If

                End If
            Else ' We have a 1 dimension array
                If TestArray.GetLongLength(0) = Dimension1 Then ' OK: Dimension 1 lengths match
                    LogMsg("CheckImage", MessageLevel.msgOK, "  ImageArray dimension 1 has the expected length:: " & Dimension1.ToString(CommaFormat))
                Else ' Error: Dimension 1 lengths do not match
                    LogMsg("CheckImage", MessageLevel.msgError, "  ImageArray dimension 1 does not has the expected length:: " & Dimension1.ToString(CommaFormat) & ", received: " & TestArray.GetLongLength(0).ToString(CommaFormat))
                End If
            End If
        Else ' Error Dimension1 should always be greater than 0
            LogMsg("CheckImage", MessageLevel.msgError, "  Dimension 1 is 0 it should never be!")
        End If


    End Sub


    Private Function TestShort(ByVal p_Type As VideoProperty, ByVal p_Min As Short, ByVal p_Max As Short, p_Mandatory As Boolean) As Short
        Dim MethodName As String

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() ' & " Read"
        Catch ex As Exception 'Deal with the possibility that we haven't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        Try
            TestShort = 0
            LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property")
            Select Case p_Type
                Case VideoProperty.GainMax
                    CanReadGainMax = False
                    TestShort = Video.GainMax
                    CanReadGainMax = True
                Case VideoProperty.GainMin
                    CanReadGainMin = False
                    TestShort = Video.GainMin
                    CanReadGainMin = True
                Case VideoProperty.Gain
                    CanReadGain = False
                    TestShort = Video.Gain
                    CanReadGain = True
                Case VideoProperty.GammaMax
                    CanReadGammaMax = False
                    TestShort = Video.GammaMax
                    CanReadGammaMax = True
                Case VideoProperty.GammaMin
                    CanReadGammaMin = False
                    TestShort = Video.GammaMin
                    CanReadGammaMin = True
                Case VideoProperty.Gamma
                    CanReadGamma = False
                    TestShort = Video.Gamma
                    CanReadGamma = True
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestShort: Unknown test type - " & p_Type.ToString)
            End Select

            'Successfully retrieved a value
            Select Case TestShort
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(MethodName, MessageLevel.msgError, "Invalid value: " & TestShort.ToString)
                Case Is > p_Max 'Higher than maximum value
                    LogMsg(MethodName, MessageLevel.msgError, "Invalid value: " & TestShort.ToString)
                Case Else 'OK value
                    LogMsg(MethodName, MessageLevel.msgOK, TestShort.ToString)
            End Select
        Catch ex As Exception
            HandleException(MethodName, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    Private Function CameraPropertyMustNotImplemented(ByVal p_Type As VideoProperty, ByVal p_Name As String) As Boolean

        CameraPropertyMustNotImplemented = True
        Try
            Select Case p_Type
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyMustNotImplemented: Unknown test type - " & p_Type.ToString)
            End Select
        Catch ex As COMException
            If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
                LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM)
            Else
                ExTest(p_Name, ex.Message, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
            End If
        Catch ex As PropertyNotImplementedException
            LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET)
        Catch ex As Exception
            ExTest(p_Name, ex.Message, EX_NET & ex.Message)
        End Try
        Return CameraPropertyMustNotImplemented ' Return success indicator, True means property did thrown the exception, False means that it did not
    End Function
    ''' <summary>
    ''' Test whether an integer is returned by a driver
    ''' </summary>
    ''' <param name="p_Type">Method to test</param>
    ''' <param name="p_Min">Lowest valid value</param>
    ''' <param name="p_Max">Highest valid value</param>
    ''' <param name="p_Mandatory">Mandatory method</param>
    ''' <returns>Integer value returned by the driver</returns>
    ''' <remarks></remarks>
    Private Function TestInteger(ByVal p_Type As VideoProperty, ByVal p_Min As Integer, ByVal p_Max As Integer, p_Mandatory As Boolean) As Integer
        Dim MethodName As String

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibility that we haven't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        Try
            TestInteger = 0
            LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property")
            Select Case p_Type
                Case VideoProperty.BitDepth
                    TestInteger = Video.BitDepth
                Case VideoProperty.CameraState
                    TestInteger = Video.CameraState
                Case VideoProperty.Height
                    TestInteger = Video.Height
                Case VideoProperty.IntegrationRate
                    CanReadIntegrationRate = False
                    TestInteger = Video.IntegrationRate
                    CanReadIntegrationRate = True
                Case VideoProperty.Width
                    TestInteger = Video.Width
                Case VideoProperty.VideoFramesBufferSize
                    TestInteger = Video.VideoFramesBufferSize
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestInteger: Unknown test type - " & p_Type.ToString)
            End Select

            'Successfully retrieved a value
            Select Case TestInteger
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(MethodName, MessageLevel.msgError, "Invalid value: " & TestInteger.ToString)
                Case Is > p_Max 'Higher than maximum value
                    LogMsg(MethodName, MessageLevel.msgError, "Invalid value: " & TestInteger.ToString)
                Case Else 'OK value
                    LogMsg(MethodName, MessageLevel.msgOK, TestInteger.ToString)
            End Select
        Catch ex As Exception
            HandleException(MethodName, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    ''' <summary>
    ''' Test whether an integer is returned by a driver
    ''' </summary>
    ''' <param name="p_Type">Method to test</param>
    ''' <param name="p_Min">Lowest valid value</param>
    ''' <param name="p_Max">Highest valid value</param>
    ''' <param name="p_Mandatory">Mandatory method</param>
    ''' <returns>Integer value returned by the driver</returns>
    ''' <remarks></remarks>
    Private Function TestLong(ByVal p_Type As VideoProperty, ByVal p_Min As Long, ByVal p_Max As Long, p_Mandatory As Boolean) As Long
        Dim MethodName As String

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibility that we haven't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        Try
            TestLong = 0
            Select Case p_Type
                Case VideoProperty.FrameNumber
                    TestLong = LastVideoFrame.FrameNumber
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestInteger: Unknown test type - " & p_Type.ToString)
            End Select

            'Successfully retrieved a value
            Select Case TestLong
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(MethodName, MessageLevel.msgError, "Invalid value: " & TestLong.ToString)
                Case Is > p_Max 'Higher than maximum value
                    LogMsg(MethodName, MessageLevel.msgError, "Invalid value: " & TestLong.ToString)
                Case Else 'OK value
                    LogMsg(MethodName, MessageLevel.msgOK, TestLong.ToString)
            End Select
        Catch ex As Exception
            HandleException(MethodName, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function

    ''' <summary>
    ''' Test whether an integer is returned by a driver
    ''' </summary>
    ''' <param name="p_Type">Method to test</param>
    ''' <param name="p_Min">Lowest valid value</param>
    ''' <param name="p_Max">Highest valid value</param>
    ''' <param name="p_Mandatory">Mandatory method</param>
    ''' <returns>Integer value returned by the driver</returns>
    ''' <remarks></remarks>
    Private Function TestVideoFrame(ByVal p_Type As VideoProperty, ByVal p_Min As Integer, ByVal p_Max As Integer, p_Mandatory As Boolean) As IVideoFrame
        Dim MethodName As String

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibility that we haven't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        TestVideoFrame = Nothing

        Try
            LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property")
            Select Case p_Type
                Case VideoProperty.LastVideoFrame
                    CanReadVideoFrame = False
                    TestVideoFrame = Video.LastVideoFrame
                    CanReadVideoFrame = True
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestVideoFrame: Unknown test type - " & p_Type.ToString)
            End Select

            'Successfully retrieved a value
            LogMsg(MethodName, MessageLevel.msgOK, "Successfully received VideoFrame")
        Catch ex As Exception
            HandleException(MethodName, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function

    Private Function TestDouble(ByVal p_Type As VideoProperty, ByVal p_Min As Double, ByVal p_Max As Double, p_Mandatory As Boolean) As Double
        Dim MethodName As String

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibility that we haven't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        Try
            TestDouble = 0.0
            LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property")
            Select Case p_Type
                Case VideoProperty.PixelSizeX
                    TestDouble = Video.PixelSizeX
                Case VideoProperty.PixelSizeY
                    TestDouble = Video.PixelSizeY
                Case VideoProperty.ExposureMax
                    TestDouble = Video.ExposureMax
                Case VideoProperty.ExposureMin
                    TestDouble = Video.ExposureMin
                Case VideoProperty.ExposureDuration
                    TestDouble = LastVideoFrame.ExposureDuration
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestDouble: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case TestDouble
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(MethodName, MessageLevel.msgError, "Invalid value: " & TestDouble.ToString)
                Case Is > p_Max 'Higher than maximum value
                    LogMsg(MethodName, MessageLevel.msgError, "Invalid value: " & TestDouble.ToString)
                Case Else 'OK value
                    LogMsg(MethodName, MessageLevel.msgOK, TestDouble.ToString)
            End Select
        Catch ex As Exception
            HandleException(MethodName, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    Private Function TestBoolean(ByVal p_Type As VideoProperty, p_Mandatory As Boolean) As Boolean
        Dim MethodName As String

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibility that we haven't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        Try
            TestBoolean = False
            Select Case p_Type
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestBoolean: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            LogMsg(MethodName, MessageLevel.msgOK, TestBoolean.ToString)
        Catch ex As Exception
            HandleException(MethodName, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    ''' <summary>
    ''' Test whether the driver returns a valid ArrayList
    ''' </summary>
    ''' <param name="p_Type">Property</param>
    ''' <param name="p_Mandatory"></param>
    ''' <param name="p_ItemType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TestArrayList(ByVal p_Type As VideoProperty, p_Mandatory As Boolean, p_ItemType As Type) As ArrayList
        Dim MethodName As String, Count As Integer

        TestArrayList = New ArrayList()

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibility that we haven't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        Try
            LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property")
            Select Case p_Type
                Case VideoProperty.Gains
                    CanReadGains = False
                    TestArrayList = Video.Gains
                    CanReadGains = True
                Case VideoProperty.Gammas
                    CanReadGammas = False
                    TestArrayList = Video.Gammas
                    CanReadGammas = True
                Case VideoProperty.SupportedIntegrationRates
                    CanReadSupportedIntegrationRates = False
                    TestArrayList = Video.SupportedIntegrationRates
                    CanReadSupportedIntegrationRates = True
                Case VideoProperty.ImageMetadata
                    TestArrayList = LastVideoFrame.ImageMetadata
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestArrayList: Unknown test type - " & p_Type.ToString)
            End Select

            'Successfully retrieved a value
            Count = 0
            LogMsg(MethodName, MessageLevel.msgOK, "Received an array containing " & TestArrayList.Count & " items.")

            For Each ListItem As Object In TestArrayList
                If ListItem.GetType.Equals(p_ItemType) Then
                    LogMsg(MethodName & "(" & Count & ")", MessageLevel.msgOK, "  " & ListItem.ToString)
                Else
                    LogMsg(MethodName, MessageLevel.msgError, "  Type of ArrayList item: " & ListItem.GetType.Name & " does not match expected type: " & p_ItemType.Name)
                End If
                Count += 1
            Next
        Catch ex As Exception
            HandleException(MethodName, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function

    Private Function TestString(ByVal p_Type As VideoProperty, ByVal p_MaxLength As Integer, p_Mandatory As Boolean) As String
        Dim MethodName As String

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibility that we haven't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        TestString = ""
        Try
            LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property")
            Select Case p_Type
                Case VideoProperty.SensorName
                    TestString = Video.SensorName
                Case VideoProperty.ExposureStartTime
                    TestString = LastVideoFrame.ExposureStartTime
                Case VideoProperty.VideoCaptureDeviceName
                    TestString = Video.VideoCaptureDeviceName
                Case VideoProperty.VideoCodec
                    TestString = Video.VideoCodec
                Case VideoProperty.VideoFileFormat
                    TestString = Video.VideoFileFormat
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestString: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case TestString
                Case Is = ""
                    LogMsg(MethodName, MessageLevel.msgOK, "The driver returned an empty string")
                Case Else
                    If Len(TestString) <= p_MaxLength Then
                        LogMsg(MethodName, MessageLevel.msgOK, TestString)
                    Else
                        LogMsg(MethodName, MessageLevel.msgError, "String exceeds " & p_MaxLength & " characters maximum length - " & TestString)
                    End If
            End Select
        Catch ex As Exception
            HandleException(MethodName, MemberType.Property, IIf(p_Mandatory, Required.Mandatory, Required.Optional), ex, "")
        End Try
    End Function
    Private Sub CameraPropertyWriteTest(ByVal p_Type As VideoProperty, ByVal p_Property As String, ByVal p_TestOK As Integer, ByVal p_TestLow As Integer, ByVal p_TestHigh As Integer)

        Try 'OK value first
            Select Case p_Type
            End Select
            LogMsg(p_Property & " write", MessageLevel.msgOK, "Successfully wrote " & p_TestOK)
        Catch ex As COMException
            LogMsg(p_Property & " write", MessageLevel.msgError, EX_COM & "exception generated when setting legal value: " & p_TestOK.ToString & " - " & ex.Message)
        Catch ex As Exception
            LogMsg(p_Property & " write", MessageLevel.msgError, EX_NET & "exception generated when setting legal value: " & p_TestOK.ToString & " - " & ex.Message)
        End Try
    End Sub

    Overrides Sub CheckMethods()

    End Sub

    Public Overrides Sub CheckPerformance()
        CameraPerformanceTest(CameraPerformance.CameraState, "CameraState")
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
                        CameraState = Video.CameraState
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
        Try : Video.StopRecordingVideoFile() : Catch : End Try
        LogMsg("PostRunCheck", MessageLevel.msgOK, "Camera returned to initial cooler temperature")
    End Sub

#End Region

End Class
