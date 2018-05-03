Imports System.Reflection
Imports System.Collections.Generic

Friend Class ObservingConditionsTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"
    Dim averageperiod, cloudCover, dewPoint, humidity, pressure, rainRate, skyBrightness, skyQuality, starFWHM, skyTemperature, temperature, windDirection, windGust, windSpeed As Double

    ' Variables to indicate whether each funciton is or is not implemented to that it is possible to check that for any given sensor, all three either are are or are not implemented
    Dim sensorisImplemented As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)
    Dim sensorHasDescription As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)
    Dim sensorHasTimeOfLastUpdate As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)

    Const ABSOLUTE_ZERO As Double = -273.15
    Const WATER_BOILING_POINT As Double = 100.0
    Const BAD_VALUE As Double = Double.NaN

    ' Valid sensor properties constants
    Private Const PROPERTY_CLOUDCOVER As String = "CloudCover"
    Private Const PROPERTY_DEWPOINT As String = "DewPoint"
    Private Const PROPERTY_HUMIDITY As String = "Humidity"
    Private Const PROPERTY_PRESSURE As String = "Pressure"
    Private Const PROPERTY_RAINRATE As String = "RainRate"
    Private Const PROPERTY_SKYBRIGHTNESS As String = "SkyBrightness"
    Private Const PROPERTY_SKYQUALITY As String = "SkyQuality"
    Private Const PROPERTY_SKYTEMPERATURE As String = "SkyTemperature"
    Private Const PROPERTY_STARFWHM As String = "StarFWHM"
    Private Const PROPERTY_TEMPERATURE As String = "Temperature"
    Private Const PROPERTY_WINDDIRECTION As String = "WindDirection"
    Private Const PROPERTY_WINDGUST As String = "WindGust"
    Private Const PROPERTY_WINDSPEED As String = "WindSpeed"

    ' Other property names
    Private Const PROPERTY_AVERAGEPERIOD As String = "AveragePeriod"
    Private Const PROPERTY_LATESTUPDATETIME As String = "LatestUpdateTime"
    Private Const PROPERTY_TIMESINCELASTUPDATE As String = "TimeSinceLastUpdate"
    Private Const PROPERTY_LASTUPDATETIME As String = "LastUpdateTime"

    ' List of valid ObservingConditions sensor properties
    Dim ValidSensors As List(Of String) = New List(Of String) From {
        PROPERTY_CLOUDCOVER,
        PROPERTY_DEWPOINT,
        PROPERTY_HUMIDITY,
        PROPERTY_PRESSURE,
        PROPERTY_RAINRATE,
        PROPERTY_SKYBRIGHTNESS,
        PROPERTY_SKYQUALITY,
        PROPERTY_SKYTEMPERATURE,
        PROPERTY_STARFWHM,
        PROPERTY_TEMPERATURE,
        PROPERTY_WINDDIRECTION,
        PROPERTY_WINDGUST,
        PROPERTY_WINDSPEED}

#If DEBUG Then
    Private m_ObservingConditions As DriverAccess.ObservingConditions ' 
#Else
    Private m_ObservingConditions As Object ' 
#End If

#End Region

#Region "Enums"
    Private Enum ObservingConditionsProperty
        ' Primary properties
        AveragePeriod
        CloudCover
        DewPoint
        Humidity
        Pressure
        RainRate
        SkyBrightness
        SkyQuality
        SkyTemperature
        StarFWHM
        Temperature
        WindDirection
        WindGust
        WindSpeed

        'TimeSinceLastUpdate method
        TimeSinceLastUpdateLatest
        TimeSinceLastUpdateCloudCover
        TimeSinceLastUpdateDewPoint
        TimeSinceLastUpdateHumidity
        TimeSinceLastUpdatePressure
        TimeSinceLastUpdateRainRate
        TimeSinceLastUpdateSkyBrightness
        TimeSinceLastUpdateSkyQuality
        TimeSinceLastUpdateSkyTemperature
        TimeSinceLastUpdateStarFWHM
        TimeSinceLastUpdateTemperature
        TimeSinceLastUpdateWindDirection
        TimeSinceLastUpdateWindGust
        TimeSinceLastUpdateWindSpeed

        ' SensorDescription method
        SensorDescriptionCloudCover
        SensorDescriptionDewPoint
        SensorDescriptionHumidity
        SensorDescriptionPressure
        SensorDescriptionRainRate
        SensorDescriptionSkyBrightness
        SensorDescriptionSkyQuality
        SensorDescriptionSkyTemperature
        SensorDescriptionStarFWHM
        SensorDescriptionTemperature
        SensorDescriptionWindDirection
        SensorDescriptionWindGust
        SensorDescriptionWindSpeed

    End Enum

#End Region

#Region "New And Dispose"
    Sub New()
        MyBase.New(False, True, True, False, False, True, False) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        For Each sensorName As String In ValidSensors
            sensorisImplemented(sensorName) = False
            sensorHasDescription(sensorName) = False
            sensorHasTimeOfLastUpdate(sensorName) = False
        Next
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If
            If True Then 'Should be True but make False to stop Conform from cleanly dropping the telescope object (useful for retaining scopesim in memory to change flags
                Try : m_ObservingConditions.Connected = False : Catch : End Try
                Try : Marshal.ReleaseComObject(m_ObservingConditions) : Catch : End Try
                m_ObservingConditions = Nothing
                GC.Collect()
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        MyBase.Dispose(disposing)
        Me.disposedValue = True
    End Sub
#End Region

#Region "Code"
    Overrides Sub CheckInitialise()
        'Set the error type numbers acording to the standards adopted by individual authors.
        'Unfortunatley these vary between drivers so I have to allow for these here in order to give meaningful
        'messages to driver authors!

        Select Case g_ObservingConditionsProgID
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040405
                g_ExInvalidValue2 = &H80040405
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_ObservingConditionsProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_IObservingConditions As Object = Nothing, l_DriverAccessObservingConditions As DriverAccess.ObservingConditions
        Dim l_DeviceObject As Object = Nothing

        MyBase.CheckAccessibility(g_ObservingConditionsProgID, DeviceType.ObservingConditions)

        Try
            l_DeviceObject = CreateObject(g_ObservingConditionsProgID)
            l_IObservingConditions = CType(l_DeviceObject, ASCOM.DeviceInterface.IObservingConditions)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through Interface IObservingConditions")
            Try
                l_IObservingConditions.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes Interface IObservingConditions")
                l_IObservingConditions.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does Not expose Interface IObservingConditions")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does Not expose Interface IObservingConditions")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IObservingConditions) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IObservingConditions = Nothing
            GC.Collect()
        End Try

        WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting For device driver To be cleaned up by operating system")

        'Try client access toolkit
        l_DriverAccessObservingConditions = Nothing
        Try
            l_DriverAccessObservingConditions = New ASCOM.DriverAccess.ObservingConditions(g_ObservingConditionsProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver Using driver access toolkit")
            Try
                l_DriverAccessObservingConditions.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected Using driver access toolkit")
                l_DriverAccessObservingConditions.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error conecting To driver Using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : l_DriverAccessObservingConditions.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DriverAccessObservingConditions) : Catch : End Try
            l_DriverAccessObservingConditions = Nothing
            GC.Collect()
        End Try
    End Sub
    Overrides Sub CreateDevice()
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.ObservingConditions to get an ObservingConditions object")
            m_ObservingConditions = New ASCOM.DriverAccess.ObservingConditions(g_ObservingConditionsProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get an ObservingConditions object")
            m_ObservingConditions = CreateObject(g_ObservingConditionsProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        End If
        g_Stop = False 'connected OK so clear stop flag to allow other tests to run
    End Sub
    Overrides Property Connected() As Boolean
        Get
            Connected = MyBase.Connected
        End Get
        Set(ByVal value As Boolean)
            m_ObservingConditions.Connected = value
            MyBase.Connected = value
        End Set
    End Property
    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(m_ObservingConditions, DeviceType.ObservingConditions)
    End Sub

    Overrides Sub CheckProperties()

        averageperiod = TestDouble(PROPERTY_AVERAGEPERIOD, ObservingConditionsProperty.AveragePeriod, 0.0, 100000.0, Required.Mandatory) ' AveragePeriod is mandatory

        'AveragePeriod Write - Mandatory
        If IsGoodValue(averageperiod) Then
            ' Try setting a bad value i.e. less than -1, this should be rejected
            Try 'Invalid low value
                m_ObservingConditions.AveragePeriod = -2.0
                LogMsgIssue("AveragePeriod Write", "No error generated on setting the average period < -1.0")
            Catch ex As Exception
                HandleInvalidValueExceptionAsOK("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "", "Invalid Value exception generated as expected on set average period < -1.0")
            End Try

            ' Try setting a good zero value, this should be accepted
            Try 'Invalid low value
                m_ObservingConditions.AveragePeriod = 0.0
                LogMsgOK("AveragePeriod Write", "Successfully set average period to 0.0")
            Catch ex As Exception
                HandleException("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "")
            End Try

            ' Try setting a good positive value, this should be accepted
            Try 'Invalid low value
                m_ObservingConditions.AveragePeriod = 5.0
                LogMsgOK("AveragePeriod Write", "Successfully set average period to 5.0")
            Catch ex As Exception
                HandleInvalidValueExceptionAsOK("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "", "Invalid value exception thrown when average period is set to 5.0, which is permitted by the specification")
            End Try

            ' Restore original value, this should be accepted
            Try 'Invalid low value
                m_ObservingConditions.AveragePeriod = averageperiod
                LogMsgOK("AveragePeriod Write", "Successfully restored original average period: " & averageperiod)
            Catch ex As Exception
                LogMsgInfo("AveragePeriod Write", "Unable to restore original average period")
                HandleException("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "")
            End Try
        Else
            LogMsgInfo("AveragePeriod Write", "Test skipped because AveragePerid cold not be read")
        End If

        cloudCover = TestDouble(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.CloudCover, 0.0, 100.0, Required.Optional)
        dewPoint = TestDouble(PROPERTY_DEWPOINT, ObservingConditionsProperty.DewPoint, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional)
        humidity = TestDouble(PROPERTY_HUMIDITY, ObservingConditionsProperty.Humidity, 0.0, 100.0, Required.Optional)

        If (IsGoodValue(dewPoint) And IsGoodValue(humidity)) Then ' Two good values so OK
            LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both implemented per the interface specification")
        ElseIf (Not IsGoodValue(dewPoint) And Not IsGoodValue(humidity)) Then ' Two not implemented so OK
            LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both not implemented per the interface specification")
        Else ' One implemented and the other not... bad acording to the Spec!
            LogMsg("DewPoint & Humidity", MessageLevel.msgIssue, "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification")
        End If

        pressure = TestDouble(PROPERTY_PRESSURE, ObservingConditionsProperty.Pressure, 0.0, 1100.0, Required.Optional)
        rainRate = TestDouble(PROPERTY_RAINRATE, ObservingConditionsProperty.RainRate, 0.0, 20000.0, Required.Optional)
        skyBrightness = TestDouble(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.SkyBrightness, 0.0, 1000000.0, Required.Optional)
        skyQuality = TestDouble(PROPERTY_SKYQUALITY, ObservingConditionsProperty.SkyQuality, -20.0, 30.0, Required.Optional)
        starFWHM = TestDouble(PROPERTY_STARFWHM, ObservingConditionsProperty.StarFWHM, 0.0, 1000.0, Required.Optional)
        skyTemperature = TestDouble(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.SkyTemperature, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional)
        temperature = TestDouble(PROPERTY_TEMPERATURE, ObservingConditionsProperty.Temperature, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional)
        windDirection = TestDouble(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.WindDirection, 0.0, 360.0, Required.Optional)
        windGust = TestDouble(PROPERTY_WINDGUST, ObservingConditionsProperty.WindGust, 0.0, 1000.0, Required.Optional)
        windSpeed = TestDouble(PROPERTY_WINDSPEED, ObservingConditionsProperty.WindSpeed, 0.0, 1000.0, Required.Optional)

        ' Additional test to confirm that the reported direction is 0.0 if the wind speed is reported as 0.0
        If (windSpeed = 0.0) Then
            If (windDirection = 0.0) Then
                LogMsg(PROPERTY_WINDSPEED, MessageLevel.msgOK, "Wind direction is reported as 0.0 when wind speed is 0.0")
            Else
                LogMsg(PROPERTY_WINDSPEED, MessageLevel.msgError, String.Format("When wind speed is reported as 0.0, wind direction should also be reported as 0.0, it is actually reported as {0}", windDirection))
            End If
        End If

    End Sub

    Overrides Sub CheckMethods()

        Dim LastUpdateTimeLatest, LastUpdateTimeCloudCover, LastUpdateTimeDewPoint, LastUpdateTimeHumidity, LastUpdateTimePressure, LastUpdateTimeRainRate, LastUpdateTimeSkyBrightness, LastUpdateTimeSkyQuality As Double
        Dim LastUpdateTimeStarFWHM, LastUpdateTimeSkyTemperature, LastUpdateTimeTemperature, LastUpdateTimeWindDirection, LastUpdateTimeWindGust, LastUpdateTimeWindSpeed As Double

        Dim SensorDescriptionCloudCover, SensorDescriptionDewPoint, SensorDescriptionHumidity, SensorDescriptionPressure, SensorDescriptionRainRate, SensorDescriptionSkyBrightness, SensorDescriptionSkyQuality As String
        Dim SensorDescriptionStarFWHM, SensorDescriptionSkyTemperature, SensorDescriptionTemperature, SensorDescriptionWindDirection, SensorDescriptionWindGust, SensorDescriptionWindSpeed As String

        ' TimeSinceLastUpdate
        LastUpdateTimeLatest = TestDouble(PROPERTY_LATESTUPDATETIME, ObservingConditionsProperty.TimeSinceLastUpdateLatest, -1.0, Double.MaxValue, Required.Mandatory)

        LastUpdateTimeCloudCover = TestDouble(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.TimeSinceLastUpdateCloudCover, -1.0, Double.MaxValue, Required.Optional)
        LastUpdateTimeDewPoint = TestDouble(PROPERTY_DEWPOINT, ObservingConditionsProperty.TimeSinceLastUpdateDewPoint, -1.0, Double.MaxValue, Required.Optional)
        LastUpdateTimeHumidity = TestDouble(PROPERTY_HUMIDITY, ObservingConditionsProperty.TimeSinceLastUpdateHumidity, -1.0, Double.MaxValue, Required.Optional)

        If (IsGoodValue(LastUpdateTimeDewPoint) And IsGoodValue(LastUpdateTimeHumidity)) Then ' Two good values so OK
            LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both implemented per the interface specification")
        ElseIf (Not IsGoodValue(LastUpdateTimeDewPoint) And Not IsGoodValue(LastUpdateTimeHumidity)) Then ' Two not implemented so OK
            LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both not implemented per the interface specification")
        Else ' One implemented and the other not... bad according to the Spec!
            LogMsg("DewPoint & Humidity", MessageLevel.msgIssue, "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification")
        End If

        LastUpdateTimePressure = TestDouble(PROPERTY_PRESSURE, ObservingConditionsProperty.TimeSinceLastUpdatePressure, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeRainRate = TestDouble(PROPERTY_RAINRATE, ObservingConditionsProperty.TimeSinceLastUpdateRainRate, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeSkyBrightness = TestDouble(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.TimeSinceLastUpdateSkyBrightness, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeSkyQuality = TestDouble(PROPERTY_SKYQUALITY, ObservingConditionsProperty.TimeSinceLastUpdateSkyQuality, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeStarFWHM = TestDouble(PROPERTY_STARFWHM, ObservingConditionsProperty.TimeSinceLastUpdateStarFWHM, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeSkyTemperature = TestDouble(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.TimeSinceLastUpdateSkyTemperature, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeTemperature = TestDouble(PROPERTY_TEMPERATURE, ObservingConditionsProperty.TimeSinceLastUpdateTemperature, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeWindDirection = TestDouble(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.TimeSinceLastUpdateWindDirection, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeWindGust = TestDouble(PROPERTY_WINDGUST, ObservingConditionsProperty.TimeSinceLastUpdateWindGust, Double.MinValue, Double.MaxValue, Required.Optional)
        LastUpdateTimeWindSpeed = TestDouble(PROPERTY_WINDSPEED, ObservingConditionsProperty.TimeSinceLastUpdateWindSpeed, Double.MinValue, Double.MaxValue, Required.Optional)

        'Refresh
        Try
            m_ObservingConditions.Refresh()
            LogMsg("Refesh", MessageLevel.msgOK, "Refreshed OK")
        Catch ex As Exception
            HandleException("Refresh", MemberType.Method, Required.Optional, ex, "")
        End Try

        'SensorDescrtiption
        SensorDescriptionCloudCover = TestSensorDescription(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.SensorDescriptionCloudCover, Integer.MaxValue, Required.Optional)
        SensorDescriptionDewPoint = TestSensorDescription(PROPERTY_DEWPOINT, ObservingConditionsProperty.SensorDescriptionDewPoint, Integer.MaxValue, Required.Optional)
        SensorDescriptionHumidity = TestSensorDescription(PROPERTY_HUMIDITY, ObservingConditionsProperty.SensorDescriptionHumidity, Integer.MaxValue, Required.Optional)

        If ((SensorDescriptionDewPoint Is Nothing) And (SensorDescriptionHumidity Is Nothing)) Then ' Two missing values so OK
            LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both not implemented per the interface specification")
        ElseIf ((Not SensorDescriptionDewPoint Is Nothing) And (Not SensorDescriptionHumidity Is Nothing)) Then ' Two are implemented so OK
            LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both implemented per the interface specification")
        Else ' One implemented and the other not... bad acording to the Spec!
            LogMsg("DewPoint & Humidity", MessageLevel.msgIssue, "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification")
        End If

        SensorDescriptionPressure = TestSensorDescription(PROPERTY_PRESSURE, ObservingConditionsProperty.SensorDescriptionPressure, Integer.MaxValue, Required.Optional)
        SensorDescriptionRainRate = TestSensorDescription(PROPERTY_RAINRATE, ObservingConditionsProperty.SensorDescriptionRainRate, Integer.MaxValue, Required.Optional)
        SensorDescriptionSkyBrightness = TestSensorDescription(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.SensorDescriptionSkyBrightness, Integer.MaxValue, Required.Optional)
        SensorDescriptionSkyQuality = TestSensorDescription(PROPERTY_SKYQUALITY, ObservingConditionsProperty.SensorDescriptionSkyQuality, Integer.MaxValue, Required.Optional)
        SensorDescriptionStarFWHM = TestSensorDescription(PROPERTY_STARFWHM, ObservingConditionsProperty.SensorDescriptionStarFWHM, Integer.MaxValue, Required.Optional)
        SensorDescriptionSkyTemperature = TestSensorDescription(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.SensorDescriptionSkyTemperature, Integer.MaxValue, Required.Optional)
        SensorDescriptionTemperature = TestSensorDescription(PROPERTY_TEMPERATURE, ObservingConditionsProperty.SensorDescriptionTemperature, Integer.MaxValue, Required.Optional)
        SensorDescriptionWindDirection = TestSensorDescription(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.SensorDescriptionWindDirection, Integer.MaxValue, Required.Optional)
        SensorDescriptionWindGust = TestSensorDescription(PROPERTY_WINDGUST, ObservingConditionsProperty.SensorDescriptionWindGust, Integer.MaxValue, Required.Optional)
        SensorDescriptionWindSpeed = TestSensorDescription(PROPERTY_WINDSPEED, ObservingConditionsProperty.SensorDescriptionWindSpeed, Integer.MaxValue, Required.Optional)

        ' Now check that the sensor value, description and lastupdated time are all either implemented or not implemented
        For Each sensorName As String In ValidSensors
            LogMsg("Consistency", MessageLevel.msgDebug, "Sensor name: " & sensorName)
            If (sensorisImplemented(sensorName) And sensorHasDescription(sensorName) And sensorHasTimeOfLastUpdate(sensorName)) Then
                LogMsg("Consistency - " & sensorName, MessageLevel.msgOK, "Sensor value, description and time since last update are all implemented as required by the specification")
            Else
                If ((Not sensorisImplemented(sensorName)) And (Not sensorHasDescription(sensorName)) And (Not sensorHasTimeOfLastUpdate(sensorName))) Then
                    LogMsg("Consistency - " & sensorName, MessageLevel.msgOK, "Sensor value, description and time since last update are all not implemented as required by the specification")
                Else
                    LogMsg("Consistency - " & sensorName, MessageLevel.msgIssue, "Sensor value is implemented: " & sensorisImplemented(sensorName) &
                           ", Sensor description is implemented: " & sensorHasDescription(sensorName) &
                           ", Sensor time since last update is implemented: " & sensorHasTimeOfLastUpdate(sensorName))
                    LogMsg("Consistency - " & sensorName, MessageLevel.msgInfo, "The ASCOM specification requires that sensor value, description and time since last update must either all be implemented or all not be implemented.")
                End If
            End If
        Next
    End Sub

    Overrides Sub CheckPerformance()
        Status(StatusType.staTest, "")
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")
    End Sub

    ''' <summary>
    ''' Tests whether a double has a good value or the NaN bad value indicator
    ''' </summary>
    ''' <param name="value">Variable to be tested</param>
    ''' <returns>Returns True if the variable has a good value, otherwise returns False</returns>
    ''' <remarks></remarks>
    Private Function IsGoodValue(value As Double) As Boolean
        Return Not Double.IsNaN(value)
    End Function

    Private Function TestDouble(ByVal p_Nmae As String, ByVal p_Type As ObservingConditionsProperty, ByVal p_Min As Double, ByVal p_Max As Double, p_Mandatory As Required) As Double
        Dim MethodName, SensorName As String, retryCount As Integer = 0, readOK As Boolean = False, methodType As MemberType = MemberType.Property, unexpectedError As Boolean = False

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibilty that we havn't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try
        If MethodName.StartsWith(PROPERTY_TIMESINCELASTUPDATE) Then ' Extract the sensor name from the TimeSinceLastUpdateXXX property name
            SensorName = MethodName.Substring(PROPERTY_TIMESINCELASTUPDATE.Length)
        Else
            SensorName = MethodName
        End If
        LogMsg("TestDouble", MessageLevel.msgDebug, "methodName: " & MethodName & ", SensorName: " & SensorName)
        sensorHasTimeOfLastUpdate(SensorName) = False

        Do
            Try
                TestDouble = BAD_VALUE
                Select Case p_Type

                    ' Primary properties
                    Case ObservingConditionsProperty.AveragePeriod
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.AveragePeriod
                    Case ObservingConditionsProperty.CloudCover
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.CloudCover
                    Case ObservingConditionsProperty.DewPoint
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.DewPoint
                    Case ObservingConditionsProperty.Humidity
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.Humidity
                    Case ObservingConditionsProperty.Pressure
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.Pressure
                    Case ObservingConditionsProperty.RainRate
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.RainRate
                    Case ObservingConditionsProperty.SkyBrightness
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.SkyBrightness
                    Case ObservingConditionsProperty.SkyQuality
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.SkyQuality
                    Case ObservingConditionsProperty.StarFWHM
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.StarFWHM
                    Case ObservingConditionsProperty.SkyTemperature
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.SkyTemperature
                    Case ObservingConditionsProperty.Temperature
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.Temperature
                    Case ObservingConditionsProperty.WindDirection
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.WindDirection
                    Case ObservingConditionsProperty.WindGust
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.WindGust
                    Case ObservingConditionsProperty.WindSpeed
                        methodType = MemberType.Property : TestDouble = m_ObservingConditions.WindSpeed

                    'TimeSincelastUpdate method
                    Case ObservingConditionsProperty.TimeSinceLastUpdateLatest
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate("")
                    Case ObservingConditionsProperty.TimeSinceLastUpdateCloudCover
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_CLOUDCOVER)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateDewPoint
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_DEWPOINT)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateHumidity
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_HUMIDITY)
                    Case ObservingConditionsProperty.TimeSinceLastUpdatePressure
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_PRESSURE)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateRainRate
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_RAINRATE)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateSkyBrightness
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYBRIGHTNESS)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateSkyQuality
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYQUALITY)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateStarFWHM
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_STARFWHM)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateSkyTemperature
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYTEMPERATURE)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateTemperature
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_TEMPERATURE)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateWindDirection
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDDIRECTION)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateWindGust
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDGUST)
                    Case ObservingConditionsProperty.TimeSinceLastUpdateWindSpeed
                        methodType = MemberType.Method : TestDouble = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDSPEED)

                    Case Else
                        LogMsg(MethodName, MessageLevel.msgError, "TestDouble: Unknown test type - " & p_Type.ToString)
                End Select
                readOK = True

                'Successfully retrieved a value so check validity
                Select Case TestDouble
                    Case Is < p_Min 'Lower than minimum value
                        LogMsg(MethodName, MessageLevel.msgError, "Invalid value (below minimum expected - " & p_Min.ToString() & "): " & TestDouble.ToString)
                    Case Is > p_Max 'Higher than maximum value
                        LogMsg(MethodName, MessageLevel.msgError, "Invalid value (above maximum expected - " & p_Max.ToString() & "): " & TestDouble.ToString)
                    Case Else 'OK value
                        LogMsg(MethodName, MessageLevel.msgOK, TestDouble.ToString)
                End Select
                If MethodName.StartsWith(PROPERTY_TIMESINCELASTUPDATE) Then ' Extract the sensor name from the TimeSinceLastUpdateXXX property name
                    sensorHasTimeOfLastUpdate(SensorName) = True
                Else
                    sensorisImplemented(SensorName) = True
                End If

            Catch ex As Exception
                If IsInvalidOperationException(p_Nmae, ex) Then ' We got an InvalidOperationException so retry the reading after the configured wait period
                    TestDouble = BAD_VALUE
                    retryCount = retryCount + 1
                    LogMsg(MethodName, MessageLevel.msgInfo, "Sensor not ready, received InvalidOperationException, waiting " & g_Settings.ObservingConditionsRetryTime & " second to retry. Attempt " & retryCount & " out of " & g_Settings.ObservingConditionsMaxRetries)
                    WaitFor(g_Settings.ObservingConditionsRetryTime * 1000)
                Else ' We got some other error so don't do any retrys and just record the error
                    unexpectedError = True
                    TestDouble = BAD_VALUE
                    HandleException(MethodName, methodType, p_Mandatory, ex, "")
                End If
            End Try
        Loop Until readOK Or (retryCount = g_Settings.ObservingConditionsMaxRetries) Or unexpectedError

        If (Not readOK) And (Not unexpectedError) Then ' We must have timed out after the specified number of retries so record this
            LogMsg(MethodName, MessageLevel.msgInfo, "InvalidOperationException persisted for longer than " & g_Settings.ObservingConditionsMaxRetries * g_Settings.ObservingConditionsRetryTime & " seconds.")
        End If

    End Function

    Private Function TestSensorDescription(ByVal p_Nmae As String, ByVal p_Type As ObservingConditionsProperty, ByVal p_MaxLength As Integer, p_Mandatory As Required) As String
        Dim MethodName As String

        'Create a text version of the calling method name
        Try
            MethodName = p_Type.ToString() '& " Read"
        Catch ex As Exception 'Deal with the possibilty that we havn't added it to the ENUM yet!
            MethodName = "?????? Read"
        End Try

        sensorHasDescription(p_Nmae) = False

        TestSensorDescription = Nothing
        Try
            Select Case p_Type
                'TimeSincelastUpdate method
                Case ObservingConditionsProperty.SensorDescriptionCloudCover
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_CLOUDCOVER)
                Case ObservingConditionsProperty.SensorDescriptionDewPoint
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_DEWPOINT)
                Case ObservingConditionsProperty.SensorDescriptionHumidity
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_HUMIDITY)
                Case ObservingConditionsProperty.SensorDescriptionPressure
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_PRESSURE)
                Case ObservingConditionsProperty.SensorDescriptionRainRate
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_RAINRATE)
                Case ObservingConditionsProperty.SensorDescriptionSkyBrightness
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_SKYBRIGHTNESS)
                Case ObservingConditionsProperty.SensorDescriptionSkyQuality
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_SKYQUALITY)
                Case ObservingConditionsProperty.SensorDescriptionStarFWHM
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_STARFWHM)
                Case ObservingConditionsProperty.SensorDescriptionSkyTemperature
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_SKYTEMPERATURE)
                Case ObservingConditionsProperty.SensorDescriptionTemperature
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_TEMPERATURE)
                Case ObservingConditionsProperty.SensorDescriptionWindDirection
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_WINDDIRECTION)
                Case ObservingConditionsProperty.SensorDescriptionWindGust
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_WINDGUST)
                Case ObservingConditionsProperty.SensorDescriptionWindSpeed
                    TestSensorDescription = m_ObservingConditions.SensorDescription(PROPERTY_WINDSPEED)
                Case Else
                    LogMsg(MethodName, MessageLevel.msgError, "TestString: Unknown test type - " & p_Type.ToString)
            End Select

            'Successfully retrieved a value
            Select Case TestSensorDescription
                Case Is = Nothing
                    LogMsg(MethodName, MessageLevel.msgError, "The driver did not return any string at all: Nothing (VB), null (C#)")
                Case Is = ""
                    LogMsg(MethodName, MessageLevel.msgOK, "The driver returned an empty string: """"")
                Case Else
                    If Len(TestSensorDescription) <= p_MaxLength Then
                        LogMsg(MethodName, MessageLevel.msgOK, TestSensorDescription)
                    Else
                        LogMsg(MethodName, MessageLevel.msgError, "String exceeds " & p_MaxLength & " characters maximum length - " & TestSensorDescription)
                    End If
            End Select

            sensorHasDescription(p_Nmae) = True

        Catch ex As Exception
            HandleException(MethodName, MemberType.Method, p_Mandatory, ex, "")
            TestSensorDescription = Nothing
        End Try

    End Function

#End Region

End Class
