Imports ASCOM.DeviceInterface
Imports ASCOM.DriverAccess

Friend Class CoverCalibratorTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Private coverState As CoverStatus
    Private calibratorState As CalibratorStatus
    Private canAsynchronousOpen As Boolean
    Private asynchronousOpenTime As Double
    Private canAsynchronousClose As Boolean
    Private asynchronousCloseTime As Double
    Private maxBrightness As Integer
    Private calibratorStateOk As Boolean
    Private coverStateOk As Boolean
    Private brightnessOk As Boolean
    Private maxBrightnessOk As Boolean

#If DEBUG Then
    Private coverCalibratorDevice As ICoverCalibratorV1
#Else
    Private coverCalibratorDevice As Object
#End If
#End Region

#Region "Enums"
    Private Enum RequiredProperty
        CalibratorState
        CoverState
    End Enum

    Private Enum PerformanceProperty
        CalibratorState
        CoverState
    End Enum
#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(False, True, True, False, True, True, False) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub

    ' IDisposable
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            If True Then 'Should be True but make False to stop Conform from cleanly dropping the SafetyMonitor object (useful for retaining driver in memory to change flags)
                Try : coverCalibratorDevice.Connected = False : Catch : End Try
                Try : Marshal.ReleaseComObject(coverCalibratorDevice) : Catch : End Try
                coverCalibratorDevice = Nothing
                GC.Collect()
            End If

        End If
        MyBase.Dispose(disposing)
        Me.disposedValue = True
    End Sub
#End Region

#Region "Code"
    Public Overrides Sub CheckInitialise()
        'Set the error type numbers according to the standards adopted by individual authors.
        'Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
        'messages to driver authors!
        Select Case g_CoverCalibratorProgID
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040405
                g_ExInvalidValue2 = &H80040405
                g_ExInvalidValue3 = &H80040405
                g_ExInvalidValue4 = &H80040405
                g_ExInvalidValue5 = &H80040405
                g_ExInvalidValue6 = &H80040405
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_CoverCalibratorProgID)

    End Sub

    Public Overrides Sub CheckAccessibility()
        Dim deviceInterface As ICoverCalibratorV1
        Dim driverAccessDevice As CoverCalibrator
        Dim deviceObject As Object

        MyBase.CheckAccessibility(g_CoverCalibratorProgID, DeviceType.CoverCalibrator)

        deviceInterface = Nothing
        deviceObject = Nothing
        driverAccessDevice = Nothing

        'Try early binding
        Try
            deviceObject = CreateObject(g_CoverCalibratorProgID)
            deviceInterface = CType(deviceObject, ICoverCalibratorV1)

            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using early binding to ICoverCalibratorV1 interface")
            Try
                deviceInterface.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using early binding to ICoverCalibratorV1 interface")
                deviceInterface.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using early binding to ICoverCalibratorV1 interface: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using early binding to ICoverCalibratorV1: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(deviceInterface) : Catch : End Try
            Try : Marshal.ReleaseComObject(deviceObject) : Catch : End Try
            GC.Collect()
        End Try

        'Try client access toolkit
        Try
            driverAccessDevice = New CoverCalibrator(g_CoverCalibratorProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                driverAccessDevice.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                driverAccessDevice.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : driverAccessDevice.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(driverAccessDevice) : Catch : End Try
            GC.Collect()
        End Try
    End Sub

    Public Overrides Sub CreateDevice()
#If DEBUG Then
        LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.CoverCalibrator to get a CoverCalibrator object")
        coverCalibratorDevice = New CoverCalibrator(g_CoverCalibratorProgID)
        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.CoverCalibrator to get a CoverCalibrator object")
            coverCalibratorDevice = New CoverCalibrator(g_CoverCalibratorProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            coverCalibratorDevice = CreateObject(g_CoverCalibratorProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        End If
#End If
        g_Stop = False 'connected OK so clear stop flag to allow other tests to run
    End Sub

    Public Overrides Sub PreConnectChecks()
        'Confirm that key properties are false when not connected
        'Try
        '    m_IsSafe = coverCalibratorDevice.IsSafe
        '    If Not m_IsSafe Then
        '        LogMsg("IsSafe", MessageLevel.msgOK, "Reports false before connection")
        '    Else
        '        LogMsg("IsSafe", MessageLevel.msgIssue, "Reports true before connection rather than false")
        '    End If
        'Catch ex As Exception
        '    LogMsg("IsSafe", MessageLevel.msgError, "Cannot confirm that IsSafe is false before connection because it threw an exception: " & ex.Message)
        'End Try
    End Sub

    Public Overrides Property Connected() As Boolean
        Get
            Connected = coverCalibratorDevice.Connected
        End Get
        Set(ByVal value As Boolean)
            coverCalibratorDevice.Connected = value
        End Set
    End Property

    Public Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(coverCalibratorDevice, DeviceType.CoverCalibrator)
    End Sub

    Public Overrides Sub CheckProperties()
        Dim brightness As Integer

        calibratorStateOk = RequiredPropertiesTest(RequiredProperty.CalibratorState, "CalibratorState")
        coverStateOk = RequiredPropertiesTest(RequiredProperty.CoverState, "CoverState")

        If calibratorStateOk Then ' It is OK to test MaxBrightness
            maxBrightnessOk = False ' Assume a bad value
            Try
                maxBrightness = 0 ' Initialise to a 'bad' value
                If g_Settings.DisplayMethodCalls Then LogMsg("MaxBrightness", MessageLevel.msgComment, "About to call MaxBrightness property")
                maxBrightness = coverCalibratorDevice.MaxBrightness

                If Not coverCalibratorDevice.CalibratorState = CalibratorStatus.NotPresent Then ' There is a calibrator so continue testing
                    If maxBrightness >= 1 Then ' We received a valid value
                        LogMsg("MaxBrightness", MessageLevel.msgOK, maxBrightness)
                        maxBrightnessOk = True
                    Else ' We received an invalid value
                        LogMsg("MaxBrightness", MessageLevel.msgIssue, $"The returned MaxBrightness value {maxBrightness} is invalid, it must be >= 1")
                    End If
                Else ' Calibrator is reported as NotPresent but no exception was thrown
                    LogMsg("MaxBrightness", MessageLevel.msgError, $" CalibratorStatus is 'NotPresent' but MaxBrightness did not throw a PropertyNotImplementedException. It returned: {maxBrightness}.")
                End If
            Catch ex As Exception
                If coverCalibratorDevice.CalibratorState = CalibratorStatus.NotPresent Then ' There is no calibrator so MaxBrightness must throw a PropertyNotImplementedException
                    HandleException("MaxBrightness", MemberType.Property, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'")
                Else ' A The device does have calibration capability so this property must be implemented.
                    HandleException("MaxBrightness", MemberType.Property, Required.MustBeImplemented, ex, "CalibratorStatus indicates that he device is a calibrator")
                End If
            End Try
        Else
            LogMsg("MaxBrightness", MessageLevel.msgIssue, $"Test skipped because CalibratorState returned an exception")
        End If

        If calibratorStateOk Then ' It is OK to test Brightness 
            brightnessOk = False ' Assume a bad value

            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("Brightness", MessageLevel.msgComment, "About to call Brightness property")
                brightness = coverCalibratorDevice.Brightness
                brightnessOk = True

                If Not coverCalibratorDevice.CalibratorState = CalibratorStatus.NotPresent Then ' There is a calibrator so continue testing
                    If brightness >= 0 Then ' We received a potentially valid value
                        If maxBrightnessOk Then ' We can test to make sure the brightness value is less or equal to MaxBrightness
                            If brightness <= maxBrightness Then ' The returned brightness is valid
                                LogMsg("Brightness", MessageLevel.msgOK, maxBrightness)
                            Else ' The returned brightness is above the returned MaxBrightness and so is invalid
                                LogMsg("Brightness", MessageLevel.msgError, $"The returned Brightness {brightness} is greater than MaxBrightness {maxBrightness}")
                            End If
                        Else ' Cannot test against MaxBrightness so skip the test
                            LogMsg("Brightness", MessageLevel.msgIssue, $"Skipping Brightness test because MaxBrightness did not return a value.")
                        End If

                    Else ' We received a negative brightness, which is invalid
                        LogMsg("Brightness", MessageLevel.msgIssue, $"The returned Brightness value {brightness} is invalid, it must be >= 0")
                    End If
                Else ' Calibrator is reported as NotPresent but no exception was thrown
                    LogMsg("Brightness", MessageLevel.msgError, $" CalibratorStatus is 'NotPresent' but Brightness did not throw a PropertyNotImplementedException. It returned: {brightness}.")
                End If

            Catch ex As Exception
                If coverCalibratorDevice.CalibratorState = CalibratorStatus.NotPresent Then ' There is no calibrator so MaxBrightness must throw a PropertyNotImplementedException
                    HandleException("Brightness", MemberType.Property, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'")
                Else ' A The device does have calibration capability so this property must be implemented.
                    HandleException("Brightness", MemberType.Property, Required.MustBeImplemented, ex, "CalibratorStatus indicates that he device is a calibrator")
                End If
            End Try
        Else
            LogMsg("Brightness", MessageLevel.msgIssue, $"Test skipped because CalibratorState returned an exception")
        End If

        If coverStateOk And calibratorStateOk Then
            If coverState = CoverStatus.NotPresent And calibratorState = CalibratorStatus.NotPresent Then
                LogMsg("DeviceCapabilities", MessageLevel.msgWarning, "Both CoverStatus and CalibratorStatus are set to 'NotPresent' - this driver won't do a lot!")
            End If
        End If


    End Sub

    Public Overrides Sub CheckMethods()
        Dim startTime As DateTime
        g_Status.Clear()  'Clear status messages

        ' Test OpenCover
        If coverStateOk Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("OpenCover", MessageLevel.msgComment, "About to call OpenCover method")
                startTime = DateTime.Now

                coverCalibratorDevice.OpenCover()
                If Not coverState = CoverStatus.NotPresent Then ' Device has cover capability so assess the outcome
                    If Not coverCalibratorDevice.CoverState = CoverStatus.Moving Then ' Synchronous behaviour
                        canAsynchronousOpen = False
                        If coverCalibratorDevice.CoverState = CoverStatus.Open Then ' Successful outcome
                            LogMsg("OpenCover", MessageLevel.msgOK, $"OpenCover was successful. The synchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        Else ' Unsuccessful outcome
                            LogMsg("OpenCover", MessageLevel.msgError, $"OpenCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Open'. The synchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        End If
                    Else ' Asynchronous behaviour
                        canAsynchronousOpen = True
                        asynchronousOpenTime = 0.0

                        ' Wait until the cover is no longer moving
                        Do Until Not (coverCalibratorDevice.CoverState = CoverStatus.Moving)
                            WaitFor(10)
                        Loop
                        If coverCalibratorDevice.CoverState = CoverStatus.Open Then ' Successful outcome
                            asynchronousOpenTime = DateTime.Now.Subtract(startTime).TotalSeconds
                            LogMsg("OpenCover", MessageLevel.msgOK, $"OpenCover was successful. The asynchronous open took {asynchronousOpenTime.ToString("0.0")} seconds")
                        Else ' Unsuccessful outcome
                            LogMsg("OpenCover", MessageLevel.msgError, $"OpenCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Open'. The asynchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        End If
                    End If
                Else ' The device does not have cover capability and so should have thrown a MethodNotImplementedException
                    LogMsg("OpenCover", MessageLevel.msgError, $"CoverStatus is 'NotPresent' but OpenCover did not throw a MethodNotImplementedException.")
                End If
            Catch ex As Exception
                If coverState = CoverStatus.NotPresent Then ' There is no cover capability so expect a MethodNotImplementedException
                    HandleException("OpenCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'")
                Else ' The device does have cover capability so test accordingly
                    HandleException("OpenCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates the device has cover capability")
                End If
            End Try
        Else
            LogMsg("OpenCover", MessageLevel.msgIssue, $"Test skipped because CoverState returned an exception")
        End If

        ' Test CloseCover
        If coverStateOk Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("CloseCover", MessageLevel.msgComment, "About to call CloseCover method")
                startTime = DateTime.Now
                asynchronousCloseTime = 0.0

                coverCalibratorDevice.CloseCover()
                If Not coverState = CoverStatus.NotPresent Then ' Device has cover capability so assess the outcome
                    If Not coverCalibratorDevice.CoverState = CoverStatus.Moving Then ' Synchronous behaviour
                        canAsynchronousOpen = False
                        If coverCalibratorDevice.CoverState = CoverStatus.Closed Then ' Successful outcome
                            LogMsg("CloseCover", MessageLevel.msgOK, $"CloseCover was successful. The synchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        Else ' Unsuccessful outcome
                            LogMsg("CloseCover", MessageLevel.msgError, $"CloseCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Closed'. The synchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        End If
                    Else ' Asynchronous behaviour
                        canAsynchronousOpen = True
                        ' Wait until the cover is no longer moving
                        Do Until Not (coverCalibratorDevice.CoverState = CoverStatus.Moving)
                            WaitFor(10)
                        Loop
                        If coverCalibratorDevice.CoverState = CoverStatus.Closed Then ' Successful outcome
                            asynchronousCloseTime = DateTime.Now.Subtract(startTime).TotalSeconds
                            LogMsg("CloseCover", MessageLevel.msgOK, $"CloseCover was successful. The asynchronous close took {asynchronousCloseTime.ToString("0.0")} seconds")
                        Else ' Unsuccessful outcome
                            LogMsg("CloseCover", MessageLevel.msgError, $"CloseCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Closed'. The asynchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        End If
                    End If
                Else ' The device does not have cover capability and so should have thrown a MethodNotImplementedException
                    LogMsg("CloseCover", MessageLevel.msgError, $"CoverStatus is 'NotPresent' but CloseCover did not throw a MethodNotImplementedException.")
                End If

            Catch ex As Exception
                If coverState = CoverStatus.NotPresent Then ' There is no cover capability so expect a MethodNotImplementedException
                    HandleException("CloseCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'")
                Else ' The device does have cover capability so test accordingly
                    HandleException("CloseCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates the device has cover capability")
                End If
            End Try
        Else
            LogMsg("CloseCover", MessageLevel.msgIssue, $"Test skipped because CoverState returned an exception")
        End If

        ' Test HaltCover
        If coverStateOk Then
            If Not coverState = CoverStatus.NotPresent Then ' Cover is present so test accordingly

                If canAsynchronousOpen Then ' The device opens the cover asynchronously
                    If (asynchronousOpenTime > 0.0) And (asynchronousCloseTime > 0.0) Then ' Cover open and close were both successful so we can test the Halt command

                        ' Initiate a cover open first
                        If g_Settings.DisplayMethodCalls Then LogMsg("HaltCover", MessageLevel.msgComment, "About to call OpenCover method")
                        coverCalibratorDevice.OpenCover()

                        ' Wait for half of the expected cover open time
                        WaitFor(asynchronousOpenTime * 1000.0 / 2.0)

                        ' Confirm that he cover is still moving
                        If coverCalibratorDevice.CoverState = CoverStatus.Moving Then ' The cover is still moving so we can issue a HaltCover command

                            Try
                                ' Issue a halt command
                                If g_Settings.DisplayMethodCalls Then LogMsg("HaltCover", MessageLevel.msgComment, "About to call HaltCover method")
                                coverCalibratorDevice.HaltCover()

                                ' Confirm that the cover is no longer moving
                                If Not coverCalibratorDevice.CoverState = CoverStatus.Moving Then ' Cover is no longer moving so give a thumbs up.
                                    LogMsg("HaltCover", MessageLevel.msgOK, "Cover is no longer moving after issuing the HaltCover command")
                                Else ' Cover is still moving so this is an error
                                    LogMsg("HaltCover", MessageLevel.msgError, "Cover is still moving after issuing the HaltCover command")
                                End If
                            Catch ex As Exception
                                HandleException("HaltCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates that the device has cover capability")
                            End Try

                        Else ' The cover should have been moving but was not
                            LogMsg("HaltCover", MessageLevel.msgIssue, "Cover should have been moving after waiting for half of the previous open time, but it was not. Test abandoned")
                        End If
                    Else ' The cover open or cover closed operations failed
                        LogMsg("HaltCover", MessageLevel.msgIssue, $"HaltCover tests skipped because either the cover could not be opened or closed successfully.")
                    End If
                Else ' The device opens the cover synchronously
                    Try
                        ' Since the cover opens synchronously the HaltCover method should return a MethodNotImplementedException
                        If g_Settings.DisplayMethodCalls Then LogMsg("HaltCover", MessageLevel.msgComment, "About to call HaltCover method")
                        coverCalibratorDevice.HaltCover()
                        LogMsg("HaltCover", MessageLevel.msgError, "The cover operates synchronously but did not throw a MethodNotImplementedException in response to the HaltCover command")
                    Catch ex As Exception
                        If coverState = CoverStatus.NotPresent Then ' There is no cover capability so expect a MethodNotImplementedException
                            HandleException("HaltCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'")
                        Else ' The device does have cover capability so test accordingly
                            HandleException("HaltCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates the device has cover capability")
                        End If
                    End Try
                End If
            Else ' Cover is not present so confirm that a MethodNotImplementedException is thrown
                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("HaltCover", MessageLevel.msgComment, "About to call HaltCover method")
                    coverCalibratorDevice.HaltCover()
                    ' Should never get here...
                    LogMsg("HaltCover", MessageLevel.msgError, "CoverStatus is 'NotPresent' but HaltCover did not throw a MethodNotImplementedException")
                Catch ex As Exception
                    HandleException("HaltCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'")
                End Try

            End If
        Else
            LogMsg("HaltCover", MessageLevel.msgIssue, $"Test skipped because CoverState returned an exception")
        End If

        ' Test CalibratorOn
        If calibratorStateOk Then ' CalibratorState returned a value
            If Not calibratorState = CalibratorStatus.NotPresent Then ' Calibrator is present so test accordingly
                If maxBrightnessOk And brightnessOk Then ' OK to test different brightness levels

                    TestCalibratorOn(-1) ' Test for invalid value -1

                    TestCalibratorOn(0) ' Test for zero brightness

                    Select Case maxBrightness
                        Case 1 ' Simple on/ off device
                            TestCalibratorOn(1)
                        Case 2 ' Two brightness level device
                            TestCalibratorOn(1)
                            TestCalibratorOn(2)
                        Case 3 ' Three brightness level device
                            TestCalibratorOn(1)
                            TestCalibratorOn(2)
                            TestCalibratorOn(3)
                        Case 4 ' Four brightness level device
                            TestCalibratorOn(1)
                            TestCalibratorOn(2)
                            TestCalibratorOn(3)
                            TestCalibratorOn(4)
                        Case Else ' Devices with more than 4 brightness levels (we only test four brightness levels in these devices)
                            TestCalibratorOn(Math.Ceiling(((maxBrightness + 1.0) / 4.0) - 1.0)) ' Round up to ensure that this value is least 1 so there is some level of brightness
                            TestCalibratorOn(((maxBrightness + 1) / 2) - 1)
                            TestCalibratorOn(Math.Floor(((maxBrightness + 1.0) * 3.0 / 4.0) - 1.0)) ' Round down to ensure that this value is different to the maxBrightness value
                            TestCalibratorOn(maxBrightness)
                    End Select

                    If maxBrightness < Integer.MaxValue Then ' MaxBrightness is less than the largest possible four byte integer value
                        TestCalibratorOn(maxBrightness + 1) ' Test for invalid value of MaxBrightness + 1
                    Else ' MaxBrightness is equal to the largest possible four byte integer value
                        LogMsg("CalibratorOn", MessageLevel.msgInfo, $"Test of a high invalid brightness value skipped because MaxBrightness is set to the largest positive integer value.")
                    End If
                Else ' Invalid maxBrightness value so skip further brightness tests
                    LogMsg("CalibratorOn", MessageLevel.msgIssue, $"Brightness tests skipped because one of the Brightness or MaxBrightness properties returned an invalid value or threw an exception.")
                End If
            Else ' Calibrator is not present so just test once to make sure that it generates a MethodNotImplementedException
                TestCalibratorOn(1)
            End If
        Else ' CalibratorState threw an exception
            LogMsg("CalibratorOn", MessageLevel.msgIssue, $"Brightness tests skipped because the CoverState property returned an invalid value or threw an exception.")
        End If

        ' Test CalibratorOff
        If calibratorStateOk Then ' CalibratorState returned a value
            If Not calibratorState = CalibratorStatus.NotPresent Then ' Calibrator is present so test accordingly

                Try
                    startTime = DateTime.Now

                    If g_Settings.DisplayMethodCalls Then LogMsg("CalibratorOff", MessageLevel.msgComment, "About to call CalibratorOff method")
                    coverCalibratorDevice.CalibratorOff()

                    If Not coverCalibratorDevice.CalibratorState = CalibratorStatus.NotReady Then ' Synchronous behaviour
                        If coverCalibratorDevice.CalibratorState = CalibratorStatus.Off Then ' Successful outcome
                            LogMsg("CalibratorOff", MessageLevel.msgOK, $"CalibratorOff was successful. The synchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")

                            ' Confirm that Brightness returns to zero when calibrator is turned off
                            If g_Settings.DisplayMethodCalls Then LogMsg("CalibratorOff", MessageLevel.msgComment, "About to call Brightness property")
                            If coverCalibratorDevice.Brightness = 0 Then ' Brightness is set to zero when the calibrator is turned off
                                LogMsg("CalibratorOff", MessageLevel.msgOK, $"Brightness is set to zero when the calibrator is turned off")
                            Else ' Brightness is not set to zero when the calibrator is turned off
                                LogMsg("CalibratorOff", MessageLevel.msgIssue, $"Brightness is not set to zero when the calibrator is turned off")
                            End If

                        Else ' Unsuccessful outcome
                            LogMsg("CalibratorOff", MessageLevel.msgError, $"CalibratorOff was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Off'. The synchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        End If
                    Else ' Asynchronous behaviour

                        ' Wait until the cover is no longer moving
                        Do Until Not (coverCalibratorDevice.CalibratorState = CalibratorStatus.NotReady)
                            WaitFor(10)
                        Loop
                        If coverCalibratorDevice.CalibratorState = CalibratorStatus.Off Then ' Successful outcome
                            LogMsg("CalibratorOff", MessageLevel.msgOK, $"CalibratorOff was successful. The asynchronous action took {asynchronousCloseTime.ToString("0.0")} seconds")

                            ' Confirm that Brightness returns to zero when calibrator is turned off
                            If g_Settings.DisplayMethodCalls Then LogMsg("CalibratorOff", MessageLevel.msgComment, "About to call Brightness property")
                            If coverCalibratorDevice.Brightness = 0 Then ' Brightness is set to zero when the calibrator is turned off
                                LogMsg("CalibratorOff", MessageLevel.msgOK, $"Brightness is set to zero when the calibrator is turned off")
                            Else ' Brightness is not set to zero when the calibrator is turned off
                                LogMsg("CalibratorOff", MessageLevel.msgIssue, $"Brightness is not set to zero when the calibrator is turned off")
                            End If

                        Else ' Unsuccessful outcome
                            LogMsg("CalibratorOff", MessageLevel.msgError, $"CalibratorOff was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Off'. The asynchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        End If
                    End If

                Catch ex As Exception
                    HandleException("CalibratorOff", MemberType.Method, Required.MustBeImplemented, ex, "CalibratorStatus indicates that the device is a calibrator")
                End Try
            Else ' Calibrator is not present so test to make sure that it generates a MethodNotImplementedException

                Try
                    If g_Settings.DisplayMethodCalls Then LogMsg("CalibratorOff", MessageLevel.msgComment, "About to call CalibratorOff method")
                    coverCalibratorDevice.CalibratorOff()
                    ' Should never get here...
                    LogMsg("CalibratorOff", MessageLevel.msgError, $"CalibratorStatus is 'NotPresent'but CalibratorOff did not throw a MethodNotImplementedException.")
                Catch ex As Exception
                    HandleException("CalibratorOff", MemberType.Method, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'")
                End Try

            End If
        Else ' CalibratorState threw an exception
            LogMsg("CalibratorOff", MessageLevel.msgIssue, $"Test skipped because the CoverState property returned an invalid value or threw an exception.")
        End If

    End Sub

    Private Sub TestCalibratorOn(requestedBrightness As Integer)
        Dim returnedBrightness As Integer
        Dim startTime As DateTime

        If Not calibratorState = CalibratorStatus.NotPresent Then ' The device has calibrator capability so carry out the test

            Try
                startTime = DateTime.Now

                If g_Settings.DisplayMethodCalls Then LogMsg("CalibratorOn", MessageLevel.msgComment, $"About to call CalibratorOn method with brightness: {requestedBrightness}")
                coverCalibratorDevice.CalibratorOn(requestedBrightness)

                If Not coverCalibratorDevice.CalibratorState = CalibratorStatus.NotReady Then ' Synchronous behaviour so respond according to the returned status

                    If (requestedBrightness < 0) Or (requestedBrightness > maxBrightness) Then ' The supplied brightness value is invalid so should have thrown an exception
                        LogMsg("CalibratorOn", MessageLevel.msgIssue, $"CalibratorOn with brightness {requestedBrightness} should have thrown an InvalidValueException but did not.")
                    Else ' The supplied value is valid so the method should have worked OK
                        If coverCalibratorDevice.CalibratorState = CalibratorStatus.Ready Then ' Successful outcome
                            LogMsg("CalibratorOn", MessageLevel.msgOK, $"CalibratorOn with brightness {requestedBrightness} was successful. The synchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")

                            ' Confirm that the brightness value is what was set
                            If g_Settings.DisplayMethodCalls Then LogMsg("CalibratorOn", MessageLevel.msgComment, $"About to call Brightness property.")
                            returnedBrightness = coverCalibratorDevice.Brightness

                            If returnedBrightness = requestedBrightness Then ' The returned Brightness is the same as what was set
                                LogMsg("CalibratorOn", MessageLevel.msgOK, $"The Brightness property does return the value that was set")
                            Else ' The returned Brightness is different to what was set
                                LogMsg("CalibratorOn", MessageLevel.msgIssue, $"The Brightness property value: {returnedBrightness} does not match the value that was set: {requestedBrightness}")
                            End If

                        Else ' Unsuccessful outcome
                            LogMsg("CalibratorOn", MessageLevel.msgError, $"CalibratorOn with brightness {requestedBrightness} was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Ready'. The synchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        End If
                    End If

                Else ' Asynchronous behaviour
                    ' Wait until the cover is no longer moving
                    Do Until Not (coverCalibratorDevice.CalibratorState = CalibratorStatus.NotReady)
                        WaitFor(10)
                    Loop

                    If (requestedBrightness < 0) Or (requestedBrightness > maxBrightness) Then ' The supplied brightness value is invalid so should have thrown an exception
                        LogMsg("CalibratorOn", MessageLevel.msgIssue, $"CalibratorOn with brightness {requestedBrightness} should have thrown an InvalidValueException but did not.")
                    Else ' The supplied value is valid so the method should have worked OK
                        If coverCalibratorDevice.CalibratorState = CalibratorStatus.Ready Then ' Successful outcome
                            LogMsg("CalibratorOn", MessageLevel.msgOK, $"CalibratorOn with brightness {requestedBrightness} was successful. The asynchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        Else ' Unsuccessful outcome
                            LogMsg("CalibratorOn", MessageLevel.msgError, $"CalibratorOn with brightness {requestedBrightness} was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Ready'. The asynchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds")
                        End If
                    End If

                End If
            Catch ex As Exception
                If (requestedBrightness < 0) Or (requestedBrightness > maxBrightness) Then ' The supplied brightness value is invalid so expect an InvalidValueException
                    If IsInvalidValueException("CalibratorOn", ex) Then ' The invalid value exception was correctly generator
                        LogMsg("CalibratorOn", MessageLevel.msgOK, $"CalibratorOn with brightness {requestedBrightness} threw an InvalidValueException as expected")
                    Else ' Some other exception was raised so flag as an error
                        LogMsg("CalibratorOn", MessageLevel.msgError, $"CalibratorOn with brightness {requestedBrightness} threw an {ex.GetType.Name} exception an InvalidValueException was expected")
                    End If
                Else ' The supplied value is valid so the method should not have generated an exception
                    HandleException("CalibratorOn", MemberType.Method, Required.MustBeImplemented, ex, "CalibratorStatus indicates that the device is a calibrator")
                End If
            End Try
        Else ' There is no calibrator capability so expect a MethodNotImplementedException
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("CalibratorOn", MessageLevel.msgComment, $"About to call CalibratorOn method with brightness: {requestedBrightness}")
                coverCalibratorDevice.CalibratorOn(requestedBrightness)
                ' Should never get here...
                LogMsg("CalibratorOn", MessageLevel.msgError, $"CalibratorStatus is 'NotPresent'but CalibratorOn did not throw a MethodNotImplementedException.")
            Catch ex As Exception
                HandleException("CalibratorOn", MemberType.Method, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'")
            End Try
        End If
    End Sub


    Public Overrides Sub CheckPerformance()
        Status(StatusType.staTest, "Performance")

        PerformanceTest(PerformanceProperty.CalibratorState, "CalibratorState")
        PerformanceTest(PerformanceProperty.CoverState, "CoverState")

        Status(StatusType.staTest, "")
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")
    End Sub

    Private Function RequiredPropertiesTest(ByVal propertyToTest As RequiredProperty, ByVal propertyName As String) As Boolean
        Dim testWasSuccessful As Boolean

        Try
            testWasSuccessful = True ' Assume success
            Select Case propertyToTest
                Case RequiredProperty.CalibratorState
                    If g_Settings.DisplayMethodCalls Then LogMsg("CalibratorState", MessageLevel.msgComment, "About to call CalibratorState property")
                    calibratorState = coverCalibratorDevice.CalibratorState
                    LogMsg(propertyName, MessageLevel.msgOK, calibratorState.ToString)
                Case RequiredProperty.CoverState
                    If g_Settings.DisplayMethodCalls Then LogMsg("CoverState", MessageLevel.msgComment, "About to call CoverState property")
                    coverState = coverCalibratorDevice.CoverState
                    LogMsg(propertyName, MessageLevel.msgOK, coverState.ToString)
                Case Else
                    testWasSuccessful = False ' Flag that an issue occurred
                    LogMsg(propertyName, MessageLevel.msgError, "RequiredPropertiesTest: Unknown test type " & propertyToTest.ToString)
            End Select
        Catch ex As Exception
            testWasSuccessful = False ' Flag that an issue occurred
            HandleException(propertyName, MemberType.Property, Required.Mandatory, ex, "")
        End Try

        Return testWasSuccessful
    End Function

    Private Sub PerformanceTest(ByVal propertyToTest As PerformanceProperty, ByVal propertyName As String)
        Dim startTime As Date, loopCount, lastElapsedTime, elapsedTime, loopRate As Double
        Dim testCalibratorState As CalibratorStatus
        Dim testCoverState As CoverStatus

        Status(StatusType.staAction, propertyName)
        Try
            startTime = Now
            loopCount = 0.0
            lastElapsedTime = 0.0
            Do
                loopCount += 1.0
                Select Case propertyToTest
                    Case PerformanceProperty.CalibratorState
                        testCalibratorState = coverCalibratorDevice.CalibratorState
                    Case PerformanceProperty.CoverState
                        testCoverState = coverCalibratorDevice.CoverState
                    Case Else
                        LogMsg(propertyName, MessageLevel.msgError, "PerformanceTest: Unknown test type " & propertyToTest.ToString)
                End Select
                elapsedTime = Now.Subtract(startTime).TotalSeconds
                If elapsedTime > lastElapsedTime + 1.0 Then
                    Status(StatusType.staStatus, loopCount & " transactions in " & Format(elapsedTime, "0") & " seconds")
                    lastElapsedTime = elapsedTime
                    Application.DoEvents()
                    If TestStop() Then Exit Sub
                End If
            Loop Until elapsedTime > PERF_LOOP_TIME

            loopRate = loopCount / elapsedTime
            Select Case loopRate
                Case Is > 10.0
                    LogMsg(propertyName, MessageLevel.msgInfo, "Transaction rate: " & Format(loopRate, "0.0") & " per second")
                    Exit Select
                Case 2.0 To 10.0
                    LogMsg(propertyName, MessageLevel.msgOK, "Transaction rate: " & Format(loopRate, "0.0") & " per second")
                    Exit Select
                Case 1.0 To 2.0
                    LogMsg(propertyName, MessageLevel.msgInfo, "Transaction rate: " & Format(loopRate, "0.0") & " per second")
                    Exit Select
                Case Else
                    LogMsg(propertyName, MessageLevel.msgInfo, "Transaction rate: " & Format(loopRate, "0.0") & " per second")
            End Select
        Catch ex As Exception
            LogMsg(propertyName, MessageLevel.msgInfo, "Unable to complete test: " & ex.ToString)
        End Try
    End Sub

#End Region

End Class