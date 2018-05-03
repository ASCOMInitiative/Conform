Friend Class FilterWheelTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"
    Const FILTER_WHEEL_TIME_OUT As Integer = 10 'Filter wheel command timeout (seconds)
#If DEBUG Then
    Private m_FilterWheel As ASCOM.DeviceInterface.IFilterWheelV2
#Else
    Private m_FilterWheel As Object
#End If
#End Region

#Region "Enums"
    Enum FilterWheelProperties
        FocusOffsets
        Names
        Position
    End Enum
#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(False, True, False, True, False, False, False) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If
            If True Then 'Should be True but make False to stop Conform from cleanly dropping the telescope object (useful for retaining scopesim in memory to change flags
                Try : m_FilterWheel.Connected = False : Catch : End Try
                Try : m_FilterWheel.Dispose() : Catch : End Try
                'Try : Marshal.ReleaseComObject(m_FilterWheel) : Catch : End Try
                m_FilterWheel = Nothing
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

        Select Case g_FilterWheelProgID
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040404
                g_ExInvalidValue2 = &H80040404
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_FilterWheelProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_IFilterWheel As Object, l_DriverAccessFilterWheel As DriverAccess.FilterWheel
        Dim l_DeviceObject As Object = Nothing
        MyBase.CheckAccessibility(g_FilterWheelProgID, DeviceType.Switch)
        'Try early binding
        l_IFilterWheel = Nothing
        Try
            l_DeviceObject = CreateObject(g_FilterWheelProgID)
            l_IFilterWheel = CType(l_DeviceObject, ASCOM.Interface.IFilterWheel)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface IFilterWheel")
            Try
                l_IFilterWheel.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface IFilterWheel")
                l_IFilterWheel.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IFilterWheel")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IFilterWheel")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IFilterWheel) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IFilterWheel = Nothing
            GC.Collect()
        End Try

        Try
            l_DeviceObject = CreateObject(g_FilterWheelProgID)
            l_IFilterWheel = CType(l_DeviceObject, ASCOM.DeviceInterface.IFilterWheelV2)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface IFilterWheelV2")
            Try
                l_IFilterWheel.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface IFilterWheelV2")
                l_IFilterWheel.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IFilterWheelV2")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IFilterWheelV2")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IFilterWheel) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IFilterWheel = Nothing
            GC.Collect()
        End Try

        'Try client access toolkit
        l_DriverAccessFilterWheel = Nothing
        Try
            l_DriverAccessFilterWheel = New ASCOM.DriverAccess.FilterWheel(g_FilterWheelProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                l_DriverAccessFilterWheel.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                l_DriverAccessFilterWheel.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error conecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : l_DriverAccessFilterWheel.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DriverAccessFilterWheel) : Catch : End Try
            l_DriverAccessFilterWheel = Nothing
            GC.Collect()
        End Try
    End Sub
    Overrides Sub CreateDevice()
#If DEBUG Then
         LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.FilterWheel to get a FilterWheel object")
       m_FilterWheel = New ASCOM.DriverAccess.FilterWheel(g_FilterWheelProgID)
        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.FilterWheel to get a FilterWheel object")
            m_FilterWheel = New ASCOM.DriverAccess.FilterWheel(g_FilterWheelProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a FilterWheel object")
            m_FilterWheel = CreateObject(g_FilterWheelProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        End If
#End If
            g_Stop = False 'connected OK so clear stop flag to allow other tests to run
    End Sub
    Overrides Property Connected() As Boolean
        Get
            Connected = MyBase.Connected
        End Get
        Set(ByVal value As Boolean)
            m_FilterWheel.Connected = value
            MyBase.Connected = value
        End Set
    End Property
    Public Overrides Sub PreRunCheck()
        Dim StartTime As Date

        'Get into a consistent state
        SetStatus("FilterWheel Pre-run Check", "Wait one second for initialisation", "")
        WaitFor(1000) 'Wait for 1 second to allow any movement to start
        StartTime = Now
        Try
            Do
                SetStatus("FilterWheel Pre-run Check", "Waiting for movement to stop", Now.Subtract(StartTime).Seconds & " second(s)")
                WaitFor(SLEEP_TIME)
            Loop Until (m_FilterWheel.Position <> FWTEST_IS_MOVING) Or (Now.Subtract(StartTime).TotalSeconds > FWTEST_TIMEOUT) ' Wait until movement has stopped or 30 seconds have passed
            If m_FilterWheel.Position <> FWTEST_IS_MOVING Then LogMsg("Pre-run Check", MessageLevel.msgOK, "Filterwheel is stationary, ready to start tests")
        Catch ex As Exception
            LogMsg("Pre-run Check", MessageLevel.msgInfo, "Unable to determine that the Filterwheel is stationary")
            LogMsg("Pre-run Check", MessageLevel.msgError, "Exception: " & ex.ToString)
        End Try
        SetStatus("", "", "")

    End Sub

    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(m_FilterWheel, DeviceType.FilterWheel)
    End Sub

    Overrides Sub CheckProperties()
        Dim l_Offsets(), l_NNames, l_NOffsets, l_Offset, l_FilterNumber, l_StartFilterNumber As Integer
        Dim l_Names(), l_Name As String
        Dim i As Short
        Dim l_StartTime, l_EndTime As Date

        'FocusOffsets - Required - Read only
        Try
            l_Offsets = m_FilterWheel.FocusOffsets
            l_NOffsets = l_Offsets.Length
            If l_NOffsets = 0 Then
                LogMsg("FocusOffsets Get", MessageLevel.msgError, "Found no offset values in the returned array")
            Else
                LogMsg("FocusOffsets Get", MessageLevel.msgOK, "Found " & l_NOffsets.ToString & " filter offset values")
            End If

            l_FilterNumber = 0
            For Each l_Offset In l_Offsets
                LogMsg("FocusOffsets Get", MessageLevel.msgInfo, "Filter " & l_FilterNumber.ToString & " Offset: " & l_Offset.ToString)
                l_FilterNumber += 1
                TestStop() : If g_Stop Then Exit Try
            Next
        Catch ex As Exception
            HandleException("FocusOffsets Get", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'Names - Required - Read only
        Try
            l_Names = m_FilterWheel.Names
            l_NNames = l_Names.Length
            If l_NNames = 0 Then
                LogMsg("Names Get", MessageLevel.msgError, "Did not find any names in the returned array")
            Else
                LogMsg("Names Get", MessageLevel.msgOK, "Found " & l_NNames.ToString & " filter names")
            End If
            l_FilterNumber = 0
            For Each l_Name In l_Names
                If l_Name Is Nothing Then
                    LogMsg("Names Get", MessageLevel.msgWarning, "Filter " & l_FilterNumber.ToString & " has a value of nothing")
                Else
                    If l_Name = "" Then
                        LogMsg("Names Get", MessageLevel.msgWarning, "Filter " & l_FilterNumber.ToString & " has a value of """"")
                    Else
                        LogMsg("Names Get", MessageLevel.msgInfo, "Filter " & l_FilterNumber.ToString & " Name: " & l_Name)
                    End If
                End If
                l_FilterNumber += 1
            Next
        Catch ex As Exception
            HandleException("Names Get", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'Confirm number of array elements in filter names and filter offsets are the same
        If l_NNames = l_NOffsets Then
            LogMsg("Names Get", MessageLevel.msgOK, "Number of filter offsets and number of names are the same: " & l_NNames.ToString)
        Else
            LogMsg("Names Get", MessageLevel.msgError, "Number of filter offsets and number of names are different: " & l_NOffsets.ToString & " " & l_NNames.ToString)
        End If

        'Position - Required - Read / Write
        Select Case l_NOffsets
            Case Is <= 0
                LogMsg("Position", MessageLevel.msgWarning, "Filter position tests skipped as number of filters appears to be 0: " & l_NOffsets.ToString)
            Case Else
                Try
                    l_StartFilterNumber = m_FilterWheel.Position
                    If (l_StartFilterNumber < 0) Or (l_StartFilterNumber >= l_NOffsets) Then 'Illegal starting position provided
                        LogMsg("Position Get", MessageLevel.msgError, "Illegal filter posiiton returned: " & l_StartFilterNumber.ToString)
                    Else 'Legal value returned
                        LogMsg("Position Get", MessageLevel.msgOK, "Currently at position: " & i.ToString)
                        For i = 0 To CShort(l_NOffsets - 1)
                            Try
                                m_FilterWheel.Position = i
                                l_StartTime = Now
                                Do
                                    System.Threading.Thread.Sleep(100)
                                    TestStop()
                                Loop Until (m_FilterWheel.Position = i) Or (Now.Subtract(l_StartTime).TotalSeconds > FILTER_WHEEL_TIME_OUT) Or g_Stop
                                If g_Stop Then Exit Sub

                                l_EndTime = Now
                                If m_FilterWheel.Position = i Then 'We did get to position within the timeout period
                                    LogMsg("Position Set", MessageLevel.msgOK, "Reached position: " & i.ToString & " in: " & Format(l_EndTime.Subtract(l_StartTime).TotalSeconds, "0.0") & " seconds")
                                Else 'Timed out so give error message
                                    LogMsg("Position Set", MessageLevel.msgError, "Filter wheel did not reach specified position: " & i.ToString & " within timeout of: " & FILTER_WHEEL_TIME_OUT.ToString)
                                End If
                                WaitFor(1000) 'Pause to allow filter wheel to stabilise
                            Catch ex As Exception
                                HandleException("Position Set", MemberType.Property, Required.Mandatory, ex, "")
                            End Try
                        Next
                        Try 'Confirm that an error is correctly generated for outside range values
                            m_FilterWheel.Position = -1 'Negative position, positions should never be negative
                            LogMsg("Position Set", MessageLevel.msgError, "Failed to generate exception when selecting filter with negative filter number")
                        Catch ex As Exception
                            HandleInvalidValueExceptionAsOK("Position Set", MemberType.Property, Required.MustBeImplemented, ex, "setting position to - 1", "Correctly rejected bad position: -1")
                        End Try
                        Try 'Confirm that an error is correctly generated for outside range values
                            m_FilterWheel.Position = CShort(l_NOffsets) 'This should be 1 above the highest array element returned
                            LogMsg("Position Set", MessageLevel.msgError, "Failed to generate exception when selecting filter outside expected range")
                        Catch ex As Exception
                            HandleInvalidValueExceptionAsOK("Position Set", MemberType.Property, Required.MustBeImplemented, ex, "setting position to " & CShort(l_NOffsets), "Correctly rejected bad position: " & CShort(l_NOffsets))
                        End Try
                    End If
                Catch ex As Exception
                    HandleException("Position Get", MemberType.Property, Required.Mandatory, ex, "")
                End Try

        End Select
    End Sub
    Overrides Sub CheckPerformance()
        FilterWheelPerformanceTest(FilterWheelProperties.FocusOffsets, "FocusOffsets")
        FilterWheelPerformanceTest(FilterWheelProperties.Names, "Names")
        FilterWheelPerformanceTest(FilterWheelProperties.Position, "Position")
    End Sub
    Private Sub FilterWheelPerformanceTest(ByVal p_Type As FilterWheelProperties, ByVal p_Name As String)
        Dim l_Offsets() As Integer, l_Names() As String, l_StartFilterNumber As Integer
        Dim l_StartTime As Date, l_Count, l_LastElapsedTime, l_ElapsedTime As Double, l_Rate As Double
        Status(StatusType.staTest, "Performance Test")
        Status(StatusType.staAction, p_Name)
        Try
            l_StartTime = Now
            l_Count = 0.0
            l_LastElapsedTime = 0.0
            Do
                l_Count += 1.0
                Select Case p_Type
                    Case FilterWheelProperties.FocusOffsets
                        l_Offsets = m_FilterWheel.FocusOffsets
                    Case FilterWheelProperties.Names
                        l_Names = m_FilterWheel.Names
                    Case FilterWheelProperties.Position
                        l_StartFilterNumber = m_FilterWheel.Position
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "FilterWheelPerformanceTest: Unknown test type " & p_Type.ToString)
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
            LogMsg(p_Name, MessageLevel.msgInfo, "Unable to complete test: " & ex.Message)
        End Try
    End Sub

#End Region
End Class
