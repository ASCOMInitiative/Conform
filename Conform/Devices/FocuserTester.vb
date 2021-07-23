Friend Class FocuserTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"
    'Focuser variables
    Private m_Absolute, m_IsMoving, m_Link, m_TempComp, m_TempCompAvailable As Boolean
    Private m_MaxIncrement, m_MaxStep, m_Position, m_PositionOrg As Integer
    Private m_StepSize, m_Temperature As Double
    Private m_TempCompTrueOK, m_TempCompFalseOK As Boolean ' Variable to confirm that TempComp can be successfully set to True
    Private m_AbsolutePositionOK As Boolean = False ' Variable to confirm that absolute position can be read OK
    Private m_CanReadIsMoving As Boolean = False ' Variable to confirm that IsMoving can be read OK
    Private m_CanReadTemperature As Boolean = False ' Variable to confirm that Temperature can be read OK

    Private Const GOOD_MOVE_TOLERANCE As Integer = 2 ' ± Position tolerance within which a move will be considered to be OK
    Private Const OUT_OF_RANGE_INCREMENT As Integer = 10 ' For absolute focusers, the position delta, below 0 or above maximum steps, to test that the focuser will not move to the specified position

#If DEBUG Then
    Private m_Focuser As ASCOM.DeviceInterface.IFocuserV2
#Else
    Private m_Focuser As Object
#End If
#End Region

#Region "Enums"
    Enum FocuserPropertyMethod
        IsMoving
        Position
        Temperature
    End Enum
#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(False, True, True, False, False, True, False) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then

                If True Then 'Should be True but make False to stop Conform from cleanly dropping the focuser object (useful for retaining it in memory to change flags
                    Try : DisposeAndReleaseObject("Focuser", m_Focuser) : Catch : End Try

                    m_Focuser = Nothing
                    GC.Collect()

                End If
            End If
        End If

        MyBase.Dispose(disposing)
        disposedValue = True
    End Sub

#End Region

#Region "Code"
    Overrides Sub CheckInitialise()
        'Set the error type numbers according to the standards adopted by individual authors.
        'Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
        'messages to driver authors!

        Select Case g_FocuserProgID
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040404
                g_ExInvalidValue2 = &H80040404
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_FocuserProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_IFocuser As Object, l_DriverAccessFocuser As DriverAccess.Focuser
        Dim l_DeviceObject As Object = Nothing
        MyBase.CheckAccessibility(g_FocuserProgID, DeviceType.Focuser)

        'Try early binding to V1 interface
        l_IFocuser = Nothing
        Try
            l_DeviceObject = CreateObject(g_FocuserProgID)
            l_IFocuser = CType(l_DeviceObject, ASCOM.Interface.IFocuser)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with IFocuser interface")
            Try
                l_IFocuser.Link = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes IFocuser interface")
                l_IFocuser.Link = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuser interface")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuser interface")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IFocuser) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IFocuser = Nothing
            GC.Collect()
        End Try
        WaitFor(DEVICE_DESTROY_WAIT) 'Wait to allow device to complete destruction

        'Try early binding to V2 interface
        l_IFocuser = Nothing
        Try
            l_DeviceObject = CreateObject(g_FocuserProgID)
            l_IFocuser = CType(l_DeviceObject, ASCOM.DeviceInterface.IFocuserV2)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with IFocuserV2 interface")
            Try
                l_IFocuser.Link = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes IFocuserV2 interface")
                l_IFocuser.Link = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuserV2 interface")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuserV2 interface")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IFocuser) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IFocuser = Nothing
            GC.Collect()
        End Try
        WaitFor(DEVICE_DESTROY_WAIT) 'Wait to allow device to complete destruction

        'Try early binding to V3 interface
        l_IFocuser = Nothing
        Try
            l_DeviceObject = CreateObject(g_FocuserProgID)
            l_IFocuser = CType(l_DeviceObject, ASCOM.DeviceInterface.IFocuserV3)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with IFocuserV3 interface")
            Try
                l_IFocuser.Link = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes IFocuserV3 interface")
                l_IFocuser.Link = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuserV3 interface")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuserV3 interface")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IFocuser) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IFocuser = Nothing
            GC.Collect()
        End Try
        WaitFor(DEVICE_DESTROY_WAIT) 'Wait to allow device to complete destruction

        'Try client access toolkit
        l_DriverAccessFocuser = Nothing
        Try
            l_DriverAccessFocuser = New ASCOM.DriverAccess.Focuser(g_FocuserProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                l_DriverAccessFocuser.Link = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                l_DriverAccessFocuser.Link = False
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully disconnected using driver access toolkit")
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("AccessChecks", MessageLevel.msgDebug, "Full error: " & ex.ToString())
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.ToString)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : l_DriverAccessFocuser.Dispose() : Catch ex As Exception : LogMsg("", MessageLevel.msgDebug, "Dispose Exception: " & ex.ToString()) : End Try
            Try
                Marshal.ReleaseComObject(l_DriverAccessFocuser)
            Catch ex As ArgumentException
                ' Ignore exceptions saying that the DriverAccess variable's type must be __ComObject or derived from __ComObject.
            Catch ex As Exception
                LogMsg("", MessageLevel.msgDebug, "Marshal Exception: " & ex.ToString())
            End Try
            l_DriverAccessFocuser = Nothing
            GC.Collect()
        End Try
        WaitFor(DEVICE_DESTROY_WAIT) 'Wait to allow device to complete destruction
    End Sub
    Overrides Sub CreateDevice()
        Try
#If DEBUG Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Focuser to get a Focuser object")
            m_Focuser = New ASCOM.DriverAccess.Focuser(g_FocuserProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
            If g_Settings.UseDriverAccess Then
                LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Focuser to get a Focuser object")
                m_Focuser = New ASCOM.DriverAccess.Focuser(g_FocuserProgID)
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
            Else
                LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Focuser object")
                m_Focuser = CreateObject(g_FocuserProgID)
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
            End If
#End If
            g_Stop = False 'connected OK so clear stop flag to allow other tests to run
        Catch ex As Exception
            LogMsg("CreateDevice", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.ToString)
        End Try
    End Sub
    Overrides Property Connected() As Boolean
        Get
            LogCallToDriver("Connected", "About to get Connected property")
            Return m_Focuser.Connected
        End Get
        Set(ByVal value As Boolean)
            Try
                LogMsg("Connected", MessageLevel.msgDebug, "Setting connected state to: " & value.ToString())
                LogCallToDriver("Connected", "About to set Link property")
                m_Focuser.Link = value
                LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully changed connected state")
            Catch ex As Exception
                LogMsg("Connected", MessageLevel.msgError, "Error changing focuser connected state: " & ex.ToString)
            End Try
        End Set
    End Property

    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(m_Focuser, DeviceType.Focuser)
    End Sub

    Overrides Sub CheckProperties()
        'Absolute - Required
        Try
            LogCallToDriver("Absolute", "About to get Absolute property")
            m_Absolute = m_Focuser.Absolute
            LogMsg("Absolute", MessageLevel.msgOK, m_Absolute.ToString)
        Catch ex As Exception
            HandleException("Absolute", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'IsMoving - Required
        Try
            LogCallToDriver("IsMoving", "About to get IsMoving property")
            m_CanReadIsMoving = False
            m_IsMoving = m_Focuser.IsMoving
            If Not m_IsMoving Then
                LogMsg("IsMoving", MessageLevel.msgOK, m_IsMoving.ToString)
                m_CanReadIsMoving = True
            Else
                LogMsg("IsMoving", MessageLevel.msgError, "IsMoving is True at start of tests and it should be false")
            End If
        Catch ex As Exception
            HandleException("IsMoving", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'MaxStep - Required
        Try
            LogCallToDriver("MaxStep", "About to get MaxStep property")
            m_MaxStep = m_Focuser.MaxStep
            LogMsg("MaxStep", MessageLevel.msgOK, m_MaxStep.ToString)
        Catch ex As Exception
            HandleException("MaxStep", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'MaxIncrement - Required
        Try
            LogCallToDriver("MaxIncrement", "About to get MaxIncrement property")
            m_MaxIncrement = m_Focuser.MaxIncrement
            'Minimum value is 1, 0 or negative must be a bad value, >maxstep is a bad value
            Select Case m_MaxIncrement
                Case Is < 1
                    LogMsg("MaxIncrement", MessageLevel.msgIssue, "MaxIncrement must be at least 1, actual value: " & m_MaxIncrement)
                Case Is > m_MaxStep
                    LogMsg("MaxIncrement", MessageLevel.msgIssue, "MaxIncrement is greater than MaxStep and shouldn't be: " & m_MaxIncrement)
                Case Else
                    LogMsg("MaxIncrement", MessageLevel.msgOK, m_MaxIncrement.ToString)
            End Select
        Catch ex As Exception
            HandleException("MaxIncrement", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'Position - Optional
        If m_Absolute Then 'Absolute focuser so this property should be supported
            Try
                m_AbsolutePositionOK = False
                LogCallToDriver("Position", "About to get Position property")
                m_Position = m_Focuser.Position
                Select Case m_Position 'Check that position is a valid value
                    Case Is < 0 'Lower than lowest position
                        LogMsg("", MessageLevel.msgWarning, "Position is < 0, actual value: " & m_Position.ToString)
                    Case Is > m_MaxStep '> highest position
                        LogMsg("", MessageLevel.msgWarning, "Position is > MaxStep, actual value: " & m_Position.ToString)
                    Case Else 'Valid value
                        LogMsg("Position", MessageLevel.msgOK, m_Position.ToString)
                        m_AbsolutePositionOK = True
                End Select
            Catch ex As Exception
                HandleException("Position", MemberType.Property, Required.MustBeImplemented, ex, "Position must be implemented for an absolute focuser")
            End Try
        Else 'Relative focuser so this should raise a not supported error
            Try
                LogCallToDriver("Position", "About to get Position property")
                m_Position = m_Focuser.Position
                LogMsg("Position", MessageLevel.msgIssue, "This is a relative focuser but it didn't raise an exception for Focuser.Position")
            Catch ex As Exception
                HandleException("Position", MemberType.Property, Required.MustNotBeImplemented, ex, "Position must not be implemented for a relative focuser")
            End Try
        End If

        'StepSize - Optional
        Try
            LogCallToDriver("StepSize", "About to get StepSize property")
            m_StepSize = m_Focuser.StepSize
            Select Case m_StepSize
                Case Is <= 0.0 'Must be >0
                    LogMsg("StepSize", MessageLevel.msgIssue, "StepSize must be > 0.0, actual value: " & m_StepSize)
                Case Else
                    LogMsg("StepSize", MessageLevel.msgOK, m_StepSize.ToString)
            End Select
        Catch ex As Exception
            HandleException("StepSize", MemberType.Property, Required.Optional, ex, "")
        End Try

        'TempCompAvailable - Required
        Try
            LogCallToDriver("TempCompAvailable", "About to get TempCompAvailable property")
            m_TempCompAvailable = m_Focuser.TempCompAvailable
            LogMsg("TempCompAvailable", MessageLevel.msgOK, m_TempCompAvailable.ToString)
        Catch ex As Exception
            HandleException("StepSize", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'TempComp Read - Required
        Try
            LogCallToDriver("TempComp Read", "About to get TempComp property")
            m_TempComp = m_Focuser.TempComp
            If m_TempComp And Not m_TempCompAvailable Then
                LogMsg("TempComp Read", MessageLevel.msgIssue, "TempComp is True when TempCompAvailable is False - this should not be so")
            Else
                LogMsg("TempComp Read", MessageLevel.msgOK, m_TempComp.ToString)
            End If
        Catch ex As Exception
            HandleException("TempComp Read", MemberType.Property, Required.Mandatory, ex, "")
        End Try

        'TempComp Write - Optional
        If m_TempCompAvailable Then
            Try
                m_TempCompTrueOK = False ' Initialise to false
                m_TempCompFalseOK = False
                'Turn compensation on 
                LogCallToDriver("TempComp Write", "About to set TempComp property")
                m_Focuser.TempComp = True
                LogMsg("TempComp Write", MessageLevel.msgOK, "Successfully turned temperature compensation on")
                m_TempCompTrueOK = True ' Set to true to indicate TempComp can be successfully set to True
                'Turn compensation off
                LogCallToDriver("TempComp Write", "About to set TempComp property")
                m_Focuser.TempComp = False
                LogMsg("TempComp Write", MessageLevel.msgOK, "Successfully turned temperature compensation off")
                m_TempCompFalseOK = True
            Catch ex As Exception
                HandleException("TempComp Write", MemberType.Property, Required.MustBeImplemented, ex, "Temperature compensation is available but")
            End Try
        Else 'Should raise an exception
            Try
                LogCallToDriver("TempComp Write", "About to set TempComp property")
                m_Focuser.TempComp = True
                LogMsg("TempComp Write", MessageLevel.msgIssue, "Temperature compensation is not available but no exception was raised when TempComp was set True")
            Catch ex As Exception
                HandleException("TempComp Write", MemberType.Property, Required.MustNotBeImplemented, ex, "Temperature compensation is not available")
            End Try
        End If

        'Restore original TempComp setting if possible
        LogCallToDriver("TempComp Write", "About to set TempComp property")
        Try : m_Focuser.TempComp = m_TempComp : Catch : End Try

        'Temperature - Optional
        Try
            m_CanReadTemperature = False
            LogCallToDriver("Temperature", "About to get Temperature property")
            m_Temperature = m_Focuser.Temperature
            Select Case m_Temperature
                Case Is <= -50.0 'Probably a bad value
                    LogMsg("Temperature", MessageLevel.msgWarning, "Temperature < -50.0, - possibly an issue, actual value: " & m_Temperature.ToString)
                Case Is >= 50.0 'Probably a bad value
                    LogMsg("Temperature", MessageLevel.msgWarning, "Temperature > 50.0, - possibly an issue, actual value: " & m_Temperature.ToString)
                Case Else
                    LogMsg("Temperature", MessageLevel.msgOK, m_Temperature.ToString)
                    m_CanReadTemperature = True
            End Select
        Catch ex As Exception
            HandleException("Temperature", MemberType.Property, Required.Optional, ex, "")
        End Try
    End Sub
    Public Overrides Sub CheckMethods()
        'Halt - optional
        Try
            LogCallToDriver("Halt", "About to call Halt method")
            m_Focuser.Halt()
            LogMsg("Halt", MessageLevel.msgOK, "Focuser halted OK")
        Catch ex As COMException
            Select Case ex.ErrorCode
                Case g_ExNotImplemented, ErrorCodes.NotImplemented
                    LogMsg("Halt", MessageLevel.msgOK, "COM Exception - Halt is not supported by this focuser")
                Case Else
                    LogMsg("Halt", MessageLevel.msgError, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
            End Select
        Catch ex As Exception
            HandleException("Halt", MemberType.Method, Required.Optional, ex, "")
        End Try

        'Move - Required
        Status(StatusType.staTest, "Focuser Move")
        Try
            'Ensure that TempComp is false so that a move should be possible
            If m_TempCompFalseOK Then
                LogCallToDriver("Move - TempComp False", "About to set TempComp property")
                m_Focuser.TempComp = False
            End If
            Call MoveFocuser("Move - TempComp False")
        Catch ex As Exception
            HandleException("Move", MemberType.Method, Required.Mandatory, ex, "")
        End Try
        Status(StatusType.staTest, "")
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")

        'Move with TempComp True (if supported) - Should throw an error
        Status(StatusType.staTest, "Focuser Move")
        If m_TempCompTrueOK Then ' Can set TempComp to True so confirm that exception is thrown on Move
            Select Case g_InterfaceVersion
                Case 0, 1, 2 ' Original test method for IFocuserV2 and earlier devices
                    Try
                        LogCallToDriver("Move - TempComp True", "About to set TempComp property")
                        m_Focuser.TempComp = True
                        Call MoveFocuser("Move - TempComp True")
                        LogMsg("Move - TempComp True", MessageLevel.msgError, "TempComp is True but no exception is thrown by the Move Method - See Focuser.TempComp entry in Platform help file")
                    Catch ex As COMException
                        LogMsg("Move - TempComp True", MessageLevel.msgOK, "COM Exception correctly raised as expected")
                    Catch ex As ASCOM.InvalidOperationException
                        LogMsg("Move - TempComp True", MessageLevel.msgOK, ".NET InvalidOperation Exception correctly raised as expected")
                    Catch ex As System.InvalidOperationException
                        LogMsg("Move - TempComp True", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException")
                    Catch ex As Exception
                        LogMsg("Move - TempComp True", MessageLevel.msgError, "Unexpected .NET Exception: " & ex.Message)
                    End Try

                Case 3 ' Test method for revised IFocuserV3 behaviour introduced in Platform 6.4
                    Try
                        LogCallToDriver("Move - TempComp True V3", "About to set TempComp property")
                        m_Focuser.TempComp = True
                        Call MoveFocuser("Move - TempComp True V3")
                    Catch ex As Exception
                        HandleException("Move - TempComp True V3", MemberType.Method, Required.Mandatory, ex, "")
                    End Try

                Case Else
                    LogMsgError("Move - TempComp True", String.Format("Unknown interface version returned {0}, Move test with temperature compensation enabled skipped.", g_InterfaceVersion))
            End Select

            ' For absolute focusers, test movement to the 0 and MaxStep limits, also that the focuser will gracefully stop at the limits if commanded to move beyond them
            If m_Absolute Then
                If m_TempCompFalseOK Then
                    LogCallToDriver("Move - To 0", "About to set TempComp property")
                    m_Focuser.TempComp = False ' Set temperature compensation off
                End If

                ' Test movement to the 0 limit
                Try
                    Call MoveFocuserToPosition("Move - To 0", 0)
                    LogCallToDriver("Move - To 0", "About to get Position property")
                    Select Case m_Focuser.Position
                        Case -GOOD_MOVE_TOLERANCE To +GOOD_MOVE_TOLERANCE ' OK if within a small tolerance of expected value
                            LogCallToDriver("Move - To 0", "About to get Position property")
                            LogMsg("Move - To 0", MessageLevel.msgOK, String.Format("Moved to {0}", m_Focuser.Position))
                        Case Else
                            LogCallToDriver("Move - To 0", "About to get Position property")
                            LogMsg("Move - To 0", MessageLevel.msgInfo, String.Format("Move was within {0} counts of desired position", m_Focuser.Position))
                    End Select
                Catch ex As Exception
                    HandleException("Move - To 0", MemberType.Method, Required.Mandatory, ex, "")
                End Try

                ' Test movement below the 0 limit
                Try
                    Call MoveFocuserToPosition("Move - Below 0", -OUT_OF_RANGE_INCREMENT)
                    LogCallToDriver("Move - Below 0", "About to get Position property")
                    Select Case m_Focuser.Position
                        Case -GOOD_MOVE_TOLERANCE To +GOOD_MOVE_TOLERANCE ' OK if within a small tolerance of expected value
                            LogCallToDriver("Move - Below 0", "About to get Position property")
                            LogMsg("Move - Below 0", MessageLevel.msgOK, String.Format("Moved to {0}", m_Focuser.Position))
                        Case Else
                            LogCallToDriver("Move - Below 0", "About to get Position property")
                            LogMsg("Move - Below 0", MessageLevel.msgError, String.Format("Move was permitted below position 0: {0} ", m_Focuser.Position))
                    End Select
                Catch ex As Exception
                    HandleException("Move - Below 0", MemberType.Method, Required.Mandatory, ex, "Move should fail gracefully by just moving to position 0; it should not throw an exception")
                End Try

                ' Test movement to the MaxSteps limit
                Try
                    Call MoveFocuserToPosition("Move - To MaxStep", m_MaxStep)
                    LogCallToDriver("Move - To MaxStep", "About to get Position property")
                    Select Case m_Focuser.Position
                        Case m_MaxStep - GOOD_MOVE_TOLERANCE To m_MaxStep + GOOD_MOVE_TOLERANCE ' OK if within a small tolerance of expected value
                            LogCallToDriver("Move - To MaxStep", "About to get Position property")
                            LogMsg("Move - To MaxStep", MessageLevel.msgOK, String.Format("Moved to {0}", m_Focuser.Position))
                        Case Else
                            LogCallToDriver("Move - To MaxStep", "About to get Position property")
                            LogMsg("Move - To MaxStep", MessageLevel.msgInfo, String.Format("Move position: {0}, within {1} counts of desired position", m_Focuser.Position, m_Focuser.Position - m_MaxStep))
                    End Select
                Catch ex As Exception
                    HandleException("Move - To MaxStep", MemberType.Method, Required.Mandatory, ex, "")
                End Try

                ' Test movement above the MaxStep limit
                Try
                    Call MoveFocuserToPosition("Move - Above Maxstep", m_MaxStep + OUT_OF_RANGE_INCREMENT)
                    LogCallToDriver("Move - Above MaxStep", "About to get Position property")
                    Select Case m_Focuser.Position
                        Case m_MaxStep - GOOD_MOVE_TOLERANCE To m_MaxStep + GOOD_MOVE_TOLERANCE ' OK if within a small tolerance of expected value
                            LogCallToDriver("Move - Above MaxStep", "About to get Position property")
                            LogMsg("Move - Above Maxstep", MessageLevel.msgOK, String.Format("Moved to {0}", m_Focuser.Position))
                        Case Else
                            LogCallToDriver("Move - Above MaxStep", "About to get Position property")
                            LogMsg("Move - Above Maxstep", MessageLevel.msgError, String.Format("Moved to {0}, {1} steps from MaxStep ", m_Focuser.Position, m_Focuser.Position - m_MaxStep))
                    End Select
                Catch ex As Exception
                    HandleException("Move - Above Maxstep", MemberType.Method, Required.Mandatory, ex, "Move should fail gracefully by just moving to position MaxStep; it should not throw an exception")
                End Try

            End If

            'Restore original TempComp value
            Try : m_Focuser.TempComp = m_TempComp : Catch : End Try

            Status(StatusType.staTest, "")
            Status(StatusType.staAction, "")
            Status(StatusType.staStatus, "")
        End If

    End Sub

    Private Sub MoveFocuser(testName As String)
        If m_Absolute Then 'This is an absolute focuser so position is an absolute value
            'Save the current absolute position
            LogCallToDriver(testName, "About to get Position property")
            m_PositionOrg = m_Focuser.Position
            'Calculate an acceptable focus position
            m_Position = m_PositionOrg + CInt(m_MaxStep / 10) 'Move by 1/10 of the maximum focus distance out 
            If m_Position >= m_MaxStep Then 'That would exceed maximum travel so move inwards instead
                m_Position = m_PositionOrg - CInt(m_MaxStep / 10) 'Move by 1/10 of the maximum focus distance in
            End If
            'Apply the MaxIncrement check
            If Math.Abs(m_Position - m_PositionOrg) > m_MaxIncrement Then m_Position = m_PositionOrg + m_MaxIncrement
        Else 'This is a relative focuser so position is the relative displacement
            m_Position = CInt(m_MaxIncrement / 10)
            'Apply the MaxIncrement check
            If m_Position > m_MaxIncrement Then m_Position = m_MaxIncrement
        End If

        MoveFocuserToPosition(testName, m_Position) ' Move the focuser to the new test position within the focuser's movement range

        'Test outcome if absolute
        If m_Absolute Then
            LogCallToDriver(testName, "About to get Position property")
            Select Case m_Focuser.Position - m_Position
                Case -GOOD_MOVE_TOLERANCE To +GOOD_MOVE_TOLERANCE ' OK if within a small tolerance of expected value
                    LogMsg(testName, MessageLevel.msgOK, "Absolute move OK")
                Case Else
                    LogCallToDriver(testName, "About to get Position property")
                    LogMsg(testName, MessageLevel.msgInfo, "Move was within " & m_Focuser.Position - m_Position & " counts of desired position")
            End Select
        Else
            LogMsg(testName, MessageLevel.msgOK, "Relative move OK")
        End If

        Status(StatusType.staStatus, "")
        Status(StatusType.staAction, "Returning to original position: " & m_PositionOrg)
        LogMsg(testName, MessageLevel.msgInfo, "Returning to original position: " & m_PositionOrg)
        If m_Absolute Then
            LogCallToDriver(testName, "About to call Move method")
            m_Focuser.Move(m_PositionOrg) 'Return to original position
        Else
            LogCallToDriver(testName, "About to call Move method")
            m_Focuser.Move(-m_Position) 'Return to original position
        End If
        Status(StatusType.staStatus, "Waiting for asynchronous move to complete")

        'Wait for asynchronous move to finish
        LogCallToDriver(testName, "About to get IsMoving and Position properties repeatedly")
        Do While m_Focuser.IsMoving And (Not g_Stop)
            If m_AbsolutePositionOK Then Status(StatusType.staStatus, "Waiting for asynchronous move to complete, Position: " & m_Focuser.Position & " / " & m_PositionOrg)
            Application.DoEvents()
            WaitFor(500)
        Loop
    End Sub

    Sub MoveFocuserToPosition(testName As String, newPosition As Integer)
        Dim l_StartTime, l_EndTime As Date

        'Confirm that the focuser is not moving
        LogCallToDriver(testName, "About to get IsMoving property")
        If m_Focuser.IsMoving Then 'This is an issue as we are expecting the focuser to be not moving
            LogMsg(testName, MessageLevel.msgIssue, "Focuser is already moving before start of Move test, rest of test skipped")
        Else 'Focuser not moving so proceed with the test
            'Move the focuser
            If m_Absolute Then
                LogMsg(testName, MessageLevel.msgComment, "Moving to position: " & newPosition.ToString)
            Else
                LogMsg(testName, MessageLevel.msgComment, "Moving by: " & newPosition.ToString)
            End If

            Status(StatusType.staAction, "Moving to new position")
            l_StartTime = Now
            LogCallToDriver(testName, "About to call Move method")
            m_Focuser.Move(newPosition) ' Move the focuser
            l_EndTime = Now

            If l_EndTime.Subtract(l_StartTime).TotalMilliseconds > 1000 Then 'Move took more than 1 second so assume a synchronous call
                'Confirm that IsMoving is false
                LogCallToDriver(testName, "About to get IsMoving property")
                If m_Focuser.IsMoving Then 'This is an issue as we are expecting the focuser to be not moving
                    LogMsg(testName, MessageLevel.msgIssue, "Synchronous move expected but focuser is moving after return from Focuser.Move")
                Else
                    LogMsg(testName, MessageLevel.msgComment, "Synchronous move found")
                End If
            Else 'Move took less than 1 second so assume an asynchronous call
                Status(StatusType.staStatus, "Waiting for asynchronous move to complete")
                LogCallToDriver(testName, "About to get IsMoving and Position properties repeatedly")
                Do While (m_Focuser.IsMoving And (Not g_Stop))
                    If m_AbsolutePositionOK Then Status(StatusType.staStatus, "Waiting for asynchronous move to complete, Position: " & m_Focuser.Position & " / " & newPosition)
                    Application.DoEvents()
                    WaitFor(500)
                Loop
                LogMsg(testName, MessageLevel.msgComment, "Asynchronous move found")
            End If

        End If

    End Sub

    Overrides Sub CheckPerformance()
        'Position
        If m_AbsolutePositionOK Then
            FocuserPerformanceTest(FocuserPropertyMethod.Position, "Position")
        Else
            LogMsg("Position", MessageLevel.msgInfo, "Skipping test as property is not supported")
        End If

        'IsMoving
        If m_CanReadIsMoving Then
            FocuserPerformanceTest(FocuserPropertyMethod.IsMoving, "IsMoving")
        Else
            LogMsg("IsMoving", MessageLevel.msgInfo, "Skipping test as property is not supported")
        End If

        'Temperature
        If m_CanReadTemperature Then
            FocuserPerformanceTest(FocuserPropertyMethod.Temperature, "Temperature")
        Else
            LogMsg("Temperature", MessageLevel.msgInfo, "Skipping test as property is not supported")
        End If
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")
        Status(StatusType.staTest, "")

    End Sub

    Private Sub FocuserPerformanceTest(ByVal p_Type As FocuserPropertyMethod, ByVal p_Name As String)
        Dim l_StartTime As Date, l_Count, l_LastElapsedTime, l_ElapsedTime As Double, l_Single As Single, l_Boolean As Boolean, l_Rate As Double
        Status(StatusType.staTest, "Performance Test")
        Status(StatusType.staAction, p_Name)
        Try
            l_StartTime = Now
            l_Count = 0.0
            l_LastElapsedTime = 0.0
            Do
                l_Count += 1.0
                Select Case p_Type
                    Case FocuserPropertyMethod.IsMoving
                        l_Boolean = m_Focuser.IsMoving
                    Case FocuserPropertyMethod.Position
                        l_Single = m_Focuser.Position
                    Case FocuserPropertyMethod.Temperature
                        l_Rate = m_Focuser.Temperature
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "FocuserPerformanceTest: Unknown test type " & p_Type.ToString)
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
