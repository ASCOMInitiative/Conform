Friend Class SwitchTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"

#If DEBUG Then ' Special values to speed up test cycles during Conform development
    Private m_Switch As ASCOM.DeviceInterface.ISwitchV2 ' Use DriverAccess for Development
#Else ' Production values so that equipment isn't damaged by rapid switch changes
    Private m_Switch As Object ' 
#End If

    Private m_InterfaceVersion, m_MaxSwitch, m_PerformanceGetSwitchName, m_PerformanceGetSwitch As Integer
    Private m_CanReadMaxSwitch As Boolean
    Private m_SwitchType As String, m_Switches As ArrayList
    Private ExtendedSwitchNumberTestRange As Integer ' Checks for usable switches either side of the expected range
    Private SWITCH_WRITE_DELAY As Integer
    Private SWITCH_READ_DELAY As Integer

    Const NUMBER_OF_SWITCH_TEST_STATES As Integer = 10
    Const BAD_SWITCH_VALUE As Double = Double.NaN ' Do not change this value, the Double.IsNaN method is used in various tests in the code below

#End Region

#Region "Enums"
    Enum SwitchPropertyMethod
        MaxSwitch
        GetSwitch
        GetSwitchName
        SwitchType
    End Enum

    Private Enum SwitchMethod
        CanWrite
        GetSwitch
        GetSwitchDescription
        GetSwitchName
        GetSwitchValue
        MaxSwitchValue
        MinSwitchValue
        SetSwitch
        SetSwitchName
        SetSwitchValue
        SwitchStep
    End Enum

#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(False, True, True, False, False, True, False) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck

        m_PerformanceGetSwitchName = Integer.MinValue 'Initialise to silly values
        m_PerformanceGetSwitch = Integer.MinValue

        SWITCH_READ_DELAY = g_Settings.SwitchReadDelay ' Get values for the two delay parameters as set by the user or the default values if not yet set
        SWITCH_WRITE_DELAY = g_Settings.SwitchWriteDelay
        ExtendedSwitchNumberTestRange = g_Settings.ExtendedSwitchNumberTestRange
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If True Then 'Should be True but make False to stop Conform from cleanly dropping the switch object (useful for retaining driver in memory to change flags)
                    Try : DisposeAndReleaseObject("Switch", m_Switch) : Catch : End Try

                    m_Switch = Nothing
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

        Select Case g_SwitchProgID
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040405
                g_ExInvalidValue2 = &H80040405
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_SwitchProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_ISwitch As Object, l_DriverAccessSwitch As DriverAccess.Switch
        Dim l_DeviceObject As Object = Nothing

        MyBase.CheckAccessibility(g_SwitchProgID, DeviceType.Switch)

        'Try early binding to V1 interface
        l_ISwitch = Nothing
        Try
#If DEBUG Then
            l_DeviceObject = CreateObject(g_SwitchProgID)
#Else
            Dim l_Type As Type
            l_Type = Type.GetTypeFromProgID(g_SwitchProgID)
            l_DeviceObject = Activator.CreateInstance(l_Type)
#End If
            l_ISwitch = CType(l_DeviceObject, ASCOM.Interface.ISwitch)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through interface ISwitch")
            Try
                l_ISwitch.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ISwitch")
                l_ISwitch.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgDebug, "Connected Exception: " & ex.ToString)
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ISwitch")
            End Try
        Catch ex As Exception
            'LogMsg("AccessChecks", MessageLevel.msgDebug, "CreateInstance Exception: " & ex.ToString)
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ISwitch")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_ISwitch) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_ISwitch = Nothing
            GC.Collect()
        End Try

        Try
            l_DeviceObject = CreateObject(g_SwitchProgID)
            l_ISwitch = CType(l_DeviceObject, ASCOM.DeviceInterface.ISwitchV2)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through interface ISwitchV2")
            Try
                l_ISwitch.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ISwitchV2")
                l_ISwitch.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ISwitchV2")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ISwitchV2")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_ISwitch) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_ISwitch = Nothing
            GC.Collect()
        End Try

        WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system")

        'Try client access toolkit
        l_DriverAccessSwitch = Nothing
        Try
            l_DriverAccessSwitch = New ASCOM.DriverAccess.Switch(g_SwitchProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                l_DriverAccessSwitch.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                l_DriverAccessSwitch.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : l_DriverAccessSwitch.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DriverAccessSwitch) : Catch : End Try
            l_DriverAccessSwitch = Nothing
            GC.Collect()
        End Try
    End Sub
    Overrides Sub CreateDevice()
#If DEBUG Then
        WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for switch to initialise")
        LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Switch to get a Switch object")
        m_Switch = New ASCOM.DriverAccess.Switch(g_SwitchProgID)
        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Switch to get a Switch object")
            m_Switch = New ASCOM.DriverAccess.Switch(g_SwitchProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Switch object")
            m_Switch = CreateObject(g_SwitchProgID)
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
            m_Switch.Connected = value
            MyBase.Connected = value
        End Set
    End Property
    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(m_Switch, DeviceType.Switch)
    End Sub

    Overrides Sub CheckProperties()
        'MaxSwitch - Mandatory
        Select Case g_InterfaceVersion
            Case 1, 2 'Original Platform 5 switch interface and ISwitchV2 have the same property
                m_MaxSwitch = SwitchPropertyTestShort(SwitchPropertyMethod.MaxSwitch, "MaxSwitch", 1, Short.MaxValue) : If TestStop() Then Exit Sub
            Case Else
                LogMsg("Switches", MessageLevel.msgError, "Unknown switch interface version: " & m_InterfaceVersion)
        End Select
    End Sub
    Overrides Sub CheckMethods()
        Dim i As Short
        Dim l_GetSwitch, l_GetSwitchOriginal, l_NewSwitchState, l_GetSwitchOK, l_SetSwitchOK, l_GetSwitchValueOK, l_SetSwitchValueMinOK, l_SetSwitchValueMaxOK, l_SwitchIsBoolean, l_SwitchCanWrite As Boolean
        Dim l_GetSwitchException, l_SetSwitchException, l_GetSwitchValueException, l_SetSwitchValueException As Exception
        Dim l_GetSwitchValue, l_GetSwitchValueOriginal, l_SwitchMinimum, l_SwitchMaximum, l_SwitchValue, l_SwitchStep, l_SwitchRange As Double
        Dim l_SwitchName, l_SwitchDescription As String

        Select Case g_InterfaceVersion
            Case 1 'Platform 5 interface v1
                If m_CanReadMaxSwitch Then 'Have got a valid value
                    Status(StatusType.staTest, "Read/Write Switches")
                    Status(StatusType.staAction, "Testing switch")
                    'Find valid GetSwitch values
                    For i = 0 To CShort(m_MaxSwitch - 1)
                        Status(StatusType.staStatus, i.ToString)
                        l_GetSwitchOK = False
                        l_SetSwitchOK = False
                        Try 'Read switch state
                            If g_Settings.DisplayMethodCalls Then LogMsg("GetSwitch " & i, MessageLevel.msgComment, String.Format("About to get switch {0}", i))
                            l_GetSwitchOriginal = m_Switch.GetSwitch(i)
                            LogMsg("GetSwitch " & i, MessageLevel.msgOK, "Found switch, state: " & l_GetSwitchOriginal.ToString)
                            l_GetSwitchOK = True
                            If i > m_MaxSwitch Then
                                LogMsg("GetSwitch " & i, MessageLevel.msgIssue, "Usable switch found above MaxSwitch!")
                            Else
                                m_PerformanceGetSwitch = i 'Save last good switch number for performance test
                            End If
                        Catch ex As Exception
                            If IsNotSetException(ex) Then
                                LogMsg("GetSwitch " & i, MessageLevel.msgInfo, "Switch read is not implemented")
                            Else
                                LogMsg("GetSwitch " & i, MessageLevel.msgInfo, "Unable to read switch: " & ex.Message)
                                LogMsg("GetSwitch " & i, MessageLevel.msgDebug, "Exception: " & ex.ToString)
                            End If
                        End Try

                        Try 'Now try to write the value
                            If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch " & i, MessageLevel.msgComment, String.Format("About to set switch {0}", i))
                            m_Switch.SetSwitch(i, Not l_GetSwitchOriginal) ' Swap the switch state
                            l_SetSwitchOK = True
                            If l_GetSwitchOK Then
                                l_NewSwitchState = m_Switch.GetSwitch(i) 'Read the new switch state to confirm that value did change
                                If l_NewSwitchState = Not l_GetSwitchOriginal Then 'State did change, test successful
                                    LogMsg("SetSwitch " & i, MessageLevel.msgOK, "Switch correctly changed state")
                                    m_Switch.SetSwitch(i, l_GetSwitchOriginal) 'Now put switch back to original state
                                Else 'State didn't change so error
                                    LogMsg("SetSwitch " & i, MessageLevel.msgIssue, "Switch did not change state, currently it is " & l_NewSwitchState.ToString)
                                End If
                            Else
                                LogMsg("SetSwitch " & i, MessageLevel.msgInfo, "You have a write only switch!")
                            End If
                            If i > m_MaxSwitch Then
                                LogMsg("SetSwitch " & i, MessageLevel.msgIssue, "Usable switch found above MaxSwitch!")
                            End If
                        Catch ex As Exception
                            If IsNotSetException(ex) Then
                                LogMsg("SetSwitch " & i, MessageLevel.msgInfo, "Switch write is not implemented")
                            Else
                                LogMsg("SetSwitch " & i, MessageLevel.msgInfo, "Unable to write to switch: " & ex.Message)
                                LogMsg("SetSwitch " & i, MessageLevel.msgDebug, "Exception: " & ex.ToString)
                            End If
                        End Try

                        Try
                            If g_Settings.DisplayMethodCalls Then LogMsg("GetSwitchName " & i, MessageLevel.msgComment, String.Format("About to get switch name {0}", i))
                            l_SwitchName = m_Switch.GetSwitchName(i)
                            If l_GetSwitchOK Or l_SetSwitchOK Then 'Name is associated with a real switch
                                If l_SwitchName = "" Then 'Switch name is a null string
                                    LogMsg("GetSwitchName " & i, MessageLevel.msgInfo, "Switch name returns null string")
                                Else 'Proper switch name returned
                                    m_PerformanceGetSwitchName = i 'Save last good name index for performance test
                                    LogMsg("GetSwitchName " & i, MessageLevel.msgOK, "Found switch, name:  " & l_SwitchName)
                                End If
                            Else 'Switch is named but has neither read nor write capability
                                If l_SwitchName = "" Then 'Switch name is a null string
                                    LogMsg("GetSwitchName " & i, MessageLevel.msgError, "Switch name returns null string but switch can neither read nor write!")
                                Else 'Proper switch name returned
                                    LogMsg("GetSwitchName " & i, MessageLevel.msgError, "Found switch, name:  " & l_SwitchName & " which can neither read nor write!")
                                End If
                            End If
                        Catch ex As Exception
                            LogMsg("GetSwitchName " & i, MessageLevel.msgDebug, "Exception: " & ex.ToString)
                        End Try

                    Next
                    Status(StatusType.staTest, "")
                    Status(StatusType.staAction, "")
                    Status(StatusType.staStatus, "")

                Else 'No valid value
                    LogMsg("SwitchCheckMethods", MessageLevel.msgError, "Skipping further tests as there is no valid value for MaxSwitch")
                End If
            Case 2 'Platform 6 interface version 2
                If m_CanReadMaxSwitch Then 'Have got a valid value
                    Status(StatusType.staTest, "Read/Write Switches")

                    CheckInaccessibleOutOfRange(SwitchMethod.CanWrite)
                    CheckInaccessibleOutOfRange(SwitchMethod.GetSwitch)
                    CheckInaccessibleOutOfRange(SwitchMethod.GetSwitchDescription)
                    CheckInaccessibleOutOfRange(SwitchMethod.GetSwitchName)
                    CheckInaccessibleOutOfRange(SwitchMethod.GetSwitchValue)
                    CheckInaccessibleOutOfRange(SwitchMethod.MaxSwitchValue)
                    CheckInaccessibleOutOfRange(SwitchMethod.MinSwitchValue)
                    CheckInaccessibleOutOfRange(SwitchMethod.SetSwitch)
                    CheckInaccessibleOutOfRange(SwitchMethod.SetSwitchValue)
                    CheckInaccessibleOutOfRange(SwitchMethod.SwitchStep)

                    'Find valid GetSwitch values
                    Status(StatusType.staAction, "Testing switch")
                    LogMsg("GetSwitchName ", MessageLevel.msgDebug, String.Format("Extended switch number test range: {0} - {1}", -ExtendedSwitchNumberTestRange, m_MaxSwitch + ExtendedSwitchNumberTestRange - 1))
                    For i = -ExtendedSwitchNumberTestRange To CShort(m_MaxSwitch + ExtendedSwitchNumberTestRange - 1)
                        Status(StatusType.staStatus, i.ToString)

                        ' Initialise status variables
                        l_GetSwitchOK = False
                        l_SetSwitchOK = False
                        l_GetSwitchValueOK = False
                        l_SetSwitchValueMinOK = False
                        l_SetSwitchValueMaxOK = False
                        l_SwitchIsBoolean = False
                        l_SwitchCanWrite = False
                        l_GetSwitchException = Nothing
                        l_GetSwitchValueException = Nothing
                        l_SetSwitchException = Nothing
                        l_SetSwitchValueException = Nothing
                        l_SwitchName = "Unknown"
                        l_SwitchDescription = "Unknown"
                        l_SwitchValue = BAD_SWITCH_VALUE
                        l_SwitchMinimum = BAD_SWITCH_VALUE
                        l_SwitchMaximum = BAD_SWITCH_VALUE
                        l_SwitchRange = BAD_SWITCH_VALUE
                        l_SwitchStep = BAD_SWITCH_VALUE

                        Try 'Read switch name to determine whether this is a valid switch
                            If g_Settings.DisplayMethodCalls Then LogMsg("GetSwitchName", MessageLevel.msgComment, String.Format("About to get switch {0} name", i))
                            l_SwitchName = m_Switch.GetSwitchName(i)
                            LogMsg("GetSwitchName ", MessageLevel.msgOK, "Found switch " & i)

                            ' Test that the switch number is in the valid range of 0..MaxSwitch-1
                            If i > (m_MaxSwitch - 1) Then
                                LogMsg("GetSwitchName ", MessageLevel.msgIssue, "Usable switch found above MaxSwitch - 1!")
                            ElseIf i < 0 Then
                                LogMsg("GetSwitchName ", MessageLevel.msgIssue, "Usable switch found below 0!")
                            Else
                                m_PerformanceGetSwitch = i 'Save last good switch number for performance test
                            End If

                            LogMsg("GetSwitchName ", MessageLevel.msgOK, "  Name: " & l_SwitchName)

                            Try 'Read switch description
                                If g_Settings.DisplayMethodCalls Then LogMsg("GetSwitchDescription", MessageLevel.msgComment, String.Format("  About to get switch {0} description", i))
                                l_SwitchDescription = m_Switch.GetSwitchDescription(i)
                                LogMsg("GetSwitchDescription ", MessageLevel.msgOK, "  Description: " & l_SwitchDescription)
                            Catch ex As Exception
                                LogMsg("GetSwitchDescription ", MessageLevel.msgError, "Mandatory parameter threw an exception: " & ex.Message)
                            End Try

                            Try 'Read switch minimum value
                                If g_Settings.DisplayMethodCalls Then LogMsg("MinSwitchValue", MessageLevel.msgComment, String.Format("  About to get switch {0} minimum value", i))
                                l_SwitchMinimum = m_Switch.MinSwitchValue(i)
                                LogMsg("MinSwitchValue ", MessageLevel.msgOK, "  Minimum: " & l_SwitchMinimum.ToString)
                            Catch ex As Exception
                                LogMsg("MinSwitchValue ", MessageLevel.msgError, "Mandatory parameter threw an exception: " & ex.Message)
                                l_SwitchMinimum = BAD_SWITCH_VALUE
                            End Try

                            Try 'Read switch maximum value
                                If g_Settings.DisplayMethodCalls Then LogMsg("MaxSwitchValue", MessageLevel.msgComment, String.Format("  About to get switch {0} maximum value", i))
                                l_SwitchMaximum = m_Switch.MaxSwitchValue(i)

                                If IsGoodValue(l_SwitchMinimum) Then ' Compare minimum and maxim values
                                    'Maximum value must be > Minimum value
                                    If l_SwitchMaximum > l_SwitchMinimum Then
                                        LogMsg("MaxSwitchValue ", MessageLevel.msgOK, "  Maximum: " & l_SwitchMaximum.ToString)
                                        l_SwitchRange = l_SwitchMaximum - l_SwitchMinimum ' Calculate the range of values the switch might take
                                    Else
                                        LogMsg("MaxSwitchValue ", MessageLevel.msgInfo, "  Maximum: " & l_SwitchMaximum.ToString)
                                        LogMsg("MaxSwitchValue ", MessageLevel.msgIssue, "MaxSwitchValue is less or equal to MinSwitchValue, it must be greater!")
                                        l_SwitchRange = BAD_SWITCH_VALUE ' Special value because the maximum or minimum values are bad
                                    End If
                                Else ' Cannot compare because of a bad minimum value
                                    LogMsg("MaxSwitchValue ", MessageLevel.msgInfo, "  Maximum: " & l_SwitchMaximum.ToString)
                                    l_SwitchRange = BAD_SWITCH_VALUE
                                    LogMsg("MaxSwitchValue ", MessageLevel.msgInfo, "  Test that switch Maximum is greater than Minimum skipped because of an error reading the Minimum value.")
                                End If
                            Catch ex As Exception
                                LogMsg("MaxSwitchValue ", MessageLevel.msgError, "Mandatory parameter threw an exception: " & ex.Message)
                                l_SwitchMaximum = BAD_SWITCH_VALUE
                                l_SwitchRange = BAD_SWITCH_VALUE
                            End Try

                            Try 'Read switch step value
                                If g_Settings.DisplayMethodCalls Then LogMsg("SwitchStep", MessageLevel.msgComment, String.Format("  About to get switch {0} step size", i))
                                l_SwitchStep = m_Switch.SwitchStep(i)
                                LogMsg("SwitchStep ", MessageLevel.msgOK, "  Step size: " & l_SwitchStep.ToString)

                                ' Step must be greater than 0
                                If l_SwitchStep > 0.0 Then
                                    LogMsg("SwitchStep ", MessageLevel.msgOK, "  Step size is greater than zero")

                                    ' Step must be less than or equal to the range of possible values
                                    If l_SwitchStep <= l_SwitchRange Then
                                        LogMsg("SwitchStep ", MessageLevel.msgOK, "  Step size is less than the range of possible values")

                                        ' Now check that the switch range is an integer multiple of the step size
                                        ' Doubles are converted to the Decimal type (which has higher precision) in order to avoid unexpected outcomes from Mod due to rounding errors
                                        Select Case Math.Abs(Decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) Mod Convert.ToDecimal(l_SwitchStep))
                                            Case 0D
                                                LogMsg("SwitchStep ", MessageLevel.msgOK, "  The switch range is an integer multiple of the step size.")
                                            Case 0D To Convert.ToDecimal(l_SwitchStep / 100)
                                                LogMsg("SwitchStep ", MessageLevel.msgInfo, "  The switch range is within 1% of being an integer multiple of the step size.")
                                            Case 0D To Convert.ToDecimal(l_SwitchStep / 10)
                                                LogMsg("SwitchStep ", MessageLevel.msgIssue, "  The switch range is not an integer multiple of the step size, but is within 10%.")
                                                'l_SwitchStep = BAD_SWITCH_VALUE
                                            Case 0D To Convert.ToDecimal(l_SwitchStep / 5)
                                                LogMsg("SwitchStep ", MessageLevel.msgIssue, "  The switch range is not an integer multiple of the step size, but is within 20%.")
                                                'l_SwitchStep = BAD_SWITCH_VALUE
                                            Case 0D To Convert.ToDecimal(l_SwitchStep / 2)
                                                LogMsg("SwitchStep ", MessageLevel.msgIssue, "  The switch range is not an integer multiple of the step size, but is within 50%.")
                                                'l_SwitchStep = BAD_SWITCH_VALUE
                                            Case Else
                                                LogMsg("SwitchStep ", MessageLevel.msgIssue, "The switch range must be an integer multiple of the step size. Remainder`: " & Decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) Mod Convert.ToDecimal(l_SwitchStep))
                                                'l_SwitchStep = BAD_SWITCH_VALUE
                                        End Select

                                    Else ' Step size is greater than the switch range
                                        LogMsg("SwitchStep ", MessageLevel.msgIssue, "Step size must be less than the range of possible values (MaxSwitchValue - MinSwitchValue")
                                        l_SwitchStep = BAD_SWITCH_VALUE
                                    End If
                                Else ' Step size is 0 or less than 
                                    LogMsg("SwitchStep ", MessageLevel.msgIssue, "Step size must be greater than zero")
                                    l_SwitchStep = BAD_SWITCH_VALUE
                                End If
                            Catch ex As Exception
                                LogMsg("SwitchStep ", MessageLevel.msgError, "Mandatory parameter threw an exception: " & ex.Message)
                            End Try

                            LogMsg("SwitchMinimum ", MessageLevel.msgDebug, l_SwitchMinimum)
                            LogMsg("SwitchMaximum ", MessageLevel.msgDebug, l_SwitchMaximum)
                            LogMsg("SwitchStep ", MessageLevel.msgDebug, l_SwitchStep)
                            LogMsg("SwitchRange ", MessageLevel.msgDebug, l_SwitchRange)

                            Try 'Read CanWrite 
                                If g_Settings.DisplayMethodCalls Then LogMsg("CanWrite", MessageLevel.msgComment, String.Format("  About to get switch {0} CanWrite status", i))
                                l_SwitchCanWrite = m_Switch.CanWrite(i)
                                LogMsg("CanWrite ", MessageLevel.msgOK, "  CanWrite: " & l_SwitchCanWrite)
                            Catch ex As Exception
                                LogMsg("CanWrite ", MessageLevel.msgError, "Mandatory parameter threw an exception: " & ex.Message)
                                LogMsg("CanWrite ", MessageLevel.msgInfo, "Assuming that CanWrite is false")
                                l_SwitchCanWrite = False ' Initialise to a default state
                            End Try

                            ' Access the Get Methods and record the outcomes
                            Try
                                If g_Settings.DisplayMethodCalls Then LogMsg("GetSwitch", MessageLevel.msgComment, String.Format("  About to call GetSwitch({0}) method", i))
                                l_GetSwitchOriginal = m_Switch.GetSwitch(i)
                                l_GetSwitchOK = True
                                LogMsg("GetSwitch ", MessageLevel.msgOK, "  " & l_GetSwitchOriginal.ToString)
                            Catch ex As Exception
                                l_GetSwitchException = ex
                                LogMsg("GetSwitch ", MessageLevel.msgDebug, "Exception: " & ex.ToString())
                                l_GetSwitchOK = False
                            End Try

                            Try
                                If g_Settings.DisplayMethodCalls Then LogMsg("GetSwitchValue", MessageLevel.msgComment, String.Format("  About to call GetSwitchValue({0}) method", i))
                                l_GetSwitchValueOriginal = m_Switch.GetSwitchValue(i)
                                l_GetSwitchValueOK = True
                                LogMsg("GetSwitchValue ", MessageLevel.msgOK, "  " & l_GetSwitchValueOriginal)
                            Catch ex As Exception
                                l_GetSwitchValueException = ex
                                LogMsg("GetSwitchValue ", MessageLevel.msgDebug, "Exception: " & ex.ToString())
                                l_GetSwitchValueOK = False
                            End Try

                            ' Now try to write to see which of these methods are available
                            If g_Settings.SwitchSet Then ' We are allowed to test write methods

                                'Try to set the two boolean values through SetSwitch
                                Try

                                    ' Try SetSwitch(False)
                                    Status(StatusType.staStatus, "Setting SetSwitch - False")
                                    If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch", MessageLevel.msgComment, String.Format("  About to call SetSwitch({0}, {1}) method", i, False))
                                    m_Switch.SetSwitch(i, False) ' Set switch false
                                    WaitFor(SWITCH_READ_DELAY)

                                    ' Check GetSwitch
                                    If l_GetSwitchOK Then ' We can read OK
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch", MessageLevel.msgComment, String.Format("  About to call GetSwitch({0}) method", i))
                                        If m_Switch.GetSwitch(i) = False Then ' We got back what we set so that's good!
                                            LogMsg("SetSwitch ", MessageLevel.msgOK, "  GetSwitch returned False after SetSwitch(False)")
                                        Else ' We didn't get back what we set
                                            LogMsg("SetSwitch ", MessageLevel.msgIssue, "  GetSwitch returned True after SetSwitch(False)")
                                        End If
                                    Else ' Issue with GetSwitch so we can't read back to test that the returned getSwitch value is correct
                                        LogMsg("SetSwitch ", MessageLevel.msgInfo, "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method")
                                    End If

                                    'Check GetSwitchValue returns the switch minimum value
                                    If l_GetSwitchValueOK And IsGoodValue(l_SwitchMinimum) Then ' We can perform this test
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch", MessageLevel.msgComment, String.Format("  About to call GetSwitchValue({0}) method", i))
                                        l_GetSwitchValue = m_Switch.GetSwitchValue(i)
                                        Select Case l_GetSwitchValue
                                            Case l_SwitchMinimum
                                                LogMsg("SetSwitch ", MessageLevel.msgOK, "  GetSwitchValue returned MINIMUM_VALUE after SetSwitch(False)")
                                            Case l_SwitchMinimum * 0.99 To l_SwitchMinimum * 1.01
                                                LogMsg("SetSwitch ", MessageLevel.msgOK, "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetSwitch(False)")
                                            Case Else
                                                LogMsg("SetSwitch ", MessageLevel.msgIssue, "  GetSwitchValue did not return MINIMUM_VALUE after SetSwitch(False): " & l_GetSwitchValue)
                                        End Select
                                        'If l_GetSwitchValue = l_SwitchMinimum Then ' We got back what we set so that's good!
                                        'Else ' We didn't get back what we set
                                        'End If
                                    Else ' Issue with GetSwitchValue GetSwitchMaximums or GetSwitchMinimum so we can't test that the returned GetSwitchValue is correct
                                        LogMsg("SetSwitch ", MessageLevel.msgInfo, "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMinimum methods")
                                    End If
                                    WaitFor(SWITCH_WRITE_DELAY)

                                    ' Try SetSwitch(True)
                                    Status(StatusType.staStatus, "Setting SetSwitch - True")
                                    If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch", MessageLevel.msgComment, String.Format("  About to call SetSwitch({0}, {1}) method", i, True))
                                    m_Switch.SetSwitch(i, True) ' Set switch true
                                    WaitFor(SWITCH_READ_DELAY)

                                    ' Check GetSwitch
                                    If l_GetSwitchOK Then ' We can read OK
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch", MessageLevel.msgComment, String.Format("  About to call GetSwitch({0}) method", i))
                                        If m_Switch.GetSwitch(i) = True Then ' We got back what we set so that's good!
                                            LogMsg("SetSwitch ", MessageLevel.msgOK, "  GetSwitch read True after SetSwitch(True)")
                                        Else ' We didn't get back what we set
                                            LogMsg("SetSwitch ", MessageLevel.msgIssue, "  GetSwitch read False after SetSwitch(True)")
                                        End If
                                    Else ' Issue with GetSwitch so we can't read back to test that the returned getSwitch value is correct
                                        LogMsg("SetSwitch ", MessageLevel.msgInfo, "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method")
                                    End If

                                    'Check GetSwitchValue returns the switch maximum value
                                    If l_GetSwitchValueOK And IsGoodValue(l_SwitchMaximum) Then ' We can perform this test
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch", MessageLevel.msgComment, String.Format("  About to call GetSwitchValue({0}) method", i))
                                        l_GetSwitchValue = m_Switch.GetSwitchValue(i)
                                        Select Case l_SwitchMaximum
                                            Case l_SwitchMaximum
                                                LogMsg("SetSwitch ", MessageLevel.msgOK, "  GetSwitchValue returned MAXIMUM_VALUE after SetSwitch(True)")
                                            Case l_SwitchMaximum * 0.99 To l_SwitchMaximum * 1.01
                                                LogMsg("SetSwitch ", MessageLevel.msgOK, "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetSwitch(True)")
                                            Case Else
                                                LogMsg("SetSwitch ", MessageLevel.msgIssue, "  GetSwitchValue did not return MAXIMUM_VALUE after SetSwitch(True): " & l_GetSwitchValue)
                                        End Select
                                        'If l_GetSwitchValue = l_SwitchMaximum Then  ' We got back what we set so that's good!
                                        'Else ' We didn't get back what we set
                                        'End If
                                    Else ' Issue with GetSwitchValue GetSwitchMaximums or GetSwitchMinimum so we can't test that the returned GetSwitchValue is correct
                                        LogMsg("SetSwitch ", MessageLevel.msgInfo, "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMaximum methods")
                                    End If
                                    WaitFor(SWITCH_WRITE_DELAY)

                                    ' Return to original state if possible,otherwise set to false
                                    If l_GetSwitchOK Then
                                        Status(StatusType.staStatus, "Returning boolean switch to its original value")
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch", MessageLevel.msgComment, String.Format("  About to call SetSwitch({0}, {1}) method", i, l_GetSwitch))
                                        m_Switch.SetSwitch(i, l_GetSwitch) ' Return to the original state
                                        WaitFor(SWITCH_WRITE_DELAY)
                                    Else
                                        Status(StatusType.staStatus, "Setting boolean switch to False")
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitch", MessageLevel.msgComment, String.Format("  About to call SetSwitch({0}, {1}) method", i, False))
                                        m_Switch.SetSwitch(i, False) ' Set to false
                                        WaitFor(SWITCH_WRITE_DELAY)
                                    End If

                                    l_SetSwitchOK = True
                                    LogMsg("SetSwitch ", MessageLevel.msgDebug, "Set value OK")
                                Catch ex As Exception ' Something went wrong with SetSwitch
                                    If l_SwitchCanWrite Then ' We can write but something went wrong
                                        LogMsg("SetSwitch ", MessageLevel.msgError, "Exception: " & ex.Message())
                                        LogMsg("SetSwitch ", MessageLevel.msgDebug, "Exception: " & ex.ToString())
                                    Else ' We can't write so we should have got a MethodNotImplementedExeption
                                        If IsMethodNotImplementedException(ex) Then ' We did get a not implemented exception and the switch is flagged as "can't write" so this is OK
                                            LogMsg("SetSwitch ", MessageLevel.msgOK, "  CanWrite is False and MethodNotImplementedException was thrown")
                                        Else ' We got some other exception and the switch is flagged as "can't write" so this is an error
                                            LogMsg("SetSwitch ", MessageLevel.msgError, "Exception: " & ex.Message())
                                            LogMsg("SetSwitch ", MessageLevel.msgDebug, "Exception: " & ex.ToString())
                                        End If
                                    End If
                                End Try

                                'Try to set the minimum and maximum values through SetSwitchValue
                                Try
                                    If IsGoodValue(l_SwitchMinimum) Then ' We have a good value to set
                                        Status(StatusType.staStatus, "Setting SetSwitchValue - MINIMUM_VALUE")
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the minimum permissible value", i, l_SwitchMinimum))
                                        m_Switch.SetSwitchValue(i, l_SwitchMinimum) ' Set switch to minimum
                                        WaitFor(SWITCH_READ_DELAY)

                                        ' Check GetSwitch
                                        If l_GetSwitchOK Then ' We can read OK
                                            If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call GetSwitch({0}) method", i))
                                            If m_Switch.GetSwitch(i) = False Then ' We got back what we set so that's good!
                                                LogMsg("SetSwitchValue", MessageLevel.msgOK, "  GetSwitch returned False after SetSwitchValue(MINIMUM_VALUE)")
                                            Else ' We didn't get back what we set
                                                LogMsg("SetSwitchValue", MessageLevel.msgIssue, "  GetSwitch returned True after SetSwitchValue(MINIMUM_VALUE)")
                                            End If
                                        Else ' Issue with GetSwitch so we can't read back to test that the returned getSwitch value is correct
                                            LogMsg("SetSwitchValue ", MessageLevel.msgInfo, "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method")
                                        End If

                                        'Check GetSwitchValue returns the switch minimum value
                                        If l_GetSwitchValueOK Then ' We can perform this test
                                            If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call GetSwitchValue({0}) method", i))
                                            l_GetSwitchValue = m_Switch.GetSwitchValue(i)
                                            Select Case l_GetSwitchValue
                                                Case l_SwitchMinimum
                                                    LogMsg("SetSwitchValue", MessageLevel.msgOK, "  GetSwitchValue returned MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE)")
                                                    l_SetSwitchValueMinOK = True
                                                Case l_SwitchMinimum * 0.99 To l_SwitchMinimum * 1.01
                                                    LogMsg("SetSwitchValue", MessageLevel.msgOK, "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE)")
                                                    l_SetSwitchValueMinOK = True
                                                Case Else
                                                    LogMsg("SetSwitchValue", MessageLevel.msgIssue, "  GetSwitchValue did not return MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE): " & l_GetSwitchValue)
                                            End Select
                                            If l_GetSwitchValue = l_SwitchMinimum Then ' We got back what we set so that's good!
                                            Else ' We didn't get back what we set
                                            End If
                                        Else ' Issue with GetSwitchValue GetSwitchMaximums or GetSwitchMinimum so we can't test that the returned GetSwitchValue is correct
                                            LogMsg("SetSwitchValue ", MessageLevel.msgInfo, "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method")
                                        End If
                                        WaitFor(SWITCH_WRITE_DELAY)

                                        ' Now try a value below minimum
                                        Try
                                            If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an invalid low value", i, l_SwitchMinimum - 1.0))
                                            m_Switch.SetSwitchValue(i, l_SwitchMinimum - 1.0)
                                            LogMsg("SetSwitchValue", MessageLevel.msgIssue, "Switch did not throw an exception when a value below SwitchMinimum was set: " & l_SwitchMinimum - 1.0)
                                        Catch ex As Exception
                                            HandleInvalidValueExceptionAsOK("SetSwitchValue", MemberType.Method, Required.Mandatory, ex, "when setting a value below SwitchMinimum - " & l_SwitchMinimum - 1.0, "  Switch threw an InvalidOperationException when a value below SwitchMinimum was set: " & l_SwitchMinimum - 1.0)
                                        End Try
                                        WaitFor(SWITCH_WRITE_DELAY)
                                    Else ' Minimum value is bad so skip the SetSwitchValue test
                                        LogMsg("SetSwitchValue ", MessageLevel.msgInfo, "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim")
                                    End If

                                    ' Try SetSwitchValue(MAXIMUM_VALUE)
                                    If IsGoodValue(l_SwitchMaximum) Then ' We have a good value to set
                                        Status(StatusType.staStatus, "Setting SetSwitchValue - MAXIMUM_VALUE")
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the maximum permissible value", i, l_SwitchMaximum))
                                        m_Switch.SetSwitchValue(i, l_SwitchMaximum) ' Set switch to minimum
                                        WaitFor(SWITCH_READ_DELAY)

                                        ' Check GetSwitch
                                        If l_GetSwitchOK Then ' We can read OK
                                            If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call GetSwitch({0}) method", i))
                                            If m_Switch.GetSwitch(i) = True Then ' We got back what we set so that's good!
                                                LogMsg("SetSwitchValue ", MessageLevel.msgOK, "  GetSwitch returned True after SetSwitchValue(MAXIMUM_VALUE)")
                                            Else ' We didn't get back what we set
                                                LogMsg("SetSwitchValue ", MessageLevel.msgIssue, "  GetSwitch returned False after SetSwitchValue(MAXIMUM_VALUE)")
                                            End If
                                        Else ' Issue with GetSwitch so we can't read back to test that the returned getSwitch value is correct
                                            LogMsg("SetSwitchValue ", MessageLevel.msgInfo, "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method")
                                        End If

                                        'Check GetSwitchValue returns the switch maximum value
                                        If l_GetSwitchValueOK Then ' We can perform this test
                                            If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call GetSwitchValue({0}) method", i))
                                            l_GetSwitchValue = m_Switch.GetSwitchValue(i)
                                            Select Case l_SwitchMaximum
                                                Case l_SwitchMaximum
                                                    LogMsg("SetSwitchValue ", MessageLevel.msgOK, "  GetSwitchValue returned MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE)")
                                                    l_SetSwitchValueMaxOK = True
                                                Case l_SwitchMaximum * 0.99 To l_SwitchMaximum * 1.01
                                                    LogMsg("SetSwitchValue ", MessageLevel.msgOK, "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE)")
                                                Case Else
                                                    LogMsg("SetSwitchValue ", MessageLevel.msgIssue, "  GetSwitchValue did not return MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE): " & l_GetSwitchValue)
                                            End Select
                                            If l_GetSwitchValue = l_SwitchMaximum Then ' We got back what we set so that's good!
                                            Else ' We didn't get back what we set
                                            End If
                                        Else ' Issue with GetSwitchValue GetSwitchMaximums or GetSwitchMinimum so we can't test that the returned GetSwitchValue is correct
                                            LogMsg("SetSwitchValue ", MessageLevel.msgInfo, "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method")
                                        End If
                                        WaitFor(SWITCH_WRITE_DELAY)

                                        ' Now try a value above maximum
                                        Try
                                            If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an invalid high value", i, l_SwitchMaximum + 1.0))
                                            m_Switch.SetSwitchValue(i, l_SwitchMaximum + 1.0)
                                            LogMsg("SetSwitchValue", MessageLevel.msgIssue, "Switch did not throw an exception when a value above SwitchMaximum was set: " & l_SwitchMaximum + 1.0)
                                        Catch ex As Exception
                                            HandleInvalidValueExceptionAsOK("SetSwitchValue", MemberType.Method, Required.Mandatory, ex, "when setting a value above SwitchMaximum was set: " & l_SwitchMaximum + 1.0, "  Switch threw an InvalidOperationException when a value above SwitchMaximum was set: " & l_SwitchMaximum + 1.0)
                                        End Try
                                        WaitFor(SWITCH_WRITE_DELAY)

                                    Else ' Minimum value is bad so skip the SetSwitchValue test
                                        LogMsg("SetSwitchValue ", MessageLevel.msgInfo, "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim")
                                    End If

                                    ' Test some positions of the multi-state switch between the minimum and maximum values
                                    If l_GetSwitchValueOK And l_SetSwitchValueMinOK And l_SetSwitchValueMaxOK And IsGoodValue(l_SwitchRange) And IsGoodValue(l_SwitchStep) Then

                                        TestSetSwitchValue(i, 0.0, l_SwitchMinimum, l_SwitchMaximum, l_SwitchRange, l_SwitchStep)
                                        TestSetSwitchValue(i, 0.25, l_SwitchMinimum, l_SwitchMaximum, l_SwitchRange, l_SwitchStep)
                                        TestSetSwitchValue(i, 0.5, l_SwitchMinimum, l_SwitchMaximum, l_SwitchRange, l_SwitchStep)
                                        TestSetSwitchValue(i, 0.75, l_SwitchMinimum, l_SwitchMaximum, l_SwitchRange, l_SwitchStep)

                                    Else
                                        LogMsg("SetSwitchValue ", MessageLevel.msgWarning, "Skipping multi state tests because of earlier errors")
                                        LogMsg("GetSwitchValueOK ", MessageLevel.msgDebug, l_GetSwitchValueOK)
                                        LogMsg("SetSwitchValueMinOK ", MessageLevel.msgDebug, l_SetSwitchValueMinOK)
                                        LogMsg("SetSwitchValueMaxOK ", MessageLevel.msgDebug, l_SetSwitchValueMaxOK)
                                        LogMsg("SwitchRange ", MessageLevel.msgDebug, l_SwitchRange)
                                        LogMsg("SwitchStep ", MessageLevel.msgDebug, l_SwitchStep.ToString)
                                    End If

                                    ' Return to original state if possible,otherwise set to false
                                    If l_GetSwitchValueOK Then
                                        Status(StatusType.staStatus, "Returning switch to its original value")
                                        If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call SetSwitchValue({0}, {1}), attempting to restore pre-test value", i, l_GetSwitchValueOriginal))
                                        m_Switch.SetSwitchValue(i, l_GetSwitchValueOriginal) ' Return to the original state
                                        LogMsg("SetSwitchValue ", MessageLevel.msgOK, "  Switch has been reset to its original state")
                                        WaitFor(SWITCH_WRITE_DELAY)
                                    Else
                                        If IsGoodValue(l_SwitchMinimum) And IsGoodValue(l_SwitchMaximum) Then ' We have valid minimum and maxim values so set the switch accordingly
                                            Status(StatusType.staStatus, "Setting switch to half its minimum to maximum range")
                                            If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the value to its mid-point", i, (l_SwitchMaximum - l_SwitchMinimum) / 2.0))
                                            m_Switch.SetSwitchValue(i, (l_SwitchMaximum - l_SwitchMinimum) / 2.0) ' Return to the half way state
                                            LogMsg("SetSwitchValue ", MessageLevel.msgOK, "  Switch has been reset to half its range")
                                            WaitFor(SWITCH_WRITE_DELAY)
                                        Else ' Can't do anything because the switch driver is too broken
                                            LogMsg("SetSwitchValue ", MessageLevel.msgWarning, "Switch can not be returned to its default state because of issues with GetSwitchValue, GetSwitchMinimum or GetSwitchMaximum")
                                        End If
                                    End If

                                Catch ex As Exception ' Something went wrong with SetSwitch
                                    If l_SwitchCanWrite Then ' We can write but something went wrong
                                        LogMsg("SetSwitchValue ", MessageLevel.msgError, "Exception: " & ex.Message())
                                        LogMsg("SetSwitchValue ", MessageLevel.msgDebug, "Exception: " & ex.ToString())
                                    Else ' We can't write so we should have got a MethodNotImplementedExeption
                                        If IsMethodNotImplementedException(ex) Then ' We did get a not implemented exception and the switch is flagged as "can't write" so this is OK
                                            LogMsg("SetSwitchValue ", MessageLevel.msgOK, "  CanWrite is False and MethodNotImplementedException was thrown")
                                        Else ' We got some other exception and the switch is flagged as "can't write" so this is an error
                                            LogMsg("SetSwitchValue ", MessageLevel.msgError, "Exception: " & ex.Message())
                                            LogMsg("SetSwitchValue ", MessageLevel.msgDebug, "Exception: " & ex.ToString())
                                        End If
                                    End If
                                End Try
                            Else
                                LogMsg("SetSwitch", MessageLevel.msgInfo, "  All write tests have been skipped because the ""Set Switches"" checkbox is unchecked")
                            End If

                            LogMsg("", MessageLevel.msgAlways, "")
                        Catch ex As Exception ' GetSwitchName threw an exception
                            If (i >= 0) And (i < m_MaxSwitch) Then
                                LogMsg("GetSwitchName ", MessageLevel.msgError, "Mandatory method GetSwitchName threw an exception: " & ex.ToString())
                            Else
                                LogMsg("GetSwitchName ", MessageLevel.msgDebug, "Exception: " & ex.ToString())
                            End If
                        End Try

                        If TestStop() Then Exit Sub
                    Next
                    Status(StatusType.staTest, "")
                    Status(StatusType.staAction, "")
                    Status(StatusType.staStatus, "")

                Else 'No valid value
                    LogMsg("SwitchCheckMethods", MessageLevel.msgInfo, "Skipping further tests as there is no valid value for MaxSwitch")
                End If

        End Select
    End Sub

    Overrides Sub CheckPerformance()
        'MaxSwitch
        If m_CanReadMaxSwitch Then 'Have got a valid value
            SwitchPerformanceTest(SwitchPropertyMethod.MaxSwitch, "MaxSwitch")
        Else
            LogMsg("MaxSwitch", MessageLevel.msgInfo, "Test skipped as unable to read value")
        End If
        'GetSwitch
        If CBool(m_PerformanceGetSwitch) Then 'Have got a valid value
            SwitchPerformanceTest(SwitchPropertyMethod.GetSwitch, "GetSwitch")
        Else
            LogMsg("GetSwitch", MessageLevel.msgInfo, "Test skipped as unable to read value")
        End If
        'GetSwitchName
        If CBool(m_PerformanceGetSwitchName) Then 'Have got a valid value
            SwitchPerformanceTest(SwitchPropertyMethod.GetSwitchName, "GetSwitchName")
        Else
            LogMsg("GetSwitchName", MessageLevel.msgInfo, "Test skipped as unable to read value")
        End If
        Status(StatusType.staTest, "")
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")
    End Sub

#End Region

#Region "Support code"
    Private Function SwitchPropertyTestString(ByVal p_Type As SwitchPropertyMethod, ByVal p_Name As String, ByVal p_MaxLength As Integer) As String
        SwitchPropertyTestString = ""
        Try
            Select Case p_Type
                Case SwitchPropertyMethod.SwitchType
                    SwitchPropertyTestString = m_Switch.SwitchType
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "SwitchPropertyTestString: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case SwitchPropertyTestString
                Case Is = ""
                    LogMsg(p_Name, MessageLevel.msgError, "Null string returned")
                Case Else
                    If Len(SwitchPropertyTestString) <= p_MaxLength Then
                        LogMsg(p_Name, MessageLevel.msgOK, SwitchPropertyTestString)
                    Else
                        LogMsg(p_Name, MessageLevel.msgError, "String exceeds " & p_MaxLength & " characters maximum length - " & SwitchPropertyTestString)
                    End If
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "")
        End Try
    End Function

    Private Function SwitchPropertyTestShort(ByVal p_Type As SwitchPropertyMethod, ByVal p_Name As String, ByVal p_Min As Short, ByVal p_Max As Short) As Short
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg(p_Name, MessageLevel.msgComment, String.Format("About to get property {0}", p_Name))
            SwitchPropertyTestShort = 0
            Select Case p_Type
                Case SwitchPropertyMethod.MaxSwitch
                    m_CanReadMaxSwitch = False
                    SwitchPropertyTestShort = m_Switch.MaxSwitch
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "SwitchPropertyTestInteger: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case SwitchPropertyTestShort
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & SwitchPropertyTestShort.ToString)
                Case Is > p_Max 'Higher than maximum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & SwitchPropertyTestShort.ToString)
                Case Else 'OK value
                    m_CanReadMaxSwitch = True 'A valid value has been found
                    LogMsg(p_Name, MessageLevel.msgOK, SwitchPropertyTestShort.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "")
        End Try
    End Function

    Private Sub SwitchPerformanceTest(ByVal p_Type As SwitchPropertyMethod, ByVal p_Name As String)
        Dim l_StartTime As Date, l_Count, l_LastElapsedTime, l_ElapsedTime As Double, l_String As String, l_Boolean As Boolean, l_Rate As Double, l_Short As Short
        Status(StatusType.staTest, "Performance Testing")
        Status(StatusType.staAction, p_Name)
        Try
            l_StartTime = Now
            l_Count = 0.0
            l_LastElapsedTime = 0.0
            Do
                l_Count += 1.0
                Select Case p_Type
                    Case SwitchPropertyMethod.MaxSwitch
                        l_Short = m_Switch.MaxSwitch
                    Case SwitchPropertyMethod.GetSwitch
                        l_Boolean = m_Switch.GetSwitch(CShort(m_PerformanceGetSwitch))
                    Case SwitchPropertyMethod.GetSwitchName
                        l_String = m_Switch.GetSwitchName(CShort(m_PerformanceGetSwitchName))
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "SwitchPerformanceTest: Unknown test type " & p_Type.ToString)
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

    ''' <summary>
    ''' Tests whether a double has a good value or the NaN bad value indicator
    ''' </summary>
    ''' <param name="value">Variable to be tested</param>
    ''' <returns>Returns True if the variable has a good value, otherwise returns False</returns>
    ''' <remarks></remarks>
    Private Function IsGoodValue(value As Double) As Boolean
        Return Not Double.IsNaN(value)
    End Function

    ''' <summary>
    ''' Test that GetSwitchValue returns the same value as was set with SetSwitchValue
    ''' </summary>
    ''' <param name="i">Switch number</param>
    ''' <param name="Offset">Offset size as a percentage of switch step</param>
    ''' <param name="SwitchMinimum">Switch minimum value</param>
    ''' <param name="SwitchRange">Switch value range</param>
    ''' <param name="SwitchStep">Size of each switch step</param>
    ''' <remarks></remarks>
    Private Sub TestSetSwitchValue(i As Integer, Offset As Double, SwitchMinimum As Double, SwitchMaximum As Double, SwitchRange As Double, SwitchStep As Double)
        Dim l_MultiStateStepSize, TestValue2, l_SwitchValue As Double
        Dim l_SetSwitchValueOK As Boolean
        Dim msgLevel0, msgLevel1, msgLevel2 As MessageLevel
        Dim l_MultiStateNumberOfSteps As Integer

        ' Test the switch at the calculated positions
        Try

            If (((SwitchMaximum - SwitchMinimum) / SwitchStep) + 1) >= NUMBER_OF_SWITCH_TEST_STATES Then
                l_MultiStateStepSize = SwitchRange / NUMBER_OF_SWITCH_TEST_STATES
                l_MultiStateNumberOfSteps = NUMBER_OF_SWITCH_TEST_STATES
            Else
                l_MultiStateStepSize = SwitchStep ' Use the specified switch step size
                l_MultiStateNumberOfSteps = Math.Floor(SwitchRange / SwitchStep)
            End If
            LogMsg("MultiStateStepSize", MessageLevel.msgDebug, l_MultiStateStepSize)
            LogMsg("MultiStateNumberOfSteps", MessageLevel.msgDebug, l_MultiStateNumberOfSteps)

            If Offset = 0.0 Then ' An offset of 0 must be tested
                msgLevel0 = MessageLevel.msgOK
                msgLevel1 = MessageLevel.msgInfo
                msgLevel2 = MessageLevel.msgIssue
            Else ' Offsets other than 0 are only for information and issues will not be reported
                msgLevel0 = MessageLevel.msgInfo
                msgLevel1 = MessageLevel.msgInfo
                msgLevel2 = MessageLevel.msgInfo
            End If

            l_SetSwitchValueOK = True
            LogMsg("SetSwitchValue", MessageLevel.msgInfo, "  Testing with steps that are " & Offset.ToString("P0") & " offset from integer SwitchStep values")

            For TestValue As Double = SwitchMinimum To SwitchMinimum + l_MultiStateStepSize * l_MultiStateNumberOfSteps Step l_MultiStateStepSize

                ' Round the test value to the nearest lowest switch step
                If TestValue = SwitchMinimum Then
                    TestValue2 = SwitchMinimum + Offset * SwitchStep
                Else
                    TestValue2 = (Math.Round((TestValue - SwitchMinimum) / SwitchStep) * SwitchStep) + SwitchMinimum + Offset * SwitchStep
                End If

                If TestValue2 <= SwitchMaximum Then ' Only test if the new value is within the switches range
                    Status(StatusType.staStatus, "Setting multi-state switch - " & TestValue2)
                    If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an intermediate value", i, TestValue2))
                    m_Switch.SetSwitchValue(i, TestValue2) ' Set the required switch value
                    WaitFor(SWITCH_READ_DELAY)
                    If g_Settings.DisplayMethodCalls Then LogMsg("SetSwitchValue", MessageLevel.msgComment, String.Format("  About to call GetSwitchValue({0})", i))
                    l_SwitchValue = m_Switch.GetSwitchValue(i) ' Read back the switch value 

                    Select Case Math.Abs(l_SwitchValue - TestValue2)
                        Case 0.0
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel0, "  Set and read match: " & TestValue2)
                        Case 0.0 To SwitchStep * 0.01
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel1, "   Set/Read differ by up to 1% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                        Case 0.0 To SwitchStep * 0.1
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "   Set/Read differ by 1-10% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 0.2
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 10-20% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 0.3
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 20-30% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 0.4
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 30-40% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 0.5
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 40-50% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 0.6
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 50-60% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 0.7
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 60-70% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 0.8
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 70-80% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 0.9
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 80-90% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case 0.0 To SwitchStep * 1.0
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 90-100% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                        Case Else
                            LogMsg("SetSwitchValue " & "Offset: " & Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by >100% of SwitchStep. Set: " & TestValue2 & ", Read: " & l_SwitchValue)
                            l_SetSwitchValueOK = False
                    End Select

                    WaitFor(SWITCH_WRITE_DELAY)
                End If
                If TestStop() Then Exit Sub
            Next
        Catch ex As Exception
            HandleException("SetSwitchValue", MemberType.Property, Required.Mandatory, ex, "")
        End Try

    End Sub

    Private Sub CheckInaccessibleOutOfRange(method As SwitchMethod)
        Const LOW_TEST_VALUE As Short = -1
        Const HIGH_TEST_VALUE As Short = 1
        Dim BoolValue As Boolean, StringValue As String, DoubleValue As Double

        If m_MaxSwitch > 0 Then

            ' Try a value below 0
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg(method.ToString(), MessageLevel.msgComment, String.Format("About to call {0} with invalid low value for switch number: {1} for ", method.ToString(), LOW_TEST_VALUE))
                Select Case method
                    Case SwitchMethod.CanWrite
                        BoolValue = m_Switch.CanWrite(LOW_TEST_VALUE)
                    Case SwitchMethod.GetSwitch
                        BoolValue = m_Switch.GetSwitch(LOW_TEST_VALUE)
                    Case SwitchMethod.GetSwitchDescription
                        StringValue = m_Switch.GetSwitchDescription(LOW_TEST_VALUE)
                    Case SwitchMethod.GetSwitchName
                        StringValue = m_Switch.GetSwitchName(LOW_TEST_VALUE)
                    Case SwitchMethod.GetSwitchValue
                        DoubleValue = m_Switch.GetSwitchValue(LOW_TEST_VALUE)
                    Case SwitchMethod.MaxSwitchValue
                        DoubleValue = m_Switch.MaxSwitchValue(LOW_TEST_VALUE)
                    Case SwitchMethod.MinSwitchValue
                        DoubleValue = m_Switch.MinSwitchValue(LOW_TEST_VALUE)
                    Case SwitchMethod.SetSwitch
                        m_Switch.SetSwitchValue(LOW_TEST_VALUE, False)
                    Case SwitchMethod.SetSwitchName
                        ' Ignoring this because it is an optional parameter
                    Case SwitchMethod.SetSwitchValue
                        m_Switch.SetSwitchValue(LOW_TEST_VALUE, 0.0)
                    Case SwitchMethod.SwitchStep
                        DoubleValue = m_Switch.SwitchStep(LOW_TEST_VALUE)
                    Case Else
                        MsgBox("Unknown value of SwitchMethod Enum: " & method.ToString, MsgBoxStyle.Critical, "Conform Application Error")
                End Select
                LogMsg("SwitchNumber", MessageLevel.msgIssue, "Switch did not throw an exception when a switch ID below 0 was used in method: " & method.ToString)
            Catch ex As Exception
                HandleInvalidValueExceptionAsOK("SwitchNumber", MemberType.Property, Required.MustBeImplemented, ex, "when a switch ID below 0 was used in method: " & method.ToString, "Switch device threw an InvalidOperationException when a switch ID below 0 was used in method: " & method.ToString)
            End Try

            ' Try a value above MaxSwitch
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg(method.ToString(), MessageLevel.msgComment, String.Format("About to call {0} with invalid high value for switch number: {1} for ", method.ToString(), m_MaxSwitch + HIGH_TEST_VALUE))
                Select Case method
                    Case SwitchMethod.CanWrite
                        BoolValue = m_Switch.CanWrite(m_MaxSwitch + HIGH_TEST_VALUE)
                    Case SwitchMethod.GetSwitch
                        BoolValue = m_Switch.GetSwitch(m_MaxSwitch + HIGH_TEST_VALUE)
                    Case SwitchMethod.GetSwitchDescription
                        StringValue = m_Switch.GetSwitchDescription(m_MaxSwitch + HIGH_TEST_VALUE)
                    Case SwitchMethod.GetSwitchName
                        StringValue = m_Switch.GetSwitchName(m_MaxSwitch + HIGH_TEST_VALUE)
                    Case SwitchMethod.GetSwitchValue
                        DoubleValue = m_Switch.GetSwitchValue(m_MaxSwitch + HIGH_TEST_VALUE)
                    Case SwitchMethod.MaxSwitchValue
                        DoubleValue = m_Switch.MaxSwitchValue(m_MaxSwitch + HIGH_TEST_VALUE)
                    Case SwitchMethod.MinSwitchValue
                        DoubleValue = m_Switch.MinSwitchValue(m_MaxSwitch + HIGH_TEST_VALUE)
                    Case SwitchMethod.SetSwitch
                        m_Switch.SetSwitchValue(m_MaxSwitch + HIGH_TEST_VALUE, False)
                    Case SwitchMethod.SetSwitchName
                        ' Ignoring this because it is an optional parameter
                    Case SwitchMethod.SetSwitchValue
                        m_Switch.SetSwitchValue(m_MaxSwitch + HIGH_TEST_VALUE, 0.0)
                    Case SwitchMethod.SwitchStep
                        DoubleValue = m_Switch.SwitchStep(m_MaxSwitch + HIGH_TEST_VALUE)
                    Case Else
                        MsgBox("Unknown value of SwitchMethod Enum: " & method.ToString, MsgBoxStyle.Critical, "Conform Application Error")
                End Select
                LogMsg("SwitchNumber", MessageLevel.msgIssue, "Switch did not throw an exception when a switch ID above MaxSwitch was used in method: " & method.ToString)
            Catch ex As Exception
                HandleInvalidValueExceptionAsOK("SwitchNumber", MemberType.Property, Required.MustBeImplemented, ex, "when a switch ID above MaxSwitch was used in method: " & method.ToString, "Switch device threw an InvalidOperationException when a switch ID above MaxSwitch was used in method: " & method.ToString)
            End Try

        Else
            LogMsg("SwitchNumber", MessageLevel.msgInfo, "Skipping range tests because MaxSwitch cannot be read")
        End If
    End Sub

#End Region

End Class
