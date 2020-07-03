'Base class from which particular device testers are derived
'Put all common elements in here
Friend Class RotatorTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"
    Const ROTATOR_WAIT_LIMIT As Double = 30.0
    Const ROTATOR_OK_TOLERANCE As Double = 1.0
    Const ROTATOR_INFO_TOLERANCE As Double = 2.0
    Const ROTATOR_POSITION_TOLERANCE As Single = 0.001 ' Degrees

    'Rotator variables
    Private m_CanReadIsMoving, canReadPosition, m_CanReadTargetPosition, m_CanReadStepSize As Boolean
    Private m_CanReverse, m_IsMoving As Boolean
    Private m_TargetPosition, m_RotatorStepSize, m_RotatorPosition, mechanicalPosition As Single
    Private m_Reverse As Boolean
    Private m_LastMoveWasAsync As Boolean
    Private canReadMechanicalPosition As Boolean

#If DEBUG Then
    Private m_Rotator As ASCOM.DeviceInterface.IRotatorV3

#Else
    Private m_Rotator As Object
#End If
#End Region

#Region "Enums"
    Private Enum RotatorPropertyMethod
        Position
        TargetPosition
        StepSize
        Move
        MoveAbsolute
        IsMoving
        MoveMechanical
    End Enum
#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(True, True, True, True, False, True, False) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If True Then 'Should be True but make False to stop Conform from cleanly dropping the rotator object (useful for retaining driver in memory to change flags)
                    Try : DisposeAndReleaseObject("Rotator", m_Rotator) : Catch : End Try

                    m_Rotator = Nothing
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

        Select Case g_RotatorProgID
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040405
                g_ExInvalidValue2 = &H80040405
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_RotatorProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_IRotator As Object, l_DriverAccessRotator As DriverAccess.Rotator
        Dim l_DeviceObject As Object = Nothing
        MyBase.CheckAccessibility(g_RotatorProgID, DeviceType.Rotator)
        'Try early binding
        l_IRotator = Nothing
        Try
            l_DeviceObject = CreateObject(g_RotatorProgID)
            l_IRotator = CType(l_DeviceObject, ASCOM.Interface.IRotator)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with interface IRotator")
            Try
                l_IRotator.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver exposes interface IRotator")
                l_IRotator.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface IRotator")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface IRotator")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IRotator) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IRotator = Nothing
            GC.Collect()
        End Try

        Try
            l_DeviceObject = CreateObject(g_RotatorProgID)
            l_IRotator = CType(l_DeviceObject, ASCOM.DeviceInterface.IRotatorV2)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with interface IRotatorV2")
            Try
                l_IRotator.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver exposes interface IRotatorV2")
                l_IRotator.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface IRotatorV2")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface IRotatorV2")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IRotator) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IRotator = Nothing
            GC.Collect()
        End Try

        'Try client access toolkit
        l_DriverAccessRotator = Nothing
        Try
            l_DriverAccessRotator = New ASCOM.DriverAccess.Rotator(g_RotatorProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                l_DriverAccessRotator.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                l_DriverAccessRotator.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : l_DriverAccessRotator.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DriverAccessRotator) : Catch : End Try
            l_DriverAccessRotator = Nothing
            GC.Collect()
        End Try

    End Sub
    Overrides Sub CreateDevice()
#If DEBUG Then
        LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Rotator to get a Rotator object")
        m_Rotator = New ASCOM.DriverAccess.Rotator(g_RotatorProgID)
        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Rotator to get a Rotator object")
            m_Rotator = New ASCOM.DriverAccess.Rotator(g_RotatorProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Rotator object")
            m_Rotator = CreateObject(g_RotatorProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        End If
#End If
        g_Stop = False 'connected OK so clear stop flag to allow other tests to run
    End Sub
    Overrides Property Connected() As Boolean
        Get
            Connected = m_Rotator.Connected
        End Get
        Set(ByVal value As Boolean)
            m_Rotator.Connected = value
        End Set
    End Property
    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(m_Rotator, DeviceType.Rotator)
    End Sub

    Public Overrides Sub ReadCanProperties()
        Try
            m_CanReverse = m_Rotator.CanReverse
            LogMsg("CanReverse", MessageLevel.msgOK, m_CanReverse.ToString)
        Catch ex As Exception
            HandleException("CanReverse", MemberType.Property, Required.Mandatory, ex, "")
        End Try

    End Sub

    Public Overrides Sub PreRunCheck()
        Dim l_Now As Date
        'Get the rotator into a standard state
        g_Stop = True
        Try : m_Rotator.Halt() : Catch : End Try 'Stop any movement
        l_Now = Now
        Try 'Confirm that rotator is not moving or wait for it to stop
            Status(StatusType.staAction, "Waiting up to " & ROTATOR_WAIT_LIMIT & " seconds for rotator to stop moving")
            Do
                WaitFor(500)
                Status(StatusType.staStatus, Now.Subtract(l_Now).TotalSeconds & "/" & ROTATOR_WAIT_LIMIT)
            Loop Until (Not m_Rotator.IsMoving) Or (Now.Subtract(l_Now).TotalSeconds > ROTATOR_WAIT_LIMIT)
            If Not m_Rotator.IsMoving Then 'Rotator is stopped so OK
                g_Stop = False 'Clear stop flag to allow other tests to run
            Else 'Report error message and don't do other tests
                LogMsg("Pre-run Check", MessageLevel.msgError, "Rotator still moving after " & ROTATOR_WAIT_LIMIT & "seconds, IsMoving stuck on?")
            End If
            LogMsg("Pre-run Check", MessageLevel.msgOK, "Rotator is stationary")
        Catch ex As Exception
            'Don't report errors at this point
        End Try

    End Sub
    Overrides Sub CheckProperties()
        'IsMoving - Optional
        Try
            m_CanReadIsMoving = False
            m_IsMoving = m_Rotator.IsMoving
            m_CanReadIsMoving = True ' Can read OK, doesn't generate an exception
            If m_IsMoving Then
                LogMsg("IsMoving", MessageLevel.msgError, "IsMoving is True before any movement has been commanded!")
                LogMsg("IsMoving", MessageLevel.msgInfo, "Further tests have been skipped")
                g_Stop = True
            Else
                LogMsg("IsMoving", MessageLevel.msgOK, m_IsMoving.ToString)
            End If
        Catch ex As Exception
            HandleException("IsMoving", MemberType.Property, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'Position - Optional
        m_RotatorPosition = RotatorPropertyTestSingle(RotatorPropertyMethod.Position, "Position", 0.0, 360.0) : If TestStop() Then Exit Sub

        'TargetPosition - Optional
        m_TargetPosition = RotatorPropertyTestSingle(RotatorPropertyMethod.TargetPosition, "TargetPosition", 0.0, 360.0) : If TestStop() Then Exit Sub

        'StepSize - Optional
        m_RotatorStepSize = RotatorPropertyTestSingle(RotatorPropertyMethod.StepSize, "StepSize", 0.0, 360.0) : If TestStop() Then Exit Sub 'Reverse Read (Boolean)

        'Reverse Read - Optional if CanReverse is False, Mandatory if CanReverse is True
        Try
            m_Reverse = m_Rotator.Reverse
            If m_CanReverse Then
                LogMsg("Reverse Read", MessageLevel.msgOK, m_Reverse.ToString)
            Else
                LogMsg("Reverse Read", MessageLevel.msgError, "CanReverse is false but no exception generated")
            End If
        Catch ex As Exception
            If m_CanReverse Then
                HandleException("Reverse Read", MemberType.Property, Required.MustBeImplemented, ex, "when CanReverse is True")
            Else
                HandleException("Reverse Read", MemberType.Property, Required.MustNotBeImplemented, ex, "when CanReverse is False")
            End If
        End Try
        If TestStop() Then Exit Sub

        'Reverse Write (Boolean)
        Try
            If m_Reverse Then 'Try and set reverse to the opposite state
                m_Rotator.Reverse = False
            Else
                m_Rotator.Reverse = True
            End If

            m_Rotator.Reverse = m_Reverse 'Restore original value

            If m_CanReverse Then
                LogMsg("Reverse Write", MessageLevel.msgOK, "Reverse state successfully changed and restored")
            Else
                LogMsg("Reverse Write", MessageLevel.msgError, "CanReverse is false but no exception generated")
            End If
        Catch ex As Exception
            If m_CanReverse Then
                HandleException("Reverse Read", MemberType.Property, Required.MustBeImplemented, ex, "when CanReverse is True")
            Else
                HandleException("Reverse Read", MemberType.Property, Required.MustNotBeImplemented, ex, "when CanReverse is False")
            End If
        End Try

        ' Test MechanicalPosition introduced in IRotatorV3
        If g_InterfaceVersion >= 3 Then

            Try
                canReadMechanicalPosition = False
                mechanicalPosition = m_Rotator.MechanicalPosition
                canReadMechanicalPosition = True 'Can read mechanical position OK, doesn't generate an exception

                'Successfully retrieved a value
                Select Case mechanicalPosition
                    Case Is < 0.0 'Lower than minimum value
                        LogMsg("MechanicalPosition", MessageLevel.msgError, "Invalid value: " & mechanicalPosition.ToString)
                    Case Is >= 360.0 'Higher than maximum value
                        LogMsg("MechanicalPosition", MessageLevel.msgError, "Invalid value: " & mechanicalPosition.ToString)
                    Case Else 'OK value
                        LogMsg("MechanicalPosition", MessageLevel.msgOK, mechanicalPosition.ToString)
                End Select

                ' For information show the sync offset, if possible, using OFFSET = SKYPOSITION - MECHANICALPOSITION
                If canReadPosition Then ' Can read synced position and mechanical position
                    LogMsgInfo("MechanicalPosition", $"Rotator sync offset: {m_RotatorPosition - mechanicalPosition}")
                End If

            Catch ex As Exception
                HandleException("MechanicalPosition", MemberType.Property, Required.Mandatory, ex, "")
            End Try

        End If

    End Sub
    Private Function RotatorPropertyTestSingle(ByVal p_Type As RotatorPropertyMethod, ByVal p_Name As String, ByVal p_Min As Single, ByVal p_Max As Single) As Single
        Try
            RotatorPropertyTestSingle = 0.0
            Select Case p_Type
                Case RotatorPropertyMethod.Position
                    canReadPosition = False
                    RotatorPropertyTestSingle = m_Rotator.Position
                    canReadPosition = True 'Can read position OK, doesn't generate an exception
                Case RotatorPropertyMethod.StepSize
                    m_CanReadStepSize = False
                    RotatorPropertyTestSingle = m_Rotator.StepSize
                    m_CanReadStepSize = True
                Case RotatorPropertyMethod.TargetPosition
                    m_CanReadTargetPosition = False
                    RotatorPropertyTestSingle = m_Rotator.TargetPosition
                    m_CanReadTargetPosition = True
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "RotatorPropertyTestSingle: Unknown test type - " & p_Type.ToString)
            End Select
            'Successfully retrieved a value
            Select Case RotatorPropertyTestSingle
                Case Is < p_Min 'Lower than minimum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & RotatorPropertyTestSingle.ToString)
                Case Is >= p_Max 'Higher than maximum value
                    LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " & RotatorPropertyTestSingle.ToString)
                Case Else 'OK value
                    LogMsg(p_Name, MessageLevel.msgOK, RotatorPropertyTestSingle.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Optional, ex, "")
        End Try
    End Function

    Overrides Sub CheckMethods()

        LogMsg("CheckMethods", MessageLevel.msgDebug, "Rotator is connected: " & m_Rotator.Connected.ToString)

        ' Halt - Optional
        Try
            m_Rotator.Halt()
            LogMsg("Halt", MessageLevel.msgOK, "Halt command successful")
        Catch ex As Exception
            HandleException("Halt", MemberType.Method, Required.Optional, ex, "")
        End Try
        If TestStop() Then Exit Sub

        ' MoveAbsolute - Optional
        RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 45.0, "") : If TestStop() Then Exit Sub
        RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 135.0, "") : If TestStop() Then Exit Sub
        RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 225.0, "") : If TestStop() Then Exit Sub
        RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 315.0, "") : If TestStop() Then Exit Sub
        RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", -405.0, "movement to large negative angle -405 degrees") : If TestStop() Then Exit Sub
        RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 405.0, "movement to large positive angle 405 degrees") : If TestStop() Then Exit Sub

        ' Move - Optional
        RelativeMoveTest(10.0) : If TestStop() Then Exit Sub
        RelativeMoveTest(40.0) : If TestStop() Then Exit Sub
        RelativeMoveTest(130.0) : If TestStop() Then Exit Sub
        RotatorMoveTest(RotatorPropertyMethod.Move, "Move", -375.0, "movement to large negative angle -375 degrees") : If TestStop() Then Exit Sub
        RotatorMoveTest(RotatorPropertyMethod.Move, "Move", 375.0, "movement to large positive angle 375 degrees") : If TestStop() Then Exit Sub

        ' Sync - Mandatory if CanSync is true - introduced in IRotatorV3
        If g_InterfaceVersion >= 3 Then ' Test for Sync method
            Try
                If canReadMechanicalPosition And canReadPosition Then ' Test new IRotaotrV3 methods

                    ' MoveMechanical
                    RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 45.0, "") : If TestStop() Then Exit Sub
                    RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 135.0, "") : If TestStop() Then Exit Sub
                    RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 225.0, "") : If TestStop() Then Exit Sub
                    RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 315.0, "") : If TestStop() Then Exit Sub
                    RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", -405.0, "movement to large negative angle -405 degrees") : If TestStop() Then Exit Sub
                    RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 405.0, "movement to large positive angle 405 degrees") : If TestStop() Then Exit Sub

                    RotatorSynctest(90.0F, 90.0F) ' Make sure that the rotator can be synced to its mechanical position
                    RotatorSynctest(120.0F, 90.0F) ' Test sync to a positive offset
                    RotatorSynctest(60.0F, 90.0F) ' Test sync to a negative offset

                    RotatorSynctest(00.0F, 00.0F) ' Test sync to zero
                    RotatorSynctest(30.0F, 00.0F) ' Test sync to a positive offset
                    RotatorSynctest(330.0F, 00.0F) ' Test sync to a negative offset that is through zero

                Else ' Just select an arbitrary sync angle

                End If

            Catch ex As Exception
                HandleException("Sync", MemberType.Method, Required.MustBeImplemented, ex, "CanSync is True")
            End Try

        End If

    End Sub

    Private Sub RotatorSynctest(SyncAngle As Single, MechanicalAngle As Single)
        Dim syncAngleDifference As Single

        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "Sync", MechanicalAngle, "") : If TestStop() Then Exit Sub

        Try
            m_Rotator.Sync(SyncAngle)
            LogMsgOK("Sync", "Synced OK")

            ' Check that Position and MechanicalPosition are now the same
            syncAngleDifference = m_Rotator.Position - SyncAngle
            If Math.Abs(syncAngleDifference) < ROTATOR_POSITION_TOLERANCE Then
                LogMsgOK("Sync", $"Rotator Position has synced to {SyncAngle} OK.")
            Else
                LogMsgIssue("Sync", $"Rotator Position is more than {ROTATOR_POSITION_TOLERANCE} different from requested position {SyncAngle}.")
            End If
        Catch ex As Exception
            HandleException("Sync", MemberType.Method, Required.MustBeImplemented, ex, "CanSync is true")
        End Try

    End Sub

    Private Sub RotatorMoveTest(ByVal p_type As RotatorPropertyMethod, ByVal p_Name As String, ByVal p_Value As Single, ByVal p_ExpectErrorMsg As String)
        Dim l_RotatorStartPosition, rotatorPosition As Single
        Dim l_OKLimit, l_PositionOffset As Double

        LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Start value, position: " & p_Value.ToString("0.000") & " " & m_Rotator.Position.ToString("0.000"))
        Try 'Move to requested position
            Select Case p_type
                Case RotatorPropertyMethod.Move
                    LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Reading rotator start position: " & canReadPosition)
                    If canReadPosition Then 'Get us to a starting point of 10 degrees
                        l_RotatorStartPosition = m_Rotator.Position
                    End If
                    LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Starting relative move")
                    m_Rotator.Move(p_Value)
                    LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Starting relative move")
                Case RotatorPropertyMethod.MoveAbsolute
                    LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Starting absolute move")
                    l_RotatorStartPosition = 0.0
                    m_Rotator.MoveAbsolute(p_Value)
                    LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Completed absolute move")
                Case RotatorPropertyMethod.MoveMechanical
                    LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Starting mechanical move")
                    l_RotatorStartPosition = 0.0
                    m_Rotator.MoveMechanical(p_Value)
                    LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Completed mechanical move")
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "RotatorMoveTest: Unknown test type - " & p_type.ToString)
            End Select
            RotatorWait(p_type, p_Name, p_Value, l_RotatorStartPosition)
            If m_LastMoveWasAsync Then 'Asynchronous move
                Select Case p_type
                    Case RotatorPropertyMethod.Move
                        If canReadPosition Then
                            LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous move successful - moved by " & p_Value & " degrees to: " & m_Rotator.Position & " degrees")
                        Else
                            LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous move successful")
                        End If
                    Case RotatorPropertyMethod.MoveMechanical
                    Case RotatorPropertyMethod.MoveAbsolute
                        If canReadPosition Then
                            LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous move successful to: " & m_Rotator.Position & " degrees")
                        Else
                            LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous move successful")
                        End If
                End Select
            Else 'Sync move
                If canReadPosition Then
                    LogMsg(p_Name, MessageLevel.msgOK, "Synchronous move successful to: " & m_Rotator.Position & " degrees")
                Else
                    LogMsg(p_Name, MessageLevel.msgOK, "Synchronous move successful")
                End If
            End If

            'Now test whether we got to where we expected to go
            If canReadPosition Then
                If m_CanReadStepSize Then
                    l_OKLimit = 1.1 * m_RotatorStepSize ' Set to 110% of step size to allow tolerance on reporting within 1 step of required location
                Else
                    l_OKLimit = ROTATOR_OK_TOLERANCE
                End If
                LogMsg(p_Name & "1", MessageLevel.msgDebug, "Position, value, start, tolerance: " & m_Rotator.Position.ToString("0.000") & " " & p_Value.ToString("0.000") & " " & l_RotatorStartPosition.ToString("0.000") & " " & l_OKLimit.ToString("0.000"))
                If m_Rotator.Position < 0.0 Then LogMsg(p_Name, MessageLevel.msgInfo, "Rotator supports angles < 0.0")
                If m_Rotator.Position > 360.0 Then LogMsg(p_Name, MessageLevel.msgInfo, "Rotator supports angles > 360.0")

                ' Get the relvant position value
                If p_type = RotatorPropertyMethod.MoveMechanical Then ' Use the MechanicalPosition property
                    rotatorPosition = m_Rotator.MechanicalPosition
                Else ' Use the Position property for all other methods
                    rotatorPosition = m_Rotator.Position
                End If
                ' Calculate the position offset from the required position
                l_PositionOffset = Math.Abs((720.0 + rotatorPosition - (p_Value + l_RotatorStartPosition)) Mod 360.0) ' Account for rotator positions that report < 0.0 or > 360.0 degrees

                If l_PositionOffset > 180.0 Then l_PositionOffset = 360.0 - l_PositionOffset ' Cope with positions that return just under the expected value

                Select Case Math.Abs(l_PositionOffset)
                    Case 0.0
                        LogMsg(p_Name, MessageLevel.msgOK, $"Rotator is at the expected position: {rotatorPosition}")
                        Exit Select
                    Case 0.0 To l_OKLimit
                        LogMsg(p_Name, MessageLevel.msgOK, $"Rotator is within {l_OKLimit.ToString("0.000")} { IIf(l_PositionOffset <= 1.0, " degree", " degrees").ToString } of the expected position: {rotatorPosition}")
                        Exit Select
                    Case 0.0 To ROTATOR_INFO_TOLERANCE
                        LogMsg(p_Name, MessageLevel.msgInfo, $"Rotator is { l_PositionOffset.ToString("0.000")} degrees from expected position: {rotatorPosition}")
                        Exit Select
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgIssue, $"Rotator is { l_PositionOffset.ToString("0.000")} degrees from expected position {rotatorPosition}, which is more than the conformance value of { ROTATOR_INFO_TOLERANCE.ToString("0.0")} degrees")
                End Select

            End If
        Catch ex As COMException
            If p_ExpectErrorMsg <> "" Then
                LogMsg(p_Name, MessageLevel.msgOK, "Exception thrown as expected on " & p_ExpectErrorMsg)
            Else
                If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
                    LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM)
                Else
                    ExTest(p_Name, ex.Message, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
                End If
            End If
        Catch ex As PropertyNotImplementedException
            If p_ExpectErrorMsg <> "" Then
                LogMsg(p_Name, MessageLevel.msgOK, "Exception thrown as expected on " & p_ExpectErrorMsg)
            Else
                LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET)
            End If
        Catch ex As Exception
            If p_ExpectErrorMsg <> "" Then
                LogMsg(p_Name, MessageLevel.msgOK, "Exception thrown as expected on " & p_ExpectErrorMsg)
            Else
                ExTest(p_Name, ex.Message, EX_NET & ex.Message)
            End If
        End Try
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")
        Status(StatusType.staTest, "")
    End Sub
    Private Sub RotatorWait(ByVal p_type As RotatorPropertyMethod, ByVal p_Name As String, ByVal p_value As Single, ByVal p_RotatorStartPosition As Single)
        LogMsg("RotatorWait", MessageLevel.msgDebug, "Entered RotatorWait")
        If m_CanReadIsMoving Then 'Can read IsMoving so test for asynchronous and synchronous behaviour
            LogMsg("RotatorWait", MessageLevel.msgDebug, "Can Read IsMoving OK")
            If m_Rotator.IsMoving Then
                LogMsg("RotatorWait", MessageLevel.msgDebug, "Rotator is moving, waiting for move to complete")
                Status(StatusType.staTest, p_Name & " test")
                Status(StatusType.staAction, "Waiting for move to complete")
                Do
                    WaitFor(500)
                    If canReadPosition Then 'Only do this if position doesn't generate an exception
                        Select Case p_type
                            Case RotatorPropertyMethod.Move
                                Status(StatusType.staStatus, System.Math.Abs(m_Rotator.Position - p_RotatorStartPosition) & "/" & p_value & " relative")
                            Case RotatorPropertyMethod.MoveAbsolute
                                Status(StatusType.staStatus, System.Math.Abs(m_Rotator.Position - p_RotatorStartPosition) & "/" & System.Math.Abs(p_value - p_RotatorStartPosition) & " absolute")
                        End Select
                    End If
                    Application.DoEvents()
                Loop Until Not m_Rotator.IsMoving
                LogMsg("RotatorWait", MessageLevel.msgDebug, "Rotator has stopped moving")
                Status(StatusType.staAction, "")
                m_LastMoveWasAsync = True
            Else
                m_LastMoveWasAsync = False
            End If
        Else 'Can only test for synchronous move
            LogMsg("RotatorWait", MessageLevel.msgDebug, "Cannot Read IsMoving")
            m_LastMoveWasAsync = False
        End If
    End Sub

    Private Sub RelativeMoveTest(ByVal p_RelativeStepSize As Single)
        Dim l_Target As Single
        If canReadPosition Then
            If m_Rotator.Position < p_RelativeStepSize Then 'Set a value that should succeed OK
                l_Target = p_RelativeStepSize
            Else
                l_Target = -p_RelativeStepSize
            End If
        Else
            l_Target = p_RelativeStepSize
        End If
        RotatorMoveTest(RotatorPropertyMethod.Move, "Move", l_Target, "") : If TestStop() Then Exit Sub
        RotatorMoveTest(RotatorPropertyMethod.Move, "Move", -l_Target, "") : If TestStop() Then Exit Sub
        'Should now be back where we started

    End Sub

    Public Overrides Sub CheckPerformance()
        'Position
        If canReadPosition Then
            RotatorPerformanceTest(RotatorPropertyMethod.Position, "Position")
        Else
            LogMsg("Position", MessageLevel.msgInfo, "Skipping test as property is not supported")
        End If

        'TargetPosition
        If m_CanReadTargetPosition Then
            RotatorPerformanceTest(RotatorPropertyMethod.TargetPosition, "TargetPosition")
        Else
            LogMsg("TargetPosition", MessageLevel.msgInfo, "Skipping test as property is not supported")
        End If

        'StepSize
        If m_CanReadStepSize Then
            RotatorPerformanceTest(RotatorPropertyMethod.StepSize, "StepSize")
        Else
            LogMsg("StepSize", MessageLevel.msgInfo, "Skipping test as property is not supported")
        End If

        'IsMoving
        If m_CanReadIsMoving Then
            RotatorPerformanceTest(RotatorPropertyMethod.IsMoving, "IsMoving")
        Else
            LogMsg("IsMoving", MessageLevel.msgInfo, "Skipping test as property is not supported")
        End If
    End Sub

    Private Sub RotatorPerformanceTest(ByVal p_Type As RotatorPropertyMethod, ByVal p_Name As String)
        Dim l_StartTime As Date, l_Count, l_LastElapsedTime, l_ElapsedTime As Double, l_Single As Single, l_Boolean As Boolean, l_Rate As Double
        Status(StatusType.staAction, p_Name)
        Try
            l_StartTime = Now
            l_Count = 0.0
            l_LastElapsedTime = 0.0
            Do
                l_Count += 1.0
                Select Case p_Type
                    Case RotatorPropertyMethod.Position
                        l_Single = m_Rotator.Position
                    Case RotatorPropertyMethod.TargetPosition
                        l_Single = m_Rotator.TargetPosition
                    Case RotatorPropertyMethod.StepSize
                        l_Single = m_Rotator.StepSize
                    Case RotatorPropertyMethod.IsMoving
                        l_Boolean = m_Rotator.IsMoving
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "RotatorPerformanceTest: Unknown test type " & p_Type.ToString)
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
