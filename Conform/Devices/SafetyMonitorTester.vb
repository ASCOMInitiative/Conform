Imports ASCOM.Interface
Friend Class SafetyMonitorTester
    Inherits DeviceTesterBaseClass

#Region "Variables and Constants"
    Private m_CanIsGood, m_CanEmergencyShutdown As Boolean
    Private m_IsSafe, m_IsGood As Boolean
    Private m_Description, m_DriverInfo, m_DriverVersion As String

#If DEBUG Then
    Private m_SafetyMonitor As ASCOM.DeviceInterface.ISafetyMonitor

#Else
    Private m_SafetyMonitor As Object
#End If
#End Region

#Region "Enums"
    Private Enum RequiredProperty
        propIsSafe
    End Enum
    Private Enum PerformanceProperty
        propIsSafe
    End Enum
#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(False, True, False, False, True, True, False) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If True Then 'Should be True but make False to stop Conform from cleanly dropping the SafetyMonitor object (useful for retaining driver in memory to change flags)
                    Try : DisposeAndReleaseObject("SafetyMonitor", m_SafetyMonitor) : Catch : End Try

                    m_SafetyMonitor = Nothing
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
        Select Case g_SafetyMonitorProgID
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
        MyBase.CheckInitialise(g_SafetyMonitorProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_ISafetyMonitor As ASCOM.DeviceInterface.ISafetyMonitor, l_DriverAccessSafetyMonitor As ASCOM.DriverAccess.SafetyMonitor = Nothing
        Dim l_DeviceObject As Object = Nothing
        MyBase.CheckAccessibility(g_SafetyMonitorProgID, DeviceType.SafetyMonitor)

        'Try early binding
        l_ISafetyMonitor = Nothing
        Try
            l_DeviceObject = CreateObject(g_SafetyMonitorProgID)
            l_ISafetyMonitor = CType(l_DeviceObject, ASCOM.DeviceInterface.ISafetyMonitor)

            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using early binding to ISafetyMonitor interface")
            Try
                LogCallToDriver("AccessChecks", "About to set Connected property")
                l_ISafetyMonitor.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using early binding to ISafetyMonitor interface")
                LogCallToDriver("AccessChecks", "About to set Connected property")
                l_ISafetyMonitor.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using early binding to ISafetyMonitor interface: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using early binding to ISafetyMonitor: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_ISafetyMonitor) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_ISafetyMonitor = Nothing
            GC.Collect()
        End Try

        'Try client access toolkit
        Try
            l_DriverAccessSafetyMonitor = New ASCOM.DriverAccess.SafetyMonitor(g_SafetyMonitorProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                LogCallToDriver("AccessChecks", "About to set Connected property")
                l_DriverAccessSafetyMonitor.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                LogCallToDriver("AccessChecks", "About to set Connected property")
                l_DriverAccessSafetyMonitor.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : l_DriverAccessSafetyMonitor.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DriverAccessSafetyMonitor) : Catch : End Try
            l_DriverAccessSafetyMonitor = Nothing
            GC.Collect()
        End Try
    End Sub

    Overrides Sub CreateDevice()
#If DEBUG Then
        LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.SafetyMonitor to get a SafetyMonitor object")
        m_SafetyMonitor = New ASCOM.DriverAccess.SafetyMonitor(g_SafetyMonitorProgID)
        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.SafetyMonitor to get a SafetyMonitor object")
            m_SafetyMonitor = New ASCOM.DriverAccess.SafetyMonitor(g_SafetyMonitorProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            m_SafetyMonitor = CreateObject(g_SafetyMonitorProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        End If
#End If
            g_Stop = False 'connected OK so clear stop flag to allow other tests to run
    End Sub
    Public Overrides Sub PreConnectChecks()
        'Confirm that key properties are false when not connected
        Try
            LogCallToDriver("IsSafe", "About to get IsSafe property")
            m_IsSafe = m_SafetyMonitor.IsSafe
            If Not m_IsSafe Then
                LogMsg("IsSafe", MessageLevel.msgOK, "Reports false before connection")
            Else
                LogMsg("IsSafe", MessageLevel.msgIssue, "Reports true before connection rather than false")
            End If
        Catch ex As Exception
            LogMsg("IsSafe", MessageLevel.msgError, "Cannot confirm that IsSafe is false before connection because it threw an exception: " & ex.Message)
        End Try
    End Sub
    Overrides Property Connected() As Boolean
        Get
            LogCallToDriver("Connected", "About to get Connected property")
            Connected = m_SafetyMonitor.Connected
        End Get
        Set(ByVal value As Boolean)
            LogCallToDriver("Connected", "About to set Connected property")
            m_SafetyMonitor.Connected = value
        End Set
    End Property

    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(m_SafetyMonitor, DeviceType.SafetyMonitor)
    End Sub

    Overrides Sub CheckProperties()
        RequiredPropertiesTest(RequiredProperty.propIsSafe, "IsSafe")
    End Sub
    Overrides Sub CheckPerformance()
        Status(StatusType.staTest, "Performance")
        PerformanceTest(PerformanceProperty.propIsSafe, "IsSafe")
        Status(StatusType.staTest, "")
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")
    End Sub

    Private Sub RequiredPropertiesTest(ByVal p_Type As RequiredProperty, ByVal p_Name As String)
        Try
            Select Case p_Type
                Case RequiredProperty.propIsSafe
                    m_IsSafe = m_SafetyMonitor.IsSafe
                    LogCallToDriver("IsSafe", "About to get IsSafe property")
                    LogMsg(p_Name, MessageLevel.msgOK, m_IsSafe.ToString)
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "RequiredPropertiesTest: Unknown test type " & p_Type.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "")
        End Try
    End Sub
    Private Sub PerformanceTest(ByVal p_Type As PerformanceProperty, ByVal p_Name As String)
        Dim l_StartTime As Date, l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate As Double
        Status(StatusType.staAction, p_Name)
        Try
            l_StartTime = Now
            l_Count = 0.0
            l_LastElapsedTime = 0.0
            Do
                l_Count += 1.0
                Select Case p_Type
                    Case PerformanceProperty.propIsSafe
                        m_IsSafe = m_SafetyMonitor.IsSafe
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "PerformanceTest: Unknown test type " & p_Type.ToString)
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

#End Region
End Class
