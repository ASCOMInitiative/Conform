Imports ASCOM.DeviceInterface
Friend Class DomeTester
    Inherits DeviceTesterBaseClass
#Region "Variables and Constants"
    'These constants have been moved to the applications setting object so that they can be changed by the user, the same default values have been retained
    'Const DOME_SHUTTER_TIMEOUT As Integer = 60 'Timeout for dome to open or close shutter
    'Const DOME_AZIMUTH_TIMEOUT As Integer = 180 'Timeout for dome to move to azimuth
    'Const DOME_ALTITUDE_TIMEOUT As Integer = 120 'Timeout for dome to move to altitude
    'Const DOME_POST_COMMAND_WAIT_TIME As Integer = 5000 '

    Const DOME_SYNC_OFFSET As Double = 45.0 'Amount to offset the azimuth when testing ability to sync
    Const DOME_ILLEGAL_ALTITUDE_LOW As Double = -10.0 'Illegal value to test dome driver exception generation
    Const DOME_ILLEGAL_ALTITUDE_HIGH As Double = 100.0 'Illegal value to test dome driver exception generation
    Const DOME_ILLEGAL_AZIMUTH_LOW As Double = -10.0 'Illegal value to test dome driver exception generation
    Const DOME_ILLEGAL_AZIMUTH_HIGH As Double = 370.0 'Illegal value to test dome driver exception generation

    'Dome variables
    Private m_CanSetAltitude, m_CanSetAzimuth, m_CanSetShutter, m_CanSlave, m_CanSyncAzimuth, m_Slaved As Boolean
    Private m_ShutterStatus As ShutterState
    Private m_CanReadAltitude, m_CanReadAtPark, m_CanReadAtHome, m_CanReadSlewing, m_CanReadSlaved, m_CanReadShutterStatus, m_CanReadAzimuth, m_CanSlewToAzimuth As Boolean
    Private m_AsyncSlewAzimuth, m_AsyncSlewAltitude As Boolean

    'General variables
    Private m_Slewing, m_AtHome, m_AtPark, m_CanFindHome, m_CanFindPark, m_CanPark, m_CanSetPark, m_Connected As Boolean
    Private m_Description, m_DriverINfo, m_Name As String
    Private m_InterfaceVersion As Short
    Private m_Altitude, m_Azimuth As Double


#If DEBUG Then
    Private m_Dome As IDomeV2
#Else
    Private m_Dome As Object
#End If
#End Region

#Region "Enums"
    Private Enum DomePropertyMethod
        'Properties
        Altitude
        AtHome
        AtPark
        Azimuth
        CanFindHome
        CanPark
        CanSetAltitude
        CanSetAzimuth
        CanSetPark
        CanSetShutter
        CanSlave
        CanSyncAzimuth
        Connected
        Description
        DriverInfo
        InterfaceVersion
        Name
        ShutterStatus
        SlavedRead
        SlavedWrite
        Slewing

        'Methods
        AbortSlew
        CloseShutter
        CommandBlind
        CommandBool
        CommandString
        FindHome
        OpenShutter
        Park
        SetPark
        SlewToAltitude
        SlewToAzimuth
        SyncToAzimuth
    End Enum
#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New(True, True, True, True, False, False, True) ' Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
    End Sub

    ' IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If True Then 'Should be True but make False to stop Conform from cleanly dropping the Dome object (useful for retaining driver in memory to change flags
                    Try : DisposeAndReleaseObject("Dome", m_Dome) : Catch : End Try

                    m_Dome = Nothing
                    GC.Collect()
                End If
            End If

        End If
        MyBase.Dispose(disposing)
        Me.disposedValue = True
    End Sub
#End Region

#Region "Code"
    Overrides Sub CheckInitialise()
        'Set the error type numbers according to the standards adopted by individual authors.
        'Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
        'messages to driver authors!
        Select Case g_DomeProgID.ToUpper
            Case Else 'I'm using the simulator values as the defaults since it is the reference platform
                g_ExNotImplemented = &H80040400
                g_ExInvalidValue1 = &H80040405
                g_ExInvalidValue2 = &H80040405
                g_ExNotSet1 = &H80040403
        End Select
        MyBase.CheckInitialise(g_DomeProgID)

    End Sub
    Overrides Sub CheckAccessibility()
        Dim l_IDome As Object, l_DriverAccessDome As DriverAccess.Dome
        Dim l_DeviceObject As Object = Nothing
        MyBase.CheckAccessibility(g_DomeProgID, DeviceType.Dome)
        'Try early binding
        l_IDome = Nothing
        Try
            l_DeviceObject = CreateObject(g_DomeProgID)
            l_IDome = CType(l_DeviceObject, ASCOM.Interface.IDome)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through interface IDome")
            Try
                l_IDome.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface IDome")
                l_IDome.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IDome")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IDome")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IDome) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IDome = Nothing
            GC.Collect()
        End Try

        Try
            l_DeviceObject = CreateObject(g_DomeProgID)
            l_IDome = CType(l_DeviceObject, ASCOM.DeviceInterface.IDomeV2)
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through interface IDomeV2")
            Try
                l_IDome.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface IDomeV2")
                l_IDome.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IDomeV2")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IDomeV2")
        Finally
            'Clean up
            Try : Marshal.ReleaseComObject(l_IDome) : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DeviceObject) : Catch : End Try
            l_DeviceObject = Nothing
            l_IDome = Nothing
            GC.Collect()
        End Try
        WaitForAbsolute(2 * DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system")

        'Try client access toolkit
        l_DriverAccessDome = Nothing
        Try
            l_DriverAccessDome = New ASCOM.DriverAccess.Dome(g_DomeProgID)
            LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit")
            Try
                l_DriverAccessDome.Connected = True
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit")
                l_DriverAccessDome.Connected = False
            Catch ex As Exception
                LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " & ex.Message)
                LogMsg("", MessageLevel.msgAlways, "")
            End Try
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " & ex.Message)
            LogMsg("", MessageLevel.msgAlways, "")
        Finally
            'Clean up
            Try : l_DriverAccessDome.Dispose() : Catch : End Try
            Try : Marshal.ReleaseComObject(l_DriverAccessDome) : Catch : End Try
            l_DriverAccessDome = Nothing
            GC.Collect()
        End Try

        WaitForAbsolute(2 * DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system")

    End Sub
    Overrides Sub CreateDevice()
#If DEBUG Then
         LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Dome to get a Dome object")
       m_Dome = New ASCOM.DriverAccess.Dome(g_DomeProgID)
        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
#Else
        If g_Settings.UseDriverAccess Then
            LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Dome to get a Dome object")
            m_Dome = New ASCOM.DriverAccess.Dome(g_DomeProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        Else
            LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Dome object")
            m_Dome = CreateObject(g_DomeProgID)
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver")
        End If
#End If
        g_Stop = False
    End Sub
    Overrides Property Connected() As Boolean
        Get
            Connected = m_Dome.Connected
        End Get
        Set(ByVal value As Boolean)
            m_Dome.Connected = value
            g_Stop = False
        End Set
    End Property
    Overrides Sub CheckCommonMethods()
        MyBase.CheckCommonMethods(m_Dome, DeviceType.Dome)
    End Sub


    Public Overrides Sub PreRunCheck()
        Dim l_VStringPtr, l_V1, l_V2, l_V3 As Integer, l_VString As String
        Dim l_MsgBoxResult As Microsoft.VisualBasic.MsgBoxResult

        'Add a test for a back level version of the Dome simulator - just abandon this process if any errors occur
        If g_DomeProgID.ToUpper = "DOMESIM.DOME" Then
            l_VString = ""
            Try
                l_VStringPtr = InStr(m_Dome.DriverInfo.ToUpper, "ASCOM DOME SIMULATOR ") 'Point at the start of the version string
                If l_VStringPtr > 0 Then 'There is a version string so read what it is
                    l_VString = Mid(m_Dome.DriverInfo.ToUpper, l_VStringPtr + 21) 'Get the version string
                    l_VStringPtr = InStr(l_VString, ".")
                    If l_VStringPtr > 1 Then 'there is a first version number part
                        l_V1 = CInt(Mid(l_VString, 1, l_VStringPtr - 1)) 'Extract the number
                        l_VString = Mid(l_VString, l_VStringPtr + 1) 'Get the second version number part
                        l_VStringPtr = InStr(l_VString, ".")
                        If l_VStringPtr > 1 Then 'there is a second version number part
                            l_V2 = CInt(Mid(l_VString, 1, l_VStringPtr - 1)) 'Extract the number
                            l_VString = Mid(l_VString, l_VStringPtr + 1) 'Get the third version number part
                            'Find the next non numeric character
                            l_VStringPtr = 0
                            Do
                                l_VStringPtr += 1
                            Loop Until Not IsNumeric(Mid(l_VString, l_VStringPtr, 1))
                            If l_VStringPtr > 1 Then 'there is a third version number part
                                l_V3 = CInt(Mid(l_VString, 1, l_VStringPtr - 1)) 'Extract the number
                                'Turn the version parts into a whole number
                                l_V1 = l_V1 * 1000000 + l_V2 * 1000 + l_V3
                                If l_V1 < 5000007 Then 'We are using a version earlier that 5.0.7 so issue a message
                                    LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the dome simulator has known conformance issues, ***")
                                    LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***")
                                    LogMsg("", MessageLevel.msgAlways, "")
                                Else
                                    LogMsg("Version Check", MessageLevel.msgDebug, "Version check OK")
                                End If
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                LogMsg("ConformanceCheck", MessageLevel.msgError, ex.ToString)
            End Try
        End If
        If Not TestStop() Then
            'Get into a consistent state
            Try
                m_Slewing = m_Dome.Slewing ' Try to read the Slewing property
                If m_Slewing Then
                    LogMsg("DomeSafety", MessageLevel.msgInfo, $"The Slewing property is true at device start-up. This could be by design or possibly Slewing logic is inverted?") ' Display a message if slewing is True
                End If
                DomeWaitForSlew(g_Settings.DomeAzimuthTimeout) 'Wait for slewing to finish
            Catch ex As Exception
                LogMsg("DomeSafety", MessageLevel.msgWarning, $"The Slewing property threw an exception and should not have: {ex.Message}") ' Display a warning message because Slewing should not throw an exception!
                LogMsg("DomeSafety", MessageLevel.msgDebug, $"{ex}") ' Log the full message in debug mode
            End Try
            If FrmConformMain.chkDomeShutter.Checked Then
                LogMsg("DomeSafety", MessageLevel.msgComment, "Attempting to open shutter as some tests may fail if it is closed...")
                Try
                    m_Dome.OpenShutter()
                    Try : DomeShutterWait(ShutterState.shutterOpen) : Catch : End Try
                    If TestStop() Then 'Stop button pressed
                        LogMsg("DomeSafety", MessageLevel.msgComment, "Stop button pressed, further testing abandoned, shutter status: " & m_Dome.ShutterStatus.ToString)
                    Else 'Got to end of test
                        If m_Dome.ShutterStatus = ShutterState.shutterOpen Then
                            LogMsg("DomeSafety", MessageLevel.msgOK, "Shutter status: " & m_Dome.ShutterStatus.ToString)
                        Else
                            LogMsg("DomeSafety", MessageLevel.msgWarning, "Shutter status: " & m_Dome.ShutterStatus.ToString)
                        End If
                    End If
                Catch ex As Exception
                    LogMsg("DomeSafety", MessageLevel.msgComment, "Unable to open shutter: " & ex.Message)
                    LogMsg("DomeSafety", MessageLevel.msgComment, "Prompting for manual intervention...")
                    l_MsgBoxResult = MsgBox("Please open shutter and press OK or press Cancel to stop tests", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal, "Dome Safety")
                    If l_MsgBoxResult = MsgBoxResult.Cancel Then 'Cancel was pressed
                        g_Stop = True 'Stop if cancel was pressed
                        LogMsg("DomeSafety", MessageLevel.msgComment, "Cancel pressed, further testing abandoned")
                    Else
                        LogMsg("DomeSafety", MessageLevel.msgComment, "OK pressed, testing continues")
                    End If

                End Try
                Status(StatusType.staTest, "")
            Else
                LogMsg("DomeSafety", MessageLevel.msgComment, "Open shutter check box is unchecked so shutter not opened")
            End If
        End If
    End Sub
    Public Overrides Sub ReadCanProperties()
        DomeMandatoryTest(DomePropertyMethod.CanFindHome, "CanFindHome") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.CanPark, "CanPark") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.CanSetAltitude, "CanSetAltitude") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.CanSetAzimuth, "CanSetAzimuth") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.CanSetPark, "CanSetPark") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.CanSetShutter, "CanSetShutter") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.CanSlave, "CanSlave") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.CanSyncAzimuth, "CanSyncAzimuth") : If TestStop() Then Exit Sub

    End Sub
    Overrides Sub CheckProperties()
        DomeOptionalTest(DomePropertyMethod.Altitude, MemberType.Property, "Altitude") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.AtHome, MemberType.Property, "AtHome") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.AtPark, MemberType.Property, "AtPark") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.Azimuth, MemberType.Property, "Azimuth") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.ShutterStatus, MemberType.Property, "ShutterStatus") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.SlavedRead, "Slaved Read") : If TestStop() Then Exit Sub
        If m_Slaved And (Not m_CanSlave) Then LogMsg("Slaved Read", MessageLevel.msgIssue, "Dome is slaved but CanSlave is false")
        DomeOptionalTest(DomePropertyMethod.SlavedWrite, MemberType.Property, "Slaved Write") : If TestStop() Then Exit Sub
        DomeMandatoryTest(DomePropertyMethod.Slewing, "Slewing") : If TestStop() Then Exit Sub
    End Sub
    Overrides Sub CheckMethods()
        DomeMandatoryTest(DomePropertyMethod.AbortSlew, "AbortSlew") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.SlewToAltitude, MemberType.Method, "SlewToAltitude") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.SlewToAzimuth, MemberType.Method, "SlewToAzimuth") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.SyncToAzimuth, MemberType.Method, "SyncToAzimuth") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.CloseShutter, MemberType.Method, "CloseShutter") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.OpenShutter, MemberType.Method, "OpenShutter") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.FindHome, MemberType.Method, "FindHome") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.Park, MemberType.Method, "Park") : If TestStop() Then Exit Sub
        DomeOptionalTest(DomePropertyMethod.SetPark, MemberType.Method, "SetPark") : If TestStop() Then Exit Sub 'SetPark must follow Park
    End Sub
    Public Overrides Sub CheckPerformance()
        If m_CanReadAltitude Then DomePerformanceTest(DomePropertyMethod.Altitude, "Altitude") : If TestStop() Then Exit Sub
        If m_CanReadAzimuth Then DomePerformanceTest(DomePropertyMethod.Azimuth, "Azimuth") : If TestStop() Then Exit Sub
        If m_CanReadShutterStatus Then DomePerformanceTest(DomePropertyMethod.ShutterStatus, "ShutterStatus") : If TestStop() Then Exit Sub
        If m_CanReadSlaved Then DomePerformanceTest(DomePropertyMethod.SlavedRead, "Slaved") : If TestStop() Then Exit Sub
        If m_CanReadSlewing Then DomePerformanceTest(DomePropertyMethod.Slewing, "Slewing") : If TestStop() Then Exit Sub
    End Sub
    Public Overrides Sub PostRunCheck()
        If FrmConformMain.chkDomeShutter.Checked Then
            If m_CanSetShutter Then
                LogMsg("DomeSafety", MessageLevel.msgInfo, "Attempting to close shutter...")
                Try 'Close shutter
                    m_Dome.CloseShutter()
                    DomeShutterWait(ShutterState.shutterClosed)
                    LogMsg("DomeSafety", MessageLevel.msgOK, "Shutter successfully closed")
                Catch ex As Exception
                    LogMsg("DomeSafety", MessageLevel.msgComment, "Exception closing shutter: " & ex.Message)
                    LogMsg("DomeSafety", MessageLevel.msgComment, "Please close shutter manually")
                End Try
            Else
                LogMsg("DomeSafety", MessageLevel.msgInfo, "CanSetShutter is false, please close the shutter manually")
            End If
        Else
            LogMsg("DomeSafety", MessageLevel.msgInfo, "Open shutter check box is unchecked so close shutter bypassed")
        End If
        '3.0.0.17 - Added check for CanPark
        If m_CanPark Then 'Can park the dome
            LogMsg("DomeSafety", MessageLevel.msgInfo, "Attempting to park dome...")
            Try 'Park
                m_Dome.Park()
                DomeWaitForSlew(g_Settings.DomeAzimuthTimeout)
                LogMsg("DomeSafety", MessageLevel.msgOK, "Dome successfully parked")
            Catch ex As Exception
                LogMsg("DomeSafety", MessageLevel.msgError, "Exception generated, unable to park dome")
            End Try
        Else
            LogMsg("DomeSafety", MessageLevel.msgInfo, "CanPark is false - skipping dome parking")
        End If

    End Sub

    Private Sub DomeSlewToAltitude(ByVal p_Name As String, ByVal p_Altitude As Double)
        Dim l_StartTime As Date

        Status(StatusType.staAction, "Slew to " & p_Altitude & " degrees")
        m_Dome.SlewToAltitude(p_Altitude)
        If m_CanReadSlewing Then 'Can read slewing so make sure dome is at rest
            l_StartTime = Now
            If m_Dome.Slewing Then 'Asynchronous slew
                DomeWaitForSlew(g_Settings.DomeAltitudeTimeout) : If TestStop() Then Exit Sub
                m_AsyncSlewAltitude = True
                LogMsg(p_Name & " " & p_Altitude, MessageLevel.msgOK, "Asynchronous slew OK")
            Else 'Sync slew
                m_AsyncSlewAltitude = False
                LogMsg(p_Name & " " & p_Altitude, MessageLevel.msgOK, "Synchronous slew OK")
            End If
        Else 'Can't read slewing so just assume synchronous
            LogMsg(p_Name & " " & p_Altitude, MessageLevel.msgOK, "Can't read Slewing so assume synchronous slew OK")
        End If
        Call DomeStabliisationWait()
    End Sub
    Private Sub DomeSlewToAzimuth(ByVal p_Name As String, ByVal p_Azimuth As Double)
        Status(StatusType.staAction, "Slew to " & p_Azimuth & " degrees")
        If p_Azimuth >= 0.0 And p_Azimuth <= 359.9999999 Then 'Only check for successful operation on legal values
            m_CanSlewToAzimuth = False
            m_Dome.SlewToAzimuth(p_Azimuth)
            m_CanSlewToAzimuth = True 'Command is supported and didn't generate an exception
        Else
            m_Dome.SlewToAzimuth(p_Azimuth)
        End If
        If m_CanReadSlewing Then 'Can read slewing so make sure dome is at rest
            If m_Dome.Slewing Then 'Asynchronous slew
                DomeWaitForSlew(g_Settings.DomeAzimuthTimeout) : If TestStop() Then Exit Sub
                m_AsyncSlewAzimuth = True
                LogMsg(p_Name & " " & p_Azimuth, MessageLevel.msgOK, "Asynchronous slew OK")
            Else 'Sync slew
                m_AsyncSlewAzimuth = False
                LogMsg(p_Name & " " & p_Azimuth, MessageLevel.msgOK, "Synchronous slew OK")
            End If
        Else 'Can't read slewing so just assume synchronous
            LogMsg(p_Name & " " & p_Azimuth, MessageLevel.msgOK, "Can't read Slewing so assume synchronous slew OK")
        End If
        Call DomeStabliisationWait()
    End Sub
    Private Sub DomeWaitForSlew(ByVal p_TimeOut As Double)
        Dim l_StartTime As Date
        l_StartTime = Now
        Do
            WaitFor(SLEEP_TIME)
            Application.DoEvents()
            Status(StatusType.staStatus, "Slewing Status: " & m_Dome.Slewing & ", Timeout: " & Format(Now.Subtract(l_StartTime).TotalSeconds, "#0") & "/" & p_TimeOut & ", press stop to abandon wait")
        Loop Until Not m_Dome.Slewing Or TestStop() Or (Now.Subtract(l_StartTime).TotalSeconds > p_TimeOut)

        Status(StatusType.staStatus, "")
        If (Now.Subtract(l_StartTime).TotalSeconds > p_TimeOut) Then
            LogMsg("DomeWaitForSlew", MessageLevel.msgError, "Timed out waiting for Dome slew, consider increasing time-outs in Options/Conform Options.")
            LogMsg("DomeWaitForSlew", MessageLevel.msgInfo, "Another cause of time-outs is if your Slewing Property logic is inverted or is not operating correctly.")
        End If
    End Sub
    Private Sub DomeMandatoryTest(ByVal p_Type As DomePropertyMethod, ByVal p_Name As String)
        Try
            Select Case p_Type
                'Properties
                Case DomePropertyMethod.CanFindHome
                    m_CanFindHome = m_Dome.CanFindHome
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanFindHome.ToString)
                Case DomePropertyMethod.CanPark
                    m_CanPark = m_Dome.CanPark
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanPark.ToString)
                Case DomePropertyMethod.CanSetAltitude
                    m_CanSetAltitude = m_Dome.CanSetAltitude
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanSetAltitude.ToString)
                Case DomePropertyMethod.CanSetAzimuth
                    m_CanSetAzimuth = m_Dome.CanSetAzimuth
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanSetAzimuth.ToString)
                Case DomePropertyMethod.CanSetPark
                    m_CanSetPark = m_Dome.CanSetPark
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanSetPark.ToString)
                Case DomePropertyMethod.CanSetShutter
                    m_CanSetShutter = m_Dome.CanSetShutter
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanSetShutter.ToString)
                Case DomePropertyMethod.CanSlave
                    m_CanSlave = m_Dome.CanSlave
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanSlave.ToString)
                Case DomePropertyMethod.CanSyncAzimuth
                    m_CanSyncAzimuth = m_Dome.CanSyncAzimuth
                    LogMsg(p_Name, MessageLevel.msgOK, m_CanSyncAzimuth.ToString)
                Case DomePropertyMethod.Connected
                    m_Connected = m_Dome.Connected
                    LogMsg(p_Name, MessageLevel.msgOK, m_Connected.ToString)
                Case DomePropertyMethod.Description
                    m_Description = m_Dome.Description
                    LogMsg(p_Name, MessageLevel.msgOK, m_Description.ToString)
                Case DomePropertyMethod.Description
                    m_Description = m_Dome.Description
                    LogMsg(p_Name, MessageLevel.msgOK, m_Description.ToString)
                Case DomePropertyMethod.DriverInfo
                    m_DriverINfo = m_Dome.DriverInfo
                    LogMsg(p_Name, MessageLevel.msgOK, m_DriverINfo.ToString)
                Case DomePropertyMethod.InterfaceVersion
                    m_InterfaceVersion = m_Dome.InterfaceVersion
                    LogMsg(p_Name, MessageLevel.msgOK, m_InterfaceVersion.ToString)
                Case DomePropertyMethod.Name
                    m_Name = m_Dome.Name
                    LogMsg(p_Name, MessageLevel.msgOK, m_Name.ToString)
                Case DomePropertyMethod.SlavedRead
                    m_CanReadSlaved = False
                    m_Slaved = m_Dome.Slaved
                    m_CanReadSlaved = True
                    LogMsg(p_Name, MessageLevel.msgOK, m_Slaved.ToString)
                Case DomePropertyMethod.Slewing
                    m_CanReadSlewing = False
                    m_Slewing = m_Dome.Slewing
                    m_CanReadSlewing = True
                    LogMsg(p_Name, MessageLevel.msgOK, m_Slewing.ToString)

                    'Methods
                Case DomePropertyMethod.AbortSlew
                    m_Dome.AbortSlew()
                    'Confirm that slaved is false
                    If m_CanReadSlaved Then 'Check it is false
                        If m_Dome.Slaved Then 'Should not be showing slaved after AbortSlew
                            LogMsg("AbortSlew", MessageLevel.msgError, "Slaved property Is true after AbortSlew")
                        Else ' Not slaved so fine
                            LogMsg("AbortSlew", MessageLevel.msgOK, "AbortSlew command issued successfully")
                        End If
                    Else 'Assume all is OK
                        LogMsg("AbortSlew", MessageLevel.msgOK, "Can't read Slaved property AbortSlew command was successful")
                    End If
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "DomeMandatoryTest: Unknown test type " & p_Type.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "")
        End Try
    End Sub
    Private Sub DomeOptionalTest(ByVal p_Type As DomePropertyMethod, p_MemberType As MemberType, ByVal p_Name As String)
        Dim l_SlewAngle, l_OriginalAzimuth, l_NewAzimuth As Double
        Try
            Select Case p_Type
                'Properties
                Case DomePropertyMethod.Altitude
                    m_CanReadAltitude = False
                    m_Altitude = m_Dome.Altitude
                    m_CanReadAltitude = True
                    LogMsg(p_Name, MessageLevel.msgOK, m_Altitude.ToString)
                Case DomePropertyMethod.AtHome
                    m_CanReadAtHome = False
                    m_AtHome = m_Dome.AtHome
                    m_CanReadAtHome = True
                    LogMsg(p_Name, MessageLevel.msgOK, m_AtHome.ToString)
                Case DomePropertyMethod.AtPark
                    m_CanReadAtPark = False
                    m_AtPark = m_Dome.AtPark
                    m_CanReadAtPark = True
                    LogMsg(p_Name, MessageLevel.msgOK, m_AtPark.ToString)
                Case DomePropertyMethod.Azimuth
                    m_CanReadAzimuth = False
                    m_Azimuth = m_Dome.Azimuth
                    m_CanReadAzimuth = True
                    LogMsg(p_Name, MessageLevel.msgOK, m_Azimuth.ToString)
                Case DomePropertyMethod.ShutterStatus
                    m_CanReadShutterStatus = False
                    m_ShutterStatus = m_Dome.ShutterStatus
                    m_CanReadShutterStatus = True
                    m_ShutterStatus = CType(m_ShutterStatus, ShutterState) 'In release mode only an integer value is returned
                    LogMsg(p_Name, MessageLevel.msgOK, m_ShutterStatus.ToString)
                Case DomePropertyMethod.SlavedWrite
                    If m_CanSlave Then 'Write test should succeed
                        If m_CanReadSlaved Then 'Can read so worth testing write
                            If m_Slaved Then 'We are slaved so try and turn it off
                                m_Dome.Slaved = False
                            Else 'Not slaved so try and turn on
                                m_Dome.Slaved = True
                            End If
                            m_Dome.Slaved = m_Slaved 'Restore original value
                            LogMsg("Slaved Write", MessageLevel.msgOK, "Slave state changed successfully")
                        Else 'Can't read so skip test
                            LogMsg("Slaved Write", MessageLevel.msgInfo, "Test skipped since Slaved property can't be read")
                        End If
                    Else 'Slaved write should generate an exception
                        m_Dome.Slaved = True
                        LogMsg(p_Name, MessageLevel.msgError, "CanSlave is false but setting Slaved true did not raise an exception")
                        m_Dome.Slaved = False 'Unslave to continue tests
                    End If
                    'Methods
                Case DomePropertyMethod.CloseShutter
                    If m_CanSetShutter Then 'CloseShutter should work OK
                        Try
                            DomeShutterTest(ShutterState.shutterClosed, p_Name)
                            Call DomeStabliisationWait()
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetShutter is True")
                        End Try
                    Else 'CloseShutter should throw an exception
                        m_Dome.CloseShutter()
                        LogMsg(p_Name, MessageLevel.msgError, "CanSetShutter is false but CloseShutter did not raise an exception")
                    End If
                Case DomePropertyMethod.CommandBlind
                    m_Dome.CommandBlind("") 'm_Dome.CommandBlind("", True)
                    LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent")
                Case DomePropertyMethod.CommandBool
                    m_Dome.CommandBool("") 'm_Dome.CommandBool("", True)
                    LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent")
                Case DomePropertyMethod.CommandString
                    m_Dome.CommandString("") 'm_Dome.CommandString("", True)
                    LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent")
                Case DomePropertyMethod.FindHome
                    If m_CanFindHome Then
                        Status(StatusType.staTest, p_Name)
                        Status(StatusType.staAction, "Waiting for movement to stop")
                        Try
                            m_Dome.FindHome()
                            If m_CanReadSlaved Then 'Check whether slaved is true, if it is then Park  should have raised an exception and didn't
                                If m_Dome.Slaved Then LogMsg(p_Name, MessageLevel.msgError, "Slaved is true but Home did not raise an exception")
                            End If
                            If m_CanReadSlewing Then 'Make sure dome is at rest
                                Do
                                    WaitFor(SLEEP_TIME)
                                    Application.DoEvents()
                                    Status(StatusType.staStatus, "Slewing Status: " & m_Dome.Slewing)
                                Loop Until Not m_Dome.Slewing Or TestStop()
                            End If
                            If Not TestStop() Then 'Only do remaining tests if stop hasn't been pressed
                                If m_CanReadAtHome Then 'Can read AtHome so confirm that it 
                                    If m_Dome.AtHome Then 'Dome shows as homed - hooray!
                                        LogMsg(p_Name, MessageLevel.msgOK, "Dome homed successfully")
                                    Else 'Home completed but apparently dome isn't homed!
                                        LogMsg(p_Name, MessageLevel.msgError, "Home command completed but AtHome is false")
                                    End If
                                Else
                                    LogMsg(p_Name, MessageLevel.msgOK, "Can't read AtHome so assume that dome has homed successfully")
                                End If
                                Call DomeStabliisationWait()
                            End If
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanFindHome is True")
                            Call DomeStabliisationWait()
                        End Try
                    Else 'CanFindHome is false so FindHome should throw a not implemented exception
                        m_Dome.FindHome()
                        LogMsg(p_Name, MessageLevel.msgError, "CanFindHome is false but FindHome did not throw an exception")
                    End If
                Case DomePropertyMethod.OpenShutter
                    If m_CanSetShutter Then 'OpenShutter should work OK
                        Try
                            DomeShutterTest(ShutterState.shutterOpen, p_Name)
                            Call DomeStabliisationWait()
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetShutter is True")
                        End Try
                    Else 'OpenShutter should throw an exception
                        m_Dome.OpenShutter()
                        LogMsg(p_Name, MessageLevel.msgError, "CanSetShutter is false but OpenShutter did not raise an exception")
                    End If

                Case DomePropertyMethod.Park
                    If m_CanPark Then 'Should be able to issue the Park command
                        Status(StatusType.staTest, p_Name)
                        Status(StatusType.staAction, "Waiting for movement to stop")
                        Try
                            m_Dome.Park()
                            If m_CanReadSlaved Then 'Check whether slaved is true, if it is then Park  should have raised an exception and didn't
                                If m_Dome.Slaved Then LogMsg(p_Name, MessageLevel.msgError, "Slaved is true but Park did not raise an exception")
                            End If
                            If m_CanReadSlewing Then 'Make sure dome is at rest
                                Do
                                    WaitFor(SLEEP_TIME)
                                    Application.DoEvents()
                                    Status(StatusType.staStatus, "Slewing Status: " & m_Dome.Slewing)
                                Loop Until Not m_Dome.Slewing Or TestStop()
                            End If
                            If Not TestStop() Then 'Only do remain tests if stop hasn't been pressed
                                If m_CanReadAtPark Then 'Can read at park so confirm that it 
                                    If m_Dome.AtPark Then 'Dome shows as parked - hooray!
                                        LogMsg(p_Name, MessageLevel.msgOK, "Dome parked successfully")
                                    Else 'Park completed but apparently dome isn't parked!
                                        LogMsg(p_Name, MessageLevel.msgError, "Park command completed but AtPark is false")
                                    End If
                                Else
                                    LogMsg(p_Name, MessageLevel.msgOK, "Can't read AtPark so assume that dome has parked successfully")
                                End If
                            End If
                            Call DomeStabliisationWait()
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanPark is True")
                            Call DomeStabliisationWait()
                        End Try
                    Else 'Park command should throw a not implemented exception
                        m_Dome.Park()
                        LogMsg(p_Name, MessageLevel.msgError, "CanPark is false but Park did not raise an exception")
                    End If
                Case DomePropertyMethod.SetPark
                    If m_CanSetPark Then 'Should be able to set park so try it
                        Try
                            m_Dome.SetPark()
                            LogMsg(p_Name, MessageLevel.msgOK, "SetPark issued OK")
                        Catch ex As Exception
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetPark is True")
                        End Try
                    Else 'Can't set park so should raise an error
                        m_Dome.SetPark()
                        LogMsg(p_Name, MessageLevel.msgError, "CanSetPath is false but SetPath did not throw an exception")
                    End If
                Case DomePropertyMethod.SlewToAltitude
                    If m_CanSetAltitude Then 'SlewToAltitude should work OK
                        Status(StatusType.staTest, p_Name)
#If DEBUG Then
                        For l_SlewAngle = 0 To 45 Step 45
#Else
                        For l_SlewAngle = 0 To 90 Step 15
#End If
                            Try
                                DomeSlewToAltitude(p_Name, l_SlewAngle)
                                If TestStop() Then Exit Sub
                            Catch ex As Exception
                                HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetAltitude is True")
                            End Try
                        Next
                        'Test out of range values -10 and 100 degrees
                        If m_CanSetAltitude Then 'Can set altitude so check out of range values
                            Try
                                DomeSlewToAltitude(p_Name, DOME_ILLEGAL_ALTITUDE_LOW)
                                LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal altitude " & DOME_ILLEGAL_ALTITUDE_LOW & " degrees")
                            Catch ex As Exception
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " & DOME_ILLEGAL_ALTITUDE_LOW & " degrees", "COM invalid value exception correctly raised for slew to " & DOME_ILLEGAL_ALTITUDE_LOW & " degrees")
                            End Try
                            Try
                                DomeSlewToAltitude(p_Name, DOME_ILLEGAL_ALTITUDE_HIGH)
                                LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal altitude " & DOME_ILLEGAL_ALTITUDE_HIGH & " degrees")
                            Catch ex As Exception
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " & DOME_ILLEGAL_ALTITUDE_HIGH & " degrees", "COM invalid value exception correctly raised for slew to " & DOME_ILLEGAL_ALTITUDE_HIGH & " degrees")
                            End Try
                        End If
                    Else 'SlewToAltitude should raise an exception
                        m_Dome.SlewToAltitude(45.0)
                        LogMsg(p_Name, MessageLevel.msgError, "CanSetAltitude is false but SlewToAltitude did not raise an exception")
                    End If

                Case DomePropertyMethod.SlewToAzimuth
                    If m_CanSetAzimuth Then 'Can set azimuth so SlewToAzimuth should succeed
                        Status(StatusType.staTest, p_Name)
#If DEBUG Then
                        For l_SlewAngle = 0 To 30 Step 30
#Else
                        For l_SlewAngle = 0 To 315 Step 45
#End If
                            Try
                                DomeSlewToAzimuth(p_Name, l_SlewAngle)
                                If TestStop() Then Exit Sub
                            Catch ex As Exception
                                HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetAzimuth is True")
                            End Try
                        Next

                        If m_CanSetAzimuth Then
                            'Test out of range values -10 and 370 degrees
                            Try
                                DomeSlewToAzimuth(p_Name, DOME_ILLEGAL_AZIMUTH_LOW)
                                LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal azimuth " & DOME_ILLEGAL_AZIMUTH_LOW & " degrees")
                            Catch ex As Exception
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " & DOME_ILLEGAL_AZIMUTH_LOW & " degrees", "COM invalid value exception correctly raised for slew to " & DOME_ILLEGAL_AZIMUTH_LOW & " degrees")
                            End Try
                            If TestStop() Then Exit Sub
                            Try
                                DomeSlewToAzimuth(p_Name, DOME_ILLEGAL_AZIMUTH_HIGH)
                                LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal azimuth " & DOME_ILLEGAL_AZIMUTH_HIGH & " degrees")
                            Catch ex As Exception
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " & DOME_ILLEGAL_AZIMUTH_HIGH & " degrees", "COM invalid value exception correctly raised for slew to " & DOME_ILLEGAL_AZIMUTH_HIGH & " degrees")
                            End Try
                            If TestStop() Then Exit Sub
                        End If
                    Else 'SlewToAzimuth should throw an exception
                        m_Dome.SlewToAzimuth(45.0)
                        LogMsg(p_Name, MessageLevel.msgError, "CanSetAzimuth is false but SlewToAzimuth did not throw an exception")
                    End If
                Case DomePropertyMethod.SyncToAzimuth
                    If m_CanSyncAzimuth Then 'Can sync azimuth so test the command
                        If m_CanSlewToAzimuth Then 'SlewToAzimuth command appears to work so use it
                            If m_CanReadAzimuth Then 'Can also read azimuth so test sync to a new azimuth
                                l_OriginalAzimuth = m_Dome.Azimuth
                                If l_OriginalAzimuth > 300.0 Then ' Choose a smaller azimuth
                                    l_NewAzimuth = l_OriginalAzimuth - DOME_SYNC_OFFSET
                                Else 'Choose a bigger azimuth
                                    l_NewAzimuth = l_OriginalAzimuth + DOME_SYNC_OFFSET
                                End If
                                m_Dome.SyncToAzimuth(l_NewAzimuth) 'Sync to new azimuth
                                ' OK Dome hasn't moved but should now show azimuth as a new value
                                Select Case Math.Abs(l_NewAzimuth - m_Dome.Azimuth)
                                    Case Is < 1.0 'very close so give it an OK
                                        LogMsg(p_Name, MessageLevel.msgOK, "Dome synced OK to within +- 1 degree")
                                    Case Is < 2.0 'close so give it an INFO
                                        LogMsg(p_Name, MessageLevel.msgInfo, "Dome synced to within +- 2 degrees")
                                    Case Is < 5.0 'Closish so give an issue
                                        LogMsg(p_Name, MessageLevel.msgIssue, "Dome only synced to within +- 5 degrees")
                                    Case (DOME_SYNC_OFFSET - 2.0) To (DOME_SYNC_OFFSET + 2) 'Hasn't really moved
                                        LogMsg(p_Name, MessageLevel.msgError, "Dome did not sync, Azimuth didn't change value after sync command")
                                    Case Else 'Something else!
                                        LogMsg(p_Name, MessageLevel.msgIssue, "Dome azimuth was " & Math.Abs(l_NewAzimuth - m_Dome.Azimuth) & " degrees away from expected value")
                                End Select
                                'Now try and restore original value
                                m_Dome.SyncToAzimuth(l_OriginalAzimuth)
                            Else 'Can't read azimuth so can only test that command completes
                                m_Dome.SyncToAzimuth(45.0) 'Sync to an arbitrary direction
                                LogMsg(p_Name, MessageLevel.msgOK, "Dome successfully synced to 45 degrees but unable to read azimuth to confirm this")
                            End If

                            'Now test sync to illegal values
                            Try
                                m_Dome.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_LOW)
                                LogMsg(p_Name, MessageLevel.msgError, "No exception generated when syncing to illegal azimuth " & DOME_ILLEGAL_AZIMUTH_LOW & " degrees")
                            Catch ex As Exception
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "sync to " & DOME_ILLEGAL_AZIMUTH_LOW & " degrees", "COM invalid value exception correctly raised for sync to " & DOME_ILLEGAL_AZIMUTH_LOW & " degrees")
                            End Try
                            If TestStop() Then Exit Sub
                            Try
                                m_Dome.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_HIGH)
                                LogMsg(p_Name, MessageLevel.msgError, "No exception generated when syncing to illegal azimuth " & DOME_ILLEGAL_AZIMUTH_HIGH & " degrees")
                            Catch ex As Exception
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "sync to " & DOME_ILLEGAL_AZIMUTH_HIGH & " degrees", "COM invalid value exception correctly raised for sync to " & DOME_ILLEGAL_AZIMUTH_HIGH & " degrees")
                            End Try
                            If TestStop() Then Exit Sub
                        Else 'SlewoAzimuth doesn't work so skip test
                            LogMsg(p_Name, MessageLevel.msgInfo, "SyncToAzimuth test skipped since SlewToAzimuth throws an exception")
                        End If
                    Else 'Can not sync azimuth so SyncAzimuth should raise an exception
                        m_Dome.SyncToAzimuth(45.0)
                        LogMsg(p_Name, MessageLevel.msgError, "CanSyncAzimuth is false but SyncToAzimuth did not raise an exception")
                    End If
                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "DomeOptionalTest: Unknown test type " & p_Type.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, p_MemberType, Required.Optional, ex, "")
        End Try
        Status(StatusType.staTest, "")
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")
    End Sub
    Private Sub DomeShutterTest(ByVal p_RequiredShutterState As ShutterState, ByVal p_Name As String)
        Dim l_MsgBoxResult As MsgBoxResult, l_ShutterState As ShutterState

        If FrmConformMain.chkDomeShutter.Checked Then 'Shutter tests are allowed
            Status(StatusType.staTest, p_Name)
            If m_CanReadShutterStatus Then 'Can read shutter status so use it
                l_ShutterState = CType(m_Dome.ShutterStatus, ShutterState)

                'Make sure we are in the required state to start the test
                Select Case l_ShutterState
                    Case ShutterState.shutterClosed
                        If p_RequiredShutterState = ShutterState.shutterClosed Then 'Testing ShutterClose
                            'Wrong state, get to the required state
                            Status(StatusType.staAction, "Opening shutter ready for close test")
                            LogMsg(p_Name, MessageLevel.msgDebug, "Opening shutter ready for close test")
                            m_Dome.OpenShutter()
                            If Not DomeShutterWait(ShutterState.shutterOpen) Then Exit Sub 'Wait for shutter to open
                            Call DomeStabliisationWait()
                        Else 'Testing ShutterOpen
                            'No action this is the desired start state
                        End If
                    Case ShutterState.shutterClosing
                        If p_RequiredShutterState = ShutterState.shutterClosed Then 'Testing ShutterClose
                            Status(StatusType.staAction, "Waiting for shutter to close before opening ready for close test")
                            LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close before opening ready for close test")
                            If Not DomeShutterWait(ShutterState.shutterClosed) Then Exit Sub 'Wait for shutter to close
                            LogMsg(p_Name, MessageLevel.msgDebug, "Opening shutter ready for close test")
                            Status(StatusType.staAction, "Opening shutter ready for close test")
                            m_Dome.OpenShutter() 'Then open it
                            If Not DomeShutterWait(ShutterState.shutterOpen) Then Exit Sub
                            Call DomeStabliisationWait()
                        Else 'Testing ShutterOpen
                            Status(StatusType.staAction, "Waiting for shutter to close ready for open test")
                            LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close ready for open test")
                            If Not DomeShutterWait(ShutterState.shutterClosed) Then Exit Sub 'Wait for shutter to close
                            Call DomeStabliisationWait()
                        End If
                    Case ShutterState.shutterOpening
                        If p_RequiredShutterState = ShutterState.shutterClosed Then 'Testing ShutterClose
                            Status(StatusType.staAction, "Waiting for shutter to open ready for close test")
                            LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open ready for close test")
                            If Not DomeShutterWait(ShutterState.shutterOpen) Then Exit Sub 'Wait for shutter to open
                            Call DomeStabliisationWait()
                        Else 'Testing for ShutterOpen
                            Status(StatusType.staAction, "Waiting for shutter to open before closing ready for open test")
                            LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open before closing ready for open test")
                            If Not DomeShutterWait(ShutterState.shutterOpen) Then Exit Sub 'Wait for shutter to open
                            LogMsg(p_Name, MessageLevel.msgDebug, "Closing shutter ready for open test")
                            Status(StatusType.staAction, "Closing shutter ready for open test")
                            m_Dome.CloseShutter() 'Then close it
                            If Not DomeShutterWait(ShutterState.shutterClosed) Then Exit Sub
                            Call DomeStabliisationWait()
                        End If
                    Case ShutterState.shutterError
                        l_MsgBoxResult = MsgBox("ShutterState is ""Error"" please prepare the shutter for the " & p_Name & " test and press OK or press Cancel to stop tests", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal, "Dome Safety")
                        If l_MsgBoxResult = MsgBoxResult.Cancel Then 'Cancel was pressed
                            g_Stop = True 'Stop if cancel was pressed
                            LogMsg("DomeSafety", MessageLevel.msgComment, "Cancel pressed, further testing abandoned")
                        Else
                            LogMsg("DomeSafety", MessageLevel.msgComment, "OK pressed, testing continues")
                        End If
                    Case ShutterState.shutterOpen
                        If p_RequiredShutterState = ShutterState.shutterClosed Then 'Testing ShutterClose
                            'No action this is the desired start state
                        Else 'Testing ShutterOpen
                            'Wrong state, get to the required state
                            Status(StatusType.staAction, "Closing shutter ready for open  test")
                            LogMsg(p_Name, MessageLevel.msgDebug, "Closing shutter ready for open test")
                            m_Dome.CloseShutter()
                            If Not DomeShutterWait(ShutterState.shutterClosed) Then Exit Sub 'Wait for shutter to open
                            Call DomeStabliisationWait()
                        End If
                    Case Else
                        LogMsg("DomeShutterTest", MessageLevel.msgError, "Unexpected shutter status: " & l_ShutterState.ToString)
                End Select

                ' Now test that we can get to the required state
                If p_RequiredShutterState = ShutterState.shutterClosed Then 'Testing ShutterClose
                    'Shutter is now open so close it
                    Status(StatusType.staAction, "Closing shutter")
                    m_Dome.CloseShutter()
                    Status(StatusType.staAction, "Waiting for shutter to close")
                    LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close")
                    If Not DomeShutterWait(ShutterState.shutterClosed) Then
                        l_ShutterState = m_Dome.ShutterStatus
                        l_ShutterState = CType(l_ShutterState, ShutterState)
                        LogMsg(p_Name, MessageLevel.msgError, "Unable to close shutter - ShutterStatus: " & m_Dome.ShutterStatus.ToString)
                        Exit Sub
                    Else
                        LogMsg(p_Name, MessageLevel.msgOK, "Shutter closed successfully")
                    End If
                    Call DomeStabliisationWait()
                Else 'Shutter is now closed so test that we can open it
                    Status(StatusType.staAction, "Opening shutter")
                    m_Dome.OpenShutter()
                    Status(StatusType.staAction, "Waiting for shutter to open")
                    LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open")
                    If Not DomeShutterWait(ShutterState.shutterOpen) Then
                        l_ShutterState = m_Dome.ShutterStatus
                        l_ShutterState = CType(l_ShutterState, ShutterState)
                        LogMsg(p_Name, MessageLevel.msgError, "Unable to open shutter - ShutterStatus: " & m_Dome.ShutterStatus.ToString)
                        Exit Sub
                    Else
                        LogMsg(p_Name, MessageLevel.msgOK, "Shutter opened successfully")
                    End If
                    Call DomeStabliisationWait()
                End If

            Else 'Can't read shutter status so don't try
                LogMsg(p_Name, MessageLevel.msgDebug, "Can't read shutter status!")
                If p_RequiredShutterState = ShutterState.shutterClosed Then 'Testing ShutterClose
                    'Just issue command to see if it doesn't generate an error
                    m_Dome.CloseShutter()
                    Call DomeStabliisationWait()
                Else 'Testing ShutterOpen
                    'Just issue command to see if it doesn't generate an error
                    m_Dome.OpenShutter()
                    Call DomeStabliisationWait()
                End If
                LogMsg(p_Name, MessageLevel.msgOK, "Command issued successfully but can't read ShutterStatus to confirm shutter is closed")
            End If
            Status(StatusType.staTest, "")
            Status(StatusType.staAction, "")
            Status(StatusType.staStatus, "")
        Else
            LogMsg("DomeSafety", MessageLevel.msgComment, "Open shutter check box is unchecked so shutter test bypassed")
        End If
    End Sub

    Private Function DomeShutterWait(ByVal p_RequiredStatus As ShutterState) As Boolean
        'Wait for shutter to reach required stats or user presses stop or timeout occurs
        'Returns true if required state is reached
        Dim l_StartTime As Date, l_ShutterState As ShutterState
        DomeShutterWait = False
        l_StartTime = Now
        Try
            Do
                WaitFor(SLEEP_TIME)
                Application.DoEvents()
                l_ShutterState = m_Dome.ShutterStatus
                l_ShutterState = CType(l_ShutterState, ShutterState)
                Status(StatusType.staStatus, "Shutter State: " & l_ShutterState.ToString & " Timeout: " & Now.Subtract(l_StartTime).Seconds & "/" & g_Settings.DomeShutterTimeout)
            Loop Until (l_ShutterState = p_RequiredStatus) Or TestStop() Or (Now.Subtract(l_StartTime).TotalSeconds > g_Settings.DomeShutterTimeout)
            If (m_Dome.ShutterStatus = p_RequiredStatus) Then DomeShutterWait = True ' All worked so return True
            If (Now.Subtract(l_StartTime).TotalSeconds > g_Settings.DomeShutterTimeout) Then ' We timed out so give error message
                LogMsg("DomeShutterWait", MessageLevel.msgError, "Timed out waiting for shutter to reach state: " & p_RequiredStatus.ToString & ", consider increasing the timeout setting in Options / Conformance Options")
            End If
        Catch ex As Exception
            LogMsg("DomeShutterWait", MessageLevel.msgError, "Unexpected exception: " & ex.ToString)
        End Try
    End Function
    Private Sub DomePerformanceTest(ByVal p_Type As DomePropertyMethod, ByVal p_Name As String)
        Dim l_StartTime As Date, l_Count, l_LastElapsedTime, l_ElapsedTime As Double, l_Rate As Double
        Dim l_Boolean As Boolean, l_Double As Double, l_ShutterState As ShutterState
        Status(StatusType.staTest, "Performance Testing")
        Status(StatusType.staAction, p_Name)
        Try
            l_StartTime = Now
            l_Count = 0.0
            l_LastElapsedTime = 0.0
            Do
                l_Count += 1.0
                Select Case p_Type
                    Case DomePropertyMethod.Altitude
                        l_Double = m_Dome.Altitude
                    Case DomePropertyMethod.Azimuth
                        l_Double = m_Dome.Azimuth
                    Case DomePropertyMethod.ShutterStatus
                        l_ShutterState = m_Dome.ShutterStatus
                    Case DomePropertyMethod.SlavedRead
                        l_Boolean = m_Dome.Slaved
                    Case DomePropertyMethod.Slewing
                        l_Boolean = m_Dome.Slewing
                    Case Else
                        LogMsg(p_Name, MessageLevel.msgError, "DomePerformanceTest: Unknown test type " & p_Type.ToString)
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

    Sub DomeStabliisationWait()
        Status(StatusType.staStatus, "") 'Clear status field
        For i As Double = 1.0 To g_Settings.DomeStabilisationWait
            Status(StatusType.staAction, "Waiting for Dome to stabilise - " & CInt(i) & "/" & g_Settings.DomeStabilisationWait & " seconds")
            WaitFor(1000) ' Wait for 1 second
        Next
    End Sub
#End Region

End Class
