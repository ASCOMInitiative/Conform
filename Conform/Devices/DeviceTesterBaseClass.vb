'Base class from which particular device testers are derrived
'Put all common elements in here
Imports ASCOM.DeviceInterface
Imports System.Threading
Imports System.Runtime.InteropServices
Imports ASCOM

''' <summary>
''' Base class for device tester classes. Contains common code and placeholders for the 
''' methods that must be implemented in the device tester class
''' </summary>
Friend Class DeviceTesterBaseClass
    Implements IDisposable

#Region "Variables and Constants"
    Dim l_Connected, l_HasProperties, l_HasCanProperties, l_HasMethods, l_HasPreRunCheck, l_HasPostRunCheck, l_HasPerformanceCheck As Boolean
    Dim l_HasPreConnectCheck As Boolean
    Dim Device As Object 'IAscomDriverV1
#End Region

#Region "Enums"
    Private Enum CommandType
        tstCommandString = 1
        tstCommandBool = 2
        tstCommandBlind = 3
        tstCommandStringRaw = 4
        tstCommandBoolRaw = 5
        tstCommandBlindRaw = 6
    End Enum

    Protected Enum MemberType
        [Property]
        [Method]
    End Enum

    Protected Enum Required
        [Optional]
        [Mandatory]
        MustBeImplemented
        MustNotBeImplemented
    End Enum
#End Region

#Region "New and Dispose"
    Sub New()
        MyBase.New()
        l_HasPreConnectCheck = False
        l_Connected = False
        l_HasPreRunCheck = False
        l_HasCanProperties = False
        l_HasProperties = True
        l_HasMethods = True
        l_HasPostRunCheck = False
        l_HasPerformanceCheck = True
    End Sub

    ''' <summary>
    ''' Initialise device tester base class
    ''' </summary>
    ''' <param name="HasCanProperties">Device has Can properties</param>
    ''' <param name="HasProperties">Device has properties</param>
    ''' <param name="HasMethods">Device has methods</param>
    ''' <param name="HasPreRunCheck">Device requires a pre run safety check </param>
    ''' <param name="HasPreConnectCheck">Device requires a pre connection check</param>
    ''' <param name="HasPerformanceCheck">Device has a performance test</param>
    ''' <param name="HasPostRunCheck">Device requiures a post run safety check</param>
    ''' <remarks></remarks>
    Sub New(HasCanProperties As Boolean, HasProperties As Boolean, HasMethods As Boolean, HasPreRunCheck As Boolean, HasPreConnectCheck As Boolean, HasPerformanceCheck As Boolean, HasPostRunCheck As Boolean)
        MyBase.New()
        l_HasPreConnectCheck = HasPreConnectCheck
        l_Connected = False
        l_HasPreRunCheck = HasPreRunCheck
        l_HasCanProperties = HasCanProperties
        l_HasProperties = HasProperties
        l_HasMethods = HasMethods
        l_HasPostRunCheck = HasPostRunCheck
        l_HasPerformanceCheck = HasPerformanceCheck
    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls
    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

#Region "Code"

    Overloads Sub CheckCommonMethods(ByVal p_DeviceObject As Object, ByVal p_DeviceType As DeviceType)
        Dim m_DriverVersion, m_DriverInfo, m_Description, m_Name As String ', m_LastResult As String
        Dim SA As ArrayList

        'Dim m_Configuration, SC() As String
        Dim m_Connected As Boolean
        Dim l_HasInterfaceVersion As Boolean = False

        LogMsg("Common Driver Methods", MessageLevel.msgAlways, "")
        Device = p_DeviceObject 'CType(DeviceObject, IAscomDriverV1)

        'InterfaceVersion - Required
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("InterfaceVersion", MessageLevel.msgComment, "About to get property InterfaceVersion")
            g_InterfaceVersion = Device.InterfaceVersion
            Select Case g_InterfaceVersion
                Case Is < 1
                    LogMsg("InterfaceVersion", MessageLevel.msgIssue, "InterfaceVersion must be 1 or greater but driver returned: " & g_InterfaceVersion.ToString)
                Case Else
                    LogMsg("InterfaceVersion", MessageLevel.msgOK, g_InterfaceVersion.ToString)
            End Select
            l_HasInterfaceVersion = True
        Catch ex As Exception
            HandleException("InterfaceVersion", MemberType.Property, Required.Mandatory, ex, "")
        End Try
        If TestStop() Then Exit Sub

        'Connected - Required
        If IncludeMethod(MandatoryMethod.Connected, p_DeviceType, g_InterfaceVersion) Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("Connected", MessageLevel.msgComment, "About to get property Connected")
                m_Connected = Device.Connected
                LogMsg("Connected", MessageLevel.msgOK, m_Connected.ToString)
            Catch ex As COMException
                LogMsg("Connected", MessageLevel.msgError, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
            Catch ex As Exception
                LogMsg("Connected", MessageLevel.msgError, ex.Message)
            End Try
            If TestStop() Then Exit Sub
        End If

        'Description - Required
        If IncludeMethod(MandatoryMethod.Description, p_DeviceType, g_InterfaceVersion) Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("Description", MessageLevel.msgComment, "About to get property Description")
                m_Description = Device.Description
                Select Case m_Description
                    Case ""
                        LogMsg("Description", MessageLevel.msgInfo, "No description string")
                    Case Else
                        If (m_Description.Length > 68) And (p_DeviceType = DeviceType.Camera) Then
                            LogMsg("Description", MessageLevel.msgError, "Maximum number of characters is 68 for compatibility with FITS headers, found: " _
                                   & m_Description.Length & " characters: " & m_Description)
                        Else
                            LogMsg("Description", MessageLevel.msgOK, m_Description.ToString)
                        End If
                End Select
            Catch ex As Exception
                HandleException("Description", MemberType.Property, Required.Mandatory, ex, "")
            End Try
            If TestStop() Then Exit Sub
        End If

        'DriverInfo - Required
        If IncludeMethod(MandatoryMethod.DriverInfo, p_DeviceType, g_InterfaceVersion) Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("DriverInfo", MessageLevel.msgComment, "About to get property DriverInfo")
                m_DriverInfo = Device.DriverInfo
                Select Case m_DriverInfo
                    Case ""
                        LogMsg("DriverInfo", MessageLevel.msgInfo, "No DriverInfo string")
                    Case Else
                        LogMsg("DriverInfo", MessageLevel.msgOK, m_DriverInfo.ToString)
                End Select
            Catch ex As Exception
                HandleException("DriverInfo", MemberType.Property, Required.Mandatory, ex, "")
            End Try
            If TestStop() Then Exit Sub
        End If

        'DriverVersion - Required
        If IncludeMethod(MandatoryMethod.DriverVersion, p_DeviceType, g_InterfaceVersion) Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("DriverVersion", MessageLevel.msgComment, "About to get property DriverVersion")
                m_DriverVersion = Device.DriverVersion
                Select Case m_DriverVersion
                    Case ""
                        LogMsg("DriverVersion", MessageLevel.msgInfo, "No DriverVersion string")
                    Case Else
                        LogMsg("DriverVersion", MessageLevel.msgOK, m_DriverVersion.ToString)
                End Select
            Catch ex As COMException
                LogMsg("DriverVersion", MessageLevel.msgError, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
            Catch ex As Exception
                LogMsg("DriverVersion", MessageLevel.msgError, ex.Message)
            End Try
            If TestStop() Then Exit Sub
        Else
            LogMsg("DriverVersion", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" & g_InterfaceVersion)
        End If

        'Name - Required
        If IncludeMethod(MandatoryMethod.Name, p_DeviceType, g_InterfaceVersion) Then
            Try
                If g_Settings.DisplayMethodCalls Then LogMsg("Name", MessageLevel.msgComment, "About to get property Name")
                m_Name = Device.Name
                Select Case m_Name
                    Case ""
                        LogMsg("Name", MessageLevel.msgInfo, "Name is empty")
                    Case Else
                        LogMsg("Name", MessageLevel.msgOK, m_Name)
                End Select
            Catch ex As Exception
                HandleException("Name", MemberType.Property, Required.Mandatory, ex, "")
            End Try
            If TestStop() Then Exit Sub
        End If

        'CommandXXX tests - Optional
        If IncludeMethod(MandatoryMethod.CommandXXX, p_DeviceType, g_InterfaceVersion) Then
            If (g_TelescopeTests.Item(TELTEST_COMMANDXXX) = CheckState.Checked) Then
                'CommandTest(CommandType.tstCommandBlind, "CommandBlind") : If TestStop() Then Exit Sub
                'CommandTest(CommandType.tstCommandBool, "CommandBool") : If TestStop() Then Exit Sub
                'CommandTest(CommandType.tstCommandString, "CommandString") : If TestStop() Then Exit Sub
                'CommandTest(CommandType.tstCommandBlindRaw, "CommandBlind Raw") : If TestStop() Then Exit Sub
                'CommandTest(CommandType.tstCommandBoolRaw, "CommandBool Raw") : If TestStop() Then Exit Sub
                'CommandTest(CommandType.tstCommandStringRaw, "CommandString Raw") : If TestStop() Then Exit Sub
                LogMsg("CommandString", MessageLevel.msgInfo, "Conform cannot test the CommandString method")
                LogMsg("CommandBlind", MessageLevel.msgInfo, "Conform cannot test the CommandBlind method")
                LogMsg("CommandBool", MessageLevel.msgInfo, "Conform cannot test the CommandBool method")
            Else
                LogMsg(TELTEST_COMMANDXXX, MessageLevel.msgInfo, "Tests skipped")
            End If
        End If

        'Action - optional but cannot be tested
        LogMsg("Action", MessageLevel.msgInfo, "Conform cannot test the Action method")

        'Supported actions - Optional but Required through Driveraccess
        Try
            If g_Settings.DisplayMethodCalls Then LogMsg("SupportedActions", MessageLevel.msgComment, "About to call method SupportedActions")
            SA = Device.SupportedActions()
            If SA.Count = 0 Then
                LogMsg("SupportedActions", MessageLevel.msgOK, "Driver returned an empty action list")
            Else
                Dim i As Integer
                For Each Action As Object In SA
                    i += 1
                    If Action.GetType.Name = "String" Then
                        Dim ActionString As String = CType(Action, String)

                        Select Case ActionString
                            Case Is = ""
                                LogMsg("SupportedActions", MessageLevel.msgError, "Supported action " & i & " is an empty string")
                            Case Else ' List the action that was found
                                LogMsg("SupportedActions", MessageLevel.msgOK, "Found action: " & ActionString)

                                ' Carry out the following Action tests only when we are testing the Observing Conditions Hub and it is configured to use the Switch and OC simulators
                                If ((p_DeviceType = DeviceType.ObservingConditions) And (g_ObservingConditionsProgID.ToUpperInvariant() = "ASCOM.OCH.OBSERVINGCONDITIONS")) Then
                                    If ((ActionString.ToUpperInvariant.StartsWith("//OCSIMULATOR:")) Or (ActionString.ToUpperInvariant.StartsWith("//SWITCHSIMULATOR:"))) Then
                                        Dim result As String
                                        Try
                                            result = Device.Action(ActionString, "")
                                            LogMsg("SupportedActions", MessageLevel.msgOK, String.Format("{0} action gave result: {1}", ActionString, result))
                                        Catch ex1 As Exception
                                            LogMsg("SupportedActions", MessageLevel.msgError, String.Format("Exception calling action {0}: {1}", ActionString, ex1.Message))
                                        End Try
                                    End If
                                End If
                        End Select
                    Else
                        LogMsg("SupportedActions", MessageLevel.msgError, "Actions must be strings. The type of action " & i & " " & Action.ToString() & " is: " & Action.GetType.Name)
                    End If
                Next
            End If
        Catch ex As Exception
            If (p_DeviceType = DeviceType.Switch) And (ex.GetType Is GetType(MissingMemberException)) Then
                LogMsg("SupportedActions", MessageLevel.msgOK, "Switch V1 Driver does not have SupportedActions")
            Else
                HandleException("SupportedActions", MemberType.Property, Required.Optional, ex, "")
                LogMsg("SupportedActions", MessageLevel.msgError, ex.Message)
            End If
        End Try
        If TestStop() Then Exit Sub
        LogMsg("", MessageLevel.msgAlways, "")
    End Sub
    Overridable Overloads Sub CheckCommonMethods()
        LogMsg("CheckCommonMethods", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub

    Overridable Property HasCanProperties() As Boolean
        Get
            Return l_HasCanProperties
        End Get
        Set(ByVal value As Boolean)
            l_HasCanProperties = value
        End Set
    End Property
    Overridable Property HasProperties() As Boolean
        Get
            Return l_HasProperties
        End Get
        Set(ByVal value As Boolean)
            l_HasProperties = value
        End Set
    End Property
    Overridable Property HasMethods() As Boolean
        Get
            Return l_HasMethods
        End Get
        Set(ByVal value As Boolean)
            l_HasMethods = value
        End Set
    End Property
    Overridable Property HasPreConnectCheck() As Boolean
        Get
            Return l_HasPreConnectCheck
        End Get
        Set(ByVal value As Boolean)
            l_HasPreConnectCheck = value
        End Set
    End Property
    Overridable Property HasPreRunCheck() As Boolean
        Get
            Return l_HasPreRunCheck
        End Get
        Set(ByVal value As Boolean)
            l_HasPreRunCheck = value
        End Set
    End Property
    Overridable Property HasPostRunCheck() As Boolean
        Get
            Return l_HasPostRunCheck
        End Get
        Set(ByVal value As Boolean)
            l_HasPostRunCheck = value
        End Set
    End Property
    Overridable Property HasPerformanceCheck() As Boolean
        Get
            Return l_HasPerformanceCheck
        End Get
        Set(ByVal value As Boolean)
            l_HasPerformanceCheck = value
        End Set
    End Property

    Overridable Sub CheckInitialise()
        LogMsg("ConformanceCheckInitialise", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub

    ''' <summary>
    ''' Get error codes.
    ''' </summary>
    ''' <param name="p_ProgID">The p_ prog ID.</param>
    Protected Sub CheckInitialise(ByVal p_ProgID As String)
        Dim Prof As ASCOM.Utilities.RegistryAccess
        Prof = New ASCOM.Utilities.RegistryAccess

        Status(StatusType.staTest, "") 'Clear status messages
        Status(StatusType.staAction, "")
        Status(StatusType.staStatus, "")

        g_Stop = True 'Initialise stop flag to stop

        Dim assembly As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim fileInfo As System.IO.FileInfo = New System.IO.FileInfo(assembly.Location)
        Dim lastModified As DateTime = fileInfo.LastWriteTime

        LogMsg("", MessageLevel.msgAlways, "") 'Blank line
        LogMsg("ConformanceCheck", MessageLevel.msgAlways, "ASCOM Device Conformance Checker Version " & My.Application.Info.Version.ToString & ", Build time: " & lastModified.ToString())
        LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Running on: " & Prof.GetProfile("Platform", "Platform Name", "Unknown") & " " & Prof.GetProfile("Platform", "Platform Version", "Unknown"))
        Prof.Dispose()
        LogMsg("", MessageLevel.msgAlways, "") 'Blank line
        LogMsg("ConformanceCheck", MessageLevel.msgAlways, DRIVER_PROGID & p_ProgID)
        LogMsg("", MessageLevel.msgAlways, "") 'Blank line

        LogMsg("Error handling", MessageLevel.msgAlways, "")
        LogMsg("Error", MessageLevel.msgAlways, "number for ""Not Implemented"" is: " & Hex(g_ExNotImplemented))
        LogMsg("Error", MessageLevel.msgAlways, "number for ""Invalid Value 1"" is: " & Hex(g_ExInvalidValue1))
        If (g_ExInvalidValue2 <> 0) And (g_ExInvalidValue2 <> g_ExInvalidValue1) Then LogMsg("Error", MessageLevel.msgAlways, "number for ""Invalid Value 2"" is: " & Hex(g_ExInvalidValue2))
        If (g_ExInvalidValue3 <> 0) And (g_ExInvalidValue3 <> g_ExInvalidValue2) Then LogMsg("Error", MessageLevel.msgAlways, "number for ""Invalid Value 3"" is: " & Hex(g_ExInvalidValue3))
        If (g_ExInvalidValue4 <> 0) And (g_ExInvalidValue4 <> g_ExInvalidValue3) Then LogMsg("Error", MessageLevel.msgAlways, "number for ""Invalid Value 4"" is: " & Hex(g_ExInvalidValue4))
        If (g_ExInvalidValue5 <> 0) And (g_ExInvalidValue5 <> g_ExInvalidValue4) Then LogMsg("Error", MessageLevel.msgAlways, "number for ""Invalid Value 5"" is: " & Hex(g_ExInvalidValue5))
        If (g_ExInvalidValue6 <> 0) And (g_ExInvalidValue6 <> g_ExInvalidValue5) Then LogMsg("Error", MessageLevel.msgAlways, "number for ""Invalid Value 6"" is: " & Hex(g_ExInvalidValue6))
        LogMsg("Error", MessageLevel.msgAlways, "number for ""Value Not Set 1"" is: " & Hex(ErrorCodes.ValueNotSet))
        LogMsg("Error", MessageLevel.msgAlways, "number for ""Value Not Set 2"" is: " & Hex(g_ExNotSet1))
        If (g_ExNotSet2 <> 0) And (g_ExNotSet2 <> g_ExNotSet1) Then LogMsg("Error", MessageLevel.msgAlways, "number for ""Value Not Set 3"" is: " & Hex(g_ExNotSet2))
        If g_Settings.InterpretErrorMessages Then
            LogMsg("Error", MessageLevel.msgAlways, "messages will be interpreted to infer state.")
        Else
            LogMsg("Error", MessageLevel.msgAlways, "messages will not be interpreted to infer state.")
        End If
        LogMsg("", MessageLevel.msgAlways, "")
    End Sub
    Overridable Sub CheckAccessibility()
        LogMsg("ConformanceCheckAccessibility", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub
    Protected Sub CheckAccessibility(ByVal p_ProgId As String, ByVal p_DeviceType As DeviceType)
        Dim l_DeviceObject As Object, l_Type As System.Type
        Dim l_TryCount As Integer, l_ErrMsg As String = ""

        LogMsg("Driver Access Checks", MessageLevel.msgOK, "")

        'Try late binding as an object
        l_DeviceObject = Nothing
        Do
            l_TryCount += 1
            Try
#If DEBUG Then
                If g_Settings.DisplayMethodCalls Then LogMsg("CreateObject", MessageLevel.msgComment, "About to create instance using CreateObject")
                l_DeviceObject = CreateObject(p_ProgId)
                LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using CreateObject")
#Else
                l_Type = Type.GetTypeFromProgID(p_ProgId)
                If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to create instance using Activator.CreateInstance")
                l_DeviceObject = Activator.CreateInstance(l_Type)
                LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using Activator.CreateInstance")
#End If
                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver initialisation")
                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using late binding")
                Try
                    Select Case p_DeviceType
                        Case DeviceType.Focuser 'Focuser uses link to connect
                            If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Link property true")
                            l_DeviceObject.Link = True
                            If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Link property false")
                            l_DeviceObject.Link = False
                        Case Else 'Everything else uses connect!
                            If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property true")
                            l_DeviceObject.Connected = True
                            If g_Settings.DisplayMethodCalls Then LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property false")
                            Try
                                l_DeviceObject.Connected = False
                            Catch ex As Exception
                                LogMsg("AccessChecks", MessageLevel.msgError, "Error disconnecting from late bound driver: " & ex.Message)
                                LogMsg("AccessChecks", MessageLevel.msgDebug, "Exception: " & ex.ToString)
                            End Try
                    End Select
                    LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using late binding")

                    Try
                        If l_DeviceObject.GetType.IsCOMObject Then
                            LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver is a COM object")
                        Else
                            LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver is a .NET object")
                            LogMsg("AccessChecks", MessageLevel.msgInfo, "The AssemblyQualifiedName is: " & Microsoft.VisualBasic.Left(l_DeviceObject.GetType.AssemblyQualifiedName.ToString, 76))
                        End If
                        For Each l_Type In l_DeviceObject.GetType.GetInterfaces()
                            LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver implements interface: " & l_Type.FullName)
                        Next
                    Catch ex As Exception
                        LogMsg("AccessChecks", MessageLevel.msgError, "Error reading driver characteristsics: " & ex.Message)
                        LogMsg("", MessageLevel.msgAlways, "")
                    End Try
                Catch ex As Exception
                    LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using late binding: " & ex.ToString)
                    LogMsg("", MessageLevel.msgAlways, "")
                End Try
            Catch ex As Exception
                l_ErrMsg = ex.ToString
                LogMsg("DeviceTesterBaseClass", MessageLevel.msgDebug, l_ErrMsg)
            End Try
            If l_DeviceObject Is Nothing Then WaitFor(200)
        Loop Until (l_TryCount = 3) Or (Not (l_DeviceObject Is Nothing))

        If l_DeviceObject Is Nothing Then
            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver object using late binding: " & l_ErrMsg)
            LogMsg("", MessageLevel.msgAlways, "")
        End If

        'Clean up
        Try
            ReleaseCOMObjects("AccessChecks", l_DeviceObject)
        Catch ex As Exception
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Error releasing driver object using ReleaseCOMObject: " & ex.ToString)
        End Try
        l_DeviceObject = Nothing
        LogMsg("AccessChecks", MessageLevel.msgDebug, "Collecting garbage")
        GC.Collect()
        LogMsg("AccessChecks", MessageLevel.msgDebug, "Collecting garbage complete")
        GC.WaitForPendingFinalizers()
        LogMsg("AccessChecks", MessageLevel.msgDebug, "Finished waiting for pending finalisers")
        WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system") 'Wait to allow device to complete destruction
    End Sub

    Overridable Sub CreateDevice()
        LogMsg("CreateDevice", MessageLevel.msgError, "DeviceTester base Class warning messagev")
    End Sub
    Overridable Sub PreConnectChecks()
        LogMsg("PreConnectChecks", MessageLevel.msgError, "DeviceTester base Class warning messagev")
    End Sub

    Overridable Property Connected() As Boolean
        Get
            Connected = l_Connected
        End Get
        Set(ByVal value As Boolean)
            l_Connected = value
        End Set
    End Property

    Overridable Sub ReadCanProperties()
        LogMsg("ReadCanProperties", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub
    Overridable Sub PreRunCheck()
        LogMsg("PreSafetyCheck", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub

    Overridable Sub CheckProperties()
        LogMsg("CheckProperties", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub
    Overridable Sub CheckMethods()
        LogMsg("CheckMethods", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub
    Overridable Sub CheckPerformance()
        LogMsg("CheckPerformance", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub

    Overridable Sub PostRunCheck()
        LogMsg("PostSafetyCheck", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub

    Sub SpecialTests(ByVal p_Test As SpecialTests)
        Select Case p_Test
            Case GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier
                SpecialTelescopeSideOfPier()
            Case GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier
                SpecialTelescopeDestinationSideOfPier()
            Case GlobalVarsAndCode.SpecialTests.TelescopeSideOfPierAnalysis
                SpecialTelescopeSideOfPierAnalysis()
            Case GlobalVarsAndCode.SpecialTests.TelescopeCommands
                SpecialTelescopeCommands()
            Case Else
                MsgBox("Devicetester BaseClass:SpecialTests - Unknown test: " & p_Test.ToString, MsgBoxStyle.Critical)
        End Select
    End Sub
    Protected Overridable Sub SpecialTelescopeSideOfPier()
        LogMsg("SpecialTelescopeSideOfPier", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub
    Protected Overridable Sub SpecialTelescopeDestinationSideOfPier()
        LogMsg("SpecialTelescopeDestinationSideOfPier", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub
    Protected Overridable Sub SpecialTelescopeSideOfPierAnalysis()
        LogMsg("SpecialTelescopeFlipRange", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub
    Protected Overridable Sub SpecialTelescopeCommands()
        LogMsg("SpecialTelescopeCommands", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!")
    End Sub

#End Region

#Region "Base class support Code"

    Private Function IncludeMethod(ByVal p_Method As MandatoryMethod, ByVal p_DeviceType As DeviceType, ByVal p_InterfaceVersion As Integer) As Boolean
        'This mechanic allows individual tests for particular devices to be skipped. It is no longer required because this is handled by DriverAccess
        'The code is left in place in case it is ever needed in the future

        Dim RetVal As Boolean = True 'Default to true as most methods will be tested , we just list the exceptions to this below

        'Matrix controlling what tests
        Select Case p_DeviceType
            Case DeviceType.Telescope
                Select Case p_InterfaceVersion
                    Case 1 ' Telescope interface V1 does not have Driver Version
                        If p_Method = MandatoryMethod.DriverVersion Then RetVal = False
                    Case Else
                        RetVal = True 'All methods in all interface versions are mandatory
                End Select
            Case DeviceType.Camera
                RetVal = True
        End Select

        Return RetVal
    End Function

    Private Sub CommandTest(ByVal p_Type As CommandType, ByVal p_Name As String)
        Dim l_CmdSent As String = "!! Unknown Value !!"
        Dim m_CommandString As String
        Dim m_CommandBool As Boolean


        g_Status.Test = p_Name
        Try
            Select Case p_Type
                Case CommandType.tstCommandBlind
                    If Not g_CmdStrings.CommandBlind Is Nothing Then
                        l_CmdSent = g_CmdStrings.CommandBlind
#If PLATFORM6 Then
                        Device.CommandBlind(l_CmdSent, False)
#Else
                        Device.CommandBlind(l_CmdSent)
#End If

                        LogMsg(p_Name, MessageLevel.msgOK, "Sent string """ & g_CmdStrings.CommandBlind & """")
                    Else
                        LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBlind test")
                    End If
                Case CommandType.tstCommandBool
                    If Not g_CmdStrings.CommandBool Is Nothing Then
                        l_CmdSent = g_CmdStrings.CommandBool
#If PLATFORM6 Then

                        m_CommandBool = Device.CommandBool(l_CmdSent, False)
#Else
                        m_CommandBool = Device.CommandBool(l_CmdSent)
#End If
                        If m_CommandBool = g_CmdStrings.ReturnBool Then
                            LogMsg(p_Name, MessageLevel.msgOK, "Sent string """ & g_CmdStrings.CommandBool & """ - Received expected return value: " & m_CommandBool)
                        Else
                            LogMsg(p_Name, MessageLevel.msgError, "Sent string """ & g_CmdStrings.CommandBool & """ - Received unexpected return value: " & m_CommandBool)
                        End If
                    Else
                        LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBool test")
                    End If
                Case CommandType.tstCommandString
                    If Not g_CmdStrings.CommandString Is Nothing Then
                        l_CmdSent = g_CmdStrings.CommandString
#If PLATFORM6 Then
                        m_CommandString = Device.CommandString(l_CmdSent, False)
#Else
                        m_CommandString = Device.CommandString(l_CmdSent)
#End If
                        If Not g_CmdStrings.ReturnString Is Nothing Then 'Valid return string toi test
                            If m_CommandString = g_CmdStrings.ReturnString Then
                                LogMsg(p_Name, MessageLevel.msgOK, "Sent string """ & g_CmdStrings.CommandString & """ - Received string: """ & m_CommandString & """")
                            Else
                                LogMsg(p_Name, MessageLevel.msgError, "Sent string """ & g_CmdStrings.CommandString & """ - Received string: """ & m_CommandString & """ - Expected string: """ & g_CmdStrings.ReturnString & """")
                            End If
                        Else 'Skip the return string test
                            LogMsg(p_Name, MessageLevel.msgOK, "Sent string """ & g_CmdStrings.CommandString & """ - Return string test skipped")
                        End If
                    Else
                        LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandString test")
                    End If
                Case CommandType.tstCommandBlindRaw
                    If Not g_CmdStringsRaw.CommandBlind Is Nothing Then
                        l_CmdSent = g_CmdStringsRaw.CommandBlind
                        Device.CommandBlind(l_CmdSent, True)
                        LogMsg(p_Name, MessageLevel.msgOK, "Sent string """ & g_CmdStringsRaw.CommandBlind & """")
                    Else
                        LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBlind Raw test")
                    End If
                Case CommandType.tstCommandBoolRaw
                    If Not g_CmdStringsRaw.CommandBool Is Nothing Then
                        l_CmdSent = g_CmdStringsRaw.CommandBool
                        m_CommandBool = Device.CommandBool(l_CmdSent, True)
                        If m_CommandBool = g_CmdStringsRaw.ReturnBool Then
                            LogMsg(p_Name, MessageLevel.msgOK, "Sent string """ & g_CmdStringsRaw.CommandBool & """ - Received expected return value: " & m_CommandBool)
                        Else
                            LogMsg(p_Name, MessageLevel.msgError, "Sent string """ & g_CmdStringsRaw.CommandBool & """ - Received unexpected return value: " & m_CommandBool)
                        End If
                    Else
                        LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBool Raw test")
                    End If
                Case CommandType.tstCommandStringRaw
                    If Not g_CmdStringsRaw.CommandString Is Nothing Then
                        l_CmdSent = g_CmdStringsRaw.CommandString
                        m_CommandString = Device.CommandString(l_CmdSent, True)
                        If Not g_CmdStringsRaw.ReturnString Is Nothing Then 'Valid return string to test
                            If m_CommandString = g_CmdStringsRaw.ReturnString Then
                                LogMsg(p_Name, MessageLevel.msgOK, "Sent string """ & g_CmdStringsRaw.CommandString & """ - Received string: """ & m_CommandString & """")
                            Else
                                LogMsg(p_Name, MessageLevel.msgError, "Sent string """ & g_CmdStringsRaw.CommandString & """ - Received string: """ & m_CommandString & """ - Expected string: """ & g_CmdStringsRaw.ReturnString & """")
                            End If
                        Else 'Skip the return string test
                            LogMsg(p_Name, MessageLevel.msgOK, "Sent string """ & g_CmdStringsRaw.CommandString & """ - Return string test skipped")
                        End If
                    Else
                        LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandString Raw test")
                    End If

                Case Else
                    LogMsg(p_Name, MessageLevel.msgError, "Conform:CommandTest: Unknown test type " & p_Type.ToString)
            End Select
        Catch ex As Exception
            HandleException(p_Name, MemberType.Method, Required.Optional, ex, "")
        End Try
        g_Status.Clear()  'Clear status messages
    End Sub

#End Region

#Region "Common methods for all device tester classes"
    Public Shared Sub ReleaseCOMObjects(testName As String, ByRef ObjectToRelease As Object)
        Dim ObjectType As Type, RemainingObjectCount, LoopCount As Integer

        LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Start of ReleaseCOMObject")
        If g_Settings.DisplayMethodCalls Then LogMsg(testName, MessageLevel.msgComment, "About to release driver instance")

        Try
            ObjectType = ObjectToRelease.GetType
            LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Unmarshalling " & ObjectType.Name)
        Catch ex1 As Exception
            LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  GetType Exception: " & ex1.Message)
        End Try

        Try
            LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Releasing COM object")
            LoopCount = 0
            Do
                LoopCount += 1
                RemainingObjectCount = Marshal.ReleaseComObject(ObjectToRelease)
                LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Remaining object count: " & RemainingObjectCount & ", LoopCount: " & LoopCount)
            Loop Until (RemainingObjectCount <= 0) Or (LoopCount = 20)
        Catch ex2 As Exception
            LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  ReleaseComObject Exception: " & ex2.Message)
        End Try


        Try
            ObjectToRelease = Nothing
        Catch ex3 As Exception
            LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Set to nothing Exception: " & ex3.Message)
        End Try
        LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  End of ReleaseCOMObject")
    End Sub

    ''' <summary>
    ''' Test a supplied exception for whether it is a MethodNotImplemented type
    ''' </summary>
    ''' <param name="deviceException">The exception sent by the driver</param>
    ''' <returns>True if the exception is a NotImplemented type</returns>
    ''' <remarks>Different tests are applied for COM and MethodNotImplemmented exceptions</remarks>
    Protected Function IsMethodNotImplementedException(deviceException As Exception) As Boolean
        Dim COMException As COMException

        IsMethodNotImplementedException = False ' Set false default value

        Try
            If (TypeOf deviceException Is COMException) Then ' This is a COM exception so test whether the error code indciates that it is a not implemented exception
                COMException = CType(deviceException, COMException)
                If (COMException.ErrorCode = g_ExNotImplemented) Or (COMException.ErrorCode = ErrorCodes.NotImplemented) Then ' This is a not implemented exception
                    IsMethodNotImplementedException = True
                End If
            End If
            If (TypeOf deviceException Is MethodNotImplementedException) Then
                IsMethodNotImplementedException = True
            End If
        Catch ex As Exception
            LogMsg("IsMethodNotImplementedException", MessageLevel.msgWarning, "Unexpected exception: " & ex.ToString())
        End Try
    End Function

    ''' <summary>
    ''' Test a supplied exception for whether it is a PropertyNotImplementedException type
    ''' </summary>
    ''' <param name="deviceException">The exception sent by the driver</param>
    ''' <returns>True if the exception is a NotImplemented type</returns>
    ''' <remarks>Different tests are applied for COM and PropertyNotImplemmented exceptions</remarks>
    Protected Function IsPropertyNotImplementedException(deviceException As Exception) As Boolean
        Dim COMException As COMException

        IsPropertyNotImplementedException = False ' Set false default value

        Try
            If (TypeOf deviceException Is COMException) Then ' This is a COM exception so test whether the error code indciates that it is a not implemented exception
                COMException = CType(deviceException, COMException)
                If (COMException.ErrorCode = g_ExNotImplemented) Or (COMException.ErrorCode = ErrorCodes.NotImplemented) Then ' This is a not implemented exception
                    IsPropertyNotImplementedException = True
                End If
            End If
            If (TypeOf deviceException Is PropertyNotImplementedException) Then
                IsPropertyNotImplementedException = True
            End If
        Catch ex As Exception
            LogMsg("IsPropertyNotImplementedException", MessageLevel.msgWarning, "Unexpected exception: " & ex.ToString())
        End Try
    End Function

    ''' <summary>
    ''' Test a supplied exception for whether it is an InvalidValueException type
    ''' </summary>
    ''' <param name="deviceException">The exception sent by the driver</param>
    ''' <returns>True if the exception is a InvalidValue type</returns>
    ''' <remarks>Different tests are applied for COM and InvalidValueException exceptions</remarks>
    Protected Function IsInvalidValueException(MemberName As String, deviceException As Exception) As Boolean
        Dim COMException As COMException, DriverException As DriverException

        IsInvalidValueException = False ' Set false default value

        Try
            If (TypeOf deviceException Is COMException) Then ' This is a COM exception so test whether the error code indciates that it is an invalid value exception
                COMException = CType(deviceException, COMException)
                If (COMException.ErrorCode = ErrorCodes.InvalidValue) Or
                   (COMException.ErrorCode = g_ExInvalidValue1) Or
                   (COMException.ErrorCode = g_ExInvalidValue2) Or
                   (COMException.ErrorCode = g_ExInvalidValue3) Or
                   (COMException.ErrorCode = g_ExInvalidValue4) Or
                   (COMException.ErrorCode = g_ExInvalidValue5) Or
                   (COMException.ErrorCode = g_ExInvalidValue6) Then ' This is an invalid value exception
                    IsInvalidValueException = True
                End If
            End If

            If (TypeOf deviceException Is InvalidValueException) Then
                IsInvalidValueException = True
            ElseIf (TypeOf deviceException Is DriverException) Then
                DriverException = CType(deviceException, DriverException)
                If (DriverException.Number = ErrorCodes.InvalidValue) Then ' This is an invalid value exception
                    LogMsg(MemberName, MessageLevel.msgIssue, "Received ASCOM.DriverException(0x" & ErrorCodes.InvalidValue.ToString("X8") & "), please use ASCOM.InvalidValueException to report invalid values")
                End If
            End If

            If (TypeOf deviceException Is System.InvalidOperationException) Then
                LogMsg(MemberName, MessageLevel.msgIssue, "Received System.InvalidOperationException rather than ASCOM.InvalidValueException")
            End If

        Catch ex As Exception
            LogMsg("IsInvalidValueException", MessageLevel.msgWarning, "Unexpected exception: " & ex.ToString())
        End Try

        Return IsInvalidValueException
    End Function

    ''' <summary>
    ''' Test a supplied exception for whether it is an InvalidValueException type
    ''' </summary>
    ''' <param name="deviceException">The exception sent by the driver</param>
    ''' <returns>True if the exception is a InvalidValue type</returns>
    ''' <remarks>Different tests are applied for COM and InvalidValueException exceptions</remarks>
    Protected Function IsInvalidOperationException(MemberName As String, deviceException As Exception) As Boolean
        Dim COMException As COMException, DriverException As DriverException

        IsInvalidOperationException = False ' Set false default value

        Try
            If (TypeOf deviceException Is COMException) Then ' This is a COM exception so test whether the error code indciates that it is an invalid operation exception
                COMException = CType(deviceException, COMException)
                If (COMException.ErrorCode = ErrorCodes.InvalidOperationException) Then ' This is an invalid operation exception
                    IsInvalidOperationException = True
                End If
            End If

            If (TypeOf deviceException Is InvalidOperationException) Then
                IsInvalidOperationException = True
            ElseIf (TypeOf deviceException Is DriverException) Then
                DriverException = CType(deviceException, DriverException)
                If (DriverException.Number = ErrorCodes.InvalidOperationException) Then ' This is an invalid operation exception
                    LogMsg(MemberName, MessageLevel.msgIssue, "Received ASCOM.DriverException(0x" & ErrorCodes.InvalidOperationException.ToString("X8") & "), please use ASCOM.InvalidOperationException to report invalid operations")
                End If
            End If

            If (TypeOf deviceException Is System.InvalidOperationException) Then
                LogMsg(MemberName, MessageLevel.msgIssue, "Received System.InvalidOperationException rather than ASCOM.InvalidOperationException")
            End If

        Catch ex As Exception
            LogMsg("IsInvalidOperationException", MessageLevel.msgWarning, "Unexpected exception: " & ex.ToString())
        End Try

        Return IsInvalidOperationException
    End Function

    ''' <summary>
    ''' Test a supplied exception for whether it is a NotSetException type
    ''' </summary>
    ''' <param name="deviceException">The exception sent by the driver</param>
    ''' <returns>True if the exception is a NotSet type</returns>
    ''' <remarks>Different tests are applied for COM and ValueNotSetException exceptions</remarks>
    Protected Function IsNotSetException(deviceException As Exception) As Boolean
        Dim COMException As COMException

        IsNotSetException = False ' Set false default value

        Try
            If (TypeOf deviceException Is COMException) Then ' This is a COM exception so test whether the error code indciates that it is a not set exception
                COMException = CType(deviceException, COMException)
                If (COMException.ErrorCode = g_ExNotSet1) Then ' This is a not set exception
                    IsNotSetException = True
                End If
            End If
            If (TypeOf deviceException Is ValueNotSetException) Then
                IsNotSetException = True
            End If
        Catch ex As Exception
            LogMsg("IsNotSetException", MessageLevel.msgWarning, "Unexpected exception: " & ex.ToString())
        End Try
    End Function

    ''' <summary>
    ''' Provides messaging when an exception is thrown by a member
    ''' </summary>
    ''' <param name="MemberName">The name of the member throwing the exception</param>
    ''' <param name="TypeOfMember">Flag indicating whether the member is a property or a method</param>
    ''' <param name="IsRequired">Flag indicating whether the member is optional or mandatory</param>
    ''' <param name="ex">The exception received from the device</param>
    ''' <param name="UserMessage">The member specific message to report</param>
    ''' <remarks></remarks>
    Protected Sub HandleException(MemberName As String, TypeOfMember As MemberType, IsRequired As Required, ex As Exception, UserMessage As String)

        ' Handle PropertyNotImplemented exceptions from properties and MethodNotImplemented exceptions from methods
        If (IsPropertyNotImplementedException(ex) And TypeOfMember = MemberType.Property) Or (IsMethodNotImplementedException(ex) And TypeOfMember = MemberType.Method) Then
            Select Case IsRequired
                Case Required.Mandatory
                    LogMsg(MemberName, MessageLevel.msgIssue, "This member is mandatory but threw a " & GetExceptionName(ex) & " exception, it must function per the ASCOM specification.")
                Case Required.MustNotBeImplemented
                    LogMsg(MemberName, MessageLevel.msgOK, UserMessage & " and a " & GetExceptionName(ex) & " exception was generated as expected")
                Case Required.MustBeImplemented
                    LogMsg(MemberName, MessageLevel.msgIssue, UserMessage & " and a " & GetExceptionName(ex) & " exception was thrown, this method must must function per the ASCOM specification.")
                Case Required.Optional
                    LogMsg(MemberName, MessageLevel.msgOK, "Optional member threw a " & GetExceptionName(ex) & " exception.")
                Case Else
                    LogMsg(MemberName, MessageLevel.msgError, "CONFORM ERROR! - Received unexpected member of 'Required' enum: " + IsRequired)
            End Select

            ' Handle wrong type of not implemnented exceptions
        ElseIf (TypeOf ex Is MethodNotImplementedException) And TypeOfMember = MemberType.Property Then ' We got a MethodNotImplementedException so this is an error
            LogMsg(MemberName, MessageLevel.msgIssue, "Received a MethodNotImplementedException instead of a PropertyNotImplementedException")
        ElseIf (TypeOf ex Is PropertyNotImplementedException) And TypeOfMember = MemberType.Method Then ' We got a PropertyNotImplementedException so this is an error
            LogMsg(MemberName, MessageLevel.msgIssue, "Received a PropertyNotImplementedException instead of a MethodNotImplementedException")
        ElseIf (TypeOf ex Is NotImplementedException) Then
            LogMsg(MemberName, MessageLevel.msgIssue, "Received a NotImplementedException instead of a " & IIf(TypeOfMember = MemberType.Property, "PropertyNotImplementedException", "MethodNotImplementedException"))
        ElseIf (TypeOf ex Is System.NotImplementedException) Then
            LogMsg(MemberName, MessageLevel.msgIssue, "Received a System.NotImplementedException instead of an ASCOM." & IIf(TypeOfMember = MemberType.Property, "PropertyNotImplementedException", "MethodNotImplementedException"))

            ' Handle all other types of error
        Else
            LogMsg(MemberName, MessageLevel.msgError, "Unexpected " & GetExceptionName(ex) & ", " & UserMessage & ": " & ex.Message)
        End If

        LogMsg(MemberName, MessageLevel.msgDebug, "Exception: " & ex.ToString)

    End Sub

    Protected Sub HandleInvalidValueExceptionAsOK(MemberName As String, TypeOfMember As MemberType, IsRequired As Required, ex As Exception, UserAction As String, Message As String)
        If IsInvalidValueException(MemberName, ex) Then
            LogMsg(MemberName, MessageLevel.msgOK, Message)
        Else
            HandleException(MemberName, TypeOfMember, IsRequired, ex, UserAction)
        End If
    End Sub

    Protected Sub HandleInvalidValueExceptionAsInfo(MemberName As String, TypeOfMember As MemberType, IsRequired As Required, ex As Exception, UserAction As String, Message As String)
        If IsInvalidValueException(MemberName, ex) Then
            LogMsg(MemberName, MessageLevel.msgInfo, Message)
        Else
            HandleException(MemberName, TypeOfMember, IsRequired, ex, UserAction)
        End If
    End Sub

    Protected Sub HandleInvalidOperationExceptionAsOK(MemberName As String, TypeOfMember As MemberType, IsRequired As Required, ex As Exception, UserAction As String, Message As String)
        If IsInvalidOperationException(MemberName, ex) Then
            LogMsg(MemberName, MessageLevel.msgOK, Message)
        Else
            HandleException(MemberName, TypeOfMember, IsRequired, ex, UserAction)
        End If
    End Sub

    ''' <summary>
    ''' Get an exception name (and number if a COM or Driver execption)
    ''' </summary>
    ''' <param name="ex">Exception whose name is required</param>
    ''' <returns>String exception name</returns>
    ''' <remarks></remarks>
    Protected Function GetExceptionName(ex As Exception) As String
        Dim ComEx As COMException, DriverEx As DriverException, RetVal As String

        ' Treat ASCOM exceptions specially
        If ex.GetType.FullName.ToUpper().Contains("ASCOM") Then
            If ex.GetType.FullName.ToUpper().Contains("DRIVEREXCEPTION") Then ' We have a driver execption so add its number
                DriverEx = CType(ex, DriverException)
                RetVal = "DriverException(0x" & DriverEx.Number.ToString("X8") & ")"
            Else ' Otherwise just use the ASCOM exception's name
                RetVal = ex.GetType.Name
            End If
        ElseIf (TypeOf ex Is COMException) Then ' Handle XOM exceptions with their error code
            ComEx = CType(ex, COMException)
            RetVal = "COMException(0x" & ComEx.ErrorCode.ToString("X8") & ")"
        Else ' We got something else so report it
            RetVal = ex.GetType().FullName & " exception"
        End If
        Return RetVal
    End Function

#End Region
End Class

