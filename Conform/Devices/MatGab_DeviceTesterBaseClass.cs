//Base class from which particular device testers are derrived
//Put all common elements in here
using ASCOM.DeviceInterface;
using System.Threading;
using System.Runtime.InteropServices;

/// <summary>
/// Base class for device tester classes. Contains common code and placeholders for the 
/// methods that must be implemented in the device tester class
/// </summary>
internal class DeviceTesterBaseClass : IDisposable
{
	bool l_Connected;
	bool l_HasProperties;
	bool l_HasCanProperties;
	bool l_HasMethods;
	bool l_HasPreSafetyCheck;
	bool l_HasPostSafetyCheck;
	bool l_HasPerformanceCheck;

	bool l_HasPreConnectCheck;
	#Region "Variables and Constants"
		//IAscomDriverV1
	object Device;
	#End Region

	#Region "Enums"
	private enum CommandType
	{
		tstCommandString = 1,
		tstCommandBool = 2,
		tstCommandBlind = 3,
		tstCommandStringRaw = 4,
		tstCommandBoolRaw = 5,
		tstCommandBlindRaw = 6
	}
	#End Region

	#Region "New and Dispose"
	private DeviceTesterBaseClass()
	{
		base.New();
		l_HasPreConnectCheck = false;
		l_Connected = false;
		l_HasPreSafetyCheck = false;
		l_HasCanProperties = false;
		l_HasProperties = true;
		l_HasMethods = true;
		l_HasPostSafetyCheck = false;
		l_HasPerformanceCheck = true;
	}

		// To detect redundant calls
	private bool disposedValue = false;
	// IDisposable
	protected virtual void Dispose(bool disposing)
	{
		if (!this.disposedValue) {
			if (disposing) {
				LogMsg("Dispose", MessageLevel.msgError, "DeviceTester base Class warning messagev");
				// TODO: free other state (managed objects).
			}

			// TODO: free your own state (unmanaged objects).
			// TODO: set large fields to null.
		}
		this.disposedValue = true;
	}

	// This code added by Visual Basic to correctly implement the disposable pattern.
	public void IDisposable.Dispose()
	{
		// Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
		Dispose(true);
		GC.SuppressFinalize(this);
	}
	#End Region

	#Region "Code"

	private void CheckCommonMethods(object p_DeviceObject, DeviceType p_DeviceType)
	{
		string m_DriverVersion;
		string m_DriverInfo;
		string m_Description;
		string m_Name;
		//, m_LastResult As String
		ArrayList SA;

		//Dim m_Configuration, SC() As String
		bool m_Connected;
		bool l_HasInterfaceVersion = false;

		LogMsg("Common Driver Methods", MessageLevel.msgAlways, "");
		Device = p_DeviceObject;
		//CType(DeviceObject, IAscomDriverV1)

		//InterfaceVersion - Required
		try {
			g_InterfaceVersion = Device.InterfaceVersion;
			switch (g_InterfaceVersion) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
1:
					LogMsg("InterfaceVersion", MessageLevel.msgIssue, "InterfaceVersion must be 1 or greater but driver returned: " + g_InterfaceVersion.ToString);
				default:
					LogMsg("InterfaceVersion", MessageLevel.msgOK, g_InterfaceVersion.ToString);
			}
			l_HasInterfaceVersion = true;
		} catch (COMException ex) {
			LogMsg("InterfaceVersion", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (PropertyNotImplementedException ex) {
			//Method not implemented so assume we are dealing wqith a pre common-methods driver
			LogMsg("", MessageLevel.msgError, "InterfaceVersion can not throw a PropertyNotImplementedException - is a mandatory parameter. Skipping tests of other common methods");
		} catch (Exception ex) {
			LogMsg("InterfaceVersion", MessageLevel.msgError, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		if (l_HasInterfaceVersion) {
			//Connected - Required
			if (IncludeMethod(MandatoryMethod.Connected, p_DeviceType, g_InterfaceVersion)) {
				try {
					m_Connected = Device.Connected;
					LogMsg("Connected", MessageLevel.msgOK, m_Connected.ToString);
				} catch (COMException ex) {
					LogMsg("Connected", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("Connected", MessageLevel.msgError, EX_NET + ex.Message);
				}
				if (TestStop())
					return;
			}

			//Description - Required
			if (IncludeMethod(MandatoryMethod.Description, p_DeviceType, g_InterfaceVersion)) {
				try {
					m_Description = Device.Description;
					switch (m_Description) {
						case "":
							LogMsg("Description", MessageLevel.msgInfo, "No description string");
						default:
							if ((m_Description.Length > 68) & (p_DeviceType == DeviceType.Camera)) {
								LogMsg("Description", MessageLevel.msgError, "Maximum number of characters is 68 for compatibility with FITS headers, found: " + m_Description.Length + " characters: " + m_Description);
							} else {
								LogMsg("Description", MessageLevel.msgOK, m_Description.ToString);
							}
					}
				} catch (COMException ex) {
					LogMsg("Description", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("Description", MessageLevel.msgError, EX_NET + ex.ToString);
				}
				if (TestStop())
					return;
			}

			//DriverInfo - Required
			if (IncludeMethod(MandatoryMethod.DriverInfo, p_DeviceType, g_InterfaceVersion)) {
				try {
					m_DriverInfo = Device.DriverInfo;
					switch (m_DriverInfo) {
						case "":
							LogMsg("DriverInfo", MessageLevel.msgInfo, "No DriverInfo string");
						default:
							LogMsg("DriverInfo", MessageLevel.msgOK, m_DriverInfo.ToString);
					}
				} catch (COMException ex) {
					LogMsg("DriverInfo", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("DriverInfo", MessageLevel.msgError, EX_NET + ex.Message);
				}
				if (TestStop())
					return;
			}

			//DriverVersion - Required
			if (IncludeMethod(MandatoryMethod.DriverVersion, p_DeviceType, g_InterfaceVersion)) {
				try {
					m_DriverVersion = Device.DriverVersion;
					switch (m_DriverVersion) {
						case "":
							LogMsg("DriverVersion", MessageLevel.msgInfo, "No DriverVersion string");
						default:
							LogMsg("DriverVersion", MessageLevel.msgOK, m_DriverVersion.ToString);
					}
				} catch (COMException ex) {
					LogMsg("DriverVersion", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("DriverVersion", MessageLevel.msgError, EX_NET + ex.Message);
				}
				if (TestStop())
					return;
			} else {
				LogMsg("DriverVersion", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
			}

			//Name - Required
			if (IncludeMethod(MandatoryMethod.Name, p_DeviceType, g_InterfaceVersion)) {
				try {
					m_Name = Device.Name;
					switch (m_Name) {
						case "":
							LogMsg("Name", MessageLevel.msgInfo, "Name is empty");
						default:
							LogMsg("Name", MessageLevel.msgOK, m_Name);
					}
				} catch (COMException ex) {
					LogMsg("Name", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("Name", MessageLevel.msgError, EX_NET + ex.Message);
				}
				if (TestStop())
					return;
			}

			//CommandXXX tests - Optional
			if (IncludeMethod(MandatoryMethod.CommandXXX, p_DeviceType, g_InterfaceVersion)) {
				if ((g_TelescopeTests.Item(TELTEST_COMMANDXXX) == CheckState.Checked)) {
					//CommandTest(CommandType.tstCommandBlind, "CommandBlind") : If TestStop() Then Exit Sub
					//CommandTest(CommandType.tstCommandBool, "CommandBool") : If TestStop() Then Exit Sub
					//CommandTest(CommandType.tstCommandString, "CommandString") : If TestStop() Then Exit Sub
					//CommandTest(CommandType.tstCommandBlindRaw, "CommandBlind Raw") : If TestStop() Then Exit Sub
					//CommandTest(CommandType.tstCommandBoolRaw, "CommandBool Raw") : If TestStop() Then Exit Sub
					//CommandTest(CommandType.tstCommandStringRaw, "CommandString Raw") : If TestStop() Then Exit Sub
					LogMsg("CommandString", MessageLevel.msgInfo, "Conform cannot test the CommandString method");
					LogMsg("CommandBlind", MessageLevel.msgInfo, "Conform cannot test the CommandBlind method");
					LogMsg("CommandBool", MessageLevel.msgInfo, "Conform cannot test the CommandBool method");
				} else {
					LogMsg(TELTEST_COMMANDXXX, MessageLevel.msgInfo, "Tests skipped");
				}
			}

			//Action - optional but cannot be tested
			LogMsg("Action", MessageLevel.msgInfo, "Conform cannot test the Action method");

			//Supported actions - Optional but Required through Driveraccess
			try {
				SA = Device.SupportedActions();
				if (SA.Count == 0) {
					LogMsg("SupportedActions", MessageLevel.msgOK, "Driver returned an empty action list");
				} else {
					int i;
					foreach (object Action in SA) {
						i += 1;
						if (Action.GetType.Name == "String") {
							switch (Action) {
								case "":
									LogMsg("SupportedActions", MessageLevel.msgError, "Supported action " + i + " is an empty string");
								default:
									LogMsg("SupportedActions", MessageLevel.msgOK, Action);
							}
						} else {
							LogMsg("SupportedActions", MessageLevel.msgError, "Actions must be strings. The type of action " + i + " " + Action.ToString + " is: " + Action.GetType.Name);
						}
					}
				}
			} catch (COMException ex) {
				LogMsg("SupportedActions", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				if ((p_DeviceType == DeviceType.Switch) & (object.ReferenceEquals(ex.GetType, typeof(MissingMemberException)))) {
					LogMsg("SupportedActions", MessageLevel.msgOK, "Switch V1 Driver does not have SupportedActions");
				} else {
					LogMsg("SupportedActions", MessageLevel.msgError, EX_NET + ex.ToString);
				}
			}
			if (TestStop())
				return;
		}
		LogMsg("", MessageLevel.msgAlways, "");
	}
	private virtual void CheckCommonMethods()
	{
		LogMsg("CheckCommonMethods", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}

	private virtual bool HasCanProperties {
		get { return l_HasCanProperties; }
		set { l_HasCanProperties = value; }
	}
	private virtual bool HasProperties {
		get { return l_HasProperties; }
		set { l_HasProperties = value; }
	}
	private virtual bool HasMethods {
		get { return l_HasMethods; }
		set { l_HasMethods = value; }
	}
	private virtual bool HasPreConnectCheck {
		get { return l_HasPreConnectCheck; }
		set { l_HasPreConnectCheck = value; }
	}
	private virtual bool HasPreRunCheck {
		get { return l_HasPreSafetyCheck; }
		set { l_HasPreSafetyCheck = value; }
	}
	private virtual bool HasPostRunCheck {
		get { return l_HasPostSafetyCheck; }
		set { l_HasPostSafetyCheck = value; }
	}
	private virtual bool HasPerformanceCheck {
		get { return l_HasPerformanceCheck; }
		set { l_HasPerformanceCheck = value; }
	}

	private virtual void CheckInitialise()
	{
		LogMsg("ConformanceCheckInitialise", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}

	/// <summary>
	/// Reads CommandXXX test confligration information through the IConform interface or returns default values..
	/// </summary>
	/// <param name="p_ProgID">The p_ prog ID.</param>
	protected void CheckInitialise(string p_ProgID)
	{
		ASCOM.Utilities.RegistryAccess Prof;
		Prof = new ASCOM.Utilities.RegistryAccess();

		Status(StatusType.staTest, "");
		//Clear status messages
		Status(StatusType.staAction, "");
		Status(StatusType.staStatus, "");

		g_Stop = true;
		//Initialise stop flag to stop

		LogMsg("", MessageLevel.msgAlways, "");
		//Blank line
		LogMsg("ConformanceCheck", MessageLevel.msgAlways, "ASCOM Device Conformance Checker Version " + My.Application.Info.Version.ToString);
		LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Running on ASCOM Platform " + g_Util.PlatformVersion + " Build: " + Prof.GetProfile("Platform", "Platform Build", "Unknown"));
		Prof.Dispose();
		LogMsg("", MessageLevel.msgAlways, "");
		//Blank line
		LogMsg("ConformanceCheck", MessageLevel.msgAlways, DRIVER_PROGID + p_ProgID);
		LogMsg("", MessageLevel.msgAlways, "");
		//Blank line

		//GetIConformValues() 'Read values from IConform if it is present
		LogMsg("Error handling", MessageLevel.msgAlways, "");
		LogMsg("Error", MessageLevel.msgAlways, "number for \"Not Implemented\" is: " + Hex(g_ExNotImplemented));
		LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 1\" is: " + Hex(g_ExInvalidValue1));
		if ((g_ExInvalidValue2 != 0) & (g_ExInvalidValue2 != g_ExInvalidValue1))
			LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 2\" is: " + Hex(g_ExInvalidValue2));
		if ((g_ExInvalidValue3 != 0) & (g_ExInvalidValue3 != g_ExInvalidValue2))
			LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 3\" is: " + Hex(g_ExInvalidValue3));
		if ((g_ExInvalidValue4 != 0) & (g_ExInvalidValue4 != g_ExInvalidValue3))
			LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 4\" is: " + Hex(g_ExInvalidValue4));
		if ((g_ExInvalidValue5 != 0) & (g_ExInvalidValue5 != g_ExInvalidValue4))
			LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 5\" is: " + Hex(g_ExInvalidValue5));
		if ((g_ExInvalidValue6 != 0) & (g_ExInvalidValue6 != g_ExInvalidValue5))
			LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 6\" is: " + Hex(g_ExInvalidValue6));
		LogMsg("Error", MessageLevel.msgAlways, "number for \"Value Not Set 1\" is: " + Hex(ErrorCodes.ValueNotSet));
		LogMsg("Error", MessageLevel.msgAlways, "number for \"Value Not Set 2\" is: " + Hex(g_ExNotSet1));
		if ((g_ExNotSet2 != 0) & (g_ExNotSet2 != g_ExNotSet1))
			LogMsg("Error", MessageLevel.msgAlways, "number for \"Value Not Set 3\" is: " + Hex(g_ExNotSet2));
		if (g_Settings.InterpretErrorMessages) {
			LogMsg("Error", MessageLevel.msgAlways, "messages will be interpreted to infer state.");
		} else {
			LogMsg("Error", MessageLevel.msgAlways, "messages will not be interpreted to infer state.");
		}
		LogMsg("", MessageLevel.msgAlways, "");
	}
	private virtual void CheckAccessibility()
	{
		LogMsg("ConformanceCheckAccessibility", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}
	protected void CheckAccessibility(string p_ProgId, DeviceType p_DeviceType)
	{
		object l_DeviceObject;
		System.Type l_Type;
		int l_TryCount;
		string l_ErrMsg = "";

		LogMsg("Driver Access Checks", MessageLevel.msgOK, "");

		//Try late binding as an object
		l_DeviceObject = null;
		do {
			l_TryCount += 1;
			#If DEBUG Then
			try {
				l_DeviceObject = CreateObject(p_ProgId);
				LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using CreateObject");
				#Else
				l_Type = Type.GetTypeFromProgID(p_ProgId);
				l_DeviceObject = Activator.CreateInstance(l_Type);
				LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using Activator.CreateInstance");
				#End If
				WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver initialisation");
				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using late binding");
				try {
					switch (p_DeviceType) {
						case DeviceType.Focuser:
							//Focuser uses link to connect
							l_DeviceObject.Link = true;
							l_DeviceObject.Link = false;
						default:
							//Everything else uses connect!
							l_DeviceObject.Connected = true;
							l_DeviceObject.Connected = false;
					}
					LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using late binding");

					try {
						if (l_DeviceObject.GetType.IsCOMObject) {
							LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver is a COM object");
						} else {
							LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver is a .NET object");
							LogMsg("AccessChecks", MessageLevel.msgInfo, "The AssemblyQualifiedName is: " + Microsoft.VisualBasic.Left(l_DeviceObject.GetType.AssemblyQualifiedName.ToString, 76));
						}
						foreach ( l_Type in l_DeviceObject.GetType.GetInterfaces()) {
							LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver implements interface: " + l_Type.FullName);
						}
					} catch (Exception ex) {
						LogMsg("AccessChecks", MessageLevel.msgError, "Error reading driver characteristsics: " + ex.Message);
						LogMsg("", MessageLevel.msgAlways, "");
					}
				} catch (Exception ex) {
					LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using late binding: " + ex.ToString);
					LogMsg("", MessageLevel.msgAlways, "");
				}
			} catch (Exception ex) {
				l_ErrMsg = ex.ToString;
				LogMsg("DeviceTesterBaseClass", MessageLevel.msgDebug, l_ErrMsg);
			}
			if (l_DeviceObject == null)
				WaitFor(200);
		} while (!((l_TryCount == 3) | (!(l_DeviceObject == null))));

		if (l_DeviceObject == null) {
			LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver object using late binding: " + l_ErrMsg);
			LogMsg("", MessageLevel.msgAlways, "");
		}

		//Clean up
		try {
			ReleaseCOMObjects(l_DeviceObject);
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Error releasing driver object using ReleaseCOMObject: " + ex.ToString);
		}
		l_DeviceObject = null;
		LogMsg("AccessChecks", MessageLevel.msgDebug, "Collecting garbage");
		GC.Collect();
		LogMsg("AccessChecks", MessageLevel.msgDebug, "Collecting garbage complete");
		GC.WaitForPendingFinalizers();
		LogMsg("AccessChecks", MessageLevel.msgDebug, "Finished waiting for pending finalisers");
		//Special handling for Gemini that takes a long time to dispose
		if ((p_ProgId.ToUpper == "ASCOM.GeminiTelescope.Telescope".ToUpper)) {
			WaitForAbsolute(GEMINI_DESTROY_WAIT, "Waiting for Telescope device driver to be cleaned up by operating system");
			//Wait to allow device to complete destruction
		} else {
			WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system");
			//Wait to allow device to complete destruction
		}
	}

	private virtual void CreateDevice()
	{
		LogMsg("CreateDevice", MessageLevel.msgError, "DeviceTester base Class warning messagev");
	}
	private virtual void PreConnectChecks()
	{
		LogMsg("PreConnectChecks", MessageLevel.msgError, "DeviceTester base Class warning messagev");
	}

	private virtual bool Connected {
		get { Connected = l_Connected; }
		set { l_Connected = value; }
	}

	private virtual void ReadCanProperties()
	{
		LogMsg("ReadCanProperties", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}
	private virtual void PreRunCheck()
	{
		LogMsg("PreSafetyCheck", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}

	private virtual void CheckProperties()
	{
		LogMsg("CheckProperties", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}
	private virtual void CheckMethods()
	{
		LogMsg("CheckMethods", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}
	private virtual void CheckPerformance()
	{
		LogMsg("CheckPerformance", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}

	private virtual void PostRunCheck()
	{
		LogMsg("PostSafetyCheck", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}

	private void SpecialTests(SpecialTests p_Test)
	{
		switch (p_Test) {
			case GlobalVarsAndCode.SpecialTests.TelescopeSideOfPier:
				SpecialTelescopeSideOfPier();
			case GlobalVarsAndCode.SpecialTests.TelescopeDestinationSideOfPier:
				SpecialTelescopeDestinationSideOfPier();
			case GlobalVarsAndCode.SpecialTests.TelescopeSideOfPierAnalysis:
				SpecialTelescopeSideOfPierAnalysis();
			case GlobalVarsAndCode.SpecialTests.TelescopeCommands:
				SpecialTelescopeCommands();
			default:
				MsgBox("Devicetester BaseClass:SpecialTests - Unknown test: " + p_Test.ToString, MsgBoxStyle.Critical);
		}
	}
	protected virtual void SpecialTelescopeSideOfPier()
	{
		LogMsg("SpecialTelescopeSideOfPier", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}
	protected virtual void SpecialTelescopeDestinationSideOfPier()
	{
		LogMsg("SpecialTelescopeDestinationSideOfPier", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}
	protected virtual void SpecialTelescopeSideOfPierAnalysis()
	{
		LogMsg("SpecialTelescopeFlipRange", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}
	protected virtual void SpecialTelescopeCommands()
	{
		LogMsg("SpecialTelescopeCommands", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
	}

	#End Region

	#Region "Support Code"

	private bool IncludeMethod(MandatoryMethod p_Method, DeviceType p_DeviceType, int p_InterfaceVersion)
	{
		//This mechanic allows individual tests for particular devices to be skipped. It is no longer required because this is handled by DriverAccess
		//The code is left in place in case it is ever needed in the future

		bool RetVal = true;
		//Default to true as most methods will be tested , we just list the exceptions to this below

		//Matrix controlling what tests
		switch (p_DeviceType) {
			case DeviceType.Telescope:
				switch (p_InterfaceVersion) {
					case 1:
						// Telescope interface V1 does not have Driver Version
						if (p_Method == MandatoryMethod.DriverVersion)
							RetVal = false;
					default:
						//All methods in all interface versions are mandatory
						RetVal = true;
				}
			case DeviceType.Camera:
				RetVal = true;
		}

		return RetVal;
	}

	private void CommandTest(CommandType p_Type, string p_Name)
	{
		string l_CmdSent = "!! Unknown Value !!";
		string m_CommandString;
		bool m_CommandBool;


		g_Status.Test = p_Name;
		try {
			switch (p_Type) {
				case CommandType.tstCommandBlind:
					if (!g_CmdStrings.CommandBlind == null) {
						l_CmdSent = g_CmdStrings.CommandBlind;
						#If PLATFORM6 Then
						Device.CommandBlind(l_CmdSent, false);
						#Else
						m_Telescope.CommandBlind(l_CmdSent);
						#End If

						LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStrings.CommandBlind + "\"");
					} else {
						LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBlind test");
					}
				case CommandType.tstCommandBool:
					if (!g_CmdStrings.CommandBool == null) {
						l_CmdSent = g_CmdStrings.CommandBool;
						#If PLATFORM6 Then

						m_CommandBool = Device.CommandBool(l_CmdSent, false);
						#Else
						m_CommandBool = Device.CommandBool(l_CmdSent);
						#End If
						if (m_CommandBool == g_CmdStrings.ReturnBool) {
							LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStrings.CommandBool + "\" - Received expected return value: " + m_CommandBool);
						} else {
							LogMsg(p_Name, MessageLevel.msgError, "Sent string \"" + g_CmdStrings.CommandBool + "\" - Received unexpected return value: " + m_CommandBool);
						}
					} else {
						LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBool test");
					}
				case CommandType.tstCommandString:
					if (!g_CmdStrings.CommandString == null) {
						l_CmdSent = g_CmdStrings.CommandString;
						#If PLATFORM6 Then
						m_CommandString = Device.CommandString(l_CmdSent, false);
						#Else
						m_CommandString = Device.CommandString(l_CmdSent);
						#End If
						//Valid return string toi test
						if (!g_CmdStrings.ReturnString == null) {
							if (m_CommandString == g_CmdStrings.ReturnString) {
								LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStrings.CommandString + "\" - Received string: \"" + m_CommandString + "\"");
							} else {
								LogMsg(p_Name, MessageLevel.msgError, "Sent string \"" + g_CmdStrings.CommandString + "\" - Received string: \"" + m_CommandString + "\" - Expected string: \"" + g_CmdStrings.ReturnString + "\"");
							}
						//Skip the return string test
						} else {
							LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStrings.CommandString + "\" - Return string test skipped");
						}
					} else {
						LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandString test");
					}
				case CommandType.tstCommandBlindRaw:
					if (!g_CmdStringsRaw.CommandBlind == null) {
						l_CmdSent = g_CmdStringsRaw.CommandBlind;
						Device.CommandBlind(l_CmdSent, true);
						LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStringsRaw.CommandBlind + "\"");
					} else {
						LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBlind Raw test");
					}
				case CommandType.tstCommandBoolRaw:
					if (!g_CmdStringsRaw.CommandBool == null) {
						l_CmdSent = g_CmdStringsRaw.CommandBool;
						m_CommandBool = Device.CommandBool(l_CmdSent, true);
						if (m_CommandBool == g_CmdStringsRaw.ReturnBool) {
							LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStringsRaw.CommandBool + "\" - Received expected return value: " + m_CommandBool);
						} else {
							LogMsg(p_Name, MessageLevel.msgError, "Sent string \"" + g_CmdStringsRaw.CommandBool + "\" - Received unexpected return value: " + m_CommandBool);
						}
					} else {
						LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBool Raw test");
					}
				case CommandType.tstCommandStringRaw:
					if (!g_CmdStringsRaw.CommandString == null) {
						l_CmdSent = g_CmdStringsRaw.CommandString;
						m_CommandString = Device.CommandString(l_CmdSent, true);
						//Valid return string to test
						if (!g_CmdStringsRaw.ReturnString == null) {
							if (m_CommandString == g_CmdStringsRaw.ReturnString) {
								LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Received string: \"" + m_CommandString + "\"");
							} else {
								LogMsg(p_Name, MessageLevel.msgError, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Received string: \"" + m_CommandString + "\" - Expected string: \"" + g_CmdStringsRaw.ReturnString + "\"");
							}
						//Skip the return string test
						} else {
							LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Return string test skipped");
						}
					} else {
						LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandString Raw test");

					}
				default:
					LogMsg(p_Name, MessageLevel.msgError, "Conform:CommandTest: Unknown test type " + p_Type.ToString);
			}
		} catch (InvalidValueException ex) {
			LogMsg(p_Name, MessageLevel.msgError, ".NET Invalid value sent to telescope: " + l_CmdSent);
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg(p_Name, MessageLevel.msgError, "COM Invalid value sent to telescope: " + l_CmdSent + "\"");
				default:
					ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (MethodNotImplementedException ex) {
			LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
		} catch (ASCOM.NotImplementedException ex) {
			LogMsg(p_Name, MessageLevel.msgIssue, "ASCOM.NotImplementedException thrown, expecting ASCOM.MethodNotImplementedException");
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.ToString);
		}
		g_Status.Clear();
		//Clear status messages
	}

	#End Region
	public static void ReleaseCOMObjects(ref object ObjectToRelease)
	{
		Type ObjectType;
		int RemainingObjectCount;
		int LoopCount;

		LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Start of ReleaseCOMObject");

		try {
			ObjectType = ObjectToRelease.GetType;
			LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Unmarshalling " + ObjectType.Name);
		} catch (Exception ex1) {
			LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  GetType Exception: " + ex1.Message);
		}

		try {
			LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Releasing COM object");
			LoopCount = 0;
			do {
				LoopCount += 1;
				RemainingObjectCount = Marshal.ReleaseComObject(ObjectToRelease);
				LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Remaining object count: " + RemainingObjectCount + ", LoopCount: " + LoopCount);
			} while (!((RemainingObjectCount <= 0) | (LoopCount == 20)));
		} catch (Exception ex2) {
			LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  ReleaseComObject Exception: " + ex2.Message);
		}


		try {
			ObjectToRelease = null;
		} catch (Exception ex3) {
			LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  Set to nothing Exception: " + ex3.Message);
		}
		LogMsg("ReleaseCOMObjects", MessageLevel.msgDebug, "  End of ReleaseCOMObject");
	}
}

