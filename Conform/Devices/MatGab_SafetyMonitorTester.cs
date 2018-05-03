using ASCOM.Interface;
internal class SafetyMonitorTester : DeviceTesterBaseClass
{

	#Region "Variables and Constants"
	private bool m_CanIsGood;
	private bool m_CanEmergencyShutdown;
	private bool m_IsSafe;
	private bool m_IsGood;
	private bool m_EmergencyShutdown;
	private bool m_Connected;
	private string m_Description;
	private string m_DriverInfo;

	private string m_DriverVersion;
	#If DEBUG Then
		#Else
	private ISafetyMonitor m_SafetyMonitor;
		#End If
	private object m_SafetyMonitor;
	#End Region

	#Region "Enums"
	private enum CanType
	{
		tstCanIsGood = 1,
		tstCanEmergencyShutdown = 2
	}
	private enum RequiredProperty
	{
		propIsSafe,
		propDescription,
		propDriverInfo,
		propDriverVersion,
		propConnected
	}
	private enum OptionalProperty
	{
		propIsGood,
		propEmergencyShutdown
	}
	private enum PerformanceProperty
	{
		propIsSafe,
		propIsGood,
		propEmergencyShutdown
	}
	#End Region

	#Region "New and Dispose"
	private SafetyMonitorTester()
	{
		this.HasMethods = false;
		this.HasCanProperties = true;
		this.HasPreConnectCheck = true;
		this.HasPerformanceCheck = true;
	}
	// IDisposable
		// To detect redundant calls
	private bool disposedValue = false;
	protected override void Dispose(bool disposing)
	{
		if (!this.disposedValue) {
			if (disposing) {
				// TODO: free other state (managed objects).
			}
			//Should be True but make False to stop Conform from cleanly dropping the telescope object (useful for retaining scopesim in memory to change flags
			if (true) {
				try {
					m_SafetyMonitor.Connected = false;
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_SafetyMonitor);
				} catch {
				}
				m_SafetyMonitor = null;
				GC.Collect();
			}

			// TODO: free your own state (unmanaged objects).
			// TODO: set large fields to null.
		}
		this.disposedValue = true;
	}
	#End Region

	#Region "Code"
	private override void CheckInitialise()
	{
		//Set the error type numbers acording to the standards adopted by individual authors.
		//Unfortunatley these vary between drivers so I have to allow for these here in order to give meaningful
		//messages to driver authors!
		switch (g_SafetyMonitorProgID) {
			default:
				//I'm using the simulator values as the defaults since it is the reference platform
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040405;
				g_ExInvalidValue2 = 0x80040405;
				g_ExInvalidValue3 = 0x80040405;
				g_ExInvalidValue4 = 0x80040405;
				g_ExInvalidValue5 = 0x80040405;
				g_ExInvalidValue6 = 0x80040405;
				g_ExNotSet = 0x80040403;
		}
		base.CheckInitialise(g_SafetyMonitorProgID);

	}
	private override void CheckAccessibility()
	{
		ISafetyMonitor l_ISafetyMonitor;
		ASCOM.DriverAccess.SafetyMonitor l_DriverAccessSafetyMonitor = null;
		object l_DeviceObject = null;
		if (false) {
			LogMsg("ConformanceCheckAccessibility", MessageLevel.msgInfo, "Early binding and Driver Access Toolkit checks omitted");
		} else {
			base.CheckAccessibility(g_SafetyMonitorProgID, DeviceType.SafetyMonitor);
			//Try early binding
			l_ISafetyMonitor = null;
			try {
				l_DeviceObject = CreateObject(g_SafetyMonitorProgID);
				l_ISafetyMonitor = (ISafetyMonitor)l_DeviceObject;

				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using early binding to ISafetyMonitor interface");
				try {
					l_ISafetyMonitor.Connected = true;
					LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using early binding to ISafetyMonitor interface");
					l_ISafetyMonitor.Connected = false;
				} catch (Exception ex) {
					LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using early binding to ISafetyMonitor interface: " + ex.Message);
					LogMsg("", MessageLevel.msgAlways, "");
				}
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using early binding to ISafetyMonitor: " + ex.Message);
				LogMsg("", MessageLevel.msgAlways, "");
			} finally {
				//Clean up
				try {
					Marshal.ReleaseComObject(l_ISafetyMonitor);
				} catch {
				}
				try {
					Marshal.ReleaseComObject(l_DeviceObject);
				} catch {
				}
				l_DeviceObject = null;
				l_ISafetyMonitor = null;
				GC.Collect();
			}

			//Try client access toolkit
			try {
				l_DriverAccessSafetyMonitor = new ASCOM.DriverAccess.SafetyMonitor(g_SafetyMonitorProgID);
				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
				try {
					l_DriverAccessSafetyMonitor.Connected = true;
					LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
					l_DriverAccessSafetyMonitor.Connected = false;
				} catch (Exception ex) {
					LogMsg("AccessChecks", MessageLevel.msgError, "Error conecting to driver using driver access toolkit: " + ex.Message);
					LogMsg("", MessageLevel.msgAlways, "");
				}
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " + ex.Message);
				LogMsg("", MessageLevel.msgAlways, "");
			} finally {
				//Clean up
				try {
					l_DriverAccessSafetyMonitor.Dispose();
				} catch {
				}
				try {
					Marshal.ReleaseComObject(l_DriverAccessSafetyMonitor);
				} catch {
				}
				l_DriverAccessSafetyMonitor = null;
				GC.Collect();
			}
		}
	}

	#If DEBUG Then
	private override void CreateDevice()
	{
		g_DeviceObject = CreateObject(g_SafetyMonitorProgID);
		m_SafetyMonitor = (ISafetyMonitor)g_DeviceObject;
		LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.ISafetyMmonitor to get a SafetyMonitor object");
		#Else
		m_SafetyMonitor = CreateObject(g_SafetyMonitorProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a SafetyMonitor object");
		#End If
		g_Stop = false;
		//connected OK so clear stop flag to allow other tests to run
	}
	public override void PreConnectChecks()
	{
		//Confirm that key properties are false when not connected
		try {
			m_IsSafe = m_SafetyMonitor.IsSafe;
			if (!m_IsSafe) {
				LogMsg("IsSafe", MessageLevel.msgOK, "Reports false before connection");
			} else {
				LogMsg("IsSafe", MessageLevel.msgIssue, "Reports true before connection rather than false");
			}
		} catch (Exception ex) {
			LogMsg("IsSafe", MessageLevel.msgError, "Cannot confirm that IsSafe is false before connection because it threw an exception: " + ex.Message);
		}
		try {
			m_IsGood = m_SafetyMonitor.IsGood;
			if (!m_IsGood) {
				LogMsg("IsGood", MessageLevel.msgOK, "Reports false before connection");
			} else {
				LogMsg("IsGood", MessageLevel.msgIssue, "Reports true before connection rather than false");
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("IsGood", MessageLevel.msgInfo, "Received PropertyNotImplemented exception");
		} catch (Exception ex) {
			LogMsg("IsGood", MessageLevel.msgError, "Cannot confirm that IsGood is false before connection because it threw an exception: " + ex.Message);
		}
	}
	private override bool Connected {
		get { Connected = m_SafetyMonitor.Connected; }
		set { m_SafetyMonitor.Connected = value; }
	}

	public override void ReadCanProperties()
	{
		SafetyMonitorCanTest(CanType.tstCanEmergencyShutdown, "CanEmergencyShutdown");
		SafetyMonitorCanTest(CanType.tstCanIsGood, "CanIsGood");
	}
	private override void CheckProperties()
	{
		RequiredPropertiesTest(RequiredProperty.propDescription, "Description");
		RequiredPropertiesTest(RequiredProperty.propDriverInfo, "DriverInfo");
		RequiredPropertiesTest(RequiredProperty.propDriverVersion, "DriverVersion");
		RequiredPropertiesTest(RequiredProperty.propConnected, "Connected");
		OptionalPropertiesTest(OptionalProperty.propEmergencyShutdown, "EmergencyShutdown", m_CanEmergencyShutdown, "CanEmergencyShutdown");
		RequiredPropertiesTest(RequiredProperty.propIsSafe, "IsSafe");
		OptionalPropertiesTest(OptionalProperty.propIsGood, "IsGood", m_CanIsGood, "CanIsGood");
	}
	private override void CheckPerformance()
	{
		Status(StatusType.staTest, "Performance");
		PerformanceTest(PerformanceProperty.propIsSafe, "IsSafe");
		if (m_CanIsGood)
			PerformanceTest(PerformanceProperty.propIsGood, "IsGood");
		if (m_CanEmergencyShutdown)
			PerformanceTest(PerformanceProperty.propEmergencyShutdown, "EmergencyShutdown");
		Status(StatusType.staTest, "");
		Status(StatusType.staAction, "");
		Status(StatusType.staStatus, "");
	}

	private void SafetyMonitorCanTest(CanType p_Type, string p_Name)
	{
		try {
			switch (p_Type) {
				case CanType.tstCanIsGood:
					m_CanIsGood = m_SafetyMonitor.CanIsGood;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanIsGood.ToString);
				case CanType.tstCanEmergencyShutdown:
					m_CanEmergencyShutdown = m_SafetyMonitor.CanEmergencyShutdown;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanEmergencyShutdown.ToString);
				default:
					LogMsg(p_Name, MessageLevel.msgError, "Conform:CanTest: Unknown test type " + p_Type.ToString);
			}
		} catch (COMException ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.Message);
		}
	}
	private void RequiredPropertiesTest(RequiredProperty p_Type, string p_Name)
	{
		try {
			switch (p_Type) {
				case RequiredProperty.propDescription:
					m_Description = m_SafetyMonitor.Description;
					LogMsg(p_Name, MessageLevel.msgOK, m_Description);
				case RequiredProperty.propDriverInfo:
					m_DriverInfo = m_SafetyMonitor.DriverInfo;
					LogMsg(p_Name, MessageLevel.msgOK, m_DriverInfo);
				case RequiredProperty.propDriverVersion:
					m_DriverVersion = m_SafetyMonitor.DriverVersion;
					LogMsg(p_Name, MessageLevel.msgOK, m_DriverVersion);
				case RequiredProperty.propIsSafe:
					m_IsSafe = m_SafetyMonitor.IsSafe;
					//Check that value is false
					if (!m_Connected) {
						if (!m_IsSafe) {
							LogMsg(p_Name, MessageLevel.msgOK, m_IsSafe.ToString);
						} else {
							LogMsg(p_Name, MessageLevel.msgIssue, "IsSafe is True when not connected, IsSafe should be False");
						}
					// We are connected so test whether emergency shutdown is true, if so IsSafes hsoiuld be false
					} else {
						if (m_CanEmergencyShutdown) {
							//EmergencyShutdown is true so IsSafe shoudl be false
							if (m_EmergencyShutdown) {
								if (!m_IsSafe) {
									LogMsg(p_Name, MessageLevel.msgOK, m_IsSafe.ToString);
								} else {
									LogMsg(p_Name, MessageLevel.msgIssue, "IsSafe is True when EmergencyShutdown is true, IsSafe should be False");
								}
							//Any value is OK
							} else {
								LogMsg(p_Name, MessageLevel.msgOK, m_IsSafe.ToString);
							}
						}
					}
				case RequiredProperty.propConnected:
					m_Connected = m_SafetyMonitor.Connected;
					LogMsg(p_Name, MessageLevel.msgOK, m_Connected.ToString);
				default:
					LogMsg(p_Name, MessageLevel.msgError, "RequiredPropertiesTest: Unknown test type " + p_Type.ToString);
			}
		} catch (COMException ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.Message);
		}
	}
	private void OptionalPropertiesTest(OptionalProperty p_Type, string p_Name, bool p_CanTest, string p_CanName)
	{
		Status(StatusType.staTest, p_Name);
		// Confirm that an error is raised if the optional command is not implemented
		if (p_CanTest) {
			try {
				switch (p_Type) {
					case OptionalProperty.propEmergencyShutdown:
						m_EmergencyShutdown = m_SafetyMonitor.EmergencyShutdown;
						LogMsg(p_Name, MessageLevel.msgOK, m_EmergencyShutdown.ToString);
					case OptionalProperty.propIsGood:
						m_IsGood = m_SafetyMonitor.IsGood;
						LogMsg(p_Name, MessageLevel.msgOK, m_IsGood.ToString);
					default:
						LogMsg(p_Name, MessageLevel.msgError, "OptionalPropertiesTest: Unknown test type " + p_Type.ToString);
				}
				if (TestStop())
					return;
			} catch (COMException ex) {
				LogMsg(p_Name, MessageLevel.msgError, p_CanName + " is true but received unexpected COM exception: " + ex.Message);
			} catch (DriverException ex) {
				LogMsg(p_Name, MessageLevel.msgError, p_CanName + " is true but received unexpected .NET exception: " + ex.Message);
			} catch (Exception ex) {
				LogMsg(p_Name, MessageLevel.msgError, p_CanName + " is true but received unexpected exception: " + ex.Message);
			}
		//Can property is false so confirm that an error is generated
		} else {
			try {
				switch (p_Type) {
					case OptionalProperty.propEmergencyShutdown:
						m_EmergencyShutdown = m_SafetyMonitor.EmergencyShutdown;
					case OptionalProperty.propIsGood:
						m_IsGood = m_SafetyMonitor.IsGood;
					default:
						LogMsg(p_Name, MessageLevel.msgError, "Conform:OptionalMethodsTest: Unknown test type " + p_Type.ToString);
				}
				LogMsg(p_Name, MessageLevel.msgIssue, p_CanName + " is false but no exception was generated on use");
			} catch (COMException ex) {
				switch (ex.ErrorCode) {
					case g_ExNotImplemented:
					case ErrorCodes.NotImplemented:
						LogMsg(p_Name, MessageLevel.msgOK, p_CanName + " is false and COM Not Implemented exception generated as expected");
					default:
						ExTest(p_Name, ex.Message, p_CanName + " is false but inappropriate COM exception generated: " + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
			} catch (DriverException ex) {
				switch (ex.Number) {
					case g_ExNotImplemented:
					case ErrorCodes.NotImplemented:
						LogMsg(p_Name, MessageLevel.msgOK, p_CanName + " is false and .NET Not Implemented exception generated as expected");
					case ErrorCodes.InvalidValue:
					case g_ExInvalidValue1:
					case g_ExInvalidValue2:
						LogMsg(p_Name, MessageLevel.msgOK, p_CanName + " is false and .NET Invalid Value exception generated");
					default:
						ExTest(p_Name, ex.Message, p_CanName + " is false but inappropriate .NET exception generated: " + ex.Message + " " + Hex(ex.Number));
				}
			} catch (Exception ex) {
				ExTest(p_Name, ex.Message, p_CanName + " is false but inappropriate exception generated: " + ex.Message);
			}
		}
	}
	private void PerformanceTest(PerformanceProperty p_Type, string p_Name)
	{
		System.DateTime l_StartTime;
		double l_Count;
		double l_LastElapsedTime;
		double l_ElapsedTime;
		double l_Rate;
		Status(StatusType.staAction, p_Name);
		try {
			l_StartTime = Now;
			l_Count = 0.0;
			l_LastElapsedTime = 0.0;
			do {
				l_Count += 1.0;
				switch (p_Type) {
					case PerformanceProperty.propEmergencyShutdown:
						m_EmergencyShutdown = m_SafetyMonitor.EmergencyShutdown;
					case PerformanceProperty.propIsGood:
						m_IsGood = m_SafetyMonitor.IsGood;
					case PerformanceProperty.propIsSafe:
						m_IsSafe = m_SafetyMonitor.IsSafe;
					default:
						LogMsg(p_Name, MessageLevel.msgError, "PerformanceTest: Unknown test type " + p_Type.ToString);
				}
				l_ElapsedTime = Now.Subtract(l_StartTime).TotalSeconds;
				if (l_ElapsedTime > l_LastElapsedTime + 1.0) {
					Status(StatusType.staStatus, l_Count + " transactions in " + Format(l_ElapsedTime, "0") + " seconds");
					l_LastElapsedTime = l_ElapsedTime;
					Application.DoEvents();
					if (TestStop())
						return;
				}
			} while (!(l_ElapsedTime > PERF_LOOP_TIME));
			l_Rate = l_Count / l_ElapsedTime;
			switch (l_Rate) {
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
10.0:
					LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + Format(l_Rate, "0.0") + " per second");
					break; // TODO: might not be correct. Was : Exit Select

				case 2.0: // TODO: to 10.0
					LogMsg(p_Name, MessageLevel.msgOK, "Transaction rate: " + Format(l_Rate, "0.0") + " per second");
					break; // TODO: might not be correct. Was : Exit Select

				case 1.0: // TODO: to 2.0
					LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + Format(l_Rate, "0.0") + " per second");
					break; // TODO: might not be correct. Was : Exit Select

				default:
					LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + Format(l_Rate, "0.0") + " per second");
			}
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgInfo, "Unable to complete test: " + ex.ToString);
		}
	}

	#End Region
}
