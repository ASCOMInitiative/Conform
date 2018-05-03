internal class SwitchTester : DeviceTesterBaseClass
{

	#Region "Variables and Constants"
	private const int SWITCH_NUMBER_TEST_RANGE = 100;
	private int m_InterfaceVersion;
	private int m_MaxSwitch;
	private int m_PerformanceGetSwitchName;
	private int m_PerformanceGetSwitch;
	private bool m_CanReadMaxSwitch;
	private string m_SwitchType;

	private ArrayList m_Switches;
	#If DEBUG Then
	//Private m_Switch As ASCOM.DeviceInterface.ISwitchV2
		#Else
	private object m_Switch;
		#End If
	private object m_Switch;
	#End Region

	#Region "Enums"
	private enum SwitchPropertyMethod
	{
		Description,
		DriverInfo,
		DriverVersion,
		Name,
		InterfaceVersion,
		MaxSwitch,
		GetSwitch,
		GetSwitchName,
		SwitchType
	}
	#End Region

	#Region "New and Dispose"
	private SwitchTester()
	{
		base.New();
		m_PerformanceGetSwitchName = int.MinValue;
		//Initialise to silly values
		m_PerformanceGetSwitch = int.MinValue;
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
					m_Switch.Connected = false;
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_Switch);
				} catch {
				}
				m_Switch = null;
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

		switch (g_SwitchProgID) {
			default:
				//I'm using the simulator values as the defaults since it is the reference platform
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040405;
				g_ExInvalidValue2 = 0x80040405;
				g_ExNotSet = 0x80040403;
		}
		base.CheckInitialise(g_SwitchProgID);

	}
	private override void CheckAccessibility()
	{
		object l_ISwitch;
		DriverAccess.Switch l_DriverAccessSwitch;
		object l_DeviceObject = null;

		base.CheckAccessibility(g_SwitchProgID, DeviceType.Switch);
		//Try early binding
		l_ISwitch = null;
		#If DEBUG Then
		try {
			l_DeviceObject = CreateObject(g_SwitchProgID);
			#Else
			Type l_Type;
			l_Type = Type.GetTypeFromProgID(g_SwitchProgID);
			l_DeviceObject = Activator.CreateInstance(l_Type);
			#End If
			l_ISwitch = (ASCOM.Interface.ISwitch)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through interface ISwitch");
			try {
				l_ISwitch.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ISwitch");
				l_ISwitch.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgDebug, "Connected Exception: " + ex.ToString);
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ISwitch");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgDebug, "CreateInstance Exception: " + ex.ToString);
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ISwitch");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_ISwitch);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_ISwitch = null;
			GC.Collect();
		}

		try {
			l_DeviceObject = CreateObject(g_SwitchProgID);
			l_ISwitch = (ASCOM.DeviceInterface.ISwitchV2)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through interface ISwitchV2");
			try {
				l_ISwitch.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ISwitchV2");
				l_ISwitch.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ISwitchV2");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ISwitchV2");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_ISwitch);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_ISwitch = null;
			GC.Collect();
		}

		WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system");

		//Try client access toolkit
		l_DriverAccessSwitch = null;
		try {
			l_DriverAccessSwitch = new ASCOM.DriverAccess.Switch(g_SwitchProgID);
			LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
			try {
				l_DriverAccessSwitch.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
				l_DriverAccessSwitch.Connected = false;
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
				l_DriverAccessSwitch.Dispose();
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DriverAccessSwitch);
			} catch {
			}
			l_DriverAccessSwitch = null;
			GC.Collect();
		}
	}
	#If DEBUG Then
	private override void CreateDevice()
	{
		m_Switch = new ASCOM.DriverAccess.Switch(g_SwitchProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Switch to get a Switch object");
		#Else
		WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for switch to initialise");
		Type l_Type;
		l_Type = Type.GetTypeFromProgID(g_SwitchProgID);
		m_Switch = Activator.CreateInstance(l_Type);
		//m_Switch = CreateObject(g_SwitchProgID)
		LogMsg("Conform", MessageLevel.msgAlways, "is using Activator.CreateInstance to get a Switch object");
		#End If
		g_Stop = false;
		//connected OK so clear stop flag to allow other tests to run
	}
	private override bool Connected {
		get { Connected = base.Connected; }
		set {
			m_Switch.Connected = value;
			base.Connected = value;
		}
	}
	private override void CheckCommonMethods()
	{
		base.CheckCommonMethods(m_Switch, DeviceType.Switch);
	}

	private override void CheckProperties()
	{
		string l_Name;
		string l_State;
		switch (g_InterfaceVersion) {
			case 1:
				//Original Platform 5 switch interface 
				m_MaxSwitch = SwitchPropertyTestShort(SwitchPropertyMethod.MaxSwitch, "MaxSwitch", 1, short.MaxValue);
				if (TestStop())
					return;

			case 2:
				m_SwitchType = SwitchPropertyTestString(SwitchPropertyMethod.SwitchType, "SwitchType", int.MaxValue);
				if (TestStop())
					return;

				try {
					m_Switches = m_Switch.Switches;
					LogMsg("Switches", MessageLevel.msgInfo, "Retrieved " + m_Switches.Count + " switches");
					foreach (object l_Switch in m_Switches) {
						l_Name = l_Switch.Name;
						if (string.IsNullOrEmpty(l_Name)) {
							LogMsg("Switches", MessageLevel.msgError, "Invalid name value: name is null or empty string");
							l_Name = "";
						//Name is valid value
						} else {
						}
						switch (m_SwitchType.ToUpper) {
							case "TOGGLESWITCH":
								l_State = l_Switch.State(0);
								if (string.IsNullOrEmpty(l_State)) {
									LogMsg("Switches", MessageLevel.msgError, l_Name + " - Invalid state value: state is null or empty string");
									l_State = "";
								}
								switch (l_State.ToUpper) {
									case "ON":
									case "OFF":
										//Valid values so no error message
										LogMsg("Switches", MessageLevel.msgOK, l_Switch.Name + " = " + l_State);
									default:
										LogMsg("Switches", MessageLevel.msgError, l_Switch.Name + " - Invalid state value: " + l_State);
								}
							case "NWAYSWITCH":
								LogMsg("Switches", MessageLevel.msgOK, l_Switch.Name);
								foreach ( l_State in l_Switch.State) {
									if (IsNumeric(l_State)) {
										LogMsg("Switches", MessageLevel.msgOK, "  " + (int)l_State);
									} else {
										LogMsg("Switches", MessageLevel.msgError, "  Invalid state value: " + l_State);
									}
								}

							default:
								//Can't test switches we don't know about so just report values returned
								LogMsg("Switches", MessageLevel.msgOK, l_Switch.Name);
								foreach ( l_State in l_Switch.State) {
									LogMsg("Switches", MessageLevel.msgOK, "  " + l_State);
								}

						}
					}
				} catch (COMException ex) {
					LogMsg("Switches", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("Switches", MessageLevel.msgError, EX_NET + ex.Message);
				}
			default:
				LogMsg("Switches", MessageLevel.msgError, "Unknown switch interface version: " + m_InterfaceVersion);
		}
	}
	private string SwitchPropertyTestString(SwitchPropertyMethod p_Type, string p_Name, int p_MaxLength)
	{
		SwitchPropertyTestString = "";
		try {
			switch (p_Type) {
				case SwitchPropertyMethod.Description:
					SwitchPropertyTestString = m_Switch.Description;
				case SwitchPropertyMethod.DriverInfo:
					SwitchPropertyTestString = m_Switch.DriverInfo;
				case SwitchPropertyMethod.DriverVersion:
					SwitchPropertyTestString = m_Switch.DriverVersion;
				case SwitchPropertyMethod.Name:
					SwitchPropertyTestString = m_Switch.Name;
				case SwitchPropertyMethod.SwitchType:
					SwitchPropertyTestString = m_Switch.SwitchType;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "SwitchPropertyTestString: Unknown test type - " + p_Type.ToString);
			}
			//Successfully retrieved a value
			switch (SwitchPropertyTestString) {
				case  // ERROR: Case labels with binary operators are unsupported : Equality
"":
					LogMsg(p_Name, MessageLevel.msgError, "Null string returned");
				default:
					if (Len(SwitchPropertyTestString) <= p_MaxLength) {
						LogMsg(p_Name, MessageLevel.msgOK, SwitchPropertyTestString);
					} else {
						LogMsg(p_Name, MessageLevel.msgError, "String exceeds " + p_MaxLength + " characters maximum length - " + SwitchPropertyTestString);
					}
			}
		} catch (COMException ex) {
			//If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
			//LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM)
			//Else
			//ExTest(p_Name, ex.Message, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
			LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		//End If
		//Catch ex As PropertyNotImplementedException
		//    LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET)
		} catch (Exception ex) {
			//ExTest(p_Name, ex.Message, EX_NET & ex.Message)
			LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.Message);
		}

	}
	private short SwitchPropertyTestShort(SwitchPropertyMethod p_Type, string p_Name, short p_Min, short p_Max)
	{
		try {
			SwitchPropertyTestShort = 0;
			switch (p_Type) {
				case SwitchPropertyMethod.InterfaceVersion:
					SwitchPropertyTestShort = m_Switch.InterfaceVersion;
				case SwitchPropertyMethod.MaxSwitch:
					m_CanReadMaxSwitch = false;
					SwitchPropertyTestShort = m_Switch.MaxSwitch;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "SwitchPropertyTestInteger: Unknown test type - " + p_Type.ToString);
			}
			//Successfully retrieved a value
			switch (SwitchPropertyTestShort) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
p_Min:
					//Lower than minimum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + SwitchPropertyTestShort.ToString);
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
p_Max:
					//Higher than maximum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + SwitchPropertyTestShort.ToString);
				default:
					//OK value
					m_CanReadMaxSwitch = true;
					//A valid value has been found
					LogMsg(p_Name, MessageLevel.msgOK, SwitchPropertyTestShort.ToString);
			}
		} catch (COMException ex) {
			//If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
			// LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM)
			// Else
			// ExTest(p_Name, ex.Message, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
			// End If
			LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		//Catch ex As PropertyNotImplementedException
		//    LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET)
		} catch (Exception ex) {
			//ExTest(p_Name, ex.Message, EX_NET & ex.Message)
			LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.Message);
		}
	}

	private override void CheckMethods()
	{
		short i;
		bool l_SwitchState;
		bool l_NewSwitchState;
		bool l_FoundSwitchRead;
		bool l_FoundSwitchWrite;
		string l_SwitchName;
		object Switch;
		string l_Name;
		string l_ExpectedState;
		switch (g_InterfaceVersion) {
			case 1:
				//Platform 5 interface v1
				//Have got a valid value
				if (m_CanReadMaxSwitch) {
					Status(StatusType.staTest, "Read/Write Switches");
					Status(StatusType.staAction, "Testing switch");
					//Find valid GetSwitch values
					for (i = -SWITCH_NUMBER_TEST_RANGE; i <= (short)m_MaxSwitch + SWITCH_NUMBER_TEST_RANGE; i++) {
						Status(StatusType.staStatus, i.ToString);
						l_FoundSwitchRead = false;
						l_FoundSwitchWrite = false;
						//Read switch state
						try {
							l_SwitchState = m_Switch.GetSwitch(i);
							LogMsg("GetSwitch " + i, MessageLevel.msgOK, "Found switch, state: " + l_SwitchState.ToString);
							l_FoundSwitchRead = true;
							if (i > m_MaxSwitch) {
								LogMsg("GetSwitch " + i, MessageLevel.msgError, "Usable switch found above MaxSwitch!");
							} else {
								m_PerformanceGetSwitch = i;
								//Save last good switch number for performance test
							}
						} catch (Exception ex) {
							LogMsg("GetSwitch " + i, MessageLevel.msgDebug, "Exception: " + ex.Message);
						}
						//Now try to write the value
						try {
							m_Switch.SetSwitch(i, !l_SwitchState);
							// Swap the switch state
							l_NewSwitchState = m_Switch.GetSwitch(i);
							//Read the new switch state to confirm that value did change
							//State did change, test successful
							if (l_NewSwitchState == !l_SwitchState) {
								LogMsg("SetSwitch " + i, MessageLevel.msgOK, "Switch correctly changed state");
								m_Switch.SetSwitch(i, l_SwitchState);
								//Now put switch back to original state
							//State didn't change so error
							} else {
								LogMsg("SetSwitch " + i, MessageLevel.msgError, "Switch did not change state, currently it is " + l_NewSwitchState.ToString);
							}
							if (!l_FoundSwitchRead) {
								LogMsg("SetSwitch " + i, MessageLevel.msgInfo, "You have a write only switch!");
							}
							if (i > m_MaxSwitch) {
								LogMsg("SetSwitch " + i, MessageLevel.msgError, "Usable switch found above MaxSwitch!");
							}
							l_FoundSwitchWrite = true;
						} catch (Exception ex) {
							if (l_FoundSwitchRead) {
								LogMsg("SetSwitch " + i, MessageLevel.msgInfo, "Unable to write to switch: " + ex.Message);
							} else {
								LogMsg("SetSwitch " + i, MessageLevel.msgDebug, "Exception: " + ex.Message);
							}
						}

						try {
							l_SwitchName = m_Switch.GetSwitchName(i);
							//Name is associated with a real switch
							if (l_FoundSwitchRead | l_FoundSwitchWrite) {
								//Switch name is a null string
								if (l_SwitchName == "") {
									LogMsg("SwitchName " + i, MessageLevel.msgInfo, "Switch name returns null string");
								//Proper switch name returned
								} else {
									m_PerformanceGetSwitchName = i;
									//Save last good name index for performance test
									LogMsg("SwitchName " + i, MessageLevel.msgOK, "Found switch, name:  " + l_SwitchName);
								}
							//Switch is named but has neither read nor write capability
							} else {
								//Switch name is a null string
								if (l_SwitchName == "") {
									LogMsg("SwitchName " + i, MessageLevel.msgError, "Switch name returns null string but no switches found!");
								//Proper switch name returned
								} else {
									LogMsg("SwitchName " + i, MessageLevel.msgError, "Found switch, name:  " + l_SwitchName + " but no switches found!");
								}
							}
						} catch (Exception ex) {
							LogMsg("SwitchName " + i, MessageLevel.msgDebug, "Exception: " + ex.Message);
						}

					}
					Status(StatusType.staTest, "");
					Status(StatusType.staAction, "");
					Status(StatusType.staStatus, "");

				//No valid value
				} else {
					LogMsg("SwitchCheckMethods", MessageLevel.msgInfo, "Skipping further tests as thereis no valid value for MaxSwitch");
				}
			case 2:
				//Platform 6 interface version 2
				try {
					m_Switches = m_Switch.Switches;
					foreach (object l_Switch in m_Switches) {
						l_Name = l_Switch.Name;
						if (string.IsNullOrEmpty(l_Name)) {
							LogMsg("Get Switch", MessageLevel.msgError, "Invalid name value: name is null or empty string");
							l_Name = "";
						//Name is valid value
						} else {
							LogMsg("Get Switch", MessageLevel.msgInfo, "Switch " + l_Name);
							Switch = m_Switch.GetSwitch(l_Name);
							switch (m_SwitchType.ToUpper) {
								case "TOGGLESWITCH":
									if (Switch.State.Length != 1)
										LogMsg("State", MessageLevel.msgError, "Only expected one state value, actually received " + l_Switch.State.Length);
									foreach (string l_State in Switch.State) {
										if (string.IsNullOrEmpty(l_State)) {
											LogMsg("Get Switch", MessageLevel.msgError, l_Name + " - Invalid state value: state is null or empty string");
											l_State = "";
										//Name and state are potentially valid so lets test the switch!
										} else {
											switch (l_State.ToUpper) {
												case "ON":
												case "OFF":
													//Valid values so no error message
													LogMsg("Get Switch", MessageLevel.msgOK, "  " + Switch.Name + " = " + Switch.State(0));
													if (g_Settings.SwitchSet) {
														//Turn it on
														if (l_State.ToUpper == "OFF") {
															l_ExpectedState = "ON";
															m_Switch.SetSwitch(l_Name, new string[] { "On" });
														//Turn it off
														} else {
															l_ExpectedState = "OFF";
															m_Switch.SetSwitch(l_Name, new string[] { "Off" });
														}
														Switch = m_Switch.GetSwitch(l_Name);
														//OK!
														if (Switch.State(0).ToUpper == l_ExpectedState) {
															LogMsg("Set Switch", MessageLevel.msgOK, "  " + Switch.Name + " = " + Switch.State(0));
															m_Switch.SetSwitch(l_Name, new string[] { l_State });
															//Reset original value
														//Bad value
														} else {
															LogMsg("Set Switch", MessageLevel.msgError, "  Unexpected switch state for " + Switch.Name + " - " + Switch.State(0));
														}
													} else {
														LogMsg("Set Switch", MessageLevel.msgInfo, "  " + Switch.Name + " Set test omitted because of Conform safety setting");
													}
												default:
													LogMsg("Get Switch", MessageLevel.msgError, l_Switch.Name + "- Invalid state value: " + l_State);
											}
										}
									}

								case "NWAYSWITCH":
								case "RHEOSTAT":
									int l_Min;
									int l_Max;
									int l_Current;
									LogMsg("Get Switch", MessageLevel.msgOK, l_Switch.Name);
									if (IsNumeric(Switch.State(0))) {
										l_Min = Switch.State(0);
										LogMsg("Get Switch", MessageLevel.msgOK, "  Minimum: " + (int)l_Min);
									} else {
										LogMsg("Get Switch", MessageLevel.msgError, "  Minimum value (State(0)) is not numeric: " + Switch.State(0));
									}
									if (IsNumeric(Switch.State(1))) {
										l_Max = Switch.State(1);
										LogMsg("Get Switch", MessageLevel.msgOK, "  Maximum: " + (int)l_Max);
									} else {
										LogMsg("Get Switch", MessageLevel.msgError, "  Maximum value (State(1)) is not numeric: " + Switch.State(1));
									}
									if (IsNumeric(Switch.State(2))) {
										l_Current = Switch.State(2);
										LogMsg("Get Switch", MessageLevel.msgOK, "  Current: " + (int)l_Current);
									} else {
										LogMsg("Get Switch", MessageLevel.msgError, "  Current value (State(2)) is not numeric: " + Switch.State(2));
									}

									if ((IsNumeric(Switch.State(0)) & IsNumeric(Switch.State(1)) & IsNumeric(Switch.State(2)))) {
										if (l_Current >= l_Min) {
											LogMsg("Get Switch", MessageLevel.msgOK, "  Current value is equal to or greater than the minimum value");
										} else {
											LogMsg("Get Switch", MessageLevel.msgError, "  Current value is less than the minimum value");
										}
										if (l_Current <= l_Max) {
											LogMsg("Get Switch", MessageLevel.msgOK, "  Current value is equal to or less than the maximum value");
										} else {
											LogMsg("Get Switch", MessageLevel.msgError, "  Current value is greater than the maximum value");
										}
									} else {
										LogMsg("Get Switch", MessageLevel.msgInfo, "Skipping numeric tests because one or more state values are not numeric");

									}

								default:
									//Unrecognised type of switch
									foreach ( l_State in l_Switch.State) {
										LogMsg("Get Switch", MessageLevel.msgOK, "  " + (int)l_State);
									}

									LogMsg("Set Switch", MessageLevel.msgInfo, "Unrecognised switch type: " + m_SwitchType + ", unable to test SetSwitch.");
							}
						}
					}
				} catch (COMException ex) {
					LogMsg("Switches", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("Switches", MessageLevel.msgError, EX_NET + ex.Message);

				}
		}
	}
	private override void CheckPerformance()
	{
		//Maxswitch
		//Have got a valid value
		if (m_CanReadMaxSwitch) {
			SwitchPerformanceTest(SwitchPropertyMethod.MaxSwitch, "MaxSwitch");
		} else {
			LogMsg("MaxSwitch", MessageLevel.msgInfo, "Test skipped as unable to read value");
		}
		//GetSwitch
		//Have got a valid value
		if ((bool)m_PerformanceGetSwitch) {
			SwitchPerformanceTest(SwitchPropertyMethod.GetSwitch, "GetSwitch");
		} else {
			LogMsg("GetSwitch", MessageLevel.msgInfo, "Test skipped as unable to read value");
		}
		//GetSwitchName
		//Have got a valid value
		if ((bool)m_PerformanceGetSwitchName) {
			SwitchPerformanceTest(SwitchPropertyMethod.GetSwitchName, "GetSwitchName");
		} else {
			LogMsg("GetSwitchName", MessageLevel.msgInfo, "Test skipped as unable to read value");
		}
		Status(StatusType.staTest, "");
		Status(StatusType.staAction, "");
		Status(StatusType.staStatus, "");
	}
	private void SwitchPerformanceTest(SwitchPropertyMethod p_Type, string p_Name)
	{
		System.DateTime l_StartTime;
		double l_Count;
		double l_LastElapsedTime;
		double l_ElapsedTime;
		string l_String;
		bool l_Boolean;
		double l_Rate;
		short l_Short;
		Status(StatusType.staTest, "Performance Testing");
		Status(StatusType.staAction, p_Name);
		try {
			l_StartTime = Now;
			l_Count = 0.0;
			l_LastElapsedTime = 0.0;
			do {
				l_Count += 1.0;
				switch (p_Type) {
					case SwitchPropertyMethod.MaxSwitch:
						l_Short = m_Switch.MaxSwitch;
					case SwitchPropertyMethod.GetSwitch:
						l_Boolean = m_Switch.GetSwitch((short)m_PerformanceGetSwitch);
					case SwitchPropertyMethod.GetSwitchName:
						l_String = m_Switch.GetSwitchName((short)m_PerformanceGetSwitchName);
					default:
						LogMsg(p_Name, MessageLevel.msgError, "SwitchPerformanceTest: Unknown test type " + p_Type.ToString);
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
			LogMsg(p_Name, MessageLevel.msgInfo, "Unable to complete test: " + ex.Message);
		}
	}
	#End Region
}
