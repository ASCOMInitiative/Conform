internal class FilterWheelTester : DeviceTesterBaseClass
{

	#Region "Variables and Constants"
		//Filter wheel command timeout (seconds)
	const int FILTER_WHEEL_TIME_OUT = 10;
	#If DEBUG Then
		#Else
	private object m_FilterWheel;
		#End If
	private ASCOM.DeviceInterface.IFilterWheelV2 m_FilterWheel;
	#End Region

	#Region "Enums"
	private enum FilterWheelProperties
	{
		FocusOffsets,
		Names,
		Position
	}
	#End Region

	#Region "New and Dispose"
	private FilterWheelTester()
	{
		base.New();
		this.HasMethods = false;
		this.HasPreRunCheck = true;
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
					m_FilterWheel.Connected = false;
				} catch {
				}
				try {
					m_FilterWheel.Dispose();
				} catch {
				}
				//Try : Marshal.ReleaseComObject(m_FilterWheel) : Catch : End Try
				m_FilterWheel = null;
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

		switch (g_FilterWheelProgID) {
			default:
				//I'm using the simulator values as the defaults since it is the reference platform
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040404;
				g_ExInvalidValue2 = 0x80040404;
				g_ExNotSet1 = 0x80040403;
		}
		base.CheckInitialise(g_FilterWheelProgID);

	}
	private override void CheckAccessibility()
	{
		object l_IFilterWheel;
		DriverAccess.FilterWheel l_DriverAccessFilterWheel;
		object l_DeviceObject = null;
		base.CheckAccessibility(g_FilterWheelProgID, DeviceType.Switch);
		//Try early binding
		l_IFilterWheel = null;
		try {
			l_DeviceObject = CreateObject(g_FilterWheelProgID);
			l_IFilterWheel = (ASCOM.Interface.IFilterWheel)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface IFilterWheel");
			try {
				l_IFilterWheel.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface IFilterWheel");
				l_IFilterWheel.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IFilterWheel");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IFilterWheel");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_IFilterWheel);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_IFilterWheel = null;
			GC.Collect();
		}

		try {
			l_DeviceObject = CreateObject(g_FilterWheelProgID);
			l_IFilterWheel = (ASCOM.DeviceInterface.IFilterWheelV2)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface IFilterWheelV2");
			try {
				l_IFilterWheel.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface IFilterWheelV2");
				l_IFilterWheel.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IFilterWheelV2");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IFilterWheelV2");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_IFilterWheel);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_IFilterWheel = null;
			GC.Collect();
		}

		//Try client access toolkit
		l_DriverAccessFilterWheel = null;
		try {
			l_DriverAccessFilterWheel = new ASCOM.DriverAccess.FilterWheel(g_FilterWheelProgID);
			LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
			try {
				l_DriverAccessFilterWheel.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
				l_DriverAccessFilterWheel.Connected = false;
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
				l_DriverAccessFilterWheel.Dispose();
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DriverAccessFilterWheel);
			} catch {
			}
			l_DriverAccessFilterWheel = null;
			GC.Collect();
		}
	}
	#If DEBUG Then
	private override void CreateDevice()
	{
		m_FilterWheel = CreateObject(g_FilterWheelProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a FilterWheel object");
		#Else
		m_FilterWheel = new ASCOM.DriverAccess.FilterWheel(g_FilterWheelProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.FilterWheel to get a FilterWheel object");
		#End If
		g_Stop = false;
		//connected OK so clear stop flag to allow other tests to run
	}
	private override bool Connected {
		get { Connected = base.Connected; }
		set {
			m_FilterWheel.Connected = value;
			base.Connected = value;
		}
	}
	public override void PreRunCheck()
	{
		System.DateTime StartTime;

		//Get into a consistent state
		SetStatus("FilterWheel Pre-run Check", "Wait one second for initialisation", "");
		WaitFor(1000);
		//Wait for 1 second to allow any movement to start
		StartTime = Now;
		try {
			do {
				SetStatus("FilterWheel Pre-run Check", "Waiting for movement to stop", Now.Subtract(StartTime).Seconds + " second(s)");
				WaitFor(SLEEP_TIME);
			} while (!((m_FilterWheel.Position != FWTEST_IS_MOVING) | (Now.Subtract(StartTime).TotalSeconds > FWTEST_TIMEOUT)));
			// Wait until movement has stopped or 30 seconds have passed
			if (m_FilterWheel.Position != FWTEST_IS_MOVING)
				LogMsg("Pre-run Check", MessageLevel.msgOK, "Filterwheel is stationary, ready to start tests");
		} catch (Exception ex) {
			LogMsg("Pre-run Check", MessageLevel.msgInfo, "Unable to determine that the Filterwheel is stationary");
			LogMsg("Pre-run Check", MessageLevel.msgError, "Exception: " + ex.ToString);
		}
		SetStatus("", "", "");

	}

	private override void CheckCommonMethods()
	{
		base.CheckCommonMethods(m_FilterWheel, DeviceType.FilterWheel);
	}


	private override void CheckProperties()
	{
		int[] l_Offsets;
		int l_NNames;
		int l_NOffsets;
		int l_Offset;
		int l_FilterNumber;
		string[] l_Names;
		string l_Name;

		//FocusOffsets - Required - Read only
		try {
			l_Offsets = m_FilterWheel.FocusOffsets;
			l_NOffsets = l_Offsets.Length;
			if (l_NOffsets == 0) {
				LogMsg("FocusOffsets Get", MessageLevel.msgError, "Found no offset values in the returned array");
			} else {
				LogMsg("FocusOffsets Get", MessageLevel.msgOK, "Found " + l_NOffsets.ToString + " filter offset values");
			}

			l_FilterNumber = 0;
			foreach ( l_Offset in l_Offsets) {
				LogMsg("FocusOffsets Get", MessageLevel.msgInfo, "Filter " + l_FilterNumber.ToString + " Offset: " + l_Offset.ToString);
				l_FilterNumber += 1;
				TestStop();
				if (g_Stop)
					break; // TODO: might not be correct. Was : Exit Try
			}
		} catch (COMException ex) {
			LogMsg("FocusOffsets Get", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("FocusOffsets Get", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//Names - Required - Read only
		try {
			l_Names = m_FilterWheel.Names;
			l_NNames = l_Names.Length;
			if (l_NNames == 0) {
				LogMsg("Names Get", MessageLevel.msgError, "Did not find any names in the returned array");
			} else {
				LogMsg("Names Get", MessageLevel.msgOK, "Found " + l_NNames.ToString + " filter names");
			}
			l_FilterNumber = 0;
			foreach ( l_Name in l_Names) {
				if (l_Name == null) {
					LogMsg("Names Get", MessageLevel.msgWarning, "Filter " + l_FilterNumber.ToString + " has a value of nothing");
				} else {
					if (l_Name == "") {
						LogMsg("Names Get", MessageLevel.msgWarning, "Filter " + l_FilterNumber.ToString + " has a value of \"\"");
					} else {
						LogMsg("Names Get", MessageLevel.msgInfo, "Filter " + l_FilterNumber.ToString + " Name: " + l_Name);
					}
				}
				l_FilterNumber += 1;
			}
		} catch (COMException ex) {
			LogMsg("Names Get", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("Names Get", MessageLevel.msgError, EX_NET + ex.ToString);
		}

		//Confirm number of array elements in filter names and filter offsets are the same
		if (l_NNames == l_NOffsets) {
			LogMsg("Names Get", MessageLevel.msgOK, "Number of filter offsets and number of names are the same: " + l_NNames.ToString);
		} else {
			LogMsg("Names Get", MessageLevel.msgError, "Number of filter offsets and number of names are different: " + l_NOffsets.ToString + " " + l_NNames.ToString);
		}

		//Position - Required - Read / Write
		int l_StartFilterNumber;
		short i;
		System.DateTime l_StartTime;
		System.DateTime l_EndTime;
		switch (l_NOffsets) {
			case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
0:
				LogMsg("Position", MessageLevel.msgWarning, "Filter position tests skipped as number of filters appears to be 0: " + l_NOffsets.ToString);
			default:
				try {
					l_StartFilterNumber = m_FilterWheel.Position;
					//Illegal starting position provided
					if ((l_StartFilterNumber < 0) | (l_StartFilterNumber >= l_NOffsets)) {
						LogMsg("Position Get", MessageLevel.msgError, "Illegal filter posiiton returned: " + l_StartFilterNumber.ToString);
					//Legal value returned
					} else {
						LogMsg("Position Get", MessageLevel.msgOK, "Currently at position: " + i.ToString);
						for (i = 0; i <= (short)l_NOffsets - 1; i++) {
							try {
								m_FilterWheel.Position = i;
								l_StartTime = Now;
								do {
									System.Threading.Thread.Sleep(100);
									TestStop();
								} while (!((m_FilterWheel.Position == i) | (Now.Subtract(l_StartTime).TotalSeconds > FILTER_WHEEL_TIME_OUT) | g_Stop));
								if (g_Stop)
									return;

								l_EndTime = Now;
								//We did get to position within the timeout period
								if (m_FilterWheel.Position == i) {
									LogMsg("Position Set", MessageLevel.msgOK, "Reached position: " + i.ToString + " in: " + Format(l_EndTime.Subtract(l_StartTime).TotalSeconds, "0.0") + " seconds");
								//Timed out so give error message
								} else {
									LogMsg("Position Set", MessageLevel.msgError, "Filter wheel did not reach specified position: " + i.ToString + " within timeout of: " + FILTER_WHEEL_TIME_OUT.ToString);
								}
								WaitFor(1000);
								//Pause to allow filter wheel to stabilise
							} catch (COMException ex) {
								LogMsg("Position Set", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
							} catch (Exception ex) {
								LogMsg("Position Set", MessageLevel.msgError, EX_NET + ex.ToString);
							}
						}
						//Confirm that an error is correctly generated for outside range values
						try {
							m_FilterWheel.Position = -1;
							//This should be 1 above the highest array element returned
							LogMsg("Position Set", MessageLevel.msgError, "Failed to generate exception when selecting filter with negative filter number");
						} catch (COMException ex) {
							switch (ex.ErrorCode) {
								case ErrorCodes.InvalidValue:
								case g_ExInvalidValue1:
								case g_ExInvalidValue2:
									LogMsg("Position Set", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on selecting filter with negative filter number");
								default:
									LogMsg("Position Set", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
							}
						} catch (DriverException ex) {
							switch (ex.Number) {
								case ErrorCodes.InvalidValue:
								case g_ExInvalidValue1:
								case g_ExInvalidValue2:
									LogMsg("Position Set", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on selecting filter with negative filter number");
								default:
									LogMsg("Position Set", MessageLevel.msgError, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
							}
						} catch (Exception ex) {
							LogMsg("Position Set", MessageLevel.msgError, "Unexpected exception on selecting filter with negative filter number: " + ex.Message);
						}
						//Confirm that an error is correctly generated for outside range values
						try {
							m_FilterWheel.Position = (short)l_NOffsets;
							//This should be 1 above the highest array element returned
							LogMsg("Position Set", MessageLevel.msgError, "Failed to generate exception when selecting filter outside expected range");
						} catch (COMException ex) {
							switch (ex.ErrorCode) {
								case ErrorCodes.InvalidValue:
								case g_ExInvalidValue1:
								case g_ExInvalidValue2:
									LogMsg("Position Set", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on selecting filter outside expected range");
								default:
									LogMsg("Position Set", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
							}
						} catch (DriverException ex) {
							switch (ex.Number) {
								case ErrorCodes.InvalidValue:
								case g_ExInvalidValue1:
								case g_ExInvalidValue2:
									LogMsg("Position Set", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on selecting filter outside expected range");
								default:
									LogMsg("Position Set", MessageLevel.msgError, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
							}
						} catch (Exception ex) {
							LogMsg("Position Set", MessageLevel.msgError, "Unexpected exception on selecting filter outside expected range: " + ex.Message);
						}
					}
				} catch (COMException ex) {
					LogMsg("Position Get", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("Position Get", MessageLevel.msgError, EX_NET + ex.ToString);

				}
		}
	}
	private override void CheckPerformance()
	{
		FilterWheelPerformanceTest(FilterWheelProperties.FocusOffsets, "FocusOffsets");
		FilterWheelPerformanceTest(FilterWheelProperties.Names, "Names");
		FilterWheelPerformanceTest(FilterWheelProperties.Position, "Position");
	}
	private void FilterWheelPerformanceTest(FilterWheelProperties p_Type, string p_Name)
	{
		int[] l_Offsets;
		string[] l_Names;
		int l_StartFilterNumber;
		System.DateTime l_StartTime;
		double l_Count;
		double l_LastElapsedTime;
		double l_ElapsedTime;
		double l_Rate;
		Status(StatusType.staTest, "Performance Test");
		Status(StatusType.staAction, p_Name);
		try {
			l_StartTime = Now;
			l_Count = 0.0;
			l_LastElapsedTime = 0.0;
			do {
				l_Count += 1.0;
				switch (p_Type) {
					case FilterWheelProperties.FocusOffsets:
						l_Offsets = m_FilterWheel.FocusOffsets;
					case FilterWheelProperties.Names:
						l_Names = m_FilterWheel.Names;
					case FilterWheelProperties.Position:
						l_StartFilterNumber = m_FilterWheel.Position;
					default:
						LogMsg(p_Name, MessageLevel.msgError, "FilterWheelPerformanceTest: Unknown test type " + p_Type.ToString);
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
