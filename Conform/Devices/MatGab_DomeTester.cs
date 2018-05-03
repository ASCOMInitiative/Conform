using ASCOM.DeviceInterface;
internal class DomeTester : DeviceTesterBaseClass
{
	#Region "Variables and Constants"
	//These constants have been moved to the applications setting object so that they can be changed by the user, the same default values have been retained
	//Const DOME_SHUTTER_TIMEOUT As Integer = 60 'Timeout for dome to open or close shutter
	//Const DOME_AZIMUTH_TIMEOUT As Integer = 180 'Timeout for dome to move to azimuth
	//Const DOME_ALTITUDE_TIMEOUT As Integer = 120 'Timeout for dome to move to altitude
	//Const DOME_POST_COMMAND_WAIT_TIME As Integer = 5000 '

		//Amount to offset the azimuth when testing ability to sync
	const double DOME_SYNC_OFFSET = 45.0;
		//Illegal value to test dome driver exception generation
	const double DOME_ILLEGAL_ALTITUDE_LOW = -10.0;
		//Illegal value to test dome driver exception generation
	const double DOME_ILLEGAL_ALTITUDE_HIGH = 100.0;
		//Illegal value to test dome driver exception generation
	const double DOME_ILLEGAL_AZIMUTH_LOW = -10.0;
		//Illegal value to test dome driver exception generation
	const double DOME_ILLEGAL_AZIMUTH_HIGH = 370.0;

	//Dome variables
	private bool m_CanSetAltitude;
	private bool m_CanSetAzimuth;
	private bool m_CanSetShutter;
	private bool m_CanSlave;
	private bool m_CanSyncAzimuth;
	private bool m_Slaved;
	private ShutterState m_ShutterStatus;
	private bool m_CanReadAltitude;
	private bool m_CanReadAtPark;
	private bool m_CanReadAtHome;
	private bool m_CanReadSlewing;
	private bool m_CanReadSlaved;
	private bool m_CanReadShutterStatus;
	private bool m_CanReadAzimuth;
	private bool m_CanSlewToAzimuth;
	private bool m_AsyncSlewAzimuth;

	private bool m_AsyncSlewAltitude;
	//General variables
	private bool m_Slewing;
	private bool m_AtHome;
	private bool m_AtPark;
	private bool m_CanFindHome;
	private bool m_CanFindPark;
	private bool m_CanPark;
	private bool m_CanSetPark;
	private bool m_Connected;
	private string m_Description;
	private string m_DriverINfo;
	private string m_Name;
	private short m_InterfaceVersion;
	private double m_Altitude;

	private double m_Azimuth;

	#If DEBUG Then
		#Else
	private object m_Dome;
		#End If
	private IDomeV2 m_Dome;
	#End Region

	#Region "Enums"
	private enum DomePropertyMethod
	{
		//Properties
		Altitude,
		AtHome,
		AtPark,
		Azimuth,
		CanFindHome,
		CanPark,
		CanSetAltitude,
		CanSetAzimuth,
		CanSetPark,
		CanSetShutter,
		CanSlave,
		CanSyncAzimuth,
		Connected,
		Description,
		DriverInfo,
		InterfaceVersion,
		Name,
		ShutterStatus,
		SlavedRead,
		SlavedWrite,
		Slewing,

		//Methods
		AbortSlew,
		CloseShutter,
		CommandBlind,
		CommandBool,
		CommandString,
		FindHome,
		OpenShutter,
		Park,
		SetPark,
		SlewToAltitude,
		SlewToAzimuth,
		SyncToAzimuth
	}
	#End Region

	#Region "New and Dispose"
	private DomeTester()
	{
		base.New();
		this.HasCanProperties = true;
		this.HasPreRunCheck = true;
		this.HasPostRunCheck = true;
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
					m_Dome.Connected = false;
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_Dome);
				} catch {
				}
				m_Dome = null;
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
		switch (g_DomeProgID.ToUpper) {
			default:
				//I'm using the simulator values as the defaults since it is the reference platform
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040405;
				g_ExInvalidValue2 = 0x80040405;
				g_ExNotSet1 = 0x80040403;
		}
		//GetIConformValues() 'Read values from IConform if it is present
		base.CheckInitialise(g_DomeProgID);

	}
	private override void CheckAccessibility()
	{
		object l_IDome;
		DriverAccess.Dome l_DriverAccessDome;
		object l_DeviceObject = null;
		base.CheckAccessibility(g_DomeProgID, DeviceType.Dome);
		//Try early binding
		l_IDome = null;
		try {
			l_DeviceObject = CreateObject(g_DomeProgID);
			l_IDome = (ASCOM.Interface.IDome)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through interface IDome");
			try {
				l_IDome.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface IDome");
				l_IDome.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IDome");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IDome");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_IDome);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_IDome = null;
			GC.Collect();
		}

		try {
			l_DeviceObject = CreateObject(g_DomeProgID);
			l_IDome = (ASCOM.DeviceInterface.IDomeV2)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver through interface IDomeV2");
			try {
				l_IDome.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface IDomeV2");
				l_IDome.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IDomeV2");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface IDomeV2");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_IDome);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_IDome = null;
			GC.Collect();
		}
		WaitForAbsolute(2 * DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system");

		//Try client access toolkit
		l_DriverAccessDome = null;
		try {
			l_DriverAccessDome = new ASCOM.DriverAccess.Dome(g_DomeProgID);
			LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
			try {
				l_DriverAccessDome.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
				l_DriverAccessDome.Connected = false;
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
				l_DriverAccessDome.Dispose();
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DriverAccessDome);
			} catch {
			}
			l_DriverAccessDome = null;
			GC.Collect();
		}

		WaitForAbsolute(2 * DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system");

	}
	#If DEBUG Then
	private override void CreateDevice()
	{
		m_Dome = CreateObject(g_DomeProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Dome object");
		#Else
		m_Dome = new ASCOM.DriverAccess.Dome(g_DomeProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Dome to get a Dome object");
		#End If
		g_Stop = false;
	}
	private override bool Connected {
		get { Connected = m_Dome.Connected; }
		set {
			m_Dome.Connected = value;
			g_Stop = false;
		}
	}
	private override void CheckCommonMethods()
	{
		base.CheckCommonMethods(m_Dome, DeviceType.Dome);
	}


	public override void PreRunCheck()
	{
		int l_VStringPtr;
		int l_V1;
		int l_V2;
		int l_V3;
		string l_VString;
		Microsoft.VisualBasic.MsgBoxResult l_MsgBoxResult;

		//Add a test for a back level version of the Dome simulator - just abandon this process if any errors occur
		if (g_DomeProgID.ToUpper == "DOMESIM.DOME") {
			l_VString = "";
			try {
				l_VStringPtr = InStr(m_Dome.DriverInfo.ToUpper, "ASCOM DOME SIMULATOR ");
				//Point at the start of the version string
				//There is a version string so read what it is
				if (l_VStringPtr > 0) {
					l_VString = Mid(m_Dome.DriverInfo.ToUpper, l_VStringPtr + 21);
					//Get the version string
					l_VStringPtr = InStr(l_VString, ".");
					//there is a first version number part
					if (l_VStringPtr > 1) {
						l_V1 = (int)Mid(l_VString, 1, l_VStringPtr - 1);
						//Extract the number
						l_VString = Mid(l_VString, l_VStringPtr + 1);
						//Get the second version number part
						l_VStringPtr = InStr(l_VString, ".");
						//there is a second version number part
						if (l_VStringPtr > 1) {
							l_V2 = (int)Mid(l_VString, 1, l_VStringPtr - 1);
							//Extract the number
							l_VString = Mid(l_VString, l_VStringPtr + 1);
							//Get the third version number part
							//Find the next non numeric character
							l_VStringPtr = 0;
							do {
								l_VStringPtr += 1;
							} while (!(!IsNumeric(Mid(l_VString, l_VStringPtr, 1))));
							//there is a third version number part
							if (l_VStringPtr > 1) {
								l_V3 = (int)Mid(l_VString, 1, l_VStringPtr - 1);
								//Extract the number
								//Turn the version parts into a whole number
								l_V1 = l_V1 * 1000000 + l_V2 * 1000 + l_V3;
								//We are using a version earlier that 5.0.7 so issue a message
								if (l_V1 < 5000007) {
									LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the dome simulator has known conformance issues, ***");
									LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site http://ascom-standards.org/Downloads/Index.htm ***");
									LogMsg("", MessageLevel.msgAlways, "");
								} else {
									LogMsg("Version Check", MessageLevel.msgDebug, "Version check OK");
								}
							}
						}
					}
				}
			} catch (Exception ex) {
				LogMsg("ConformanceCheck", MessageLevel.msgError, ex.ToString);
			}
		}
		if (!TestStop()) {
			//Get into a consistent state
			try {
				DomeWaitForSlew(g_Settings.DomeAzimuthTimeout);
				//Wait for slewing to finish
			} catch (Exception ex) {
			}
			if (frmConformMain.chkDomeShutter.Checked) {
				LogMsg("DomeSafety", MessageLevel.msgComment, "Attempting to open shutter as some tests may fail if it is closed...");
				try {
					m_Dome.OpenShutter();
					try {
						DomeShutterWait(ShutterState.shutterOpen);
					} catch {
					}
					//Stop button pressed
					if (TestStop()) {
						LogMsg("DomeSafety", MessageLevel.msgComment, "Stop button pressed, further testing abandoned, shutter status: " + m_Dome.ShutterStatus.ToString);
					//Got to end of test
					} else {
						if (m_Dome.ShutterStatus == ShutterState.shutterOpen) {
							LogMsg("DomeSafety", MessageLevel.msgOK, "Shutter status: " + m_Dome.ShutterStatus.ToString);
						} else {
							LogMsg("DomeSafety", MessageLevel.msgWarning, "Shutter status: " + m_Dome.ShutterStatus.ToString);
						}
					}
				} catch (Exception ex) {
					LogMsg("DomeSafety", MessageLevel.msgComment, "Unable to open shutter: " + ex.Message);
					LogMsg("DomeSafety", MessageLevel.msgComment, "Prompting for manual intervention...");
					l_MsgBoxResult = MsgBox("Please open shutter and press OK or press Cancel to stop tests", MsgBoxStyle.OkCancel | MsgBoxStyle.SystemModal, "Dome Safety");
					//Cancel was pressed
					if (l_MsgBoxResult == MsgBoxResult.Cancel) {
						g_Stop = true;
						//Stop if cancel was pressed
						LogMsg("DomeSafety", MessageLevel.msgComment, "Cancel pressed, further testing abandoned");
					} else {
						LogMsg("DomeSafety", MessageLevel.msgComment, "OK pressed, testing continues");
					}

				}
				Status(StatusType.staTest, "");
			} else {
				LogMsg("DomeSafety", MessageLevel.msgComment, "Open shutter check box is unchecked so shutter not opened");
			}
		}
	}
	public override void ReadCanProperties()
	{
		DomeMandatoryTest(DomePropertyMethod.CanFindHome, "CanFindHome");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.CanPark, "CanPark");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.CanSetAltitude, "CanSetAltitude");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.CanSetAzimuth, "CanSetAzimuth");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.CanSetPark, "CanSetPark");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.CanSetShutter, "CanSetShutter");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.CanSlave, "CanSlave");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.CanSyncAzimuth, "CanSyncAzimuth");
		if (TestStop())
			return;

	}
	private override void CheckProperties()
	{
		DomeOptionalTest(DomePropertyMethod.Altitude, "Altitude");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.AtHome, "AtHome");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.AtPark, "AtPark");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.Azimuth, "Azimuth");
		if (TestStop())
			return;
		//DomeMandatoryTest(DomePropertyMethod.Connected, "Connected") : If TestStop() Then Exit Sub
		//DomeMandatoryTest(DomePropertyMethod.Description, "Description") : If TestStop() Then Exit Sub
		//DomeMandatoryTest(DomePropertyMethod.DriverInfo, "DriverInfo") : If TestStop() Then Exit Sub
		//DomeMandatoryTest(DomePropertyMethod.InterfaceVersion, "InterfaceVersion") : If TestStop() Then Exit Sub
		//DomeMandatoryTest(DomePropertyMethod.Name, "Name") : If TestStop() Then Exit Sub
		DomeOptionalTest(DomePropertyMethod.ShutterStatus, "ShutterStatus");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.SlavedRead, "Slaved Read");
		if (TestStop())
			return;
		if (m_Slaved & (!m_CanSlave))
			LogMsg("Slaved Read", MessageLevel.msgError, "Dome is slaved but CanSlave is false");
		DomeOptionalTest(DomePropertyMethod.SlavedWrite, "Slaved Write");
		if (TestStop())
			return;
		DomeMandatoryTest(DomePropertyMethod.Slewing, "Slewing");
		if (TestStop())
			return;
	}
	private override void CheckMethods()
	{
		DomeMandatoryTest(DomePropertyMethod.AbortSlew, "AbortSlew");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.SlewToAltitude, "SlewToAltitude");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.SlewToAzimuth, "SlewToAzimuth");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.SyncToAzimuth, "SyncToAzimuth");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.CloseShutter, "CloseShutter");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.OpenShutter, "OpenShutter");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.FindHome, "FindHome");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.Park, "Park");
		if (TestStop())
			return;
		DomeOptionalTest(DomePropertyMethod.SetPark, "SetPark");
		if (TestStop())
			return;
		//SetPark must follow Park
		//DomeOptionalTest(DomePropertyMethod.CommandBlind, "CommandBlind") : If TestStop() Then Exit Sub
		//DomeOptionalTest(DomePropertyMethod.CommandBool, "CommandBool") : If TestStop() Then Exit Sub
		//DomeOptionalTest(DomePropertyMethod.CommandString, "CommandString") : If TestStop() Then Exit Sub
	}
	public override void CheckPerformance()
	{
		if (m_CanReadAltitude){DomePerformanceTest(DomePropertyMethod.Altitude, "Altitude");if (TestStop())
			return;
}
		if (m_CanReadAzimuth){DomePerformanceTest(DomePropertyMethod.Azimuth, "Azimuth");if (TestStop())
			return;
}
		if (m_CanReadShutterStatus){DomePerformanceTest(DomePropertyMethod.ShutterStatus, "ShutterStatus");if (TestStop())
			return;
}
		if (m_CanReadSlaved){DomePerformanceTest(DomePropertyMethod.SlavedRead, "Slaved");if (TestStop())
			return;
}
		if (m_CanReadSlewing){DomePerformanceTest(DomePropertyMethod.Slewing, "Slewing");if (TestStop())
			return;
}
	}
	public override void PostRunCheck()
	{
		if (frmConformMain.chkDomeShutter.Checked) {
			if (m_CanSetShutter) {
				LogMsg("DomeSafety", MessageLevel.msgInfo, "Attempting to close shutter...");
				//Close shutter
				try {
					m_Dome.CloseShutter();
					DomeShutterWait(ShutterState.shutterClosed);
					LogMsg("DomeSafety", MessageLevel.msgOK, "Shutter successfully closed");
				} catch (Exception ex) {
					LogMsg("DomeSafety", MessageLevel.msgComment, "Exception closing shutter: " + ex.Message);
					LogMsg("DomeSafety", MessageLevel.msgComment, "Please close shutter manually");
				}
			} else {
				LogMsg("DomeSafety", MessageLevel.msgInfo, "CanSetShutter is false, please close the shutter manually");
			}
		} else {
			LogMsg("DomeSafety", MessageLevel.msgInfo, "Open shutter check box is unchecked so close shutter bypassed");
		}
		//3.0.0.17 - Added check for CanPark
		//Can park the dome
		if (m_CanPark) {
			LogMsg("DomeSafety", MessageLevel.msgInfo, "Attempting to park dome...");
			//Park
			try {
				m_Dome.Park();
				DomeWaitForSlew(g_Settings.DomeAzimuthTimeout);
				LogMsg("DomeSafety", MessageLevel.msgOK, "Dome successfully parked");
			} catch (Exception ex) {
				LogMsg("DomeSafety", MessageLevel.msgError, "Exception generated, unable to park dome");
			}
		} else {
			LogMsg("DomeSafety", MessageLevel.msgInfo, "CanPark is false - skipping dome parking");
		}

	}

	private void DomeSlewToAltitude(string p_Name, double p_Altitude)
	{
		System.DateTime l_StartTime;
		Status(StatusType.staAction, "Slew to " + p_Altitude + " degrees");
		m_Dome.SlewToAltitude(p_Altitude);
		//Can read slewing so make sure dome is at rest
		if (m_CanReadSlewing) {
			l_StartTime = Now;
			//Async slew
			if (m_Dome.Slewing) {
				DomeWaitForSlew(g_Settings.DomeAltitudeTimeout);
				if (TestStop())
					return;
				m_AsyncSlewAltitude = true;
				LogMsg(p_Name + " " + p_Altitude, MessageLevel.msgOK, "Asynchronous slew OK");
			//Sync slew
			} else {
				m_AsyncSlewAltitude = false;
				LogMsg(p_Name + " " + p_Altitude, MessageLevel.msgOK, "Synchronous slew OK");
			}
		//Can't read slewing so just assume synchronous
		} else {
			LogMsg(p_Name + " " + p_Altitude, MessageLevel.msgOK, "Can't read Slewing so assume synchronous slew OK");
		}
		DomeStabliisationWait();
	}
	private void DomeSlewToAzimuth(string p_Name, double p_Azimuth)
	{
		Status(StatusType.staAction, "Slew to " + p_Azimuth + " degrees");
		//Only check for successful operation on legal values
		if (p_Azimuth >= 0.0 & p_Azimuth <= 359.9999999) {
			m_CanSlewToAzimuth = false;
			m_Dome.SlewToAzimuth(p_Azimuth);
			m_CanSlewToAzimuth = true;
			//Command is supported and didn't generate an exception
		} else {
			m_Dome.SlewToAzimuth(p_Azimuth);
		}
		//Can read slewing so make sure dome is at rest
		if (m_CanReadSlewing) {
			//Async slew
			if (m_Dome.Slewing) {
				DomeWaitForSlew(g_Settings.DomeAzimuthTimeout);
				if (TestStop())
					return;
				m_AsyncSlewAzimuth = true;
				LogMsg(p_Name + " " + p_Azimuth, MessageLevel.msgOK, "Asynchronous slew OK");
			//Sync slew
			} else {
				m_AsyncSlewAzimuth = false;
				LogMsg(p_Name + " " + p_Azimuth, MessageLevel.msgOK, "Synchronous slew OK");
			}
		//Can't read slewing so just assume synchronous
		} else {
			LogMsg(p_Name + " " + p_Azimuth, MessageLevel.msgOK, "Can't read Slewing so assume synchronous slew OK");
		}
		DomeStabliisationWait();
	}
	private void DomeWaitForSlew(double p_TimeOut)
	{
		System.DateTime l_StartTime;
		l_StartTime = Now;
		do {
			WaitFor(SLEEP_TIME);
			Application.DoEvents();
			Status(StatusType.staStatus, "Slewing Status: " + m_Dome.Slewing + ", Timeout: " + Format(Now.Subtract(l_StartTime).TotalSeconds, "#0") + "/" + p_TimeOut + ", press stop to abandon wait");
		} while (!(!m_Dome.Slewing | TestStop() | (Now.Subtract(l_StartTime).TotalSeconds > p_TimeOut)));
		Status(StatusType.staStatus, "");
		if ((Now.Subtract(l_StartTime).TotalSeconds > p_TimeOut)) {
			LogMsg("DomeWaitForSlew", MessageLevel.msgError, "Timed out waiting for Dome slew, consider increasing timeouts in Options/Conform Options");
		}
	}
	private void DomeMandatoryTest(DomePropertyMethod p_Type, string p_Name)
	{
		try {
			switch (p_Type) {
				//Properties
				case DomePropertyMethod.CanFindHome:
					m_CanFindHome = m_Dome.CanFindHome;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanFindHome.ToString);
				case DomePropertyMethod.CanPark:
					m_CanPark = m_Dome.CanPark;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanPark.ToString);
				case DomePropertyMethod.CanSetAltitude:
					m_CanSetAltitude = m_Dome.CanSetAltitude;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSetAltitude.ToString);
				case DomePropertyMethod.CanSetAzimuth:
					m_CanSetAzimuth = m_Dome.CanSetAzimuth;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSetAzimuth.ToString);
				case DomePropertyMethod.CanSetPark:
					m_CanSetPark = m_Dome.CanSetPark;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSetPark.ToString);
				case DomePropertyMethod.CanSetShutter:
					m_CanSetShutter = m_Dome.CanSetShutter;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSetShutter.ToString);
				case DomePropertyMethod.CanSlave:
					m_CanSlave = m_Dome.CanSlave;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSlave.ToString);
				case DomePropertyMethod.CanSyncAzimuth:
					m_CanSyncAzimuth = m_Dome.CanSyncAzimuth;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSyncAzimuth.ToString);
				case DomePropertyMethod.Connected:
					m_Connected = m_Dome.Connected;
					LogMsg(p_Name, MessageLevel.msgOK, m_Connected.ToString);
				case DomePropertyMethod.Description:
					m_Description = m_Dome.Description;
					LogMsg(p_Name, MessageLevel.msgOK, m_Description.ToString);
				case DomePropertyMethod.Description:
					m_Description = m_Dome.Description;
					LogMsg(p_Name, MessageLevel.msgOK, m_Description.ToString);
				case DomePropertyMethod.DriverInfo:
					m_DriverINfo = m_Dome.DriverInfo;
					LogMsg(p_Name, MessageLevel.msgOK, m_DriverINfo.ToString);
				case DomePropertyMethod.InterfaceVersion:
					m_InterfaceVersion = m_Dome.InterfaceVersion;
					LogMsg(p_Name, MessageLevel.msgOK, m_InterfaceVersion.ToString);
				case DomePropertyMethod.Name:
					m_Name = m_Dome.Name;
					LogMsg(p_Name, MessageLevel.msgOK, m_Name.ToString);
				case DomePropertyMethod.SlavedRead:
					m_CanReadSlaved = false;
					m_Slaved = m_Dome.Slaved;
					m_CanReadSlaved = true;
					LogMsg(p_Name, MessageLevel.msgOK, m_Slaved.ToString);
				case DomePropertyMethod.Slewing:
					m_CanReadSlewing = false;
					m_Slewing = m_Dome.Slewing;
					m_CanReadSlewing = true;

					LogMsg(p_Name, MessageLevel.msgOK, m_Slewing.ToString);
				//Methods
				case DomePropertyMethod.AbortSlew:
					m_Dome.AbortSlew();
					//Confirm that slaved is false
					//Check it is false
					if (m_CanReadSlaved) {
						//Should not be showing slaved after abortslew
						if (m_Dome.Slaved) {
							LogMsg("AbortSlew", MessageLevel.msgError, "Slaved property is true after AbortSlew");
						// Not slaved so fine
						} else {
							LogMsg("AbortSlew", MessageLevel.msgOK, "AbortSlew command issued successfully");
						}
					//Assume all is OK
					} else {
						LogMsg("AbortSlew", MessageLevel.msgOK, "Can't read Slaved property AbortSlew command was successful");
					}
				default:
					LogMsg(p_Name, MessageLevel.msgError, "DomeMandatoryTest: Unknown test type " + p_Type.ToString);
			}
		} catch (COMException ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.Message);
		}
	}
	private void DomeOptionalTest(DomePropertyMethod p_Type, string p_Name)
	{
		double l_SlewAngle;
		double l_OriginalAzimuth;
		double l_NewAzimuth;
		try {
			switch (p_Type) {
				//Properties
				case DomePropertyMethod.Altitude:
					m_CanReadAltitude = false;
					m_Altitude = m_Dome.Altitude;
					m_CanReadAltitude = true;
					LogMsg(p_Name, MessageLevel.msgOK, m_Altitude.ToString);
				case DomePropertyMethod.AtHome:
					m_CanReadAtHome = false;
					m_AtHome = m_Dome.AtHome;
					m_CanReadAtHome = true;
					LogMsg(p_Name, MessageLevel.msgOK, m_AtHome.ToString);
				case DomePropertyMethod.AtPark:
					m_CanReadAtPark = false;
					m_AtPark = m_Dome.AtPark;
					m_CanReadAtPark = true;
					LogMsg(p_Name, MessageLevel.msgOK, m_AtPark.ToString);
				case DomePropertyMethod.Azimuth:
					m_CanReadAzimuth = false;
					m_Azimuth = m_Dome.Azimuth;
					m_CanReadAzimuth = true;
					LogMsg(p_Name, MessageLevel.msgOK, m_Azimuth.ToString);
				case DomePropertyMethod.ShutterStatus:
					m_CanReadShutterStatus = false;
					m_ShutterStatus = m_Dome.ShutterStatus;
					m_CanReadShutterStatus = true;
					m_ShutterStatus = (ShutterState)m_ShutterStatus;
					//In relase mode only an integer value is returned
					LogMsg(p_Name, MessageLevel.msgOK, m_ShutterStatus.ToString);
				case DomePropertyMethod.SlavedWrite:
					//Write test should succeed
					if (m_CanSlave) {
						//Can read so worth testing write
						if (m_CanReadSlaved) {
							//We are slaved so try and turn it off
							if (m_Slaved) {
								m_Dome.Slaved = false;
							//Not slaved so try and turn on
							} else {
								m_Dome.Slaved = true;
							}
							m_Dome.Slaved = m_Slaved;
							//Restore original value
							LogMsg("Slaved Write", MessageLevel.msgOK, "Slave state changed successfully");
						//Can't read so skip test
						} else {
							LogMsg("Slaved Write", MessageLevel.msgInfo, "Test skipped since Slaved property can't be read");
						}
					//Slaved write should generate an exception
					} else {
						m_Dome.Slaved = true;
						LogMsg(p_Name, MessageLevel.msgError, "CanSlave is false but setting Slaved true did not raise an exception");
						m_Dome.Slaved = false;
						//Unslave to continue tests
					}
				//Methods
				case DomePropertyMethod.CloseShutter:
					//CloseShutter should work OK
					if (m_CanSetShutter) {
						DomeShutterTest(ShutterState.shutterClosed, p_Name);
						DomeStabliisationWait();
					//CloseShutter should throw an exception
					} else {
						//2.1.5 next line changed because close shutter was never tested
						//DomeShutterTest(ShutterState.shutterClosed, p_Name)
						m_Dome.CloseShutter();
						LogMsg(p_Name, MessageLevel.msgError, "CanSetShutter is false but CloseShutter did not raise an exception");
					}
				case DomePropertyMethod.CommandBlind:
					m_Dome.CommandBlind("");
					//m_Dome.CommandBlind("", True)
					LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent");
				case DomePropertyMethod.CommandBool:
					m_Dome.CommandBool("");
					//m_Dome.CommandBool("", True)
					LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent");
				case DomePropertyMethod.CommandString:
					m_Dome.CommandString("");
					//m_Dome.CommandString("", True)
					LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent");
				case DomePropertyMethod.FindHome:
					if (m_CanFindHome) {
						Status(StatusType.staTest, p_Name);
						Status(StatusType.staAction, "Waiting for movement to stop");
						m_Dome.FindHome();
						//Check whether slaved is true, if it is then Park  should have raised an exception and didn't
						if (m_CanReadSlaved) {
							if (m_Dome.Slaved)
								LogMsg(p_Name, MessageLevel.msgError, "Slaved is true but Home did not raise an exception");
						}
						//Make sure dome is at rest
						if (m_CanReadSlewing) {
							do {
								WaitFor(SLEEP_TIME);
								Application.DoEvents();
								Status(StatusType.staStatus, "Slewing Status: " + m_Dome.Slewing);
							} while (!(!m_Dome.Slewing | TestStop()));
						}
						//Only do remaining tests if stop hasn't been pressed
						if (!TestStop()) {
							//Can read athome so confirm that it 
							if (m_CanReadAtHome) {
								//Dome shows as homed - hooray!
								if (m_Dome.AtHome) {
									LogMsg(p_Name, MessageLevel.msgOK, "Dome homed successfully");
								//Home completed but apparently dome isn't homed!
								} else {
									LogMsg(p_Name, MessageLevel.msgError, "Home command completed but AtHome is false");
								}
							} else {
								LogMsg(p_Name, MessageLevel.msgOK, "Can't read AtHome so assume that dome has homed successfully");
							}
							DomeStabliisationWait();
						}
					//CanFindHome is false so FindHome should throw an exception
					} else {
						m_Dome.FindHome();
						LogMsg(p_Name, MessageLevel.msgError, "CanFindHome is false but FindHome did not throw an exception");
					}
				case DomePropertyMethod.OpenShutter:
					if (m_CanSetShutter) {
						DomeShutterTest(ShutterState.shutterOpen, p_Name);
					} else {
						//2.1.5 Next line changed to ensure that an open shutter is tried
						//DomeShutterTest(ShutterState.shutterOpen, p_Name)
						m_Dome.OpenShutter();
						LogMsg(p_Name, MessageLevel.msgError, "CanSetShutter is false but OpenShutter did not raise an exception");

					}
				case DomePropertyMethod.Park:
					//Should be able to issue the Park command
					if (m_CanPark) {
						Status(StatusType.staTest, p_Name);
						Status(StatusType.staAction, "Waiting for movement to stop");
						m_Dome.Park();
						//Check whether slaved is true, if it is then Park  should have raised an exception and didn't
						if (m_CanReadSlaved) {
							if (m_Dome.Slaved)
								LogMsg(p_Name, MessageLevel.msgError, "Slaved is true but Park did not raise an exception");
						}
						//Make sure dome is at rest
						if (m_CanReadSlewing) {
							do {
								WaitFor(SLEEP_TIME);
								Application.DoEvents();
								Status(StatusType.staStatus, "Slewing Status: " + m_Dome.Slewing);
							} while (!(!m_Dome.Slewing | TestStop()));
						}
						//Only do remain tests if stop hasn't been pressed
						if (!TestStop()) {
							//Can read at park so confirm that it 
							if (m_CanReadAtPark) {
								//Dome shows as parked - hooray!
								if (m_Dome.AtPark) {
									LogMsg(p_Name, MessageLevel.msgOK, "Dome parked successfully");
								//Park completed but apparently dome isn't parked!
								} else {
									LogMsg(p_Name, MessageLevel.msgError, "Park command completed but AtPark is false");
								}
							} else {
								LogMsg(p_Name, MessageLevel.msgOK, "Can't read AtPark so assume that dome has parked successfully");
							}
						}
						DomeStabliisationWait();
					//Park command should raise an error
					} else {
						m_Dome.Park();
						LogMsg(p_Name, MessageLevel.msgError, "CanPark is false but Park did not raise an exception");
					}
				case DomePropertyMethod.SetPark:
					//Should be able to set park so try it
					if (m_CanSetPark) {
						try {
							m_Dome.SetPark();
							LogMsg(p_Name, MessageLevel.msgOK, "SetPark issued OK");
						} catch (COMException ex) {
							LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
						} catch (Exception ex) {
							LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.Message);
						}
					//Can't set park so should raise an error
					} else {
						m_Dome.SetPark();
						LogMsg(p_Name, MessageLevel.msgError, "CanSetPath is false but SetPath did not throw an exception");
					}
				case DomePropertyMethod.SlewToAltitude:
					//SlewToAltitude should work OK
					if (m_CanSetAltitude) {
						Status(StatusType.staTest, p_Name);
						#If DEBUG Then
						#Else
						for (l_SlewAngle = 0; l_SlewAngle <= 45; l_SlewAngle += 45) {
							#End If
							for (l_SlewAngle = 0; l_SlewAngle <= 90; l_SlewAngle += 15) {
								DomeSlewToAltitude(p_Name, l_SlewAngle);
								if (TestStop())
									return;
							}
							//Test out of range values -10 and 100 degrees
							//Can set altitude so check out of range values
							if (m_CanSetAltitude) {
								try {
									DomeSlewToAltitude(p_Name, DOME_ILLEGAL_ALTITUDE_LOW);
									LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal altitude " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees");
								} catch (COMException ex) {
									//COM Invalid value exception
									if (ex.ErrorCode == ErrorCodes.InvalidValue | ex.ErrorCode == g_ExInvalidValue1 | ex.ErrorCode == g_ExInvalidValue2) {
										LogMsg(p_Name, MessageLevel.msgOK, "COM invalid value exception correctly raised for slew to " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees");
									} else {
										LogMsg(p_Name, MessageLevel.msgError, "Unexpected exception: " + ex.ToString);
									}
								} catch (Exception ex) {
									LogMsg(p_Name, MessageLevel.msgOK, "Exception correctly raised for slew to " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees");
								}
								try {
									DomeSlewToAltitude(p_Name, DOME_ILLEGAL_ALTITUDE_HIGH);
									LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal altitude " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees");
								} catch (COMException ex) {
									//COM Invalid value exception
									if (ex.ErrorCode == ErrorCodes.InvalidValue | ex.ErrorCode == g_ExInvalidValue1 | ex.ErrorCode == g_ExInvalidValue2) {
										LogMsg(p_Name, MessageLevel.msgOK, "COM invalid value exception correctly raised for slew to " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees");
									} else {
										LogMsg(p_Name, MessageLevel.msgError, "Unexpected exception: " + ex.ToString);
									}
								} catch (Exception ex) {
									LogMsg(p_Name, MessageLevel.msgOK, "Exception correctly raised for slew to " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees");
								}
							}
						}
						//SlewToAltitude should raise an exception
						m_Dome.SlewToAltitude(45.0);
						LogMsg(p_Name, MessageLevel.msgError, "CanSetAltitude is false but SlewToAltitude did not raise an exception");

					}
				case DomePropertyMethod.SlewToAzimuth:
					//Can set azimuth so SlewToAzimuth sould succeed
					if (m_CanSetAzimuth) {
						Status(StatusType.staTest, p_Name);
						#If DEBUG Then
						#Else
						for (l_SlewAngle = 0; l_SlewAngle <= 30; l_SlewAngle += 30) {
							#End If
							for (l_SlewAngle = 0; l_SlewAngle <= 315; l_SlewAngle += 45) {
								DomeSlewToAzimuth(p_Name, l_SlewAngle);
								if (TestStop())
									return;
							}
							//Test out of range values -10 and 370 degrees
							try {
								DomeSlewToAzimuth(p_Name, DOME_ILLEGAL_AZIMUTH_LOW);
								LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
							} catch (COMException ex) {
								//COM Invalid value exception
								if (ex.ErrorCode == ErrorCodes.InvalidValue | ex.ErrorCode == g_ExInvalidValue1 | ex.ErrorCode == g_ExInvalidValue2) {
									LogMsg(p_Name, MessageLevel.msgOK, "COM invalid value exception correctly raised for slew to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
								} else {
									LogMsg(p_Name, MessageLevel.msgError, "Unexpected exception: " + ex.ToString);
								}
							} catch (Exception ex) {
								LogMsg(p_Name, MessageLevel.msgOK, "Exception correctly raised for slew to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
							}
							if (TestStop())
								return;
							try {
								DomeSlewToAzimuth(p_Name, DOME_ILLEGAL_AZIMUTH_HIGH);
								LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
							} catch (COMException ex) {
								//COM Invalid value exception
								if (ex.ErrorCode == ErrorCodes.InvalidValue | ex.ErrorCode == g_ExInvalidValue1 | ex.ErrorCode == g_ExInvalidValue2) {
									LogMsg(p_Name, MessageLevel.msgOK, "COM invalid value exception correctly raised for slew to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
								} else {
									LogMsg(p_Name, MessageLevel.msgError, "Unexpected exception: " + ex.ToString);
								}
							} catch (Exception ex) {
								LogMsg(p_Name, MessageLevel.msgOK, "Exception correctly raised for slew to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
							}
							if (TestStop())
								return;
						}
						//SlewToAzimuth should throw an exception
						m_Dome.SlewToAzimuth(45.0);
						LogMsg(p_Name, MessageLevel.msgError, "CanSetAzimuth is false but SlewToAzimuth did not throw an exception");
					}
				case DomePropertyMethod.SyncToAzimuth:
					//Can sync azimuth so test the command
					if (m_CanSyncAzimuth) {
						//SlewToAzimuth command appears to work so use it
						if (m_CanSlewToAzimuth) {
							//Can also read azimuth so test sync to a new azimuth
							if (m_CanReadAzimuth) {
								l_OriginalAzimuth = m_Dome.Azimuth;
								// Choose a smaller azimuth
								if (l_OriginalAzimuth > 300.0) {
									l_NewAzimuth = l_OriginalAzimuth - DOME_SYNC_OFFSET;
								//Choose a bigger azimuth
								} else {
									l_NewAzimuth = l_OriginalAzimuth + DOME_SYNC_OFFSET;
								}
								m_Dome.SyncToAzimuth(l_NewAzimuth);
								//Sync to new azimuth
								// OK Dome hasn't moved but should now show azimuth as a new value
								switch (Math.Abs(l_NewAzimuth - m_Dome.Azimuth)) {
									case  // ERROR: Case labels with binary operators are unsupported : LessThan
1.0:
										//very close so give it an OK
										LogMsg(p_Name, MessageLevel.msgOK, "Dome sync'd OK to within +- 1 degree");
									case  // ERROR: Case labels with binary operators are unsupported : LessThan
2.0:
										//close so give it an INFO
										LogMsg(p_Name, MessageLevel.msgInfo, "Dome sync'd to within +- 2 degrees");
									case  // ERROR: Case labels with binary operators are unsupported : LessThan
5.0:
										//Closeish so give an issue
										LogMsg(p_Name, MessageLevel.msgIssue, "Dome only sync'd to within +- 5 degrees");
									case (DOME_SYNC_OFFSET - 2.0): // TODO: to (DOME_SYNC_OFFSET + 2)
										//Hasn't really moved
										LogMsg(p_Name, MessageLevel.msgError, "Dome did not sync, Azimuth didn't change value after sync command");
									default:
										//Something else!
										LogMsg(p_Name, MessageLevel.msgIssue, "Dome azimuth was " + Math.Abs(l_NewAzimuth - m_Dome.Azimuth) + " degrees away from expected value");
								}
								//Now try and restore original value
								m_Dome.SyncToAzimuth(l_OriginalAzimuth);
							//Can't read azimuth so can only test that command completes
							} else {
								m_Dome.SyncToAzimuth(45.0);
								//Sync to an arbitary direction
								LogMsg(p_Name, MessageLevel.msgOK, "Dome successfully sync'd to 45 degrees but unable to read azimuth to confirm this");
							}

							//Now test sync to illegal values
							try {
								m_Dome.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_LOW);
								LogMsg(p_Name, MessageLevel.msgError, "No exception generated when syncing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
							} catch (COMException ex) {
								//COM Invalid value exception
								if (ex.ErrorCode == ErrorCodes.InvalidValue | ex.ErrorCode == g_ExInvalidValue1 | ex.ErrorCode == g_ExInvalidValue2) {
									LogMsg(p_Name, MessageLevel.msgOK, "COM invalid value exception correctly raised for sync to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
								} else {
									LogMsg(p_Name, MessageLevel.msgError, "Unexpected exception: " + ex.ToString);
								}
							} catch (Exception ex) {
								LogMsg(p_Name, MessageLevel.msgOK, "Exception correctly raised for sync to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
							}
							if (TestStop())
								return;
							try {
								m_Dome.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_HIGH);
								LogMsg(p_Name, MessageLevel.msgError, "No exception generated when syncing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
							} catch (COMException ex) {
								//COM Invalid value exception
								if (ex.ErrorCode == ErrorCodes.InvalidValue | ex.ErrorCode == g_ExInvalidValue1 | ex.ErrorCode == g_ExInvalidValue2) {
									LogMsg(p_Name, MessageLevel.msgOK, "COM invalid value exception correctly raised for sync to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
								} else {
									LogMsg(p_Name, MessageLevel.msgError, "Unexpected exception: " + ex.ToString);
								}
							} catch (Exception ex) {
								LogMsg(p_Name, MessageLevel.msgOK, "Exception correctly raised for sync to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
							}
							if (TestStop())
								return;
						//SlewoAzimuth doesn't work so skip test
						} else {
							LogMsg(p_Name, MessageLevel.msgInfo, "SyncToAzimuth test skipped since SlewToAzimuth throws an exception");
						}
					//Can not sync azimuth so SyncAzimuth should raise an exception
					} else {
						m_Dome.SyncToAzimuth(45.0);
						LogMsg(p_Name, MessageLevel.msgError, "CanSyncAzimuth is false but SyncToAzimuth did not raise an exception");
					}
				default:
					LogMsg(p_Name, MessageLevel.msgError, "DomeOptionalTest: Unknown test type " + p_Type.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (MethodNotImplementedException ex) {
			LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
		} catch (PropertyNotImplementedException ex) {
			LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
		Status(StatusType.staTest, "");
		Status(StatusType.staAction, "");
		Status(StatusType.staStatus, "");
	}
	private void DomeShutterTest(ShutterState p_RequiredShutterState, string p_Name)
	{
		MsgBoxResult l_MsgBoxResult;
		ShutterState l_ShutterState;
		//Shutter tests are allowed
		if (frmConformMain.chkDomeShutter.Checked) {
			Status(StatusType.staTest, p_Name);
			//Can read shutter status so use it
			if (m_CanReadShutterStatus) {
				l_ShutterState = m_Dome.ShutterStatus;
				l_ShutterState = (ShutterState)l_ShutterState;
				switch (l_ShutterState) {
					//Make sure we are in the required state to start the test
					case ShutterState.shutterClosed:
						//Testing ShutterClose
						if (p_RequiredShutterState == ShutterState.shutterClosed) {
							//Wrong state, get to the required state
							Status(StatusType.staAction, "Opening shutter ready for close test");
							LogMsg(p_Name, MessageLevel.msgDebug, "Opening shutter ready for close test");
							m_Dome.OpenShutter();
							if (!DomeShutterWait(ShutterState.shutterOpen))
								return;
							//Wait for shutter to open
							DomeStabliisationWait();
						//Testing ShutterOpen
						} else {
							//No action this is the desired start state
						}
					case ShutterState.shutterClosing:
						//Testing ShutterClose
						if (p_RequiredShutterState == ShutterState.shutterClosed) {
							Status(StatusType.staAction, "Waiting for shutter to close before opening ready for close test");
							LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close before opening ready for close test");
							if (!DomeShutterWait(ShutterState.shutterClosed))
								return;
							//Wait for shutter to close
							LogMsg(p_Name, MessageLevel.msgDebug, "Opening shutter ready for close test");
							Status(StatusType.staAction, "Opening shutter ready for close test");
							m_Dome.OpenShutter();
							//Then open it
							if (!DomeShutterWait(ShutterState.shutterOpen))
								return;
							DomeStabliisationWait();
						//Testing ShutterOpen
						} else {
							Status(StatusType.staAction, "Waiting for shutter to close ready for open test");
							LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close ready for open test");
							if (!DomeShutterWait(ShutterState.shutterClosed))
								return;
							//Wait for shutter to close
							DomeStabliisationWait();
						}
					case ShutterState.shutterOpening:
						//Testing ShutterClose
						if (p_RequiredShutterState == ShutterState.shutterClosed) {
							Status(StatusType.staAction, "Waiting for shutter to open ready for close test");
							LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open ready for close test");
							if (!DomeShutterWait(ShutterState.shutterOpen))
								return;
							//Wait for shutter to open
							DomeStabliisationWait();
						//Testing for ShutterOpen
						} else {
							Status(StatusType.staAction, "Waiting for shutter to open before closing ready for open test");
							LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open before closing ready for open test");
							if (!DomeShutterWait(ShutterState.shutterOpen))
								return;
							//Wait for shutter to open
							LogMsg(p_Name, MessageLevel.msgDebug, "Closing shutter ready for open test");
							Status(StatusType.staAction, "Closing shutter ready for open test");
							m_Dome.CloseShutter();
							//Then close it
							if (!DomeShutterWait(ShutterState.shutterClosed))
								return;
							DomeStabliisationWait();
						}
					case ShutterState.shutterError:
						l_MsgBoxResult = MsgBox("ShutterState is \"Error\" please prepare the shutter for the " + p_Name + " test and press OK or press Cancel to stop tests", MsgBoxStyle.OkCancel | MsgBoxStyle.SystemModal, "Dome Safety");
						//Cancel was pressed
						if (l_MsgBoxResult == MsgBoxResult.Cancel) {
							g_Stop = true;
							//Stop if cancel was pressed
							LogMsg("DomeSafety", MessageLevel.msgComment, "Cancel pressed, further testing abandoned");
						} else {
							LogMsg("DomeSafety", MessageLevel.msgComment, "OK pressed, testing continues");
						}
					case ShutterState.shutterOpen:
						//Testing ShutterClose
						if (p_RequiredShutterState == ShutterState.shutterClosed) {
						//No action this is the desired start state
						//Testing ShutterOpen
						} else {
							//Wrong state, get to the required state
							Status(StatusType.staAction, "Closing shutter ready for open  test");
							LogMsg(p_Name, MessageLevel.msgDebug, "Closing shutter ready for open test");
							m_Dome.CloseShutter();
							if (!DomeShutterWait(ShutterState.shutterClosed))
								return;
							//Wait for shutter to open
							DomeStabliisationWait();
						}
					default:
						LogMsg("DomeShutterTest", MessageLevel.msgError, "Unexpected shutter status: " + l_ShutterState.ToString);
				}
				//Testing ShutterClose
				if (p_RequiredShutterState == ShutterState.shutterClosed) {
					//Shutter is now open so close it
					Status(StatusType.staAction, "Closing shutter");
					m_Dome.CloseShutter();
					Status(StatusType.staAction, "Waiting for shutter to close");
					LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close");
					if (!DomeShutterWait(ShutterState.shutterClosed)) {
						l_ShutterState = m_Dome.ShutterStatus;
						l_ShutterState = (ShutterState)l_ShutterState;
						LogMsg(p_Name, MessageLevel.msgError, "Unable to close shutter - ShutterStatus: " + m_Dome.ShutterStatus.ToString);
						return;
					} else {
						LogMsg(p_Name, MessageLevel.msgOK, "Shutter closed successfully");
					}
					DomeStabliisationWait();
				} else {
					//Shutter is now closed so open it
					Status(StatusType.staAction, "Opening shutter");
					m_Dome.OpenShutter();
					Status(StatusType.staAction, "Waiting for shutter to open");
					LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open");
					if (!DomeShutterWait(ShutterState.shutterOpen)) {
						l_ShutterState = m_Dome.ShutterStatus;
						l_ShutterState = (ShutterState)l_ShutterState;
						LogMsg(p_Name, MessageLevel.msgError, "Unable to open shutter - ShutterStatus: " + m_Dome.ShutterStatus.ToString);
						return;
					} else {
						LogMsg(p_Name, MessageLevel.msgOK, "Shutter opened successfully");
					}
					DomeStabliisationWait();
				}

			//Can't read shutter status so don't try
			} else {
				LogMsg(p_Name, MessageLevel.msgDebug, "Can't read shutter status!");
				//Testing ShutterClose
				if (p_RequiredShutterState == ShutterState.shutterClosed) {
					//Just issue command to seee if it doesn't generate an error
					m_Dome.CloseShutter();
					DomeStabliisationWait();
				//Testing ShutterOpen
				} else {
					//Just issue command to seee if it doesn't generate an error
					m_Dome.OpenShutter();
					DomeStabliisationWait();
				}
				LogMsg(p_Name, MessageLevel.msgOK, "Command issued successfully but can't read ShutterStatus to confirm shutter is closed");
			}
			Status(StatusType.staTest, "");
			Status(StatusType.staAction, "");
			Status(StatusType.staStatus, "");
		} else {
			LogMsg("DomeSafety", MessageLevel.msgComment, "Open shutter check box is unchecked so shutter test bypassed");
		}
	}
	private bool DomeShutterWait(ShutterState p_RequiredStatus)
	{
		//Wait for shutter to reach required stats or user presses stop or timeout occurs
		//Returns true if required state is reached
		System.DateTime l_StartTime;
		ShutterState l_ShutterState;
		DomeShutterWait = false;
		l_StartTime = Now;
		try {
			do {
				WaitFor(SLEEP_TIME);
				Application.DoEvents();
				l_ShutterState = m_Dome.ShutterStatus;
				l_ShutterState = (ShutterState)l_ShutterState;
				Status(StatusType.staStatus, "Shutter State: " + l_ShutterState.ToString + " Timeout: " + Now.Subtract(l_StartTime).Seconds + "/" + g_Settings.DomeShutterTimeout);
			} while (!((l_ShutterState == p_RequiredStatus) | TestStop() | (Now.Subtract(l_StartTime).TotalSeconds > g_Settings.DomeShutterTimeout)));
			if ((m_Dome.ShutterStatus == p_RequiredStatus))
				DomeShutterWait = true;
			// All worked so return True
			// We timed out so give error message
			if ((Now.Subtract(l_StartTime).TotalSeconds > g_Settings.DomeShutterTimeout)) {
				LogMsg("DomeShutterWait", MessageLevel.msgError, "Timed out waiting for shutter to reach state: " + p_RequiredStatus.ToString + ", consider increasing the timeout setting in Options / Conformance Options");
			}
		} catch (Exception ex) {
			LogMsg("DomeShutterWait", MessageLevel.msgError, "Unexpected exception: " + ex.ToString);
		}
	}
	private void DomePerformanceTest(DomePropertyMethod p_Type, string p_Name)
	{
		System.DateTime l_StartTime;
		double l_Count;
		double l_LastElapsedTime;
		double l_ElapsedTime;
		double l_Rate;
		bool l_Boolean;
		double l_Double;
		ShutterState l_ShutterState;
		Status(StatusType.staTest, "Performance Testing");
		Status(StatusType.staAction, p_Name);
		try {
			l_StartTime = Now;
			l_Count = 0.0;
			l_LastElapsedTime = 0.0;
			do {
				l_Count += 1.0;
				switch (p_Type) {
					case DomePropertyMethod.Altitude:
						l_Double = m_Dome.Altitude;
					case DomePropertyMethod.Azimuth:
						l_Double = m_Dome.Azimuth;
					case DomePropertyMethod.ShutterStatus:
						l_ShutterState = m_Dome.ShutterStatus;
					case DomePropertyMethod.SlavedRead:
						l_Boolean = m_Dome.Slaved;
					case DomePropertyMethod.Slewing:
						l_Boolean = m_Dome.Slewing;
					default:
						LogMsg(p_Name, MessageLevel.msgError, "DomePerformanceTest: Unknown test type " + p_Type.ToString);
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

	private void DomeStabliisationWait()
	{
		Status(StatusType.staStatus, "");
		//Clear status field
		for (int i = 1; i <= g_Settings.DomeStabilisationWait; i++) {
			Status(StatusType.staAction, "Waiting for Dome to stabilise - " + i + "/" + g_Settings.DomeStabilisationWait + " seconds");
			WaitFor(1000);
			// Wait for 1 second
		}
	}
	#End Region

}
