using ASCOM.DeviceInterface;
using ASCOM.DriverAccess;
internal class CameraTester : DeviceTesterBaseClass
{
	#Region "Variables and Constants"
		//Duration of camera pulse guide test (ms)
	const int CAMERA_PULSE_DURATION = 2000;
		//Tolerance for acceptab;e performance (ms)
	const int CAMERA_PULSE_TOLERANCE = 300;

	//Camera variables
	private bool m_CanAbortExposure;
	private bool m_CanAsymmetricBin;
	private bool m_CanGetCoolerPower;
	private bool m_CanSetCCDTemperature;
	private bool m_CanStopExposure;
	private bool m_CanFastReadout;
	private bool m_CoolerOn;
	private bool m_HasShutter;
	private bool m_ImageReady;
	private int m_CameraXSize;
	private int m_CameraYSize;
	private int m_MaxADU;
	private int m_NumX;
	private int m_NumY;
	private int m_StartX;
	private int m_StartY;
	private short m_MaxBinX;
	private short m_MaxBinY;
	private short m_BinX;
	private short m_BinY;
	private double m_CCDTemperature;
	private double m_CoolerPower;
	private double m_ElectronsPerADU;
	private double m_FullWellCapacity;
	private double m_HeatSinkTemperature;
	private double m_LastExposureDuration;
	private double m_PixelSizeX;
	private double m_PixelSizeY;
	private double m_SetCCDTemperature;
	private string m_LastError;
	private string m_LastExposureStartTime;
	private ASCOM.DeviceInterface.CameraStates m_CameraState;
	private bool m_LastErrorStatus1;
	private bool m_LastErrorStatus2;
	private Array m_ImageArray;
	private Array m_ImageArrayVariant;
		//Confirm that IsPulseGuiding command will work
	private bool m_IsPulseGuidingSupported;
	private bool m_CanPulseGuide;
	private string m_Description;
	private bool m_IsPulseGuiding;
	//ICamera V2 properties
	private short m_BayerOffsetX;
	private short m_BayerOffsetY;
	private short m_Gain;
	private short m_GainMax;
	private short m_GainMin;
	private short m_PercentCompleted;
	private short m_ReadoutMode;
	private double m_ExposureMax;
	private double m_ExposureMin;
	private double m_ExposureResolution;
	private bool m_FastReadout;
	private bool m_CanReadGain;
	private bool m_CanReadGainMax;
	private bool m_CanReadGainMin;
	private bool m_CanReadGains;
	private bool m_CanReadReadoutModes;
	private ArrayList m_Gains;
	private ArrayList m_ReadoutModes;
	private string m_SensorName;
	private SensorType m_SensorType;

	private bool m_CanReadSensorType = false;
	#If DEBUG Then
		#Else
	private object m_Camera;
		#End If
	private ASCOM.DriverAccess.Camera m_Camera;
	#End Region

	#Region "Enums"
	private enum CanType
	{
		tstCanFindHome = 1,
		tstCanPark = 2,
		tstCanPulseGuide = 3,
		tstCanSetDeclinationRate = 4,
		tstCanSetGuideRates = 5,
		tstCanSetPark = 6,
		tstCanSetPierSide = 7,
		tstCanSetRightAscensionRate = 8,
		tstCanSetTracking = 9,
		tstCanSlew = 10,
		tstCanSlewAsync = 11,
		tstCanSlewAltAz = 12,
		tstCanSlewAltAzAsync = 13,
		tstCanSync = 14,
		tstCanSyncAltAz = 15,
		tstCanUnPark = 16,
		tstCanAbortExposure = 17,
		tstCanAsymmetricBin = 18,
		tstCanGetCoolerPower = 19,
		tstCanSetCCDTemperature = 20,
		tstCanStopExposure = 21,
		// ICameraV2 property
		tstCanFastReadout = 22
	}
	private enum CameraPerformance : int
	{
		CameraState,
		CCDTemperature,
		CoolerPower,
		HeatSinkTemperature,
		ImageReady,
		IsPulseGuiding,
		ImageArray,
		ImageArrayVariant
	}
	private enum CamPropertyType
	{
		// ICameraV1 Properties
		BinX = 1,
		BinY = 2,
		CameraState = 3,
		CameraXSize = 4,
		CameraYSize = 5,
		CCDTemperature = 6,
		CoolerOn = 7,
		CoolerPower = 8,
		Description = 9,
		DriverInfo = 10,
		ElectronsPerADU = 11,
		FullWellCapacity = 12,
		HasShutter = 13,
		HeatSinkTemperature = 14,
		ImageReady = 15,
		IsPulseGuiding = 16,
		MaxADU = 17,
		MaxBinX = 18,
		MaxBinY = 19,
		NumX = 20,
		NumY = 21,
		PixelSizeX = 22,
		PixelSizeY = 23,
		SetCCDTemperature = 24,
		StartX = 25,
		StartY = 26,

		// ICameraV2 Properties
		BayerOffsetX = 27,
		BayerOffsetY = 28,
		ExposureMax = 29,
		ExposureMin = 30,
		ExposureResolution = 31,
		FastReadout = 32,
		Gain = 33,
		GainMax = 34,
		GainMin = 35,
		Gains = 36,
		PercentCompleted = 37,
		ReadoutMode = 38,
		ReadoutModes = 39,
		SensorName = 40,
		SensorType = 41

	}

	#End Region

	#Region "New and Dispose"
	private CameraTester()
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
					m_Camera.Connected = false;
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_Camera);
				} catch {
				}
				m_Camera = null;
				GC.Collect();
			}

			// TODO: free your own state (unmanaged objects).
			// TODO: set large fields to null.
		}
		this.disposedValue = true;
	}
	#End Region

	#Region "Code"
	private override void CheckCommonMethods()
	{
		base.CheckCommonMethods(m_Camera, DeviceType.Camera);
	}

	private override void CheckInitialise()
	{
		//Set the error type numbers acording to the standards adopted by individual authors.
		//Unfortunatley these vary between drivers so I have to allow for these here in order to give meaningful
		//messages to driver authors!
		switch (g_CameraProgID.ToUpper) {
			default:
				//I'm using the simulator values as the defaults since it is the reference platform
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040405;
				g_ExInvalidValue2 = 0x80040405;
				g_ExNotSet1 = 0x80040403;
		}
		base.CheckInitialise(g_CameraProgID);

	}
	private override void CheckAccessibility()
	{
		object l_DeviceObject;
		object l_ICamera;
		ASCOM.DriverAccess.Camera l_DriverAccessCamera;
		base.CheckAccessibility(g_CameraProgID, DeviceType.Camera);

		//Try early binding to V1 interface
		l_ICamera = null;
		l_DeviceObject = null;
		try {
			l_DeviceObject = CreateObject(g_CameraProgID);
			l_ICamera = (ASCOM.Interface.ICamera)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface ICamera");
			try {
				l_ICamera.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ICamera");
				l_ICamera.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ICamera");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface ICamera");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_ICamera);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_ICamera = null;
			GC.Collect();
		}

		//Try early binding to V2 interface
		try {
			l_DeviceObject = CreateObject(g_CameraProgID);
			l_ICamera = (ASCOM.DeviceInterface.ICameraV2)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using interface ICameraV2");
			try {
				l_ICamera.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface ICameraV2");
				l_ICamera.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface ICameraV2");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface ICameraV2");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_ICamera);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_ICamera = null;
			GC.Collect();
		}

		//Try client access toolkit 
		l_DriverAccessCamera = null;
		try {
			l_DriverAccessCamera = new ASCOM.DriverAccess.Camera(g_CameraProgID);
			LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
			try {
				l_DriverAccessCamera.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
				l_DriverAccessCamera.Connected = false;
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
				l_DriverAccessCamera.Dispose();
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DriverAccessCamera);
			} catch {
			}
			l_DriverAccessCamera = null;
			GC.Collect();
		}

	}
	#If DEBUG Then
	private override void CreateDevice()
	{
		m_Camera = CreateObject(g_CameraProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Camera object");
		#Else
		m_Camera = new ASCOM.DriverAccess.Camera(g_CameraProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Camera to get a Camera object");
		#End If
		g_Stop = false;
	}
	private override bool Connected {
		get { Connected = m_Camera.Connected; }
		set {
			m_Camera.Connected = value;
			g_Stop = false;
		}
	}
	public override void ReadCanProperties()
	{
		// ICameraV1 properties
		CameraCanTest(CanType.tstCanAbortExposure, "CanAbortExposure");
		CameraCanTest(CanType.tstCanAsymmetricBin, "CanAsymmetricBin");
		CameraCanTest(CanType.tstCanGetCoolerPower, "CanGetCoolerPower");
		CameraCanTest(CanType.tstCanPulseGuide, "CanPulseGuide");
		CameraCanTest(CanType.tstCanSetCCDTemperature, "CanSetCCDTemperature");
		CameraCanTest(CanType.tstCanStopExposure, "CanStopExposure");

		// ICameraV2 properties
		CameraCanTest(CanType.tstCanFastReadout, "CanFastReadout");
	}
	private void CameraCanTest(CanType p_Type, string p_Name)
	{
		try {
			switch (p_Type) {
				case CanType.tstCanAbortExposure:
					m_CanAbortExposure = m_Camera.CanAbortExposure;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanAbortExposure.ToString);
				case CanType.tstCanAsymmetricBin:
					m_CanAsymmetricBin = m_Camera.CanAsymmetricBin;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanAsymmetricBin.ToString);
				case CanType.tstCanGetCoolerPower:
					m_CanGetCoolerPower = m_Camera.CanGetCoolerPower;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanGetCoolerPower.ToString);
				case CanType.tstCanPulseGuide:
					m_CanPulseGuide = m_Camera.CanPulseGuide;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanPulseGuide.ToString);
				case CanType.tstCanSetCCDTemperature:
					m_CanSetCCDTemperature = m_Camera.CanSetCCDTemperature;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSetCCDTemperature.ToString);
				case CanType.tstCanStopExposure:
					m_CanStopExposure = m_Camera.CanStopExposure;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanStopExposure.ToString);
				case CanType.tstCanFastReadout:
					m_CanFastReadout = m_Camera.CanFastReadout;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanFastReadout.ToString);
				default:
					LogMsg(p_Name, MessageLevel.msgError, "Conform:CanTest: Unknown test type " + p_Type.ToString);
			}
		} catch (COMException ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.Message);
		}
	}
	public override void PreRunCheck()
	{
		int l_VStringPtr;
		int l_V1;
		int l_V2;
		int l_V3;
		string l_VString;
		//Add a test for a back level version of the Camera simulator - just abandon this process if any errors occur
		if (g_CameraProgID.ToUpper == "CCDSIMULATOR.CAMERA") {
			l_VString = "";
			try {
				l_VStringPtr = InStr(m_Camera.Description.ToUpper, "VERSION ");
				//Point at the start of the version string
				//There is a version string so read what it is
				if (l_VStringPtr > 0) {
					l_VString = Mid(m_Camera.Description.ToUpper, l_VStringPtr + 8);
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
							l_V3 = (int)l_VString;
							//Extract the number
							//Turn the version parts into a whole number
							l_V1 = l_V1 * 1000000 + l_V2 * 1000 + l_V3;
							//We are using a version earlier that 5.0.6 so issue a message
							if (l_V1 < 5000008) {
								LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the camera simulator has known conformance issues, ***");
								LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site http://ascom-standards.org/Downloads/Index.htm ***");
								LogMsg("", MessageLevel.msgAlways, "");
							}
						}
					}
				//Must be an early version 5.0.1 or earlier
				} else {
					LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the camera simulator has known conformance issues, ***");
					LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site http://ascom-standards.org/Downloads/Index.htm ***");
					LogMsg("", MessageLevel.msgAlways, "");
				}
			} catch (Exception ex) {
				LogMsg("ConformanceCheck", MessageLevel.msgError, ex.ToString);
			}
		}

		//Run camera tests
		if (!TestStop()) {
			LogMsg("", MessageLevel.msgAlways, "");
			//Check LastError throws an exception
			LogMsg("Last Tests", MessageLevel.msgAlways, "");
			try {
				m_LastExposureDuration = m_Camera.LastExposureDuration;
				LogMsg("LastExposureDuration", MessageLevel.msgError, "LastExposureDuration did not generate an exception when called before an exposure was made");
			} catch (COMException ex) {
				LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration correctly generated a COM exception before an exposure was made");
			} catch (Exception ex) {
				LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration correctly generated a .NET exception before an exposure was made");
			}

			try {
				m_LastExposureStartTime = m_Camera.LastExposureStartTime;
				LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime did not generate an exception when called before an exposure was made");
			} catch (COMException ex) {
				LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime correctly generated a COM exception before an exposure was made");
			} catch (Exception ex) {
				LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime correctly generated a .NET exception before an exposure was made");
			}
		}
	}

	private override void CheckProperties()
	{
		int l_BinX;
		int l_BinY;
		//Basic read tests
		m_MaxBinX = (short)CameraPropertyTestInteger(CamPropertyType.MaxBinX, "MaxBinX", 1, 10);
		if (TestStop())
			return;
		m_MaxBinY = (short)CameraPropertyTestInteger(CamPropertyType.MaxBinY, "MaxBinY", 1, 10);
		if (TestStop())
			return;

		// Only symetric binning is supported so confirm MaxBinX and Y match
		if (!m_CanAsymmetricBin) {
			if (m_MaxBinX != m_MaxBinY)
				LogMsg("CanAsymmetricBin", MessageLevel.msgError, "CanAsymmetricBin is false but MaxBinX and MaxBinY are not equal!");
		}

		m_BinX = (short)CameraPropertyTestInteger(CamPropertyType.BinX, "BinX Read", 1, 1);
		if (TestStop())
			return;
		//Must default to 1 on startup
		m_BinY = (short)CameraPropertyTestInteger(CamPropertyType.BinY, "BinY Read", 1, 1);
		if (TestStop())
			return;
		//Must default to 1 on startup
		// Only symetric binning is supported so confirm MaxBinX and Y match
		if (!m_CanAsymmetricBin) {
			if (m_BinX != m_BinY)
				LogMsg("CanAsymmetricBin", MessageLevel.msgError, "CanAsymmetricBin is false but BinX and BinY are not equal!");
		}

		//Test writing low and high Bin values outside maximum range
		//Invalid low value
		try {
			m_Camera.BinX = 0;
			LogMsg("BinX Write", MessageLevel.msgError, "Invalid value 0 written but no exception generated");
		} catch (Exception ex) {
			LogMsg("BinX Write", MessageLevel.msgOK, "Exception correctly generated on setting BinX to 0");
		}
		//Invalid highvalue
		try {
			m_Camera.BinX = (short)m_MaxBinX + 1;
			LogMsg("BinX Write", MessageLevel.msgError, "Invalid value " + m_MaxBinX + 1 + " written but no exception generated");
		} catch (Exception ex) {
			LogMsg("BinX Write", MessageLevel.msgOK, "Exception correctly generated on setting BinX to " + m_MaxBinX + 1);
		}
		//Invalid low value
		try {
			m_Camera.BinY = 0;
			LogMsg("BinY Write", MessageLevel.msgError, "Invalid value 0 written but no exception generated");
		} catch (Exception ex) {
			LogMsg("BinY Write", MessageLevel.msgOK, "Exception correctly generated on setting BinY to 0");
		}
		//Invalid high value
		try {
			m_Camera.BinY = (short)m_MaxBinY + 1;
			LogMsg("BinY Write", MessageLevel.msgError, "Invalid value " + m_MaxBinY + 1 + " written but no exception generated");
		} catch (Exception ex) {
			LogMsg("BinY Write", MessageLevel.msgOK, "Exception correctly generated on setting BinY to " + m_MaxBinY + 1);
		}

		//Write BinX and BinY
		//Test all combinations
		if (m_CanAsymmetricBin) {
			for (l_BinY = 1; l_BinY <= m_MaxBinY; l_BinY++) {
				for (l_BinX = 1; l_BinX <= m_MaxBinX; l_BinX++) {
					try {
						m_Camera.BinY = (short)l_BinY;
						try {
							m_Camera.BinX = (short)l_BinX;
							LogMsg("BinXY Write", MessageLevel.msgOK, "Successfully set asymmetric xy binning: " + l_BinX + " x " + l_BinY);
						} catch (COMException ex) {
							switch (ex.ErrorCode) {
								case ErrorCodes.InvalidValue:
								case g_ExInvalidValue1:
								case g_ExInvalidValue2:
								case g_ExInvalidValue3:
								case g_ExInvalidValue4:
								case g_ExInvalidValue5:
								case g_ExInvalidValue6:
									LogMsg("BinX Write", MessageLevel.msgInfo, "COM InvalidValueException received when setting BinX to " + l_BinX);
								default:
									LogMsg("BinX Write", MessageLevel.msgError, EX_COM + "error setting BinX - " + ex.Message);
							}
						} catch (InvalidValueException ex) {
							LogMsg("BinX Write", MessageLevel.msgInfo, ".NET InvalidValueException received when setting BinX to " + l_BinX);
						} catch (Exception ex) {
							LogMsg("BinX Write", MessageLevel.msgError, EX_NET + "error setting BinX - " + ex.Message);
						}
					} catch (COMException ex) {
						switch (ex.ErrorCode) {
							case ErrorCodes.InvalidValue:
							case g_ExInvalidValue1:
							case g_ExInvalidValue2:
							case g_ExInvalidValue3:
							case g_ExInvalidValue4:
							case g_ExInvalidValue5:
							case g_ExInvalidValue6:
								LogMsg("BinY Write", MessageLevel.msgInfo, "COM InvalidValueException received when setting BinX to " + l_BinX);
							default:
								LogMsg("BinY Write", MessageLevel.msgError, EX_COM + "error setting BinX - " + ex.Message);
						}
					} catch (InvalidValueException ex) {
						LogMsg("BinY Write", MessageLevel.msgInfo, ".NET InvalidValueException received when setting BinY to " + l_BinY);
					} catch (Exception ex) {
						LogMsg("BinY Write", MessageLevel.msgError, EX_NET + "error setting BinY - " + ex.Message);
					}
				}
			}

		//Test same bin combinations only
		} else {
			for (l_BinX = 1; l_BinX <= m_MaxBinX; l_BinX++) {
				try {
					m_Camera.BinY = (short)l_BinX;
					try {
						m_Camera.BinX = (short)l_BinX;
						LogMsg("BinXY Write", MessageLevel.msgOK, "Successfully set symmetric xy binning: " + l_BinX + " x " + l_BinX);
					} catch (COMException ex) {
						switch (ex.ErrorCode) {
							case ErrorCodes.InvalidValue:
							case g_ExInvalidValue1:
							case g_ExInvalidValue2:
							case g_ExInvalidValue3:
							case g_ExInvalidValue4:
							case g_ExInvalidValue5:
							case g_ExInvalidValue6:
								LogMsg("BinX Write", MessageLevel.msgInfo, "COM InvalidValueException received when setting BinX to " + l_BinX);
							default:
								LogMsg("BinX Write", MessageLevel.msgError, EX_COM + "error setting BinX - " + ex.Message);
						}
					} catch (InvalidValueException ex) {
						LogMsg("BinX Write", MessageLevel.msgInfo, ".NET InvalidValueException received when setting BinX to " + l_BinX);
					} catch (Exception ex) {
						LogMsg("BinX Write", MessageLevel.msgError, EX_NET + "error setting BinX - " + ex.Message);
					}
				} catch (COMException ex) {
					switch (ex.ErrorCode) {
						case ErrorCodes.InvalidValue:
						case g_ExInvalidValue1:
						case g_ExInvalidValue2:
						case g_ExInvalidValue3:
						case g_ExInvalidValue4:
						case g_ExInvalidValue5:
						case g_ExInvalidValue6:
							LogMsg("BinY Write", MessageLevel.msgInfo, "COM InvalidValueException received when setting BinY to " + l_BinY);
						default:
							LogMsg("BinY Write", MessageLevel.msgError, EX_COM + "error setting BinY - " + ex.Message);
					}
				} catch (InvalidValueException ex) {
					LogMsg("BinY Write", MessageLevel.msgInfo, ".NET InvalidValueException received when setting BinY to " + l_BinY);
				} catch (Exception ex) {
					LogMsg("BinY Write", MessageLevel.msgError, EX_NET + "error setting BinY - " + ex.Message);
				}
			}

		}
		//Reset binning to 1x1 state
		try {
			m_Camera.BinX = 1;
		} catch (Exception ex) {
		}
		try {
			m_Camera.BinY = 1;
		} catch (Exception ex) {
		}
		m_CameraState = CameraPropertyTestCameraState(CamPropertyType.CameraState, "CameraState");
		if (TestStop())
			return;
		m_CameraXSize = CameraPropertyTestInteger(CamPropertyType.CameraXSize, "CameraXSize", 1, int.MaxValue);
		if (TestStop())
			return;
		m_CameraYSize = CameraPropertyTestInteger(CamPropertyType.CameraYSize, "CameraYSize", 1, int.MaxValue);
		if (TestStop())
			return;

		m_CCDTemperature = CameraPropertyTestDouble(CamPropertyType.CCDTemperature, "CCDTemperature", -273.0, 100.0, false);
		if (TestStop())
			return;
		m_CoolerOn = CameraPropertyTestBoolean(CamPropertyType.CoolerOn, "CoolerOn Read", false);
		if (TestStop())
			return;

		//Write CoolerOn
		bool l_OriginalCoolerState;
		string l_TargetCoolerState;
		try {
			l_OriginalCoolerState = m_Camera.CoolerOn;
			//Set up error message element in case needed
			if (l_OriginalCoolerState) {
				l_TargetCoolerState = "off";
			} else {
				l_TargetCoolerState = "on";
			}
			try {
				//Cooler on so switch off
				if (l_OriginalCoolerState) {
					m_Camera.CoolerOn = false;
				//Cooler off so switch on
				} else {
					m_Camera.CoolerOn = true;
				}
				LogMsg("CoolerOn Write", MessageLevel.msgOK, "Successfully changed CoolerOn state");
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("CoolerOn", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("CoolerOn", ex.Message, EX_COM + "turning Cooler " + l_TargetCoolerState + " - " + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("CoolerOn", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("CoolerOn", ex.Message, EX_NET + "turning Cooler " + l_TargetCoolerState + " - " + ex.Message);
			}
			//Restore Cooler state
			try {
				m_Camera.CoolerOn = l_OriginalCoolerState;
			} catch {
			}

		} catch (Exception ex) {
			//Skip tests as we can't even read the cooler state
			LogMsg("CoolerOn Write", MessageLevel.msgInfo, "Skipping test as unable to read cooler state");
		}

		m_CoolerPower = CameraPropertyTestDouble(CamPropertyType.CoolerPower, "CoolerPower", 0.0, 100.0, false);
		if (TestStop())
			return;
		//m_Description = CameraPropertyTestString(CamPropertyType.Description, "Description", 68) : If TestStop() Then Exit Sub
		m_ElectronsPerADU = CameraPropertyTestDouble(CamPropertyType.ElectronsPerADU, "ElectronsPerADU", 1E-05, double.PositiveInfinity, false);
		if (TestStop())
			return;
		m_FullWellCapacity = CameraPropertyTestDouble(CamPropertyType.FullWellCapacity, "FullWellCapacity", 0.0, double.PositiveInfinity, false);
		if (TestStop())
			return;
		m_HasShutter = CameraPropertyTestBoolean(CamPropertyType.HasShutter, "HasShutter", false);
		if (TestStop())
			return;
		m_HeatSinkTemperature = CameraPropertyTestDouble(CamPropertyType.HeatSinkTemperature, "HeatSinkTemperature", -273.0, 100, false);
		if (TestStop())
			return;

		m_ImageReady = CameraPropertyTestBoolean(CamPropertyType.ImageReady, "ImageReady", false);
		if (TestStop())
			return;
		if (m_ImageReady)
			LogMsg("ImageReady", MessageLevel.msgError, "Image is flagged as ready but no exposure has been started!");
		//ImageReady is incorrectly flagged so don't know what to expect!
		if (m_ImageReady) {
			try {
				m_ImageArray = (int[,])m_Camera.ImageArray;
				LogMsg("ImageArray", MessageLevel.msgError, "No image has been taken but ImageArray has not generated an exception");
			} catch (COMException ex) {
				LogMsg("ImageArray", MessageLevel.msgOK, "Exception correctly generated before an image has been taken");
			} catch (Exception ex) {
				LogMsg("ImageArray", MessageLevel.msgOK, "Exception correctly generated before an image has been taken");
			}
		//Imageready is false so should throw an exception
		} else {
			try {
				m_ImageArray = (int[,])m_Camera.ImageArray;
				LogMsg("ImageArray", MessageLevel.msgError, "ImageReady is false and no image has been taken but ImageArray has not generated an exception");
			} catch (COMException ex) {
				LogMsg("ImageArray", MessageLevel.msgOK, "Exception correctly generated when ImageReady is false");
			} catch (Exception ex) {
				LogMsg("ImageArray", MessageLevel.msgOK, "Exception correctly generated when ImageReady is false");
			}
		}

		//ImageReady is incorrectly flagged so don't know what to expect!
		if (m_ImageReady) {
			try {
				m_ImageArrayVariant = (int[,])m_Camera.ImageArrayVariant;
				LogMsg("ImageArrayVariant", MessageLevel.msgError, "No image has been taken but ImageArray has not generated an exception");
			} catch (COMException ex) {
				LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Exception correctly generated before an image has been taken");
			} catch (Exception ex) {
				LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Exception correctly generated before an image has been taken");
			}
		//Imageready is false so should throw an exception
		} else {
			try {
				m_ImageArrayVariant = (int[,])m_Camera.ImageArrayVariant;
				LogMsg("ImageArrayVariant", MessageLevel.msgError, "ImageReady is false and no image has been taken but ImageArray has not generated an exception");
			} catch (COMException ex) {
				LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Exception correctly generated when ImageReady is false");
			} catch (Exception ex) {
				LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Exception correctly generated when ImageReady is false");
			}
		}


		m_IsPulseGuiding = CameraPropertyTestBoolean(CamPropertyType.IsPulseGuiding, "IsPulseGuiding", false);
		if (TestStop())
			return;
		if (m_IsPulseGuiding)
			LogMsg("IsPulseGuiding", MessageLevel.msgError, "Camera is showing pulse guiding underway although no PulseGuide command has been issued!");

		m_MaxADU = CameraPropertyTestInteger(CamPropertyType.MaxADU, "MaxADU", 1, int.MaxValue);
		if (TestStop())
			return;

		m_NumX = CameraPropertyTestInteger(CamPropertyType.NumX, "NumX Read", 1, m_CameraXSize);
		if (TestStop())
			return;
		CameraPropertyWriteTest(CamPropertyType.NumX, "NumX", (int)m_CameraXSize / 2, 0, m_CameraXSize + 1);

		m_NumY = CameraPropertyTestInteger(CamPropertyType.NumY, "NumY Read", 1, m_CameraYSize);
		if (TestStop())
			return;
		CameraPropertyWriteTest(CamPropertyType.NumY, "NumY", (int)m_CameraYSize / 2, 0, m_CameraYSize + 1);

		m_PixelSizeX = CameraPropertyTestDouble(CamPropertyType.PixelSizeX, "PixelSizeX", 1.0, double.PositiveInfinity, false);
		if (TestStop())
			return;
		m_PixelSizeY = CameraPropertyTestDouble(CamPropertyType.PixelSizeY, "PixelSizeY", 1.0, double.PositiveInfinity, false);
		if (TestStop())
			return;

		m_SetCCDTemperature = CameraPropertyTestDouble(CamPropertyType.SetCCDTemperature, "SetCCDTemperature Read", -273.0, 100.0, false);
		if (TestStop())
			return;
		//Should work OK
		if (m_CanSetCCDTemperature) {
			try {
				m_Camera.SetCCDTemperature = 0;
				//Try an acceptable value
				LogMsg("SetCCDTemperature Write", MessageLevel.msgOK, "Successfully wrote 0");
				//Find lowest acceptable temperature
				bool l_OK;
				double l_SetPoint;

				//Find low setpoint at which an exception is generated, stop at -280 as this is unphysical
				l_OK = false;
				l_SetPoint = 0.0;
				do {
					l_SetPoint = l_SetPoint - 5;
					try {
						m_Camera.SetCCDTemperature = l_SetPoint;
					} catch (Exception ex) {
						l_OK = true;
						//Reached lower limit so exit loop
					}
				} while (!(l_OK | (l_SetPoint == -280.0)));
				// No lower limit so generate an issue
				if (l_SetPoint == -280.0) {
					LogMsg("SetCCDTemperature Write", MessageLevel.msgIssue, "Setpoint could be set below absolute zero!");
				//Have found a limit so log what it is
				} else {
					LogMsg("SetCCDTemperature Write", MessageLevel.msgInfo, "Setpoint lower limit (to nearest 5 degrees) found: " + l_SetPoint + 5.0);
				}

				//Find high setpoint at which an exception is generated, stop at 100 as this is a rather high value
				l_OK = false;
				l_SetPoint = 0.0;
				do {
					l_SetPoint = l_SetPoint + 5;
					try {
						m_Camera.SetCCDTemperature = l_SetPoint;
					} catch (Exception ex) {
						l_OK = true;
						//Reached lower limit so exit loop
					}
				} while (!(l_OK | (l_SetPoint == 100.0)));
				// No higher limit so generate an issue
				if (l_SetPoint == 100.0) {
					LogMsg("SetCCDTemperature Write", MessageLevel.msgIssue, "Setpoint could be set to 100 degrees!");
				//Have found a limit so log what it is
				} else {
					LogMsg("SetCCDTemperature Write", MessageLevel.msgInfo, "Setpoint high limit (to nearest 5 degrees) found: " + l_SetPoint - 5.0);
				}
			} catch (COMException ex) {
				LogMsg("SetCCDTemperature Write", MessageLevel.msgError, EX_COM + "exception generated when setting legal value 0");
			} catch (Exception ex) {
				LogMsg("SetCCDTemperature Write", MessageLevel.msgError, EX_NET + "exception generated when setting legal value 0");
			}
			//Restore original value
			try {
				m_Camera.SetCCDTemperature = m_SetCCDTemperature;
			} catch {
			}
		//CanSetCCDTemperature is false, should generate an exception
		} else {
			try {
				m_Camera.SetCCDTemperature = 0;
				LogMsg("SetCCDTemperature Write", MessageLevel.msgError, "CanSetCCDTemperature is false but no exception generated on write");
			} catch (COMException ex) {
				LogMsg("SetCCDTemperature Write", MessageLevel.msgOK, "CanSetCCDTemperature is false and exception correctly generated on write");
			} catch (Exception ex) {
				LogMsg("SetCCDTemperature Write", MessageLevel.msgOK, "CanSetCCDTemperature is false and exception correctly generated on write");
			}
		}

		m_StartX = CameraPropertyTestInteger(CamPropertyType.StartX, "StartX Read", 0, m_CameraXSize - 1);
		if (TestStop())
			return;
		CameraPropertyWriteTest(CamPropertyType.StartX, "StartX", (int)m_CameraXSize / 2, -1, m_CameraXSize);
		m_StartY = CameraPropertyTestInteger(CamPropertyType.StartY, "StartY Read", 0, m_CameraYSize - 1);
		if (TestStop())
			return;
		CameraPropertyWriteTest(CamPropertyType.StartY, "StartY", (int)m_CameraYSize / 2, -1, m_CameraYSize);

		// Test ICameraV2 Properties
		// Only for ICameraV2 and later
		if (m_Camera.InterfaceVersion > 1) {
			// SensorType - Mandatory
			// This must be tested before BayerOffset because BayerOffset is mandatory for colour and optional for monchrome cameras
			try {
				m_SensorType = m_Camera.SensorType;
				m_CanReadSensorType = true;
				// Set a flag to indicate that we have got a valid SensorType value
				//Successfully retrieved a value
				LogMsg("SensorType Read", MessageLevel.msgOK, m_SensorType.ToString);
			} catch (COMException ex) {
				switch (ex.ErrorCode) {
					case ErrorCodes.NotImplemented:
					case g_ExNotImplemented:
						LogMsg("SensorType Read", MessageLevel.msgError, "COM - This property must not throw a NotImplementedException, it must return a value");
					default:
						LogMsg("SensorType Read", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
						LogMsg("SensorType Read", MessageLevel.msgDebug, ex.ToString);
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("SensorType Read", MessageLevel.msgError, ".NET - This property must not throw a PropertyNotImplementedException, it must return a value");
			} catch (Exception ex) {
				LogMsg("SensorType Read", MessageLevel.msgError, EX_NET + ex.Message);
				LogMsg("SensorType Read", MessageLevel.msgDebug, ex.ToString);
			}

			// BayerOffset Read
			// SensorType value is available
			if (m_CanReadSensorType) {
				// Monchrome chip
				if (m_SensorType == SensorType.Monochrome) {
					// Monochrome so both BayerOffset properties should throw not implemented exceptions
					CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetX, "BayerOffsetX Read");
					CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetY, "BayerOffsetY Read");
				// Bayer colour chip
				} else {
					m_BayerOffsetX = CameraPropertyTestShort(CamPropertyType.BayerOffsetX, "BayerOffsetX Read", 0, 10000, true);
					m_BayerOffsetY = CameraPropertyTestShort(CamPropertyType.BayerOffsetY, "BayerOffsetY Read", 0, 10000, true);
				}
			// Unable to read SensorType so assume that the sensor is Monochrome
			} else {
				LogMsg("BayerOffset Read", MessageLevel.msgInfo, "Unable to read SensorType value; assuming that the sensor is Monochrome");
				// Monochrome so both BayerOffset properties should throw not implemented exceptions
				CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetX, "BayerOffsetX Read");
				CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetY, "BayerOffsetY Read");
			}

			// ExposureMin and ExpoosureMax Read
			m_ExposureMax = CameraPropertyTestDouble(CamPropertyType.ExposureMax, "ExposureMax Read", 0.0001, double.MaxValue, true);
			m_ExposureMin = CameraPropertyTestDouble(CamPropertyType.ExposureMin, "ExposureMin Read", 0.0, double.MaxValue, true);
			if (m_ExposureMin <= m_ExposureMax) {
				LogMsg("ExposureMin", MessageLevel.msgOK, "ExposureMin is less than or equal to ExposureMax");
			} else {
				LogMsg("ExposureMin", MessageLevel.msgError, "ExposureMin is greater than ExposureMax");
			}

			// ExposureResolution Read
			m_ExposureResolution = CameraPropertyTestDouble(CamPropertyType.ExposureResolution, "ExposureResolution Read", 0.0, double.MaxValue, true);
			if (m_ExposureResolution <= m_ExposureMax) {
				LogMsg("ExposureResolution", MessageLevel.msgOK, "ExposureResolution is less than or equal to ExposureMax");
			} else {
				LogMsg("ExposureResolution", MessageLevel.msgError, "ExposureResolution is greater than ExposureMax");
			}

			//FastReadout Read Optional
			// Should be implemented and not throw a not impemented exception
			if (m_CanFastReadout) {
				m_FastReadout = CameraPropertyTestBoolean(CamPropertyType.FastReadout, "FastReadout Read", true);
			// Should throw a not implememented exception
			} else {
				try {
					m_FastReadout = m_Camera.FastReadout;
					LogMsg("FastReadout Read", MessageLevel.msgError, "CanFastReadout is False but a PropertyNotImplementedException was not thrown");
				} catch (COMException ex) {
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("FastReadout Read", MessageLevel.msgOK, NOT_IMP_COM);
					} else {
						ExTest("FastReadout Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				} catch (PropertyNotImplementedException ex) {
					LogMsg("FastReadout Read", MessageLevel.msgOK, NOT_IMP_NET);
				} catch (Exception ex) {
					ExTest("FastReadout Read", ex.Message, EX_NET + ex.Message);
				}
			}

			// FastReadout Write Optional
			// Should be implemented and not throw a not impemented exception
			if (m_CanFastReadout) {
				try {
					m_Camera.FastReadout = !m_FastReadout;
					m_Camera.FastReadout = m_FastReadout;
					LogMsg("FastReadout Write", MessageLevel.msgOK, "Able to change the FastReadout state OK");
				} catch (COMException ex) {
					LogMsg("FastReadout Write", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					LogMsg("FastReadout Write", MessageLevel.msgDebug, ex.ToString);
				} catch (Exception ex) {
					LogMsg("FastReadout Write", MessageLevel.msgError, EX_NET + ex.Message);
					LogMsg("FastReadout Write", MessageLevel.msgDebug, ex.ToString);
				}

			// Should throw a not implememented exception
			} else {
				try {
					m_Camera.FastReadout = true;
					LogMsg("FastReadout Write", MessageLevel.msgError, "CanFastReadout is False but a PropertyNotImplementedException was not thrown");
				} catch (COMException ex) {
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("FastReadout Write", MessageLevel.msgOK, NOT_IMP_COM);
					} else {
						ExTest("FastReadout Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
						LogMsg("FastReadout Write", MessageLevel.msgDebug, ex.ToString);
					}
				} catch (PropertyNotImplementedException ex) {
					LogMsg("FastReadout Write", MessageLevel.msgOK, NOT_IMP_NET);
				} catch (Exception ex) {
					ExTest("FastReadout Write", ex.Message, EX_NET + ex.Message);
				}
			}

			// GainMin Read - Optional
			try {
				m_CanReadGainMin = false;
				m_GainMin = m_Camera.GainMin;
				//Successfully retrieved a value
				m_CanReadGainMin = true;
				LogMsg("GainMin Read", MessageLevel.msgOK, m_GainMin);
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("GainMin Read", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("GainMin Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (ASCOM.InvalidOperationException ex) {
				LogMsg("GainMin Read", MessageLevel.msgOK, "InvalidOperationxception correctly thrown");
			} catch (PropertyNotImplementedException ex) {
				LogMsg("GainMin Read", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("GainMin Read", ex.Message, EX_NET + ex.Message);
			}

			// GainMax Read - Optional
			try {
				m_CanReadGainMax = false;
				m_GainMax = m_Camera.GainMax;
				//Successfully retrieved a value
				m_CanReadGainMax = true;
				LogMsg("GainMax Read", MessageLevel.msgOK, m_GainMax);
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("GainMax Read", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("GainMax Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (ASCOM.InvalidOperationException ex) {
				LogMsg("GainMax Read", MessageLevel.msgOK, "InvalidOperationxception correctly thrown");
			} catch (PropertyNotImplementedException ex) {
				LogMsg("GainMax Read", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("GainMax Read", ex.Message, EX_NET + ex.Message);
			}

			if (m_CanReadGainMin ^ m_CanReadGainMax) {
				if (m_CanReadGainMin) {
					LogMsg("GainMinMax", MessageLevel.msgError, "Can read GainMin but GainMax threw an exception");
				} else {
					LogMsg("GainMinMax", MessageLevel.msgError, "Can read GainMax but GainMax threw an exception");
				}
			} else {
				LogMsg("GainMinMax", MessageLevel.msgOK, "Both GainMin and GainMax are readable or both throw exceptions");
			}

			// Gains Read - Optional
			try {
				m_CanReadGains = false;
				m_Gains = m_Camera.Gains;
				//Successfully retrieved a value
				m_CanReadGains = true;
				foreach (string Gain in m_Gains) {
					LogMsg("Gains Read", MessageLevel.msgOK, Gain.ToString);
				}
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("Gains Read", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("Gains Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					LogMsg("Gains Read", MessageLevel.msgDebug, ex.ToString);
				}
			} catch (ASCOM.InvalidOperationException ex) {
				LogMsg("Gains Read", MessageLevel.msgOK, "InvalidOperationxception correctly thrown");
			} catch (PropertyNotImplementedException ex) {
				LogMsg("Gains Read", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("Gains Read", ex.Message, EX_NET + ex.Message);
				LogMsg("Gains Read", MessageLevel.msgDebug, ex.ToString);
			}

			// Both GainMin/Max and Gains are enbaled but only one mechanic should be used
			if (m_CanReadGainMax & m_CanReadGainMin & m_CanReadGains) {
				LogMsg("Gains", MessageLevel.msgError, "GainMin, GainMax and Gains are all readable. Only one of GainMin/Max as a pair or Gains should be used, the other should throw a PropertyNotImplementedException");
			// Only one mechanic is actyive or no mechanic is active so no action
			} else {

			}

			// Gain Read - Optional 
			m_CanReadGain = false;
			// Set default value to indicate can't read gain
			try {
				m_Gain = m_Camera.Gain;
				m_CanReadGain = true;
				// Flag that we can read Gain OK
				if (m_CanReadGains) {
					LogMsg("Gain Read", MessageLevel.msgOK, m_Gain + " " + m_Gains(0).ToString);
				} else {
					LogMsg("Gain Read", MessageLevel.msgOK, m_Gain);
				}
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("Gain Read", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("Gain Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("Gain Read", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("Gain Read", ex.Message, EX_NET + ex.Message);
			}

			// PercentCompleted Read - Mandatory
			try {
				m_PercentCompleted = m_Camera.PercentCompleted;
				switch (m_PercentCompleted) {
					case  // ERROR: Case labels with binary operators are unsupported : LessThan
0:
						//Lower than minimum value
						LogMsg("PercentCompleted Read", MessageLevel.msgError, "Invalid value: " + m_PercentCompleted.ToString);
					case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
100:
						//Higher than maximum value
						LogMsg("PercentCompleted Read", MessageLevel.msgError, "Invalid value: " + m_PercentCompleted.ToString);
					default:
						//OK value
						LogMsg("PercentCompleted Read", MessageLevel.msgOK, m_PercentCompleted.ToString);
				}
			} catch (COMException ex) {
				switch (ex.ErrorCode) {
					case ErrorCodes.NotImplemented:
					case g_ExNotImplemented:
						LogMsg("PercentCompleted Read", MessageLevel.msgError, "COM - This property must not throw a NotImplementedException, it must return a value");
					case ErrorCodes.InvalidOperationException:
						LogMsg("PercentCompleted Read", MessageLevel.msgOK, "COM - Operation is invalid when camera is not imaging or downloading");
					default:
						LogMsg("PercentCompleted Read", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
						LogMsg("PercentCompleted Read", MessageLevel.msgDebug, ex.ToString);
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("PercentCompleted Read", MessageLevel.msgError, ".NET - This property must not throw a PropertyNotImplementedException, it must return a value");
			} catch (ASCOM.InvalidOperationException ex) {
				LogMsg("PercentCompleted Read", MessageLevel.msgOK, ".NET - Operation is invalid when camera is not imaging or downloading");
			} catch (Exception ex) {
				LogMsg("PercentCompleted Read", MessageLevel.msgError, EX_NET + ex.Message);
				LogMsg("PercentCompleted Read", MessageLevel.msgDebug, ex.ToString);
			}

			// ReadoutModes - Mandatory
			try {
				m_CanReadReadoutModes = false;
				m_ReadoutModes = m_Camera.ReadoutModes;
				//Successfully retrieved a value
				m_CanReadReadoutModes = true;
				foreach (string ReadoutMode in m_ReadoutModes) {
					LogMsg("ReadoutModes Read", MessageLevel.msgOK, ReadoutMode.ToString);
				}
			} catch (COMException ex) {
				switch (ex.ErrorCode) {
					case ErrorCodes.NotImplemented:
					case g_ExNotImplemented:
						LogMsg("ReadoutModes Read", MessageLevel.msgError, "COM - This property must not throw a NotImplementedException, it must return a value");
					default:
						LogMsg("ReadoutModes Read", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
						LogMsg("ReadoutModes Read", MessageLevel.msgDebug, ex.ToString);
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("ReadoutModes Read", MessageLevel.msgError, ".NET - This property must not throw a PropertyNotImplementedException, it must return a value");
			} catch (Exception ex) {
				LogMsg("ReadoutModes Read", MessageLevel.msgError, EX_NET + ex.Message);
				LogMsg("ReadoutModes Read", MessageLevel.msgDebug, ex.ToString);
			}

			// ReadoutMode Read - Mandatory
			m_ReadoutMode = CameraPropertyTestShort(CamPropertyType.ReadoutMode, "ReadoutMode Read", 0, short.MaxValue, true);
			try {
				if (m_CanReadReadoutModes) {
					if (m_ReadoutMode < m_ReadoutModes.Count) {
						LogMsg("ReadoutMode Index", MessageLevel.msgOK, "ReadReadoutMode is within the bounds of the ReadoutModes ArrayList");
						LogMsg("ReadoutMode Index", MessageLevel.msgInfo, "Current value: " + m_ReadoutModes(m_ReadoutMode).ToString);
					} else {
						LogMsg("ReadoutMode Index", MessageLevel.msgError, "ReadReadoutMode is outside the bounds of the ReadoutModes ArrayList");
					}
				} else {
					LogMsg("ReadoutMode Index", MessageLevel.msgInfo, "Skipping ReadReadoutMode index test because ReadoutModes is unavailable");
				}
			} catch (Exception ex) {
				LogMsg("ReadoutMode Index", MessageLevel.msgError, "Exception: " + ex.Message);
				LogMsg("ReadoutMode Read", MessageLevel.msgDebug, ex.ToString);
			}

			// SensorName
			m_SensorName = CameraPropertyTestString(CamPropertyType.SensorName, "SensorName Read", 250, true);

		}
	}
	private CameraStates CameraPropertyTestCameraState(CamPropertyType p_Type, string p_Name)
	{
		try {
			CameraPropertyTestCameraState = CameraStates.cameraIdle;
			switch (p_Type) {
				case CamPropertyType.CameraState:
					CameraPropertyTestCameraState = m_Camera.CameraState;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestCameraState: Unknown test type - " + p_Type.ToString);
			}
			LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestCameraState.ToString);
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
	}
	private short CameraPropertyTestShort(CamPropertyType p_Type, string p_Name, int p_Min, int p_Max, bool p_Mandatory)
	{
		try {
			CameraPropertyTestShort = 0;
			switch (p_Type) {
				case CamPropertyType.BayerOffsetX:
					CameraPropertyTestShort = m_Camera.BayerOffsetX;
				case CamPropertyType.BayerOffsetY:
					CameraPropertyTestShort = m_Camera.BayerOffsetY;
				case CamPropertyType.PercentCompleted:
					CameraPropertyTestShort = m_Camera.PercentCompleted;
				case CamPropertyType.ReadoutMode:
					CameraPropertyTestShort = m_Camera.ReadoutMode;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestShort: Unknown test type - " + p_Type.ToString);
			}
			//Successfully retrieved a value
			switch (CameraPropertyTestShort) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
p_Min:
					//Lower than minimum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + CameraPropertyTestShort.ToString);
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
p_Max:
					//Higher than maximum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + CameraPropertyTestShort.ToString);
				default:
					//OK value
					LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestShort.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				if (p_Mandatory) {
					LogMsg(p_Name, MessageLevel.msgError, "COM - This property must not throw a PropertyNotImplementedException, it must return a value");
				} else {
					LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
				}
			} else {
				ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			if (p_Mandatory) {
				LogMsg(p_Name, MessageLevel.msgError, ".NET - This property must not throw a PropertyNotImplementedException, it must return a value");
			} else {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
			}
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
	}
	private bool CameraPropertyMustNotImplemented(CamPropertyType p_Type, string p_Name)
	{
		short TestShort;
		// Dummy variable to hold value that should never be returned

		CameraPropertyMustNotImplemented = true;
		try {
			switch (p_Type) {
				case CamPropertyType.BayerOffsetX:
					TestShort = m_Camera.BayerOffsetX;
					CameraPropertyMustNotImplemented = false;
					// Property should throw an exception but did not so record that fact
					LogMsg(p_Name, MessageLevel.msgError, "Sensortype is Monochrome so this property must throw a PropertyNotImplementedException; it must not return a value");
				case CamPropertyType.BayerOffsetY:
					TestShort = m_Camera.BayerOffsetY;
					CameraPropertyMustNotImplemented = false;
					// Property should throw an exception but did not so record that fact
					LogMsg(p_Name, MessageLevel.msgError, "Sensortype is Monochrome so this property must throw a PropertyNotImplementedException; it must not return a value");
				default:
					LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyMustNotImplemented: Unknown test type - " + p_Type.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
		return CameraPropertyMustNotImplemented;
		// Return success indicator, True means property did thrown the exception, False means that it did not
	}
	private int CameraPropertyTestInteger(CamPropertyType p_Type, string p_Name, int p_Min, int p_Max)
	{
		try {
			CameraPropertyTestInteger = 0;
			switch (p_Type) {
				case CamPropertyType.BinX:
					CameraPropertyTestInteger = m_Camera.BinX;
				case CamPropertyType.BinY:
					CameraPropertyTestInteger = m_Camera.BinY;
				case CamPropertyType.CameraState:
					CameraPropertyTestInteger = m_Camera.CameraState;
				case CamPropertyType.CameraXSize:
					CameraPropertyTestInteger = m_Camera.CameraXSize;
				case CamPropertyType.CameraYSize:
					CameraPropertyTestInteger = m_Camera.CameraYSize;
				case CamPropertyType.CameraYSize:
					CameraPropertyTestInteger = m_Camera.CameraYSize;
				case CamPropertyType.MaxADU:
					CameraPropertyTestInteger = m_Camera.MaxADU;
				case CamPropertyType.MaxBinX:
					CameraPropertyTestInteger = m_Camera.MaxBinX;
				case CamPropertyType.MaxBinY:
					CameraPropertyTestInteger = m_Camera.MaxBinY;
				case CamPropertyType.NumX:
					CameraPropertyTestInteger = m_Camera.NumX;
				case CamPropertyType.NumY:
					CameraPropertyTestInteger = m_Camera.NumY;
				case CamPropertyType.StartX:
					CameraPropertyTestInteger = m_Camera.StartX;
				case CamPropertyType.StartY:
					CameraPropertyTestInteger = m_Camera.StartY;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestInteger: Unknown test type - " + p_Type.ToString);
			}
			//Successfully retrieved a value
			switch (CameraPropertyTestInteger) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
p_Min:
					//Lower than minimum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + CameraPropertyTestInteger.ToString);
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
p_Max:
					//Higher than maximum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + CameraPropertyTestInteger.ToString);
				default:
					//OK value
					LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestInteger.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
	}
	private double CameraPropertyTestDouble(CamPropertyType p_Type, string p_Name, double p_Min, double p_Max, bool p_Mandatory)
	{
		try {
			CameraPropertyTestDouble = 0.0;
			switch (p_Type) {
				case CamPropertyType.CCDTemperature:
					CameraPropertyTestDouble = m_Camera.CCDTemperature;
				case CamPropertyType.CoolerPower:
					CameraPropertyTestDouble = m_Camera.CoolerPower;
				case CamPropertyType.ElectronsPerADU:
					CameraPropertyTestDouble = m_Camera.ElectronsPerADU;
				case CamPropertyType.FullWellCapacity:
					CameraPropertyTestDouble = m_Camera.FullWellCapacity;
				case CamPropertyType.HeatSinkTemperature:
					CameraPropertyTestDouble = m_Camera.HeatSinkTemperature;
				case CamPropertyType.PixelSizeX:
					CameraPropertyTestDouble = m_Camera.PixelSizeX;
				case CamPropertyType.PixelSizeY:
					CameraPropertyTestDouble = m_Camera.PixelSizeY;
				case CamPropertyType.SetCCDTemperature:
					CameraPropertyTestDouble = m_Camera.SetCCDTemperature;
				case CamPropertyType.ExposureMax:
					CameraPropertyTestDouble = m_Camera.ExposureMax;
				case CamPropertyType.ExposureMin:
					CameraPropertyTestDouble = m_Camera.ExposureMin;
				case CamPropertyType.ExposureResolution:
					CameraPropertyTestDouble = m_Camera.ExposureResolution;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestDouble: Unknown test type - " + p_Type.ToString);
			}
			//Successfully retrieved a value
			switch (CameraPropertyTestDouble) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
p_Min:
					//Lower than minimum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + CameraPropertyTestDouble.ToString);
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
p_Max:
					//Higher than maximum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + CameraPropertyTestDouble.ToString);
				default:
					//OK value
					LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestDouble.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				if (p_Mandatory) {
					LogMsg(p_Name, MessageLevel.msgError, "COM - This property must not throw a PropertyNotImplementedException, it must return a value");
				} else {
					LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
				}
			} else {
				ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			if (p_Mandatory) {
				LogMsg(p_Name, MessageLevel.msgError, ".NET - This property must not throw a PropertyNotImplementedException, it must return a value");
			} else {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
			}
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
	}
	private bool CameraPropertyTestBoolean(CamPropertyType p_Type, string p_Name, bool p_Mandatory)
	{
		try {
			CameraPropertyTestBoolean = false;
			switch (p_Type) {
				case CamPropertyType.CoolerOn:
					CameraPropertyTestBoolean = m_Camera.CoolerOn;
				case CamPropertyType.HasShutter:
					CameraPropertyTestBoolean = m_Camera.HasShutter;
				case CamPropertyType.ImageReady:
					CameraPropertyTestBoolean = m_Camera.ImageReady;
				case CamPropertyType.IsPulseGuiding:
					m_IsPulseGuidingSupported = false;
					CameraPropertyTestBoolean = m_Camera.IsPulseGuiding;
					//Command works properly and doesn't cause a not implemented exception
					m_IsPulseGuidingSupported = true;
				case CamPropertyType.FastReadout:
					CameraPropertyTestBoolean = m_Camera.FastReadout;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestBoolean: Unknown test type - " + p_Type.ToString);
			}
			//Successfully retrieved a value
			LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestBoolean.ToString);
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				if (p_Mandatory) {
					LogMsg(p_Name, MessageLevel.msgError, "COM - This property must not throw a PropertyNotImplementedException, it must return a value");
				} else {
					LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
				}
			} else {
				ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			if (p_Mandatory) {
				LogMsg(p_Name, MessageLevel.msgError, ".NET - This property must not throw a PropertyNotImplementedException, it must return a value");
			} else {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
			}
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
	}
	private string CameraPropertyTestString(CamPropertyType p_Type, string p_Name, int p_MaxLength, bool p_Mandatory)
	{
		CameraPropertyTestString = "";
		try {
			switch (p_Type) {
				case CamPropertyType.Description:
					CameraPropertyTestString = m_Camera.Description;
				case CamPropertyType.SensorName:
					CameraPropertyTestString = m_Camera.SensorName;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "CameraPropertyTestString: Unknown test type - " + p_Type.ToString);
			}
			//Successfully retrieved a value
			switch (CameraPropertyTestString) {
				case  // ERROR: Case labels with binary operators are unsupported : Equality
"":
					LogMsg(p_Name, MessageLevel.msgOK, "The driver returned an empty string");
				default:
					if (Len(CameraPropertyTestString) <= p_MaxLength) {
						LogMsg(p_Name, MessageLevel.msgOK, CameraPropertyTestString);
					} else {
						LogMsg(p_Name, MessageLevel.msgError, "String exceeds " + p_MaxLength + " characters maximum length - " + CameraPropertyTestString);
					}
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				if (p_Mandatory) {
					LogMsg(p_Name, MessageLevel.msgError, "COM - This property must not throw a PropertyNotImplementedException, it must return a value");
				} else {
					LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
				}
			} else {
				ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			if (p_Mandatory) {
				LogMsg(p_Name, MessageLevel.msgError, ".NET - This property must not throw a PropertyNotImplementedException, it must return a value");
			} else {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
			}
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
	}

	private void CameraPropertyWriteTest(CamPropertyType p_Type, string p_Property, int p_TestOK, int p_TestLow, int p_TestHigh)
	{
		//OK value first
		try {
			m_LastErrorStatus2 = false;
			switch (p_Type) {
				case CamPropertyType.NumX:
					m_Camera.NumX = p_TestOK;
				case CamPropertyType.NumY:
					m_Camera.NumY = p_TestOK;
				case CamPropertyType.StartX:
					m_Camera.StartX = p_TestOK;
				case CamPropertyType.StartY:
					m_Camera.StartY = p_TestOK;
			}
			LogMsg(p_Property + " write", MessageLevel.msgOK, "Successfully wrote " + p_TestOK);
		} catch (COMException ex) {
			m_LastErrorStatus2 = true;
			LogMsg(p_Property + " write", MessageLevel.msgError, EX_COM + "exception generated when setting legal value: " + p_TestOK.ToString + " - " + ex.Message);
		} catch (Exception ex) {
			m_LastErrorStatus2 = true;
			LogMsg(p_Property + " write", MessageLevel.msgError, EX_NET + "exception generated when setting legal value: " + p_TestOK.ToString + " - " + ex.Message);
		}
		//Try 'Illegal low value
		// Select Case p_Type
		//     Case CamPropertyType.NumX
		// m_Camera.NumX = p_TestLow
		//     Case CamPropertyType.NumY
		// m_Camera.NumY = p_TestLow
		//     Case CamPropertyType.StartX
		// m_Camera.StartX = p_TestLow
		//     Case CamPropertyType.StartY
		// m_Camera.StartY = p_TestLow
		// End Select
		//LogMsg(p_Property & " write", MessageLevel.msgError, "No exception generated for illegal low value: " & p_TestLow.ToString)
		//Catch ex As COMException
		// LogMsg(p_Property & " write", MessageLevel.msgOK, "Exception correctly generated for illegal low value: " & p_TestLow.ToString)
		// Catch ex As Exception
		// LogMsg(p_Property & " write", MessageLevel.msgOK, "Exception correctly generated for illegal low value: " & p_TestLow.ToString)
		// End Try
		//Try 'Illegal highvalue
		// Select Case p_Type
		//     Case CamPropertyType.NumX
		// m_Camera.NumX = p_TestHigh
		//     Case CamPropertyType.NumY
		// m_Camera.NumY = p_TestHigh
		//     Case CamPropertyType.StartX
		// m_Camera.StartX = p_TestHigh
		//     Case CamPropertyType.StartY
		// m_Camera.StartY = p_TestHigh
		// End Select
		// LogMsg(p_Property & " write", MessageLevel.msgError, "No exception generated for illegal high value: " & p_TestHigh.ToString)
		// m_LastErrorStatus2 = False
		// Catch ex As COMException
		// LogMsg(p_Property & " write", MessageLevel.msgOK, "Exception correctly generated for illegal high value: " & p_TestHigh.ToString)
		// m_LastErrorStatus2 = True
		// Catch ex As Exception
		// LogMsg(p_Property & " write", MessageLevel.msgOK, "Exception correctly generated for illegal high value: " & p_TestHigh.ToString)
		// m_LastErrorStatus2 = True
		// End Try
		//CameraCheckLastError(p_Property & " write", m_LastErrorStatus2)
	}

	private override void CheckMethods()
	{
		int l_i;
		int l_j;
		//AbortExposure - Mandatory
		try {
			m_CameraState = m_Camera.CameraState;
			switch (m_CameraState) {
				case CameraStates.cameraIdle:
					//All is OK so test a call to AbortExposure
					try {
						m_Camera.AbortExposure();
						//Can abort and no exception so OK
						if (m_CanAbortExposure) {
							LogMsg("AbortExposure", MessageLevel.msgOK, "No exception generated when camera is already idle");
						//Can not abort and no exception - so violates specification
						} else {
							LogMsg("AbortExposure", MessageLevel.msgError, "CanAbortExposure is false but no exception is generated when AbortExposure is called");
						}
					} catch (COMException ex) {
						//Can abort so should not have generated an error
						if (m_CanAbortExposure) {
							LogMsg("AbortExposure", MessageLevel.msgError, EX_COM + "exception incorrectly generated when camera is idle");
						//Can not abort so should have generated an error
						} else {
							LogMsg("AbortExposure", MessageLevel.msgOK, "CanAbortExposure is false and COM exception correctly generated");
						}
					} catch (Exception ex) {
						//Can abort so should not have generated an error
						if (m_CanAbortExposure) {
							LogMsg("AbortExposure", MessageLevel.msgError, EX_NET + "exception incorrectly generated when camera is idle");
						//Can not abort so should have generated an error
						} else {
							LogMsg("AbortExposure", MessageLevel.msgOK, "CanAbortExposure is false and .NET exception correctly generated");
						}
					}
				default:
					LogMsg("AbortExposure", MessageLevel.msgError, EX_COM + "camera is not idle, further AbortExposure tests skipped: " + m_CameraState.ToString);
			}
		} catch (COMException ex) {
			LogMsg("AbortExposure", MessageLevel.msgError, EX_COM + "exception generated when reading camera state, further AbortExposure tests skipped");
		} catch (Exception ex) {
			LogMsg("AbortExposure", MessageLevel.msgError, EX_NET + "exception generated when reading camera state, further AbortExposure tests skipped");
		}
		//Pulseguide
		//Should work OK
		if (m_CanPulseGuide) {
			try {
				CameraPulseGuideTest(GuideDirections.guideNorth);
				if (TestStop())
					return;
				CameraPulseGuideTest(GuideDirections.guideSouth);
				if (TestStop())
					return;
				CameraPulseGuideTest(GuideDirections.guideEast);
				if (TestStop())
					return;
				CameraPulseGuideTest(GuideDirections.guideWest);
				if (TestStop())
					return;
			} catch (COMException ex) {
				LogMsg("PulseGuide", MessageLevel.msgError, EX_COM + "CanPulseGuide is true but exception generated when calling method - " + ex.ToString);
			} catch (Exception ex) {
				LogMsg("PulseGuide", MessageLevel.msgError, EX_NET + "CanPulseGuide is true but exception generated when calling method - " + ex.ToString);
			}
		//CanPulseGuide is false, should generate an exception
		} else {
			try {
				m_Camera.PulseGuide(GuideDirections.guideNorth, 0);
				LogMsg("PulseGuide", MessageLevel.msgError, "CanPulseGuide is false but no exception generated when calling method");
			} catch (COMException ex) {
				LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and exception correctly generated when calling method");
			} catch (MethodNotImplementedException ex) {
				LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and Pulseguide is not implemented in this driver");
			} catch (Exception ex) {
				LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and exception correctly generated when calling method");
			}
		}

		//Stopexposure
		try {
			m_CameraState = m_Camera.CameraState;
			switch (m_CameraState) {
				case CameraStates.cameraIdle:
					//all is OK so test that a call to StopExposure doesn't generate an exception
					try {
						m_Camera.StopExposure();
						//Can abort and no exception so OK
						if (m_CanStopExposure) {
							LogMsg("StopExposure", MessageLevel.msgOK, "No exception generated when camera is already idle");
						//Can not abort and no exception - so violates specification
						} else {
							LogMsg("StopExposure", MessageLevel.msgError, "CanStopExposure is false but no exception is generated when AbortExposure is called");
						}
					} catch (COMException ex) {
						//Can abort so should not have generated an error
						if (m_CanStopExposure) {
							LogMsg("StopExposure", MessageLevel.msgError, EX_COM + "exception incorrectly generated when camera is idle");
						//Can not abort so should have generated an error
						} else {
							LogMsg("StopExposure", MessageLevel.msgOK, "CanStopExposure is false and COM exception correctly generated");
						}
					} catch (Exception ex) {
						//Can abort so should not have generated an error
						if (m_CanStopExposure) {
							LogMsg("StopExposure", MessageLevel.msgError, EX_NET + "exception incorrectly generated when camera is idle");
						//Can not abort so should have generated an error
						} else {
							LogMsg("StopExposure", MessageLevel.msgOK, "CanStopExposure is false and .NET exception correctly generated");
						}
					}
				default:
					LogMsg("StopExposure", MessageLevel.msgError, EX_COM + "camera is not idle, further StopExposure tests skipped: " + m_CameraState.ToString);
			}
		} catch (COMException ex) {
			LogMsg("StopExposure", MessageLevel.msgError, EX_COM + "exception generated when reading camera state, further StopExposure tests skipped");
		} catch (Exception ex) {
			LogMsg("StopExposure", MessageLevel.msgError, EX_NET + "exception generated when reading camera state, further StopExposure tests skipped");
		}

		//StartExposure - Confirm that correct operation occurs
		int l_BinX;
		int l_BinY;
		//Test all combinations
		if (m_CanAsymmetricBin) {
			for (l_BinY = 1; l_BinY <= m_MaxBinY; l_BinY++) {
				for (l_BinX = 1; l_BinX <= m_MaxBinX; l_BinX++) {
					CameraExposure("Take image full frame " + l_BinX + " x " + l_BinY + " bin", l_BinX, l_BinY, 0, 0, m_CameraXSize / l_BinX, m_CameraYSize / l_BinY, 2.0, "");
					if (TestStop())
						return;
				}
			}

		//Test same bin combinations only
		} else {
			for (l_BinX = 1; l_BinX <= m_MaxBinX; l_BinX++) {
				CameraExposure("Take image full frame " + l_BinX + " x " + l_BinX + " bin", l_BinX, l_BinX, 0, 0, m_CameraXSize / l_BinX, m_CameraYSize / l_BinX, 2.0, "");
				if (TestStop())
					return;
			}
		}

		//StartExposure - Confirm error cases
		LogMsg("", MessageLevel.msgAlways, "");
		LogMsg("StartExposure error cases", MessageLevel.msgAlways, "");

		//StartExposure - Invalid Bin values
		//Test writing low and high values outside maximum range
		//If m_CanAsymmetricBin Then
		// CameraExposure("", 0, 1, 0, 0, m_CameraXSize, m_CameraYSize, 0.1, "illegal X Binning: 0x1") : If TestStop() Then Exit Sub 'Bin size < 1
		// CameraExposure("", m_MaxBinX + 1, 1, 0, 0, m_CameraXSize, m_CameraYSize, 0.1, "illegal X Binning: " & m_MaxBinX + 1 & "x1") : If TestStop() Then Exit Sub 'Bin size > MaxBin
		// CameraExposure("", 1, 0, 0, 0, m_CameraXSize, m_CameraYSize, 0.1, "illegal Y Binning: 1x0") : If TestStop() Then Exit Sub 'Bin size < 1
		// CameraExposure("", 1, m_MaxBinY + 1, 0, 0, m_CameraXSize, m_CameraYSize, 0.1, "illegal Y Binning: 1x" & m_MaxBinY + 1) : If TestStop() Then Exit Sub 'Bin size > MaxBin
		// Else
		// CameraExposure("", 0, 0, 0, 0, m_CameraXSize, m_CameraYSize, 0.1, "illegal Binning: 0x0") : If TestStop() Then Exit Sub 'Bin size < 1
		// CameraExposure("", m_MaxBinX + 1, m_MaxBinX + 1, 0, 0, m_CameraXSize, m_CameraYSize, 0.1, "illegal Binning: " & m_MaxBinX + 1 & "x" & m_MaxBinX + 1) : If TestStop() Then Exit Sub 'Bin size > MaxBin
		// End If

		//StartExposure - Negative time
		CameraExposure("", 1, 1, 0, 0, m_CameraXSize, m_CameraYSize, -1.0, "negative duration");
		if (TestStop())
			return;
		//Test that negative duration generates an error

		//StartExposure - Outside expected values for different binnings
		for (l_i = 1; l_i <= m_MaxBinX; l_i++) {
			for (l_j = 1; l_j <= m_MaxBinY; l_j++) {
				//Carry out for all X and Y bin values
				if (m_CanAsymmetricBin) {
					CameraExposure("", l_i, l_j, 0, 0, (int)(m_CameraXSize / l_i) + 1, (int)m_CameraYSize / l_j, 0.1, "X size larger than binned chip size, Bin " + l_i + "x" + l_j);
					if (TestStop())
						return;
					//X size too large for binned size
					CameraExposure("", l_i, l_j, 0, 0, (int)m_CameraXSize / l_i, (int)(m_CameraYSize / l_j) + 1, 0.1, "Y size larger than binned chip size, Bin " + l_i + "x" + l_j);
					if (TestStop())
						return;
					//Y size too large for binned size
					CameraExposure("", l_i, l_j, (int)(m_CameraXSize / l_i) + 1, 0, (int)m_CameraXSize / l_i, (int)m_CameraYSize / l_j, 0.1, "X start outside binned chip size, Bin " + l_i + "x" + l_j);
					if (TestStop())
						return;
					//X start outside binned chip dimensions
					CameraExposure("", l_i, l_j, 0, (int)(m_CameraYSize / l_j) + 1, (int)m_CameraXSize / l_i, (int)m_CameraYSize / l_j, 0.1, "Y start outside binned chip size, Bin " + l_i + "x" + l_j);
					if (TestStop())
						return;
					//Y start outside binned chip dimensions
				//Only carry out where X and Y bin are equal
				} else {
					//X and Y bin are same
					if (l_i == l_j) {
						CameraExposure("", l_i, l_j, 0, 0, (int)(m_CameraXSize / l_i) + 1, (int)m_CameraYSize / l_j, 0.1, "X size larger than binned chip size, Bin " + l_i + "x" + l_j);
						if (TestStop())
							return;
						//X size too large for binned size
						CameraExposure("", l_i, l_j, 0, 0, (int)m_CameraXSize / l_i, (int)(m_CameraYSize / l_j) + 1, 0.1, "Y size larger than binned chip size, Bin " + l_i + "x" + l_j);
						if (TestStop())
							return;
						//Y size too large for binned size
						CameraExposure("", l_i, l_j, (int)(m_CameraXSize / l_i) + 1, 0, (int)m_CameraXSize / l_i, (int)m_CameraYSize / l_j, 0.1, "X start outside binned chip size, Bin " + l_i + "x" + l_j);
						if (TestStop())
							return;
						//X start outside binned chip dimensions
						CameraExposure("", l_i, l_j, 0, (int)(m_CameraYSize / l_j) + 1, (int)m_CameraXSize / l_i, (int)m_CameraYSize / l_j, 0.1, "Y start outside binned chip size, Bin " + l_i + "x" + l_j);
						if (TestStop())
							return;
						//Y start outside binned chip dimensions
					}
				}
			}
		}
	}
	private void CameraExposure(string p_Description, int p_BinX, int p_BinY, int p_StartX, int p_StartY, int p_NumX, int p_NumY, double p_Duration, string p_ExpectedErrorMessage)
	{
		string l_NumPlanes;
		string l_VariantType;
		string l_PercentCompletedMessage;
		bool l_ExposeOK;
		//Flag to determine whether we were successful or something failed
		System.DateTime l_StartTime;
		System.DateTime l_StartTimeUTC;
		System.DateTime l_EndTime;
		short l_PercentCompleted;

		if (p_Description != "") {
			LogMsg("", MessageLevel.msgAlways, "");
			//Blank Line
			LogMsg(p_Description, MessageLevel.msgAlways, "");

		}
		try {
			m_Camera.BinX = (short)p_BinX;
			m_Camera.BinY = (short)p_BinY;
			try {
				l_ExposeOK = false;
				//Start off by assuming the worst
				m_Camera.StartX = p_StartX;
				m_Camera.StartY = p_StartY;
				m_Camera.NumX = p_NumX;
				m_Camera.NumY = p_NumY;
				try {
					Status(StatusType.staAction, "Start " + p_Duration.ToString + " second synchronous exposure");
					l_StartTime = Now;
					l_StartTimeUTC = System.DateTime.UtcNow;
					m_Camera.StartExposure(p_Duration, true);
					//Not expecting an error and didn't get one
					if (p_ExpectedErrorMessage == "") {
						l_EndTime = Now;
						//Probably a synchrous camera
						if (m_Camera.ImageReady & (m_Camera.CameraState == CameraStates.cameraIdle)) {
							//Is a synchronous camera
							if (l_EndTime.Subtract(l_StartTime).TotalSeconds >= p_Duration) {
								LogMsg("StartExposure", MessageLevel.msgOK, "Synchronous exposure found OK: " + p_Duration + " seconds");
								CameraTestLast(p_Duration, l_StartTimeUTC);
							// Is synchronous but returned early!
							} else {
								LogMsg("StartExposure", MessageLevel.msgError, "Synchronous exposure found but image was returned before exposure time was complete");
							}
						//Assume asynchronous
						} else {
							Status(StatusType.staAction, "Waiting for exposure to start");
							//Wait for exposing state
							do {
								WaitFor(CAMERA_SLEEP_TIME);
								if (TestStop())
									return;
							} while (!((m_Camera.CameraState == CameraStates.cameraExposing) | (m_Camera.CameraState == CameraStates.cameraError)));

							l_StartTime = Now;
							l_StartTimeUTC = System.DateTime.UtcNow;
							//Wait for state following exposing
							do {
								l_PercentCompletedMessage = "Not present in a V1 driver";
								// Initialise PercentCompleted message
								if (m_Camera.InterfaceVersion > 1) {
									try {
										l_PercentCompleted = m_Camera.PercentCompleted;
										l_PercentCompletedMessage = l_PercentCompleted.ToString;
										// Operation completed OK
									} catch (COMException ex) {
										switch (ex.ErrorCode) {
											case ErrorCodes.NotImplemented:
												// Not implemented
												l_PercentCompletedMessage = "COM - Not implemented";
											case ErrorCodes.InvalidOperationException:
												// Not valid at this time
												l_PercentCompletedMessage = "COM - Invalid operation";
											default:
												// Something bad happened!
												l_PercentCompletedMessage = "COM - Exception: " + ex.Message + " 0x" + ex.ErrorCode.ToString("X8");
										}
									} catch (ASCOM.InvalidOperationException ex) {
										l_PercentCompletedMessage = ".NET - Invalid operation";
										// Not valid at this time
									} catch (PropertyNotImplementedException ex) {
										l_PercentCompletedMessage = "Not implemented";
										// Not implemented
									} catch (Exception ex) {
										l_PercentCompletedMessage = "Exception: " + ex.Message;
										// Something bad happened!
									}
								}
								Status(StatusType.staAction, "Waiting for " + p_Duration.ToString + " second exposure to complete: " + Int(Now.Subtract(l_StartTime).TotalSeconds) + ",   PercentComplete: " + l_PercentCompletedMessage);
								WaitFor(CAMERA_SLEEP_TIME);
								if (TestStop())
									return;
							} while (!((m_Camera.CameraState != CameraStates.cameraExposing) | (m_Camera.CameraState == CameraStates.cameraError)));

							l_EndTime = Now;
							Status(StatusType.staAction, "Waiting for cameraidle state, reading/downloading image");
							//Wait for image to become ready
							do {
								WaitFor(CAMERA_SLEEP_TIME);
								if (TestStop())
									return;
							} while (!((m_Camera.CameraState == CameraStates.cameraIdle) | (m_Camera.CameraState == CameraStates.cameraError)));

							Status(StatusType.staAction, "Waiting for image ready");
							//Wait for image to become ready
							do {
								WaitFor(CAMERA_SLEEP_TIME);
								if (TestStop())
									return;
							} while (!((m_Camera.ImageReady) | (m_Camera.CameraState == CameraStates.cameraError)));

							//Worked OK
							if (m_Camera.ImageReady) {
								LogMsg("StartExposure", MessageLevel.msgOK, "Asynchronous exposure found OK: " + p_Duration + " seconds");
								CameraTestLast(p_Duration, l_StartTimeUTC);
							//Failed somehow
							} else {
								LogMsg("StartExposure", MessageLevel.msgError, "Camerastate is CameraError");
							}
						}
						l_ExposeOK = true;
						//Camera exposed OK and didn't generate an exception

						//Check image array dimensions
						try {
							m_ImageArray = (Array)m_Camera.ImageArray;
							if ((m_ImageArray.GetLength(0) == p_NumX) & (m_ImageArray.GetLength(1) == p_NumY)) {
								if (m_ImageArray.GetType.ToString == "System.Int32[,]" | m_ImageArray.GetType.ToString == "System.Int32[,,]") {
									//Single plane image be definition
									if (m_ImageArray.Rank == 2) {
										l_NumPlanes = "1 plane";
									//Read the number of image planes from the maximum value of the third aray index
									} else {
										l_NumPlanes = "1 plane";
										//More than 1 plane
										if (m_ImageArray.GetUpperBound(2) > 0) {
											l_NumPlanes = (string)m_ImageArray.GetUpperBound(2) + 1 + " planes";
										}
									}
									LogMsg("ImageArray", MessageLevel.msgOK, "Successfully read 32 bit integer array (" + l_NumPlanes + ") " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1) + " pixels");
								} else {
									LogMsg("ImageArray", MessageLevel.msgError, "Expected 32 bit integer array, actually got: " + m_ImageArray.GetType.ToString);
								}
							} else {
								if ((m_ImageArray.GetLength(0) == p_NumY) & (m_ImageArray.GetLength(1) == p_NumX)) {
									LogMsg("ImageArray", MessageLevel.msgError, "Camera image dimensions swapped, expected values: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1));
								} else {
									LogMsg("ImageArray", MessageLevel.msgError, "Camera image does not have the expected dimensions of: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1));
								}
							}
						} catch (COMException ex) {
							LogMsg("StartExposure", MessageLevel.msgError, EX_COM + "exception when reading ImageArray" + ex.ToString);
						} catch (Exception ex) {
							LogMsg("StartExposure", MessageLevel.msgError, EX_NET + "exception when reading ImageArray" + ex.ToString);
						}

						//Check image array variant dimensions
						try {
							m_ImageArrayVariant = (Array)m_Camera.ImageArrayVariant;
							if ((m_ImageArrayVariant.GetLength(0) == p_NumX) & (m_ImageArrayVariant.GetLength(1) == p_NumY)) {
								if (m_ImageArrayVariant.GetType.ToString == "System.Object[,]" | m_ImageArrayVariant.GetType.ToString == "System.Object[,,]") {
									//Single plane image be definition
									if (m_ImageArrayVariant.Rank == 2) {
										l_NumPlanes = "1 plane";
										l_VariantType = m_ImageArrayVariant(0, 0).GetType.ToString();
									//Read the number of image planes from the maximum value of the third aray index
									} else {
										l_NumPlanes = "1 plane";
										//More than 1 plane
										if (m_ImageArrayVariant.GetUpperBound(2) > 0) {
											l_NumPlanes = (string)m_ImageArrayVariant.GetUpperBound(2) + 1 + " planes";
											l_VariantType = m_ImageArrayVariant(0, 0, 0).GetType.ToString();
										//Just one plane
										} else {
											l_VariantType = m_ImageArrayVariant(0, 0).GetType.ToString();
										}
									}
									LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Successfully read variant array (" + l_NumPlanes + ") with " + l_VariantType + " elements " + m_ImageArrayVariant.GetLength(0) + " x " + m_ImageArrayVariant.GetLength(1) + " pixels");
								} else {
									LogMsg("ImageArrayVariant", MessageLevel.msgError, "Expected variant array, actually got: " + m_ImageArrayVariant.GetType.ToString);
								}
							} else {
								if ((m_ImageArrayVariant.GetLength(0) == p_NumY) & (m_ImageArrayVariant.GetLength(1) == p_NumX)) {
									LogMsg("ImageArrayVariant", MessageLevel.msgError, "Camera image dimensions swapped, expected values: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArrayVariant.GetLength(0) + " x " + m_ImageArrayVariant.GetLength(1));
								} else {
									LogMsg("ImageArrayVariant", MessageLevel.msgError, "Camera image does not have the expected dimensions of: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArrayVariant.GetLength(0) + " x " + m_ImageArrayVariant.GetLength(1));
								}
							}
						} catch (COMException ex) {
							LogMsg("ImageArrayVariant", MessageLevel.msgError, EX_COM + "exception when reading ImageArrayVariant" + ex.ToString);
						} catch (Exception ex) {
							LogMsg("ImageArrayVariant", MessageLevel.msgError, EX_NET + "exception when reading ImageArrayVariant" + ex.ToString);
						}
					//Expecting an error and didn't get one!
					} else {
						LogMsg("StartExposure", MessageLevel.msgComment, "Test: " + p_ExpectedErrorMessage);
						LogMsg("StartExposure", MessageLevel.msgError, "Expected an exception and didn't get one - BinX:" + p_BinX + " BinY:" + p_BinY + " StartX:" + p_StartX + " StartY:" + p_StartY + " NumX:" + p_NumX + " NumY:" + p_NumY);
						try {
							m_Camera.StopExposure();
						} catch (Exception ex) {
						}
						// Try and do some clean up
						try {
							m_Camera.AbortExposure();
						} catch (Exception ex) {
						}
					}

				} catch (COMException ex) {
					if (p_ExpectedErrorMessage != "") {
						LogMsg("StartExposure", MessageLevel.msgOK, "Exception correctly generated for " + p_ExpectedErrorMessage);
					} else {
						LogMsg("", MessageLevel.msgInfo, p_ExpectedErrorMessage);
						LogMsg("StartExposure", MessageLevel.msgError, EX_COM + "exception generated when exposing, further StartExposure tests skipped - " + ex.ToString);
					}
				} catch (Exception ex) {
					if (p_ExpectedErrorMessage != "") {
						LogMsg("StartExposure", MessageLevel.msgOK, "Exception correctly generated for " + p_ExpectedErrorMessage);
					} else {
						LogMsg("StartExposure", MessageLevel.msgError, EX_NET + "exception generated when exposing, further StartExposure tests skipped - " + ex.ToString);
					}
				}

			} catch (COMException ex) {
				LogMsg("StartExposure", MessageLevel.msgInfo, "COM Exception: " + ex.Message);
				LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped");
			} catch (Exception ex) {
				LogMsg("StartExposure", MessageLevel.msgInfo, ".NET Exception: " + ex.Message);
				LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped");
			}
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("StartExposure", MessageLevel.msgInfo, ex.Message);
				default:
					LogMsg("StartExposure", MessageLevel.msgInfo, "COM Exception: " + ex.Message);
					LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped");
			}
		} catch (InvalidValueException ex) {
			LogMsg("BinXY Write", MessageLevel.msgInfo, "Exposure skipped because BinX or BinY cannot be set. Exception message:");
			LogMsg("BinXY Write", MessageLevel.msgComment, ex.Message);
		} catch (Exception ex) {
			LogMsg("StartExposure", MessageLevel.msgInfo, ".NET Exception: " + ex.Message);
			LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped");

		}

		Status(StatusType.staAction, "");
	}
	private void CameraTestLast(double p_Duration, System.DateTime p_Start)
	{
		System.DateTime l_StartTime;

		//LastExposureDuration
		try {
			m_LastExposureDuration = m_Camera.LastExposureDuration;
			// Measured is within 2% of requested
			if ((Math.Abs(m_LastExposureDuration - p_Duration) / p_Duration) < 0.02) {
				LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration is: " + m_LastExposureDuration + " seconds");
			} else {
				LogMsg("LastExposureDuration", MessageLevel.msgIssue, "LastExposureDuration is not the same as image duration: " + m_LastExposureDuration + " " + p_Duration);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("LastExposureDuration", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("LastExposureDuration", ex.Message, EX_COM + "reading LastExposureDuration - " + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("LastExposureDuration", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("LastExposureDuration", ex.Message, EX_NET + "reading LastExposureDuration - " + ex.Message);
		}

		//LastExposurestartTime
		//Confirm that it can be read
		try {
			m_LastExposureStartTime = m_Camera.LastExposureStartTime;
			//Confirm that the format is as expected
			int l_i;
			bool l_FormatOK;
			l_FormatOK = true;
			//String is at least of correct minimum length
			if (Len(m_LastExposureStartTime) >= 19) {
				for (l_i = 1; l_i <= 19; l_i++) {
					switch (l_i) {
						case 5:
						case 8:
							//"-"
							if (Mid(m_LastExposureStartTime, l_i, 1) != "-")
								l_FormatOK = false;
						case 11:
							// "T"
							if (Mid(m_LastExposureStartTime, l_i, 1) != "T")
								l_FormatOK = false;
						case 14:
						case 17:
							//":"
							if (Mid(m_LastExposureStartTime, l_i, 1) != ":")
								l_FormatOK = false;
						default:
							// Numeric
							if (!IsNumeric(Mid(m_LastExposureStartTime, l_i, 1)))
								l_FormatOK = false;
					}
				}
				//Passed format check so nopw try and read as a datetime to compare with expected value
				if (l_FormatOK) {
					//Confirm that it parses as a valid date and check for correct value
					try {
						l_StartTime = System.DateTime.Parse(m_LastExposureStartTime);
						if (p_Start.Subtract(l_StartTime).TotalSeconds < 2.0) {
							LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime is correct to within 2 seconds: " + m_LastExposureStartTime + " UTC");
						} else {
							LogMsg("LastExposureStartTime", MessageLevel.msgIssue, "LastExposureStartTime is more than 2 seconds inaccurate : " + m_LastExposureStartTime + ", expected: " + Format(p_Start, "yyyy-MM-ddTHH:mm:ss") + " UTC");
						}
					} catch (COMException ex) {
						LogMsg("LastExposureStartTime", MessageLevel.msgError, EX_COM + "parsing LastExposureStartTime - " + ex.Message + " " + Hex(ex.ErrorCode));
					} catch (Exception ex) {
						LogMsg("LastExposureStartTime", MessageLevel.msgError, EX_NET + "parsing LastExposureStartTime - " + ex.Message + " " + m_LastExposureStartTime);
					}
				//Give format error message
				} else {
					LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime not in the expected format yyyy-mm-ddThh:mm:ss - " + m_LastExposureStartTime);
				}
			//Not of required length so give an error message
			} else {
				//Empty string
				if (m_LastExposureStartTime == "") {
					LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime has returned an empty string - expected yyyy-mm-ddThh:mm:ss");
				//Some characters but less than 19
				} else {
					LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime is less than 19 characters - expected yyyy-mm-ddThh:mm:ss - " + m_LastExposureStartTime);
				}
			}

		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("LastExposureStartTime", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("LastExposureStartTime", ex.Message, EX_COM + "reading LastExposureStartTime - " + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (Exception ex) {
			ExTest("LastExposureStartTime", ex.Message, EX_NET + "reading LastExposureStartTime - " + ex.Message + " " + m_LastExposureStartTime);
		}


	}
	private void CameraPulseGuideTest(GuideDirections p_Direction)
	{
		System.DateTime l_StartTime;
		System.DateTime l_EndTime;
		l_StartTime = Now;
		Status(StatusType.staAction, "Start " + CAMERA_PULSE_DURATION / 1000 + " second pulse guide " + p_Direction.ToString);
		m_Camera.PulseGuide(p_Direction, CAMERA_PULSE_DURATION);
		//Start a 2 second pulse
		l_EndTime = Now;
		// IsPulseGuiding is supported so go ahead and use it
		if (m_IsPulseGuidingSupported) {
			//If less than 1.5 seconds then assume we have returned early
			if (l_EndTime.Subtract(l_StartTime).TotalMilliseconds < (CAMERA_PULSE_DURATION - 500)) {
				if (m_Camera.IsPulseGuiding) {
					do {
						WaitFor(SLEEP_TIME);
						Application.DoEvents();
						if (TestStop())
							return;
					} while (!((!m_Camera.IsPulseGuiding) | (Now.Subtract(l_StartTime).TotalMilliseconds > 3000)));
					//Wait for up to 3 seconds
					if (!m_Camera.IsPulseGuiding) {
						LogMsg("PulseGuide " + p_Direction.ToString, MessageLevel.msgOK, "Asynchronous pulseguide found OK");
					} else {
						LogMsg("PulseGuide " + p_Direction.ToString, MessageLevel.msgIssue, "Asynchronous pulseguide expected but IsPulseGuiding is TRUE beyond expected time of 2 seconds");
					}
				} else {
					LogMsg("PulseGuide " + p_Direction.ToString, MessageLevel.msgIssue, "Asynchronous pulseguide expected but IsPulseGuiding has returned FALSE");
				}
			//Assume synchronous pulseguide and that ispulseguideing is false
			} else {
				if (!m_Camera.IsPulseGuiding) {
					LogMsg("PulseGuide " + p_Direction.ToString, MessageLevel.msgOK, "Synchronous pulseguide found OK");
				} else {
					LogMsg("PulseGuide " + p_Direction.ToString, MessageLevel.msgIssue, "Synchronous pulseguide expected but IsPulseGuiding has returned TRUE");
				}
			}
		//IsPulseGuiding is not supported so test for synchronous move
		} else {
			switch (l_EndTime.Subtract(l_StartTime).TotalMilliseconds - CAMERA_PULSE_DURATION) {
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
CAMERA_PULSE_TOLERANCE:
					//Duration was more than 0.5 seconds longer than expected
					LogMsg("PulseGuide " + p_Direction.ToString, MessageLevel.msgIssue, "Synchronous pulseguide longer than expected " + (CAMERA_PULSE_DURATION) / 1000 + " seconds: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
20:
					//Duration was more than 20ms shorter than expected
					LogMsg("PulseGuide " + p_Direction.ToString, MessageLevel.msgIssue, "Synchronous pulseguide shorter than expected " + (CAMERA_PULSE_DURATION) / 1000 + " seconds: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
				default:
					//Within acceptable tolerance
					LogMsg("PulseGuide " + p_Direction.ToString, MessageLevel.msgOK, "Synchronous pulseguide found OK: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
			}
		}
	}

	public override void CheckPerformance()
	{
		CameraPerformanceTest(CameraPerformance.CameraState, "CameraState");
		CameraPerformanceTest(CameraPerformance.CCDTemperature, "CCDTemperature");
		CameraPerformanceTest(CameraPerformance.CoolerPower, "CoolerPower");
		CameraPerformanceTest(CameraPerformance.HeatSinkTemperature, "HeatSinkTemperature");
		CameraPerformanceTest(CameraPerformance.ImageReady, "ImageReady");
		if (m_CanPulseGuide)
			CameraPerformanceTest(CameraPerformance.IsPulseGuiding, "IsPulseGuiding");
		Status(StatusType.staAction, "Exposure for ImageArray Test");
		Status(StatusType.staStatus, "Start");
		m_Camera.BinX = 1;
		m_Camera.BinY = 1;
		m_Camera.StartX = 0;
		m_Camera.StartY = 0;
		m_Camera.NumX = m_Camera.MaxBinX;
		m_Camera.NumY = m_Camera.MaxBinY;
		m_Camera.StartExposure(1, true);
		//1 second exposure
		do {
			Status(StatusType.staStatus, "Waiting for ImageReady");
		} while (!(m_Camera.ImageReady));
		Status(StatusType.staStatus, "Finished");
		CameraPerformanceTest(CameraPerformance.ImageArray, "ImageArray");
		CameraPerformanceTest(CameraPerformance.ImageArrayVariant, "ImageArrayVariant");
	}
	private void CameraPerformanceTest(CameraPerformance p_Type, string p_Name)
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
					case CameraPerformance.CameraState:
						m_CameraState = m_Camera.CameraState;
					case CameraPerformance.CCDTemperature:
						m_CCDTemperature = m_Camera.CCDTemperature;
					case CameraPerformance.CoolerPower:
						m_CoolerPower = m_Camera.CoolerPower;
					case CameraPerformance.HeatSinkTemperature:
						m_HeatSinkTemperature = m_Camera.HeatSinkTemperature;
					case CameraPerformance.ImageReady:
						m_ImageReady = m_Camera.ImageReady;
					case CameraPerformance.IsPulseGuiding:
						m_IsPulseGuiding = m_Camera.IsPulseGuiding;
					case CameraPerformance.ImageArray:
						m_ImageArray = (Array)m_Camera.ImageArray;
					case CameraPerformance.ImageArrayVariant:
						m_ImageArrayVariant = (Array)m_Camera.ImageArrayVariant;
					default:
						LogMsg(p_Name, MessageLevel.msgError, "Conform:PerformanceTest: Unknown test type " + p_Type.ToString);
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

	public override void PostRunCheck()
	{
		if (m_CanAbortExposure)
			try {
				m_Camera.AbortExposure();
			} catch {
			}
		if (m_CanStopExposure)
			try {
				m_Camera.StopExposure();
			} catch {
			}
		if (m_CanSetCCDTemperature)
			try {
				m_Camera.SetCCDTemperature = m_SetCCDTemperature;
			} catch {
			}
		try {
			m_Camera.CoolerOn = m_CoolerOn;
		} catch {
		}
		LogMsg("PostRunCheck", MessageLevel.msgOK, "Camera returned to intitial cooler temperature");
	}

	#End Region
}
