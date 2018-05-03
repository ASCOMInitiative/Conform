//Option Strict On
using ASCOM.DeviceInterface;
using System.Reflection;

internal class TelescopeTester : DeviceTesterBaseClass
{

	#Region "Variables and Constants"
		//Time to wait between changing Tracking state
	private const int TRACKING_COMMAND_DELAY = 1000;
		//Performance loop run time in seconds
	private const int PERF_LOOP_TIME = 5;
		// Number of milliseconds for each move axis command
	private const int MOVE_AXIS_TIME = 2000;
	private const int NUM_AXIS_RATES = 1000;
		// Mnemonics for the axis rate array second dimension
	private const int AXIS_RATE_MINIMUM = 0;
	private const int AXIS_RATE_MAXIMUM = 1;
		// Minimum number of seconds to wait before declaring an async slew is finished (allows for mounts that dopn't set IsSlewing straight away!)
	private const int WAIT_FOR_SLEW_MINIMUM_DURATION = 5;
		// 7 minutes in milliseocnds
	private const int SIDEOFPIER_MERIDIAN_TRACKING_PERIOD = 7 * 60 * 1000;

	private bool m_CanFindHome;
	private bool m_CanPark;
	private bool m_CanPulseGuide;
	private bool m_CanSetDeclinationRate;
	private bool m_CanSetGuideRates;
	private bool m_CanSetPark;
	private bool m_CanSetPierside;
	private bool m_CanSetRightAscensionRate;
	private bool m_CanSetTracking;
	private bool m_CanSlew;
	private bool m_CanSlewAltAz;
	private bool m_CanSlewAltAzAsync;
	private bool m_CanSlewAsync;
	private bool m_CanSync;
	private bool m_CanSyncAltAz;
	private bool m_CanUnpark;
	private ASCOM.DeviceInterface.AlignmentModes m_AlignmentMode;
	private double m_Altitude;
	private double m_ApertureArea;
	private double m_ApertureDiameter;
	private bool m_AtHome;
	private bool m_AtPark;
	private double m_Azimuth;
	private double m_Declination;
	private double m_DeclinationRate;
	private bool m_DoesRefraction;
	private ASCOM.DeviceInterface.EquatorialCoordinateType m_EquatorialSystem;
	private double m_FocalLength;
	private double m_GuideRateDeclination;
	private double m_GuideRateRightAscension;
	private bool m_IsPulseGuiding;
	private double m_RightAscension;
	private double m_RightAscensionRate;
	private ASCOM.DeviceInterface.PierSide m_SideOfPier;
	private double m_SiderealTimeScope;
	private double m_SiteElevation;
	private double m_SiteLatitude;
	private double m_SiteLongitude;
	private bool m_Slewing;
	private short m_SlewSettleTime;
	private double m_TargetDeclination;
	private double m_TargetRightAscension;
	private bool m_Tracking;
	private System.DateTime m_UTCDate;
	private bool m_CanMoveAxisPrimary;
	private bool m_CanMoveAxisSecondary;
	private bool m_CanMoveAxisTertiary;
	private ASCOM.DeviceInterface.PierSide m_DestinationSideOfPier;
	private ASCOM.DeviceInterface.PierSide m_DestinationSideOfPierEast;
	private ASCOM.DeviceInterface.PierSide m_DestinationSideOfPierWest;
	private double m_SiderealTimeASCOM;
	private double m_CurrentRA;
	private double m_CurrentDEC;
	private double m_CurrentAz;
	private double m_CurrentAlt;
	private double m_NewRA;
	private double m_NewDEC;
	private double m_NewAlt;
	private double m_NewAz;
	private double m_SyncRA;
	private double m_SyncDEC;
	private double m_SyncAlt;
	private double m_SyncAz;
	private System.DateTime m_StartTime;
	private System.DateTime m_EndTime;
	private System.DateTime m_StartTimeUTC;
	private string m_LastLogFileDirectory;
	private bool m_CanReadSideOfPier;
	private double m_TargetAltitude;
	private double m_TargetAzimuth;
	private bool m_CanReadAltitide;

	private bool m_CanReadAzimuth;

	private DeviceType m_TelescopeType;
	#If DEBUG Then
	private object m_Telescope;
	private object m_AxisRatesPrimary;
	private object m_AxisRatesSecondary;
	private object m_AxisRatesTertiary;
	private object m_Rate;
	private object m_TrackingRates;
		#Else
	private object m_TrackingRate;
	private ASCOM.DriverAccess.Telescope m_Telescope;
	private ASCOM.DeviceInterface.IAxisRates m_AxisRatesPrimary;
	private ASCOM.DeviceInterface.IAxisRates m_AxisRatesSecondary;
	private ASCOM.DeviceInterface.IAxisRates m_AxisRatesTertiary;
	private ASCOM.DeviceInterface.IRate m_Rate;
	private ASCOM.DeviceInterface.ITrackingRates m_TrackingRates;
		#End If
	private ASCOM.DeviceInterface.DriveRates m_TrackingRate;
	//Axisrate checks
	private bool m_AxisRatePrimaryOverlap;
	private bool m_AxisRateSecondaryOverlap;
	private bool m_AxisRateTertiaryOverlap;
	private double[,] m_AxisRatesPrimaryArray = new[];
		#End Region
	private double[,] m_AxisRatesArray = new double[NUM_AXIS_RATES - 1, 0];

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
		tstCanStopExposure = 21
	}
	private enum OptionalMethodType
	{
		tstAbortSlew = 1,
		tstDestinationSideOfPier = 2,
		tstFindHome = 3,
		tstMoveAxisPrimary = 4,
		tstMoveAxisSecondary = 5,
		tstMoveAxisTertiary = 6,
		tstPulseGuide = 7,
		tstSideOfPierWrite = 8
	}
	private enum RequiredMethodType
	{
		tstAxisrates = 1,
		tstCanMoveAxisPrimary = 2,
		tstCanMoveAxisSecondary = 3,
		tstCanMoveAxisTertiary = 4
	}
	private enum ParkedExceptionType
	{
		tstPExcepAbortSlew = 1,
		tstPExcepFindHome = 2,
		tstPExcepMoveAxisPrimary = 3,
		tstPExcepMoveAxisSecondary = 4,
		tstPExcepMoveAxisTertiary = 5,
		tstPExcepSlewToCoordinates = 6,
		tstPExcepSlewToCoordinatesAsync = 7,
		tstPExcepSlewToTarget = 8,
		tstPExcepSlewToTargetAsync = 9,
		tstPExcepSyncToCoordinates = 10,
		tstPExcepSyncToTarget = 11,
		tstPExcepPulseGuide = 12
	}
	private enum SyncType
	{
		tstSyncToCoordinates = 1,
		tstSyncToTarget = 2,
		tstSyncToAltAz = 3
	}
	private enum SlewType
	{
		tstSlewToCoordinates = 1,
		tstSlewToCoordinatesAsync = 2,
		tstSlewToTarget = 3,
		tstSlewToTargetAsync = 4,
		tstSlewToAltAz = 5,
		tstSlewToAltAzAsync = 6
	}
	private enum PerformanceType
	{
		tstPerfAltitude = 1,
		tstPerfAtHome = 1,
		tstPerfAtPark = 2,
		tstPerfAzimuth = 3,
		tstPerfDeclination = 4,
		tstPerfIsPulseGuiding = 5,
		tstPerfRightAscension = 6,
		tstPerfSideOfPier = 7,
		tstPerfSiderealTime = 8,
		tstPerfSlewing = 9,
		tstPerfUTCDate = 10
	}
	private enum FlipTestType
	{
		DestinationSideOfPier,
		SideOfPier
	}
	private enum InterfaceType
	{
		ITelescope,
		ITelescopeV2,
		ITelescopeV3
	}
	#End Region

	#Region "New and Dispose"
	private TelescopeTester()
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
		LogMsg("Dispose", MessageLevel.msgDebug, "Disposing of Telescope driver: " + disposing.ToString + " " + disposedValue.ToString);
		if (!this.disposedValue) {
			if (disposing) {
				// TODO: free other state (managed objects).
			}
			//Should be True but make False to stop Conform from cleanly dropping the telescope object (useful for retaining scopesim in memory to change flags
			if (true) {
				try {
					m_Telescope.Connected = false;
				} catch {
				}
				#If DEBUG Then
				try {
					Marshal.ReleaseComObject(m_Telescope);
				} catch {
				}
				#Else
				try {
					m_Telescope.Dispose();
				} catch {
				}
				try {
					m_Telescope = null;
				} catch {
				}
				#End If
				try {
					Marshal.ReleaseComObject(m_TrackingRates);
				} catch {
				}
				//Ensure each object is cleaned up if possible
				try {
					Marshal.ReleaseComObject(g_DeviceObject);
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_Rate);
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_AxisRatesPrimary);
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_AxisRatesSecondary);
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_AxisRatesTertiary);
				} catch {
				}

				m_Telescope = null;
				m_TrackingRates = null;
				g_DeviceObject = null;
				m_Rate = null;
				m_AxisRatesPrimary = null;
				m_AxisRatesSecondary = null;
				m_AxisRatesTertiary = null;
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
		base.CheckCommonMethods(m_Telescope, DeviceType.Telescope);
	}

	private override void CheckInitialise()
	{
		//Set the error type numbers acording to the standards adopted by individual authors.
		//Unfortunatley these vary between drivers so I have to allow for these here in order to give meaningful
		//messages to driver authors!
		switch (g_TelescopeProgID) {
			case "Hub.Telescope":
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040401;
				g_ExInvalidValue2 = 0x80040402;
				g_ExInvalidValue3 = 0x80040405;
				g_ExInvalidValue4 = 0x80040402;
				g_ExInvalidValue5 = 0x80040402;
				g_ExInvalidValue6 = 0x80040402;

				g_ExNotSet1 = 0x80040403;
			case "ASCOM.MI250SA.Telescope":
			case "Celestron.Telescope":
			case "ASCOM.MI250.Telescope":
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040401;
				g_ExInvalidValue2 = 0x80040402;
				g_ExInvalidValue3 = 0x80040402;
				g_ExInvalidValue4 = 0x80040402;
				g_ExInvalidValue5 = 0x80040402;
				g_ExInvalidValue6 = 0x80040402;
				g_ExNotSet1 = 0x80040403;
			case "TemmaLite.Telescope":
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040410;
				g_ExInvalidValue2 = 0x80040418;
				g_ExInvalidValue3 = 0x80040418;
				g_ExInvalidValue4 = 0x80040418;
				g_ExInvalidValue5 = 0x80040418;
				g_ExInvalidValue6 = 0x80040418;
				g_ExNotSet1 = 0x80040417;
			case "Gemini.Telescope":
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040410;
				g_ExInvalidValue2 = 0x80040418;
				g_ExInvalidValue3 = 0x80040419;
				g_ExInvalidValue4 = 0x80040420;
				g_ExInvalidValue5 = 0x80040420;
				g_ExInvalidValue6 = 0x80040420;
				g_ExNotSet1 = 0x80040417;
			case "POTH.Telescope":
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040405;
				g_ExInvalidValue2 = 0x80040406;
				g_ExInvalidValue3 = 0x80040406;
				g_ExInvalidValue4 = 0x80040406;
				g_ExInvalidValue5 = 0x80040406;
				g_ExInvalidValue6 = 0x80040406;
				g_ExNotSet1 = 0x80040403;
			case "ServoCAT.Telescope":
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = ErrorCodes.InvalidValue;
				g_ExInvalidValue2 = 0x80040405;
				g_ExInvalidValue3 = 0x80040405;
				g_ExInvalidValue4 = 0x80040405;
				g_ExInvalidValue5 = 0x80040405;
				g_ExInvalidValue6 = 0x80040405;
				g_ExNotSet1 = 0x80040403;
				g_ExNotSet2 = 0x80040404;
			default:
				//I'm using the simulator values as the defaults since it is the reference platform
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = ErrorCodes.InvalidValue;
				g_ExInvalidValue2 = 0x80040405;
				g_ExInvalidValue3 = 0x80040405;
				g_ExInvalidValue4 = 0x80040405;
				g_ExInvalidValue5 = 0x80040405;
				g_ExInvalidValue6 = 0x80040405;
				g_ExNotSet1 = 0x80040403;
		}
		base.CheckInitialise(g_TelescopeProgID);

	}
	private override void CheckAccessibility()
	{
		try {
			ASCOM.DriverAccess.Telescope l_DriverAccessTelescope = null;
			string l_ErrMsg = "";
			int l_TryCount = 0;
			try {
				LogMsg("AccessChecks", MessageLevel.msgDebug, "Before MyBase.CheckAccessibility");
				base.CheckAccessibility(g_TelescopeProgID, DeviceType.Telescope);
				LogMsg("AccessChecks", MessageLevel.msgDebug, "After MyBase.CheckAccessibility");
				try {
					TestEarlyBinding(InterfaceType.ITelescopeV2);
					TestEarlyBinding(InterfaceType.ITelescopeV3);

					//Try client access toolkit
					l_DriverAccessTelescope = null;
					l_TryCount = 0;
					do {
						l_TryCount += 1;
						try {
							l_DriverAccessTelescope = new ASCOM.DriverAccess.Telescope(g_TelescopeProgID);
							WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
							LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
							try {
								l_DriverAccessTelescope.Connected = true;
								LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
								l_DriverAccessTelescope.Connected = false;
								LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully disconnected using driver access toolkit");
							} catch (Exception ex) {
								LogMsg("AccessChecks", MessageLevel.msgError, "Error conecting to driver using driver access toolkit: " + ex.ToString);
								LogMsg("", MessageLevel.msgAlways, "");
							}
						} catch (Exception ex) {
							l_ErrMsg = ex.ToString;
						}
						if (l_DriverAccessTelescope == null)
							WaitFor(200);
					} while (!(((l_TryCount == 3) | (!l_DriverAccessTelescope == null))));
					//Exit if created OK

					if (l_DriverAccessTelescope == null) {
						LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " + l_ErrMsg);
						LogMsg("", MessageLevel.msgAlways, "");
					} else {
						LogMsg("Telescope:CreateDevice", MessageLevel.msgDebug, "Created telescope on attempt: " + l_TryCount.ToString);
					}
				} catch (Exception ex) {
					LogMsg("Telescope:CheckAcc.EX3", MessageLevel.msgError, ex.ToString);
				}
			} catch (Exception ex) {
				LogMsg("Telescope:CheckAcc.EX2", MessageLevel.msgError, ex.ToString);
			}

			//Clean up
			try {
				l_DriverAccessTelescope.Dispose();
			} catch {
			}
			//Try : Marshal.ReleaseComObject(l_DriverAccessTelescope) : Catch : End Try
			l_DriverAccessTelescope = null;
			GC.Collect();
			//Special handling for Gemini that takes a long time to dispose
			if ((g_TelescopeProgID.ToUpper == "ASCOM.GeminiTelescope.Telescope".ToUpper)) {
				WaitForAbsolute(GEMINI_DESTROY_WAIT, "Waiting for Accessibility Telescope Object to Dispose");
			} else {
				WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for Accessibility Telescope Object to Dispose");
			}
		} catch (Exception ex) {
			LogMsg("Telescope:CheckAcc.EX1", MessageLevel.msgError, ex.ToString);
		}
	}
	private override void CreateDevice()
	{
		int l_TryCount = 0;
		do {
			l_TryCount += 1;
			try {
				LogMsg("Telescope:CreateDevice", MessageLevel.msgDebug, "Creating ProgID: " + g_TelescopeProgID);
				#If DEBUG Then
				m_Telescope = CreateObject(g_TelescopeProgID);
				LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Telescope object");
				#Else
				m_Telescope = new ASCOM.DriverAccess.Telescope(g_TelescopeProgID);
				LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Telescope to get a Telescope object");
				#End If
				WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
				g_Stop = false;
			} catch (Exception ex) {
				LogMsg("", MessageLevel.msgDebug, "Attempt " + l_TryCount + " - exception thrown: " + ex.Message);
				if (l_TryCount == 3)
					throw;
				//Rethrow exception if on our third attempt
			}
			if (g_Stop)
				WaitFor(200);
		} while (!((!g_Stop)));
		//Exit if created OK
		LogMsg("Telescope:CreateDevice", MessageLevel.msgDebug, "Created telescope on attempt: " + l_TryCount.ToString);
	}
	private override bool Connected {
		get { Connected = m_Telescope.Connected; }
		set {
			m_Telescope.Connected = value;
			g_Stop = false;
		}
	}

	public override void PreRunCheck()
	{
		//Get into a consistent state
		if (g_InterfaceVersion > 1) {
			if (m_Telescope.AtPark) {
				if (m_CanUnpark) {
					m_Telescope.Unpark();
					LogMsg("Mount Safety", MessageLevel.msgInfo, "Scope is parked, so it has been unparked for testing");
				} else {
					LogMsg("Mount Safety", MessageLevel.msgError, "Scope reports that it is parked but CanUnPark is false - please manually unpark the scope");
					g_Stop = true;
				}
			} else {
				LogMsg("Mount Safety", MessageLevel.msgInfo, "Scope is not parked, continuing testing");
			}
		} else {
			LogMsg("Mount Safety", MessageLevel.msgInfo, "Skipping AtPark test as this method is not supported in interface V" + g_InterfaceVersion);
			try {
				if (m_CanUnpark) {
					m_Telescope.Unpark();
					LogMsg("Mount Safety", MessageLevel.msgOK, "Scope has been unparked for testing");
				} else {
					LogMsg("Mount Safety", MessageLevel.msgOK, "Scope reports that it cannot unpark, unparking skipped");
				}
			} catch (Exception ex) {
				LogMsg("Mount Safety", MessageLevel.msgError, "Driver threw an exception while unparking: " + ex.Message);
			}
		}

		if ((!TestStop()) & m_CanSetTracking) {
			m_Telescope.Tracking = true;
			LogMsg("Mount Safety", MessageLevel.msgInfo, "Scope tracking has been enabled");
		}

		if (!TestStop()) {
			LogMsg("TimeCheck", MessageLevel.msgInfo, "PC Time Zone:  " + g_Util.TimeZoneName + ", offset " + g_Util.TimeZoneOffset.ToString + " hours.");
			LogMsg("TimeCheck", MessageLevel.msgInfo, "PC UTCDate:    " + Format(g_Util.UTCDate, "dd-MMM-yyyy HH:mm:ss"));
			//v1.0.12.0 Added catch logic for any UTCDate issues
			try {
				LogMsg("TimeCheck", MessageLevel.msgDebug, "Mount UTCDate Unformatted: " + m_Telescope.UTCDate);
				LogMsg("TimeCheck", MessageLevel.msgInfo, "Mount UTCDate: " + Format(m_Telescope.UTCDate, "dd-MMM-yyyy HH:mm:ss"));
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: COM exception - UTCDate not implemented in this driver");
				} else {
					LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: COM Exception - " + ex.Message);
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: .NET exception - UTCDate not implemented in this driver");
			} catch (Exception ex) {
				LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: .NET Exception - " + ex.Message);
			}

		}
	}

	public override void ReadCanProperties()
	{
		TelescopeCanTest(CanType.tstCanFindHome, "CanFindHome");
		TelescopeCanTest(CanType.tstCanPark, "CanPark");
		TelescopeCanTest(CanType.tstCanPulseGuide, "CanPulseGuide");
		TelescopeCanTest(CanType.tstCanSetDeclinationRate, "CanSetDeclinationRate");
		TelescopeCanTest(CanType.tstCanSetGuideRates, "CanSetGuideRates");
		TelescopeCanTest(CanType.tstCanSetPark, "CanSetPark");
		TelescopeCanTest(CanType.tstCanSetPierSide, "CanSetPierSide");
		try {
			if ((m_Telescope.AlignmentMode != AlignmentModes.algGermanPolar) & m_CanSetPierside)
				LogMsg("CanSetPierSide", MessageLevel.msgIssue, "AlignmentMode is not GermanPolar but CanSetPierSide is true - contrary to ASCOM specification");
		} catch (Exception ex) {
		}
		TelescopeCanTest(CanType.tstCanSetRightAscensionRate, "CanSetRightAscensionRate");
		TelescopeCanTest(CanType.tstCanSetTracking, "CanSetTracking");
		TelescopeCanTest(CanType.tstCanSlew, "CanSlew");
		TelescopeCanTest(CanType.tstCanSlewAltAz, "CanSlewltAz");
		TelescopeCanTest(CanType.tstCanSlewAltAzAsync, "CanSlewAltAzAsync");
		TelescopeCanTest(CanType.tstCanSlewAsync, "CanSlewAsync");
		TelescopeCanTest(CanType.tstCanSync, "CanSync");
		TelescopeCanTest(CanType.tstCanSyncAltAz, "CanSyncAltAz");
		TelescopeCanTest(CanType.tstCanUnPark, "CanUnPark");
		if (m_CanUnpark & (!m_CanPark))
			LogMsg("CanUnPark", MessageLevel.msgIssue, "CanUnPark is true but CanPark is false - this does not comply with ASCOM specification");

	}
	private override void CheckProperties()
	{
		bool l_OriginalTrackingState;
		ASCOM.DeviceInterface.DriveRates l_DriveRate;
		double l_TimeDifference;

		//AlignmentMode - Optional
		try {
			m_AlignmentMode = m_Telescope.AlignmentMode;
			LogMsg("AlignmentMode", MessageLevel.msgOK, m_AlignmentMode.ToString);
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("AlignmentMode", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("AlignmentMode", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("AlignmentMode", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("AlignmentMode", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//Altitude - Optional
		try {
			m_CanReadAltitide = false;
			m_Altitude = m_Telescope.Altitude;
			m_CanReadAltitide = true;
			//Read successfully
			switch (m_Altitude) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
					LogMsg("Altitude", MessageLevel.msgWarning, "Altitude is <0.0 degrees: " + Format(m_Altitude, "0.00"));
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
90.0000001:
					LogMsg("Altitude", MessageLevel.msgWarning, "Altitude is >90.0 degrees: " + Format(m_Altitude, "0.00"));
				default:
					LogMsg("Altitude", MessageLevel.msgOK, Format(m_Altitude, "0.00"));
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("Altitude", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("Altitude", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("Altitude", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("Altitude", ex.Message, EX_NET + ex.Message + ex.GetType.ToString);
		}
		if (TestStop())
			return;

		//ApertureArea - Optional
		try {
			m_ApertureArea = m_Telescope.ApertureArea;
			switch (m_ApertureArea) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
0:
					LogMsg("ApertureArea", MessageLevel.msgWarning, "ApertureArea is < 0.0 : " + m_ApertureArea.ToString);
					break; // TODO: might not be correct. Was : Exit Select

				case 0.0:
					LogMsg("ApertureArea", MessageLevel.msgInfo, "ApertureArea is 0.0");
				default:
					LogMsg("ApertureArea", MessageLevel.msgOK, m_ApertureArea.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("ApertureArea", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("ApertureArea", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("ApertureArea", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("ApertureArea", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//ApertureDiameter - Optional
		try {
			m_ApertureDiameter = m_Telescope.ApertureDiameter;
			switch (m_ApertureDiameter) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
					LogMsg("ApertureDiameter", MessageLevel.msgWarning, "ApertureDiameter is < 0.0 : " + m_ApertureDiameter.ToString);
					break; // TODO: might not be correct. Was : Exit Select

				case 0.0:
					LogMsg("ApertureDiameter", MessageLevel.msgInfo, "ApertureDiameter is 0.0");
				default:
					LogMsg("ApertureDiameter", MessageLevel.msgOK, m_ApertureDiameter.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("ApertureDiameter", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("ApertureDiameter", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("ApertureDiameter", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("ApertureDiameter", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//AtHome - Required
		if (g_InterfaceVersion > 1) {
			try {
				m_AtHome = m_Telescope.AtHome;
				LogMsg("AtHome", MessageLevel.msgOK, m_AtHome.ToString);
			} catch (COMException ex) {
				LogMsg("AtHome", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("AtHome", MessageLevel.msgError, EX_NET + ex.Message);
			}
		} else {
			LogMsg("AtHome", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		if (TestStop())
			return;

		//AtPark - Required
		if (g_InterfaceVersion > 1) {
			try {
				m_AtPark = m_Telescope.AtPark;
				LogMsg("AtPark", MessageLevel.msgOK, m_AtPark.ToString);
			} catch (COMException ex) {
				LogMsg("AtPark", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("AtPark", MessageLevel.msgError, EX_NET + ex.Message);
			}
		} else {
			LogMsg("AtPark", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		if (TestStop())
			return;

		//Azimuth - Optional
		try {
			m_CanReadAzimuth = false;
			m_Azimuth = m_Telescope.Azimuth;
			m_CanReadAzimuth = true;
			// Read successfully
			switch (m_Azimuth) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
					LogMsg("Azimuth", MessageLevel.msgWarning, "Azimuth is <0.0 degrees: " + Format(m_Azimuth, "0.00"));
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
360.0000000001:
					LogMsg("Azimuth", MessageLevel.msgWarning, "Azimuth is >360.0 degrees: " + Format(m_Azimuth, "0.00"));
				default:
					LogMsg("Azimuth", MessageLevel.msgOK, Format(m_Azimuth, "0.00"));
			}

		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("Azimuth", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("Azimuth", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		//v1.0.12.0 Next two lines added to catch permitted exception
		} catch (PropertyNotImplementedException ex) {
			LogMsg("Azimuth", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("Azimuth", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//Declination - Required
		try {
			m_Declination = m_Telescope.Declination;
			switch (m_Declination) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
-90.0:
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
90.0:
					LogMsg("Declination", MessageLevel.msgWarning, "Declination is <-90 or >90 degrees: " + g_Util.DegreesToDMS(m_Declination, ":", ":", ""));
				default:
					LogMsg("Declination", MessageLevel.msgOK, g_Util.DegreesToDMS(m_Declination, ":", ":", ""));
			}
		} catch (COMException ex) {
			LogMsg("Declination", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("Declination", MessageLevel.msgError, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//DeclinationRate Read - Optional
		try {
			m_DeclinationRate = m_Telescope.DeclinationRate;
			//Read has been successful
			//Any value is acceptable
			if (m_CanSetDeclinationRate) {
				switch (m_DeclinationRate) {
					case  // ERROR: Case labels with binary operators are unsupported : GreaterThanOrEqual
0.0:
						LogMsg("DeclinationRate Read", MessageLevel.msgOK, Format(m_DeclinationRate, "0.00"));
					default:
						LogMsg("DeclinationRate Read", MessageLevel.msgWarning, "Negative DeclinatioRate: " + Format(m_DeclinationRate, "0.00"));
				}
			//Only zero is acceptable
			} else {
				switch (m_DeclinationRate) {
					case 0.0:
						LogMsg("DeclinationRate Read", MessageLevel.msgOK, Format(m_DeclinationRate, "0.00"));
					default:
						LogMsg("DeclinationRate Read", MessageLevel.msgIssue, "DeclinationRate is non zero when CanSetDeclinationRate is False " + Format(m_DeclinationRate, "0.00"));
				}
			}
		//v1.0.12.0 catch revised because its mandatory that a number is returned even if its zero
		} catch (COMException ex) {
			//            If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
			//LogMsg("DeclinationRate Read", MessageLevel.msgOK, NOT_IMP_COM)
			//Else
			if (!m_CanSetDeclinationRate)
				LogMsg("DeclinationRate Read", MessageLevel.msgInfo, "DeclinationRate must return 0 if CanSetDeclinationRate is false.");
			LogMsg("DeclinationRate Read", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		//End If
		//Catch ex As PropertyNotImplementedException
		//    LogMsg("DeclinationRate Read", MessageLevel.msgOK, NOT_IMP_NET)
		} catch (Exception ex) {
			//   If ex.InnerException.GetType.ToString = "ASCOM.PropertyNotImplementedException" Then
			// LogMsg("DeclinationRate Read", MessageLevel.msgOK, NOT_IMP_NET)
			// Else
			if (!m_CanSetDeclinationRate)
				LogMsg("DeclinationRate Read", MessageLevel.msgInfo, "DeclinationRate must return 0 if CanSetDeclinationRate is false.");
			LogMsg("DeclinationRate Read", MessageLevel.msgError, EX_NET + ex.Message);
			//End If
		}
		if (TestStop())
			return;

		//DeclinationRate Write - Optional
		if (g_InterfaceVersion > 1) {
			//Any value is acceptable
			if (m_CanSetDeclinationRate) {
				try {
					m_Telescope.DeclinationRate = 0.0;
					//Set to a harmless value
					LogMsg("DeclinationRate Write", MessageLevel.msgOK, Format(m_DeclinationRate, "0.00"));
				} catch (COMException ex) {
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("DeclinationRate Write", MessageLevel.msgIssue, "COM CanSetDeclinationRate is True but a \"Not Implemented\" error was raised");
					} else {
						ExTest("DeclinationRate Write", ex.Message, "COM CanSetDeclinationRate is True but: " + EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				} catch (PropertyNotImplementedException ex) {
					LogMsg("DeclinationRate Write", MessageLevel.msgOK, ".NET CanSetDeclinationRate is True but a \"Not Implemented\" error was raised");
				} catch (Exception ex) {
					ExTest("DeclinationRate Write", ex.Message, "CanSetDeclinationRate is True but: " + EX_NET + ex.Message);
				}
			//Should generate an error
			} else {
				try {
					m_Telescope.DeclinationRate = 0.0;
					//Set to a harmless value
					LogMsg("DeclinationRate", MessageLevel.msgIssue, "CanSetDeclinationRate is False but setting DeclinationRate did not generate an error");
				} catch (COMException ex) {
					//Not implemented error so OK
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("DeclinationRate Write", MessageLevel.msgOK, NOT_IMP_COM);
					} else {
						ExTest("DeclinationRate Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				} catch (PropertyNotImplementedException ex) {
					LogMsg("DeclinationRate Write", MessageLevel.msgOK, NOT_IMP_NET);
				} catch (Exception ex) {
					ExTest("DeclinationRate Write", ex.Message, EX_NET + ex.ToString);
				}
			}
		} else {
			try {
				m_Telescope.DeclinationRate = 0.0;
				//Set to a harmless value
				LogMsg("DeclinationRate Write", MessageLevel.msgOK, Format(m_DeclinationRate, "0.00"));
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("DeclinationRate Write", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("DeclinationRate Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("DeclinationRate Write", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("DeclinationRate Write", ex.Message, EX_NET + ex.Message);
			}
		}

		if (TestStop())
			return;

		//DoesRefraction Read - Optional
		if (g_InterfaceVersion > 1) {
			try {
				m_DoesRefraction = m_Telescope.DoesRefraction;
				LogMsg("DoesRefraction Read", MessageLevel.msgOK, m_DoesRefraction.ToString);
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("DoesRefraction Read", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("DoesRefraction Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("DoesRefraction Read", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("DoesRefraction Read", ex.Message, EX_NET + ex.Message);
			}
			if (TestStop())
				return;
		} else {
			LogMsg("DoesRefraction Read", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		//DoesRefraction Write - Optional
		if (g_InterfaceVersion > 1) {
			//Try opposite value
			if (m_DoesRefraction) {
				try {
					m_Telescope.DoesRefraction = false;
					LogMsg("DoesRefraction Write", MessageLevel.msgOK, "Can set DoesRefraction to False");
				} catch (COMException ex) {
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("DoesRefraction Write", MessageLevel.msgOK, NOT_IMP_COM);
					} else {
						ExTest("DoesRefraction Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				} catch (PropertyNotImplementedException ex) {
					LogMsg("DoesRefraction Write", MessageLevel.msgOK, NOT_IMP_NET);
				} catch (Exception ex) {
					ExTest("DoesRefraction Write", ex.Message, EX_NET + ex.Message);
				}
			//Try other opposite value
			} else {
				try {
					m_Telescope.DoesRefraction = true;
					LogMsg("DoesRefraction Write", MessageLevel.msgOK, "Can set DoesRefraction to True");
				} catch (COMException ex) {
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("DoesRefraction Write", MessageLevel.msgOK, NOT_IMP_COM);
					} else {
						ExTest("DoesRefraction Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				} catch (PropertyNotImplementedException ex) {
					LogMsg("DoesRefraction Write", MessageLevel.msgOK, NOT_IMP_NET);
				} catch (Exception ex) {
					ExTest("DoesRefraction Write", ex.Message, EX_NET + ex.Message);
				}
			}
		} else {
			LogMsg("DoesRefraction Write", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		if (TestStop())
			return;

		//EquatorialSystem - Required
		if (g_InterfaceVersion > 1) {
			try {
				m_EquatorialSystem = m_Telescope.EquatorialSystem;
				LogMsg("EquatorialSystem", MessageLevel.msgOK, m_EquatorialSystem.ToString);
			} catch (COMException ex) {
				LogMsg("EquatorialSystem", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("EquatorialSystem", MessageLevel.msgError, EX_NET + ex.Message);
			}
		} else {
			LogMsg("EquatorialSystem", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		if (TestStop())
			return;

		//FocalLength - Optional
		try {
			m_FocalLength = m_Telescope.FocalLength;
			switch (m_FocalLength) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
					LogMsg("FocalLength", MessageLevel.msgWarning, "FocalLength is <0.0 : " + m_FocalLength.ToString);
					break; // TODO: might not be correct. Was : Exit Select

				case 0.0:
					LogMsg("FocalLength", MessageLevel.msgInfo, "FocalLength is 0.0");
				default:
					LogMsg("FocalLength", MessageLevel.msgOK, m_FocalLength.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("FocalLength", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("FocalLength", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("FocalLength", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("FocalLength", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//GuideRateDeclination - Optional
		if (g_InterfaceVersion > 1) {
			if (m_CanSetGuideRates) {
				try {
					m_GuideRateDeclination = m_Telescope.GuideRateDeclination;
					//Read guiderateDEC
					switch (m_GuideRateDeclination) {
						case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
							LogMsg("GuideRateDeclination Read", MessageLevel.msgWarning, "GuideRateDeclination is < 0.0 " + Format(m_GuideRateDeclination, "0.00"));
						default:
							LogMsg("GuideRateDeclination Read", MessageLevel.msgOK, Format(m_GuideRateDeclination, "0.00"));
					}
				//Read failed
				} catch (COMException ex) {
					LogMsg("GuideRateDeclination Read", MessageLevel.msgError, "CanSetGuideRates is True but " + EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				//Read failed
				} catch (Exception ex) {
					LogMsg("GuideRateDeclination Read", MessageLevel.msgError, "CanSetGuideRates is True but " + EX_NET + ex.Message);
				}
				//Read OK so now try to write
				try {
					m_Telescope.GuideRateDeclination = m_GuideRateDeclination;
					LogMsg("GuideRateDeclination Write", MessageLevel.msgOK, "Can write Declination Guide Rate OK");
				//Write failed
				} catch (COMException ex) {
					LogMsg("GuideRateDeclination Write", MessageLevel.msgError, "CanSetGuideRates is True but " + EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				//Write failed
				} catch (Exception ex) {
					LogMsg("GuideRateDeclination Write", MessageLevel.msgError, "CanSetGuideRates is True but " + EX_NET + ex.Message);
				}
			} else {
				//Cannot set guide rates so read may generate an error
				try {
					m_GuideRateDeclination = m_Telescope.GuideRateDeclination;
					switch (m_GuideRateDeclination) {
						case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
							LogMsg("GuideRateDeclination Read", MessageLevel.msgWarning, "GuideRateDeclination is < 0.0 " + Format(m_GuideRateDeclination, "0.00"));
						default:
							LogMsg("GuideRateDeclination Read", MessageLevel.msgOK, Format(m_GuideRateDeclination, "0.00"));
					}
				} catch (COMException ex) {
					//Read not implemented
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("GuideRateDeclination Read", MessageLevel.msgOK, NOT_IMP_COM);
					//Some other error
					} else {
						ExTest("GuideRateDeclination Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				//Read not implemented
				} catch (PropertyNotImplementedException ex) {
					LogMsg("GuideRateDeclination Read", MessageLevel.msgOK, NOT_IMP_NET);
				//Some other error
				} catch (Exception ex) {
					ExTest("GuideRateDeclination Read", ex.Message, EX_NET + ex.Message);
				}
				//Write should definitely raise an error
				try {
					m_Telescope.GuideRateDeclination = m_GuideRateDeclination;
					LogMsg("GuideRateDeclination Write", MessageLevel.msgIssue, "CanSetGuideRates is false but no exception generated; value returned: " + Format(m_GuideRateDeclination, "0.00"));
				} catch (COMException ex) {
					//Write not implemented
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("GuideRateDeclination Write", MessageLevel.msgOK, NOT_IMP_COM);
					//Some other error so OK
					} else {
						ExTest("GuideRateDeclination Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				//Write not implemented
				} catch (PropertyNotImplementedException ex) {
					LogMsg("GuideRateDeclination Write", MessageLevel.msgOK, NOT_IMP_NET);
				//Some other error so OK
				} catch (Exception ex) {
					ExTest("GuideRateDeclination Write", ex.Message, EX_NET + ex.Message);
				}
			}
		} else {
			LogMsg("GuideRateDeclination", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		if (TestStop())
			return;

		//GuideRateRightAscension - Optional
		if (g_InterfaceVersion > 1) {
			if (m_CanSetGuideRates) {
				try {
					m_GuideRateRightAscension = m_Telescope.GuideRateRightAscension;
					//Read guiderateRA
					switch (m_GuideRateDeclination) {
						case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
							LogMsg("GuideRateRightAscension Read", MessageLevel.msgWarning, "GuideRateRightAscension is < 0.0 " + Format(m_GuideRateRightAscension, "0.00"));
						default:
							LogMsg("GuideRateRightAscension Read", MessageLevel.msgOK, Format(m_GuideRateRightAscension, "0.00"));
					}
				//Read failed
				} catch (COMException ex) {
					LogMsg("GuideRateRightAscension Read", MessageLevel.msgError, "CanSetGuideRates is True but " + EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				//Read failed
				} catch (Exception ex) {
					LogMsg("GuideRateRightAscension Read", MessageLevel.msgError, "CanSetGuideRates is True but " + EX_NET + ex.Message);
				}
				//Read OK so now try to write
				try {
					m_Telescope.GuideRateRightAscension = m_GuideRateRightAscension;
					LogMsg("GuideRateRightAscension Write", MessageLevel.msgOK, "Can set RightAscension Guide OK");
				//Write failed
				} catch (COMException ex) {
					LogMsg("GuideRateRightAscension Write", MessageLevel.msgError, "CanSetGuideRates is True but " + EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				//Write failed
				} catch (Exception ex) {
					LogMsg("GuideRateRightAscension Write", MessageLevel.msgError, "CanSetGuideRates is True but " + EX_NET + ex.Message);
				}
			} else {
				//Cannot set guide rates so read may generate an error
				try {
					m_GuideRateRightAscension = m_Telescope.GuideRateRightAscension;
					//Read guiderateRA
					switch (m_GuideRateDeclination) {
						case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
							LogMsg("GuideRateRightAscension Read", MessageLevel.msgWarning, "GuideRateRightAscension is < 0.0 " + Format(m_GuideRateRightAscension, "0.00"));
						default:
							LogMsg("GuideRateRightAscension Read", MessageLevel.msgOK, Format(m_GuideRateRightAscension, "0.00"));
					}
				} catch (COMException ex) {
					//Read not implemented
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("GuideRateRightAscension Read", MessageLevel.msgOK, NOT_IMP_COM);
					//Some other error
					} else {
						ExTest("GuideRateRightAscension Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				//Read not implemented
				} catch (PropertyNotImplementedException ex) {
					LogMsg("GuideRateRightAscension Read", MessageLevel.msgOK, NOT_IMP_NET);
				//Some other error
				} catch (Exception ex) {
					ExTest("GuideRateRightAscension Read", ex.Message, EX_NET + ex.Message);
				}
				//Write should definitely raise an error
				try {
					m_Telescope.GuideRateRightAscension = m_GuideRateRightAscension;
					LogMsg("GuideRateRightAscension Write", MessageLevel.msgIssue, "CanSetGuideRates is false but no exception generated; value returned: " + Format(m_GuideRateRightAscension, "0.00"));
				} catch (COMException ex) {
					//Write not implemented
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("GuideRateRightAscension Write", MessageLevel.msgOK, NOT_IMP_COM);
					//Some other error so OK
					} else {
						ExTest("GuideRateRightAscension Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				//Write not implemented
				} catch (PropertyNotImplementedException ex) {
					LogMsg("GuideRateRightAscension Write", MessageLevel.msgOK, NOT_IMP_NET);
				//Some other error so OK
				} catch (Exception ex) {
					ExTest("GuideRateRightAscension Write", ex.Message, EX_NET + ex.Message);
				}
			}
		} else {
			LogMsg("GuideRateRightAscension", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		if (TestStop())
			return;

		//IsPulseGuideing - Optional
		if (g_InterfaceVersion > 1) {
			//Can pulseguide so test if we can successfully read ispulseguiding
			if (m_CanPulseGuide) {
				try {
					m_IsPulseGuiding = m_Telescope.IsPulseGuiding;
					LogMsg("IsPulseGuideing", MessageLevel.msgOK, m_IsPulseGuiding.ToString);
				//v1.0.12.0 Catch statemennts revised as not supprted exception is not permitted if CanPulseGuide is true
				//Catch ex As COMException
				//If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
				//LogMsg("IsPulseGuideing", MessageLevel.msgOK, NOT_IMP_COM)
				//Else
				//ExTest("IsPulseGuideing", ex.Message, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
				//End If
				//Catch ex As PropertyNotImplementedException
				//LogMsg("IsPulseGuideing", MessageLevel.msgOK, NOT_IMP_NET)
				//Catch ex As Exception
				//ExTest("IsPulseGuideing", ex.Message, EX_NET & ex.Message)
				//Read failed
				} catch (COMException ex) {
					LogMsg("IsPulseGuideing", MessageLevel.msgError, "CanPulseGuide is True but " + EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				//Read failed
				} catch (Exception ex) {
					LogMsg("IsPulseGuideing", MessageLevel.msgError, "CanPulseGuide is True but " + EX_NET + ex.Message);
				}
			//Cant pulseguide so confirm that an error is raised
			} else {
				try {
					m_IsPulseGuiding = m_Telescope.IsPulseGuiding;
					LogMsg("IsPulseGuideing", MessageLevel.msgIssue, "CanPulseGuide is False but no error was raised on calling IsPulseGuideing");
				} catch (COMException ex) {
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("IsPulseGuideing", MessageLevel.msgOK, NOT_IMP_COM);
					} else {
						ExTest("IsPulseGuideing", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				} catch (PropertyNotImplementedException ex) {
					LogMsg("IsPulseGuideing", MessageLevel.msgOK, NOT_IMP_NET);
				} catch (Exception ex) {
					ExTest("IsPulseGuideing", ex.Message, EX_NET + ex.Message);
				}
			}
		} else {
			LogMsg("IsPulseGuideing", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		if (TestStop())
			return;

		//RightAscension - Required
		try {
			m_RightAscension = m_Telescope.RightAscension;
			switch (m_RightAscension) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThanOrEqual
24.0:
					LogMsg("RightAscension", MessageLevel.msgWarning, "RightAscension is <0 or >=24 hours: " + m_RightAscension + " " + g_Util.HoursToHMS(m_RightAscension, ":", ":", ""));
				default:
					LogMsg("RightAscension", MessageLevel.msgOK, g_Util.HoursToHMS(m_RightAscension, ":", ":", ""));
			}
		} catch (COMException ex) {
			LogMsg("RightAscension", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("RightAscension", MessageLevel.msgError, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//RightAscensionRate Read - Optional - Read must always return 0.0
		try {
			m_RightAscensionRate = m_Telescope.RightAscensionRate;
			//Read has been successful
			//Any value is acceptable
			if (m_CanSetRightAscensionRate) {
				switch (m_DeclinationRate) {
					case  // ERROR: Case labels with binary operators are unsupported : GreaterThanOrEqual
0.0:
						LogMsg("RightAscensionRate Read", MessageLevel.msgOK, Format(m_RightAscensionRate, "0.00"));
					default:
						LogMsg("RightAscensionRate Read", MessageLevel.msgWarning, "Negative RightAscensionRate: " + Format(m_RightAscensionRate, "0.00"));
				}
			//Only zero is acceptable
			} else {
				switch (m_RightAscensionRate) {
					case 0.0:
						LogMsg("RightAscensionRate Read", MessageLevel.msgOK, Format(m_RightAscensionRate, "0.00"));
					default:
						LogMsg("RightAscensionRate Read", MessageLevel.msgIssue, "RightAscensionRate is non zero when CanSetRightAscensionRate is False " + Format(m_DeclinationRate, "0.00"));
				}
			}
		//v1.0.12.0 Catch revised to properly throw an error, property not supported is not allowed since read must always return 0
		//Catch ex As COMException
		//If ex.ErrorCode = g_ExNotImplemented Or ex.ErrorCode = ErrorCodes.NotImplemented Then
		//LogMsg("RightAscensionRate Read", MessageLevel.msgOK, NOT_IMP_COM)
		//Else
		//ExTest("RightAscensionRate Read", ex.Message, EX_COM & ex.Message & " " & Hex(ex.ErrorCode))
		//End If
		//Catch ex As PropertyNotImplementedException
		//LogMsg("RightAscensionRate Read", MessageLevel.msgOK, NOT_IMP_NET)
		//Catch ex As Exception
		//ExTest("RightAscensionRate Read", ex.Message, EX_NET & ex.Message)
		} catch (COMException ex) {
			if (!m_CanSetRightAscensionRate)
				LogMsg("RightAscensionRate Read", MessageLevel.msgInfo, "RightAscensionRate must return 0 if CanSetRightAscensionRate is false.");
			LogMsg("RightAscensionRate Read", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			if (!m_CanSetRightAscensionRate)
				LogMsg("RightAscensionRate Read", MessageLevel.msgInfo, "RightAscensionRate must return 0 if CanSetRightAscensionRate is false.");
			LogMsg("RightAscensionRate Read", MessageLevel.msgError, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//RightAscensionRate Write - Optional
		if (g_InterfaceVersion > 1) {
			//Any value is acceptable and should not generate an error
			if (m_CanSetRightAscensionRate) {
				try {
					m_Telescope.RightAscensionRate = 0.0;
					//Set to a harmless value
					LogMsg("RightAscensionRate Write", MessageLevel.msgOK, Format(m_RightAscensionRate, "0.00"));
				} catch (COMException ex) {
					LogMsg("RightAscensionRate Write", MessageLevel.msgError, "CanSetRightAscensionRate is True but: " + EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (Exception ex) {
					LogMsg("RightAscensionRate Write", MessageLevel.msgError, "CanSetRightAscensionRate is True but: " + EX_NET + ex.Message);
				}
			//Should generate an error
			} else {
				try {
					m_Telescope.RightAscensionRate = 0.0;
					//Set to a harmless value
					LogMsg("RightAscensionRateWrite", MessageLevel.msgIssue, "CanSetRightAscensionRate is False but setting RightAscensionRate did not generate an error");
				} catch (COMException ex) {
					//Not implemented error so OK
					if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
						LogMsg("RightAscensionRate Write", MessageLevel.msgOK, NOT_IMP_COM);
					} else {
						ExTest("RightAscensionRate Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					}
				} catch (PropertyNotImplementedException ex) {
					LogMsg("RightAscensionRate Write", MessageLevel.msgOK, NOT_IMP_NET);
				} catch (Exception ex) {
					ExTest("RightAscensionRate Write", ex.Message, EX_NET + ex.Message);
				}
			}
		//V1 has no Can property so just test anyway, it is optional
		} else {
			try {
				m_Telescope.RightAscensionRate = 0.0;
				//Set to a harmless value
				LogMsg("RightAscensionRate Write", MessageLevel.msgOK, Format(m_RightAscensionRate, "0.00"));
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("RightAscensionRate Write", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("RightAscensionRate Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("RightAscensionRate Write", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("RightAscensionRate Write", ex.Message, EX_NET + ex.Message);
			}
		}

		if (TestStop())
			return;

		//SiteElevation Read - Optional
		try {
			m_SiteElevation = m_Telescope.SiteElevation;
			switch (m_SiteElevation) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
-300.0:
					LogMsg("SiteElevation Read", MessageLevel.msgIssue, "SiteElevation is <-300m");
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
10000.0:
					LogMsg("SiteElevation Read", MessageLevel.msgIssue, "SiteElevation is >10,000m");
				default:
					LogMsg("SiteElevation Read", MessageLevel.msgOK, m_SiteElevation.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("SiteElevation Read", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("SiteElevation Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteElevation Read", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("SiteElevation Read", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//SiteElevation Write - Optional
		try {
			m_Telescope.SiteElevation = -301.0;
			LogMsg("SiteElevation Write", MessageLevel.msgIssue, "No error generated on set site elevation < -300m");
		} catch (InvalidValueException ex) {
			LogMsg("SiteElevation Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set site elevation < -300m");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SiteElevation Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteElevation Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set site elevation < -300m");
				default:
					ExTest("SiteElevation Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteElevation Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteElevation Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set site elevation < -300m");
				default:
					ExTest("SiteElevation Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("SiteElevation Write", ex.Message, EX_NET + ex.Message);
		}
		try {
			m_Telescope.SiteElevation = 10001.0;
			LogMsg("SiteElevation Write", MessageLevel.msgIssue, "No error generated on set site elevation > 10,000m");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SiteElevation Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteElevation Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set site elevation > 10,000m");
				default:
					ExTest("SiteElevation Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteElevation Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteElevation Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set site elevation > 10,000m");
				default:
					ExTest("SiteElevation Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("SiteElevation Write", ex.Message, EX_NET + ex.Message);
		}
		try {
			if ((m_SiteElevation < -300.0) | (m_SiteElevation > 10000.0))
				m_SiteElevation = 1000;
			m_Telescope.SiteElevation = m_SiteElevation;
			//Restore original value
			LogMsg("SiteElevation Write", MessageLevel.msgOK, "Legal value " + m_SiteElevation.ToString + "m written successfully");
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("SiteElevation Write", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("SiteElevation Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteElevation Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("SiteElevation Write", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//SiteLatitude Read - Optional
		try {
			m_SiteLatitude = m_Telescope.SiteLatitude;
			switch (m_SiteLatitude) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
-90.0:
					LogMsg("SiteLatitude Read", MessageLevel.msgWarning, "SiteLatitude is < -90 degrees");
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
90.0:
					LogMsg("SiteLatitude Read", MessageLevel.msgWarning, "SiteLatitude is > 90 degrees");
				default:
					LogMsg("SiteLatitude Read", MessageLevel.msgOK, g_Util.DegreesToDMS(m_SiteLatitude, ":", ":", ""));
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("SiteLatitude Read", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("SiteLatitude Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteLatitude Read", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("SiteLatitude Read", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//SiteLatitude Write - Optional
		try {
			m_Telescope.SiteLatitude = -91.0;
			LogMsg("SiteLatitude Write", MessageLevel.msgIssue, "No error generated on set site latitude < -90 degrees");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SiteLatitude Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteLatitude Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set site latitude < -90 degrees");
				default:
					ExTest("SiteLatitude Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteLatitude Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteLatitude Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set latitude < -90 degrees");
				default:
					ExTest("SiteLatitude Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("SiteLatitude Write", ex.Message, EX_NET + ex.Message);
		}
		try {
			m_Telescope.SiteLatitude = 91.0;
			LogMsg("SiteLatitude Write", MessageLevel.msgIssue, "No error generated on set site latitude > 90 degrees");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SiteLatitude Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteLatitude Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set site latitude > 90 degrees");
				default:
					ExTest("SiteLatitude Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteLatitude Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteLatitude Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set site latitude > 90 degrees");
				default:
					ExTest("SiteLatitude Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("SiteLatitude Write", ex.Message, EX_NET + ex.Message);
		}
		try {
			if ((m_SiteLatitude < -90.0) | (m_SiteLatitude > 90.0))
				m_SiteLatitude = 45.0;
			m_Telescope.SiteLatitude = m_SiteLatitude;
			//Restore original value
			LogMsg("SiteLatitude Write", MessageLevel.msgOK, "Legal value " + g_Util.DegreesToDMS(m_SiteLatitude, ":", ":", "") + " degrees written successfully");
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("SiteLatitude Write", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("SiteLatitude Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteLatitude Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("SiteLatitude Write", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//SiteLongitude Read - Optional
		try {
			m_SiteLongitude = m_Telescope.SiteLongitude;
			switch (m_SiteLongitude) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
-180.0:
					LogMsg("SiteLongitude Read", MessageLevel.msgWarning, "SiteLongitude is < -180 degrees");
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
180.0:
					LogMsg("SiteLongitude Read", MessageLevel.msgWarning, "SiteLongitude is > 180 degrees");
				default:
					LogMsg("SiteLongitude Read", MessageLevel.msgOK, g_Util.DegreesToDMS(m_SiteLongitude, ":", ":", ""));
			}
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SiteLongitude Read", MessageLevel.msgOK, NOT_IMP_COM);
				default:
					ExTest("SiteLongitude Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteLongitude Read", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("SiteLongitude Read", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//SiteLongitude Write - Optional
		try {
			m_Telescope.SiteLongitude = -181.0;
			LogMsg("SiteLongitude Write", MessageLevel.msgIssue, "No error generated on set site Longitude < -180 degrees");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SiteLongitude Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteLongitude Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set site longitude < -180 degrees");
				default:
					ExTest("SiteLongitude Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteLongitude Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteLongitude Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set site longitude < - 180 degrees");
				default:
					ExTest("SiteLongitude Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("SiteLongitude Write", ex.Message, EX_NET + ex.Message);
		}

		try {
			m_Telescope.SiteLongitude = 181.0;
			LogMsg("SiteLongitude Write", MessageLevel.msgIssue, "No error generated on set site longitude > 180 degrees");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SiteLongitude Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteLongitude Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set site longitude > 180 degrees");
				default:
					ExTest("SiteLongitude Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteLongitude Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SiteLongitude Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set site longitude > 180 degrees");
				default:
					ExTest("SiteLongitude Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("SiteLongitude Write", ex.Message, "Error generated OK on set site longitude > 180 degrees");
		}

		try {
			if ((m_SiteLongitude < -180.0) | (m_SiteLongitude > 180.0))
				m_SiteLongitude = 60.0;
			m_Telescope.SiteLongitude = m_SiteLongitude;
			//Restore original value
			LogMsg("SiteLongitude Write", MessageLevel.msgOK, "Legal value " + g_Util.DegreesToDMS(m_SiteLongitude, ":", ":", "") + " degrees written successfully");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SiteLongitude Write", MessageLevel.msgOK, NOT_IMP_COM);
				default:
					ExTest("SiteLongitude Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SiteLongitude Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("SiteLongitude Write", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//Slewing - Optional
		try {
			m_Slewing = m_Telescope.Slewing;
			switch (m_Slewing) {
				case false:
					LogMsg("Slewing", MessageLevel.msgOK, m_Slewing.ToString);
				case true:
					LogMsg("Slewing", MessageLevel.msgIssue, "Slewing should be false and it reads as " + m_Slewing.ToString);
			}
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("Slewing", MessageLevel.msgOK, NOT_IMP_COM);
				default:
					ExTest("Slewing", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("Slewing", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("Slewing", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//SlewSettleTime Read - Optional
		try {
			m_SlewSettleTime = m_Telescope.SlewSettleTime;
			switch (m_SlewSettleTime) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
(short)0:
					LogMsg("SlewSettleTime Read", MessageLevel.msgWarning, "SlewSettleTime is < 0 seconds");
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
(short)30.0:
					LogMsg("SlewSettleTime Read", MessageLevel.msgInfo, "SlewSettleTime is > 30 seconds");
				default:
					LogMsg("SlewSettleTime Read", MessageLevel.msgOK, m_SlewSettleTime.ToString);
			}
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SlewSettleTime Read", MessageLevel.msgOK, NOT_IMP_COM);
				default:
					ExTest("SlewSettleTime Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SlewSettleTime Read", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("SlewSettleTime Read", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//SlewSettleTime Write - Optional
		try {
			m_Telescope.SlewSettleTime = (short)-1;
			LogMsg("SlewSettleTime Write", MessageLevel.msgIssue, "No error generated on set SlewSettleTime < 0 seconds");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SlewSettleTime Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SlewSettleTime Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set slew settle time < 0");
				default:
					ExTest("SlewSettleTime Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SlewSettleTime Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("SlewSettleTime Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set slew settle time < 0");
				default:
					ExTest("SlewSettleTime Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("SlewSettleTime Write", ex.Message, EX_NET + ex.Message);
		}
		try {
			if ((m_SlewSettleTime < 0))
				m_SlewSettleTime = 0;
			m_Telescope.SlewSettleTime = m_SlewSettleTime;
			//Restore original value
			LogMsg("SlewSettleTime Write", MessageLevel.msgOK, "Legal value " + m_SlewSettleTime.ToString + " seconds written successfully");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("SlewSettleTime Write", MessageLevel.msgOK, NOT_IMP_COM);
				default:
					ExTest("SlewSettleTime Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("SlewSettleTime Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("SlewSettleTime Write", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//SideOfPier - Optional
		//v1.0.12.0 Changed false true logic to only flag OK when test is successful
		//m_CanReadSideOfPier = True 'Start out assuming we can read side of pier OK
		m_CanReadSideOfPier = false;
		//Start out assuming that we actually can't read side of pier so the performance test can be omitted

		//v1.0.11.12 Telescope interface definition was updated to remove the limitation that SideOfPier was only valid for German mounts
		//Consequently changes are made here to remove that selectivity.
		//Select Case m_AlignmentMode
		//   Case ASCOM.DeviceInterface.AlignmentModes.algGermanPolar 'SideOfPier is only valid for German mounts

		//SideOfPier Read
		if (g_InterfaceVersion > 1) {
			try {
				m_SideOfPier = m_Telescope.SideOfPier;
				LogMsg("SideOfPier Read", MessageLevel.msgOK, m_SideOfPier.ToString);
				m_CanReadSideOfPier = true;
				//Flag that it is OK to read sideofpier
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("SideOfPier Read", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("SideOfPier Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("SideOfPier Read", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("SideOfPier Read", ex.Message, EX_NET + ex.Message);
			}
		} else {
			LogMsg("SideOfPier Read", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		//SideOfPier Write
		//Moved to methods section as this really is a method rather than a property


		//SiderealTime - Required
		try {
			m_SiderealTimeScope = m_Telescope.SiderealTime;
			m_SiderealTimeASCOM = (18.697374558 + 24.0657098244191 * (g_Util.DateLocalToJulian(Now) - 2451545.0) + (m_SiteLongitude / 15.0)) % 24.0;
			switch (m_SiderealTimeScope) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
0.0:
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThanOrEqual
24.0:
					LogMsg("SiderealTime", MessageLevel.msgWarning, "SiderealTime is <0 or >=24 hours: " + g_Util.HoursToHMS(m_SiderealTimeScope, ":", ":", ""));
				default:
					//Valid time returned
					//Now do a sense check on the received value
					LogMsg("SiderealTime", MessageLevel.msgOK, g_Util.HoursToHMS(m_SiderealTimeScope, ":", ":", ""));
					l_TimeDifference = System.Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM);
					//Get time difference between scope and PC
					//Process edge cases where the two clocks are on either side of 0:0:0/24:0:0
					if ((m_SiderealTimeASCOM > 23.0 & m_SiderealTimeASCOM < 23.999) & (m_SiderealTimeScope > 0.0 & m_SiderealTimeScope < 1.0)) {
						l_TimeDifference = System.Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM + 24.0);
					}
					if ((m_SiderealTimeScope > 23.0 & m_SiderealTimeScope < 23.999) & (m_SiderealTimeASCOM > 0.0 & m_SiderealTimeASCOM < 1.0)) {
						l_TimeDifference = System.Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM - 24.0);
					}

					switch (l_TimeDifference) {
						case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / 3600.0):
							//1 seconds
							LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sideral times agree to better than 1 second, Scope: " + g_Util.HoursToHMS(m_SiderealTimeScope) + ", ASCOM: " + g_Util.HoursToHMS(m_SiderealTimeASCOM));
							break; // TODO: might not be correct. Was : Exit Select

						case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / 3600.0):
							//2 seconds
							LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sideral times agree to better than 2 seconds, Scope: " + g_Util.HoursToHMS(m_SiderealTimeScope) + ", ASCOM: " + g_Util.HoursToHMS(m_SiderealTimeASCOM));
							break; // TODO: might not be correct. Was : Exit Select

						case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(5.0 / 3600.0):
							//5 seconds
							LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sideral times agree to better than 5 seconds, Scope: " + g_Util.HoursToHMS(m_SiderealTimeScope) + ", ASCOM: " + g_Util.HoursToHMS(m_SiderealTimeASCOM));
							break; // TODO: might not be correct. Was : Exit Select

						case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / 60.0):
							//1 minute
							LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sideral times agree to better than 1 minute, Scope: " + g_Util.HoursToHMS(m_SiderealTimeScope) + ", ASCOM: " + g_Util.HoursToHMS(m_SiderealTimeASCOM));
							break; // TODO: might not be correct. Was : Exit Select

						case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(5.0 / 60.0):
							//5 minutes
							LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sideral times agree to better than 5 minutes, Scope: " + g_Util.HoursToHMS(m_SiderealTimeScope) + ", ASCOM: " + g_Util.HoursToHMS(m_SiderealTimeASCOM));
							break; // TODO: might not be correct. Was : Exit Select

						case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
0.5:
							//0.5 an hour
							LogMsg("SiderealTime", MessageLevel.msgInfo, "Scope and ASCOM sideral times are up to 0.5 hour different, Scope: " + g_Util.HoursToHMS(m_SiderealTimeScope) + ", ASCOM: " + g_Util.HoursToHMS(m_SiderealTimeASCOM));
							break; // TODO: might not be correct. Was : Exit Select

						case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
1.0:
							//1.0 an hour
							LogMsg("SiderealTime", MessageLevel.msgInfo, "Scope and ASCOM sideral times are up to 1.0 hour different, Scope: " + g_Util.HoursToHMS(m_SiderealTimeScope) + ", ASCOM: " + g_Util.HoursToHMS(m_SiderealTimeASCOM));
							break; // TODO: might not be correct. Was : Exit Select

						default:
							LogMsg("SiderealTime", MessageLevel.msgError, "Scope and ASCOM sideral times are more than 1 hour apart, Scope: " + g_Util.HoursToHMS(m_SiderealTimeScope) + ", ASCOM: " + g_Util.HoursToHMS(m_SiderealTimeASCOM));
							MsgBox("Following tests rely on correct sidereal time to calculate target RAs. The sidereal time returned by this driver is more than 1 hour from the expected value based on your computer clock and site longitude, so this program will end now to protect your mount from potential harm caused by slewing to an inappropriate location." + vbCrLf + vbCrLf + "Please check the longitude set by the driver and your PC clock (time, time zone and summer time) before checking the sideral time code in your driver or your mount. Thanks, Peter", MsgBoxStyle.Critical, "CONFORM - MOUNT PROTECTION");
							g_Stop = true;
							return;

					}
			}
		} catch (COMException ex) {
			LogMsg("SiderealTime", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("SiderealTime", MessageLevel.msgError, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//TargetDeclination Read - Optional
		//First read should fail!
		try {
			m_TargetDeclination = m_Telescope.TargetDeclination;
			LogMsg("TargetDeclination Read", MessageLevel.msgIssue, "Read before write should generate an error and didn't");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("TargetDeclination Read", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.ValueNotSet:
				case g_ExNotSet1:
				case g_ExNotSet2:
					LogMsg("TargetDeclination Read", MessageLevel.msgOK, "COM Not Set exception generated on read before write");
				default:
					ExTest("TargetDeclination Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("TargetDeclination Read", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (ASCOM.InvalidOperationException ex) {
			LogMsg("TargetDeclination Read", MessageLevel.msgOK, ".NET Not InvalidOperationException generated on read before write");
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.ValueNotSet:
				case g_ExNotSet1:
				case g_ExNotSet2:
					LogMsg("TargetDeclination Read", MessageLevel.msgOK, ".NET Not Set exception generated on read before write");
				default:
					ExTest("TargetDeclination Read", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (System.InvalidOperationException ex) {
			LogMsg("TargetDeclination Read", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException");
		} catch (Exception ex) {
			ExTest("TargetDeclination Read", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//TargetDeclination Write - Optional
		try {
			m_Telescope.TargetDeclination = -91.0;
			LogMsg("TargetDeclination Write", MessageLevel.msgIssue, "No error generated on set TargetDeclination < -90 degrees");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("TargetDeclination Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("TargetDeclination Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
				default:
					ExTest("TargetDeclination Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("TargetDeclination Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("TargetDeclination Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
				default:
					ExTest("TargetDeclination Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("TargetDeclination Write", ex.Message, EX_NET + ex.Message);
		}

		try {
			m_Telescope.TargetDeclination = 91.0;
			LogMsg("TargetDeclination Write", MessageLevel.msgIssue, "No error generated on set TargetDeclination > 90 degrees");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("TargetDeclination Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("TargetDeclination Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set TargetDeclination > 90 degrees");
				default:
					ExTest("TargetDeclination Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("TargetDeclination Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("TargetDeclination Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
				default:
					ExTest("TargetDeclination Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("TargetDeclination Write", ex.Message, EX_NET + ex.Message);
		}

		try {
			m_TargetDeclination = 1.0;
			m_Telescope.TargetDeclination = m_TargetDeclination;
			//Set a valid value
			try {
				switch (System.Math.Abs(m_Telescope.TargetDeclination - m_TargetDeclination)) {
					case 0.0:
						LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Legal value " + g_Util.DegreesToDMS(m_TargetDeclination, ":", ":") + " DD:MM:SS written successfully");
					case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / 3600.0):
						//1 seconds
						LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 1 second of the value set: " + g_Util.DegreesToDMS(m_TargetDeclination, ":", ":"));
						break; // TODO: might not be correct. Was : Exit Select

					case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / 3600.0):
						//2 seconds
						LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 2 seconds of the value set: " + g_Util.DegreesToDMS(m_TargetDeclination, ":", ":"));
						break; // TODO: might not be correct. Was : Exit Select

					case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(5.0 / 3600.0):
						//5 seconds
						LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 5 seconds of the value set: " + g_Util.DegreesToDMS(m_TargetDeclination, ":", ":"));
						break; // TODO: might not be correct. Was : Exit Select

					default:
						LogMsg("TargetDeclination Write", MessageLevel.msgInfo, "Target Declination: " + g_Util.DegreesToDMS(m_TargetDeclination, ":", ":"));
				}
			} catch (COMException ex) {
				switch (ex.ErrorCode) {
					case g_ExNotImplemented:
					case ErrorCodes.NotImplemented:
						LogMsg("TargetDeclination Write", MessageLevel.msgOK, NOT_IMP_COM);
					default:
						ExTest("TargetDeclination Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("TargetDeclination Write", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("TargetDeclination Write", ex.Message, EX_NET + ex.Message);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("TargetDeclination Write", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("TargetDeclination Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("TargetDeclination Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("TargetDeclination Write", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//TargetRightAscension Read - Optional
		//First read should fail!
		try {
			m_TargetRightAscension = m_Telescope.TargetRightAscension;
			LogMsg("TargetRightAscension Read", MessageLevel.msgIssue, "Read before write should generate an error and didn't");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("TargetRightAscension Read", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.ValueNotSet:
				case g_ExNotSet1:
				case g_ExNotSet2:
					LogMsg("TargetRightAscension Read", MessageLevel.msgOK, "COM Not Set exception generated on read before write");
				default:
					ExTest("TargetRightAscension Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("TargetRightAscension Read", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (ASCOM.InvalidOperationException ex) {
			LogMsg("TargetDeclination Read", MessageLevel.msgOK, ".NET Not InvalidOperationException generated on read before write");
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.ValueNotSet:
				case g_ExNotSet1:
				case g_ExNotSet2:
					LogMsg("TargetRightAscension Read", MessageLevel.msgOK, ".NET Not Set exception generated on read before write");
				default:
					ExTest("TargetRightAscension Read", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (System.InvalidOperationException ex) {
			LogMsg("TargetRightAscension Read", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException");
		} catch (Exception ex) {
			ExTest("TargetRightAscension Read", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//TargetRightAscension Write - Optional
		try {
			m_Telescope.TargetRightAscension = -1.0;
			LogMsg("TargetRightAscension Write", MessageLevel.msgIssue, "No error generated on set TargetRightAscension < 0 hours");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("TargetRightAscension Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set TargetRightAscension < 0 hours");
				default:
					ExTest("TargetRightAscension Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("TargetRightAscension Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("TargetRightAscension Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set TargetRightAscension < 0 hours");
				default:
					ExTest("TargetRightAscension Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("TargetRightAscension Write", ex.Message, EX_NET + ex.Message);
		}

		try {
			m_Telescope.TargetRightAscension = 25.0;
			LogMsg("TargetRightAscension Write", MessageLevel.msgIssue, "No error generated on set TargetRightAscension > 24 hours");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("TargetRightAscension Write", MessageLevel.msgOK, NOT_IMP_COM);
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "COM Invalid Value exception generated as expected on set TargetRightAscension > 24 hours");
				default:
					ExTest("TargetRightAscension Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("TargetRightAscension Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (DriverException ex) {
			switch (ex.Number) {
				case ErrorCodes.InvalidValue:
				case g_ExInvalidValue1:
				case g_ExInvalidValue2:
				case g_ExInvalidValue3:
				case g_ExInvalidValue4:
				case g_ExInvalidValue5:
				case g_ExInvalidValue6:
					LogMsg("TargetRightAscension Write", MessageLevel.msgOK, ".NET Invalid Value exception generated as expected on set TargetRightAscension > 24 hours");
				default:
					ExTest("TargetRightAscension Write", ex.Message, EX_DRV_NET + ex.Message + " " + Hex(ex.Number));
			}
		} catch (Exception ex) {
			ExTest("TargetRightAscension Write", ex.Message, EX_NET + ex.Message);
		}

		try {
			m_TargetRightAscension = TelescopeRAFromSiderealTime(-4.0);
			m_Telescope.TargetRightAscension = m_TargetRightAscension;
			//Set a valid value
			try {
				switch (System.Math.Abs(m_Telescope.TargetRightAscension - m_TargetRightAscension)) {
					case 0.0:
						LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Legal value " + g_Util.HoursToHMS(m_Telescope.TargetRightAscension, ":", ":") + " HH:MM:SS written successfully");
					case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / 3600.0):
						//1 seconds
						LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 1 second of the value set: " + g_Util.HoursToHMS(m_Telescope.TargetRightAscension, ":", ":"));
						break; // TODO: might not be correct. Was : Exit Select

					case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / 3600.0):
						//2 seconds
						LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 2 seconds of the value set: " + g_Util.HoursToHMS(m_Telescope.TargetRightAscension, ":", ":"));
						break; // TODO: might not be correct. Was : Exit Select

					case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(5.0 / 3600.0):
						//5 seconds
						LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 5 seconds of the value set: " + g_Util.HoursToHMS(m_Telescope.TargetRightAscension, ":", ":"));
						break; // TODO: might not be correct. Was : Exit Select

					default:
						LogMsg("TargetRightAscension Write", MessageLevel.msgInfo, "Target RightAscension: " + g_Util.HoursToHMS(m_Telescope.TargetRightAscension, ":", ":"));
				}
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("TargetRightAscension Write", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("TargetRightAscension Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("TargetRightAscension Write", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("TargetRightAscension Write", ex.Message, EX_NET + ex.Message);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("TargetRightAscension Write", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("TargetRightAscension Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("TargetRightAscension Write", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("TargetRightAscension Write", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//Tracking Read - Required
		try {
			m_Tracking = m_Telescope.Tracking;
			//Read of tracking state is mandatory
			LogMsg("Tracking Read", MessageLevel.msgOK, m_Tracking.ToString);
		} catch (COMException ex) {
			LogMsg("Tracking Read", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("Tracking Read", MessageLevel.msgError, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//Tracking Write - Required
		l_OriginalTrackingState = m_Tracking;
		// Set should work OK
		if (m_CanSetTracking) {
			try {
				//OK try turning tracking off
				if (m_Tracking) {
					m_Telescope.Tracking = false;
				//OK try turning tracking on
				} else {
					m_Telescope.Tracking = true;
				}
				WaitFor(TRACKING_COMMAND_DELAY);
				//Wait for a short time to allow mounts to implement the tracking state change
				m_Tracking = m_Telescope.Tracking;
				if (m_Tracking != l_OriginalTrackingState) {
					LogMsg("Tracking Write", MessageLevel.msgOK, m_Tracking.ToString);
				} else {
					LogMsg("Tracking Write", MessageLevel.msgIssue, "Tracking didn't change state on write: " + m_Tracking.ToString);
				}
				m_Telescope.Tracking = l_OriginalTrackingState;
				//Restore original state
				WaitFor(TRACKING_COMMAND_DELAY);
				//Wait for a short time to allow mounts to implement the tracking state change
			} catch (COMException ex) {
				LogMsg("Tracking Write", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("Tracking Write", MessageLevel.msgError, EX_NET + ex.Message);
			}
		//Can read OK but Set tracking should fail
		} else {
			try {
				//OK try turning tracking off
				if (m_Tracking) {
					m_Telescope.Tracking = false;
				//OK try turning tracking on
				} else {
					m_Telescope.Tracking = true;
				}
				m_Tracking = m_Telescope.Tracking;
				LogMsg("Tracking Write", MessageLevel.msgIssue, "CanSetTracking is false but no error generated when value is set");
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("Tracking Write", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("Tracking Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("Tracking Write", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("Tracking Write", ex.Message, EX_NET + ex.Message);
			}
		}
		if (TestStop())
			return;

		//TrackingRates - Required
		if (g_InterfaceVersion > 1) {
			int l_Count = 0;

			try {
				m_TrackingRates = m_Telescope.TrackingRates;
				if (m_TrackingRates == null) {
					LogMsg("TrackingRates", MessageLevel.msgDebug, "ERROR: The driver did NOT return an TrackingRates object!");
				} else {
					LogMsg("TrackingRates", MessageLevel.msgDebug, "OK - the driver returned an TrackingRates object");
				}

				l_Count = m_TrackingRates.Count;
				//Save count for use later if no members are returned in the for each loop test
				LogMsg("TrackingRates Count", MessageLevel.msgDebug, l_Count);

				int i;
				for (i = 1; i <= m_TrackingRates.Count; i++) {
					LogMsg("TrackingRates Count", MessageLevel.msgDebug, "Found drive rate: " + Enum.GetName(typeof(DriveRates), m_TrackingRates.Item(i)));
				}
			} catch (COMException ex) {
				LogMsg("TrackingRates", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("TrackingRates", MessageLevel.msgError, EX_NET + ex.ToString);
			}

			try {
				System.Collections.IEnumerator l_Enum;
				object l_Obj;
				DriveRates l_Drv;
				l_Enum = m_TrackingRates.GetEnumerator();
				if (l_Enum == null) {
					LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "ERROR: The driver did NOT return an Enumerator object!");
				} else {
					LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "OK - the driver returned an Enumerator object");
				}

				l_Enum.Reset();
				LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Reset Enumerator");
				while (l_Enum.MoveNext()) {
					LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Reading Current");
					l_Obj = l_Enum.Current;
					LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Read Current OK, Type: " + l_Obj.GetType.Name);
					l_Drv = (DriveRates)l_Obj;
					LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Found drive rate: " + Enum.GetName(typeof(DriveRates), l_Drv));
				}

				l_Enum.Reset();
				l_Enum = null;
				m_TrackingRates = null;
			} catch (COMException ex) {
				LogMsg("TrackingRates", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("TrackingRates", MessageLevel.msgError, EX_NET + ex.ToString);
			}

			try {
				m_TrackingRates = m_Telescope.TrackingRates;
				LogMsg("TrackingRates", MessageLevel.msgDebug, "Read TrackingRates OK, Count: " + m_TrackingRates.Count);

				int l_RateCount = 0;
				foreach ( l_DriveRate in m_TrackingRates) {
					LogMsg("TrackingRates", MessageLevel.msgComment, "Found drive rate: " + l_DriveRate.ToString());
					l_RateCount += 1;
				}
				if (l_RateCount > 0) {
					LogMsg("TrackingRates", MessageLevel.msgOK, "Drive rates read OK");
				} else {
					//We did get some members on the first call, but now they have disappeared!
					if (l_Count > 0) {
						//This can be due to the drvier returning the same trackingrates object on every TrackingRates call but not resetting the iterator pointer
						LogMsg("TrackingRates", MessageLevel.msgError, "Multiple calls to TrackingRates returned different answers!");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "The first call to TrackingRates returned " + l_Count + " drive rates; the next call appeared to return no rates.");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "This can arise when the SAME TrackingRates object is returned on every TrackingRates call.");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "The root cause is usually that the enumeration pointer in the object is set to the end of the");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "collection through the application's use of the first object; subsequent uses see the pointer at the end");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "of the collection, which indicates no more members and is interpreted as meaning the collection is empty.");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "It is recommended to return a new TrackingRates object on each call. Alternatively, you could reset the");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "object's enumeration pointer every time the GetEnumerator method is called.");
						LogMsg("TrackingRates", MessageLevel.msgInfo, "");
					} else {
						LogMsg("TrackingRates", MessageLevel.msgIssue, "No drive rates returned");
					}
				}
			} catch (COMException ex) {
				LogMsg("TrackingRates", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("TrackingRates", MessageLevel.msgError, EX_NET + ex.ToString);
			}
			if (TestStop())
				return;

			//Trackingrate - Test after trackingrates so we kow what the valid values are - Optional
			try {
				m_TrackingRate = m_Telescope.TrackingRate;
				LogMsg("TrackingRate Read", MessageLevel.msgOK, m_TrackingRate.ToString);
				//Now test trying to set each tracking rate in turn
				foreach ( l_DriveRate in m_TrackingRates) {
					Application.DoEvents();
					if (TestStop())
						return;
					try {
						m_Telescope.TrackingRate = l_DriveRate;
						if (m_Telescope.TrackingRate == l_DriveRate) {
							LogMsg("TrackingRate Write", MessageLevel.msgOK, "Successfully set drive rate: " + l_DriveRate.ToString);
						} else {
							LogMsg("TrackingRate Write", MessageLevel.msgIssue, "Unable to set drive rate: " + l_DriveRate.ToString);
						}
					} catch (COMException ex) {
						if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
							LogMsg("TrackingRate Write", MessageLevel.msgOK, NOT_IMP_COM);
						} else {
							ExTest("TrackingRate Write", ex.Message, "Exception setting tracking rate: " + l_DriveRate + " " + ex.Message + " " + Hex(ex.ErrorCode));
						}
					} catch (PropertyNotImplementedException ex) {
						LogMsg("TrackingRate Write", MessageLevel.msgOK, NOT_IMP_NET);
					} catch (Exception ex) {
						ExTest("TrackingRate Write", ex.Message, "Exception setting tracking rate: " + l_DriveRate + " " + ex.Message);
					}
				}
				//Finally restore original trackingrate
				try {
					m_Telescope.TrackingRate = m_TrackingRate;
				} catch (Exception ex) {
					LogMsg("TrackingRate Write", MessageLevel.msgError, "Unable to restore original tracking rate. Exception: " + ex.Message);
				}
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("TrackingRate Read", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest("TrackingRate Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("TrackingRate Read", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("TrackingRate Read", ex.Message, EX_NET + ex.Message);
			}
		} else {
			LogMsg("TrackingRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		if (TestStop())
			return;

		//UTCDate Read - Required
		try {
			m_UTCDate = m_Telescope.UTCDate;
			//Save starting value
			LogMsg("UTCDate Read", MessageLevel.msgOK, m_UTCDate.ToString);
			// UTCDate Write is optional since if you are using the PC time as UTCTime then you should not write to the PC clock!
			try {
				m_Telescope.UTCDate = m_UTCDate.AddHours(1.0);
				//Try and write a new utcdate in the future
				LogMsg("UTCDate Write", MessageLevel.msgOK, "New UTCDate written successfully: " + m_Telescope.UTCDate.ToString);
				m_Telescope.UTCDate = m_UTCDate;
				//Restore original value
			//v1.0.12.0 Added code to allowably throw a not implemented exception
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg("UTCDate Write", MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					LogMsg("UTCDate Write", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("UTCDate Write", MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				ExTest("UTCDate Write", ex.Message, EX_NET + ex.Message);
			}
		} catch (COMException ex) {
			LogMsg("UTCDate Read", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("UTCDate Read", MessageLevel.msgError, EX_NET + ex.Message);
		}
		if (TestStop())
			return;
	}

	private override void CheckMethods()
	{
		//CanMoveAxis - Required - This must be first test as Parked tests use its results
		if (g_InterfaceVersion > 1) {
			if ((g_TelescopeTests.Item(TELTEST_CAN_MOVE_AXIS) == CheckState.Checked) | (g_TelescopeTests.Item(TELTEST_MOVE_AXIS) == CheckState.Checked) | (g_TelescopeTests.Item(TELTEST_PARK_UNPARK) == CheckState.Checked)) {
				TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisPrimary, "CanMoveAxis:Primary");
				if (TestStop())
					return;
				TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisSecondary, "CanMoveAxis:Secondary");
				if (TestStop())
					return;
				TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisTertiary, "CanMoveAxis:Tertiary");
				if (TestStop())
					return;
			} else {
				LogMsg(TELTEST_CAN_MOVE_AXIS, MessageLevel.msgInfo, "Tests skipped");
			}
		} else {
			LogMsg("CanMoveAxis", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		//Test Park, Unpark - Optional
		if (g_InterfaceVersion > 1) {
			if (g_TelescopeTests.Item(TELTEST_PARK_UNPARK) == CheckState.Checked) {
				//Can Park
				if (m_CanPark) {
					try {
						// OK We are unparked so check that no error is generated
						if (!m_Telescope.AtPark) {
							Status(StatusType.staTest, "Park");
							try {
								Status(StatusType.staAction, "Park scope");
								m_Telescope.Park();
								Status(StatusType.staStatus, "Waiting for scope to park");
								do {
									WaitFor(SLEEP_TIME);
									Application.DoEvents();
								} while (!(m_Telescope.AtPark | TestStop()));
								if (TestStop())
									return;
								Status(StatusType.staStatus, "Scope parked");
								LogMsg("Park", MessageLevel.msgOK, "Success");

								//Scope Parked OK
								//Confirm second park is harmless
								try {
									m_Telescope.Park();
									LogMsg("Park", MessageLevel.msgOK, "Success if already parked");
								} catch (COMException ex) {
									LogMsg("Park", MessageLevel.msgIssue, "Exception when calling Park two times in succession: " + ex.Message + " " + Hex(ex.ErrorCode));
								} catch (Exception ex) {
									LogMsg("Park", MessageLevel.msgIssue, "Exception when calling Park two times in succession: " + ex.Message);
								}

								//Confirm that methods do raise exceptions when scope is parked
								if (m_CanSlew | m_CanSlewAsync | m_CanSlewAltAz | m_CanSlewAltAzAsync){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepAbortSlew, "AbortSlew");if (TestStop())
									return;
}
								if (m_CanFindHome){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepFindHome, "FindHome");if (TestStop())
									return;
}
								if (m_CanMoveAxisPrimary){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisPrimary, "MoveAxis Primary");if (TestStop())
									return;
}
								if (m_CanMoveAxisSecondary){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisSecondary, "MoveAxis Secondary");if (TestStop())
									return;
}
								if (m_CanMoveAxisTertiary){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisTertiary, "MoveAxis Tertiary");if (TestStop())
									return;
}
								if (m_CanPulseGuide){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepPulseGuide, "PulseGuide");if (TestStop())
									return;
}
								if (m_CanSlew){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToCoordinates, "SlewToCoordinates");if (TestStop())
									return;
}
								if (m_CanSlewAsync){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToCoordinatesAsync, "SlewToCoordinatesAsync");if (TestStop())
									return;
}
								if (m_CanSlew){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToTarget, "SlewToTarget");if (TestStop())
									return;
}
								if (m_CanSlewAsync){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToTargetAsync, "SlewToTargetAsync");if (TestStop())
									return;
}
								if (m_CanSync){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSyncToCoordinates, "SyncToCoordinates");if (TestStop())
									return;
}
								if (m_CanSync){TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSyncToTarget, "SyncToTarget");if (TestStop())
									return;
}

								//Test unpark after park
								if (m_CanUnpark) {
									try {
										Status(StatusType.staAction, "UnPark scope after park");
										m_Telescope.Unpark();
										do {
											WaitFor(SLEEP_TIME);
											Application.DoEvents();
										} while (!(!m_Telescope.AtPark | TestStop()));
										if (TestStop())
											return;
										//Make sure tracking doesn't generate an error if it is not implemented
										try {
											m_Telescope.Tracking = true;
										} catch (Exception ex) {
										}
										Status(StatusType.staStatus, "Scope UnParked");
										LogMsg("UnPark", MessageLevel.msgOK, "Success");

										//Scope Unparked
										//Confirm UnPark is harmless if already unparked
										try {
											m_Telescope.Unpark();
											LogMsg("UnPark", MessageLevel.msgOK, "Success if already unparked");
										} catch (COMException ex) {
											LogMsg("UnPark", MessageLevel.msgIssue, "Exception when calling UnPark two times in succession: " + ex.Message + " " + Hex(ex.ErrorCode));
										} catch (Exception ex) {
											LogMsg("UnPark", MessageLevel.msgIssue, "Exception when calling UnPark two times in succession: " + ex.Message);
										}
									} catch (COMException ex) {
										LogMsg("UnPark", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
									} catch (Exception ex) {
										LogMsg("UnPark", MessageLevel.msgError, EX_NET + ex.Message);
									}
								// Can't unpark
								} else {
									//Confirm that Unpark generates an error
									//v1.0.12.0 added test for Unpark generating exception when CanUnPark is false
									try {
										m_Telescope.Unpark();
										LogMsg("UnPark", MessageLevel.msgIssue, "No exception thrown by UnPark when CanUnPark is false");
									} catch (COMException ex) {
										if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
											LogMsg("UnPark", MessageLevel.msgOK, NOT_IMP_COM);
										} else {
											ExTest("UnPark", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
										}
									} catch (MethodNotImplementedException ex) {
										LogMsg("UnPark", MessageLevel.msgOK, NOT_IMP_NET);
									} catch (Exception ex) {
										ExTest("UnPark", ex.Message, EX_NET + ex.Message);
									}
									//Create user interface message asking for manual scope unpark
									LogMsg("UnPark", MessageLevel.msgComment, "CanUnPark is false so you need to unpark manually");
									MsgBox("This scope cannot be unparked automatically, please unpark it now", MsgBoxStyle.Critical, "UnPark");
								}
							} catch (COMException ex) {
								LogMsg("Park", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
							} catch (Exception ex) {
								LogMsg("Park", MessageLevel.msgError, EX_NET + ex.Message);
							}
						// We are still in parked status despite a successul UnPark
						} else {
							LogMsg("Park", MessageLevel.msgError, "AtPark still true despite an earlier successful UnPark");
						}
					} catch (COMException ex) {
						LogMsg("Park", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
					} catch (Exception ex) {
						LogMsg("Park", MessageLevel.msgError, EX_NET + ex.Message);
					}
				//Can't park
				} else {
					try {
						m_Telescope.Park();
						LogMsg("Park", MessageLevel.msgError, "CanPark is false but no exception was generated on use");
					} catch (COMException ex) {
						if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
							LogMsg("Park", MessageLevel.msgOK, "CanPark is false and COM Not Implemented exception generated as expected");
						} else {
							ExTest("Park", ex.Message, "CanPark is false but inappropriate COM exception generated: " + ex.Message + " " + Hex(ex.ErrorCode));
						}
					} catch (MethodNotImplementedException ex) {
						LogMsg("Park", MessageLevel.msgOK, "CanPark is false and .NET MethodNotImplemented exception generated as expected");
					} catch (Exception ex) {
						ExTest("Park", ex.Message, "CanPark is false but inappropriate exception generated: " + ex.Message);
					}
					//v1.0.12.0 Added test for unpark if canpark is false
					//Now test unpark
					// We should already be unparked so confirm that unpark works fine
					if (m_CanUnpark) {
						try {
							m_Telescope.Unpark();
							LogMsg("UnPark", MessageLevel.msgOK, "CanPark is false and CanUnPark is true; no exception generated as expected");
						} catch (COMException ex) {
							LogMsg("UnPark", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
						} catch (Exception ex) {
							LogMsg("UnPark", MessageLevel.msgError, EX_NET + ex.Message);
						}
					//Can't unpark so confirm an exception is raised
					} else {
						try {
							m_Telescope.Unpark();
							LogMsg("UnPark", MessageLevel.msgError, "CanPark and CanUnPark are false but no exception was generated on use");
						} catch (COMException ex) {
							if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
								LogMsg("UnPark", MessageLevel.msgOK, "CanUnPark is false and COM Not Implemented exception generated as expected");
							} else {
								ExTest("UnPark", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
							}
						} catch (MethodNotImplementedException ex) {
							LogMsg("UnPark", MessageLevel.msgOK, "CanUnPark is false and MethodNotImplemented exception generated as expected");
						} catch (Exception ex) {
							ExTest("UnPark", ex.Message, EX_NET + ex.Message);
						}
					}
				}
				g_Status.Clear();
				//Clear status messages
				if (TestStop())
					return;
			} else {
				LogMsg(TELTEST_PARK_UNPARK, MessageLevel.msgInfo, "Tests skipped");
			}
		} else {
			LogMsg("Park", MessageLevel.msgInfo, "Skipping tests since behaviour of this method is not well defined in interface V" + g_InterfaceVersion);
		}

		//AbortSlew - Optional
		if (g_TelescopeTests.Item(TELTEST_ABORT_SLEW) == CheckState.Checked) {
			TelescopeOptionalMethodsTest(OptionalMethodType.tstAbortSlew, "AbortSlew", true);
			if (TestStop())
				return;
		} else {
			LogMsg(TELTEST_ABORT_SLEW, MessageLevel.msgInfo, "Tests skipped");
		}

		//AxisRates - Required
		if (g_InterfaceVersion > 1) {
			if ((g_TelescopeTests.Item(TELTEST_AXIS_RATE) == CheckState.Checked) | (g_TelescopeTests.Item(TELTEST_MOVE_AXIS) == CheckState.Checked)) {
				TelescopeAxisRateTest("AxisRate:Primary", TelescopeAxes.axisPrimary);
				TelescopeAxisRateTest("AxisRate:Secondary", TelescopeAxes.axisSecondary);
				TelescopeAxisRateTest("AxisRate:Tertiary", TelescopeAxes.axisTertiary);
			} else {
				LogMsg(TELTEST_AXIS_RATE, MessageLevel.msgInfo, "Tests skipped");
			}
		} else {
			LogMsg("AxisRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		//FindHome - Optional
		if (g_TelescopeTests.Item(TELTEST_FIND_HOME) == CheckState.Checked) {
			TelescopeOptionalMethodsTest(OptionalMethodType.tstFindHome, "FindHome", m_CanFindHome);
			if (TestStop())
				return;
		} else {
			LogMsg(TELTEST_FIND_HOME, MessageLevel.msgInfo, "Tests skipped");
		}

		//MoveAxis - Optional
		if (g_InterfaceVersion > 1) {
			if (g_TelescopeTests.Item(TELTEST_MOVE_AXIS) == CheckState.Checked) {
				TelescopeOptionalMethodsTest(OptionalMethodType.tstMoveAxisPrimary, "MoveAxis Primary", m_CanMoveAxisPrimary);
				if (TestStop())
					return;
				TelescopeOptionalMethodsTest(OptionalMethodType.tstMoveAxisSecondary, "MoveAxis Secondary", m_CanMoveAxisSecondary);
				if (TestStop())
					return;
				TelescopeOptionalMethodsTest(OptionalMethodType.tstMoveAxisTertiary, "MoveAxis Tertiary", m_CanMoveAxisTertiary);
				if (TestStop())
					return;
			} else {
				LogMsg(TELTEST_MOVE_AXIS, MessageLevel.msgInfo, "Tests skipped");
			}
		} else {
			LogMsg("MoveAxis", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		//PulseGuide - Optional
		if (g_TelescopeTests.Item(TELTEST_PULSE_GUIDE) == CheckState.Checked) {
			TelescopeOptionalMethodsTest(OptionalMethodType.tstPulseGuide, "PulseGuide", m_CanPulseGuide);
			if (TestStop())
				return;
		} else {
			LogMsg(TELTEST_PULSE_GUIDE, MessageLevel.msgInfo, "Tests skipped");
		}

		//Test Equatorial slewing to coordinates - Optional
		if (g_TelescopeTests.Item(TELTEST_SLEW_TO_COORDINATES) == CheckState.Checked) {
			TelescopeSlewTest(SlewType.tstSlewToCoordinates, "SlewToCoordinates", m_CanSlew, "CanSlew");
			if (TestStop())
				return;
		} else {
			LogMsg(TELTEST_SLEW_TO_COORDINATES, MessageLevel.msgInfo, "Tests skipped");
		}

		//Test Equatorial slewing to coordinates async - Optional
		if (g_TelescopeTests.Item(TELTEST_SLEW_TO_COORDINATES_ASYNC) == CheckState.Checked) {
			TelescopeSlewTest(SlewType.tstSlewToCoordinatesAsync, "SlewToCoordinatesAsync", m_CanSlewAsync, "CanSlewAsync");
			if (TestStop())
				return;
		} else {
			LogMsg(TELTEST_SLEW_TO_COORDINATES_ASYNC, MessageLevel.msgInfo, "Tests skipped");
		}

		//Test Equatorial target slewing - Optional
		if (g_TelescopeTests.Item(TELTEST_SLEW_TO_TARGET) == CheckState.Checked) {
			TelescopeSlewTest(SlewType.tstSlewToTarget, "SlewToTarget", m_CanSlew, "CanSlew");
			if (TestStop())
				return;
		} else {
			LogMsg(TELTEST_SLEW_TO_TARGET, MessageLevel.msgInfo, "Tests skipped");
		}

		//Test Equatorial target slewing async - Optional
		if (g_TelescopeTests.Item(TELTEST_SLEW_TO_TARGET_ASYNC) == CheckState.Checked) {
			TelescopeSlewTest(SlewType.tstSlewToTargetAsync, "SlewToTargetAsync", m_CanSlewAsync, "CanSlewAsync");
			if (TestStop())
				return;
		} else {
			LogMsg(TELTEST_SLEW_TO_TARGET_ASYNC, MessageLevel.msgInfo, "Tests skipped");
		}

		//DestinationSideOfPier - Optional
		if (g_InterfaceVersion > 1) {
			if (g_TelescopeTests.Item(TELTEST_DESTINATION_SIDE_OF_PIER) == CheckState.Checked) {
				if (m_AlignmentMode == ASCOM.DeviceInterface.AlignmentModes.algGermanPolar) {
					TelescopeOptionalMethodsTest(OptionalMethodType.tstDestinationSideOfPier, "DestinationSideOfPier", true);
					if (TestStop())
						return;
				} else {
					LogMsg("DestinationSideOfPier", MessageLevel.msgComment, "Test skipped as AligmentMode is not German Polar");
				}
			} else {
				LogMsg(TELTEST_DESTINATION_SIDE_OF_PIER, MessageLevel.msgInfo, "Tests skipped");
			}
		} else {
			LogMsg("DestinationSideOfPier", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		//Test AltAz Slewing - Optional
		if (g_InterfaceVersion > 1) {
			if (g_TelescopeTests.Item(TELTEST_SLEW_TO_ALTAZ) == CheckState.Checked) {
				TelescopeSlewTest(SlewType.tstSlewToAltAz, "SlewToAltAz", m_CanSlewAltAz, "CanSlewAltAz");
				if (TestStop())
					return;
			} else {
				LogMsg(TELTEST_SLEW_TO_ALTAZ, MessageLevel.msgInfo, "Tests skipped");
			}
		} else {
			LogMsg("SlewToAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		//Test AltAz Slewing async - Optional
		if (g_InterfaceVersion > 1) {
			if (g_TelescopeTests.Item(TELTEST_SLEW_TO_ALTAZ_ASYNC) == CheckState.Checked) {
				TelescopeSlewTest(SlewType.tstSlewToAltAzAsync, "SlewToAltAzAsync", m_CanSlewAltAzAsync, "CanSlewAltAzAsync");
				if (TestStop())
					return;
			} else {
				LogMsg(TELTEST_SLEW_TO_ALTAZ_ASYNC, MessageLevel.msgInfo, "Tests skipped");
			}
		} else {
			LogMsg("SlewToAltAzAsync", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		//Equatorial Sync - Optional
		if (g_TelescopeTests.Item(TELTEST_SYNC_TO_COORDINATES) == CheckState.Checked) {
			TelescopeSyncTest(SyncType.tstSyncToCoordinates, "SyncToCoordinates", m_CanSync, "CanSync");
			if (TestStop())
				return;
			TelescopeSyncTest(SyncType.tstSyncToTarget, "SyncToTarget", m_CanSync, "CanSync");
			if (TestStop())
				return;
		} else {
			LogMsg(TELTEST_SYNC_TO_COORDINATES, MessageLevel.msgInfo, "Tests skipped");
		}

		//AltAz Sync - Optional
		if (g_InterfaceVersion > 1) {
			if (g_TelescopeTests.Item(TELTEST_SYNC_TO_ALTAZ) == CheckState.Checked) {
				TelescopeSyncTest(SyncType.tstSyncToAltAz, "SyncToAltAz", m_CanSyncAltAz, "CanSyncAltAz");
				if (TestStop())
					return;
			} else {
				LogMsg(TELTEST_SYNC_TO_ALTAZ, MessageLevel.msgInfo, "Tests skipped");
			}
		} else {
			LogMsg("SyncToAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}

		if (g_Settings.TestSideOfPier) {
			LogMsg("", MessageLevel.msgAlways, "");
			LogMsg("SideOfPier Model Tests", MessageLevel.msgAlways, "");

			if (g_InterfaceVersion > 1) {
				//3.0.0.14 - Skip these tests if unable to read SideOfPier
				if (m_CanReadSideOfPier) {
					//SideOfPier write property test - Optional
					TelescopeOptionalMethodsTest(OptionalMethodType.tstSideOfPierWrite, "SideOfPier Write", m_CanSetPierside);
					if (TestStop())
						return;

					//Further side of pier tests
					if (m_Telescope.AlignmentMode == AlignmentModes.algGermanPolar) {
						SideOfPierTests();
						//Only run these for German mounts
					} else {
						LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Test skipped as this is not a German equatorial mount");
					}
				} else {
					LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Tests skipped as this driver does not support SideOfPier Read");
				}
			} else {
				LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
			}
		}
		g_Status.Clear();
		//Clear status messages

	}

	public override void CheckPerformance()
	{
		Status(StatusType.staTest, "Performance");
		//Clear status messages
		LogMsg("Transactions", MessageLevel.msgAlways, "per second: <1=Rather low, 1-2=OK, 2-10=Good, >10=Rather high");
		TelescopePerformanceTest(PerformanceType.tstPerfAltitude, "Altitude");
		if (TestStop())
			return;
		if (g_InterfaceVersion > 1) {
			TelescopePerformanceTest(PerformanceType.tstPerfAtHome, "AtHome");
			if (TestStop())
				return;
		} else {
			LogMsg("Performance: AtHome", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		if (g_InterfaceVersion > 1) {
			TelescopePerformanceTest(PerformanceType.tstPerfAtPark, "AtPark");
			if (TestStop())
				return;
		} else {
			LogMsg("Performance: AtPark", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		TelescopePerformanceTest(PerformanceType.tstPerfAzimuth, "Azimuth");
		if (TestStop())
			return;
		TelescopePerformanceTest(PerformanceType.tstPerfDeclination, "Declination");
		if (TestStop())
			return;
		if (g_InterfaceVersion > 1) {
			if (m_CanPulseGuide) {
				TelescopePerformanceTest(PerformanceType.tstPerfIsPulseGuiding, "IsPulseGuiding");
				if (TestStop())
					return;
			} else {
				LogMsg("Performance: IsPulseGuiding", MessageLevel.msgOK, "Test omitted since IsPulseGuiding is not implemented");
			}
		} else {
			LogMsg("Performance: IsPulseGuiding", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		TelescopePerformanceTest(PerformanceType.tstPerfRightAscension, "RightAscension");
		if (TestStop())
			return;
		if (g_InterfaceVersion > 1) {
			if (m_AlignmentMode == AlignmentModes.algGermanPolar) {
				if (m_CanReadSideOfPier) {
					TelescopePerformanceTest(PerformanceType.tstPerfSideOfPier, "SideOfPier");
					if (TestStop())
						return;
				} else {
					LogMsg("Performance: SideOfPier", MessageLevel.msgOK, "Test omitted since SideOfPier is not implemented");
				}
			} else {
				LogMsg("Performance: SideOfPier", MessageLevel.msgOK, "Test omitted since alignment mode is not German Polar");
			}
		} else {
			LogMsg("Performance: SideOfPier", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
		}
		TelescopePerformanceTest(PerformanceType.tstPerfSiderealTime, "SiderealTime");
		if (TestStop())
			return;
		TelescopePerformanceTest(PerformanceType.tstPerfSlewing, "Slewing");
		if (TestStop())
			return;
		TelescopePerformanceTest(PerformanceType.tstPerfUTCDate, "UTCDate");
		if (TestStop())
			return;
		g_Status.Clear();
	}
	public override void PostRunCheck()
	{
		//Make things safe
		// LogMsg("", MessageLevel.msgAlways, "") 'Blank line
		try {
			if (m_Telescope.CanSetTracking) {
				m_Telescope.Tracking = false;
				LogMsg("Mount Safety", MessageLevel.msgOK, "Tracking stopped to protect your mount.");
			} else {
				LogMsg("Mount Safety", MessageLevel.msgInfo, "Tracking can't be turned off for this mount, please switch off manually.");
			}
		} catch (Exception ex) {
			LogMsg("Mount Safety", MessageLevel.msgError, "Exception when disabling tracking to protect mount: " + ex.ToString);
		}
	}

	private void TelescopeSyncTest(SyncType p_Test, string p_Name, bool p_CanDoIt, string p_CanDoItName)
	{
		double l_Difference;
		try {
			LogMsg(p_Name, MessageLevel.msgDebug, "Reading RightAscension");
			m_CurrentRA = m_Telescope.RightAscension;
			LogMsg(p_Name, MessageLevel.msgDebug, "Read RightAscension, Reading Declination");
			m_CurrentDEC = m_Telescope.Declination;
			LogMsg(p_Name, MessageLevel.msgDebug, "Read Declination, RA/DEC: " + g_Util.HoursToHMS(m_CurrentRA) + " " + g_Util.DegreesToDMS(m_CurrentDEC, ":", ":", ""));
			m_SyncRA = m_CurrentRA - 0.02;
			if (m_SyncRA < 0.0)
				m_SyncRA = m_SyncRA + 24.0;
			//Ensure legal RA
			//Ensure legal DEC
			if (m_CurrentDEC < 89.0) {
				m_SyncDEC = m_CurrentDEC + 0.02;
			} else {
				m_SyncDEC = m_CurrentDEC - 0.02;
			}
			switch (p_Test) {
				case SyncType.tstSyncToCoordinates:
					//SyncToCoordinates
					if (m_CanSetTracking & (!m_Telescope.Tracking))
						m_Telescope.Tracking = true;
					LogMsg(p_Name, MessageLevel.msgDebug, "SyncToCoordinates: " + g_Util.HoursToHMS(m_SyncRA) + " " + g_Util.DegreesToDMS(m_SyncDEC, ":", ":", ""));
					m_Telescope.SyncToCoordinates(m_SyncRA, m_SyncDEC);
					//Sync to slightly different coordinates
					LogMsg(p_Name, MessageLevel.msgDebug, "Completed SyncToCoordinates");
				case SyncType.tstSyncToTarget:
					//SyncToTarget
					if (m_CanSetTracking & (!m_Telescope.Tracking))
						m_Telescope.Tracking = true;
					try {
						LogMsg(p_Name, MessageLevel.msgDebug, "Setting TargetRightAscension: " + g_Util.HoursToHMS(m_SyncRA));
						m_Telescope.TargetRightAscension = m_SyncRA;
						LogMsg(p_Name, MessageLevel.msgDebug, "Set TargetRightAscension, Setting TargetDeclination: " + g_Util.DegreesToDMS(m_SyncDEC, ":", ":", ""));
						m_Telescope.TargetDeclination = m_SyncDEC;
						LogMsg(p_Name, MessageLevel.msgDebug, "Set TargetDeclination, SyncToTarget");
					} catch (PropertyNotImplementedException ex) {
						LogMsg(p_Name, MessageLevel.msgInfo, "TargetRightAscension or TargetDeclination is throwing a PropertyNotImplementedException");
					}

					m_Telescope.SyncToTarget();
					//Sync to slightly different coordinates
					LogMsg(p_Name, MessageLevel.msgDebug, "Completed SyncToTarget");
				case SyncType.tstSyncToAltAz:
					if (m_CanReadAltitide)
						m_CurrentAlt = m_Telescope.Altitude;
					if (m_CanReadAzimuth)
						m_CurrentAz = m_Telescope.Azimuth;
					m_SyncAlt = m_CurrentAlt - 1.0;
					m_SyncAz = m_CurrentAz + 1.0;
					if (m_SyncAlt < 0.0)
						m_SyncAlt = 1.0;
					//Ensure legal Alt
					if (m_SyncAz > 359.0)
						m_SyncAz = 358.0;
					//Ensure legal Az
					if (m_CanSetTracking & m_Telescope.Tracking)
						m_Telescope.Tracking = false;
					//Sync to new Alt Az
					m_Telescope.SyncToAltAz(m_SyncAz, m_SyncAlt);
				default:
					LogMsg(p_Name, MessageLevel.msgError, "Conform:SyncTest: Unknown test type " + p_Test.ToString);
			}
			//This should have worked so report the outcome
			if (p_CanDoIt) {
				switch (p_Test) {
					case SyncType.tstSyncToCoordinates:
					case SyncType.tstSyncToTarget:
						//Only do this for euqatorial syncs
						m_NewRA = m_Telescope.RightAscension;
						m_NewDEC = m_Telescope.Declination;
						LogMsg(p_Name, MessageLevel.msgComment, "RA Original: " + g_Util.HoursToHMS(m_CurrentRA) + " Sync to: " + g_Util.HoursToHMS(m_SyncRA) + " New:" + g_Util.HoursToHMS(m_NewRA));
						LogMsg(p_Name, MessageLevel.msgComment, "DEC Original: " + g_Util.DegreesToDMS(m_CurrentDEC, ":", ":") + " Sync to: " + g_Util.DegreesToDMS(m_SyncDEC, ":", ":") + " New:" + g_Util.DegreesToDMS(m_NewDEC, ":", ":"));
						//Compare old and new values
						l_Difference = System.Math.Abs(m_SyncRA - m_NewRA);
						switch (l_Difference) {
							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / (60 * 60)):
								//Within 1 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Synced RA OK");
								break; // TODO: might not be correct. Was : Exit Select

							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / (60 * 60)):
								//Within 2 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Synced within 2 seconds of RA");
								break; // TODO: might not be correct. Was : Exit Select

							default:
								LogMsg(p_Name, MessageLevel.msgInfo, "Synced to within " + g_Util.HoursToHMS(l_Difference) + " HH:MM:SS of expected RA co-ordinate: " + g_Util.HoursToHMS(m_SyncRA));
						}
						l_Difference = System.Math.Abs(m_SyncDEC - m_NewDEC);
						switch (l_Difference) {
							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / (60 * 60)):
								//Within 1 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Synced DEC OK");
								break; // TODO: might not be correct. Was : Exit Select

							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / (60 * 60)):
								//Within 2 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Synced within 2 seconds of DEC");
								break; // TODO: might not be correct. Was : Exit Select

							default:
								LogMsg(p_Name, MessageLevel.msgInfo, "Synced to within " + g_Util.DegreesToDMS(l_Difference, ":", ":") + " DD:MM:SS of expected DEC co-ordinate: " + g_Util.DegreesToDMS(m_SyncDEC, ":", ":"));
						}
					case SyncType.tstSyncToAltAz:
						//Can check effects of a sync
						if (m_CanReadAltitide & m_CanReadAzimuth) {
							m_NewAlt = m_Telescope.Altitude;
							m_NewAz = m_Telescope.Azimuth;
							LogMsg(p_Name, MessageLevel.msgComment, "Altitude Original: " + g_Util.DegreesToDMS(m_CurrentAlt, ":", ":") + " Sync to: " + g_Util.DegreesToDMS(m_SyncAlt, ":", ":") + " New:" + g_Util.DegreesToDMS(m_NewAlt, ":", ":"));
							LogMsg(p_Name, MessageLevel.msgComment, "Azimuth Original: " + g_Util.DegreesToDMS(m_CurrentAz, ":", ":") + " Sync to: " + g_Util.DegreesToDMS(m_SyncAz, ":", ":") + " New:" + g_Util.DegreesToDMS(m_NewAz, ":", ":"));
							//Compare old and new values
							l_Difference = System.Math.Abs(m_SyncAlt - m_NewAlt);
							switch (l_Difference) {
								case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / (60 * 60)):
									//Within 1 seconds
									LogMsg(p_Name, MessageLevel.msgOK, "Synced Altitude OK");
									break; // TODO: might not be correct. Was : Exit Select

								case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / (60 * 60)):
									//Within 2 seconds
									LogMsg(p_Name, MessageLevel.msgOK, "Synced within 2 seconds of Altitude");
									break; // TODO: might not be correct. Was : Exit Select

								default:
									LogMsg(p_Name, MessageLevel.msgInfo, "Synced to within " + g_Util.DegreesToDMS(l_Difference, ":", ":") + " DD:MM:SS of expected Altitude: " + g_Util.DegreesToDMS(m_SyncAlt, ":", ":"));
							}
							l_Difference = System.Math.Abs(m_SyncAz - m_NewAz);
							switch (l_Difference) {
								case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / (60 * 60)):
									//Within 1 seconds
									LogMsg(p_Name, MessageLevel.msgOK, "Synced Azimuth OK");
									break; // TODO: might not be correct. Was : Exit Select

								case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / (60 * 60)):
									//Within 2 seconds
									LogMsg(p_Name, MessageLevel.msgOK, "Synced within 2 seconds of Azimuth");
									break; // TODO: might not be correct. Was : Exit Select

								default:
									LogMsg(p_Name, MessageLevel.msgInfo, "Synced to within " + g_Util.DegreesToDMS(l_Difference, ":", ":") + " DD:MM:SS of expected Azimuth: " + g_Util.DegreesToDMS(m_SyncAz, ":", ":"));
							}
						//Can't test effects of a sync
						} else {
							LogMsg(p_Name, MessageLevel.msgInfo, "Can't test SyncToAltAz because Altitude or Azimuth values are not implemented");
						}
						//Do nothing
					default:
				}
			//Shouldn't have been able to do this but no error was generated
			} else {
				LogMsg(p_Name, MessageLevel.msgIssue, p_CanDoItName + " is false but no exception was generated on use");
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				if (!p_CanDoIt) {
					LogMsg(p_Name, MessageLevel.msgOK, p_CanDoItName + " is False and COM Not Implemented exception generated as expected");
				} else {
					LogMsg(p_Name, MessageLevel.msgIssue, p_CanDoItName + " is True but COM Not Implemented exception was generated");
				}
			} else {
				ExTest(p_Name, ex.Message, p_CanDoItName + " is " + p_CanDoIt.ToString + " but inappropriate COM exception generated: " + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (MethodNotImplementedException ex) {
			if (!p_CanDoIt) {
				LogMsg(p_Name, MessageLevel.msgOK, p_CanDoItName + " is False and .NET Not Implemented exception generated as expected");
			} else {
				LogMsg(p_Name, MessageLevel.msgIssue, p_CanDoItName + " is True but .NET Not Implemented exception was generated");
			}
		} catch (Exception ex) {
			ExTest(p_Name, ex.Message, EX_NET + ex.ToString);
		}
	}
	private void TelescopeSlewTest(SlewType p_Test, string p_Name, bool p_CanDoIt, string p_CanDoItName)
	{
		double l_Difference;
		Status(StatusType.staTest, p_Name);
		//3.0.0.12 - Added test for ability to set tracking
		if (m_CanSetTracking)
			m_Telescope.Tracking = true;
		//Enable tracking for these tests

		try {
			switch (p_Test) {
				case SlewType.tstSlewToCoordinates:
					//Next line changed in v1.0.12 to ensure that no error is generated from trying to change the tracking state if the driver does not allow tracking to be changed
					//If Not m_Telescope.Tracking Then m_Telescope.Tracking = True
					if (m_CanSetTracking & (!m_Telescope.Tracking))
						m_Telescope.Tracking = true;
					m_TargetRightAscension = TelescopeRAFromSiderealTime(-1.0);
					m_TargetDeclination = 1.0;
					Status(StatusType.staAction, "Slewing");
					m_Telescope.SlewToCoordinates(m_TargetRightAscension, m_TargetDeclination);
				case SlewType.tstSlewToCoordinatesAsync:
					//Next line changed in v1.0.12 to ensure that no error is generated from trying to change the tracking state if the driver does not allow tracking to be changed
					//If Not m_Telescope.Tracking Then m_Telescope.Tracking = True
					if (m_CanSetTracking & (!m_Telescope.Tracking))
						m_Telescope.Tracking = true;
					m_TargetRightAscension = TelescopeRAFromSiderealTime(-2.0);
					m_TargetDeclination = 2.0;
					Status(StatusType.staAction, "Slewing");
					m_Telescope.SlewToCoordinatesAsync(m_TargetRightAscension, m_TargetDeclination);
					WaitForSlew();
				case SlewType.tstSlewToTarget:
					//Next line changed in v1.0.12 to ensure that no error is generated from trying to change the tracking state if the driver does not allow tracking to be changed
					//If Not m_Telescope.Tracking Then m_Telescope.Tracking = True
					if (m_CanSetTracking & (!m_Telescope.Tracking))
						m_Telescope.Tracking = true;
					m_TargetRightAscension = TelescopeRAFromSiderealTime(-3.0);
					m_TargetDeclination = 3.0;

					try {
						m_Telescope.TargetRightAscension = m_TargetRightAscension;
						m_Telescope.TargetDeclination = m_TargetDeclination;
					} catch (PropertyNotImplementedException ex) {
						LogMsg(p_Name, MessageLevel.msgInfo, "TargetRightAscension or TargetDeclination is throwing a PropertyNotImplementedException");
					}
					Status(StatusType.staAction, "Slewing");
					m_Telescope.SlewToTarget();
				case SlewType.tstSlewToTargetAsync:
					//SlewToTargetAsync
					//Next line changed in v1.0.12 to ensure that no error is generated from trying to change the tracking state if the driver does not allow tracking to be changed
					//If Not m_Telescope.Tracking Then m_Telescope.Tracking = True
					if (m_CanSetTracking & (!m_Telescope.Tracking))
						m_Telescope.Tracking = true;
					m_TargetRightAscension = TelescopeRAFromSiderealTime(-4.0);
					m_TargetDeclination = 4.0;
					try {
						m_Telescope.TargetRightAscension = m_TargetRightAscension;
						m_Telescope.TargetDeclination = m_TargetDeclination;
					} catch (PropertyNotImplementedException ex) {
						LogMsg(p_Name, MessageLevel.msgInfo, "TargetRightAscension or TargetDeclination is throwing a PropertyNotImplementedException");
					}
					Status(StatusType.staAction, "Slewing");
					m_Telescope.SlewToTargetAsync();
					WaitForSlew();
				case SlewType.tstSlewToAltAz:
					//Next line changed in v1.0.12 to ensure that no error is generated from trying to change the tracking state if the driver does not allow tracking to be changed
					//If m_Telescope.Tracking Then m_Telescope.Tracking = False
					LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 1: " + m_Telescope.Tracking);
					if (m_CanSetTracking & m_Telescope.Tracking) {
						m_Telescope.Tracking = false;
						LogMsg(p_Name, MessageLevel.msgDebug, "Tracking turned off");
					}
					LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 2: " + m_Telescope.Tracking);
					m_TargetAltitude = 50.0;
					m_TargetAzimuth = 150.0;
					Status(StatusType.staAction, "Slewing to Alt/Az: " + g_Util.DegreesToDMS(m_TargetAltitude, ":", ":", "") + " " + g_Util.DegreesToDMS(m_TargetAzimuth, ":", ":", ""));
					m_Telescope.SlewToAltAz(m_TargetAzimuth, m_TargetAltitude);
					LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 3: " + m_Telescope.Tracking);
				case SlewType.tstSlewToAltAzAsync:
					//Next line changed in v1.0.12 to ensure that no error is generated from trying to change the tracking state if the driver does not allow tracking to be changed
					//If m_Telescope.Tracking Then m_Telescope.Tracking = False
					LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 1: " + m_Telescope.Tracking);
					if (m_CanSetTracking & m_Telescope.Tracking) {
						m_Telescope.Tracking = false;
						LogMsg(p_Name, MessageLevel.msgDebug, "Tracking turned off");
					}
					LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 2: " + m_Telescope.Tracking);
					m_TargetAltitude = 55.0;
					m_TargetAzimuth = 155.0;
					Status(StatusType.staAction, "Slewing to Alt/Az: " + g_Util.DegreesToDMS(m_TargetAltitude, ":", ":", "") + " " + g_Util.DegreesToDMS(m_TargetAzimuth, ":", ":", ""));
					m_Telescope.SlewToAltAzAsync(m_TargetAzimuth, m_TargetAltitude);
					LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 3: " + m_Telescope.Tracking);
					WaitForSlew();
					LogMsg(p_Name, MessageLevel.msgDebug, "Tracking 4: " + m_Telescope.Tracking);
				default:
					LogMsg(p_Name, MessageLevel.msgError, "Conform:SlewTest: Unknown test type " + p_Test.ToString);
			}
			if (TestStop())
				return;
			//Should be able to do this so report what happened
			if (p_CanDoIt) {
				switch (p_Test) {
					case SlewType.tstSlewToCoordinates:
					case SlewType.tstSlewToCoordinatesAsync:
					case SlewType.tstSlewToTarget:
					case SlewType.tstSlewToTargetAsync:
						Status(StatusType.staAction, "Slew completed");
						//Test how close the slew was to the required coordinates
						l_Difference = System.Math.Abs(m_Telescope.RightAscension - m_TargetRightAscension);
						if (l_Difference > 23.0)
							l_Difference = 24.0 - l_Difference;
						// Deal with the case where the two elements are on different sides of 24hrs
						switch (l_Difference) {
							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / 3600.0):
								// seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Slewed to target RA OK: " + g_Util.HoursToHMS(m_TargetRightAscension));
								break; // TODO: might not be correct. Was : Exit Select

							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / 3600.0):
								//2 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Slewed to within 2 seconds of RA target: " + g_Util.HoursToHMS(m_TargetRightAscension) + " Actual RA " + g_Util.HoursToHMS(m_Telescope.RightAscension));
								break; // TODO: might not be correct. Was : Exit Select

							default:
								LogMsg(p_Name, MessageLevel.msgInfo, "Slewed to within " + g_Util.HoursToHMS(l_Difference) + " HH:MM:SS of expected RA co-ordinate: " + g_Util.HoursToHMS(m_TargetRightAscension));
						}
						l_Difference = System.Math.Abs(m_Telescope.Declination - m_TargetDeclination);
						switch (l_Difference) {
							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / 3600.0):
								//<1 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Slewed to target DEC OK: " + g_Util.DegreesToDMS(m_TargetDeclination, ":", ":"));
								break; // TODO: might not be correct. Was : Exit Select

							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / 3600.0):
								//2 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Slewed to within 2 seconds of DEC target: " + g_Util.DegreesToDMS(m_TargetDeclination, ":", ":") + " Actual DEC " + g_Util.DegreesToDMS(m_Telescope.Declination, ":", ":"));
								break; // TODO: might not be correct. Was : Exit Select

							default:
								LogMsg(p_Name, MessageLevel.msgInfo, "Slewed to within " + g_Util.DegreesToDMS(l_Difference, ":", ":") + " DD:MM:SS of expected DEC co-ordinate: " + g_Util.DegreesToDMS(m_TargetDeclination, ":", ":"));
						}
					case SlewType.tstSlewToAltAz:
					case SlewType.tstSlewToAltAzAsync:
						Status(StatusType.staAction, "Slew completed");
						//Test how close the slew was to the required coordinates
						l_Difference = System.Math.Abs(m_Telescope.Azimuth - m_TargetAzimuth);
						if (l_Difference > 350.0)
							l_Difference = 360.0 - l_Difference;
						// Deal with the case where the two elements are on different sides of 360 degrees
						switch (l_Difference) {
							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / 3600.0):
								// seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Slewed to target Azimuth OK: " + g_Util.DegreesToDMS(m_TargetAzimuth, ":", ":"));
								break; // TODO: might not be correct. Was : Exit Select

							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / 3600.0):
								//2 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Slewed to within 2 seconds of Azimuth target: " + g_Util.DegreesToDMS(m_TargetAzimuth, ":", ":") + " Actual Azimuth " + g_Util.DegreesToDMS(m_Telescope.Azimuth, ":", ":"));
								break; // TODO: might not be correct. Was : Exit Select

							default:
								LogMsg(p_Name, MessageLevel.msgInfo, "Slewed to within " + g_Util.DegreesToDMS(l_Difference, ":", ":") + " DD:MM:SS of expected Azimuth: " + g_Util.DegreesToDMS(m_TargetAzimuth, ":", ":"));
						}
						l_Difference = System.Math.Abs(m_Telescope.Altitude - m_TargetAltitude);
						switch (l_Difference) {
							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(1.0 / 3600.0):
								//<1 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Slewed to target Altitude OK: " + g_Util.DegreesToDMS(m_TargetAltitude, ":", ":"));
								break; // TODO: might not be correct. Was : Exit Select

							case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
(2.0 / 3600.0):
								//2 seconds
								LogMsg(p_Name, MessageLevel.msgOK, "Slewed to within 2 seconds of Altitude target: " + g_Util.DegreesToDMS(m_TargetAltitude, ":", ":") + " Actual Altitude " + g_Util.DegreesToDMS(m_Telescope.Altitude, ":", ":"));
								break; // TODO: might not be correct. Was : Exit Select

							default:
								LogMsg(p_Name, MessageLevel.msgInfo, "Slewed to within " + g_Util.DegreesToDMS(l_Difference, ":", ":") + " DD:MM:SS of expected Altitude: " + g_Util.DegreesToDMS(m_TargetAltitude, ":", ":"));
						}
						//Do nothing
					default:
				}
			//Not supposed to be able to do this but no error generated so report an error
			} else {
				LogMsg(p_Name, MessageLevel.msgIssue, p_CanDoItName + " is false but no exception was generated on use");
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				if (!p_CanDoIt) {
					LogMsg(p_Name, MessageLevel.msgOK, p_CanDoItName + " is False and COM Not Implemented exception generated as expected");
				} else {
					LogMsg(p_Name, MessageLevel.msgIssue, p_CanDoItName + " is True but COM Not Implemented exception was generated");
				}
			} else {
				LogMsg(p_Name, MessageLevel.msgDebug, ex.ToString);
				ExTest(p_Name, ex.Message, p_CanDoItName + " is " + p_CanDoIt.ToString + " but inappropriate COM exception generated: " + ex.ToString + " " + Hex(ex.ErrorCode));
			}
		} catch (MethodNotImplementedException ex) {
			if (!p_CanDoIt) {
				LogMsg(p_Name, MessageLevel.msgOK, p_CanDoItName + " is False and .NET Not Implemented exception generated as expected");
			} else {
				LogMsg(p_Name, MessageLevel.msgIssue, p_CanDoItName + " is True but .NET Not Implemented exception was generated");
			}
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgDebug, ex.ToString);
			ExTest(p_Name, ex.Message, EX_NET + ex.Message);
		}
		g_Status.Clear();
		//Clear status messages
	}

	private void TelescopePerformanceTest(PerformanceType p_Type, string p_Name)
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
					case PerformanceType.tstPerfAltitude:
						m_Altitude = m_Telescope.Altitude;
					case PerformanceType.tstPerfAtHome:
						m_AtHome = m_Telescope.AtHome;
					case PerformanceType.tstPerfAtPark:
						m_AtPark = m_Telescope.AtPark;
					case PerformanceType.tstPerfAzimuth:
						m_Azimuth = m_Telescope.Azimuth;
					case PerformanceType.tstPerfDeclination:
						m_Declination = m_Telescope.Declination;
					case PerformanceType.tstPerfIsPulseGuiding:
						m_IsPulseGuiding = m_Telescope.IsPulseGuiding;
					case PerformanceType.tstPerfRightAscension:
						m_RightAscension = m_Telescope.RightAscension;
					case PerformanceType.tstPerfSideOfPier:
						m_SideOfPier = m_Telescope.SideOfPier;
					case PerformanceType.tstPerfSiderealTime:
						m_SiderealTimeScope = m_Telescope.SiderealTime;
					case PerformanceType.tstPerfSlewing:
						m_Slewing = m_Telescope.Slewing;
					case PerformanceType.tstPerfUTCDate:
						m_UTCDate = m_Telescope.UTCDate;
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
					LogMsg("Performance: " + p_Name, MessageLevel.msgInfo, "Transaction rate: " + Format(l_Rate, "0.0") + " per second");
					break; // TODO: might not be correct. Was : Exit Select

				case 2.0: // TODO: to 10.0
					LogMsg("Performance: " + p_Name, MessageLevel.msgOK, "Transaction rate: " + Format(l_Rate, "0.0") + " per second");
					break; // TODO: might not be correct. Was : Exit Select

				case 1.0: // TODO: to 2.0
					LogMsg("Performance: " + p_Name, MessageLevel.msgInfo, "Transaction rate: " + Format(l_Rate, "0.0") + " per second");
					break; // TODO: might not be correct. Was : Exit Select

				default:
					LogMsg("Performance: " + p_Name, MessageLevel.msgInfo, "Transaction rate: " + Format(l_Rate, "0.0") + " per second");
			}
		} catch (Exception ex) {
			LogMsg("Performance: " + p_Name, MessageLevel.msgError, EX_NET + ex.Message);
		}
	}
	private void TelescopeParkedExceptionTest(ParkedExceptionType p_Type, string p_Name)
	{
		g_Status.Action = p_Name;
		g_Status.Test = p_Type.ToString;
		// We are still parked so test abortslew
		if (m_Telescope.AtPark) {
			try {
				switch (p_Type) {
					case ParkedExceptionType.tstPExcepAbortSlew:
						m_Telescope.AbortSlew();
					case ParkedExceptionType.tstPExcepFindHome:
						m_Telescope.FindHome();
					case ParkedExceptionType.tstPExcepMoveAxisPrimary:
						m_Telescope.MoveAxis(TelescopeAxes.axisPrimary, 0);
					case ParkedExceptionType.tstPExcepMoveAxisSecondary:
						m_Telescope.MoveAxis(TelescopeAxes.axisSecondary, 0);
					case ParkedExceptionType.tstPExcepMoveAxisTertiary:
						m_Telescope.MoveAxis(TelescopeAxes.axisTertiary, 0);
					case ParkedExceptionType.tstPExcepPulseGuide:
						m_Telescope.PulseGuide(ASCOM.DeviceInterface.GuideDirections.guideEast, 0);
					case ParkedExceptionType.tstPExcepSlewToCoordinates:
						m_Telescope.SlewToCoordinates(TelescopeRAFromSiderealTime(1.0), 0.0);
					case ParkedExceptionType.tstPExcepSlewToCoordinatesAsync:
						m_Telescope.SlewToCoordinatesAsync(TelescopeRAFromSiderealTime(1.0), 0.0);
						WaitForSlew();
					case ParkedExceptionType.tstPExcepSlewToTarget:
						m_Telescope.TargetRightAscension = TelescopeRAFromSiderealTime(1.0);
						m_Telescope.TargetDeclination = 0.0;
						m_Telescope.SlewToTarget();
					case ParkedExceptionType.tstPExcepSlewToTargetAsync:
						m_Telescope.TargetRightAscension = TelescopeRAFromSiderealTime(1.0);
						m_Telescope.TargetDeclination = 0.0;
						m_Telescope.SlewToTargetAsync();
						WaitForSlew();
					case ParkedExceptionType.tstPExcepSyncToCoordinates:
						m_Telescope.SyncToCoordinates(TelescopeRAFromSiderealTime(1.0), 0.0);
					case ParkedExceptionType.tstPExcepSyncToTarget:
						m_Telescope.TargetRightAscension = TelescopeRAFromSiderealTime(1.0);
						m_Telescope.TargetDeclination = 0.0;
						m_Telescope.SyncToTarget();
					default:
						LogMsg("Park:" + p_Name, MessageLevel.msgError, "Conform:ParkedExceptionTest: Unknown test type " + p_Type.ToString);
				}
				LogMsg("Park:" + p_Name, MessageLevel.msgIssue, p_Name + " didn't raise an error when Parked as required");
			} catch (Exception ex) {
				LogMsg("Park:" + p_Name, MessageLevel.msgOK, p_Name + " did raise an exception when Parked as required");
			}
			//Check that Telescope is still parked after issuing the command!
			if (!m_Telescope.AtPark)
				LogMsg("Park:" + p_Name, MessageLevel.msgIssue, "Telescope was unparked by the " + p_Name + " command. This should not happen!");

		} else {
			LogMsg("Park:" + p_Name, MessageLevel.msgIssue, "Not parked after Telescope.Park command, " + p_Name + " when parked test skipped");
		}
		g_Status.Clear();
	}

	private void TelescopeAxisRateTest(string p_Name, ASCOM.DeviceInterface.TelescopeAxes p_Axis)
	{
		int l_NAxisRates;
		int l_i;
		int l_j;
		bool l_AxisRateOverlap;
		bool l_AxisRateDuplicate;
		int l_Count = 0;

		#If DEBUG Then
		object l_AxisRates;
		#Else
		ASCOM.DeviceInterface.IAxisRates l_AxisRates;
		#End If

		try {
			l_NAxisRates = 0;
			l_AxisRates = null;
			switch (p_Axis) {
				case TelescopeAxes.axisPrimary:
					m_AxisRatesPrimary = m_Telescope.AxisRates(ASCOM.DeviceInterface.TelescopeAxes.axisPrimary);
					//Get primary axis rates
					l_AxisRates = m_AxisRatesPrimary;
					m_AxisRatePrimaryOverlap = false;
				case TelescopeAxes.axisSecondary:
					m_AxisRatesSecondary = m_Telescope.AxisRates(ASCOM.DeviceInterface.TelescopeAxes.axisSecondary);
					//Get secondary axis rates
					l_AxisRates = m_AxisRatesSecondary;
					m_AxisRateSecondaryOverlap = false;
				case TelescopeAxes.axisTertiary:
					m_AxisRatesTertiary = m_Telescope.AxisRates(ASCOM.DeviceInterface.TelescopeAxes.axisTertiary);
					//Get tertiary axis rates
					l_AxisRates = m_AxisRatesTertiary;
					m_AxisRateTertiaryOverlap = false;
				default:
					LogMsg("TelescopeAxisRateTest", MessageLevel.msgError, "Unknown telescope axis: " + p_Axis.ToString);
			}


			try {
				if (l_AxisRates == null) {
					LogMsg(p_Name, MessageLevel.msgDebug, "ERROR: The driver did NOT return an AxisRates object!");
				} else {
					LogMsg(p_Name, MessageLevel.msgDebug, "OK - the driver returned an AxisRates object");
				}

				l_Count = l_AxisRates.Count;
				//Save count for use later if no members are returned in the for each loop test
				LogMsg(p_Name + " Count", MessageLevel.msgDebug, "The driver returned " + l_Count + " rates");

				int i;
				for (i = 1; i <= l_Count; i++) {
					IRate l_Rate;
					l_Rate = l_AxisRates.Item(i);
					LogMsg(p_Name + " Count", MessageLevel.msgDebug, "Rate " + i + " - Minimum: " + l_Rate.Minimum.ToString + ", Maximum: " + l_Rate.Maximum.ToString);
				}
			} catch (COMException ex) {
				LogMsg(p_Name + " Count", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg(p_Name + " Count", MessageLevel.msgError, EX_NET + ex.ToString);
			}

			try {
				System.Collections.IEnumerator l_Enum;
				object l_Obj;
				IRate l_Rate;

				l_Enum = l_AxisRates.GetEnumerator();
				if (l_Enum == null) {
					LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "ERROR: The driver did NOT return an Enumerator object!");
				} else {
					LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "OK - the driver returned an Enumerator object");
				}

				l_Enum.Reset();
				LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "Reset Enumerator");

				while (l_Enum.MoveNext()) {
					LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "Reading Current");
					l_Obj = l_Enum.Current;
					LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "Read Current OK, Type: " + l_Obj.GetType.Name);
					l_Rate = (IRate)l_Obj;
					LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "Found axis rate - Minimum: " + l_Rate.Minimum.ToString + ", Maximum: " + l_Rate.Maximum.ToString);
				}

				l_Enum.Reset();
				l_Enum = null;
				l_Rate = null;
			} catch (COMException ex) {
				LogMsg(p_Name + " Enum", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg(p_Name + " Enum", MessageLevel.msgError, EX_NET + ex.ToString);
			}

			if (l_AxisRates.Count > 0) {
				try {
					foreach ( Each in this.m_Rate) {
						// Error because negative values are not allowed
						if ((m_Rate.Minimum < 0) | (m_Rate.Maximum < 0)) {
							LogMsg(p_Name, MessageLevel.msgError, "Minimum or maximum rate is negative: " + m_Rate.Minimum.ToString + ", " + m_Rate.Maximum.ToString);
						// All positive values so continue tests
						} else {
							//Minimum <= Maximum so OK
							if (m_Rate.Minimum <= m_Rate.Maximum) {
								LogMsg(p_Name, MessageLevel.msgOK, "Axis rate minimum: " + m_Rate.Minimum.ToString + " Axis rate maximum: " + m_Rate.Maximum.ToString);
							//Minimum > Maximum so error!
							} else {
								LogMsg(p_Name, MessageLevel.msgError, "Maximum rate is less than minimum rate - minimum: " + m_Rate.Minimum.ToString + " maximum: " + m_Rate.Maximum.ToString);
							}
						}

						//Save rates for overlap testing
						l_NAxisRates += 1;
						m_AxisRatesArray(l_NAxisRates, AXIS_RATE_MINIMUM) = m_Rate.Minimum;
						m_AxisRatesArray(l_NAxisRates, AXIS_RATE_MAXIMUM) = m_Rate.Maximum;
					}
				} catch (COMException ex) {
					LogMsg(p_Name, MessageLevel.msgError, "COM Unable to read AxisRates object - Exeption: " + ex.Message + " " + Hex(ex.ErrorCode));
				} catch (DriverException ex) {
					LogMsg(p_Name, MessageLevel.msgError, ".NET Unable to read AxisRates object - Exeption: " + ex.Message + " " + Hex(ex.Number));
				} catch (Exception ex) {
					LogMsg(p_Name, MessageLevel.msgError, "Unable to read AxisRates object - Exeption: " + ex.Message);
				}

				//Overlap testing
				//Confirm whether there are overlaps if number of axisrate pairs exceeds 1
				if (l_NAxisRates > 1) {
					for (l_i = 1; l_i <= l_NAxisRates; l_i++) {
						for (l_j = 1; l_j <= l_NAxisRates; l_j++) {
							//Only test different lines, shouldn't compare same lines!
							if (l_i != l_j) {
								if ((m_AxisRatesArray(l_i, AXIS_RATE_MINIMUM) >= m_AxisRatesArray(l_j, AXIS_RATE_MINIMUM)) & (m_AxisRatesArray(l_i, AXIS_RATE_MINIMUM) <= m_AxisRatesArray(l_j, AXIS_RATE_MAXIMUM)))
									l_AxisRateOverlap = true;
							}
						}
					}
				}
				if (l_AxisRateOverlap) {
					LogMsg(p_Name, MessageLevel.msgIssue, "Overlapping axis rates found, suggest these be rationalised to remove overlaps");
				} else {
					LogMsg(p_Name, MessageLevel.msgOK, "No overlapping axis rates found");
				}

				//Duplicate testing
				l_AxisRateDuplicate = false;
				//Confirm whether there are overlaps if number of axisrate pairs exceeds 1
				if (l_NAxisRates > 1) {
					for (l_i = 1; l_i <= l_NAxisRates; l_i++) {
						for (l_j = 1; l_j <= l_NAxisRates; l_j++) {
							//Only test different lines, shouldn't compare same lines!
							if (l_i != l_j) {
								if ((m_AxisRatesArray(l_i, AXIS_RATE_MINIMUM) == m_AxisRatesArray(l_j, AXIS_RATE_MINIMUM)) & (m_AxisRatesArray(l_i, AXIS_RATE_MAXIMUM) == m_AxisRatesArray(l_j, AXIS_RATE_MAXIMUM)))
									l_AxisRateDuplicate = true;
							}
						}
					}
				}
				if (l_AxisRateDuplicate) {
					LogMsg(p_Name, MessageLevel.msgIssue, "Duplicate axis rates found, suggest these be removed");
				} else {
					LogMsg(p_Name, MessageLevel.msgOK, "No duplcate axis rates found");
				}

			} else {
				LogMsg(p_Name, MessageLevel.msgOK, "Empty axis rate returned");
			}
		} catch (COMException ex) {
			LogMsg(p_Name, MessageLevel.msgError, "COM Unable to get an AxisRates object - Exeption: " + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (DriverException ex) {
			LogMsg(p_Name, MessageLevel.msgError, ".NET Unable to get an AxisRates object - Exeption: " + ex.Message + " " + Hex(ex.Number));
		//Report null objects returned by the driver that are caughtby DriverAccess.
		} catch (NullReferenceException ex) {
			LogMsg(p_Name, MessageLevel.msgError, ex.Message);
			LogMsg(p_Name, MessageLevel.msgDebug, ex.ToString());
			//If debug then give full information
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgError, "Unable to get an AxisRates object - Exeption: " + ex.ToString);
		}

		//Clean up and release each object after use
		if (!(m_Rate == null))
			try {
				Marshal.ReleaseComObject(m_Rate);
			} catch {
			}
		m_Rate = null;
		g_Status.Clear();
		//Clear status messages

	}

	private void TelescopeAxisRateCheckup()
	{
		string p_Name = "TelescopeAxisRateCheckup";
		//########### DUPLICATE REMOVE AFTER TEST
		try {
			LogMsg(p_Name, MessageLevel.msgDebug, "m_AxisRatePrimary.Count = " + m_AxisRatesPrimary.Count);
			foreach ( Each in this.m_Rate) {
				// Error because negative valuies are not allowed
				if ((m_Rate.Minimum < 0) | (m_Rate.Maximum < 0)) {
					LogMsg(p_Name, MessageLevel.msgError, "Minimum or maximum rate is negative: " + m_Rate.Minimum.ToString + ", " + m_Rate.Maximum.ToString);
				// All positive values so continue tests
				} else {
					//Minimum <= Maximum so OK
					if (m_Rate.Minimum <= m_Rate.Maximum) {
						LogMsg(p_Name, MessageLevel.msgOK, "Axis rate minimum: " + m_Rate.Minimum.ToString + " Axis rate maximum: " + m_Rate.Maximum.ToString);
					//Minimum > Maximum so error!
					} else {
						LogMsg(p_Name, MessageLevel.msgError, "Maximum rate is less than minimum rate - minimum: " + m_Rate.Minimum.ToString + " maximum: " + m_Rate.Maximum.ToString);
					}
				}
			}
			LogMsg(p_Name, MessageLevel.msgDebug, "End of For each m_Rate In m_AxisRatesPrimary loop");
		} catch (COMException ex) {
			LogMsg(p_Name, MessageLevel.msgError, "COM Unable to read AxisRates object - Exeption: " + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (DriverException ex) {
			LogMsg(p_Name, MessageLevel.msgError, ".NET Unable to read AxisRates object - Exeption: " + ex.Message + " " + Hex(ex.Number));
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgError, "Unable to read AxisRates object - Exeption: " + ex.Message);
		}
		//################# DUPLICATE

	}

	private void TelescopeRequiredMethodsTest(RequiredMethodType p_Type, string p_Name)
	{
		try {
			g_Status.Test = p_Name;
			switch (p_Type) {
				case RequiredMethodType.tstAxisrates:
				//This is now done by TelescopeAxisRateTest subroutine 
				case RequiredMethodType.tstCanMoveAxisPrimary:
					m_CanMoveAxisPrimary = m_Telescope.CanMoveAxis(ASCOM.DeviceInterface.TelescopeAxes.axisPrimary);
					LogMsg(p_Name, MessageLevel.msgOK, p_Name + " " + m_CanMoveAxisPrimary.ToString);
				case RequiredMethodType.tstCanMoveAxisSecondary:
					m_CanMoveAxisSecondary = m_Telescope.CanMoveAxis(ASCOM.DeviceInterface.TelescopeAxes.axisSecondary);
					LogMsg(p_Name, MessageLevel.msgOK, p_Name + " " + m_CanMoveAxisSecondary.ToString);
				case RequiredMethodType.tstCanMoveAxisTertiary:
					m_CanMoveAxisTertiary = m_Telescope.CanMoveAxis(ASCOM.DeviceInterface.TelescopeAxes.axisTertiary);
					LogMsg(p_Name, MessageLevel.msgOK, p_Name + " " + m_CanMoveAxisTertiary.ToString);
				default:
					LogMsg(p_Name, MessageLevel.msgError, "Conform:RequiredMethodsTest: Unknown test type " + p_Type.ToString);
			}
		} catch (COMException ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.ToString);
		}

		//Clean up and release each object after use
		//If Not (m_Rate Is Nothing) Then Try : Marshal.ReleaseComObject(m_Rate) : Catch : End Try
		//m_Rate = Nothing
		g_Status.Clear();
		//Clear status messages
	}
	private void TelescopeOptionalMethodsTest(OptionalMethodType p_Type, string p_Name, bool p_CanTest)
	{
		int l_ct;
		double l_TestDec;
		double l_TestRA;

		//v1.0.12.0 removed redundant try catch block
		Status(StatusType.staTest, p_Name);
		// Confirm that an error is raised if the optional command is not implemented
		if (p_CanTest) {
			try {
				l_TestDec = 45.0;
				l_TestRA = 3.0;
				switch (p_Type) {
					case OptionalMethodType.tstAbortSlew:
						m_Telescope.AbortSlew();

						LogMsg("AbortSlew", MessageLevel.msgOK, "AbortSlew Ok when not slewing");
					case OptionalMethodType.tstDestinationSideOfPier:
						m_TargetRightAscension = TelescopeRAFromSiderealTime(-l_TestRA);
						m_DestinationSideOfPierEast = m_Telescope.DestinationSideOfPier(m_TargetRightAscension, l_TestDec);
						LogMsg("DestinationSideOfPier", MessageLevel.msgDebug, "East: " + g_Util.HoursToHMS(m_TargetRightAscension) + " " + g_Util.DegreesToDMS(l_TestDec, ":", ":", "") + " " + m_DestinationSideOfPier.ToString);
						m_TargetRightAscension = TelescopeRAFromSiderealTime(l_TestRA);
						m_DestinationSideOfPierWest = m_Telescope.DestinationSideOfPier(m_TargetRightAscension, l_TestDec);
						LogMsg("DestinationSideOfPier", MessageLevel.msgDebug, "West: " + g_Util.HoursToHMS(m_TargetRightAscension) + " " + g_Util.DegreesToDMS(l_TestDec, ":", ":", "") + " " + m_DestinationSideOfPier.ToString);
						if ((m_DestinationSideOfPierEast == PierSide.pierUnknown) | (m_DestinationSideOfPierWest == PierSide.pierUnknown)) {
							LogMsg("DestinationSideOfPier", MessageLevel.msgError, "Invalid SideOfPier value received, East: " + m_DestinationSideOfPierEast.ToString + ", West: " + m_DestinationSideOfPierWest.ToString);
						} else {
							if (m_DestinationSideOfPierEast == m_DestinationSideOfPierWest) {
								LogMsg("DestinationSideOfPier", MessageLevel.msgIssue, "Same value for DestinationSideOfPier received on both sides of the meridian: " + m_DestinationSideOfPierEast);
							} else {
								LogMsg("DestinationSideOfPier", MessageLevel.msgOK, "DestinationSideOfPier is different on either side of the meridian");
							}

						}
					case OptionalMethodType.tstFindHome:
						if (g_InterfaceVersion > 1) {
							m_Telescope.FindHome();
							m_StartTime = Now;
							Status(StatusType.staAction, "Waiting for mount to home");
							l_ct = 0;
							do {
								WaitFor(SLEEP_TIME);
								l_ct += 1;
								g_Status.Status = l_ct.ToString;
							} while (!(m_Telescope.AtHome | TestStop() | (Now.Subtract(m_StartTime).TotalMilliseconds > 60000)));
							// Wait up to a minute to find home
							if (m_Telescope.AtHome) {
								LogMsg(p_Name, MessageLevel.msgOK, "Found home OK.");
							} else {
								LogMsg(p_Name, MessageLevel.msgInfo, "Failed to Find home within 1 minute");
							}
							if (m_Telescope.AtPark) {
								LogMsg(p_Name, MessageLevel.msgIssue, "FindHome has parked the scope as well as finding home");
								m_Telescope.Unpark();
								//Unpark it ready for further tests
							}
						} else {
							Status(StatusType.staAction, "Waiting for mount to home");
							m_Telescope.FindHome();
							g_Status.Clear();
							LogMsg(p_Name, MessageLevel.msgOK, "Found home OK.");
							m_Telescope.Unpark();
							// Make sure we are still unparked!

						}
					case OptionalMethodType.tstMoveAxisPrimary:
						LogMsg(p_Name, MessageLevel.msgDebug, "Number of primary rates: " + m_AxisRatesPrimary.Count);
						m_AxisRatesPrimary.Dispose();
						m_AxisRatesPrimary = m_Telescope.AxisRates(TelescopeAxes.axisPrimary);

						TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisPrimary, m_AxisRatesPrimary);
					case OptionalMethodType.tstMoveAxisSecondary:
						m_AxisRatesSecondary.Dispose();
						m_AxisRatesSecondary = m_Telescope.AxisRates(TelescopeAxes.axisSecondary);
						LogMsg(p_Name, MessageLevel.msgDebug, "Number of secondary rates: " + m_AxisRatesSecondary.Count);

						TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisSecondary, m_AxisRatesSecondary);
					case OptionalMethodType.tstMoveAxisTertiary:
						LogMsg(p_Name, MessageLevel.msgDebug, "Number of tertiary rates: " + m_AxisRatesTertiary.Count);
						m_AxisRatesTertiary.Dispose();
						m_AxisRatesTertiary = m_Telescope.AxisRates(TelescopeAxes.axisTertiary);

						TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisTertiary, m_AxisRatesTertiary);
					case OptionalMethodType.tstPulseGuide:
						// IsPulseGuiding is true before we've started so this is an error and voids a real test
						if (m_Telescope.IsPulseGuiding) {
							LogMsg("PulseGuide", MessageLevel.msgError, "IsPulseGuiding is True when not pulse guiding - PulseGuide test omitted");
						//Ok to test pulse guiding
						} else {
							m_StartTime = Now;
							m_Telescope.PulseGuide(ASCOM.DeviceInterface.GuideDirections.guideEast, 2000);
							//Start a 2 second pulse
							m_EndTime = Now;
							Status(StatusType.staAction, "Start PulseGuide");
							//LogMsg("Pulseguide", MessageLevel.msgDebug, m_EndTime.Subtract(m_StartTime).TotalMilliseconds.ToString)
							//If less than 1.5 seconds then assume we have returned early
							if (m_EndTime.Subtract(m_StartTime).TotalMilliseconds < 1500) {
								//LogMsg("Pulseguide", MessageLevel.msgDebug, "before wait loop" & m_Telescope.IsPulseGuiding.ToString)
								//Thread.Sleep(SLEEP_TIME)
								l_ct = 0;
								if (m_Telescope.IsPulseGuiding) {
									do {
										WaitFor(SLEEP_TIME);
										l_ct += 1;
										g_Status.Status = l_ct.ToString;
										if (TestStop())
											return;
									} while (!((!m_Telescope.IsPulseGuiding) | (Now.Subtract(m_StartTime).TotalMilliseconds > 6000)));
									//Wait for up to 6 seconds
									//LogMsg("Pulseguide", MessageLevel.msgDebug, m_Telescope.IsPulseGuiding.ToString & " " & Now.Subtract(m_StartTime).TotalMilliseconds.ToString)

									if (!m_Telescope.IsPulseGuiding) {
										LogMsg("PulseGuide", MessageLevel.msgOK, "Asynchronous pulseguide found OK");
									} else {
										LogMsg("PulseGuide", MessageLevel.msgIssue, "Asynchronous pulseguide expected but IsPulseGuiding is still TRUE 4 seconds beyond expected time of 2 seconds");
									}
								} else {
									LogMsg("PulseGuide", MessageLevel.msgIssue, "Asynchronous pulseguide expected but IsPulseGuiding has returned FALSE");
								}
							//Assume synchronous pulseguide and that ispulseguideing is false
							} else {
								if (!m_Telescope.IsPulseGuiding) {
									LogMsg("PulseGuide", MessageLevel.msgOK, "Synchronous pulseguide found OK");
								} else {
									LogMsg("PulseGuide", MessageLevel.msgIssue, "Synchronous pulseguide expected but IsPulseGuiding has returned TRUE");
								}
							}

						}
					case OptionalMethodType.tstSideOfPierWrite:
						//SideOfPier Write
						//Can set pierside so test if we can
						if (m_CanSetPierside) {
							SlewScope(TelescopeRAFromHourAngle(-3.0), 0.0, "Slewing to far start point");
							if (TestStop())
								return;

							SlewScope(TelescopeRAFromHourAngle(-0.03), 0.0, "Slewing to near start point");
							//2 minutes from zenith
							if (TestStop())
								return;

							// We are now 2 minutes from the meridian looking east so allow the mount to track for 7 minutes 
							// so it passes through the meridian and ends up 5 minutes past the meridian
							LogMsg("SideOfPier Write", MessageLevel.msgInfo, "This test will now wait for 7 minutes while the mount tracks through the Meridian");

							//Wait for mount to move
							m_StartTime = Now;
							do {
								System.Threading.Thread.Sleep(SLEEP_TIME);
								Application.DoEvents();
								SetStatus("SideOfPier Write", "Waiting for transit through Meridian", Convert.ToInt32(Now.Subtract(m_StartTime).TotalSeconds) + "/" + SIDEOFPIER_MERIDIAN_TRACKING_PERIOD / 1000 + " seconds");
							} while (!((Now.Subtract(m_StartTime).TotalMilliseconds > SIDEOFPIER_MERIDIAN_TRACKING_PERIOD) | TestStop()));

							//SlewScope(TelescopeRAFromHourAngle(+0.0833333), 0.0, "Slewing to flip point") '5 minutes past zenith
							if (TestStop())
								return;

							switch (m_Telescope.SideOfPier) {
								case ASCOM.DeviceInterface.PierSide.pierEast:
									//We are on pierEast so try pierWest
									try {
										LogMsg("SideOfPier Write", MessageLevel.msgDebug, "Scope is pierEast so flipping West");
										m_Telescope.SideOfPier = ASCOM.DeviceInterface.PierSide.pierWest;
										WaitForSlew();
										if (TestStop())
											return;
										m_SideOfPier = m_Telescope.SideOfPier;
										if (m_SideOfPier == ASCOM.DeviceInterface.PierSide.pierWest) {
											LogMsg("SideOfPier Write", MessageLevel.msgOK, "Successfully flipped pierEast to pierWest");
										} else {
											LogMsg("SideOfPier Write", MessageLevel.msgIssue, "Failed to set SideOfPier to pierWest, got: " + m_SideOfPier.ToString);
										}
									//CanSetSideOfPier is True so any exception is an error
									} catch (COMException ex) {
										LogMsg("SideOfPier Write pierWest", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
									} catch (Exception ex) {
										LogMsg("SideOfPier Write pierWest", MessageLevel.msgError, EX_NET + ex.Message);
									}
								case ASCOM.DeviceInterface.PierSide.pierWest:
									//We are on pierWest so try pierEast
									try {
										LogMsg("SideOfPier Write", MessageLevel.msgDebug, "Scope is pierWest so flipping East");
										m_Telescope.SideOfPier = ASCOM.DeviceInterface.PierSide.pierEast;
										WaitForSlew();
										if (TestStop())
											return;
										m_SideOfPier = m_Telescope.SideOfPier;
										if (m_SideOfPier == ASCOM.DeviceInterface.PierSide.pierEast) {
											LogMsg("SideOfPier Write", MessageLevel.msgOK, "Successfully flipped pierWest to pierEast");
										} else {
											LogMsg("SideOfPier Write", MessageLevel.msgIssue, "Failed to set SideOfPier to pierEast, got: " + m_SideOfPier.ToString);
										}
									} catch (COMException ex) {
										LogMsg("SideOfPier Write pierEast", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
									} catch (Exception ex) {
										LogMsg("SideOfPier Write pierEast", MessageLevel.msgError, EX_NET + ex.Message);
									}
								default:
									//Unknown pier side
									LogMsg("SideOfPier Write pierUnknown", MessageLevel.msgError, "Unknown PierSide: " + m_SideOfPier.ToString);
							}
						//Can't set pierside so it should generate an error
						} else {
							try {
								LogMsg("SideofPier Write", MessageLevel.msgDebug, "Attempting to set SideOfPier");
								m_Telescope.SideOfPier = ASCOM.DeviceInterface.PierSide.pierEast;
								LogMsg("SideofPier Write", MessageLevel.msgDebug, "SideOfPier set OK to pierEast but should have thrown an error");
								WaitForSlew();
								LogMsg("SideOfPier Write", MessageLevel.msgIssue, "CanSetPierSide is false but no exception was generated when set was attempted");
							} catch (COMException ex) {
								if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
									LogMsg("SideOfPier Write", MessageLevel.msgOK, "COM CanSetPierSide is false and exception was properly generated");
								} else {
									ExTest("SideOfPier Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
								}
							} catch (PropertyNotImplementedException ex) {
								LogMsg("SideOfPier Write", MessageLevel.msgOK, ".NET CanSetPierSide is false and exception was properly generated");
							} catch (Exception ex) {
								ExTest("SideOfPier Write", ex.Message, EX_NET + ex.Message);
							//v1.0.12.0 Added finally statement
							} finally {
								WaitForSlew();
								//Make sure slewing is stopped if an exception was thrown
							}
						}

						m_Telescope.Tracking = false;
						if (TestStop())

							return;

					default:
						LogMsg(p_Name, MessageLevel.msgError, "Conform:OptionalMethodsTest: Unknown test type " + p_Type.ToString);
				}
				if (TestStop())
					return;
			} catch (COMException ex) {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (MethodNotImplementedException ex) {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
			} catch (Exception ex) {
				LogMsg(p_Name, MessageLevel.msgDebug, ex.ToString);
				ExTest(p_Name, ex.Message, EX_NET + ex.Message);
			}
		//Can property is false so confirm that an error is generated
		} else {
			try {
				switch (p_Type) {
					case OptionalMethodType.tstAbortSlew:
						m_Telescope.AbortSlew();
					case OptionalMethodType.tstDestinationSideOfPier:
						m_TargetRightAscension = TelescopeRAFromSiderealTime(-1.0);
						m_DestinationSideOfPier = m_Telescope.DestinationSideOfPier(m_TargetRightAscension, 0.0);
					case OptionalMethodType.tstFindHome:
						m_Telescope.FindHome();
					case OptionalMethodType.tstMoveAxisPrimary:
						m_Telescope.MoveAxis(TelescopeAxes.axisPrimary, 0.0);
					case OptionalMethodType.tstMoveAxisSecondary:
						m_Telescope.MoveAxis(TelescopeAxes.axisSecondary, 0.0);
					case OptionalMethodType.tstMoveAxisTertiary:
						m_Telescope.MoveAxis(TelescopeAxes.axisTertiary, 0.0);
					case OptionalMethodType.tstPulseGuide:
						m_Telescope.PulseGuide(GuideDirections.guideEast, 0);
					case OptionalMethodType.tstSideOfPierWrite:
						m_Telescope.SideOfPier = PierSide.pierEast;
					default:
						LogMsg(p_Name, MessageLevel.msgError, "Conform:OptionalMethodsTest: Unknown test type " + p_Type.ToString);
				}
				LogMsg(p_Name, MessageLevel.msgIssue, "Can" + p_Name + " is false but no exception was generated on use");
			} catch (COMException ex) {
				switch (ex.ErrorCode) {
					case g_ExNotImplemented:
					case ErrorCodes.NotImplemented:
						LogMsg(p_Name, MessageLevel.msgOK, "Can" + p_Name + " is false and COM Not Implemented exception generated as expected");
					case ErrorCodes.InvalidValue:
					case g_ExInvalidValue1:
					case g_ExInvalidValue2:
					case g_ExInvalidValue3:
					case g_ExInvalidValue4:
					case g_ExInvalidValue5:
					case g_ExInvalidValue6:
						LogMsg(p_Name, MessageLevel.msgOK, "Can" + p_Name + " is false and COM Invalid Value exception generated");
					default:
						ExTest(p_Name, ex.Message, "Can" + p_Name + " is false but inappropriate COM exception generated: " + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (MethodNotImplementedException ex) {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
			} catch (DriverException ex) {
				switch (ex.Number) {
					case g_ExNotImplemented:
					case ErrorCodes.NotImplemented:
						LogMsg(p_Name, MessageLevel.msgOK, "Can" + p_Name + " is false and .NET Not Implemented exception generated as expected");
					case ErrorCodes.InvalidValue:
					case g_ExInvalidValue1:
					case g_ExInvalidValue2:
					case g_ExInvalidValue3:
					case g_ExInvalidValue4:
					case g_ExInvalidValue5:
					case g_ExInvalidValue6:
						LogMsg(p_Name, MessageLevel.msgOK, "Can" + p_Name + " is false and .NET Invalid Value exception generated");
					default:
						ExTest(p_Name, ex.Message, "Can" + p_Name + " is false but inappropriate .NET exception generated: " + ex.Message + " " + Hex(ex.Number));
				}
			} catch (Exception ex) {
				ExTest(p_Name, ex.Message, "Can" + p_Name + " is false but inappropriate exception generated: " + ex.Message);
			}
		}
		g_Status.Clear();
		//Clear status messages
	}

	private void TelescopeCanTest(CanType p_Type, string p_Name)
	{
		try {
			switch (p_Type) {
				case CanType.tstCanFindHome:
					m_CanFindHome = m_Telescope.CanFindHome;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanFindHome.ToString);
				case CanType.tstCanPark:
					m_CanPark = m_Telescope.CanPark;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanPark.ToString);
				case CanType.tstCanPulseGuide:
					m_CanPulseGuide = m_Telescope.CanPulseGuide;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanPulseGuide.ToString);
				case CanType.tstCanSetDeclinationRate:
					if (g_InterfaceVersion > 1) {
						m_CanSetDeclinationRate = m_Telescope.CanSetDeclinationRate;
						LogMsg(p_Name, MessageLevel.msgOK, m_CanSetDeclinationRate.ToString);
					} else {
						LogMsg("CanSetDeclinationRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
					}
				case CanType.tstCanSetGuideRates:
					if (g_InterfaceVersion > 1) {
						m_CanSetGuideRates = m_Telescope.CanSetGuideRates;
						LogMsg(p_Name, MessageLevel.msgOK, m_CanSetGuideRates.ToString);
					} else {
						LogMsg("CanSetGuideRates", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
					}
				case CanType.tstCanSetPark:
					m_CanSetPark = m_Telescope.CanSetPark;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSetPark.ToString);
				case CanType.tstCanSetPierSide:
					if (g_InterfaceVersion > 1) {
						m_CanSetPierside = m_Telescope.CanSetPierSide;
						LogMsg(p_Name, MessageLevel.msgOK, m_CanSetPierside.ToString);
					} else {
						LogMsg("CanSetPierSide", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
					}
				case CanType.tstCanSetRightAscensionRate:
					if (g_InterfaceVersion > 1) {
						m_CanSetRightAscensionRate = m_Telescope.CanSetRightAscensionRate;
						LogMsg(p_Name, MessageLevel.msgOK, m_CanSetRightAscensionRate.ToString);
					} else {
						LogMsg("CanSetRightAscensionRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
					}
				case CanType.tstCanSetTracking:
					m_CanSetTracking = m_Telescope.CanSetTracking;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSetTracking.ToString);
				case CanType.tstCanSlew:
					m_CanSlew = m_Telescope.CanSlew;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSlew.ToString);
				case CanType.tstCanSlewAltAz:
					if (g_InterfaceVersion > 1) {
						m_CanSlewAltAz = m_Telescope.CanSlewAltAz;
						LogMsg(p_Name, MessageLevel.msgOK, m_CanSlewAltAz.ToString);
					} else {
						LogMsg("CanSlewAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
					}
				case CanType.tstCanSlewAltAzAsync:
					if (g_InterfaceVersion > 1) {
						m_CanSlewAltAzAsync = m_Telescope.CanSlewAltAzAsync;
						LogMsg(p_Name, MessageLevel.msgOK, m_CanSlewAltAzAsync.ToString);
					} else {
						LogMsg("CanSlewAltAzAsync", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
					}
				case CanType.tstCanSlewAsync:
					m_CanSlewAsync = m_Telescope.CanSlewAsync;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSlewAsync.ToString);
				case CanType.tstCanSync:
					m_CanSync = m_Telescope.CanSync;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanSync.ToString);
				case CanType.tstCanSyncAltAz:
					if (g_InterfaceVersion > 1) {
						m_CanSyncAltAz = m_Telescope.CanSyncAltAz;
						LogMsg(p_Name, MessageLevel.msgOK, m_CanSyncAltAz.ToString);
					} else {
						LogMsg("CanSyncAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
					}
				case CanType.tstCanUnPark:
					m_CanUnpark = m_Telescope.CanUnpark;
					LogMsg(p_Name, MessageLevel.msgOK, m_CanUnpark.ToString);
				default:
					LogMsg(p_Name, MessageLevel.msgError, "Conform:CanTest: Unknown test type " + p_Type.ToString);
			}
		} catch (COMException ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg(p_Name, MessageLevel.msgError, EX_NET + ex.Message);
		}
	}
	#If DEBUG Then
	#Else
	private void TelescopeMoveAxisTest(string p_Name, TelescopeAxes p_Axis, object p_AxisRates)
	{
	}
	#End If
	double l_MoveRate;
	double l_RateMinimum;
	double l_RateMaximum;
	bool l_TrackingStart;
	bool l_TrackingEnd;
	bool l_CanSetZero;

	int l_RateCount;
	//Determine lowest and highest tracking rates
		//Set to invalid values
	 l_RateMinimum = double.PositiveInfinity;
	 l_RateMaximum = double.NegativeInfinity;

	[,,] LogMsg = new[];
}
