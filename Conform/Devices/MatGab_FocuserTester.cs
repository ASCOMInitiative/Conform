internal class FocuserTester : DeviceTesterBaseClass
{

	#Region "Variables and Constants"
	//Focuser variables
	private bool m_Absolute;
	private bool m_IsMoving;
	private bool m_Link;
	private bool m_TempComp;
	private bool m_TempCompAvailable;
	private int m_MaxIncrement;
	private int m_MaxStep;
	private int m_Position;
	private int m_PositionOrg;
	private double m_StepSize;
	private double m_Temperature;
	private bool m_TempCompTrueOK;
		// Variable to confirm that TempComp can be successfully set to True
	private bool m_TempCompFalseOK;
		// Variable to confoirm that absolute position can be read OK
	private bool m_AbsolutePositionOK = false;

	#If DEBUG Then
		#Else
	private object m_Focuser;
		#End If
	private ASCOM.DeviceInterface.IFocuserV2 m_Focuser;
	#End Region

	#Region "Enums"
	private enum FocuserPropertyMethod
	{
		IsMoving,
		Position,
		Temperature
	}
	#End Region

	#Region "New and Dispose"
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
					m_Focuser.Link = false;
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_Focuser);
				} catch {
				}
				m_Focuser = null;
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

		switch (g_FocuserProgID) {
			default:
				//I'm using the simulator values as the defaults since it is the reference platform
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040404;
				g_ExInvalidValue2 = 0x80040404;
				g_ExNotSet1 = 0x80040403;
		}
		base.CheckInitialise(g_FocuserProgID);

	}
	private override void CheckAccessibility()
	{
		object l_IFocuser;
		DriverAccess.Focuser l_DriverAccessFocuser;
		object l_DeviceObject = null;
		base.CheckAccessibility(g_FocuserProgID, DeviceType.Focuser);
		//Try early binding
		l_IFocuser = null;
		try {
			l_DeviceObject = CreateObject(g_FocuserProgID);
			l_IFocuser = (ASCOM.Interface.IFocuser)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with IFocuser interface");
			try {
				l_IFocuser.Link = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes IFocuser interface");
				l_IFocuser.Link = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuser interface");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuser interface");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_IFocuser);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_IFocuser = null;
			GC.Collect();
		}
		WaitFor(DEVICE_DESTROY_WAIT);
		//Wait to allow device to complete destruction

		//Try early binding
		l_IFocuser = null;
		try {
			l_DeviceObject = CreateObject(g_FocuserProgID);
			l_IFocuser = (ASCOM.DeviceInterface.IFocuserV2)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with IFocuserV2 interface");
			try {
				l_IFocuser.Link = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes IFocuserV2 interface");
				l_IFocuser.Link = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuserV2 interface");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose IFocuserV2 interface");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_IFocuser);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_IFocuser = null;
			GC.Collect();
		}
		WaitFor(DEVICE_DESTROY_WAIT);
		//Wait to allow device to complete destruction

		//Try client access toolkit
		l_DriverAccessFocuser = null;
		try {
			l_DriverAccessFocuser = new ASCOM.DriverAccess.Focuser(g_FocuserProgID);
			LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
			try {
				l_DriverAccessFocuser.Link = true;
				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
				l_DriverAccessFocuser.Link = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgError, "Error conecting to driver using driver access toolkit: " + ex.Message);
				LogMsg("", MessageLevel.msgAlways, "");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " + ex.ToString);
			LogMsg("", MessageLevel.msgAlways, "");
		} finally {
			//Clean up
			try {
				l_DriverAccessFocuser.Dispose();
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DriverAccessFocuser);
			} catch {
			}
			l_DriverAccessFocuser = null;
			GC.Collect();
		}
		WaitFor(DEVICE_DESTROY_WAIT);
		//Wait to allow device to complete destruction
	}
	#If DEBUG Then
	private override void CreateDevice()
	{
		LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Focuser object");
		m_Focuser = CreateObject(g_FocuserProgID);
		#Else
		LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Focuser to get a Focuser object");
		m_Focuser = new ASCOM.DriverAccess.Focuser(g_FocuserProgID);
		#End If
		g_Stop = false;
		//connected OK so clear stop flag to allow other tests to run
	}
	private override bool Connected {
		get { Connected = base.Connected; }
		set {
			m_Focuser.Link = value;
			base.Connected = value;
		}
	}

	private override void CheckCommonMethods()
	{
		base.CheckCommonMethods(m_Focuser, DeviceType.Focuser);
	}

	private override void CheckProperties()
	{
		//Absolute - Required
		try {
			m_Absolute = m_Focuser.Absolute;
			LogMsg("Absolute", MessageLevel.msgOK, m_Absolute.ToString);
		} catch (COMException ex) {
			LogMsg("Absolute", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("Absolute", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//IsMoving - Required
		try {
			m_IsMoving = m_Focuser.IsMoving;
			if (!m_IsMoving) {
				LogMsg("IsMoving", MessageLevel.msgOK, m_IsMoving.ToString);
			} else {
				LogMsg("IsMoving", MessageLevel.msgError, "IsMoving is True at start of tests and it should be false");
			}
		} catch (COMException ex) {
			LogMsg("IsMoving", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("IsMoving", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//MaxStep - Required
		try {
			m_MaxStep = m_Focuser.MaxStep;
			LogMsg("MaxStep", MessageLevel.msgOK, m_MaxStep.ToString);
		} catch (COMException ex) {
			LogMsg("MaxStep", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("MaxStep", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//MaxIncrement - Required
		try {
			m_MaxIncrement = m_Focuser.MaxIncrement;
			//Minimum value is 1, 0 or negative must be a bad value, >maxstep is a bad value
			switch (m_MaxIncrement) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
1:
					LogMsg("MaxIncrement", MessageLevel.msgIssue, "MaxIncrement must be at least 1, actual value: " + m_MaxIncrement);
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
m_MaxStep:
					LogMsg("MaxIncrement", MessageLevel.msgIssue, "MaxIncrement is greater than MaxStep and shouldn't be: " + m_MaxIncrement);
				default:
					LogMsg("MaxIncrement", MessageLevel.msgOK, m_MaxIncrement.ToString);
			}
		} catch (COMException ex) {
			LogMsg("MaxIncrement", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("MaxIncrement", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//Position - Optional
		//Absolute focuser so this property shold be supported
		if (m_Absolute) {
			try {
				m_AbsolutePositionOK = false;
				m_Position = m_Focuser.Position;
				switch (m_Position) {
					//Check that position is a valid value
					case  // ERROR: Case labels with binary operators are unsupported : LessThan
0:
						//Lower than lowest position
						LogMsg("", MessageLevel.msgWarning, "Position is < 0, actual value: " + m_Position.ToString);
					case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
m_MaxStep:
						//> higest position
						LogMsg("", MessageLevel.msgWarning, "Position is > MaxStep, actual value: " + m_Position.ToString);
					default:
						//Valid value
						LogMsg("Position", MessageLevel.msgOK, m_Position.ToString);
						m_AbsolutePositionOK = true;
				}
			} catch (COMException ex) {
				LogMsg("Position", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("Position", MessageLevel.msgError, EX_NET + ex.Message);
			}
		//Relative focuser so this should raise a not supported error
		} else {
			try {
				m_Position = m_Focuser.Position;
				LogMsg("Position", MessageLevel.msgIssue, "This is a relative focuser but it didn't raise an exception for Focuser.Position");
			} catch (COMException ex) {
				switch (ex.ErrorCode) {
					case g_ExNotImplemented:
					case ErrorCodes.NotImplemented:
					case ErrorCodes.NotImplemented:
						LogMsg("Position", MessageLevel.msgOK, "COM Not implemented exception correctly raised as expected");
					default:
						LogMsg("Position", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("Position", MessageLevel.msgOK, ".NET Not implemented exception correctly raised as expected");
			} catch (Exception ex) {
				LogMsg("Position", MessageLevel.msgError, EX_NET + ex.Message);
			}
		}

		//StepSize - Optional
		try {
			m_StepSize = m_Focuser.StepSize;
			switch (m_StepSize) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
0.0:
					//Must be >0
					LogMsg("StepSize", MessageLevel.msgIssue, "StepSize must be > 0.0, actual value: " + m_StepSize);
				default:
					LogMsg("StepSize", MessageLevel.msgOK, m_StepSize.ToString);
			}
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("StepSize", MessageLevel.msgOK, "COM StepSize is not supported by this focuser");
				default:
					LogMsg("StepSize", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("StepSize", MessageLevel.msgOK, ".NET StepSize is not supported by this focuser");
		} catch (Exception ex) {
			LogMsg("StepSize", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//TempCompAvailable - Required
		try {
			m_TempCompAvailable = m_Focuser.TempCompAvailable;
			LogMsg("TempCompAvailable", MessageLevel.msgOK, m_TempCompAvailable.ToString);
		} catch (COMException ex) {
			LogMsg("TempCompAvailable", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("TempCompAvailable", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//TempComp Read - Required
		try {
			m_TempComp = m_Focuser.TempComp;
			if (m_TempComp & !m_TempCompAvailable) {
				LogMsg("TempComp Read", MessageLevel.msgIssue, "TempComp is True when TempCompAvailable is False - this should not be");
			} else {
				LogMsg("TempComp Read", MessageLevel.msgOK, m_TempComp.ToString);
			}
		} catch (COMException ex) {
			LogMsg("TempComp Read", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("TempComp Read", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//TempComp Write - Optional
		if (m_TempCompAvailable) {
			try {
				m_TempCompTrueOK = false;
				// Initialise to false
				m_TempCompFalseOK = false;
				//Turn compensation on 
				m_Focuser.TempComp = true;
				LogMsg("TempComp Write", MessageLevel.msgOK, "Successfully turned temperature compensation on");
				m_TempCompTrueOK = true;
				// Set to true to indicate TempComp acn be successfully set to True
				//Turn compensation off
				m_Focuser.TempComp = false;
				LogMsg("TempComp Write", MessageLevel.msgOK, "Successfully turned temperature compensation off");
				m_TempCompFalseOK = true;
			} catch (COMException ex) {
				LogMsg("TempComp Write", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			} catch (Exception ex) {
				LogMsg("TempComp Write", MessageLevel.msgError, EX_NET + ex.Message);
			}
		//Should raise an exception
		} else {
			try {
				m_Focuser.TempComp = true;
				LogMsg("TempComp Write", MessageLevel.msgIssue, "Temperature compensation is not available but no exception was raised when TempComp was set True");
			} catch (COMException ex) {
				switch (ex.ErrorCode) {
					case g_ExNotImplemented:
					case ErrorCodes.NotImplemented:
						LogMsg("TempComp Write", MessageLevel.msgOK, "COM Not implemented exception correctly raised as expected");
					default:
						LogMsg("TempComp Write", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			} catch (PropertyNotImplementedException ex) {
				LogMsg("TempComp Write", MessageLevel.msgOK, ".NET Not implemented exception correctly raised as expected");
			} catch (Exception ex) {
				LogMsg("TempComp Write", MessageLevel.msgError, EX_NET + ex.Message);
			}
		}

		//Restore original TempComp setting if possible
		try {
			m_Focuser.TempComp = m_TempComp;
		} catch {
		}

		//Temperature - Optional
		try {
			m_Temperature = m_Focuser.Temperature;
			switch (m_Temperature) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThanOrEqual
-50.0:
					//Probably a bad value
					LogMsg("Temperature", MessageLevel.msgWarning, "Temperature < -50.0, - possibly an issue, actual value: " + m_Temperature.ToString);
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThanOrEqual
50.0:
					//Probably a bad value
					LogMsg("Temperature", MessageLevel.msgWarning, "Temperature > 50.0, - possibly an issue, actual value: " + m_Temperature.ToString);
				default:
					LogMsg("Temperature", MessageLevel.msgOK, m_Temperature.ToString);
			}
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("Temperature", MessageLevel.msgOK, "COM Temperature is not supported by this focuser");
				default:
					LogMsg("Temperature", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("Temperature", MessageLevel.msgOK, ".NET Temperature is not supported by this focuser");
		} catch (Exception ex) {
			LogMsg("Temperature", MessageLevel.msgError, EX_NET + ex.Message);
		}
	}
	public override void CheckMethods()
	{
		//Halt - optional
		try {
			m_Focuser.Halt();
			LogMsg("Halt", MessageLevel.msgOK, "Focuser halted OK");
		} catch (COMException ex) {
			switch (ex.ErrorCode) {
				case g_ExNotImplemented:
				case ErrorCodes.NotImplemented:
					LogMsg("Halt", MessageLevel.msgOK, "COM Exception - Halt is not supported by this focuser");
				default:
					LogMsg("Halt", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (MethodNotImplementedException ex) {
			LogMsg("Halt", MessageLevel.msgOK, ".NET Exception - Halt is not supported by this focuser");
		} catch (Exception ex) {
			LogMsg("Halt", MessageLevel.msgError, EX_NET + ex.Message);
		}

		//Move - Required
		Status(StatusType.staTest, "Focuser Move");
		try {
			//Ensure that TempComp is false so that a move should be possible
			if (m_TempCompFalseOK)
				m_Focuser.TempComp = false;
			MoveFocuser();
		} catch (COMException ex) {
			LogMsg("Move", MessageLevel.msgError, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
		} catch (Exception ex) {
			LogMsg("Move", MessageLevel.msgError, EX_NET + ex.Message);
		}
		Status(StatusType.staTest, "");
		Status(StatusType.staAction, "");
		Status(StatusType.staStatus, "");

		//Move with TempComp True (if supported) - Should throw an error
		Status(StatusType.staTest, "Focuser Move");
		// Can set TempComp to True so cofirm that exception is thrown on Move
		if (m_TempCompTrueOK) {
			try {
				m_Focuser.TempComp = true;
				MoveFocuser();
				LogMsg("Move - TempComp True", MessageLevel.msgError, "TempComp is True but no exception is thrown by the Move Method - See TempComp help remarks");
			} catch (COMException ex) {
				LogMsg("Move - TempComp True", MessageLevel.msgOK, "COM Exception correctly raised as expected");
			// Conform 3.0.0.16 - Added InvalidOperationException and removed InvalidValueException
			//Catch ex As InvalidValueException
			//LogMsg("Move - TempComp True", MessageLevel.msgOK, ".NET InvalidValue Exception correctly raised as expected")
			} catch (ASCOM.InvalidOperationException ex) {
				LogMsg("Move - TempComp True", MessageLevel.msgOK, ".NET InvalidOperation Exception correctly raised as expected");
			} catch (System.InvalidOperationException ex) {
				LogMsg("Move - TempComp True", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException");
			} catch (Exception ex) {
				LogMsg("Move - TempComp True", MessageLevel.msgError, "Unexpected .NET Exception: " + ex.Message);
			}

			//Restore original TempComp value
			try {
				m_Focuser.TempComp = m_TempComp;
			} catch {
			}

			Status(StatusType.staTest, "");
			Status(StatusType.staAction, "");
			Status(StatusType.staStatus, "");
		}

	}

	private void MoveFocuser()
	{
		System.DateTime l_StartTime;
		System.DateTime l_EndTime;
		//This is an absolute focuser so position is an absolute value
		if (m_Absolute) {
			//Save the current absolute position
			m_PositionOrg = m_Focuser.Position;
			//Calculate an acceptable focus position
			m_Position = m_PositionOrg + (int)m_MaxStep / 10;
			//Move by 1/10 of the maximum focus distance out 
			//That would exceed maximum travel so move inwards instead
			if (m_Position >= m_MaxStep) {
				m_Position = m_PositionOrg - (int)m_MaxStep / 10;
				//Move by 1/10 of the maximum focus distance in
			}
			//Apply the MaxIncrement check
			if (System.Math.Abs(m_Position - m_PositionOrg) > m_MaxIncrement)
				m_Position = m_PositionOrg + m_MaxIncrement;
		//This is a relative focuser so position is the relative displacement
		} else {
			m_Position = (int)m_MaxIncrement / 10;
			//Apply the MaxIncrement check
			if (m_Position > m_MaxIncrement)
				m_Position = m_MaxIncrement;
		}

		//Confirm that the focuser is not moving
		//This is an issue as we are expecting the focuser to be not moving
		if (m_Focuser.IsMoving) {
			LogMsg("Move", MessageLevel.msgIssue, "Focuser is already moving before start of Move test, rest of test skipped");
		//Focuser not moving so proceed with the test
		} else {
			//Move the focuser
			if (m_Absolute) {
				LogMsg("Move", MessageLevel.msgComment, "Moving to position: " + m_Position.ToString);
			} else {
				LogMsg("Move", MessageLevel.msgComment, "Moving by: " + m_Position.ToString);
			}
			Status(StatusType.staAction, "Moving to new position");
			l_StartTime = Now;
			m_Focuser.Move(m_Position);
			l_EndTime = Now;
			//Assume a synchronous call
			if (l_EndTime.Subtract(l_StartTime).TotalMilliseconds > 1000) {
				//Confirm that ismoving is false
				//This is an issue as we are expecting the focuser to be not moving
				if (m_Focuser.IsMoving) {
					LogMsg("Move", MessageLevel.msgIssue, "Synchronous move expected but focuser is moving after return from Focuser.Move");
				} else {
					LogMsg("Move", MessageLevel.msgOK, "Synchronous move found");
				}
			//Assume an asynchronous call
			} else {
				Status(StatusType.staStatus, "Waiting for asynchronous move to complete");
				while ((m_Focuser.IsMoving & (!g_Stop))) {
					if (m_AbsolutePositionOK)
						Status(StatusType.staStatus, "Waiting for asynchronous move to complete, Position: " + m_Focuser.Position + " / " + m_Position);
					Application.DoEvents();
					WaitFor(500);
				}
				LogMsg("Move", MessageLevel.msgOK, "Asynchronous move found");
			}

			//Test outcome if absolute
			if (m_Absolute) {
				switch (m_Focuser.Position - m_Position) {
					case -2: // TODO: to +2
						//OK if within a small tolerance of expected value
						LogMsg("Move", MessageLevel.msgOK, "Absolute move OK");
					default:
						LogMsg("Move", MessageLevel.msgInfo, "Move was within " + m_Focuser.Position - m_Position + " counts of desired position");
				}
			} else {
				LogMsg("Move", MessageLevel.msgOK, "Relative move OK");
			}
		}
		Status(StatusType.staStatus, "");
		Status(StatusType.staAction, "Returning to original position: " + m_PositionOrg);
		LogMsg("Move", MessageLevel.msgInfo, "Returning to original position: " + m_PositionOrg);
		if (m_Absolute) {
			m_Focuser.Move(m_PositionOrg);
			//Return to original position
		} else {
			m_Focuser.Move(-m_Position);
			//Return to original position
		}
		Status(StatusType.staStatus, "Waiting for asynchronous move to complete");
		//Wait for asynchronous move to finish
		while (m_Focuser.IsMoving & (!g_Stop)) {
			if (m_AbsolutePositionOK)
				Status(StatusType.staStatus, "Waiting for asynchronous move to complete, Position: " + m_Focuser.Position + " / " + m_PositionOrg);
			Application.DoEvents();
			WaitFor(500);
		}
	}

	private override void CheckPerformance()
	{
		//Position
		if (true) {
			FocuserPerformanceTest(FocuserPropertyMethod.Position, "Position");
		} else {
			LogMsg("Position", MessageLevel.msgInfo, "Skipping test as property is not supported");
		}

		//IsMoving
		if (true) {
			FocuserPerformanceTest(FocuserPropertyMethod.IsMoving, "IsMoving");
		} else {
			LogMsg("IsMoving", MessageLevel.msgInfo, "Skipping test as property is not supported");
		}

		//IsMoving
		if (true) {
			FocuserPerformanceTest(FocuserPropertyMethod.Temperature, "Temperature");
		} else {
			LogMsg("Temperature", MessageLevel.msgInfo, "Skipping test as property is not supported");
		}
		Status(StatusType.staAction, "");
		Status(StatusType.staStatus, "");
		Status(StatusType.staTest, "");

	}

	private void FocuserPerformanceTest(FocuserPropertyMethod p_Type, string p_Name)
	{
		System.DateTime l_StartTime;
		double l_Count;
		double l_LastElapsedTime;
		double l_ElapsedTime;
		float l_Single;
		bool l_Boolean;
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
					case FocuserPropertyMethod.IsMoving:
						l_Boolean = m_Focuser.IsMoving;
					case FocuserPropertyMethod.Position:
						l_Single = m_Focuser.Position;
					case FocuserPropertyMethod.Temperature:
						l_Rate = m_Focuser.Temperature;
					default:
						LogMsg(p_Name, MessageLevel.msgError, "FocuserPerformanceTest: Unknown test type " + p_Type.ToString);
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
