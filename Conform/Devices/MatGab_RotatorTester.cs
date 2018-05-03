internal class RotatorTester : DeviceTesterBaseClass
{

	#Region "Variables and Constants"
	const double ROTATOR_WAIT_LIMIT = 30.0;
	const double ROTATOR_OK_TOLERANCE = 1.0;

	const double ROTATOR_INFO_TOLERANCE = 2.0;
	//Rotator variables
	private bool m_CanReadIsMoving;
	private bool m_CanReadPosition;
	private bool m_CanReadTargetPosition;
	private bool m_CanReadStepSize;
	private bool m_CanReverse;
	private bool m_IsMoving;
	private float m_TargetPosition;
	private float m_RotatorStepSize;
	private float m_RotatorPosition;
	private bool m_Reverse;

	private bool m_LastMoveWasAsync;
	#If DEBUG Then
		#Else
	private object m_Rotator;
		#End If
	private ASCOM.DeviceInterface.IRotatorV2 m_Rotator;
	#End Region

	#Region "Enums"
	private enum RotatorPropertyMethod
	{
		Position,
		TargetPosition,
		StepSize,
		Move,
		MoveAbsolute,
		IsMoving
	}
	#End Region

	#Region "New and Dispose"
	private RotatorTester()
	{
		base.New();
		this.HasCanProperties = true;
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
					m_Rotator.Connected = false;
				} catch {
				}
				try {
					Marshal.ReleaseComObject(m_Rotator);
				} catch {
				}
				m_Rotator = null;
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

		switch (g_RotatorProgID) {
			default:
				//I'm using the simulator values as the defaults since it is the reference platform
				g_ExNotImplemented = 0x80040400;
				g_ExInvalidValue1 = 0x80040405;
				g_ExInvalidValue2 = 0x80040405;
				g_ExNotSet1 = 0x80040403;
		}
		base.CheckInitialise(g_RotatorProgID);

	}
	private override void CheckAccessibility()
	{
		object l_IRotator;
		DriverAccess.Rotator l_DriverAccessRotator;
		object l_DeviceObject = null;
		base.CheckAccessibility(g_RotatorProgID, DeviceType.Rotator);
		//Try early binding
		l_IRotator = null;
		try {
			l_DeviceObject = CreateObject(g_RotatorProgID);
			l_IRotator = (ASCOM.Interface.IRotator)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with interface IRotator");
			try {
				l_IRotator.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver exposes interface IRotator");
				l_IRotator.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface IRotator");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface IRotator");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_IRotator);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_IRotator = null;
			GC.Collect();
		}

		try {
			l_DeviceObject = CreateObject(g_RotatorProgID);
			l_IRotator = (ASCOM.DeviceInterface.IRotatorV2)l_DeviceObject;
			LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with interface IRotatorV2");
			try {
				l_IRotator.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver exposes interface IRotatorV2");
				l_IRotator.Connected = false;
			} catch (Exception ex) {
				LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface IRotatorV2");
			}
		} catch (Exception ex) {
			LogMsg("AccessChecks", MessageLevel.msgInfo, "Driver does not expose interface IRotatorV2");
		} finally {
			//Clean up
			try {
				Marshal.ReleaseComObject(l_IRotator);
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DeviceObject);
			} catch {
			}
			l_DeviceObject = null;
			l_IRotator = null;
			GC.Collect();
		}

		//Try client access toolkit
		l_DriverAccessRotator = null;
		try {
			l_DriverAccessRotator = new ASCOM.DriverAccess.Rotator(g_RotatorProgID);
			LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
			try {
				l_DriverAccessRotator.Connected = true;
				LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
				l_DriverAccessRotator.Connected = false;
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
				l_DriverAccessRotator.Dispose();
			} catch {
			}
			try {
				Marshal.ReleaseComObject(l_DriverAccessRotator);
			} catch {
			}
			l_DriverAccessRotator = null;
			GC.Collect();
		}

	}
	#If DEBUG Then
	private override void CreateDevice()
	{
		m_Rotator = CreateObject(g_RotatorProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Rotator object");
		#Else
		m_Rotator = new ASCOM.DriverAccess.Rotator(g_RotatorProgID);
		LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Rotator to get a Rotator object");
		#End If
		g_Stop = false;
		//connected OK so clear stop flag to allow other tests to run
	}
	private override bool Connected {
		get { Connected = m_Rotator.Connected; }
		set { m_Rotator.Connected = value; }
	}
	private override void CheckCommonMethods()
	{
		base.CheckCommonMethods(m_Rotator, DeviceType.Rotator);
	}


	public override void ReadCanProperties()
	{
		try {
			m_CanReverse = m_Rotator.CanReverse;
			LogMsg("CanReverse", MessageLevel.msgOK, m_CanReverse.ToString);
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("CanReverse", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("CanReverse", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("CanReverse", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("CanReverse", ex.Message, EX_NET + ex.Message);
		}
	}

	public override void PreRunCheck()
	{
		System.DateTime l_Now;
		//Get the rotator into a standard state
		g_Stop = true;
		try {
			m_Rotator.Halt();
		} catch {
		}
		//Stop any movement
		l_Now = Now;
		//Confirm that rotator is not moving or wait for it to stop
		try {
			Status(StatusType.staAction, "Waiting up to " + ROTATOR_WAIT_LIMIT + " seconds for rotator to stop moving");
			do {
				WaitFor(500);
				Status(StatusType.staStatus, Now.Subtract(l_Now).TotalSeconds + "/" + ROTATOR_WAIT_LIMIT);
			} while (!((!m_Rotator.IsMoving) | (Now.Subtract(l_Now).TotalSeconds > ROTATOR_WAIT_LIMIT)));
			//Rotator is stoped so OK
			if (!m_Rotator.IsMoving) {
				g_Stop = false;
				//Clear stop flag to allow other tests to run
			//Report error message and don't do other tests
			} else {
				LogMsg("Pre-run Check", MessageLevel.msgError, "Rotator still moving after " + ROTATOR_WAIT_LIMIT + "seconds, IsMoving stuck on?");
			}
			LogMsg("Pre-run Check", MessageLevel.msgOK, "Rotator is stationary");
		} catch (Exception ex) {
			//Don't report errors at this point
		}

	}
	private override void CheckProperties()
	{
		//IsMoving
		try {
			m_CanReadIsMoving = false;
			m_IsMoving = m_Rotator.IsMoving;
			m_CanReadIsMoving = true;
			// Can read OK, doesn't generate an exception
			if (m_IsMoving) {
				LogMsg("IsMoving", MessageLevel.msgError, "IsMoving is True before any movement has been commanded!");
				LogMsg("IsMoving", MessageLevel.msgInfo, "Further tests have been skipped");
				g_Stop = true;
			} else {
				LogMsg("IsMoving", MessageLevel.msgOK, m_IsMoving.ToString);
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("IsMoving", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				ExTest("IsMoving", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (PropertyNotImplementedException ex) {
			LogMsg("IsMoving", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("IsMoving", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//Position (Single)
		m_RotatorPosition = RotatorPropertyTestSingle(RotatorPropertyMethod.Position, "Position", 0.0, 359.9999999999);
		if (TestStop())
			return;

		//TargetPosition (Single)
		m_TargetPosition = RotatorPropertyTestSingle(RotatorPropertyMethod.TargetPosition, "TargetPosition", 0.0, 359.9999999999);
		if (TestStop())
			return;

		//StepSize (Single)
		m_RotatorStepSize = RotatorPropertyTestSingle(RotatorPropertyMethod.StepSize, "StepSize", 0.0, 359.9999999999);
		if (TestStop())
			return;
		//Reverse Read (Boolean)

		//Reverse Read (Boolean)
		try {
			m_Reverse = m_Rotator.Reverse;
			if (m_CanReverse) {
				LogMsg("Reverse Read", MessageLevel.msgOK, m_Reverse.ToString);
			} else {
				LogMsg("Reverse Read", MessageLevel.msgError, "CanReverse is false but no exception generated");
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				//Should not have generated an exception
				if (m_CanReverse) {
					LogMsg("Reverse Read", MessageLevel.msgError, "CanReverse is true but COM not implemented error thrown");
				// Can not reverse so should have generated an error
				} else {
					LogMsg("Reverse Read", MessageLevel.msgOK, NOT_IMP_COM);
				}
			} else {
				//Should not have generated an exception
				if (m_CanReverse) {
					LogMsg("Reverse Read", MessageLevel.msgError, "CanReverse is true but COM not implemented error thrown");
				// Can not reverse so should have generated an error
				} else {
					ExTest("Reverse Read", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			}
		} catch (PropertyNotImplementedException ex) {
			//Should not have generated an exception
			if (m_CanReverse) {
				LogMsg("Reverse Read", MessageLevel.msgError, "CanReverse is true but .NET not implemented error thrown");
			// Can not reverse so should have generated an error
			} else {
				LogMsg("Reverse Read", MessageLevel.msgOK, NOT_IMP_NET);
			}
		} catch (Exception ex) {
			//Should not have generated an exception
			if (m_CanReverse) {
				LogMsg("Reverse Read", MessageLevel.msgError, "CanReverse is true but .NET not implemented error thrown");
			// Can not reverse so should have generated an error
			} else {
				ExTest("Reverse Read", ex.Message, EX_NET + ex.Message);
			}
		}
		if (TestStop())
			return;

		//Reverse Write (Boolean)
		try {
			//Try and set reverse to the opposite state
			if (m_Reverse) {
				m_Rotator.Reverse = false;
			} else {
				m_Rotator.Reverse = true;
			}

			m_Rotator.Reverse = m_Reverse;
			//Restore original value

			if (m_CanReverse) {
				LogMsg("Reverse Write", MessageLevel.msgOK, "Reverse state successfully changed and restored");
			} else {
				LogMsg("Reverse Write", MessageLevel.msgError, "CanReverse is false but no exception generated");
			}
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				//Should not have generated an exception
				if (m_CanReverse) {
					LogMsg("Reverse Write", MessageLevel.msgError, "CanReverse is true but COM not implemented error thrown");
				// Can not reverse so should have generated an error
				} else {
					LogMsg("Reverse Write", MessageLevel.msgOK, NOT_IMP_COM);
				}
			} else {
				//Should not have generated an exception
				if (m_CanReverse) {
					LogMsg("Reverse Write", MessageLevel.msgError, "CanReverse is true but COM not implemented error thrown");
				// Can not reverse so should have generated an error
				} else {
					ExTest("Reverse Write", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			}
		} catch (PropertyNotImplementedException ex) {
			//Should not have generated an exception
			if (m_CanReverse) {
				LogMsg("Reverse Write", MessageLevel.msgError, "CanReverse is true but COM not implemented error thrown");
			// Can not reverse so should have generated an error
			} else {
				LogMsg("Reverse Write", MessageLevel.msgOK, NOT_IMP_NET);
			}
		} catch (Exception ex) {
			//Should not have generated an exception
			if (m_CanReverse) {
				LogMsg("Reverse Write", MessageLevel.msgError, "CanReverse is true but COM not implemented error thrown");
			// Can not reverse so should have generated an error
			} else {
				ExTest("Reverse Write", ex.Message, EX_NET + ex.Message);
			}
		}

	}
	private float RotatorPropertyTestSingle(RotatorPropertyMethod p_Type, string p_Name, float p_Min, float p_Max)
	{
		try {
			RotatorPropertyTestSingle = 0.0;
			switch (p_Type) {
				case RotatorPropertyMethod.Position:
					m_CanReadPosition = false;
					RotatorPropertyTestSingle = m_Rotator.Position;
					//Can read position OK, doesn't generate an exception
					m_CanReadPosition = true;
				case RotatorPropertyMethod.StepSize:
					m_CanReadStepSize = false;
					RotatorPropertyTestSingle = m_Rotator.StepSize;
					m_CanReadStepSize = true;
				case RotatorPropertyMethod.TargetPosition:
					m_CanReadTargetPosition = false;
					RotatorPropertyTestSingle = m_Rotator.TargetPosition;
					m_CanReadTargetPosition = true;
				default:
					LogMsg(p_Name, MessageLevel.msgError, "RotatorPropertyTestSingle: Unknown test type - " + p_Type.ToString);
			}
			//Successfully retrieved a value
			switch (RotatorPropertyTestSingle) {
				case  // ERROR: Case labels with binary operators are unsupported : LessThan
p_Min:
					//Lower than minimum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + RotatorPropertyTestSingle.ToString);
				case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
p_Max:
					//Higher than maximum value
					LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + RotatorPropertyTestSingle.ToString);
				default:
					//OK value
					LogMsg(p_Name, MessageLevel.msgOK, RotatorPropertyTestSingle.ToString);
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

	private override void CheckMethods()
	{
		LogMsg("CheckMethods", MessageLevel.msgDebug, "Rotator is connected: " + m_Rotator.Connected.ToString);
		//Halt
		try {
			m_Rotator.Halt();
			LogMsg("Halt", MessageLevel.msgOK, "Halt command successful");
		} catch (COMException ex) {
			if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
				LogMsg("Halt", MessageLevel.msgOK, NOT_IMP_COM);
			} else {
				LogMsg("Halt", MessageLevel.msgInfo, "Only not implemented exceptions are valid here as any exception is used to disable client halt buttons");
				ExTest("Halt", ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
			}
		} catch (MethodNotImplementedException ex) {
			LogMsg("Halt", MessageLevel.msgOK, NOT_IMP_NET);
		} catch (Exception ex) {
			ExTest("Halt", ex.Message, EX_NET + ex.Message);
		}
		if (TestStop())
			return;

		//MoveAbsolute(Single)
		RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 45.0, "");
		if (TestStop())
			return;
		RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 135.0, "");
		if (TestStop())
			return;
		RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 225.0, "");
		if (TestStop())
			return;
		RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 315.0, "");
		if (TestStop())
			return;
		RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", -405.0, "movement to large negative angle -405 degrees");
		if (TestStop())
			return;
		RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 405.0, "movement to large positive angle 405 degrees");
		if (TestStop())
			return;

		//Move(Single)
		RelativeMoveTest(10.0);
		if (TestStop())
			return;
		RelativeMoveTest(40.0);
		if (TestStop())
			return;
		RelativeMoveTest(130.0);
		if (TestStop())
			return;
		RotatorMoveTest(RotatorPropertyMethod.Move, "Move", -375.0, "movement to large negative angle -375 degrees");
		if (TestStop())
			return;
		RotatorMoveTest(RotatorPropertyMethod.Move, "Move", 375.0, "movement to large positive angle 375 degrees");
		if (TestStop())
			return;

	}
	private void RotatorMoveTest(RotatorPropertyMethod p_type, string p_Name, float p_Value, string p_ExpectErrorMsg)
	{
		float l_RotatorStartPosition;
		double l_OKLimit;
		double l_PositionOffset;

		LogMsg("RotatorMoveTest", MessageLevel.msgDebug, "Start value, position: " + p_Value.ToString("0.000") + " " + m_Rotator.Position.ToString("0.000"));
		//Move to requested position
		try {
			switch (p_type) {
				case RotatorPropertyMethod.Move:
					//Get us to a startig point of 10 degrees
					if (m_CanReadPosition) {
						l_RotatorStartPosition = m_Rotator.Position;
					}
					m_Rotator.Move(p_Value);
				case RotatorPropertyMethod.MoveAbsolute:
					l_RotatorStartPosition = 0.0;
					m_Rotator.MoveAbsolute(p_Value);
				default:
					LogMsg(p_Name, MessageLevel.msgError, "RotatorMoveTest: Unknown test type - " + p_type.ToString);
			}
			RotatorWait(p_type, p_Name, p_Value, l_RotatorStartPosition);
			//Async move
			if (m_LastMoveWasAsync) {
				switch (p_type) {
					case RotatorPropertyMethod.Move:
						if (m_CanReadPosition) {
							LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous move successful - moved by " + p_Value + " degrees to: " + m_Rotator.Position + " degrees");
						} else {
							LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous move successful");
						}
					case RotatorPropertyMethod.MoveAbsolute:
						if (m_CanReadPosition) {
							LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous move successful to: " + m_Rotator.Position + " degrees");
						} else {
							LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous move successful");
						}
				}
			//Sync move
			} else {
				if (m_CanReadPosition) {
					LogMsg(p_Name, MessageLevel.msgOK, "Synchronous move successful to: " + m_Rotator.Position + " degrees");
				} else {
					LogMsg(p_Name, MessageLevel.msgOK, "Synchronous move successful");
				}
			}

			//Now test whether we got to where we expected to go
			if (m_CanReadPosition) {
				if (m_CanReadStepSize) {
					l_OKLimit = 1.1 * m_RotatorStepSize;
					// Set to 110% of stepsize to allow tolerance on reporting within 1 setp of required location
				} else {
					l_OKLimit = ROTATOR_OK_TOLERANCE;
				}
				LogMsg(p_Name + "1", MessageLevel.msgDebug, "Position, value, start, tolerance: " + m_Rotator.Position.ToString("0.000") + " " + p_Value.ToString("0.000") + " " + l_RotatorStartPosition.ToString("0.000") + " " + l_OKLimit.ToString("0.000"));
				if (m_Rotator.Position < 0.0)
					LogMsg(p_Name, MessageLevel.msgInfo, "Rotator supports angles < 0.0");
				if (m_Rotator.Position > 360.0)
					LogMsg(p_Name, MessageLevel.msgInfo, "Rotator supports angles > 360.0");

				// Calculate the position offset from the required position
				l_PositionOffset = Math.Abs((720.0 + m_Rotator.Position - (p_Value + l_RotatorStartPosition)) % 360.0);
				// Account for rotator positions that report < 0.0 or > 360.0 degrees

				switch (l_PositionOffset) {
					case 0.0:
						LogMsg(p_Name, MessageLevel.msgOK, "Rotator is at the expected position");
						break; // TODO: might not be correct. Was : Exit Select

					case 0.0: // TODO: to l_OKLimit
						LogMsg(p_Name, MessageLevel.msgOK, "Rotator is within " + l_OKLimit.ToString("0.000") + IIf(l_PositionOffset <= 1.0, " degree", " degrees") + " of the expected position");
						break; // TODO: might not be correct. Was : Exit Select

					case 0.0: // TODO: to ROTATOR_INFO_TOLERANCE
						LogMsg(p_Name, MessageLevel.msgInfo, "Rotator is " + l_PositionOffset.ToString("0.000") + " degrees from expected position");
						break; // TODO: might not be correct. Was : Exit Select

					default:
						LogMsg(p_Name, MessageLevel.msgIssue, "Rotator is " + l_PositionOffset.ToString("0.000") + " degrees from expected position, which is more than the conformance value of " + ROTATOR_INFO_TOLERANCE.ToString("0.0") + " degrees");
				}

			}
		} catch (COMException ex) {
			if (p_ExpectErrorMsg != "") {
				LogMsg(p_Name, MessageLevel.msgOK, "Exception thrown as expected on " + p_ExpectErrorMsg);
			} else {
				if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented) {
					LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_COM);
				} else {
					ExTest(p_Name, ex.Message, EX_COM + ex.Message + " " + Hex(ex.ErrorCode));
				}
			}
		} catch (PropertyNotImplementedException ex) {
			if (p_ExpectErrorMsg != "") {
				LogMsg(p_Name, MessageLevel.msgOK, "Exception thrown as expected on " + p_ExpectErrorMsg);
			} else {
				LogMsg(p_Name, MessageLevel.msgOK, NOT_IMP_NET);
			}
		} catch (Exception ex) {
			if (p_ExpectErrorMsg != "") {
				LogMsg(p_Name, MessageLevel.msgOK, "Exception thrown as expected on " + p_ExpectErrorMsg);
			} else {
				ExTest(p_Name, ex.Message, EX_NET + ex.Message);
			}
		}
		Status(StatusType.staAction, "");
		Status(StatusType.staStatus, "");
		Status(StatusType.staTest, "");
	}
	private void RotatorWait(RotatorPropertyMethod p_type, string p_Name, float p_value, float p_RotatorStartPosition)
	{
		//Can read ismoving so test for async and sync behaviour
		if (m_CanReadIsMoving) {
			if (m_Rotator.IsMoving) {
				Status(StatusType.staTest, p_Name + " test");
				Status(StatusType.staAction, "Waiting for move to complete");
				do {
					WaitFor(500);
					//Only do this if position doesn't generate an exception
					if (m_CanReadPosition) {
						switch (p_type) {
							case RotatorPropertyMethod.Move:
								Status(StatusType.staStatus, System.Math.Abs(m_Rotator.Position - p_RotatorStartPosition) + "/" + p_value + " releative");
							case RotatorPropertyMethod.MoveAbsolute:
								Status(StatusType.staStatus, System.Math.Abs(m_Rotator.Position - p_RotatorStartPosition) + "/" + System.Math.Abs(p_value - p_RotatorStartPosition) + " absolute");
						}
					}
					Application.DoEvents();
				} while (!(!m_Rotator.IsMoving));
				Status(StatusType.staAction, "");
				m_LastMoveWasAsync = true;
			} else {
				m_LastMoveWasAsync = false;
			}
		//Can only test for synchroous move
		} else {
			m_LastMoveWasAsync = false;
		}
	}
	private void RelativeMoveTest(float p_RelativeStepSize)
	{
		float l_Target;
		if (m_CanReadPosition) {
			//Set a value that should succeed OK
			if (m_Rotator.Position < p_RelativeStepSize) {
				l_Target = p_RelativeStepSize;
			} else {
				l_Target = -p_RelativeStepSize;
			}
		} else {
			l_Target = p_RelativeStepSize;
		}
		RotatorMoveTest(RotatorPropertyMethod.Move, "Move", l_Target, "");
		if (TestStop())
			return;
		RotatorMoveTest(RotatorPropertyMethod.Move, "Move", -l_Target, "");
		if (TestStop())
			return;
		//Should now be back where we started

	}

	public override void CheckPerformance()
	{
		//Position
		if (m_CanReadPosition) {
			RotatorPerformanceTest(RotatorPropertyMethod.Position, "Position");
		} else {
			LogMsg("Position", MessageLevel.msgInfo, "Skipping test as property is not supported");
		}

		//TargetPosition
		if (m_CanReadTargetPosition) {
			RotatorPerformanceTest(RotatorPropertyMethod.TargetPosition, "TargetPosition");
		} else {
			LogMsg("TargetPosition", MessageLevel.msgInfo, "Skipping test as property is not supported");
		}

		//StepSize
		if (m_CanReadStepSize) {
			RotatorPerformanceTest(RotatorPropertyMethod.StepSize, "StepSize");
		} else {
			LogMsg("StepSize", MessageLevel.msgInfo, "Skipping test as property is not supported");
		}

		//IsMoving
		if (m_CanReadIsMoving) {
			RotatorPerformanceTest(RotatorPropertyMethod.IsMoving, "IsMoving");
		} else {
			LogMsg("IsMoving", MessageLevel.msgInfo, "Skipping test as property is not supported");
		}
	}
	private void RotatorPerformanceTest(RotatorPropertyMethod p_Type, string p_Name)
	{
		System.DateTime l_StartTime;
		double l_Count;
		double l_LastElapsedTime;
		double l_ElapsedTime;
		float l_Single;
		bool l_Boolean;
		double l_Rate;
		Status(StatusType.staAction, p_Name);
		try {
			l_StartTime = Now;
			l_Count = 0.0;
			l_LastElapsedTime = 0.0;
			do {
				l_Count += 1.0;
				switch (p_Type) {
					case RotatorPropertyMethod.Position:
						l_Single = m_Rotator.Position;
					case RotatorPropertyMethod.TargetPosition:
						l_Single = m_Rotator.TargetPosition;
					case RotatorPropertyMethod.StepSize:
						l_Single = m_Rotator.StepSize;
					case RotatorPropertyMethod.IsMoving:
						l_Boolean = m_Rotator.IsMoving;
					default:
						LogMsg(p_Name, MessageLevel.msgError, "RotatorPerformanceTest: Unknown test type " + p_Type.ToString);
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
