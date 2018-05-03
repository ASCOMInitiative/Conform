Friend NotInheritable Class NativeMethods
#Region "External API Calls"
    ' Possible values for process error modes
    Friend Enum ErrorModes As UInteger
        ERRORMODE_SYSTEM_DEFAULT = &H0
        ERRORMODE_SEM_FAILCRITICALERRORS = &H1
        ERRORMODE_SEM_NOALIGNMENTFAULTEXCEPT = &H4
        ERRORMODE_SEM_NOGPFAULTERRORBOX = &H2
        ERRORMODE_SEM_NOOPENFILEERRORBOX = &H8000
        ERRORMODE_NO_WINDOWS_ERROR_DIALOGUES = &H8003
    End Enum

    ' Functions to get and set the process error mode.
    ' Default behaviour is for low level Windows errors to present a Windows dialogue to the user rather than returning the 
    ' exception to the calling application. SetErroMode allows this behaviour to be changed, GetErrorMode retrieves the current value.
    Friend Declare Function SetErrorMode Lib "kernel32" (ByVal uMode As ErrorModes) As ErrorModes
    Friend Declare Function GetErrorMode Lib "kernel32" () As ErrorModes
#End Region

End Class
