''' <summary>
''' The device specific commands and expected responses to be used by Conform when testing the 
''' CommandXXX commands.
''' </summary>
''' <remarks></remarks>
Public Class ConformCommandStrings
    Private cmdBlind, cmdBool, cmdString As String
    Private rtnString As String, rtnBool As Boolean

    ''' <summary>
    ''' Set all Conform CommandXXX commands and expected responses in one call
    ''' </summary>
    ''' <param name="CommandString">String to be sent through CommandString method</param>
    ''' <param name="ReturnString">Expected return value from CommandString command</param>
    ''' <param name="CommandBlind">String to be sent through CommandBling method</param>
    ''' <param name="CommandBool">Expected return value from CommandBlind command</param>
    ''' <param name="ReturnBool">Expected boolean response from CommandBool command</param>
    ''' <remarks>To suppress a Command XXX test, set the command and return values to Nothing (VB) 
    ''' or null (C#). To accept any response to a command just set the return value to Nothing or null
    ''' as appropriate.</remarks>
    Sub New(ByVal CommandString As String, ByVal ReturnString As String, ByVal CommandBlind As String, ByVal CommandBool As String, ByVal ReturnBool As Boolean)
        Me.CommandString = CommandString
        Me.CommandBlind = CommandBlind
        Me.CommandBool = CommandBool
        Me.ReturnString = ReturnString
        Me.ReturnBool = ReturnBool
    End Sub

    ''' <summary>
    ''' Create a new CommandStrings object with all fields set to Nothing (VB) / null (C#), use 
    ''' other properties to set commands and expected return values.
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()
        Me.CommandString = Nothing
        Me.CommandBlind = Nothing
        Me.CommandBool = Nothing
        Me.ReturnString = Nothing
        Me.ReturnBool = False
    End Sub

    ''' <summary>
    ''' Set the command to be sent by the Conform CommandString test
    ''' </summary>
    ''' <value>Device command</value>
    ''' <returns>String device command</returns>
    ''' <remarks></remarks>
    Property CommandString() As String
        Get
            CommandString = cmdString
        End Get
        Set(ByVal value As String)
            cmdString = value
        End Set
    End Property

    ''' <summary>
    ''' Set the expected return value from the CommandString command
    ''' </summary>
    ''' <value>Device response</value>
    ''' <returns>String device response</returns>
    ''' <remarks></remarks>
    Property ReturnString() As String
        Get
            ReturnString = rtnString
        End Get
        Set(ByVal value As String)
            rtnString = value
        End Set
    End Property

    ''' <summary>
    ''' Set the command to be sent by the Conform CommandBlind test
    ''' </summary>
    ''' <value>Device command</value>
    ''' <returns>String device command</returns>
    ''' <remarks></remarks>
    Property CommandBlind() As String
        Get
            CommandBlind = cmdBlind
        End Get
        Set(ByVal value As String)
            cmdBlind = value
        End Set
    End Property

    ''' <summary>
    ''' Set the command to be sent by the Conform CommandBool test
    ''' </summary>
    ''' <value>Device command</value>
    ''' <returns>String device command</returns>
    ''' <remarks></remarks>
    Property CommandBool() As String
        Get
            CommandBool = cmdBool
        End Get
        Set(ByVal value As String)
            cmdBool = value
        End Set
    End Property

    ''' <summary>
    ''' Set the expected return value from the CommandBool command
    ''' </summary>
    ''' <value>Device response</value>
    ''' <returns>Boolean device response</returns>
    ''' <remarks></remarks>
    Property ReturnBool() As Boolean
        Get
            ReturnBool = rtnBool
        End Get
        Set(ByVal value As Boolean)
            rtnBool = value
        End Set
    End Property
End Class