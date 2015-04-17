''' <summary>
''' Provides facilities for an application to determine the number and values of
''' arguments and switches in a string (normally the arguments part of the command
''' used to start the application).
''' </summary>
''' <remarks>
''' The format of the argument string passed to the <code>CreateCommandLineParser</code> method
''' is as follows:
'''
''' <pre>
'''   [&lt;argument&gt; | &lt;switch&gt;] [&lt;sep&gt; (&lt;argument&gt; | &lt;switch&gt;)]...
''' </pre>
'''
''' ie, there is a sequence of arguments or switches, separated by separator characters. The
''' separator Character is specified in the call to <code>CreateCommandLineParser</code>.
'''
''' Arguments that contain the separator Character must be enclosed in double quotes. Double quotes
''' appearing within an argument must be repeated.
'''
''' Switches have the following format:
'''
''' <pre>
'''   ( "/" | "-")&lt;identifier&gt; [":"&lt;switchValue&gt;]
''' </pre>
'''
''' ie the switch starts with a forward slash or a hyphen followed by an identifier, and
''' optionally followed by a colon and the switch Value. Switch identifiers are not
''' case-sensitive. Switch values that contain the separator Character must be enclosed in
''' double quotes. Double quotes appearing within a switch Value must be repeated.
'''
''' Examples (these examples use a space as the separator Character):
''' <pre>
'''   anArgument -sw1 anotherArg -sw2:42
''' </pre>
''' <pre>
'''   "C:\Program Files\MyApp\myapp.ini" -out:C:\MyLogs\myapp.Log
''' </pre>
''' </remarks>
Public NotInheritable Class CommandLineParser

    '@================================================================================
    ' Interfaces
    '@================================================================================

    '@================================================================================
    ' Events
    '@================================================================================

    '@================================================================================
    ' Constants
    '@================================================================================

    '@================================================================================
    ' Enums
    '@================================================================================

    '@================================================================================
    ' Types
    '@================================================================================

    ''' <summary>
    ''' Contains details of a command line switch.
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure SwitchEntry
        ''' <summary>
        ''' The switch identifier.
        ''' </summary>
        ''' <remarks></remarks>
        Public Switch As String

        ''' <summary>
        ''' The switch Value.
        ''' </summary>
        ''' <remarks></remarks>
        Public Value As String
    End Structure

    '@================================================================================
    ' Member variables
    '@================================================================================

    Private mArgs As List(Of String) = New List(Of String)
    Private mSwitches As List(Of SwitchEntry) = New List(Of SwitchEntry)
    Private mSep As String
    Private mCommandLine As String

    '@================================================================================
    ' Constructors
    '@================================================================================

    Private Sub New()
    End Sub

    ''' <summary>
    '''  Initialises a new instance of the <see cref="CommandLineParser"></see> class.
    ''' </summary>
    ''' <param name="commandLine">The command line arguments to be parsed. For a Visual Basic 6 program,
    '''  this Value may be obtained using the <code>Command</code> function.</param>
    ''' <param name="separator">A single Character used as the separator between command line arguments.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal commandLine As String, ByVal separator As String)
        mCommandLine = commandLine.Trim
        mSep = separator
        getArgs()
    End Sub

    '@================================================================================
    ' XXXX Interface Members
    '@================================================================================

    '@================================================================================
    ' XXXX Event Handlers
    '@================================================================================

    '@================================================================================
    ' Properties
    '@================================================================================

    ''' <summary>
    ''' Gets the nth argument, where n is the value of the <paramref>i</paramref> parameter.
    ''' </summary>
    ''' <param name="i"> The number of the argument to be returned. The first argument is number 0.</param>
    ''' <value></value>
    ''' <returns>A String Value containing the nth argument, where n is the value of the <paramref>i</paramref> parameter.</returns>
    ''' <remarks>If the requested argument has not been supplied, an empty string is returned.</remarks>
    Public ReadOnly Property Arg(ByVal i As Integer) As String
        Get
            Try
                Return mArgs(i)
            Catch ex As Exception
                Return String.Empty
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Gets an array of strings containing the arguments.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A String array containing the arguments.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Args() As String()
        Get
            Return mArgs.ToArray
        End Get
    End Property

    ''' <summary>
    ''' Gets the number of arguments.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The number of arguments.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property NumberOfArgs() As Integer
        Get
            Return mArgs.Count
        End Get
    End Property

    ''' <summary>
    ''' Gets the number of switches.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The number of switches.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property NumberOfSwitches() As Integer
        Get
            Return mSwitches.Count
        End Get
    End Property

    ''' <summary>
    ''' Gets a value which indicates whether the specified switch was included.
    ''' </summary>
    ''' <param name="s">The identifier of the switch whose inclusion is to be indicated.</param>
    ''' <value></value>
    ''' <returns>If the specified switch was included, <code>True</code> is
    ''' returned. Otherwise <code>False</code> is returned.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsSwitchSet(ByVal s As String) As Boolean
        Get
            For Each switchEntry As SwitchEntry In mSwitches
                If switchEntry.Switch.ToUpper = s.ToUpper Then Return True
            Next
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Gets an array of <code>SwitchEntry</code>s containing the
    ''' switch identifiers and values.
    ''' </summary>
    ''' <value></value>
    ''' <returns>An array of <code>SwitchEntry</code>s containing the
    ''' switch identifiers and values.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Switches() As SwitchEntry()
        Get
            Return mSwitches.ToArray
        End Get
    End Property

    ''' <summary>
    ''' Gets the value of the specified switch.
    ''' </summary>
    ''' <param name="s">The identifier of the switch whose value is to be returned.</param>
    ''' <value></value>
    ''' <returns>A String containing the value for the specified switch.</returns>
    ''' <remarks>If the requested switch has not been supplied, or no value
    ''' was supplied for the switch, an empty string is returned.</remarks>
    Public ReadOnly Property SwitchValue(ByVal s As String) As String
        Get
            For Each switchEntry As SwitchEntry In mSwitches
                If switchEntry.Switch.ToUpper = s.ToUpper Then Return switchEntry.Value
            Next
            Return String.Empty
        End Get
    End Property

    '@================================================================================
    ' Methods
    '@================================================================================

    '@================================================================================
    ' Helper Functions
    '@================================================================================


    Private Function ContainsUnbalancedQuotes(ByVal inString As String) As Boolean
        Dim pos = inString.Length - 2 ' exclude the last char
        If pos = -1 Then Return False
        pos = inString.LastIndexOf("""", pos)
        Dim balanced = False
        Do While pos <> -1
            balanced = Not balanced
            If pos = 0 Then Exit Do
            pos = inString.LastIndexOf("""", pos - 1)
        Loop
        Return balanced
    End Function

    Private Sub getArgs()
        If mCommandLine = "" Then Exit Sub

        Dim inQuotedArg As Boolean
        Dim unbalancedQuotes As Boolean
        Dim quotedArg As String = ""

        For Each argument In mCommandLine.Split({mSep}, StringSplitOptions.RemoveEmptyEntries)
            argument = argument.Replace("""""", """")
            If Not inQuotedArg Then
                If argument = "" And mSep = "" Then
                Else
                    If argument.Substring(0, 1) = """" Or _
                        ((argument.Substring(0, 1) = "/" Or argument.Substring(0, 1) = "-") And argument.IndexOf(":""") <> -1) _
                    Then
                        inQuotedArg = True
                        quotedArg = ReplaceFirst(argument, """", "") ' Right$(argument, Len(argument) - 1)
                        If ContainsUnbalancedQuotes(quotedArg) Then
                            unbalancedQuotes = Not unbalancedQuotes
                        End If
                        If quotedArg.Substring(quotedArg.Length - 1, 1) = """" And Not unbalancedQuotes Then
                            ' the ending quote is also in this arg
                            inQuotedArg = False
                            quotedArg = quotedArg.Substring(0, quotedArg.Length - 1)
                            If quotedArg.Substring(0, 1) = "/" Or _
                                quotedArg.Substring(0, 1) = "-" _
                            Then
                                setSwitch(quotedArg.Substring(1))
                            Else
                                mArgs.Add(quotedArg)
                            End If
                        Else
                            If argument.Substring(argument.Length - 1) = """" Then
                                unbalancedQuotes = Not unbalancedQuotes
                            End If
                        End If
                    ElseIf (argument.Substring(0, 1) = "/" Or _
                        argument.Substring(0, 1) = "-") And _
                        argument.Length >= 2 _
                    Then
                        setSwitch(argument.Substring(1))
                    Else
                        mArgs.Add(argument)
                    End If
                End If
            Else
                If ContainsUnbalancedQuotes(argument) Then
                    unbalancedQuotes = Not unbalancedQuotes
                End If
                If argument.Substring(argument.Length - 1) = """" And Not unbalancedQuotes Then
                    inQuotedArg = False
                    quotedArg = quotedArg & mSep & argument.Substring(0, argument.Length - 1)
                    If quotedArg.Substring(0, 1) = "/" Or _
                        quotedArg.Substring(0, 1) = "-" _
                    Then
                        setSwitch(quotedArg.Substring(1))
                    Else
                        mArgs.Add(quotedArg)
                    End If
                Else
                    If argument.Substring(argument.Length - 1) = """" Then
                        unbalancedQuotes = Not unbalancedQuotes
                    End If
                    quotedArg = quotedArg & mSep & argument
                End If
            End If
        Next

    End Sub

    Private Sub setSwitch(ByVal val As String)
        Dim i = val.IndexOf(":")
        Dim switchEntry = New SwitchEntry

        If i >= 0 Then
            switchEntry.Switch = val.Substring(0, i).ToUpper
            switchEntry.Value = val.Substring(i + 1, val.Length - (i + 1))
        Else
            switchEntry.Switch = val.ToUpper
        End If

        mSwitches.Add(switchEntry)
    End Sub

    Private Function ReplaceFirst(ByVal text As String, ByVal search As String, ByVal replace As String) As String
        Dim pos = text.IndexOf(search)
        If pos >= 0 Then Return text.Substring(0, pos) + replace + text.Substring(pos + search.Length)
        Return text
    End Function

End Class
