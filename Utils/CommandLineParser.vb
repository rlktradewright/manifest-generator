#Region "License"

' The MIT License (MIT)
'
' Copyright (c) 2017-2018 Richard L King (TradeWright Software Systems)
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

#End Region

''' <summary>
''' Provides facilities for an application to determine the number and values of
''' arguments and switches in a string (normally the arguments part of the command
''' used to start the application).
''' </summary>
''' <remarks>
''' The format of the argument string passed to the <c>CreateCommandLineParser</c> method
''' is as follows:
'''
''' <pre>
'''   [&lt;argument&gt; | &lt;switch&gt;] [&lt;sep&gt; (&lt;argument&gt; | &lt;switch&gt;)]...
''' </pre>
'''
''' ie, there is a sequence of arguments or switches, separated by separator characters. The
''' separator character is specified in the constructor.
'''
''' Arguments that contain the separator character must be enclosed in double quotes. Double quotes
''' appearing within an argument must be repeated.
'''
''' Switches have the following format:
'''
''' <pre>
'''   ( "/" | "-")&lt;identifier&gt; [":"&lt;switchValue&gt;]
''' </pre>
'''
''' ie the switch starts with a forward slash or a hyphen followed by an identifier, and
''' optionally followed by a colon and the switch value. Switch identifiers are not
''' case-sensitive. Switch values that contain the separator character must be enclosed in
''' double quotes. Double quotes appearing within a switch value must be repeated.
'''
''' Examples (these examples use a space as the separator character):
''' <pre>
'''   anArgument -sw1 anotherArg -sw2:42
''' </pre>
''' <pre>
'''   "C:\Program Files\MyApp\myapp.ini" -out:C:\MyLogs\myapp.Log
''' </pre>
''' </remarks>
Public NotInheritable Class CommandLineParser

    ''' <summary>
    ''' Contains details of a command line switch.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SwitchEntry
        ''' <summary>
        ''' The switch identifier.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Name As String

        ''' <summary>
        ''' The switch value.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Value As String = String.Empty

        Friend Sub New(value As String)
            Dim i = value.IndexOf(":")
            If i >= 0 Then
                Name = value.Substring(0, i)
                Me.Value = value.Substring(i + 1)
            Else
                Name = value
            End If
        End Sub
    End Class

    '@================================================================================
    ' Member variables
    '@================================================================================

    Private mArgs As List(Of String) = New List(Of String)
    Private mSwitches As List(Of SwitchEntry) = New List(Of SwitchEntry)
    Private mSep As String
    Private mInputString As String
    Private mCaseSensitive As Boolean

    Private Sub New()
    End Sub

    ''' <summary>
    '''  Initialises a new instance of the <see cref="CommandLineParser"></see> class.
    ''' </summary>
    ''' <param name="inputString">The command line arguments to be parsed. For a Visual Basic 6 program,
    '''  this value may be obtained using the <c>Command</c> function.</param>
    ''' <param name="separator">A single character used as the separator between command line arguments.</param>
    ''' <remarks></remarks>
    Public Sub New(inputString As String, Optional separator As String = " ", Optional caseSensitive As Boolean = False)
        mInputString = inputString.Trim
        mSep = separator
        mCaseSensitive = caseSensitive
        getArgs()
    End Sub

    ''' <summary>
    ''' Gets the nth argument, where n is the value of the <paramref>index</paramref> parameter.
    ''' </summary>
    ''' <param name="index"> The number of the argument to be returned. The first argument is number 0.</param>
    ''' <value></value>
    ''' <returns>A String value containing the nth argument, where n is the value of the <paramref>index</paramref> parameter.</returns>
    ''' <remarks>If the requested argument has not been supplied, an empty string is returned.</remarks>
    Public ReadOnly Property Arg(index As Integer) As String
        Get
            If index < 0 Then Throw New ArgumentException("Argument must be >= 0", "i")
            If index < mArgs.Count Then
                Return mArgs(index)
            Else
                Return String.Empty
            End If
        End Get
    End Property

    ''' <summary>
    ''' Gets a List(Of string) containing the arguments.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A List(Of string) containing the arguments.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Args() As List(Of String)
        Get
            Return mArgs
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
    ''' <param name="name">The identifier of the switch whose inclusion is to be indicated.</param>
    ''' <value></value>
    ''' <returns>If the specified switch was included, <c>True</c> is
    ''' returned. Otherwise <c>False</c> is returned.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsSwitchSet(name As String) As Boolean
        Get
            Dim i = findSwitch(name)
            Return i <> -1
        End Get
    End Property

    ''' <summary>
    ''' Gets a <c>List(Of SwitchEntry)</c> containing the
    ''' switch identifiers and values.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A <c>List(Of SwitchEntry)</c>s containing the
    ''' switch identifiers and values.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Switches() As List(Of SwitchEntry)
        Get
            Return mSwitches
        End Get
    End Property

    ''' <summary>
    ''' Gets the zero-based index (within the set of switches) of the specified switch.
    ''' </summary>
    ''' <param name="name">The identifier of the switch whose index is to be returned.</param>
    ''' <value></value>
    ''' <returns>The index of the specified switch.</returns>
    ''' <remarks>If the requested switch has not been supplied, -1 is returned.</remarks>
    Public ReadOnly Property SwitchIndex(name As String) As Integer
        Get
            Return findSwitch(name)
        End Get
    End Property

    ''' <summary>
    ''' Gets the value of the specified switch.
    ''' </summary>
    ''' <param name="name">The identifier of the switch whose value is to be returned.</param>
    ''' <value></value>
    ''' <returns>A String containing the value for the specified switch.</returns>
    ''' <remarks>If the requested switch has not been supplied, or no value
    ''' was supplied for the switch, an empty string is returned.</remarks>
    Public ReadOnly Property SwitchValue(name As String) As String
        Get
            Dim i = findSwitch(name)
            If i = -1 Then Return String.Empty
            Return mSwitches(i).Value
        End Get
    End Property


    Private Function ContainsUnbalancedQuotes(inString As String) As Boolean
        Dim pos = inString.LastIndexOf("""")
        Dim unBalanced = False
        Do While pos <> -1
            unBalanced = Not unBalanced
            If pos = 0 Then Exit Do
            pos = inString.LastIndexOf("""", pos - 1)
        Loop
        Return unBalanced
    End Function

    Private Function findSwitch(name As String) As Integer
        For i = 0 To mSwitches.Count - 1
            If If(mCaseSensitive,
                    name.Equals(mSwitches(i).Name),
                    name.Equals(mSwitches(i).Name, StringComparison.CurrentCultureIgnoreCase)) Then Return i
        Next
        Return -1
    End Function

    Private Sub getArgs()
        If mInputString = "" Then Exit Sub

        Dim partialArg As String = String.Empty
        For Each argument In mInputString.Split({mSep}, StringSplitOptions.None)
            If String.IsNullOrEmpty(partialArg) And argument = String.Empty And mSep = " " Then
                ' discard spaces when the separator is a space and we don't have unbalanced quotes
            Else
                If Not String.IsNullOrEmpty(partialArg) Then partialArg &= mSep
                partialArg &= argument
                If Not ContainsUnbalancedQuotes(partialArg) Then
                    setSwitchOrArg(partialArg.Trim())
                    partialArg = String.Empty
                End If
            End If
        Next

        If Not String.IsNullOrEmpty(partialArg) Then
            setSwitchOrArg(partialArg)
        End If
    End Sub

    Private Function isAlphaChar(value As Char) As Boolean
        If value >= "A" AndAlso value <= "Z" Then Return True
        If value >= "a" AndAlso value <= "z" Then Return True
        Return False
    End Function

    Private Sub setSwitchOrArg(value As String)
        If value = "--" OrElse
            ((value.StartsWith("/") OrElse value.StartsWith("-")) AndAlso
                value.Length > 1 AndAlso
                isAlphaChar(CChar(value.Substring(1, 1)))) Then
            setSwitch(value.Substring(1))
        Else
            mArgs.Add(value)
        End If
    End Sub

    Private Sub setSwitch(val As String)
        Dim switchEntry = New SwitchEntry(val)
        mSwitches.Add(switchEntry)
    End Sub

End Class

