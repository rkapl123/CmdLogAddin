Imports System.Runtime.InteropServices
Imports ExcelDna.Integration


''' <summary>All procedures for fetching from the command line of excel and starting a given macro</summary>
Public Module CmdLineFetcher
    Public StartMakroDone As Boolean = False ' used in Application_WorkbookOpen to suppress further invocations when opening workbooks
    Public ArgsProhibited As Boolean = False ' prohibit Argument fetching during opening workbooks when App.Run Macro is loaded
    Public Args As Object
    Public CmdLineArgs() As String

    ''' <summary>get excel arguments from command line of excel and start the macro given after Start or StartExt</summary>
    ''' <param name="argStart">argument starting portion to scan for ("/e" is the most harmless choice for excel)</param>
    ''' <param name="argSep">separator used for further separation of arguments being passed to macro</param>
    Sub getArgumentsAndStartMakro(Optional argStart As String = "/e", Optional argSep As String = "/")
        Dim CmdCaller As String = "", CallingWB As String = ""
        ' get the host app object afresh to avoid dangling references to excel application, prohibiting quit
        Dim theHostApp As Object = ExcelDnaUtil.Application

        Try
            ' get Array of CmdLine Arguments
            CmdLineArgs = Environment.GetCommandLineArgs()
            If CmdLineArgs.Count > 0 Then LogToEventViewer("CmdLineArgs:" & Join(CmdLineArgs, " "), EventLogEntryType.Information)
            ' get CmdLine Argument starting with argStart ("/e")
            Dim ExcelMakroArg As String = FlaggedArg(argStart, CmdLineArgs, True)
            If ExcelMakroArg <> "" Then LogToEventViewer("ExcelMakroArg:" & ExcelMakroArg, EventLogEntryType.Information)
            ' get actual passed arguments, following "/e", separated by "/"
            Args = Split(ExcelMakroArg, argSep)
            ' /e/START    /invokedMacro/arg1        /arg2   /arg3  ....
            ' /e/STARTEXT /containedWB /invokedMacro/arg1   /arg2  ....
            ' / /Args(0)  /Args(1)     /Args(2)     /Args(3)/Args(4) ..
            If UBound(Args) >= 1 Then
                If UCase$(Args(0)) = "START" Or UCase$(Args(0)) = "STARTEXT" Then
                    ' CmdCaller is (usually) the calling workbook (second Cmdline argument, first is Excel itself)
                    CmdCaller = CmdLineArgs(1)
                    ' if second cmdline argument is a switch passed to excel (like /r for readonly) then calling workbook is third cmdline argument
                    If Left(CmdCaller, 1) = "/" Then CmdCaller = CmdLineArgs(2)
                    LogToEventViewer("CmdCaller:" & CmdCaller & ",ExcelMakroArg:" & ExcelMakroArg & ",Args(0):" & Args(0) & ",Args(1):" & Args(1), EventLogEntryType.Information)

                    If UCase$(Args(0)) = "START" Then ' called sub within calling workbook or loaded addins: Start
                        If InStr(1, Args(1), "!") = 0 Then CallingWB = "'" & CmdCaller & "'!"
                    Else ' called sub outside of calling workbook or loaded addins/workbooks: StartExt
                        ' If we have a full path in Args(1) then use it, however caller must provide also correct excel calling convention: '<FullPath>\File.xlsm'!Macro
                        If InStr(1, Args(1), "\") = 0 Then
                            ' no full path: take from CmdCaller
                            Args(1) = Replace(Args(1), "!", "'!")
                            ' If we have a full path in CmdCaller then use it, given workbook/addin in Args(1) assumed to be in same directory
                            If InStr(1, CmdCaller, "\") > 0 Then
                                CallingWB = "'" & Trim$(Mid$(CmdCaller, 1, InStrRev(CmdCaller, "\")))
                                ' no path, rely on CurrentDirectory
                            Else
                                CallingWB = "'" & FileIO.FileSystem.CurrentDirectory() & "\"
                            End If
                        End If
                    End If
                    Select Case UBound(Args)
                        Case 1 : LogToEventViewer("Calling: " & CallingWB & Args(1), EventLogEntryType.Information)
                        Case 2 : LogToEventViewer("Calling: " & CallingWB & Args(1) & "," & Args(2), EventLogEntryType.Information)
                        Case 3 : LogToEventViewer("Calling: " & CallingWB & Args(1) & "," & Args(2) & "," & Args(3), EventLogEntryType.Information)
                        Case 4 : LogToEventViewer("Calling: " & CallingWB & Args(1) & "," & Args(2) & "," & Args(3) & "," & Args(4), EventLogEntryType.Information)
                    End Select
                    ' prohibit Argument fetching during opening workbooks when App.Run Macro is loaded
                    ArgsProhibited = True
                    Select Case UBound(Args)
                        Case 1 : theHostApp.Run(CallingWB & Args(1))
                        Case 2 : theHostApp.Run(CallingWB & Args(1), Args(2))
                        Case 3 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3))
                        Case 4 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4))
                        Case 5 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5))
                        Case 6 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6))
                        Case 7 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7))
                        Case 8 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8))
                        Case 9 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9))
                        Case 10 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10))
                        Case 11 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11))
                        Case 12 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12))
                        Case 13 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13))
                        Case 14 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14))
                        Case 15 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15))
                        Case 16 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16))
                        Case 17 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17))
                        Case 18 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18))
                        Case 19 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19))
                        Case 20 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20))
                        Case 21 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21))
                        Case 22 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22))
                        Case 23 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23))
                        Case 24 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24))
                        Case 25 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25))
                        Case 26 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26))
                        Case 27 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27))
                        Case 28 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27), Args(28))
                        Case 29 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27), Args(28), Args(29))
                        Case 30 : theHostApp.Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27), Args(28), Args(29), Args(30))
                    End Select
                    ArgsProhibited = False
                    StartMakroDone = True
                End If
            End If
        Catch ex As Exception
            ArgsProhibited = False
            LogToEventViewer("Error: " & ex.Message & " in Fetcher.getArguments, CallingWB: " & CallingWB & ",CmdCaller: " & CmdCaller, EventLogEntryType.Error)
        End Try
    End Sub

    <ExcelCommand(Name:="getCmdLineArgs")>
    Public Function getCmdLineArgs() As Object
        getArgumentsAndStartMakro()
        Return CmdLineArgs
    End Function

    <ExcelCommand(Name:="getExcelPassedArgs")>
    Public Function getExcelPassedArgs() As Object
        getArgumentsAndStartMakro()
        Return Args
    End Function

    ''' <summary>scans the argument list, looking for one that starts with the passed flag. If it's found, and the passed flag is the entire argument, the following
    ''' argument is returned. If the passed flag isn't the entire argument, the portion following the flag is returned.</summary>
    ''' <param name="Flag"></param>
    ''' <param name="Arguments"></param>
    ''' <param name="CaseSensitive"></param>
    ''' <returns></returns>
    Public Function FlaggedArg(ByVal Flag As String, Arguments() As String, CaseSensitive As Boolean) As String
        Dim i As Long
        Dim sRet As String = ""
        Dim CompareFlag As CompareMethod

        ' Convert flag to lowercase if case isn't important.
        If CaseSensitive Then
            CompareFlag = CompareMethod.Binary
        Else
            CompareFlag = CompareMethod.Text
        End If

        ' Scan arglist, looking for passed flag.
        For i = 1 To UBound(Arguments)
            If InStr(1, Arguments(i), Flag, CompareFlag) = 1 Then
                ' Base return on whether argument follows directly
                ' after flag, or with slash delimiter.
                If Len(Arguments(i)) > Len(Flag) Then
                    sRet = Mid$(Arguments(i), Len(Flag) + 1)
                    If Len(sRet) > 1 Then
                        If InStr("/", Left$(sRet, 1)) Then
                            ' Trim first character.
                            sRet = Mid$(sRet, 2)
                        End If
                    End If
                Else
                    If i < UBound(Arguments) Then
                        sRet = Arguments(i + 1)
                    End If
                End If
                ' All done here.
                Exit For
            End If
        Next i
        ' Return results
        FlaggedArg = sRet
    End Function

    ''' <summary>scans the argument list, looking for the passed flag, returns resulting position.</summary>
    ''' <param name="Flag"></param>
    ''' <param name="Arguments"></param>
    ''' <param name="CaseSensitive"></param>
    ''' <returns>position of argument</returns>
    Public Function FlagPresent(ByVal Flag As String, Arguments() As String, CaseSensitive As Boolean) As Long
        Dim i As Long
        Dim CompareFlag As CompareMethod

        ' Convert flag to lowercase if case isn't important.
        If CaseSensitive Then
            CompareFlag = CompareMethod.Binary
        Else
            CompareFlag = CompareMethod.Text
        End If

        FlagPresent = -1
        ' Scan arglist, looking for passed flag.
        For i = 1 To UBound(Arguments)
            If StrComp(Arguments(i), Flag, CompareFlag) = 0 Then
                ' Found it, return matching index.
                FlagPresent = i
                Exit For
            End If
        Next i
    End Function

    ''' <summary>Logs sErrMsg of eEventType in eCategory to EventLog</summary>
    ''' <param name="sErrMsg"></param>
    ''' <param name="eEventType"></param>
    Private Sub LogToEventViewer(sErrMsg As String, Optional eEventType As EventLogEntryType = EventLogEntryType.Error)
        Dim eventLog As EventLog = New EventLog("Application")
        ' .Net Runtime is always there if .Net is installed
        EventLog.WriteEntry(".NET Runtime", sErrMsg, eEventType, 1000)
    End Sub

End Module
