Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Diagnostics ' needed for EventLogEntryType !!


''' <summary>All procedures for fetching from the command line of excel and starting a given macro</summary>
Public Module CmdLineFetcher
    Public StartMakroDone As Boolean = False ' used in Application_WorkbookOpen to suppress further invocations when opening workbooks
    Public ArgsProhibited As Boolean = False ' prohibit Argument fetching during opening workbooks when App.Run Macro is loaded
    Public Args As Object
    Public CmdLineArgs() As String
    Public AppVisible As Boolean
    Public quittingApp As Boolean = False

    ''' <summary>necessary to run in main app thread as excel can't quit otherwise</summary>
    <ExcelCommand(Name:="QuitApp")>
    Public Sub QuitApp()
        ExcelDnaUtil.Application.DisplayAlerts = False
        ExcelDnaUtil.Application.Quit()
    End Sub

    ''' <summary>Logs internal sErrMsg of eEventType to Application EventLog, source .NET Runtime</summary>
    ''' <param name="sErrMsg"></param>
    ''' <param name="eEventType"></param>
    Public Sub internalLogToEventViewer(sErrMsg As String, Optional eEventType As EventLogEntryType = EventLogEntryType.Error)
        Dim eventLog As EventLog = New EventLog("Application")
        ' .Net Runtime is always there if .Net is installed
        EventLog.WriteEntry(".NET Runtime", sErrMsg, eEventType, 1000)
    End Sub

    ''' <summary>encapsulates setting fetching (currently registry)</summary>
    ''' <param name="Key">key of setting</param>
    ''' <param name="defaultValue">default value to be used if no setting given</param>
    ''' <returns>setting value</returns>
    Public Function fetchSetting(Key As String, defaultValue As Object) As Object
        fetchSetting = GetSetting("LogAddin", "Settings", Key, defaultValue)
    End Function

    ''' <summary>get excel arguments from command line of excel and start the macro given after Start or StartExt</summary>
    ''' <param name="argStart">argument starting portion to scan for ("/e" is the most harmless choice for excel)</param>
    ''' <param name="argSep">separator used for further separation of arguments being passed to macro</param>
    Sub getArgumentsAndStartMakro(Optional argStart As String = "/e", Optional argSep As String = "/", Optional calledByGetter As Boolean = False, Optional debugInfo As Boolean = False)
        Dim CmdCaller As String = "", CallingWB As String = ""

        Try
            ' get Array of CmdLine Arguments
            CmdLineArgs = Environment.GetCommandLineArgs()
            If CmdLineArgs.Length > 0 And debugInfo Then internalLogToEventViewer("CmdLineArgs:" & Join(CmdLineArgs, " "), EventLogEntryType.Information)
            ' get CmdLine Argument starting with argStart ("/e")
            Dim ExcelMakroArg As String = FlaggedArg(argStart, CmdLineArgs, True)
            If ExcelMakroArg <> "" And debugInfo Then internalLogToEventViewer("ExcelMakroArg:" & ExcelMakroArg, EventLogEntryType.Information)
            ' get actual passed arguments, following "/e", separated by "/"
            Args = Split(ExcelMakroArg, argSep)
            ' /e/START    /invokedMacro/arg1        /arg2   /arg3  ....
            ' /e/STARTEXT /containedWB /invokedMacro/arg1   /arg2  ....
            ' / /Args(0)  /Args(1)     /Args(2)     /Args(3)/Args(4) ..
            AppVisible = True
            If UBound(Args) >= 1 And Not calledByGetter Then
                Dim startSwitch As String = UCase$(Args(0))
                If Left(startSwitch, 5) = "START" Then
                    ' we like to be unobtrusive when starting from the commandline...
                    ExcelDnaUtil.Application.WindowState = Excel.XlWindowState.xlMinimized
                    ' in case we want to really be unobtrusive, specify hidden after your start
                    If Right(UCase$(startSwitch), 6) = "HIDDEN" Then
                        ExcelDnaUtil.Application.Visible = False
                        AppVisible = False
                        startSwitch = Replace(startSwitch, "HIDDEN", "")
                    End If
                    ' CmdCaller is (usually) the calling workbook (second Cmdline argument, first is Excel itself)
                    CmdCaller = CmdLineArgs(1)
                    ' if second/third/fourth cmdline argument is a switch passed to excel (like /r for readonly) then calling workbook is third/fourth/fifth cmdline argument
                    If Left(CmdCaller, 1) = "/" And CmdLineArgs.Length > 2 Then
                        CmdCaller = CmdLineArgs(2)
                        If Left(CmdCaller, 1) = "/" And CmdLineArgs.Length > 3 Then
                            CmdCaller = CmdLineArgs(3)
                            If Left(CmdCaller, 1) = "/" And CmdLineArgs.Length > 4 Then
                                CmdCaller = CmdLineArgs(4)
                            End If
                        End If
                    End If
                    If debugInfo Then internalLogToEventViewer("CmdCaller:" & CmdCaller & ",ExcelMakroArg:" & ExcelMakroArg & ",startSwitch:" & startSwitch & ",Args(1):" & Args(1), EventLogEntryType.Information)

                    If startSwitch = "START" Then ' called sub within calling workbook or loaded addins: Start
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
                    If debugInfo Then
                        Dim debugArgs As String = ""
                        For i As Integer = 1 To UBound(Args)
                            debugArgs += Args(i) + ","
                        Next
                        internalLogToEventViewer("Calling: " & CallingWB & debugArgs, EventLogEntryType.Information)
                    End If
                    ' prohibit Argument fetching during opening workbooks when App.Run Macro is loaded
                    ArgsProhibited = True
                    Try
                        With ExcelDnaUtil.Application
                            Select Case UBound(Args)
                                Case 1 : .Run(CallingWB & Args(1))
                                Case 2 : .Run(CallingWB & Args(1), Args(2))
                                Case 3 : .Run(CallingWB & Args(1), Args(2), Args(3))
                                Case 4 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4))
                                Case 5 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5))
                                Case 6 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6))
                                Case 7 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7))
                                Case 8 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8))
                                Case 9 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9))
                                Case 10 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10))
                                Case 11 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11))
                                Case 12 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12))
                                Case 13 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13))
                                Case 14 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14))
                                Case 15 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15))
                                Case 16 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16))
                                Case 17 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17))
                                Case 18 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18))
                                Case 19 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19))
                                Case 20 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20))
                                Case 21 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21))
                                Case 22 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22))
                                Case 23 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23))
                                Case 24 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24))
                                Case 25 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25))
                                Case 26 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26))
                                Case 27 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27))
                                Case 28 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27), Args(28))
                                Case 29 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27), Args(28), Args(29))
                                Case 30 : .Run(CallingWB & Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27), Args(28), Args(29), Args(30))
                            End Select
                        End With
                    Catch ex As Exception
                        internalLogToEventViewer("Error: " & ex.Message & " in Application.Run, CallingWB: " & CallingWB & ",CmdCaller: " & CmdCaller, EventLogEntryType.Error)
                    End Try
                    ArgsProhibited = False
                    StartMakroDone = True
                End If
            End If
        Catch ex As Exception
            ArgsProhibited = False
            internalLogToEventViewer("Error: " & ex.Message & " in Fetcher.getArguments, CallingWB: " & CallingWB & ",CmdCaller: " & CmdCaller, EventLogEntryType.Error)
        End Try
        ' if we were not quit by now, make excel visible again to say we are here...
        If Not quittingApp And Not ExcelDnaUtil.Application.Visible Then ExcelDnaUtil.Application.Visible = True
    End Sub

    <ExcelCommand(Name:="getCmdLineArgs")>
    Public Function getCmdLineArgs(Optional debugInfo As Boolean = False) As Object
        getArgumentsAndStartMakro(calledByGetter:=True, debugInfo:=debugInfo)
        Return CmdLineArgs
    End Function

    <ExcelCommand(Name:="getExcelPassedArgs")>
    Public Function getExcelPassedArgs(Optional debugInfo As Boolean = False) As Object
        getArgumentsAndStartMakro(calledByGetter:=True, debugInfo:=debugInfo)
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

End Module
