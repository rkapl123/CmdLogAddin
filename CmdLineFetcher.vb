Module CmdLineFetcher

    Sub getArgumentsAndStartMakro(Optional argStart As String = "/e", Optional argSep As String = "/")
        Dim CmdCaller As String, CallingWB As String
        Dim Args As Object, ArgPos As Long
        Dim CmdLineArgs() As String

        Try
            CmdLineArgs = Environment.GetCommandLineArgs()
            Dim ExcelMakroArg As String = FlaggedArg(argStart, CmdLineArgs, True)
            ' actual cmdline args
            ' Excel-passed cmd args, following "/e", separated by "/"
            Args = Split(ExcelMakroArg, argSep)
            ArgPos = FlagPresent(argStart & argSep & ExcelMakroArg, CmdLineArgs, True)
            ' /e/START|STARTINT/invokedMacro/arg1        /arg2   /arg3  ....
            ' /e/STARTEXT      /containedWB /invokedMacro/arg1   /arg2  ....
            ' / /Args(0)       /Args(1)     /Args(2)     /Args(3)/Args(4) ..
            If UBound(Args) >= 1 Then
                If UCase$(Args(0)) = "START" Or UCase$(Args(0)) = "STARTEXT" Or UCase$(Args(0)) = "STARTINT" Then
                    CallingWB = vbNullString
                    ' CmdCaller is calling workbook
                    CmdCaller = CmdLineArgs(ArgPos - 1)

                    If UCase$(Args(0)) = "START" Then ' called sub within calling workbook:
                        CallingWB = "'" & CmdCaller & "'!"
                    ElseIf UCase$(Args(0)) = "STARTEXT" Then ' called sub outside of calling workbook:
                        ' If we have a full path in args(1) then use it, however caller must provide also correct excel calling convention: '<FullPath>\File.xlsm'!Macro
                        If InStr(1, Args(1), "\") = 0 Then
                            ' no full path: take from CmdCaller (invoker of getArguments in Workbook_Open)
                            Args(1) = Replace(Args(1), "!", "'!")
                            ' If we have a full path in CmdCaller then use it, given workbook/addin in args(1) assumed to be in same directory
                            If InStr(1, CmdCaller, "\") Then
                                CallingWB = "'" & Trim$(Mid$(CmdCaller, 1, InStrRev(CmdCaller, "\")))
                                ' no path, rely on CurrentDirectory
                            Else
                                CallingWB = "'" & FileIO.FileSystem.CurrentDirectory() & "\"
                            End If
                        End If
                    Else
                        ' in case of STARTINT, either rely on called proc already having been loaded (Workbook/Addin in XLSTART) or invokedMacro is given with full path ('<FullPath>\File.xlsm'!Macro ; CallingWB is empty !)
                    End If
                    Select Case UBound(Args)
                        Case 1 : aLogger.LogToEventViewer("Calling: " & CallingWB & Args(1), EventLogEntryType.Information)
                        Case 2 : aLogger.LogToEventViewer("Calling: " & CallingWB & Args(1) & "," & Args(2), EventLogEntryType.Information)
                        Case 3 : aLogger.LogToEventViewer("Calling: " & CallingWB & Args(1) & "," & Args(2) & "," & Args(3), EventLogEntryType.Information)
                        Case 4 : aLogger.LogToEventViewer("Calling: " & CallingWB & Args(1) & "," & Args(2) & "," & Args(3) & "," & Args(4), EventLogEntryType.Information)
                    End Select
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
                End If
            End If
            StartMakroDone = True
        Catch ex As Exception
            aLogger.LogToEventViewer("Error: " & Err.Description & " in Fetcher.getArguments, " & Erl() & ",CallingWB: " & CallingWB & ",CmdCaller: " & CmdCaller,, True)
        End Try
    End Sub

    Public Function FlaggedArg(ByVal Flag As String, Arguments() As String, CaseSensitive As Boolean) As String
        ' This function will scan the argument list, looking for
        ' one that starts with the passed flag. If it's found, and
        ' the passed flag is the entire argument, the following
        ' argument is returned. If the passed flag isn't the entire
        ' argument, the portion following the flag is returned.
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

    Public Function FlagPresent(ByVal Flag As String, Arguments() As String, CaseSensitive As Boolean) As Long
        ' This function simply scans the argument list,
        ' looking for the passed flag, returns result.
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
