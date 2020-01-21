' create the logger ...
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("Logger.vbs", 1).ReadAll
If InStr(1, Wscript.ScriptFullName, "Test") > 0 Then env = "Test"
env = ""

' ... and init logging
'
' theCallingObject .. the calling object (excel workbook, word document, access project, etc..), 
'          must have a "Name" and "Path" property and a "Quit" method (to allow LogFatal to end it)...
' theLogLevel ..  (ERROR 1,  WARN 2, INFO 4, DEBUG 8), default = 4
' theLogFilePath .. where to write the logfile (LogFilePath), defaults to calling script's path
' theEnv .. environment, empty if production, used to append env to LogFilePath for test/other environments
' theCaller .. if caller is not the callingObject (commonCaller) then this can be used to
'               identify the active caller (in case of a script library handling other scripts ..).
'               Can include the full path to the calling script..,
'               the Caller's name will be extracted by using last "\" as separator
' theMailRecipients .. comma separated list of the error mail recipients
' theSubject .. the error mail's subject
' writeToEventLog .. should messages be written to the windows event log (true) or to a file (false)
' theSender .. the Sender of the sent error mails
' theMailIntro .. the intro for the error mail's body
' theMailGreetings .. the greetings for the error mail's body, body looks as follows:
' <MailIntro> (executed in: <commonCaller>, current caller: <Caller>):
' <logLine>
' <logPathMsg>
' <MailGreetings>
' overrideCommonCaller .. whether to override CallingObjectName (filename to log to) with theCaller
' doMirrorToStdOut .. whether to mirror log messages to the standard output (requires cscript execution of scripts) or a separate debug window (VB/VBA/WSCRIPT !)
theLogger.setProperties Wscript, 8, "C:\dev\CmdLogAddin", env, ,"your.address@gmail.at", , , , , , ,True
' now write some log messages
theLogger.LogError "testLog.vbs: logging error"
theLogger.LogWarn "testLog.vbs: logging warning"
theLogger.LogInfo "testLog.vbs: logging info"
theLogger.LogDebug "testLog.vbs: logging debug"

'' change caller settings, log file is in directory of calling script
theLogger.setProperties Wscript, , , ,"theTestLog.ext", , , , , , , ,True
theLogger.LogError "theTestLog.ext: logging error"
theLogger.LogWarn "theTestLog.ext: logging warning"
theLogger.LogInfo "theTestLog.ext: logging info"
theLogger.LogDebug "theTestLog.ext: logging debug"

' write to windows event log (source is WSH)
theLogger.setProperties Wscript, , , , , , ,True, , , , ,True
theLogger.LogError "windows event log: logging error"
theLogger.LogWarn "windows event log: logging warning"
theLogger.LogInfo "windows event log: logging info"
theLogger.LogDebug "windows event log: logging debug"

' capture standard output and standard error of cmd... log file is in directory of calling script
theLogger.setProperties Wscript, , , , , , ,False, , , , ,True
set WshShell = WScript.CreateObject("WScript.Shell")
theLogger.LogStream WshShell.Exec ("cmd.exe /C echo this is the standard output of an executed command (output via LogStream) !!")
theLogger.LogStream WshShell.Exec ("cmd.exe /C echo this is the standard error of an executed command (output via LogStream) !! 1>&2")
theLogger.LogFatal "logging fatal (ends execution)"
theLogger.LogDebug "shouldn't come here !"
