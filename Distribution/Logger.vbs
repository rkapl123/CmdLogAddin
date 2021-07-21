'' add logger with:
' loggerHome = "place\where\Logger.vbs\is\located"
' ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(loggerHome & "Logger.vbs", 1).ReadAll

'' after adding logger, set environment based on folder name and set the properties of the logger:
' If InStr(1, Wscript.ScriptFullName, "Test") > 0 Then theEnv = "Test"
'' Here theMailRecipients is the current user. Can also be some other hardcoded mail address...
' theMailRecipients = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERNAME%") & "@yourdomain.com"
' theLogger.setProperties theCallingObject = Wscript, theLogLevel = 4, theLogFilePath, theEnv, theCaller, theMailRecipients, theSubject, writeToEventLog, theSender, theMailIntro, theMailGreetings, overrideCommonCaller, doMirrorToStdOut

' in the code, add logging as follows:
'theLogger.LogError "error" ' also sends an errormail
'theLogger.LogWarn "warning"
'theLogger.LogInfo "info"
'theLogger.LogDebug "debug"
'theLogger.LogFatal "fatal error (ends execution)" ' also sends an errormail

Option Explicit
Const RootRegPath = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\LogAddin\Settings\"
Public Const LogErr = 1
Public Const LogWrn = 2
Public Const LogInf = 4
Public Const LogDbg = 8

' cscript based simple logger
' Copyright © 2020, MIT License, Roland Kapl
' https://rkapl123.github.io/CmdLogAddin

Public theLogger
set theLogger = New Logger

Class Logger
	' settings done at initialization
	Private cdoUserID                
	Private cdoPassword              
	Private cdoUseSSL                
	Private cdoAuthentRequired       'false
	Private cdoServerName            '= "SomeSMTPServer"   'Name or IP of Remote SMTP Server
	Private cdoServerPort            '= 25             'Server port (typically 25)
	Private cdoInternalErrMailRcpt   '
	Private defaultSubject           '"Batch Process Error"
	Private defaultSender            '"Administrator"
	Private defaultMailIntro         '"Folgender Fehler trat in batch process auf "
	Private defaultMailGreetings     '"liebe Grüße schickt der Fehleradmin..."
	Private logentry()               

	' properties set using setProperties
	Private callingObject    
	Private callingObjectPath
	Private LogLevel         
	Private LogFilePath      
	Private env              
	Private commonCaller     
	Private Caller           
	Private callerFullPath   
	Private mailRecipients   
	Private Subject          
	Private doEventLog       
	Private Sender           
	Private MailIntro        
	Private MailGreetings    
	Private mirrorToStdOut

	Private AlreadySent
	Private shell
	Private fso
	Private fh
	
	Private Sub Class_Initialize()
		Set shell = CreateObject("WScript.Shell")
		' initializes configuration variables from registry
		cdoUserID = fetchSetting("cdoUserID", "")
		cdoPassword = fetchSetting("cdoPassword", "")
		cdoUseSSL = CBool(fetchSetting("cdoUseSSL", False))
		cdoAuthentRequired = CBool(fetchSetting("cdoAuthentRequired", False))
		cdoServerName = fetchSetting("cdoServerName", "YOURSMTPSERVERNAME")    'Name or IP of Remote SMTP Server
		cdoServerPort = CInt(fetchSetting("cdoServerPort", 25))               'Server port (eg 25)
		cdoInternalErrMailRcpt = fetchSetting("cdoInternalErrMailRcpt", "internalErrAdmin1,internalErrAdmin2")
		defaultSubject = fetchSetting("defaultSubject", "Batch Process Error")
		defaultSender = fetchSetting("defaultSender", "Administrator")
		defaultMailIntro = fetchSetting("defaultMailIntro", "Folgender Fehler trat in batch process auf ")
		defaultMailGreetings = fetchSetting("defaultMailGreetings", "liebe Grüße schickt der Fehleradmin...")

		Dim i
		i = 0
		Do
			ReDim Preserve logentry(i)
			logentry(i) = fetchSetting("logentry" & i, vbNullString)
			i = i + 1
		Loop Until LenB(logentry(i - 1)) = 0
		mirrorToStdOut = False
	End Sub
	
	Private Sub Class_Terminate(  )
		Set shell = Nothing
		on error resume next
		fh.Close
		set fh = Nothing
		set fso = Nothing
	End Sub

	' fetchSetting
	'
	' encapsulates setting fetching (currently registry)
	Private Function fetchSetting(Key, defaultValue)
		Dim value
		On Error Resume Next
		value = shell.RegRead( RootRegPath & Key )

		If err.number <> 0 Then
			fetchSetting = defaultValue 
		Else
			fetchSetting = value
		End If
	End Function

	' setProperties
	'
	' sets properties for the Logger object, all parameters optional !
	' theCallingObject .. the calling script host (kept for compatibility reasons, defaults to current script host)
	' theLogLevel .. (ERROR 1,  WARN 2, INFO 4, DEBUG 8), default = 4
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
	' doMirrorToStdOut .. whether to mirror log messages to the standard output (requires cscript execution of scripts)
	Public Sub setProperties(theCallingObject, theLogLevel, theLogFilePath, theEnv, theCaller, theMailRecipients, theSubject, writeToEventLog, theSender, theMailIntro, theMailGreetings, overrideCommonCaller, doMirrorToStdOut)
		Dim theLogFile

		AlreadySent = False
		On Error Resume Next
		Set callingObject = theCallingObject
		commonCaller = callingObject.ScriptName
		callingObjectPath = Replace(callingObject.ScriptFullName, "\" & callingObject.ScriptName, "")
		If VarType(theCaller) <> vbError Then
			callerFullPath = theCaller
			Caller = theCaller
			If InStr(1, callerFullPath, "\") Then Caller = Mid(callerFullPath, InStrRev(callerFullPath, "\") + 1)
		Else
			callerFullPath = callingObject.ScriptFullName
		End If
		If VarType(theLogFilePath) <> vbError Then 
			LogFilePath = theLogFilePath
		Else
			LogFilePath = callingObjectPath
		End If
		' if no absolute path was given (drive mapping or unc), prepend callingObjectPath
		If Left(LogFilePath, 2) <> "\\" And InStr(1, LogFilePath, ":") = 0 Then LogFilePath = callingObjectPath & "\" & LogFilePath
		'LogToIntEventViewer "Caller: " & Caller & ",LogFilePath: " & LogFilePath
		If LogLevel = 0 then LogLevel = 4
		If VarType(theLogLevel)<> vbError Then LogLevel = theLogLevel
		If VarType(theEnv)<> vbError Then env = theEnv
		If VarType(theMailRecipients)<> vbError Then mailRecipients = theMailRecipients
		If VarType(theSubject)<> vbError Then Subject = theSubject
		If VarType(writeToEventLog)<> vbError Then doEventLog = writeToEventLog
		If VarType(theSender)<> vbError Then Sender = theSender
		If VarType(theMailIntro)<> vbError Then MailIntro = theMailIntro
		If VarType(theMailGreetings)<> vbError Then MailGreetings = theMailGreetings
		If VarType(overrideCommonCaller)<> vbError Then
			If overrideCommonCaller Then commonCaller = theCaller
		End If
		If VarType(doMirrorToStdOut)<> vbError Then mirrorToStdOut = doMirrorToStdOut
		on error resume next
		fh.Close
		set fh = nothing
		set fso = nothing
		Err.Clear
		set fso = CreateObject("Scripting.FileSystemObject")
		If Err <> 0 Then LogToIntEventViewer "Error: " & Err.Description & " after CreateObject(Scripting.FileSystemObject) in setProperties"
		theLogFile = IIf(Left(LogFilePath, 2) = "\\" Or Mid(LogFilePath, 2, 2) = ":\", "", callingObjectPath & "\") & LogFilePath & "\" & IIf(LenB(env) > 0, env & "\", vbNullString) & commonCaller & ".log"
		set fh = fso.OpenTextFile(theLogFile, 8, True)
		If Err <> 0 Then LogToIntEventViewer "Error: " & Err.Description & " after fso.OpenTextFile in setProperties"
	End Sub

	' LogStream, logs stdout/err from WshShell.Exec object obj to LogFile
	'
	' obj: WshShell.Exec object whose stdout should be logged as INFO and stderr as ERROR
	Public Sub LogStream(obj)
		Do While Not obj.StdOut.AtEndOfStream
			If LogLevel >= LogInf Then LogWrite obj.StdOut.ReadLine, LogInf
		Loop
		Do While Not obj.StdErr.AtEndOfStream
			LogWrite obj.StdErr.ReadLine, LogErr
		Loop
	End Sub

	' LogWrite
	' writes log information messages to defined Logfiles
	'
	' Args:
	' LogMessage .. Message to be logged
	' LogPrio .. (optional) priority level (ERROR 1,  WARN 2, INFO 4, DEBUG 8)
	Private Sub LogWrite(LogMessage, LogPrio)
		Dim FileNum, timestamp, MailFileLink, FileMessage

		timestamp = FormatDateTime(Now(), 0)
		Dim LogChoice
		Select Case LogPrio
			Case LogErr
				LogChoice = "ERROR"
			Case LogWrn
				LogChoice = "WARN"
			Case LogInf
				LogChoice = "INFO"
			Case LogDbg
				LogChoice = "DEBUG"
			Case Else
		End Select
		Dim i
		For i = 0 To UBound(logentry) - 1
			If LCase(logentry(i)) = "timestamp" Then FileMessage = FileMessage & timestamp & vbTab
			If LCase(logentry(i)) = "loglevel" Then FileMessage = FileMessage & LogChoice & vbTab
			If LCase(logentry(i)) = "caller" Then FileMessage = FileMessage & "file://" & Replace(callerFullPath, " ", "%20") & vbTab
			If Left(LCase(logentry(i)), 2) = "e:" Then FileMessage = FileMessage & shell.ExpandEnvironmentStrings("%" & Mid(logentry(i), 3) & "%") & vbTab
			If LCase(logentry(i)) = "logmessage" Then FileMessage = FileMessage & LogMessage & vbTab
		Next
		FileMessage = Left(FileMessage, Len(FileMessage) - 1)
		
		On Error Resume Next
		If mirrorToStdOut Then
			If InStr(1, UCase(callingObject.FullName), "CSCRIPT") > 0 Then
				callingObject.StdOut.WriteLine timestamp & vbTab & LogChoice & vbTab & LogMessage
				If Err<>0 then LogToIntEventViewer "Error: " & Err.Description & " after callingObject.StdOut.WriteLine in Logger.LogWrite"
			End If
		End If
		If doEventLog Then
			shell.LogEvent LogPrio, LogMessage
			MailFileLink = "Log in Event Viewer on machine \\" & shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		Else
			fh.WriteLine FileMessage
			MailFileLink = "Logfile in file://" & Replace(LogFilePath, " ", "%20") & "\" & IIf(LenB(env) > 0, env & "\", vbNullString) & commonCaller & ".log"
		End If
		If LenB(mailRecipients) > 0 And LogPrio = LogErr And Not AlreadySent Then sendMail FileMessage, MailFileLink
	End Sub

	' LogError, LogWarn, LogInfo, LogDebug
	'
	' writes a log message LogMessage having appropriate priority
	' (As shown in funtion name, usually INFO, ERROR.., etc) depending on log level set
	' in global const Loglevel (1 = ERROR, 2 = WARN, 4 = INFO, 8 = DEBUG)
	Public Sub LogError(LogMessage)
		LogWrite LogMessage, LogErr
	End Sub

	Public Sub LogWarn(LogMessage)
		If LogLevel >= LogWrn Then LogWrite LogMessage, LogWrn
	End Sub

	Public Sub LogInfo(LogMessage)
		If LogLevel >= LogInf Then LogWrite LogMessage, LogInf
	End Sub

	Public Sub LogDebug(LogMessage)
		 If LogLevel >= LogDbg Then LogWrite LogMessage, LogDbg
	End Sub

	' LogFatal
	'
	' writes a log message LogMessage with Error priority and quits
	Public Sub LogFatal(LogMessage)
		LogError LogMessage
		On Error Resume Next
		callingObject.Quit
	End Sub

	' sendMail
	'
	' sends an error mail containing the logged line (logLine) and a hyperlink to the logfile (logPathMsg)
	Private Sub sendMail(logLine, logPathMsg)
		dim mailProcExit, mailCmd, curPath, fsm, mb
		' write mail message to file, being picked up by mailCmd later in curPath
		curPath = Replace(Wscript.ScriptFullName, "\" & Wscript.ScriptName, "")
		set fsm = CreateObject("Scripting.FileSystemObject")
		set mb = fsm.OpenTextFile(curPath & "\mailbody.txt", 2, True)
		mb.WriteLine IIf(LenB(MailIntro) = 0, defaultMailIntro, MailIntro) & "(executed in: " & commonCaller & ", current caller: " & Caller & "):" & vbLf & vbLf _
			& logLine & vbLf & vbLf _
			& logPathMsg & vbLf & vbLf _
			& IIf(LenB(MailGreetings) = 0, defaultMailGreetings, MailGreetings)
		mb.Close
		
		set mb = Nothing
		set fsm = Nothing
		mailCmd = fetchSetting("mailCmd", "")
		if mailCmd = "" then LogToIntEventViewer("couldn't find setting entry for mailCmd in registry path " & RootRegPath)
		mailCmd = replace(mailCmd, "<serverName>", cdoServerName)
		mailCmd = replace(mailCmd, "<serverPort>", cdoServerPort)
		mailCmd = replace(mailCmd, "<userID>", cdoUserID)
		mailCmd = replace(mailCmd, "<passwd>", cdoPassword)
		mailCmd = replace(mailCmd, "<curPath>", curPath)
		mailCmd = replace(mailCmd, "<fromAddr>", IIf(LenB(Sender) = 0, defaultSender, Sender))
		mailCmd = replace(mailCmd, "<toAddr>", mailRecipients)
		mailCmd = replace(mailCmd, "<subject>", IIf(LenB(Subject) = 0, defaultSubject, Subject))
		mailCmd = replace(mailCmd, "<useSSL>", iif(cdoUseSSL, "-ssl", ""))
		
		mailProcExit = Exec(mailCmd, 10)
		If mailProcExit <> "" Then LogToIntEventViewer "Error after sending Mail with defined mailCmd: " & mailCmd & " returned:" & mailProcExit
		AlreadySent = True
	End Sub
	
	' Calls WshShell.Exec with c and kills the process tree after the specified timeout t (seconds)
	' Returns the created WshScriptExec object
	Private Function Exec(c, t)
		Dim e
		On Error Resume Next
		Set e = shell.Exec(c)
		If Err<>0 then 
			Exec = "Error executing """ & c & """: " & Err.Description
			Exit Function
		End If
		Do While e.Status = 0
			Wscript.Sleep 1000
			t = t - 1
			If 0 >= t Then
				Call shell.Run("taskkill /t /f /pid " & e.ProcessId, 0, True)
				Exec = "timeout reached, killed process..."
				Exit Function
			End If
		Loop
		If e.ExitCode <> 0 then
			Exec = "ExitCode=" & e.ExitCode
		Else
			Exec = ""
		End If
	End Function

	Private Function IIf( expr, truepart, falsepart )
		If expr Then 
			IIf = truepart
		Else
			IIf = falsepart
		End if
	End Function

	Private Sub LogToIntEventViewer(sErrMsg)
		wscript.echo "internal error:" & sErrMsg
		fh.WriteLine "internal error:" & sErrMsg
		shell.LogEvent LogErr, "Logger.vbs internal Error: " & sErrMsg
	End Sub

End Class