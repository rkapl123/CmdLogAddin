Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

'---------------------------------------------------------------------------------------
' Class     : Logger
' Purpose   : main class to be used in clients: set theLogger = createObject("LogAddin.Logger")
'             doing logging with LogError, LogWarn, LogInfo, LogDebug and LogStream
'             Logger properties are set with setProperties.
'---------------------------------------------------------------------------------------
<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDispatch)>
Public Class Logger '<ProgID("LogAddin.Logger")>

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function AttachConsole(dwProcessId As Integer) As Boolean

    End Function

    Private CallerInfo As String
    Private callingObject As Object
    Private callingObjectPath As String
    Private LogLevel As EventLogEntryType
    Private LogFilePath As String
    Private env As String
    Private commonCaller As String
    Private Caller As String
    Private callerFullPath As String
    Private mailRecipients As String
    Private Subject As String
    Private doEventLog As Boolean
    Private Sender As String
    Private MailIntro As String
    Private MailGreetings As String
    Private mirrorToStdOut As Boolean
    Private AlreadySent As Boolean

    Private cdoUserID As String
    Private cdoPassword As String
    Private cdoUseSSL As Boolean
    Private cdoConnectiontimeout As Integer
    Private cdoAuthentRequired As Boolean 'false
    Private cdoServerName As String  '= "SomeSMTPServer"   'Name or IP of Remote SMTP Server
    Private cdoServerPort As Integer '= 25             'Server port (typically 25)
    Private cdoInternalErrMailRcpt As String  '

    Private defaultSubject As String  '"Batch Process Error"
    Private defaultSender As String  '"Administrator"
    Private defaultMailIntro As String  '"Folgender Fehler trat in batch process auf "
    Private defaultMailGreetings As String  '"liebe Grüße schickt der Fehleradmin..."

    Private timeStampFormat As String  '"DD.MM.YYYY HH:mm:ss"
    Private logentry() As String

    Public Sub New()
        cdoUserID = fetchSetting("cdoUserID", "")
        cdoPassword = fetchSetting("cdoPassword", "")
        cdoUseSSL = CBool(fetchSetting("cdoUseSSL", False))
        cdoConnectiontimeout = CInt(fetchSetting("cdoConnectiontimeout", 60))
        cdoAuthentRequired = CBool(fetchSetting("cdoAuthentRequired", False))
        cdoServerName = fetchSetting("cdoServerName", "YOURSMTPSERVERNAME")    'Name or IP of Remote SMTP Server
        cdoServerPort = CInt(fetchSetting("cdoServerPort", 25))               'Server port (typically 25)
        cdoInternalErrMailRcpt = fetchSetting("cdoInternalErrMailRcpt", "internalErrAdmin1,internalErrAdmin2")

        defaultSubject = fetchSetting("defaultSubject", "Batch Process Error")
        defaultSender = fetchSetting("defaultSender", "Administrator")
        defaultMailIntro = fetchSetting("defaultMailIntro", "Folgender Fehler trat in batch process auf ")
        defaultMailGreetings = fetchSetting("defaultMailGreetings", "liebe Grüße schickt der Fehleradmin...")

        timeStampFormat = fetchSetting("timeStampFormat", "dd.MM.yyyy HH:mm:ss")
        Dim i As Integer
        i = 0
        Dim entry As String = fetchSetting("logentry" & i, vbNullString)
        While Len(entry) > 0
            ReDim Preserve logentry(i)
            logentry(i) = entry
            i += 1
            entry = fetchSetting("logentry" & i, vbNullString)
        End While
        mirrorToStdOut = False
    End Sub

    ' fetchSetting
    '
    ' encapsulates setting fetching (currently registry)
    '---------------------------------------------------------------------------------------
    Private Function fetchSetting(Key As String, defaultValue As Object) As Object
        fetchSetting = GetSetting("LogAddin", "Settings", Key, defaultValue)
    End Function

    ' setProperties
    '
    ' sets properties for the Logger object, all parameters optional except theCallingObject !
    ' theCallingObject .. the calling object (excel workbook, word document, access project, etc..),
    '                  must have a "Name" and "Path" property and a "Quit" method (to allow LogFatal to end it)...
    '
    ' The other arguments are optional:
    ' theLogLevel ..  (ERROR 1,  WARN 2, INFO 4, DEBUG 8), default = 4
    ' theLogFilePath .. where to write the logfile (LogFilePath), defaults to callingObject's path
    ' theEnv .. environment, empty if production, used to append env to LogFilePath for test/other environments
    ' theCaller .. if caller is not the callingObject (commonCaller) then this can be used to
    '               identify the active caller (in case of an addin handling multiple workbooks/documents/..).
    '               Can include the full path to the calling workbook/document/..,
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
    Public Sub setProperties(Optional theCallingObject As Excel.Workbook = Nothing, Optional theLogLevel As Integer = 4, Optional theLogFilePath As String = Nothing,
        Optional theEnv As String = Nothing, Optional theCaller As String = Nothing, Optional theMailRecipients As String = Nothing,
        Optional theSubject As String = Nothing, Optional writeToEventLog As Boolean = False, Optional theSender As String = Nothing, Optional theMailIntro As String = Nothing,
        Optional theMailGreetings As String = Nothing, Optional overrideCommonCaller As Boolean = False, Optional doMirrorToStdOut As Boolean = False)

        AlreadySent = False
        If theCallingObject Is Nothing Then
            MsgBox("Logger.setProperties sets properties for the Logger object" & vbCrLf & vbCrLf &
               " - theCallingObject .. the calling excel workbook," & vbCrLf & vbCrLf &
               "These arguments are optional (default=empty/false if not specified):" & vbCrLf &
               " - theLogLevel ..  (ERROR 1,  WARN 2, INFO 4, DEBUG 8), default = 4" & vbCrLf &
               " - theLogFilePath .. where to write the logfile (LogFilePath), defaults to callingObject's path" & vbCrLf &
               " - theEnv .. environment, empty if production, used to append env to LogFilePath for test/other environments" & vbCrLf &
               " - theCaller .. if caller is not the callingObject (commonCaller) then this can be used to" & vbCrLf &
               "      * identify the active caller (in case of an addin handling multiple workbooks..)." & vbCrLf &
               "      * Can include the full path to the calling workbook..," & vbCrLf &
               "      * the Caller's name will be extracted by using last \ as separator" & vbCrLf &
               " - theMailRecipients .. comma separated list of the error mail recipients" & vbCrLf &
               " - theSubject .. the error mail's subject" & vbCrLf &
               " - writeToEventLog .. should messages be written to the windows event log (true) or to a file (false)" & vbCrLf &
               " - theSender .. the Sender of the sent error mails" & vbCrLf &
               " - theMailIntro .. the intro for the error mail's body" & vbCrLf &
               " - theMailGreetings .. the greetings for the error mail's body, body looks as follows:" & vbCrLf &
               "   <MailIntro> (executed in: <commonCaller>, current caller: <Caller>):" & vbCrLf &
               "   <logLine>" & vbCrLf &
               "   <logPathMsg>" & vbCrLf &
               "   <MailGreetings>" & vbCrLf &
               " - overrideCommonCaller .. override CallingObjectName (filename to log to) with given parameter theCaller (true)" & vbCrLf &
               " - doMirrorToStdOut .. mirror log messages to the standard output (true)", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Logger.setProperties")
            Exit Sub
        End If
        On Error Resume Next
        callingObject = theCallingObject
        commonCaller = callingObject.Name
        callingObjectPath = callingObject.Path
        If Not IsNothing(theCaller) Then
            callerFullPath = theCaller
            Caller = theCaller
            If InStr(1, callerFullPath, "\") Then Caller = Mid$(callerFullPath, InStrRev(callerFullPath, "\") + 1)
        End If
        If Len(callerFullPath) = 0 Then
            Caller = commonCaller
            callerFullPath = callingObjectPath & "\" & commonCaller
        End If
        CallerInfo = "Caller: " & Caller & ",callerFullPath: " & callerFullPath
        On Error GoTo setProperties_Err
        If Not (IsNothing(theLogFilePath)) Then LogFilePath = theLogFilePath
        If Len(LogFilePath) = 0 Then LogFilePath = callingObjectPath
        ' if no absolute path was given (drive mapping or unc), prepend callingObjectPath
        If Left(LogFilePath, 2) <> "\\" And InStr(1, LogFilePath, ":") = 0 Then LogFilePath = callingObjectPath & "\" & LogFilePath

        LogLevel = theLogLevel
        If Not (IsNothing(theEnv)) Then env = theEnv
        If Not (IsNothing(theMailRecipients)) Then mailRecipients = theMailRecipients
        If Not (IsNothing(theSubject)) Then Subject = theSubject
        doEventLog = writeToEventLog
        If Not (IsNothing(theSender)) Then Sender = theSender
        If Not (IsNothing(theMailIntro)) Then MailIntro = theMailIntro
        If Not (IsNothing(theMailGreetings)) Then MailGreetings = theMailGreetings
        If overrideCommonCaller Then commonCaller = theCaller
        mirrorToStdOut = doMirrorToStdOut
        If doMirrorToStdOut Then AttachConsole(-1)
        Exit Sub

setProperties_Err:
        Dim ErrDesc As String : ErrDesc = "Error: " & Err.Description & ", line " & Erl() & " in Logger.setProperties"
        LogToEventViewer("commonCaller = " & commonCaller & ", callingObjectPath = " & callingObjectPath & ", callerFullPath = " & callerFullPath & "Caller = " & Caller & ", callingObject.Name = " & callingObject.Name & ", mirrorToStdOut = " & mirrorToStdOut & ", LogLevel = " & LogLevel & ", env =" & env & ", mailRecipients = " & mailRecipients & ", Subject = " & Subject & ", doEventLog = " & doEventLog & ", Sender = " & Sender & ", MailIntro = " & MailIntro & ", MailGreetings = " & MailGreetings)
        Err.Raise(513, "Logger.setProperties", ErrDesc)
    End Sub


    ' LogWrite
    ' writes log information messages to defined Logfiles
    '
    ' Args:
    ' LogMessage .. Message to be logged
    ' LogPrio .. (optional) priority level (ERROR 1,  WARN 2, INFO 4, DEBUG 8)
    Private Sub LogWrite(LogMessage As String, LogPrio As EventLogEntryType)

        Dim theLogFileStr, MailFileLink, FileMessage As String
        FileMessage = ""

        On Error GoTo LogWrite_Err
1:      Dim timestamp = Date.Now().ToString(timeStampFormat, System.Globalization.CultureInfo.InvariantCulture)
        Dim i As Integer
2:      For i = 0 To UBound(logentry)
3:          If LCase$(logentry(i)) = "timestamp" Then FileMessage = FileMessage & timestamp & vbTab
4:          If LCase$(logentry(i)) = "loglevel" Then FileMessage = FileMessage & Choose(LogPrio, "ERROR", "WARN", "", "INFO", "", "", "", "DEBUG") & vbTab
5:          If LCase$(logentry(i)) = "caller" Then FileMessage = FileMessage & "file://" & Replace(callerFullPath, " ", "%20") & vbTab
6:          If Left$(LCase$(logentry(i)), 2) = "e:" Then FileMessage = FileMessage & Environ$(Mid$(logentry(i), 3)) & vbTab
7:          If LCase$(logentry(i)) = "logmessage" Then FileMessage = FileMessage & LogMessage & vbTab
        Next
8:      FileMessage = Left$(FileMessage, Len(FileMessage) - 1)
        If mirrorToStdOut Then System.Console.WriteLine(FileMessage)
        If doEventLog Then
9:          LogToEventViewer(LogMessage, LogPrio, 0)
10:         MailFileLink = "Log in Event Viewer on machine \\" & Environ$("COMPUTERNAME")
        Else
11:         theLogFileStr = IIf(Left$(LogFilePath, 2) = "\\" Or Mid$(LogFilePath, 2, 2) = ":\", "", callingObjectPath & "\") & LogFilePath & "\" & IIf(Len(env) > 0, env & "\", vbNullString) & commonCaller & ".log"
12:         Dim outputFile As StreamWriter = New StreamWriter(theLogFileStr, True, System.Text.Encoding.GetEncoding(1252))
13:         outputFile.WriteLine(FileMessage)
14:         outputFile.Close() ' close the file
15:         MailFileLink = "Logfile in file://" & Replace(LogFilePath, " ", "%20") & "\" & IIf(Len(env) > 0, env & "\", vbNullString) & commonCaller & ".log"
        End If
16:     If Len(mailRecipients) > 0 And LogPrio = EventLogEntryType.Error And Not AlreadySent Then sendMail(FileMessage, MailFileLink)
        Exit Sub

LogWrite_Err:
        LogToEventViewer(IIf(doEventLog, "", "trying to Log to file: " & theLogFileStr) & ", Error: " & Err.Description & ", line " & Erl() & " in Logger.LogWrite")
    End Sub

    ' LogError, LogWarn, LogInfo, LogDebug
    '
    ' writes a log message LogMessage having appropriate priority
    ' (As shown in funtion name, usually INFO, ERROR.., etc) depending on log level set
    ' in global const Loglevel (0 = ERROR, 1 = WARN, 2 = INFO, 3 = DEBUG)
    Public Sub LogError(LogMessage As String)
        On Error GoTo LogError_Err
        Err.Clear()
        LogWrite(LogMessage, EventLogEntryType.Error)
        Exit Sub

LogError_Err:
        LogToEventViewer("Error: " & Err.Description & ", line " & Erl() & " in Logger.LogError")
    End Sub

    Public Sub LogWarn(LogMessage As String)
        On Error GoTo LogWarn_Err
        If LogLevel >= EventLogEntryType.Warning Then
            LogWrite(LogMessage, EventLogEntryType.Warning)
        End If
        Exit Sub

LogWarn_Err:
        LogToEventViewer("Error: " & Err.Description & ", line " & Erl() & " in Logger.LogWarn")
    End Sub

    Public Sub LogInfo(LogMessage As String)
        On Error GoTo LogInfo_Err
        If LogLevel >= EventLogEntryType.Information Then
            LogWrite(LogMessage, EventLogEntryType.Information)
        End If
        Exit Sub

LogInfo_Err:
        LogToEventViewer("Error: " & Err.Description & ", line " & Erl() & " in Logger.LogInfo")
    End Sub

    Public Sub LogDebug(LogMessage As String)
        On Error GoTo LogDebug_Err
        If LogLevel >= EventLogEntryType.SuccessAudit Then
            LogWrite(LogMessage, EventLogEntryType.SuccessAudit)
        End If
        Exit Sub

LogDebug_Err:
        LogToEventViewer("Error: " & Err.Description & ", line " & Erl() & " in Logger.LogDebug")
    End Sub

    Public Sub LogFatal(LogMessage As String)
        LogError(LogMessage)
        callingObject.Parent.DisplayAlerts = False
        callingObject.Parent.Quit
    End Sub

    ' LogToEventViewer
    '
    ' Logs sErrMsg of eEventType in eCategory to EventLog
    Public Sub LogToEventViewer(sErrMsg As String, Optional eEventType As EventLogEntryType = EventLogEntryType.Error, Optional sendIntErrMail As Boolean = False)
        Dim eventLog As EventLog = New EventLog("Application")
        EventLog.WriteEntry(".NET Runtime", sErrMsg, eEventType, 1000)
        If sendIntErrMail Then sendMail(sErrMsg, "", True)
    End Sub

    ' sendMail
    '
    ' sends an error mail containing the logged line (logLine) and a hyperlink to the logfile (logPathMsg)
    Private Sub sendMail(logLine As String, logPathMsg As String, Optional internalErr As Boolean = False)

        ' construct mail message
        Dim Subject As String = IIf(internalErr, "LogAddin Internal Error !", IIf(Len(Subject) = 0, defaultSubject, Subject))
        Dim FromAddr As String = IIf(Len(Sender) = 0, defaultSender, Sender)
        If Len(FromAddr) = 0 Then
            LogToEventViewer("objMessage.From is empty, please ensure valid sender !")
            Exit Sub
        End If
        Dim ToAddr As String = IIf(internalErr, cdoInternalErrMailRcpt, mailRecipients)
        If Len(ToAddr) = 0 Then
            LogToEventViewer("objMessage.To is empty, please ensure valid receiver !")
            Exit Sub
        End If
        Dim TextBody As String = IIf(internalErr, "LogAddin Internal Error occurred (COMPUTERNAME:" & Environ$("COMPUTERNAME") & ",USERNAME:" & Environ$("USERNAME") & ",CallerInfo: " & CallerInfo & "): " & vbLf & vbLf & logLine, IIf(Len(MailIntro) = 0, defaultMailIntro, MailIntro) & "(executed in: " & commonCaller & ", current caller: " & Caller & "):" & vbLf & vbLf _
           & logLine & vbLf & vbLf _
           & logPathMsg & vbLf & vbLf _
           & IIf(Len(MailGreetings) = 0, defaultMailGreetings, MailGreetings))
        Dim objMessage As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage(FromAddr, ToAddr, Subject, TextBody)

        ' send the message
        Dim client As System.Net.Mail.SmtpClient = New System.Net.Mail.SmtpClient(cdoServerName, cdoServerPort)
        ' Add credentials if the SMTP server requires them.
        Try
            client.Timeout = cdoConnectiontimeout
            client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials
            client.Send(objMessage)
        Catch ex As Exception
            Try
                client.EnableSsl = cdoUseSSL
                client.UseDefaultCredentials = False
                If cdoAuthentRequired Then client.Credentials = New System.Net.NetworkCredential(cdoUserID, cdoPassword)
                client.Send(objMessage)
            Catch ex1 As Exception
                LogToEventViewer("Error: " & ex1.Message & " in Logger.sendMail; mailRecipients: " & mailRecipients & ",defaultSender: " & defaultSender & ",Sender: " & Sender & "objMessage.From: " & FromAddr)
            End Try
        End Try
        AlreadySent = True
    End Sub

End Class