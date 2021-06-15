Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Diagnostics ' needed for EventLogEntryType !!
Imports System.IO
Imports System.Runtime.InteropServices

Public Module Globals
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

End Module

''' <summary>main class to be used in clients: set theLogger = createObject("LogAddin.Logger")
'''          doing logging with LogError, LogWarn, LogInfo, LogDebug and LogFatal
'''          Logger properties are set with setProperties.
''' </summary>
<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDispatch)>
<ProgId("LogAddin.Logger")>
Public Class Logger
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function AttachConsole(dwProcessId As Integer) As Boolean
        ' might implement this again, unclear how this works in mixed 32/64 bit environments...
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
    Private timeStampCulture As String '""
    Private timeStampFormat As String  '"dd.MM.yyyy HH:mm:ss"
    Private logentry() As String

    ''' <summary></summary>
    Public Sub New()
        cdoUserID = fetchSetting("cdoUserID", "")
        cdoPassword = fetchSetting("cdoPassword", "")
        cdoUseSSL = CBool(fetchSetting("cdoUseSSL", False))
        cdoConnectiontimeout = CInt(fetchSetting("cdoConnectiontimeout", 0))
        cdoAuthentRequired = CBool(fetchSetting("cdoAuthentRequired", False))
        cdoServerName = fetchSetting("cdoServerName", "YOURSMTPSERVERNAME")    'Name or IP of Remote SMTP Server
        cdoServerPort = CInt(fetchSetting("cdoServerPort", 25))               'Server port (typically 25)
        cdoInternalErrMailRcpt = fetchSetting("cdoInternalErrMailRcpt", "internalErrAdmin1,internalErrAdmin2")

        defaultSubject = fetchSetting("defaultSubject", "Batch Process Error")
        defaultSender = fetchSetting("defaultSender", "Administrator")
        defaultMailIntro = fetchSetting("defaultMailIntro", "Folgender Fehler trat in batch process auf ")
        defaultMailGreetings = fetchSetting("defaultMailGreetings", "liebe Grüße schickt der Fehleradmin...")

        timeStampFormat = fetchSetting("timeStampFormat", "dd.MM.yyyy HH:mm:ss")
        timeStampCulture = fetchSetting("timeStampCulture", "")
        Dim i As Integer
        i = 0
        Dim entry As String = fetchSetting("logentry" & i, vbNullString)
        While Len(entry) > 0
            ReDim Preserve logentry(i)
            logentry(i) = entry
            i += 1
            entry = fetchSetting("logentry" & i, vbNullString)
        End While
        If i = 0 Then MsgBox("No logentry defined in registry settings !")
        mirrorToStdOut = False
    End Sub

    ''' <summary>encapsulates setting fetching (currently registry)</summary>
    ''' <param name="Key">key of setting</param>
    ''' <param name="defaultValue">default value to be used if no setting given</param>
    ''' <returns>setting value</returns>
    Private Function fetchSetting(Key As String, defaultValue As Object) As Object
        fetchSetting = GetSetting("LogAddin", "Settings", Key, defaultValue)
    End Function

    ''' <summary>sets properties for the Logger object, all parameters optional except theCallingObject</summary>
    ''' <param name="theCallingObject">the calling object (excel workbook)</param>
    ''' <param name="theLogLevel">log level (ERROR 1,  WARN 2, INFO 4, DEBUG 8), default = 4</param>
    ''' <param name="theLogFilePath">where to write the logfile (LogFilePath), defaults to callingObject's path</param>
    ''' <param name="theEnv">environment, empty if production, used to append env to LogFilePath for test/other environments</param>
    ''' <param name="theCaller">if caller is not the callingObject (commonCaller) then this can be used to identify the active caller (in case of an addin handling multiple workbooks/documents/..).
    '''               Can include the full path to the calling workbook/document/.., the Caller's name will be extracted by using last "\" as separator</param>
    ''' <param name="theMailRecipients">comma separated list of the error mail recipients</param>
    ''' <param name="theSubject">the error mail's subject</param>
    ''' <param name="writeToEventLog">should messages be written to the windows event log (true) or to a file (false)</param>
    ''' <param name="theSender">the Sender of the sent error mails</param>
    ''' <param name="theMailIntro">the intro for the error mail's body</param>
    ''' <param name="theMailGreetings">the greetings for the error mail's body, body looks as follows:
    ''' [MailIntro] (executed in: [commonCaller], current caller: [Caller]):
    ''' [logLine]
    ''' [logPathMsg]
    ''' [MailGreetings]</param>
    ''' <param name="overrideCommonCaller">whether to override CallingObjectName (filename to log to) with theCaller</param>
    ''' <param name="doMirrorToStdOut">used for mirroring to stdout, not implemented now (32/64 bit problems), left for backward compatibility</param>
    Public Sub setProperties(Optional theCallingObject As Excel.Workbook = Nothing, Optional theLogLevel As Integer = 4, Optional theLogFilePath As String = Nothing,
        Optional theEnv As String = Nothing, Optional theCaller As String = Nothing, Optional theMailRecipients As String = Nothing,
        Optional theSubject As String = Nothing, Optional writeToEventLog As Boolean = False, Optional theSender As String = Nothing, Optional theMailIntro As String = Nothing,
        Optional theMailGreetings As String = Nothing, Optional overrideCommonCaller As Boolean = False, Optional doMirrorToStdOut As Boolean = False)

        AlreadySent = False
        If theCallingObject Is Nothing Then
            Dim sModuleInfo As String = vbNullString

            ' get module info for buildtime (FileDateTime of xll):
            For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
                Dim sModule As String = tModule.FileName
                If sModule.ToUpper.Contains("CMDLOGADDIN") Then
                    sModuleInfo = FileDateTime(sModule).ToString()
                End If
            Next

            MsgBox("Logger.setProperties sets properties for the Logger object" & vbCrLf & vbCrLf &
               "- theCallingObject .. the calling excel workbook," & vbCrLf & vbCrLf &
               "These arguments are optional (default=empty/false if not specified):" & vbCrLf &
               "- theLogLevel ..  (ERROR 1,  WARN 2, INFO 4, DEBUG 8), default = 4" & vbCrLf &
               "- theLogFilePath .. where to write the logfile (LogFilePath), defaults to callingObject's path" & vbCrLf &
               "- theEnv .. environment, empty if production, used to append env to LogFilePath for test/other environments" & vbCrLf &
               "- theCaller .. if caller is not the callingObject (commonCaller) then this can be used to" & vbCrLf &
               "     * identify the active caller (in case of an addin handling multiple workbooks..)." & vbCrLf &
               "     * Can include the full path to the calling workbook..," & vbCrLf &
               "     * the Caller's name will be extracted by using last \ as separator" & vbCrLf &
               "- theMailRecipients .. comma separated list of the error mail recipients" & vbCrLf &
               "- theSubject .. the error mail's subject" & vbCrLf &
               "- writeToEventLog .. should messages be written to the windows event log (true) or to a file (false)" & vbCrLf &
               "- theSender .. the Sender of the sent error mails" & vbCrLf &
               "- theMailIntro .. the intro for the error mail's body" & vbCrLf &
               "- theMailGreetings .. the greetings for the error mail's body, body looks as follows:" & vbCrLf &
               "   <MailIntro> (executed in: <commonCaller>, current caller: <Caller>):" & vbCrLf &
               "   <logLine>" & vbCrLf &
               "   <logPathMsg>" & vbCrLf &
               "   <MailGreetings>" & vbCrLf &
               "- overrideCommonCaller .. override CallingObjectName (filename to log to) with given parameter theCaller (true)",
                   MsgBoxStyle.Information + MsgBoxStyle.OkOnly, String.Format("CmdLogAddin Version {0} Buildtime {1}", My.Application.Info.Version.ToString, sModuleInfo))
            MsgBox("Logging is then done with following five methods:" & vbCrLf &
               "- Logger.LogDebug(msg) .. writes msg if debug level (theLogLevel in setProperties) = 8" & vbCrLf &
               "- Logger.LogInfo(msg) .. writes msg if debug level >= 4" & vbCrLf &
               "- Logger.LogWarn(msg) .. writes msg if debug level >= 2" & vbCrLf &
               "- Logger.LogError(msg) .. writes msg if debug level >= 1, additionally an error mail is sent to theMailRecipients" & vbCrLf &
               "- Logger.LogFatal(msg) .. writes msg if debug level >= 1, additionally to the error mail the host application is shut down" & vbCrLf & vbCrLf &
               "Author/Website: " & My.Application.Info.CompanyName.ToString & vbCrLf &
               "License: " & My.Application.Info.Copyright.ToString,
                   MsgBoxStyle.Information + MsgBoxStyle.OkOnly, String.Format("CmdLogAddin Version {0} Buildtime {1}", My.Application.Info.Version.ToString, sModuleInfo))
            Exit Sub
        End If
        On Error Resume Next
1:      callingObject = theCallingObject
2:      commonCaller = callingObject.Name
3:      callingObjectPath = callingObject.Path
4:      If Not IsNothing(theCaller) Then
5:          callerFullPath = theCaller
6:          Caller = theCaller
7:          If InStr(1, callerFullPath, "\") Then Caller = Mid$(callerFullPath, InStrRev(callerFullPath, "\") + 1)
        End If
8:      If Len(callerFullPath) = 0 Then
9:          Caller = commonCaller
10:         callerFullPath = callingObjectPath & "\" & commonCaller
        End If
11:     CallerInfo = "Caller: " & Caller & ",callerFullPath: " & callerFullPath
        On Error GoTo setProperties_Err
12:     If Not (IsNothing(theLogFilePath)) Then LogFilePath = theLogFilePath
13:     If Len(LogFilePath) = 0 Then LogFilePath = callingObjectPath
        ' if no absolute path was given (drive mapping or unc), prepend callingObjectPath
14:     If Left(LogFilePath, 2) <> "\\" And InStr(1, LogFilePath, ":") = 0 Then LogFilePath = callingObjectPath & "\" & LogFilePath

        LogLevel = theLogLevel
15:     If Not (IsNothing(theEnv)) Then env = theEnv
16:     If Not (IsNothing(theMailRecipients)) Then mailRecipients = theMailRecipients
17:     If Not (IsNothing(theSubject)) Then Subject = theSubject
18:     doEventLog = writeToEventLog
19:     If Not (IsNothing(theSender)) Then Sender = theSender
20:     If Not (IsNothing(theMailIntro)) Then MailIntro = theMailIntro
21:     If Not (IsNothing(theMailGreetings)) Then MailGreetings = theMailGreetings
22:     If overrideCommonCaller Then commonCaller = theCaller
        Exit Sub

setProperties_Err:
        Dim ErrDesc As String : ErrDesc = "Error: " & Err.Description & ", line " & Erl() & " in Logger.setProperties"
        LogToEventViewer(ErrDesc & ", commonCaller = " & commonCaller & ", callingObjectPath = " & callingObjectPath & ", callerFullPath = " & callerFullPath & "Caller = " & Caller & ", callingObject.Name = " & callingObject.Name & ", mirrorToStdOut = " & mirrorToStdOut & ", LogLevel = " & LogLevel & ", env =" & env & ", mailRecipients = " & mailRecipients & ", Subject = " & Subject & ", doEventLog = " & doEventLog & ", Sender = " & Sender & ", MailIntro = " & MailIntro & ", MailGreetings = " & MailGreetings)
    End Sub


    ''' <summary>writes log information messages to defined Logfiles</summary>
    ''' <param name="LogMessage">Message to be logged</param>
    ''' <param name="LogPrio">priority level (ERROR 1,  WARN 2, INFO 4, DEBUG 8)</param>
    Private Sub LogWrite(LogMessage As String, LogPrio As EventLogEntryType)
        Dim theLogFileStr, MailFileLink, FileMessage As String
        FileMessage = ""

        On Error GoTo LogWrite_Err
1:      Dim timestamp = Date.Now().ToString(timeStampFormat, System.Globalization.CultureInfo.CreateSpecificCulture(timeStampCulture))
        Dim i As Integer
2:      For i = 0 To UBound(logentry)
3:          If LCase$(logentry(i)) = "timestamp" Then FileMessage = FileMessage & timestamp & vbTab
4:          If LCase$(logentry(i)) = "loglevel" Then FileMessage = FileMessage & Choose(LogPrio, "ERROR", "WARN", "", "INFO", "", "", "", "DEBUG", "", "", "", "", "", "", "", "FATAL") & vbTab
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
12:         Dim outputFile As StreamWriter = New StreamWriter(theLogFileStr, True, System.Text.Encoding.Default)
13:         outputFile.WriteLine(FileMessage)
14:         outputFile.Close() ' close the file
15:         MailFileLink = "Logfile in file://" & Replace(LogFilePath, " ", "%20") & "\" & IIf(Len(env) > 0, env & "\", vbNullString) & commonCaller & ".log"
        End If
16:     If Len(mailRecipients) > 0 And LogPrio = EventLogEntryType.Error And Not AlreadySent Then sendMail(FileMessage, MailFileLink)
        Exit Sub

LogWrite_Err:
        LogToEventViewer(IIf(doEventLog, "", "trying to Log to file: " & theLogFileStr) & ", Error: " & Err.Description & ", line " & Erl() & " in Logger.LogWrite")
    End Sub

    ''' <summary>writes the log message LogMessage having appropriate priority (as shown in function name) and ends Excel</summary>
    ''' <param name="LogMessage"></param>
    Public Sub LogFatal(LogMessage As String)
        quittingApp = True
        LogWrite(LogMessage, EventLogEntryType.FailureAudit)
        ExcelDnaUtil.Application.OnTime(DateTime.Now, "QuitApp")
    End Sub

    ''' <summary>writes the log message LogMessage having appropriate priority (as shown in function name)</summary>
    ''' <param name="LogMessage"></param>
    Public Sub LogError(LogMessage As String)
        LogWrite(LogMessage, EventLogEntryType.Error)
    End Sub

    ''' <summary>writes the log message LogMessage having appropriate priority (as shown in function name) depending on log level set</summary>
    ''' <param name="LogMessage"></param>
    Public Sub LogWarn(LogMessage As String)
        If LogLevel >= EventLogEntryType.Warning Then
            LogWrite(LogMessage, EventLogEntryType.Warning)
        End If
    End Sub

    ''' <summary>writes the log message LogMessage having appropriate priority (as shown in function name) depending on log level set</summary>
    ''' <param name="LogMessage"></param>
    Public Sub LogInfo(LogMessage As String)
        If LogLevel >= EventLogEntryType.Information Then
            LogWrite(LogMessage, EventLogEntryType.Information)
        End If
    End Sub

    ''' <summary>writes the log message LogMessage having appropriate priority (as shown in function name) depending on log level set</summary>
    ''' <param name="LogMessage"></param>
    Public Sub LogDebug(LogMessage As String)
        If LogLevel >= EventLogEntryType.SuccessAudit Then
            LogWrite(LogMessage, EventLogEntryType.SuccessAudit)
        End If
    End Sub

    ''' <summary>Logs sErrMsg of eEventType in eCategory to EventLog</summary>
    ''' <param name="sErrMsg"></param>
    ''' <param name="eEventType"></param>
    ''' <param name="sendIntErrMail"></param>
    Public Sub LogToEventViewer(sErrMsg As String, Optional eEventType As EventLogEntryType = EventLogEntryType.Error, Optional sendIntErrMail As Boolean = False)
        Dim eventLog As EventLog = New EventLog("Application")
        ' .Net Runtime is always there if .Net is installed
        EventLog.WriteEntry(".NET Runtime", sErrMsg, eEventType, 1000)
        If sendIntErrMail Then sendMail(sErrMsg, "", True)
    End Sub

    ''' <summary>sends an Error mail containing the logged line (logLine) And a hyperlink To the logfile (logPathMsg)</summary>
    ''' <param name="logLine"></param>
    ''' <param name="logPathMsg"></param>
    ''' <param name="internalErr"></param>
    Private Sub sendMail(logLine As String, logPathMsg As String, Optional internalErr As Boolean = False)

        ' construct mail message
        Dim theSubject As String = IIf(internalErr, "LogAddin Internal Error !", IIf(Len(Subject) = 0, defaultSubject, Subject))
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
        Dim objMessage As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage(FromAddr, ToAddr, theSubject, TextBody)

        ' send the message
        Dim client As System.Net.Mail.SmtpClient = New System.Net.Mail.SmtpClient(cdoServerName, cdoServerPort)
        ' Add credentials if the SMTP server requires them.
        Try
            If cdoConnectiontimeout > 0 Then client.Timeout = cdoConnectiontimeout * 1000
            client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials
            client.Send(objMessage)
        Catch ex As Exception
            Try
                client.EnableSsl = cdoUseSSL
                client.UseDefaultCredentials = False
                If cdoAuthentRequired Then client.Credentials = New System.Net.NetworkCredential(cdoUserID, cdoPassword)
                client.Send(objMessage)
            Catch ex1 As Exception
                LogToEventViewer("Error: " & ex1.Message & " in Logger.sendMail; mailRecipients: " & mailRecipients & ",defaultSender: " & defaultSender & ",Sender: " & Sender & ",FromAddr: " & FromAddr)
            End Try
        End Try
        AlreadySent = True
    End Sub

End Class