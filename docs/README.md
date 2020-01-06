# CmdLogAddin

## Cmdline Arguments

CmdLogAddin starts a macro passed in the commandline of Excel and passes any arguments given after that macro.  

Usage: Call Excel with a the filename (for opening readonly after /r) and provide args to be passed after /e:  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /r TestExcelCmdArgFetching.xls /e/<start|startExt>/<MakroToStart>/<arg1 for Macro>/<arg2 for Macro>/.../<arg28 for Macro>`

In starting command:  
* pass arguments to workbook_open (see below) with arguments arg1, arg2, arg3  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" TestExcelCmdArgFetching.xls /e/arg1/arg2/arg3`
* start excel procedure testsub in TestExcelCmdArgFetching.xls with arguments arg1, arg2, arg3  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /r TestExcelCmdArgFetching.xls /e/start/testsub/arg1/arg2/arg3`
* start "internal" excel procedure start which is already loaded in excel (TestStart.xla already loaded, if necessary move TestStart.xla into XLSTART folder)  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /r TestExcelCmdArgFetching.xls /e/start/TestStart.xla!start`
* start excel external procedure "start" (Workbook in same directory as TestExcelCmdArgFetching.xls, assumed to be current directory (has to be set before calling getArguments())  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /r TestExcelCmdArgFetching.xls /e/startExt/Test.xla!start`
* start excel external procedure (Workbook in same directory as TestExcelCmdArgFetching.xls)  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" %~dp0TestExcelCmdArgFetching.xls /e/startExt/Test.xla!start`
* start excel external procedure (Workbook in same directory as TestExcelCmdArgFetching.xls) with argument arg1  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" %~dp0TestExcelCmdArgFetching.xls /e/startExt/Test.xla!start/arg1`
* start excel external procedure (Workbook in different directory as TestExcelCmdArgFetching.xls)  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" %~dp0TestExcelCmdArgFetching.xls /e/startExt/'C:\dev\CmdLogAddin\TestExcelCmdArgFetchingExt.xls'!testMacro/arg1`

## Logging

CmdLogAddin provides a logging tool to be used in VBA.  

Usage: First create a logger object:  
`Set theLogger = CreateObject("LogAddin.Logger")`

and initialise this object using the setProperties Method (all arguments are optional and have default values, except the CallingObject):  
* theCallingObject .. The calling excel workbook ...
* theLogLevel ..  (ERROR 1,  WARN 2, INFO 4, DEBUG 8), default = 4
* theLogFilePath .. where to write the logfile (LogFilePath), defaults to theCallingObject's path
* theEnv .. environment, empty if production, used to append env to LogFilePath for test/other environments
* theCaller .. if caller is not the callingObject (commonCaller) then this can be used to identify the active caller (in case of an addin handling multiple workbooks..).  
 Can include the full path to the calling workbook/document/..,  
 the Caller's name will be extracted by using last "\" as separator  
* theMailRecipients .. comma separated list of the error mail recipients
* theSubject .. the error mail's subject
* writeToEventLog .. should messages be written to the windows event log (true) or to a file (false)
* theSender .. the Sender of the sent error mails
* theMailIntro .. the intro for the error mail's body
* theMailGreetings .. the greetings for the error mail's body, body looks as follows:  
    [MailIntro] (executed in: [commonCaller], current caller: [Caller]):  
    [LogLine]  
    [logPathMsg]  
    [MailGreetings]  
* overrideCommonCaller .. whether to override CallingObjectName (filename to log to) with theCaller

Example:  
`theLogger.setProperties ThisWorkbook, theEnv:=env, theLoglevel:=8, theLogFilePath:="Logs", theMailRecipients:="admin@somewhere.com"`

Log messages are written by using methods LogError, LogWarn, LogInfo, LogDebug and LogFatal (ends excel application):  
`theLogger.LogError "testLog logging error"`  
`theLogger.LogWarn "testLog logging warning"`  
`theLogger.LogInfo "testLog logging info"`  
`theLogger.LogDebug "testLog logging debug"`  

Caller settings can also be changed within the active session:  
`theLogger.setProperties ThisWorkbook, , , ,"theTestLog.ext"`  
`theLogger.LogError "theTestLog.ext: logging error"`  
`theLogger.LogWarn "theTestLog.ext: logging warning"`  
`theLogger.LogInfo "theTestLog.ext: logging info"`  
`theLogger.LogDebug "theTestLog.ext: logging debug"`  

### Registry Settings 

Default Values are taken from the registry, located in `[HKCU\Software\VB and VBA Program Settings\LogAddin\Settings]`:  

Is Authentication required, then we need below 5 settings, otherwise do not authenticate  
`"cdoAuthentRequired"="False"`

UserID/Password for SMTP Authentication, if required  
`"cdoUserID"=""`
`"cdoPassword"=""`

SSL Authentication used?  
`"cdoUseSSL"="False"`

Maximum time to try to establish a connection to the SMTP server in seconds  
`"cdoConnectiontimeout"="60"`

SMTP Servername  
`"cdoServerName"="YourSMTPServerName"`

SMTP Serverport (default unsecure: 25)  
`"cdoServerPort"="25"`

In case of internal errors or problesm with settings, try to send to this  
`"cdoInternalErrMailRcpt"="MAIL-address1@domain, MAIL-address2@domain"`

Default Subject, Sender, Intro and Greetings for error mails...  
`"defaultSubject"="Batch Process Error"`  
`"defaultSender"="Administrator@domain"`  
`"defaultMailIntro"="Following error occured in batch process"`  
`"defaultMailGreetings"="regards, your Errorlog..."`  

Format for logentry timestamp  
`"timeStampFormat"="dd.MM.yyyy HH:mm:ss"`

Layout for logentries: first column logentry0, then logentry1, .. logentryN. The values (timestamp, loglevel, caller, logmessage) are fixed in the code but can be arranged differently, additional columns can be added as well.  

e:USERNAME is indicating an environment variable that can be fetched (e.g. e:COMPUTERNAME)
`"logentry0"="timestamp"`  
`"logentry1"="loglevel"`  
`"logentry2"="caller"`  
`"logentry3"="e:USERNAME"`  
`"logentry4"="logmessage"`  

CmdLogAddin is distributed under the [GNU Public License V3](http://www.gnu.org/copyleft/gpl.html).
