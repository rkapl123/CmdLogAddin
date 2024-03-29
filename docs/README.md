# CmdLogAddin

## Cmdline Arguments

CmdLogAddin starts a macro passed in the commandline of Excel and passes any arguments given after that macro. Additionally it has a builtin logger that writes to a logfile, windows eventlog and sends mails in case of errors.  

### Installation

* Dependencies/Prerequisites
	* .NET 4 or higher (usually distributed with Windows)

Download the zip package from the latest release in [https://github.com/rkapl123/CmdLogAddin/tags](https://github.com/rkapl123/CmdLogAddin/tags), unzip to any location and run deployAddin.cmd in the folder Distribution.
This copies CmdLogAddin32/64.xll (depending on the bitness of your Office Installation) to your %appdata%\Microsoft\AddIns folder and starts Excel for activating CmdLogAddin (adding it to the registered Addins).


### Usage

Call Excel with a the filename (for opening readonly after /r) and provide args to be passed after /e:  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /r /e/<start|startExt>/<MakroToStart>/<arg1 for Macro>/<arg2 for Macro>/.../<arg28 for Macro> TestExcelCmdArgFetching.xls`

In the starting commandline (can be in a cmd script or in the task scheduler):  
* arguments arg1, arg2, arg3 simply passed to Excel(how to fetch them see below)  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /e/arg1/arg2/arg3 TestExcelCmdArgFetching.xls`
* start excel procedure testsub in TestExcelCmdArgFetching.xls with arguments arg1, arg2, arg3  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /r /e/start/testsub/arg1/arg2/arg3 TestExcelCmdArgFetching.xls`
* start "internal" excel procedure start which is already loaded in excel (TestExcelAddin.xlam assumed to be loaded at startup, if necessary move xlam into XLSTART folder)  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /r /e/start/TestExcelAddin.xlam!start TestExcelCmdArgFetching.xls`
* start excel external procedure (Workbook in same directory as TestExcelCmdArgFetching.xls)  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /e/startExt/Test.xla!start %~dp0TestExcelCmdArgFetching.xls`
* start excel external procedure (Workbook in same directory as TestExcelCmdArgFetching.xls) with argument arg1  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /e/startExt/Test.xla!start/arg1 %~dp0TestExcelCmdArgFetching.xls`
* start excel external procedure (Workbook in different directory as TestExcelCmdArgFetching.xls)  
`"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" /e/startExt/'C:\dev\CmdLogAddin\TestExcelCmdArgFetchingExt.xls'!testMacro/arg1 %~dp0TestExcelCmdArgFetching.xls`

A maximum of three switches between EXCEL.EXE and the workbook are accepted by the Addin (e.g. /r /x /t), switches after the workbook are no problem.  

When using the first method to get commandline arguments, you have to call either  

<pre lang="vb">
    CmdlineArgs = Application.Run("getCmdlineArgs")
    For Each arg In CmdlineArgs
        MsgBox ("CmdlineArg:" & arg)
    Next
</pre>
to get the excel command line (including excel binary path itself and all switches passed to it), or  

<pre lang="vb">
    ExcelPassedArgs = Application.Run("getExcelPassedArgs")
    For Each arg In ExcelPassedArgs
        MsgBox ("ExcelPassedArg:" & arg)
    Next
</pre>

to get the specially flagged (/e) excel arguments. Both `getCmdlineArgs` and `getExcelPassedArgs` have an optional debugInfo Parameter
that can be set to True to allow for additional logging of the CmdLine Arguments and the parsed excel arguments to the event log (source is .NET Runtime).

The Workbook.Open of the Workbook's VBA is called BEFORE the procedures defined in `start` or `startExt` have been executed, this is by Excel's design.  

Generally, Excel will be minimized when using the `start` switches to be unobstrusive for an unexpecting user (when fetching the arguments with `getCmdlineArgs` or `getExcelPassedArgs` on Workbook_Open, Excel is not minimized).
To further "hide" Excel, you can add `hidden` to the start switches (so `starthidden` or `startExthidden`). 
In this case, after briefly being opened, Excel will turn off visible mode and thus "hide" from the desktop and the taskbar. In case the started macro didn't quit Excel (or excel was closed by LogFatal), 
visible mode will be turned on again after finishing the called macro.  

### Known Issues

Quitting Excel from the Workbook_Open event procedure (or any subsequently called procedure) is only possible by calling a procedure on a different thread by using Application.OnTime (Now, "NameOfQuittingProcedure")  
The same applies to the procedures invoked with the `start` switches, so for quitting Excel use `Application.OnTime`. The LogFatal call (see below) already uses this method, so there is nothing to do in this case.

Office 2019 seems to modify the command line (removing everything after the called workbook), so it is advisable to place the arguments first and the called workbook last.

## Logging

CmdLogAddin provides a logging tool to be used in VBA.  

Usage: First create a logger object:  
`Set theLogger = CreateObject("LogAddin.Logger")`

and initialize this object using the setProperties Method (all arguments are optional and have default values, except the CallingObject):  
* theCallingObject .. The calling excel workbook ...
* theLogLevel ..  (ERROR 1,  WARN 2, INFO 4, DEBUG 8), default = 4
* theLogFilePath .. where to write the log file (LogFilePath), defaults to theCallingObject's path
* theEnv .. environment, empty if production, used to append theEnv to LogFilePath for test/other environments
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
* doMirrorToStdOut .. Boolean used for mirroring logging to a trace dialog

Calling setProperties without any argument brings two helper message boxes that display the usage information.

Example:  
`theLogger.setProperties ThisWorkbook, theEnv:="Test", theLoglevel:=8, theLogFilePath:="Logs", theMailRecipients:="admin@somewhere.com"`

Log messages are written by using methods LogDebug, LogInfo, LogWarn, LogError (sends an error mail using `System.Net.Mail`) and LogFatal (ends excel application):  
`theLogger.LogDebug "testLog logging debug"`  
`theLogger.LogInfo "testLog logging info"`  
`theLogger.LogWarn "testLog logging warning"`  
`theLogger.LogError "testLog logging error"`  
`theLogger.LogFatal "testLog logging fatal error"`  

Caller settings can also be changed within the active session (here setting the Caller to `theTestLog.ext`):  
`theLogger.setProperties ThisWorkbook, , , ,"theTestLog.ext"`  
`theLogger.LogError "theTestLog.ext: logging error"`  
`theLogger.LogWarn "theTestLog.ext: logging warning"`  
`theLogger.LogInfo "theTestLog.ext: logging info"`  
`theLogger.LogDebug "theTestLog.ext: logging debug"`  

### Registry Settings 

Configurations for `System.Net.Mail` are defined in the registry values starting with `cdo` (legacy naming)
Default Values are taken from the registry, located in `[HKCU\Software\VB and VBA Program Settings\LogAddin\Settings]`:  

Is Authentication required, then we need below 3 settings, otherwise do not authenticate  
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

Format for logentry timestamp (has to conform to [.NET Custom Date and Time Format Strings](https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings)) and is 
used with the `timeStampCulture` culture settings (default: empty culture = invariant).  
`"timeStampFormat"="dd.MM.yyyy HH:mm:ss"`
`"timeStampCulture"="de-DE"`

Layout for logentries: first column logentry0, then logentry1, .. logentryN. The values (timestamp, loglevel, caller, logmessage) are fixed in the code but can be arranged differently, additional columns can be added as well.  
`e:` is indicating an environment variable (e.g. e:COMPUTERNAME or e:USERNAME) that can be fetched in this context. Example:  
`"logentry0"="timestamp"`  
`"logentry1"="loglevel"`  
`"logentry2"="caller"`  
`"logentry3"="e:USERNAME"`  
`"logentry4"="logmessage"`  

Debugging into event viewer: To add debug trace messages into the event viewer (source being .NET Runtime) add the following registry setting:  
`"debug"="true"`  

# VB-script Logger

To also have a logger for vb-script (the old Log-Addin provided an Active-X loadable COM addin here), there is now a pure vb-script based Class in the file Logger.vbs

You can add a logger with:
```VBScript 
loggerHome = "place\where\Logger.vbs\is\located"
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(loggerHome & "Logger.vbs", 1).ReadAll
```

After adding the logger, set environment based on folder name and set the properties of the logger:
```VBScript 
If InStr(1, Wscript.ScriptFullName, "Test") > 0 Then theEnv = "Test"
' Here theMailRecipients is the current user. Can also be some other hardcoded mail address...
theMailRecipients = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERNAME%") & "@yourdomain.com"
theLogger.setProperties theCallingObject = Wscript, theLogLevel = 4, theLogFilePath, theEnv, theCaller, theMailRecipients, theSubject, writeToEventLog, theSender, theMailIntro, theMailGreetings, overrideCommonCaller, doMirrorToStdOut
``` 

In the code, add logging as follows:
```VBScript 
theLogger.LogError "error" ' also sends an errormail
theLogger.LogWarn "warning"
theLogger.LogInfo "info"
theLogger.LogDebug "debug"
theLogger.LogFatal "fatal error (ends execution)" ' also sends an errormail
```

CmdLogAddin is distributed under the [GNU Public License V3](http://www.gnu.org/copyleft/gpl.html).
