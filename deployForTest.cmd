Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
	set source=bin\Release
)
rem copy /Y Test\TestExcelAddin.xlam "%appdata%\Microsoft\Excel\XLSTART\"
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y %source%\CmdLogAddin-AddIn64-packed.xll "%appdata%\Microsoft\AddIns\CmdLogAddin.xll"
	copy /Y %source%\CmdLogAddin.pdb "%appdata%\Microsoft\AddIns"
	rem copy /Y %source%\CmdLogAddin.dll.config "%appdata%\Microsoft\AddIns\CmdLogAddin.xll.config"
) else (
	echo 32bit office
	copy /Y %source%\CmdLogAddin-AddIn-packed.xll "%appdata%\Microsoft\AddIns\CmdLogAddin.xll"
	copy /Y %source%\CmdLogAddin.pdb "%appdata%\Microsoft\AddIns"
	rem copy /Y %source%\CmdLogAddin.dll.config "%appdata%\Microsoft\AddIns\CmdLogAddin.xll.config"
)
echo copying Distribution
copy /Y %source%\CmdLogAddin-AddIn-packed.xll Distribution\CmdLogAddin32.xll
copy /Y %source%\CmdLogAddin-AddIn64-packed.xll Distribution\CmdLogAddin64.xll
copy /Y vbscriptLogger\testLog.vbs Distribution\
copy /Y vbscriptLogger\Logger.vbs Distribution\
copy /Y vbscriptLogger\mailsend.exe Distribution\
copy /Y vbscriptLogger\mailsend-go.exe Distribution\
copy /Y Test\TestExcelAddin.xlam Distribution\
copy /Y Test\TestExcelCmdArgFetching.cmd Distribution\
copy /Y Test\TestExcelCmdArgFetching.xls Distribution\
copy /Y Test\TestExcelCmdArgFetchingExt.xls Distribution\
pause
