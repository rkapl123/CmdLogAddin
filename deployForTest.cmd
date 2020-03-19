Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
	set source=bin\Release
)
copy /Y Test\TestExcelAddin.xlam "%appdata%\Microsoft\Excel\XLSTART\"
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
pause
