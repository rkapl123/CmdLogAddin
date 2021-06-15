rem copy Addin and settings...
@echo off
set /p answer="Enter Y to stop Excel (if running) and continue deployment of CmdLogAddin:"
if "%answer:~,1%" NEQ "Y" exit /b
taskkill /IM "Excel.exe" /F
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y CmdLogAddin64.xll "%appdata%\Microsoft\AddIns\CmdLogAddin.xll"
) else (
	echo 32bit office
	copy /Y CmdLogAddin32.xll "%appdata%\Microsoft\AddIns\CmdLogAddin.xll"
)
rem start Excel and install Addin there..
cscript //nologo switchToCmdLogAddin.vbs
pause
