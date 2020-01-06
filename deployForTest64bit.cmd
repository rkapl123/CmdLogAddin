Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
set source=bin\Release
)
copy /Y %source%\CmdLogAddin-AddIn64-packed.xll "%appdata%\Microsoft\AddIns\CmdLogAddin.xll"
copy /Y %source%\CmdLogAddin.pdb "%appdata%\Microsoft\AddIns"
copy /Y %source%\CmdLogAddin.dll.config "%appdata%\Microsoft\AddIns\CmdLogAddin.xll.config"
pause
