set source=bin\Release

copy /Y %source%\CmdLogAddin-AddIn-packed.xll \\oebfacoat\dfs\SOFTWARE\Makro\InstalliertInAppDataRoamingMicrosoftAddins
copy /Y %source%\CmdLogAddin.dll.config \\oebfacoat\dfs\SOFTWARE\Makro\InstalliertInAppDataRoamingMicrosoftAddins\CmdLogAddin-AddIn-packed.xll.config
pause

