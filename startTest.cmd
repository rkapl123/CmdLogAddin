mkdir Test
@echo off
if exist "C:\Program Files\Microsoft Office\root\Office16" (
	echo 64bit office
	C:\Windows\Syswow64\cmd.exe /C C:\dev\CmdLogAddin\TestExcelCmdArgFetching.cmd 64
) else (
	echo 32bit office
	C:\Windows\Syswow64\cmd.exe /C C:\dev\CmdLogAddin\TestExcelCmdArgFetching.cmd 32
)