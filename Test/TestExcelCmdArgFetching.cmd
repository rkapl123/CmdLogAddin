rem need to make this folder for logfiles
mkdir ..\..\Test
@echo off
if exist "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" (
	echo 64bit office 16
	set excelexe="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
) 
if exist "C:\Program Files\Microsoft Office\root\Office15\EXCEL.EXE" (
	echo 64bit office 15
	set excelexe="C:\Program Files\Microsoft Office\root\Office15\EXCEL.EXE"
)
if exist "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" (
	echo 32bit office 16
	set excelexe="C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
)
if exist "C:\Program Files (x86)\Microsoft Office\root\Office14\EXCEL.EXE" (
	echo 32bit office 14
	set excelexe="C:\Program Files (x86)\Microsoft Office\root\Office14\EXCEL.EXE"
)
if exist "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" (
	echo 32bit office 14
	set excelexe="C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE"
)
@echo pass args to workbook open
%excelexe% /e/arg1/arg2/arg3 TestExcelCmdArgFetching.xls 
@echo start excel procedure testsub
%excelexe% /e/start/testsub/arg1/arg2/arg3 TestExcelCmdArgFetching.xls
@echo start excel procedure testsub with excel hidden. This is in a different workbook because of the Workbook.Open proc logging the messages visibly in TestExcelCmdArgFetching.xls !
%excelexe% /e/starthidden/testsub/arg1/arg2/arg3 TestExcelCmdArgFetchingExt.xls 
@echo start excel external procedure with loaded addin
%excelexe% /e/start/TestExcelAddin.xlam!testsub/passedArg TestExcelCmdArgDoNothing.xls 
@echo start excel external procedure with external workbook
%excelexe% /e/startExt/TestExcelCmdArgFetchingExt.xls!testMacro TestExcelCmdArgDoNothing.xls 
@echo start excel external procedure with external workbook having absolute path
%excelexe% /e/startExt/'%~dp0TestExcelCmdArgFetchingExt.xls'!testMacro/passedArgument TestExcelCmdArgDoNothing.xls 
