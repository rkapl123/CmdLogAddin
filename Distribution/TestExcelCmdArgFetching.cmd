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
if exist "C:\Program Files (x86)\Microsoft Office\root\Office14\EXCEL.EXE" (
	echo 32bit office 14
	set excelexe="C:\Program Files (x86)\Microsoft Office\root\Office14\EXCEL.EXE"
)
if exist "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" (
	echo 32bit office 14
	set excelexe="C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE"
)
@echo pass args to workbook open
%excelexe% TestExcelCmdArgFetching.xls /e/arg1/arg2/arg3
@echo start excel procedure testsub
%excelexe% TestExcelCmdArgFetching.xls /e/start/testsub/arg1/arg2/arg3
@echo start excel procedure testsub with excel hidden. This is in a different workbook because of the Workbook.Open proc logging the messages visibly in TestExcelCmdArgFetching.xls !
%excelexe% TestExcelCmdArgFetchingExt.xls /e/starthidden/testsub/arg1/arg2/arg3
@echo start excel external procedure with loaded addin
%excelexe% TestExcelCmdArgDoNothing.xls /e/start/TestExcelAddin.xlam!testsub/passedArg
@echo start excel external procedure with external workbook
%excelexe% TestExcelCmdArgDoNothing.xls /e/startExt/TestExcelCmdArgFetchingExt.xls!testMacro
@echo start excel external procedure with external workbook having absolute path
%excelexe% TestExcelCmdArgDoNothing.xls /e/startExt/'%~dp0TestExcelCmdArgFetchingExt.xls'!testMacro/passedArgument
