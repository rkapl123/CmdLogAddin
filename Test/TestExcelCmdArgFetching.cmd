rem need to make this folder for logfiles
mkdir ..\..\Test
@echo off
if exist "C:\Program Files\Microsoft Office\root\Office16" (
	echo 64bit office
	set excelexe="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
) else (
	echo 32bit office
	set excelexe="C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE"
)
@echo pass args to workbook open
%excelexe% TestExcelCmdArgFetching.xls /e/arg1/arg2/arg3
@echo start excel procedure testsub
%excelexe% TestExcelCmdArgFetching.xls /e/start/testsub/arg1/arg2/arg3
@echo start excel procedure testsub with excel hidden
%excelexe% TestExcelCmdArgFetching.xls /e/starthidden/testsub/arg1/arg2/arg3
@echo start excel external procedure with loaded addin
%excelexe% TestExcelCmdArgFetching.xls /e/start/TestExcelAddin.xlam!testsub/passedArg
@echo start excel external procedure with external workbook
%excelexe% TestExcelCmdArgFetching.xls /e/startExt/TestExcelCmdArgFetchingExt.xls!testMacro
@echo start excel external procedure with external workbook having absolute path
%excelexe% TestExcelCmdArgFetching.xls /e/startExt/'%~dp0TestExcelCmdArgFetchingExt.xls'!testMacro/passedArgument
