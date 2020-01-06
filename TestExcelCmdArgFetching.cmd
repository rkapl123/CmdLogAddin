if "%1" == "64" (
	set excelexe="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
) else (
	set excelexe="C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE"
)
rem start excel procedure testsub
%excelexe% TestExcelCmdArgFetching.xls /e/start/testsub/arg1/arg2/arg3
rem start excel external procedure with loaded addin
%excelexe% TestExcelCmdArgFetching.xls /e/start/OebfaBar.xla!LogoDeutsch
rem start excel external procedure with external workbook
%excelexe% TestExcelCmdArgFetching.xls /e/startExt/TestExcelCmdArgFetchingExt.xls!testMacro
rem start excel external procedure with external workbook having absolute path
%excelexe% TestExcelCmdArgFetching.xls /e/startExt/'%~dp0TestExcelCmdArgFetchingExt.xls'!testMacro/passedArgument
