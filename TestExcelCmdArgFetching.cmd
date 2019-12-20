rem start excel procedure testsub
"C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" TestExcelCmdArgFetching.xls /e/start/testsub/arg1/arg2/arg3
rem start excel external procedure with loaded addin
"C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" TestExcelCmdArgFetching.xls /e/start/OebfaBar.xla!LogoDeutsch
rem start excel external procedure with external workbook
"C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" TestExcelCmdArgFetching.xls /e/startExt/TestExcelCmdArgFetchingExt.xls!testMacro
rem start excel external procedure with external workbook having absolute path
"C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" TestExcelCmdArgFetching.xls /e/startExt/'C:\dev\CmdLogAddin\trunk\TestExcelCmdArgFetchingExt.xls'!testMacro/passedArgument
