Imports System.Runtime.InteropServices
Imports ExcelDna.ComInterop
Imports ExcelDna.Integration

''' <summary>All global Variables for CmdLogAddin</summary>
<ComVisible(False)>
Module Globals
    ' Global objects/variables for all
    Public StartMakroDone As Boolean = False
    Public ArgsProhibited As Boolean = False ' prohibit Argument fetching during opening workbooks when App.Run Macro is loaded

    Public Sub QuitApp()
        ExcelDnaUtil.Application.DisplayAlerts = False
        ExcelDnaUtil.Application.Quit()
    End Sub
End Module
