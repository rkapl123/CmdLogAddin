Imports Microsoft.Office.Interop
Imports ExcelDna.Integration
Imports ExcelDna.ComInterop
Imports System.Diagnostics ' needed for Trace.Listeners !!

''' <summary>AddIn Connection class, handling Open/Close Events from Addin</summary>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the app object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ComServer.DllRegisterServer()
        Application = ExcelDnaUtil.Application
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Try
            ComServer.DllUnregisterServer()
            Application = Nothing
        Catch ex As Exception
            internalLogToEventViewer("CmdLogAddin unloading error: " + ex.Message)
        End Try
    End Sub

    ''' <summary>open workbook: check get Arguments and start a Macro</summary>
    ''' <param name="Wb">opened workbook</param>
    Private Sub Application_WorkbookOpen(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin And Not StartMakroDone And Not ArgsProhibited Then
            getArgumentsAndStartMakro(debugInfo:=Boolean.Parse(CmdLineFetcher.fetchSetting("debug", "false")))
        End If
    End Sub

End Class
