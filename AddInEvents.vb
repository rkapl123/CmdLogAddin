﻿Imports Microsoft.Office.Interop
Imports ExcelDna.Integration
Imports System.Diagnostics
Imports System.Runtime.InteropServices

''' <summary>AddIn Connection class, also handling Events from Excel (Open, Close, Activate)</summary>
<ComVisible(True)>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the app object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application
        theHostApp = ExcelDnaUtil.Application
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
        aLogger = New Logger()
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Try
            theHostApp = Nothing
        Catch ex As Exception
            aLogger.LogToEventViewer("DBAddin unloading error: " + ex.Message)
        End Try
    End Sub


    ''' <summary>open workbook: check get Arguments and start a Macro</summary>
    ''' <param name="Wb"></param>
    Private Sub Application_WorkbookOpen(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin And Not StartMakroDone Then
            getArgumentsAndStartMakro()
        End If
    End Sub

End Class
