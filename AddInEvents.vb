Imports Microsoft.Office.Interop
Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports ExcelDna.ComInterop


''' <summary>AddIn Connection class, handling Open/Close Events from Addin</summary>
<ComVisible(False)>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the app object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ComServer.DllRegisterServer()
        Application = ExcelDnaUtil.Application
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Try
            Application = Nothing
            ComServer.DllUnregisterServer()
        Catch ex As Exception
            MsgBox("CmdLogAddin unloading error: " + ex.Message)
        End Try
    End Sub

    ''' <summary>open workbook: check get Arguments and start a Macro</summary>
    ''' <param name="Wb">opened workbook</param>
    Private Sub Application_WorkbookOpen(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin And Not StartMakroDone And Not ArgsProhibited Then
            getArgumentsAndStartMakro()
        End If
    End Sub

End Class
