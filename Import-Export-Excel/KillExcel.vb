'Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Runtime.InteropServices

'Public Class KillExcel
'    <DllImport("user32.dll", SetLastError:=True)> _
'    Private Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, _
'                                                     ByRef lpdwProcessId As Integer) As Integer
'    End Function

'    Private xlApp As Excel.Application

'    'Private Sub btnStartExcel_Click(ByVal sender As System.Object, _
'    '                                ByVal e As System.EventArgs) Handles btnStartExcel.Click
'    '    xlApp = New Excel.Application
'    '    xlApp.Visible = True
'    'End Sub

'    Private Sub btnKillExcel()

'        If xlApp IsNot Nothing Then
'            Dim excelProcessId As Integer
'            GetWindowThreadProcessId(New IntPtr(xlApp.Hwnd), excelProcessId)

'            If excelProcessId > 0 Then
'                Dim ExcelProc As Process = Process.GetProcessById(excelProcessId)
'                If ExcelProc IsNot Nothing Then ExcelProc.Kill()
'            End If
'        End If
'    End Sub
'End Class
