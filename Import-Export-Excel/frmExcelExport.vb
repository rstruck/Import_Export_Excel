'Option Strict On
'Option Infer On
Imports FlexCel.Core
Imports FlexCel.XlsAdapter

Imports System.Runtime.InteropServices
Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.Configuration

Imports System.Collections.Generic
Imports System.Linq

Imports System.Text.RegularExpressions

'Imports Microsoft.Office

'Imports Excel
Imports System.IO



'Public Enum XlCopyAction
'    Paste
'    Insert
'End Enum

'Public Enum XlBorders
'    xlDiagonalDown = MsExcel.XlBordersIndex.xlDiagonalDown
'    xlDiagonalUp = MsExcel.XlBordersIndex.xlDiagonalUp
'    xlEdgeLeft = MsExcel.XlBordersIndex.xlEdgeLeft
'    xlEdgeTop = MsExcel.XlBordersIndex.xlEdgeTop
'    xlEdgeBottom = MsExcel.XlBordersIndex.xlEdgeBottom
'    xlEdgeRigth = MsExcel.XlBordersIndex.xlEdgeRight
'    xlInsideHorizontal = MsExcel.XlBordersIndex.xlInsideHorizontal
'    xlInsideVertical = MsExcel.XlBordersIndex.xlInsideVertical
'End Enum

'Public Enum XlLineStyle
'    xlContinuous = MsExcel.XlLineStyle.xlContinuous
'    xlDash = MsExcel.XlLineStyle.xlDash
'    xlDashDot = MsExcel.XlLineStyle.xlDashDot
'    xlDashDotDot = MsExcel.XlLineStyle.xlDashDotDot
'    xlDot = MsExcel.XlLineStyle.xlDot
'    xlDouble = MsExcel.XlLineStyle.xlDouble
'    xlLineStyleNone = MsExcel.XlLineStyle.xlLineStyleNone
'    xlSlantDashDot = MsExcel.XlLineStyle.xlSlantDashDot
'End Enum

'Public Enum XlBorderWight
'    xlHairline = MsExcel.XlBorderWeight.xlHairline
'    xlMedium = MsExcel.XlBorderWeight.xlMedium
'    xlThick = MsExcel.XlBorderWeight.xlThick
'    xlThin = MsExcel.XlBorderWeight.xlThin
'End Enum

'Public Enum XlFontStyle
'    xlStrikethrough
'    xlSuperscript
'    xlSubscript
'    xlOutlineFont
'    xlShadow
'    xlBold
'    xlItalic
'    xlUnderlineDouble
'    xlUnderlineSingle
'    xlNone
'End Enum

'Public Enum XlCellFormat
'    xlWrapTest
'    xlShrinkToFit
'    xlNone
'End Enum

'Public Enum XlHAlign
'    xlCenter = MsExcel.XlHAlign.xlHAlignCenter
'    xlCenterAcrossSelection = MsExcel.XlHAlign.xlHAlignCenterAcrossSelection
'    xlDistributed = MsExcel.XlHAlign.xlHAlignDistributed
'    xlFill = MsExcel.XlHAlign.xlHAlignFill
'    xlGeneral = MsExcel.XlHAlign.xlHAlignGeneral
'    xlJustify = MsExcel.XlHAlign.xlHAlignJustify
'    xlLeft = MsExcel.XlHAlign.xlHAlignLeft
'    xlRight = MsExcel.XlHAlign.xlHAlignRight
'End Enum

'Public Enum XlVAlign
'    xlBottom = MsExcel.XlVAlign.xlVAlignBottom
'    xlCenter = MsExcel.XlVAlign.xlVAlignCenter
'    xlDistributed = MsExcel.XlVAlign.xlVAlignDistributed
'    xlJustify = MsExcel.XlVAlign.xlVAlignJustify
'    xlTop = MsExcel.XlVAlign.xlVAlignTop
'End Enum

'Public Enum XlFillPattern
'    xlNone = MsExcel.XlPattern.xlPatternNone
'    xlSolid = MsExcel.XlPattern.xlPatternSolid
'    xlAuto = MsExcel.XlPattern.xlPatternAutomatic
'    xlChecker = MsExcel.XlPattern.xlPatternChecker
'    xlCrissCross = MsExcel.XlPattern.xlPatternCrissCross
'    xlDown = MsExcel.XlPattern.xlPatternDown
'    xlUp = MsExcel.XlPattern.xlPatternUp
'    xlHorizontal = MsExcel.XlPattern.xlPatternHorizontal
'    xlVertical = MsExcel.XlPattern.xlPatternVertical
'    xlGrid = MsExcel.XlPattern.xlPatternGrid
'    xlGray8 = MsExcel.XlPattern.xlPatternGray8
'    xlGray16 = MsExcel.XlPattern.xlPatternGray16
'    xlGray25 = MsExcel.XlPattern.xlPatternGray25
'    xlGray50 = MsExcel.XlPattern.xlPatternGray50
'    xlGray75 = MsExcel.XlPattern.xlPatternGray75
'    xlLightDown = MsExcel.XlPattern.xlPatternLightDown
'    xlLightHorizontal = MsExcel.XlPattern.xlPatternLightHorizontal
'    xlLightUp = MsExcel.XlPattern.xlPatternLightUp
'    xlLightVertical = MsExcel.XlPattern.xlPatternLightVertical
'    xlSemiGray75 = MsExcel.XlPattern.xlPatternSemiGray75
'End Enum

Public Class frmExcelExport
    Inherits Form
    ' Private _spreadsheetControl As SpreadsheetControl = Nothing
    Dim strm As System.IO.Stream
    'Dim oApp As New Excel.Application
    'Dim oBooks As Excel.Workbooks

    'Dim oBook As Excel.Workbook
    'Dim oSheet As Excel.Worksheet
    'Dim oRange As Excel.Range

    'Dim conOle As OleDbConnection
    'Dim daOle As OleDbDataAdapter, ds As DataSet

    '<DllImport("user32.dll", SetLastError:=True)> _
    'Private Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, _
    '                                                 ByRef lpdwProcessId As Integer) As Integer
    ''End Function
    Public processtype As Integer
    Public dsTableStructure As New DataSet
    Public dsChanges As New DataSet
    Public dsImportSplit As New DataSet
    Public dsSqlTables As New DataSet
    ''Public ds = New DataSet
    'Private Property strm As Object

    Public itemarrlist As ArrayList = New ArrayList
    Public itemspecarrlist As ArrayList = New ArrayList
    Public itemassortmentarrlist As ArrayList = New ArrayList
    Public itemmaterialarrlist As ArrayList = New ArrayList


    ' Public spreadsheet As Spreadsheet = New Bytescout.Spreadsheet.Spreadsheet
    ' Public worksheet As Bytescout.Spreadsheet.Worksheet
    'Private Property Processtype As String



    'Private Sub btnStartExcel_Click(ByVal sender As System.Object, _
    '                                ByVal e As System.EventArgs) Handles btnStartExcel.Click
    '    xlApp = New Excel.Application
    '    xlApp.Visible = True
    'End Sub

    Private Sub btnKillExcel()

        'If xlApp IsNot Nothing Then
        '    Dim excelProcessId As Integer
        '    GetWindowThreadProcessId(New IntPtr(xlApp.Hwnd), excelProcessId)

        '    If excelProcessId > 0 Then
        '        Dim ExcelProc As Process = Process.GetProcessById(excelProcessId)
        '        If ExcelProc IsNot Nothing Then ExcelProc.Kill()
        '    End If


        'End If

        Dim proc As System.Diagnostics.Process

        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next
    End Sub

    Private Sub frmExcelExport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Loads the table structures of the four tables
        ' they will be used to get the column name and verify that against the columnname in the import table.
        ' changes will be written out to another dataset that will be used to update SQL and the EXCEL spreadsheet.
        Timer1.Enabled = True
        Try
            GetTableschema("Item", dsTableStructure)
            GetTableschema("ItemSpecs", dsTableStructure)
            GetTableschema("Item_Assortments", dsTableStructure)
            GetTableschema("ItemMaterial", dsTableStructure)

            'Dim itemarrlist As ArrayList = New ArrayList
            'Dim itemspecarrlist As ArrayList = New ArrayList
            'Dim iteassortmentarrlist As ArrayList = New ArrayList
            'Dim itemmaterialarrlist As ArrayList = New ArrayList
            Me.btnOpenExcel.Enabled = False
            For Each row In dsTableStructure.Tables("Item").Rows
                itemarrlist.Add(row.Item(1))
            Next
            For Each row In dsTableStructure.Tables("ItemSpecs").Rows
                itemspecarrlist.Add(row.Item(1))
            Next
            For Each row In dsTableStructure.Tables("Item_Assortments").Rows
                itemassortmentarrlist.Add(row.Item(1))
            Next
            For Each row In dsTableStructure.Tables("ItemMaterial").Rows
                itemmaterialarrlist.Add(row.Item(1))
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnImportToExcel_Click(sender As Object, e As EventArgs) Handles btnImportToExcel.Click
        ProcessExcel(1)
    End Sub

    Private Sub btnRefreshExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefreshExcel.Click
        ' btnKillExcel()
        ProcessExcel(2)
    End Sub

    Public Sub stopExcel()
        'If xlapp IsNot Nothing Then
        '    Dim pProcess As Process()
        '    pProcess = System.Diagnostics.Process.GetProcessesByName("Excel")
        '    pProcess(0).Kill()
        'End If
    End Sub
    Private Sub ImportLinkExcel(ByRef dtimport As DataTable)
        'Dim xlApp As New MsExcel.Application
        'Dim xlWorkbook As MsExcel.Workbook = Nothing
        'Dim xlWorksheet As MsExcel.Worksheet = Nothing
        Dim lRow As Long, lCol As Long
        Dim dr As DataRow
        Dim i As Long, c As Long
        Dim fname As String = ""
        'Open the Excel file.
        Dim xls As New XlsFile(False)
        Dim StartOpen As Date = Date.Now

        Dim dialog As OpenFileDialog = New OpenFileDialog
        dialog.Filter = "Excel document(*.xlsx)|*.xlsx"
        Dim result As DialogResult = dialog.ShowDialog
        If (result = DialogResult.OK) Then
            fname = dialog.FileName
        End If
        dialog.Reset()
        xls.Open(fname)

        xls.ActiveSheet = 1
        Dim formatted As Boolean
        Dim rs As TRichString

        formatted = False
        'sheetCombo.Items.Add(xls.SheetName)

        'Dim Data As DataTable = dataSet1.Tables.Add("Sheet" & sheet.ToString())
        dtimport.BeginLoadData()
        Try
            Dim ColCount As Integer = xls.ColCountOnlyData
            'Add one column on the dataset for each used column on Excel.
            For c = 0 To ColCount - 1
                ' dtimport.Columns.Add(TCellAddress.EncodeColumn(c), GetType(String)) 'Here we will add all strings, since we do not know what we are waiting for.
                dtimport.Columns.Add(xls.GetStringFromCell(1, c + 1))
            Next c
            dtimport.Columns.Add("FirstRowFormatted")
            dtimport.Columns.Add("SecondRowFormatted")
            dtimport.Columns.Add("thirdRowFormatted")
            'ColCount = ColCount + 3
            Dim drs(ColCount + 2) As String
            Dim cellcolor
            Dim XF As Integer
            Dim RowCount As Integer = xls.RowCount
            Dim hasBKColor As Boolean
            Dim f As TFlxFormat
            For r As Integer = 2 To RowCount
                hasBKColor = False
                Array.Clear(drs, 0, drs.Length)
                'This loop will only loop on used cells. It is more efficient than looping on all the columns.
                For cIndex As Integer = xls.ColCountInRow(r) To 1 Step -1 'reverse the loop to avoid calling ColCountInRow more than once.
                    Dim Col As Integer = xls.ColFromIndex(r, cIndex)
                    ' If Col < 4 Then MsgBox("Stop and debug")

                    f = xls.GetCellVisibleFormatDef(r, Col)
                    If f.FillPattern.Pattern <> TFlxPatternStyle.Solid = True Then
                        hasBKColor = True
                    Else
                        f.FillPattern.FgColor.ColorType.ToString()
                    End If

                    If formatted Then
                        rs = xls.GetStringFromCell(r, Col)
                        drs(Col - 1) = rs.Value.ToString
                        XF = xls.GetCellFormat(r, cIndex)
                    Else
                        ' XF = 0 'This is the cell format, we will not use it here.
                        XF = xls.GetCellFormat(r, cIndex)
                        ' Debug.Print(XF)
                        Dim val As Object = xls.GetCellValueIndexed(r, cIndex, XF)

                        Dim Fmla As TFormula = TryCast(val, TFormula)
                        If Fmla IsNot Nothing Then
                            'When we have formulas, we want to write the formula result. 
                            'If we wanted the formula text, we would not need this part.
                            drs(Col - 1) = Convert.ToString(Fmla.Result)
                        Else
                            drs(Col - 1) = Convert.ToString(val)
                        End If
                    End If
                    ' f = xls.GetCellVisibleFormatDef(r, Col)
                    Dim xfs As Object
                    Dim xfcolor As Object
                    If Col < 4 Then
                        If hasBKColor = True Then
                            If f.FillPattern.BgColor.IsAutomatic Then
                                xfs = f.FillPattern.BgColor.AutomaticFillType
                            Else
                                xfs = f.FillPattern.BgColor.RGB
                            End If
                        Else
                            If f.FillPattern.FgColor.IsAutomatic Then
                                xfs = f.FillPattern.FgColor.Index
                            Else
                                xfs = f.FillPattern.FgColor.RGB
                            End If

                        End If
                        xfcolor = xfs

                        'TFlxFormat.FillPattern.BgColor.ColorType

                        'Dim f As Object

                        ' Dim cellSourceColor = f.FillPattern.FgColor.ToColor(s.Xls).ToArgb()

                        If xfs.ToString = "None" Or xfs.ToString = "4294967295" Then
                            If Col = 1 Then
                                drs(drs.Length - 3) = ""
                            End If
                            If Col = 2 Then
                                drs(drs.Length - 2) = ""
                            End If
                            If Col = 3 Then
                                drs(drs.Length - 1) = ""
                            End If
                        Else
                            ' If xfs <> "0" Then 'backgroundcolor.ToString Then
                            If Col = 1 Then
                                drs(drs.Length - 3) = "Colored"
                            End If
                            If Col = 2 Then
                                drs(drs.Length - 2) = "Colored"
                            End If
                            If Col = 3 Then
                                drs(drs.Length - 1) = "Colored"
                                ' End If
                            End If
                        End If
                        ' xfs = xls.GetCellVisibleFormat(r, Col)
                        'f =
                        'f = fmt.FillPattern.BgColor
                        'xfs = TFlxFormat(xls.GetCellFormat(r, cIndex).ToString)
                        'cellcolor = sheet.Range(row, dtcolumn).Style.Color.ToArgb.ToString
                        'cellcolorb = sheet.Range(row, dtcolumn).Style.Color.ToArgb.ToString
                        'If xfs <> "0" Then 'backgroundcolor.ToString Then
                        '    If Col = 0 Then
                        '        drs(drs.Length - 3) = "Colored"
                        '    End If
                        '    If Col = 1 Then
                        '        drs(drs.Length - 2) = "Colored"
                        '    End If
                        '    If Col = 2 Then
                        '        drs(drs.Length - 1) = "Colored"
                        '    End If
                        'Else
                        '    If Col = 0 Then
                        '        drs(drs.Length - 3) = ""
                        '    End If
                        '    If Col = 1 Then
                        '        drs(drs.Length - 2) = ""
                        '    End If
                        '    If Col = 2 Then
                        '        drs(drs.Length - 1) = ""
                        '    End If
                        'End If
                    End If
                Next cIndex
                dtimport.Rows.Add(drs)
            Next r

        Finally
            dtimport.EndLoadData()
        End Try

        Dim EndFill As Date = Date.Now
        'StatusBar.Text = String.Format("Time to load file: {0}    Time to fill dataset: {1}     Total time: {2}", (EndOpen.Subtract(StartOpen)).ToString(), (EndFill.Subtract(EndOpen)).ToString(), (EndFill.Subtract(StartOpen)).ToString())

        ' Dim workbook As Workbook = New Workbook

        ' Dim sheet As Worksheet = workbook.Worksheets(0)

        txtFile.Text = fname

        'If (File.Exists(txtFile.Text)) = True Then
        '    For column As Integer = 0 To sheet.LastColumn - 1
        '        dtimport.Columns.Add(sheet.Range(1, column + 1).Value)
        '    Next
        '    dtimport.Columns.Add("FirstRowFormatted")
        '    dtimport.Columns.Add("SecondRowFormatted")
        '    dtimport.Columns.Add("thirdRowFormatted")
        '    'For cCnt = 1 To lCol 'xlRange.Columns.Count
        '    '    dtimport.Columns.Add(CStr(xlWorksheet.Columns.Cells(1, cCnt).value))
        '    'Next
        '    Dim dtcolumn As Integer
        '    Dim backgroundcolor As String
        '    backgroundcolor = "-1"
        '    Dim cellcolor As String
        '    Dim cellcolorb As String
        '    For row As Integer = 2 To sheet.LastRow
        '        Dim data() As Object

        '        [Array].Resize(data, sheet.LastColumn + 3)

        '        For column As Integer = 0 To sheet.LastColumn - 1
        '            dtcolumn = column + 1
        '            data(column) = sheet.Range(row, dtcolumn).Value
        '            'Next
        '            'For column As Integer = 0 To 2
        '            If column < 4 Then
        '                cellcolor = sheet.Range(row, dtcolumn).Style.Color.ToArgb.ToString
        '                cellcolorb = sheet.Range(row, dtcolumn).Style.Color.ToArgb.ToString
        '                If sheet.Range(row, dtcolumn).Style.Color.ToArgb.ToString <> "-1" Then 'backgroundcolor.ToString Then
        '                    If column = 1 Then
        '                        data(sheet.LastColumn + 1) = "Colored"
        '                    End If
        '                    If column = 2 Then
        '                        data(sheet.LastColumn + 2) = "Colored"
        '                    End If
        '                    If column = 3 Then
        '                        '  data(sheet.LastColumn + 3) = "Colored"
        '                    End If
        '                Else
        '                    If column = 1 Then
        '                        data(sheet.LastColumn) = ""
        '                    End If
        '                    If column = 2 Then
        '                        data(sheet.LastColumn + 1) = ""
        '                    End If
        '                    If column = 3 Then
        '                        data(sheet.LastColumn + 2) = ""
        '                    End If
        '                End If
        '            End If
        '        Next
        '        dtimport.Rows.Add(data)
        '        'Next
        '    Next
        '    'Dim dt As New DataTable
        Dim l As Integer, n As Integer
        Dim valuesarr As String = String.Empty
        For l = 0 To dtimport.Rows.Count - 1
            Dim lst As New List(Of Object)(dtimport.Rows(l).ItemArray)
            valuesarr = ""
            For Each s As Object In lst
                valuesarr &= s.ToString
            Next
            If String.IsNullOrEmpty(valuesarr) Then
                'Remove row here, this row do not have any value
                Exit For
            End If
        Next
        ' need to delete all rows from l to the end of the table.
        For n = dtimport.Rows.Count - 1 To l Step -1
            dtimport.Rows.RemoveAt(n)
        Next

        ' workbook.SaveToFile(txtFile.Text)

        'Else
        ' MsgBox("No file to import")
        ' End If
    End Sub



    Public Sub ProcessExcel(ByRef processtype As Integer)

        Dim column As Integer = 0
        Dim dtype As Object
        Dim itemcolumn As DataColumnCollection = dsTableStructure.Tables("Item").Columns
        Dim itemspeccolumn As DataColumnCollection = dsTableStructure.Tables("ItemSpecs").Columns
        Dim itemmaterialcolumn As DataColumnCollection = dsTableStructure.Tables("ItemMaterial").Columns
        Dim item_assortmentcolumn As DataColumnCollection = dsTableStructure.Tables("Item_Assortments").Columns
        Dim itemrow() As DataRow
        Dim expression As String
        Dim value As Object
        Dim dt As New DataTable, dtimport As New DataTable
        Dim dtColumnheader As New DataTable
        Dim dsColumnheader As New DataSet

        ImportLinkExcel(dtimport)
        Try
            dsColumnheader.Tables.Add(dtimport)

            ToolStripStatusLabel1.Text = "Adding header column"
            For column = 0 To dsColumnheader.Tables(0).Columns.Count - 1
                expression = "Column_Name = '" & dsColumnheader.Tables(0).Columns(column).ColumnName & "'"
                dtype = ""
                value = ""
                itemrow = dsTableStructure.Tables("Item_Assortments").Select(expression)
                dtype = itemrow.Length
                If itemrow.Length > 0 Then
                    CheckDataTypedatatable(dt, column, dsColumnheader.Tables(0), itemrow(0)(2))
                Else
                    ' Item Material
                    itemrow = dsTableStructure.Tables("ItemMaterial").Select(expression)
                    If itemrow.Length > 0 Then
                        CheckDataTypedatatable(dt, column, dsColumnheader.Tables(0), itemrow(0)(2))
                    Else
                        ' Item Specs
                        itemrow = dsTableStructure.Tables("ItemSpecs").Select(expression)
                        If itemrow.Length > 0 Then
                            CheckDataTypedatatable(dt, column, dsColumnheader.Tables(0), itemrow(0)(2))
                        Else
                            ' Item
                            itemrow = dsTableStructure.Tables("Item").Select(expression)
                            If itemrow.Length > 0 Then
                                CheckDataTypedatatable(dt, column, dsColumnheader.Tables(0), itemrow(0)(2))
                            Else
                                dt.Columns.Add(dsColumnheader.Tables(0).Columns(column).ColumnName, System.Type.GetType("System.String"))
                            End If
                        End If
                    End If
                End If
            Next

            dtype = ""
            For Each row As DataRow In dsColumnheader.Tables(0).Rows 'firstBlankRow(xlRange)  'oRange.Rows.Count
                If IsDBNull(row(dsColumnheader.Tables(0).Columns(0))) = True And IsDBNull(row(dsColumnheader.Tables(0).Columns(1))) = True _
                        And IsDBNull(row(dsColumnheader.Tables(0).Columns(2))) = True And IsDBNull(row(dsColumnheader.Tables(0).Columns("FirstRowFormatted"))) = True _
                        And IsDBNull(row(dsColumnheader.Tables(0).Columns("SecondRowFormatted"))) = True And IsDBNull(row(dsColumnheader.Tables(0).Columns("ThirdRowFormatted"))) = True Then
                    Exit For
                Else
                    Dim dr As DataRow = dt.NewRow()
                    For column = 0 To dsColumnheader.Tables(0).Columns.Count - 1
                        dtype = dt.Columns(column).DataType
                        Select Case dtype.ToString
                            Case "System.Byte"
                                ' convert from YES/NO to 1/0
                                If row(dsColumnheader.Tables(0).Columns(column)).ToString = "YES" Then
                                    dr(column) = 1
                                Else
                                    dr(column) = 0
                                End If
                            Case "System.String"
                                dr(column) = row(dsColumnheader.Tables(0).Columns(column))
                            Case Else
                                'Case "System.Int32", "System.Int16"
                                If row(dsColumnheader.Tables(0).Columns(column)).ToString = "" Then
                                    dr(column) = DBNull.Value
                                Else
                                    dr(column) = row(dsColumnheader.Tables(0).Columns(column))
                                End If
                                'Case Else
                                'dr(column) = row(dsColumnheader.Tables(0).Columns(column))
                        End Select
                        'If CStr(dtype.ToString) = "System.Byte" Then
                        '    ' convert from YES/NO to 1/0
                        '    If row(dsColumnheader.Tables(0).Columns(column)).ToString = "YES" Then
                        '        dr(column) = 1
                        '    Else
                        '        dr(column) = 0
                        '    End If
                        'Else
                        '    If row(dsColumnheader.Tables(0).Columns(column)).ToString = "" Then
                        '        dr(column) = DBNull.Value
                        '    Else
                        '        dr(column) = row(dsColumnheader.Tables(0).Columns(column))
                        '    End If
                        'End If
                    Next
                    dt.Rows.Add(dr)
                    dt.AcceptChanges()
                End If
            Next

            dsColumnheader.Tables(0).TableName = "ImportedTable"
            dt.TableName = "GoodImporttable"
            CreateDataTables(dt, dsTableStructure, dsColumnheader.Tables(0), processtype)
        Catch ex As Exception
            MsgBox(ex.Message)
            dgvDataToExport = Nothing
        Finally

        End Try
        ' Me.Cursor = Cursors.Arrow
        Application.DoEvents()
    End Sub

    'Function firstBlankRow(ByVal rngToSearch As Excel.Range) As Long

    '    Dim R As Excel.Range
    '    Dim C As Excel.Range
    '    Dim RowIsBlank As Boolean
    '    Dim numofrows As Long = 0
    '    For Each R In rngToSearch.Rows
    '        RowIsBlank = True
    '        'If R Then
    '        For Each C In R.Cells
    '            If C.Column > 3 Then Exit For
    '            If String.IsNullOrEmpty(CStr(C.Value)) = True Then RowIsBlank = False
    '        Next C
    '        If RowIsBlank = False Then
    '            numofrows = R.Row
    '            Exit For
    '        End If
    '    Next R
    '    If numofrows = 0 Then
    '        firstBlankRow = 2
    '    Else
    '        firstBlankRow = numofrows
    '    End If
    'End Function

    'Public Sub releasemyobjects(ByRef myExcelobj)
    '    releaseObject(myExcelobj)
    'End Sub

    Public Sub CreateDataTables(ByRef dt As DataTable, ByRef dsTableStructure As DataSet, ByRef dtColumnheader As DataTable, processtype As Integer)
        Dim DtSet As System.Data.DataSet = New DataSet
        'dt is the imported spreadsheet
        Try
            dt.Columns.Add("RowNumber", System.Type.GetType("System.Int32"))
            dt.Columns.Add("TableName", System.Type.GetType("System.String"))
            dsChanges.Tables.Add()
            dsChanges.Tables(0).Columns.Add("RowNumber", System.Type.GetType("System.Int32"))
            dsChanges.Tables(0).Columns.Add("TableName", System.Type.GetType("System.String"))
            dsChanges.Tables(0).Columns.Add("ColumnName", System.Type.GetType("System.String"))
            dsChanges.Tables(0).Columns.Add("SQLValue", System.Type.GetType("System.String"))
            dsChanges.Tables(0).Columns.Add("SpreadSheetValue", System.Type.GetType("System.String"))
            dsChanges.Tables(0).Columns.Add("ColumnNumber", System.Type.GetType("System.String"))
            dsChanges.Tables(0).Columns.Add("Proposalnumber", System.Type.GetType("System.Int32"))
            dsChanges.Tables(0).Columns.Add("Rev", System.Type.GetType("System.Int16"))
            DtSet.Tables.Add(dt)

            'dt = DtSet.Tables(0)
            ' dsChanges.Tables.Add(dt.Copy)
            dsChanges.Tables.Add(dtColumnheader.Copy)
            With Me.dgvDataToExport
                .DataSource = dt
                .Refresh()
            End With
            'MyCommand = Nothing
            'MyConnection.Close()
            dtColumnheader.Columns.Add("RowNumber", System.Type.GetType("System.Int32"))
            populaterownumber(dtColumnheader)
            '  removeblankrows(dt)
            dt.AcceptChanges()
            populaterownumber(dt)



            ' ''loop thru data table 
            Dim i As Integer = dt.Rows.Count
            AcceptOrReject(dt)
            dgvDataToExport.Refresh()
            Dim ri As Integer = 0
            For Each dtrow As DataRow In dt.Rows
                If dtrow(dt.Columns("FirstRowFormatted")).ToString = "Colored" AndAlso dtrow(dt.Columns("SecondRowFormatted")).ToString = "Colored" AndAlso dtrow(dt.Columns("ThirdRowFormatted")).ToString = "Colored" Then
                Else
                    ri += 1
                    populateItemDataGrid(CLng(dtrow.Item("proposalnumber")), CInt(dtrow.Item("rev")), dsSqlTables)
                    populateItemSpecsDataGrid(CLng(dtrow.Item("proposalnumber")), CInt(dtrow.Item("rev")), dsSqlTables)
                    populateItemAssortmentsDataGrid(CLng(dtrow.Item("proposalnumber")), CInt(dtrow.Item("rev")), dsSqlTables)
                    populateItemMaterialDataGrid(CLng(dtrow.Item("proposalnumber")), CInt(dtrow.Item("rev")), dsSqlTables)
                End If
            Next

            'If processtype = 10 Then  'making sure that this will be skipped
            Splitimportedvalues(DtSet, dsTableStructure, dsImportSplit)

            Dim dColumn() As DataColumn = {dsSqlTables.Tables("Item").Columns("ProposalNumber")}
            dsSqlTables.Tables("Item").PrimaryKey = dColumn
            Dim dimportColumn() As DataColumn = {dsImportSplit.Tables("ImportItem").Columns("ProposalNumber")}
            dsImportSplit.Tables("ImportItem").PrimaryKey = dimportColumn
            ' PrintValues(dsSqlTables.Tables("Item"), "Merged With itemsTable, Schema added")
            ' CompareDatatables(dsSqlTables.Tables("Item"), dsImportSplit.Tables("ImportItem"))
            'dsSqlTables.Tables("Item").AcceptChanges()
            ' dsImportSplit.Tables("ImportItem").AcceptChanges()
            ' Add RowChanged event handler for the table. 

            '            Dim query1 = dsSqlTables.Tables("Item").AsEnumerable().[Select](Function(a) New With { _
            '    Key .ID = a("ProposalNumber").ToString() _
            '})

            '            Dim query2 = dsImportSplit.Tables("ImportItem").AsEnumerable().[Select](Function(b) New With { _
            '                Key .ID = b("ProposalNumber").ToString() _
            '            })

            '            Dim exceptResultsAB = query1.Except(query2)
            '            Dim exceptResultsBA = query2.Except(query1)

            'For Each row As DataRow In dsImportSplit.Tables("ImportItem").Rows
            '    row.SetModified()
            'Next row
            'AddHandler dsSqlTables.Tables("Item").RowChanged, AddressOf Row_Changed
            'dsSqlTables.Tables("Item").Merge(dsImportSplit.Tables("ImportItem")) '', False, MissingSchemaAction.Add)
            'PrintValues(dsSqlTables.Tables("Item"), "Merged With itemsTable, Schema added")
            'Dim xdatatable As DataTable = dsSqlTables.Tables("Item").GetChanges(DataRowState.Modified)

            compareitemchanges(DtSet, dsImportSplit, dsChanges, dsSqlTables)

            Me.dgvChanges.DataSource = dsChanges.Tables(0)
            Me.dgvChanges.Refresh()

            'If processtype = 10 Then  'making sure that this will be skipped
            compareitemspecchanges(DtSet, dsImportSplit, dsChanges, dsSqlTables)
            Me.dgvChanges.DataSource = dsChanges.Tables(0)
            Me.dgvChanges.Refresh()

            'If processtype = 10 Then  'making sure that this will be skipped
            compareitemmaterialchanges(DtSet, dsImportSplit, dsChanges, dsSqlTables)
            Me.dgvChanges.DataSource = dsChanges.Tables(0)
            Me.dgvChanges.Refresh()
            '     If processtype = 10 Then  'making sure that this will be skipped

            compareitemAssortmantchanges(DtSet, dsImportSplit, dsChanges, dsSqlTables)
            Me.dgvChanges.DataSource = dsChanges.Tables(0)
            Me.dgvChanges.Refresh()

            Select Case processtype
                Case 1  ' import
                    ImportExcel(txtFile.Text.ToString, dsChanges)
                Case 2 '' refresh
                    ExportExcel(txtFile.Text.ToString, dsChanges)
                Case 3  ' validate
                    ValidateExcel(txtFile.Text.ToString, dsChanges, dtColumnheader)
            End Select

        Catch ex As Exception
            MsgBox(ex.Message)

            dgvDataToExport = Nothing
        End Try
    End Sub
    Private Sub CompareDatatables(dtSQL As DataTable, dtImport As DataTable)
        Dim tablesAreIdentical As Boolean = True
        Dim foundIdenticalRow As Boolean = False
        Dim allFieldsAreIdentical As Boolean = False
        ' loop through first table
        For Each row As DataRow In dtSQL.Rows
            foundIdenticalRow = False

            ' loop through tempTable to find an identical row
            For Each tempRow As DataRow In dtImport.Rows
                allFieldsAreIdentical = True

                ' compare fields, if any fields are different move on to next row in tempTable
                Dim i As Integer = 0
                While i < row.ItemArray.Length AndAlso allFieldsAreIdentical
                    If Not row(i).Equals(tempRow(i)) Then

                        allFieldsAreIdentical = False
                    End If
                    i += 1
                End While

                ' if an identical row is found, remove this row from tempTable 
                '  (in case of duplicated row exist in firstTable, so tempTable needs
                '   to have the same number of duplicated rows to be considered equivalent)
                ' and move on to next row in firstTable
                'If allFieldsAreIdentical Then
                '    dtImport.Rows.Remove(tempRow)
                '    foundIdenticalRow = True
                '    Exit For
                'End If
            Next
            ' if no identical row is found for current row in firstTable, 
            ' the two tables are different
            If Not foundIdenticalRow Then
                tablesAreIdentical = False
                Exit For
            End If
        Next

        'Return tablesAreIdentical

    End Sub

    Private Sub PrintValues(ByVal table As DataTable, ByVal label As String)
        ' Display the values in the supplied DataTable:
        Debug.Print(label)
        For Each row As DataRow In table.Rows
            Debug.Print(row.RowState.ToString)
            'For Each col As DataColumn In table.Columns
            '    Debug.Print(ControlChars.Tab + " " + row(col).ToString())
            'Next col
            'Console.WriteLine()
        Next row
    End Sub

    Private Sub Row_Changed(ByVal sender As Object, _
    ByVal e As DataRowChangeEventArgs)

        Console.WriteLine("Row changed {0}{1}{2}", _
          e.Action, ControlChars.Tab, e.Row.ItemArray(0))
        ' MsgBox("Row changed {0}{1}{2}", e.Action, ControlChars.Tab, e.Row.ItemArray(0))
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk


        strm = OpenFileDialog1.OpenFile()
        ' txtFile.Text = OpenFileDialog1.FileName.ToString()
        ' txtFile.Text = "c:\Users\rstruck\Desktop\AprilTestFile.xlsx"
        '  If Not (strm Is Nothing) Then
        'insert code to read the file data
        ' strm.Close()
        ' MessageBox.Show("file closed")
        ' End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks> reads the datatable and adds on a column with the row number.</remarks>
    Private Sub populaterownumber(ByRef dt As DataTable)
        Dim i As Integer
        Try
            For Each dr As DataRow In dt.Rows
                i = i + 1
                dr.Item("RowNumber") = i
            Next
        Catch ex As Exception
            MsgBox("there was an error setting the row number")
        End Try
    End Sub

    ''' <summary>
    '''  This will populate the field TableName that was added to the imported spreadsheet.
    ''' depending on the column name it will loop through the imported rows and then throuh the dataset for the
    ''' table structures and then through every table in the dsTableStructure dataset.
    ''' </summary>
    ''' <param name="dtSet"></param>
    ''' <param name="dsTableStructure"></param>
    ''' <remarks></remarks>
    Private Sub populatetablename(ByRef dtSet As DataSet, dsTableStructure As DataSet)
        ' Dim i As Integer
        Try
            'For Each dr As DataRow In dt.Rows
            '    'i = i + 1
            '    'dr.Item("RowNumber") = i
            'Next
        Catch ex As Exception
            MsgBox("there was an error setting the row number")
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks> reads the datatable and adds on a column with the row number.</remarks>
    Private Sub removeblankrows(ByRef dt As DataTable)
        Try
            For Each dr As DataRow In dt.Rows
                If dr.Item(0).ToString = "" Then
                    dr.Delete()
                End If
            Next
        Catch ex As Exception
            MsgBox("there was an error deleting blank rows")
        End Try
    End Sub

    Private Sub populateItemDataGrid(ByRef proposalnumber As Long, ByRef rev As Integer, ByRef dtset As DataSet)
        Itemdatatable(proposalnumber, rev, dtset)
    End Sub

    Private Sub populateItemSpecsDataGrid(ByRef proposalnumber As Long, ByRef rev As Integer, ByRef dtset As DataSet)
        ItemSpecsdatatable(proposalnumber, rev, dtset)
    End Sub

    Private Sub populateItemAssortmentsDataGrid(ByRef proposalnumber As Long, ByRef rev As Integer, ByRef dtset As DataSet)
        ItemAssortmentsdatatable(proposalnumber, rev, dtset)
    End Sub

    Private Sub populateItemMaterialDataGrid(ByRef proposalnumber As Long, ByRef rev As Integer, ByRef dtset As DataSet)
        ItemMaterialdatatable(proposalnumber, rev, dtset)
    End Sub

    Public Function SQLConnection() As SqlConnection
        Dim conn As New SqlConnection("Data Source=SS-SQL\SQL2012;Initial Catalog=SSDEV;Integrated Security=True")
        SQLConnection = conn
    End Function

    Public Sub Itemdatatable(ByRef proposalnumber As Long, ByRef rev As Integer, ByRef dtset As DataSet)
        Try

            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            Dim dtItem As New DataTable
            ' Dim i As Integer
            connection = SQLConnection()
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "spItemProposalRev"   '' create stored procedure to bring back the 
            command.Parameters.AddWithValue("ProposalNumber", proposalnumber)
            command.Parameters.AddWithValue("REV", rev)
            adapter = New SqlDataAdapter(command)
            adapter.Fill(dtItem)

            dtset.Tables.Add(dtItem)
            If dtset.Tables.Contains("Item") Then
                dtset.Tables("Item").Merge(dtItem)
            Else
                dtset.Tables(dtset.Tables.Count - 1).TableName = "Item"
            End If
            With Me.DataGridView1
                .DataSource = dtset.Tables("Item")
                .Refresh()
            End With
            connection.Close()

        Catch ex As Exception
            MsgBox("there was an error " & ex.Message)
        End Try
    End Sub

    Public Sub ItemSpecsdatatable(ByRef proposalnumber As Long, ByRef rev As Integer, ByRef dtset As DataSet)

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            Dim dtItemSpecs As New DataTable

            connection = SQLConnection()
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "spItemSpecsProposalRev"   '' create stored procedure to bring back the 
            command.Parameters.AddWithValue("ProposalNumber", proposalnumber)
            command.Parameters.AddWithValue("REV", rev)
            adapter = New SqlDataAdapter(command)
            adapter.Fill(dtItemSpecs)

            dtset.Tables.Add(dtItemSpecs)
            If dtset.Tables.Contains("ItemSpecs") Then
                dtset.Tables("ItemSpecs").Merge(dtItemSpecs)
            Else
                dtset.Tables(dtset.Tables.Count - 1).TableName = "ItemSpecs"
            End If
            With Me.DataGridView2
                .DataSource = dtset.Tables("ItemSpecs")
                .Refresh()
            End With
            connection.Close()

        Catch ex As Exception
            MsgBox("there was an error " & ex.Message)
        End Try
    End Sub

    Public Sub ItemAssortmentsdatatable(ByRef proposalnumber As Long, ByRef rev As Integer, ByRef dtset As DataSet)
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            Dim dtItemAssortments As New DataTable

            connection = SQLConnection()
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "spItemAssortmentsProposalRev"   '' create stored procedure to bring back the 
            command.Parameters.AddWithValue("ProposalNumber", proposalnumber)
            command.Parameters.AddWithValue("REV", rev)
            adapter = New SqlDataAdapter(command)
            adapter.Fill(dtItemAssortments)

            dtset.Tables.Add(dtItemAssortments)
            If dtset.Tables.Contains("ItemAssortments") Then
                dtset.Tables("ItemAssortments").Merge(dtItemAssortments)
            Else
                dtset.Tables(dtset.Tables.Count - 1).TableName = "ItemAssortments"
            End If
            With Me.DataGridView3
                .DataSource = dtset.Tables("ItemAssortments")
                .Refresh()
            End With
            connection.Close()

        Catch ex As Exception
            MsgBox("there was an error " & ex.Message)
        End Try
    End Sub

    Public Sub ItemMaterialdatatable(ByRef proposalnumber As Long, ByRef rev As Integer, ByRef dtset As DataSet)

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            Dim dtItemMaterial As New DataTable

            ' Dim i As Integer
            connection = SQLConnection()
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "spItemMaterialProposalRev"   '' create stored procedure to bring back the 
            command.Parameters.AddWithValue("ProposalNumber", proposalnumber)
            command.Parameters.AddWithValue("REV", rev)
            adapter = New SqlDataAdapter(command)
            adapter.Fill(dtItemMaterial)

            dtset.Tables.Add(dtItemMaterial)
            If dtset.Tables.Contains("ItemMaterial") Then
                dtset.Tables("ItemMaterial").Merge(dtItemMaterial)
            Else
                dtset.Tables(dtset.Tables.Count - 1).TableName = "ItemMaterial"
            End If
            With Me.DataGridView4
                .DataSource = dtset.Tables("ItemMaterial")
                .Refresh()
            End With
            connection.Close()

        Catch ex As Exception
            MsgBox("there was an error " & ex.Message)
        End Try
    End Sub

    Private Sub AcceptOrReject(table As DataTable)
        ' If there are errors, try to reconcile. 
        If (table.HasErrors) Then
            If (Reconcile(table)) Then
                ' Fixed all errors.
                table.AcceptChanges()
            Else
                ' Couldn'table fix all errors.
                table.RejectChanges()
            End If
        Else
            ' If no errors, AcceptChanges.
            table.AcceptChanges()
        End If
    End Sub

    Private Function Reconcile(thisTable As DataTable) As Boolean
        Dim row As DataRow
        For Each row In thisTable.Rows
            'Insert code to try to reconcile error. 

            ' If there are still errors return immediately 
            ' since the caller rejects all changes upon error. 
            If row.HasErrors Then
                Reconcile = False
                Exit Function
            End If
        Next row
        Reconcile = True
    End Function

    Public Sub GetTableschema(ByRef tablename As String, ByRef dsTableStructure As DataSet)

        Dim sql As String = "SELECT Ordinal_Position,Column_Name,Data_Type,Is_Nullable,Character_Maximum_Length,Numeric_Precision FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" & tablename & "'"
        Dim connection As SqlConnection
        Dim adapter As SqlDataAdapter
        Dim command As New SqlCommand
        Dim dtTableStructure As New DataTable
        Dim s As Integer
        Try
            connection = SQLConnection()
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.Text
            command.CommandText = sql  '' create stored procedure to bring back the 
            adapter = New SqlDataAdapter(command)
            adapter.Fill(dtTableStructure)

            dsTableStructure.Tables.Add(dtTableStructure)
            s = (dsTableStructure.Tables.Count)

            If dsTableStructure.Tables.Contains(tablename) Then
                dsTableStructure.Tables(tablename).Merge(dtTableStructure)
            Else
                dsTableStructure.Tables(s - 1).TableName = tablename
            End If

        Catch ex As SqlException

        Finally
            connection.Close()
        End Try
    End Sub

    Public Sub Splitimportedvalues(ByRef DtSet As DataSet, ByRef dsTableStructure As DataSet, ByRef dsImportSplit As DataSet)
        'DTSet is the imported spreadsheet

        Try
            Dim destTable As DataTable = DtSet.Tables(0).Clone()

            For index As Integer = destTable.Columns.Count - 1 To 0 Step -1
                Dim columnName As String = destTable.Columns(index).ColumnName
                If Not itemarrlist.Contains(columnName) Then
                    destTable.Columns.RemoveAt(index)
                End If
            Next
            destTable.Merge(DtSet.Tables(0), False, MissingSchemaAction.Ignore)
            dsImportSplit.Tables.Add(destTable)
            dsImportSplit.Tables(0).TableName = "ImportItem"
            destTable = Nothing
            destTable = DtSet.Tables(0).Clone()
            For index As Integer = destTable.Columns.Count - 1 To 0 Step -1
                Dim columnName As String = destTable.Columns(index).ColumnName
                If Not itemspecarrlist.Contains(columnName) Then
                    destTable.Columns.RemoveAt(index)
                End If
            Next
            destTable.Merge(DtSet.Tables(0), False, MissingSchemaAction.Ignore)
            dsImportSplit.Tables.Add(destTable)
            dsImportSplit.Tables(1).TableName = "ImportItemSpecs"

            destTable = Nothing
            destTable = DtSet.Tables(0).Clone()
            For index As Integer = destTable.Columns.Count - 1 To 0 Step -1
                Dim columnName As String = destTable.Columns(index).ColumnName
                If Not itemassortmentarrlist.Contains(columnName) Then
                    destTable.Columns.RemoveAt(index)
                End If
            Next
            destTable.Merge(DtSet.Tables(0), False, MissingSchemaAction.Ignore)
            dsImportSplit.Tables.Add(destTable)
            dsImportSplit.Tables(2).TableName = "ImportItemAssortment"

            destTable = Nothing
            destTable = DtSet.Tables(0).Clone()
            For index As Integer = destTable.Columns.Count - 1 To 0 Step -1
                Dim columnName As String = destTable.Columns(index).ColumnName
                If Not itemmaterialarrlist.Contains(columnName) Then
                    destTable.Columns.RemoveAt(index)
                End If
            Next
            destTable.Merge(DtSet.Tables(0), False, MissingSchemaAction.Ignore)
            dsImportSplit.Tables.Add(destTable)
            dsImportSplit.Tables(3).TableName = "ImportItemMaterial"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub AddMissingColumns(ByRef dtImportSplitItem As DataTable, ByRef dtTableStructureItem As DataTable)
        ' Loop thru the imported table that has been split out and add any missing columns.
        ' . then we can display them.


    End Sub

    Public Sub compareitemchanges(ByRef DtSet As DataSet, ByRef dsImport As DataSet, ByRef dsChanges As DataSet, ByRef dsimportitems As DataSet)
        ' compare the values in dtset item table to the item table in the dsimport. write out any changes to the dschanges.
        ' . then we can display them.
        'Dim rows() As DataRow = DataTable.Select("ColumnName1 = 'value3'")
        'DTSet is the imported spreadsheet
        'dsimport the fields on the excel spreadsheet split on the different tables
        ' dschanges the changes between SQL and Excle
        ' dsimporteditems is the 4 tables from SQL based on the proposalnumber and rev
        Dim importproposalnumber As Long
        Dim importrev As Long
        Dim looprows As DataRow, itemrows As DataRow
        Dim loopcolumns As DataColumn
        ' Dim columnname As String, mycolumnname As String
        Dim sErrorcolumnname As String = ""
        Dim columnnumber As String
        Dim looprowsnumber As Long
        Dim foundRows() As Data.DataRow
        Try
            '  CheckDataType(DtSet.Tables(0))
            ' CheckDataType(dsimportitems.Tables("Item"))
            For Each looprows In DtSet.Tables(0).Rows

                looprowsnumber = CLng(looprows.Item("RowNumber"))
                importproposalnumber = CLng(looprows.Item("ProposalNumber"))
                importrev = CInt(looprows.Item("Rev"))
                itemrows = dsimportitems.Tables("Item").Select("ProposalNumber = " & importproposalnumber & " and Rev = " & importrev).FirstOrDefault()
                'If String.IsNullOrEmpty(looprows.Item("FirstRowFormatted")) = False Then
                '    MsgBox("the formatiing is colored")
                'End If
                'If String.IsNullOrEmpty(looprows.Item("SecondRowFormatted").ToString) = False Then
                '    MsgBox("the formatiing is colored")
                'End If
                'If String.IsNullOrEmpty(looprows.Item("thirdRowFormatted")) = False Then
                '    MsgBox("the formatiing is colored")
                'End If
                If String.IsNullOrEmpty(looprows.Item("FirstRowFormatted").ToString) = True And String.IsNullOrEmpty(looprows.Item("SecondRowFormatted").ToString) = True And String.IsNullOrEmpty(looprows.Item("ThirdRowFormatted").ToString) = True Then
                    ' find and compare entire row first.
                    foundRows = dsimportitems.Tables("Item").Select("ProposalNumber = '" & importproposalnumber & "'")
                    If Not looprows.Equals(foundRows) Then
                        For Each loopcolumns In DtSet.Tables(0).Columns

                            sErrorcolumnname = loopcolumns.ColumnName

                            'columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                            If dsimportitems.Tables("Item").Columns.Contains(loopcolumns.ColumnName) Then
                                If dsChanges.Tables(1).Columns.Contains(loopcolumns.ColumnName) Then
                                    If dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "String" Then
                                        If Not (IsDBNull(DtSet.Tables(0).Rows(looprowsnumber - 1).Item(sErrorcolumnname))) = False Then
                                            ' If Not (IsDBNull(dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname))) = False Then
                                            columnnumber = loopcolumns.Ordinal + 1
                                            looprows.Item(loopcolumns.ColumnName) = ""
                                        Else
                                            ' columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                                            columnnumber = loopcolumns.Ordinal + 1
                                        End If
                                    Else
                                        'columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                                        columnnumber = loopcolumns.Ordinal + 1
                                    End If
                                Else
                                    columnnumber = "0"
                                End If
                                If String.Equals(looprows.Item(loopcolumns.ColumnName).ToString, itemrows.Item(loopcolumns.ColumnName).ToString) = False Then
                                    'If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                                    'convert values to the same type boolean to 1-0 or yes no and all to three decimals.
                                    If dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Byte" Then
                                        If looprows.Item(loopcolumns.ColumnName).ToString = "YES" Then
                                            looprows.Item(loopcolumns.ColumnName) = "1"
                                        Else
                                            If looprows.Item(loopcolumns.ColumnName).ToString = "NO" Then
                                                looprows.Item(loopcolumns.ColumnName) = "0"
                                            Else
                                                looprows.Item(loopcolumns.ColumnName) = DBNull.Value
                                            End If
                                        End If
                                    End If
                                    If dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal" Then
                                        Try
                                            looprows.Item(loopcolumns.ColumnName) = System.Convert.ToDecimal(looprows.Item(loopcolumns.ColumnName).ToString)
                                        Catch exception As System.OverflowException
                                            MsgBox("Overflow in string-to-decimal conversion.")
                                        Catch exception As System.FormatException
                                            ' MsgBox("The string is not formatted as a decimal.")
                                            If Len(looprows.Item(loopcolumns.ColumnName).ToString) = 0 Then
                                                looprows.Item(loopcolumns.ColumnName) = 0
                                            End If
                                        Catch exception As System.ArgumentException
                                            MsgBox("The string is null.")
                                        End Try
                                    End If
                                    If loopcolumns.ColumnName = "Class" Then
                                        looprows.Item(loopcolumns.ColumnName) = looprows.Item(loopcolumns.ColumnName).ToString.PadLeft(2, CChar("0"))
                                    End If
                                    ' MsgBox("Here are the values: " & looprows.Item(loopcolumns.ColumnName).ToString & " and " & itemrows.Item(loopcolumns.ColumnName).ToString & "  " & dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString)
                                    Select Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString
                                        Case "Decimal"
                                            If Decimal.Equals(looprows.Item(loopcolumns.ColumnName), itemrows.Item(loopcolumns.ColumnName)) = False Then
                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "Item", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                            End If
                                            ' Case "Integer"
                                        Case "String"
                                            If String.Equals(Trim(looprows.Item(loopcolumns.ColumnName).ToString), Trim(itemrows.Item(loopcolumns.ColumnName).ToString)) = False Then
                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "Item", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                            End If
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                        Case Else
                                            If String.Equals(Trim(looprows.Item(loopcolumns.ColumnName).ToString), Trim(itemrows.Item(loopcolumns.ColumnName).ToString)) = False Then
                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "Item", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                            End If
                                    End Select
                                    'If String.Equals(looprows.Item(loopcolumns.ColumnName).ToString, itemrows.Item(loopcolumns.ColumnName).ToString) = False Then
                                    '    'If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                                    '    'MsgBox("Here are the values: " & looprows.Item(loopcolumns.ColumnName).ToString & " and " & itemrows.Item(loopcolumns.ColumnName).ToString)
                                    '    'If dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal" Then
                                    '    '    'MsgBox("helpme")
                                    '    '    If IsDBNull(looprows.Item(loopcolumns.ColumnName)) = True Then
                                    '    '        itemrows.Item(loopcolumns.ColumnName) = 0
                                    '    '        System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName)).ToString()
                                    '    '    Else
                                    '    '        'looprows.Item(loopcolumns.ColumnName) = 0
                                    '    '        ' System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName))
                                    '    '    End If
                                    '    'Else
                                    '    '    'End If
                                    '    '    'If itemrows.Item(loopcolumns.ColumnName).ToString = "" Or String.IsNullOrEmpty(CStr(looprows.Item(loopcolumns.ColumnName).ToString)) = True Then
                                    '    '    '    dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "Item", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                    '    '    'Else

                                    '    '    '    'If System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName)) = looprows.Item(loopcolumns.ColumnName) Then
                                    '    '    '    'Else
                                    '    '    '    '    ' MsgBox(itemrows.Item(loopcolumns.ColumnName).dataType.ToString)
                                    '    '    '    '    ' MsgBox(dsimportitems.Tables("Item").Columns("Lighted").DataType.Name.ToString)
                                    '    '    '    '    ' End If

                                    '    '    '    dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "Item", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                    '    '    '    'End If
                                    '    '    'End If
                                    '    'End If
                                    'End If
                                End If
                            End If

                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & "  " & sErrorcolumnname & "   " & looprowsnumber & "   " & importproposalnumber)
        End Try
    End Sub

    Public Sub compareitemspecchanges(ByRef DtSet As DataSet, ByRef dsImport As DataSet, ByRef dsChanges As DataSet, ByRef dsimportitems As DataSet)
        ' compare the values in dtset item table to the item table in the dsimport. write out any changes to the dschanges.
        ' . then we can display them.
        'Dim rows() As DataRow = DataTable.Select("ColumnName1 = 'value3'")
        Dim importproposalnumber As Long
        Dim importrev As Long
        Dim looprows As DataRow, itemrows As DataRow
        Dim loopcolumns As DataColumn
        ' Dim columnname As String, mycolumnname As String
        Dim sErrorcolumnname As String = ""
        Dim columnnumber As String = ""
        Dim looprowsnumber As Long
        Dim foundRows() As Data.DataRow
        Try
            '  CheckDataType(DtSet.Tables(0))
            '  CheckDataType(dsimportitems.Tables("ItemSpecs"))
            For Each looprows In DtSet.Tables(0).Rows
                looprowsnumber = CLng(looprows.Item("RowNumber"))
                importproposalnumber = CLng(looprows.Item("ProposalNumber"))
                importrev = CInt(looprows.Item("Rev"))
                itemrows = dsimportitems.Tables("ItemSpecs").Select("ProposalNumber = " & importproposalnumber & " and Rev = " & importrev).FirstOrDefault()
                If String.IsNullOrEmpty(looprows.Item("FirstRowFormatted").ToString) = True And String.IsNullOrEmpty(looprows.Item("SecondRowFormatted").ToString) = True And String.IsNullOrEmpty(looprows.Item("ThirdRowFormatted").ToString) = True Then

                    ' find and compare entire row first.
                    foundRows = dsimportitems.Tables("ItemSpecs").Select("ProposalNumber = '" & importproposalnumber & "'")
                    If Not looprows.Equals(itemrows) Then
                        For Each loopcolumns In DtSet.Tables(0).Columns
                            sErrorcolumnname = loopcolumns.ColumnName

                            'columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                            If dsimportitems.Tables("ItemSpecs").Columns.Contains(loopcolumns.ColumnName) Then
                                If dsChanges.Tables(1).Columns.Contains(loopcolumns.ColumnName) Then
                                    If dsimportitems.Tables("ItemSpecs").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "String" Then
                                        If Not (IsDBNull(DtSet.Tables(0).Rows(looprowsnumber - 1).Item(sErrorcolumnname))) = False Then
                                            columnnumber = loopcolumns.Ordinal + 1
                                            looprows.Item(loopcolumns.ColumnName) = ""
                                        Else
                                            ' columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                                            columnnumber = loopcolumns.Ordinal + 1
                                        End If
                                    Else
                                        'columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                                        columnnumber = loopcolumns.Ordinal + 1
                                    End If

                                Else
                                    columnnumber = "0"
                                End If
                                If String.Equals(looprows.Item(loopcolumns.ColumnName).ToString, itemrows.Item(loopcolumns.ColumnName).ToString) = False Then
                                    ' If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                                    'If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                                    ''sErrorcolumnname = loopcolumns.ColumnName
                                    ''If dsChanges.Tables(1).Columns.Contains(loopcolumns.ColumnName) Then
                                    ''    columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                                    ''Else
                                    ''    columnnumber = "0"
                                    ''End If
                                    'sErrorcolumnname = loopcolumns.ColumnName

                                    ''columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                                    ''If dsimportitems.Tables("ItemSpecs").Columns.Contains(loopcolumns.ColumnName) Then




                                    'If dsimportitems.Tables("ItemSpecs").Columns.Contains(loopcolumns.ColumnName) Then
                                    '    If dsChanges.Tables(1).Columns.Contains(loopcolumns.ColumnName) Then
                                    '        If dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "String" Then
                                    '            If Not (IsDBNull(dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname))) = False Then
                                    '                columnnumber = loopcolumns.Ordinal + 1
                                    '                looprows.Item(loopcolumns.ColumnName) = ""
                                    '            Else
                                    '                ' columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                                    '                columnnumber = loopcolumns.Ordinal + 1
                                    '            End If
                                    '        Else
                                    '            'columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                                    '            columnnumber = loopcolumns.Ordinal + 1
                                    '        End If

                                    '    Else
                                    '        columnnumber = "0"
                                    '    End If
                                    '    If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                                    'convert values to the same type boolean to 1-0 or yes no and all to three decimals.
                                    If dsimportitems.Tables("ItemSpecs").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Byte" Then
                                        If looprows.Item(loopcolumns.ColumnName).ToString = "YES" Then
                                            looprows.Item(loopcolumns.ColumnName) = "1"
                                        Else
                                            If looprows.Item(loopcolumns.ColumnName).ToString = "NO" Then
                                                looprows.Item(loopcolumns.ColumnName) = "0"
                                            Else
                                                looprows.Item(loopcolumns.ColumnName) = DBNull.Value
                                            End If
                                        End If
                                    End If
                                    If dsimportitems.Tables("ItemSpecs").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal" Then
                                        Try
                                            looprows.Item(loopcolumns.ColumnName) = System.Convert.ToDecimal(looprows.Item(loopcolumns.ColumnName).ToString)
                                        Catch exception As System.OverflowException
                                            MsgBox("Overflow in string-to-decimal conversion.")
                                        Catch exception As System.FormatException
                                            ' MsgBox("The string is not formatted as a decimal.")
                                            If Len(looprows.Item(loopcolumns.ColumnName).ToString) = 0 Then
                                                looprows.Item(loopcolumns.ColumnName) = 0
                                            End If
                                        Catch exception As System.ArgumentException
                                            MsgBox("The string is null.")
                                        End Try
                                    End If
                                    Select Case dsimportitems.Tables("ItemSpecs").Columns(loopcolumns.ColumnName).DataType.Name.ToString
                                        Case "Decimal"
                                            If Decimal.Equals(looprows.Item(loopcolumns.ColumnName), itemrows.Item(loopcolumns.ColumnName)) = False Then
                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemSpecs", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                            End If
                                            ' Case "Integer"
                                        Case "String"
                                            If String.Equals(Trim(looprows.Item(loopcolumns.ColumnName).ToString), Trim(itemrows.Item(loopcolumns.ColumnName).ToString)) = False Then
                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemSpecs", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                            End If
                                        Case "Byte"
                                            If Byte.Equals(looprows.Item(loopcolumns.ColumnName).ToString, itemrows.Item(loopcolumns.ColumnName).ToString) = False Then
                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemSpecs", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                            End If
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                            'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                        Case Else
                                            If String.Equals(Trim(looprows.Item(loopcolumns.ColumnName).ToString), Trim(itemrows.Item(loopcolumns.ColumnName).ToString)) = False Then
                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemSpecs", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                            End If
                                    End Select
                                    '        --If String.Equals(looprows.Item(loopcolumns.ColumnName).ToString, itemrows.Item(loopcolumns.ColumnName).ToString) = False Then
                                    '        '   If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                                    '        If dsimportitems.Tables("ItemSpecs").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal" Then
                                    '            'MsgBox("helpme")
                                    '            If IsDBNull(looprows.Item(loopcolumns.ColumnName)) = True Then
                                    '                itemrows.Item(loopcolumns.ColumnName) = 0
                                    '                System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName)).ToString()
                                    '            Else
                                    '                'looprows.Item(loopcolumns.ColumnName) = 0
                                    '                ' System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName))
                                    '            End If
                                    '        Else
                                    '            'End If
                                    '            If itemrows.Item(loopcolumns.ColumnName).ToString = "" Or String.IsNullOrEmpty(CStr(looprows.Item(loopcolumns.ColumnName))) = True Then
                                    '                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemSpecs", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                    '            Else

                                    '                'If System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName)) = looprows.Item(loopcolumns.ColumnName) Then
                                    '                'Else
                                    '                '    ' MsgBox(itemrows.Item(loopcolumns.ColumnName).dataType.ToString)
                                    '                '    ' MsgBox(dsimportitems.Tables("Item").Columns("Lighted").DataType.Name.ToString)
                                    '                '    ' End If

                                    '                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemSpecs", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                    '                'End If
                                    '            End If
                                    '        End If
                                    '    End If
                                    'Select Case dsimportitems.Tables("ItemSpecs").Columns(loopcolumns.ColumnName).DataType.Name.ToString
                                    '    Case "Decimal"
                                    '        If Decimal.Equals(looprows.Item(loopcolumns.ColumnName), itemrows.Item(loopcolumns.ColumnName)) = False Then
                                    '            dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "Item", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                    '        End If
                                    '        ' Case "Integer"
                                    '    Case "String"
                                    '        If String.Equals(Trim(looprows.Item(loopcolumns.ColumnName).ToString), Trim(itemrows.Item(loopcolumns.ColumnName).ToString)) = False Then
                                    '            dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "Item", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                    '        End If
                                    '        'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                    '        'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                    '        'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                    '        'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                    '        'Case dsimportitems.Tables("Item").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal"
                                    '    Case Else
                                    '        If String.Equals(Trim(looprows.Item(loopcolumns.ColumnName).ToString), Trim(itemrows.Item(loopcolumns.ColumnName).ToString)) = False Then
                                    '            dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemSpecs", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                                    '        End If
                                    'End Select
                                End If
                            End If

                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & "  " & sErrorcolumnname)


        End Try


        'dsChanges.Tables(0).Columns.Add("RowNumber", System.Type.GetType("System.Int32"))
        'dsChanges.Tables(0).Columns.Add("TableName", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("ColumnName", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("OldValue", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("NewValue", System.Type.GetType("System.String"))
    End Sub

    Public Sub compareitemmaterialchanges(ByRef DtSet As DataSet, ByRef dsImport As DataSet, ByRef dsChanges As DataSet, ByRef dsimportitems As DataSet)
        ' compare the values in dtset item table to the item table in the dsimport. write out any changes to the dschanges.
        ' . then we can display them.
        'Dim rows() As DataRow = DataTable.Select("ColumnName1 = 'value3'")
        Dim importproposalnumber As Long
        Dim importrev As Long
        Dim looprows As DataRow, itemrows As DataRow
        Dim loopcolumns As DataColumn
        ' Dim columnname As String, mycolumnname As String
        Dim sErrorcolumnname As String = ""
        Dim columnnumber As String
        Dim looprowsnumber As Long
        ' Dim foundRows() As Data.DataRow
        Try
            ' CheckDataType(DtSet.Tables(0))
            ' CheckDataType(dsimportitems.Tables("ItemMaterial"))
            For Each looprows In DtSet.Tables(0).Rows
                looprowsnumber = CLng(looprows.Item("RowNumber"))
                importproposalnumber = CLng(looprows.Item("ProposalNumber"))
                importrev = CInt(looprows.Item("Rev"))
                itemrows = dsimportitems.Tables("ItemMaterial").Select("ProposalNumber = " & importproposalnumber & " and Rev = " & importrev).FirstOrDefault()
                '
                ' find and compare entire row first.
                'foundRows = dsimportitems.Tables("ItemSpecs").Select("ProposalNumber = '" & importproposalnumber & "'")
                If Not looprows.Equals(itemrows) Then
                    For Each loopcolumns In DtSet.Tables(0).Columns
                        sErrorcolumnname = loopcolumns.ColumnName

                        'columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                        'If dsimportitems.Tables("ItemMaterial").Columns.Contains(loopcolumns.ColumnName) Then
                        '    If dsChanges.Tables(1).Columns.Contains(loopcolumns.ColumnName) Then
                        '        If dsimportitems.Tables("ItemMaterial").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "String" Then
                        '            If Not (IsDBNull(DtSet.Tables(0).Rows(looprowsnumber - 1).Item(sErrorcolumnname))) = False Then
                        '                ' If Not (IsDBNull(dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname))) = False Then
                        '                columnnumber = loopcolumns.Ordinal + 1
                        '                looprows.Item(loopcolumns.ColumnName) = ""
                        '            Else
                        '                ' columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                        '                columnnumber = loopcolumns.Ordinal + 1
                        '            End If
                        '        Else
                        '            'columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                        '            columnnumber = loopcolumns.Ordinal + 1
                        '        End If

                        '    Else
                        '        columnnumber = "0"
                        '    End If




                        'If dsimportitems.Tables("ItemMaterial").Columns.Contains(loopcolumns.ColumnName) Then
                        '    If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                        'convert values to the same type boolean to 1-0 or yes no and all to three decimals.
                        'If dsimportitems.Tables("ItemMaterial").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Byte" Then
                        '    If looprows.Item(loopcolumns.ColumnName).ToString = "YES" Then
                        '        looprows.Item(loopcolumns.ColumnName) = "1"
                        '    Else
                        '        looprows.Item(loopcolumns.ColumnName) = "0"
                        '    End If
                        'End If
                        'If dsimportitems.Tables("ItemMaterial").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal" Then
                        '    Try
                        '        looprows.Item(loopcolumns.ColumnName) = System.Convert.ToDecimal(looprows.Item(loopcolumns.ColumnName).ToString)
                        '    Catch exception As System.OverflowException
                        '        MsgBox("Overflow in string-to-decimal conversion.")
                        '    Catch exception As System.FormatException
                        '        ' MsgBox("The string is not formatted as a decimal.")
                        '        If Len(looprows.Item(loopcolumns.ColumnName).ToString) = 0 Then
                        '            looprows.Item(loopcolumns.ColumnName) = 0
                        '        End If
                        '    Catch exception As System.ArgumentException
                        '        MsgBox("The string is null.")
                        '    End Try
                        'End If
                        'If loopcolumns.ColumnName = "Class" Then
                        '    looprows.Item(loopcolumns.ColumnName) = looprows.Item(loopcolumns.ColumnName).ToString.PadLeft(2, CChar("0"))
                        'End If
                        '    If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                        '        If dsimportitems.Tables("ItemMaterial").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal" Then
                        '            'MsgBox("helpme")
                        '            If IsDBNull(looprows.Item(loopcolumns.ColumnName)) = True Then
                        '                itemrows.Item(loopcolumns.ColumnName) = 0
                        '                System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName)).ToString()
                        '            Else
                        '                'looprows.Item(loopcolumns.ColumnName) = 0
                        '                ' System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName))
                        '            End If
                        '        Else
                        '            'End If
                        '            If itemrows.Item(loopcolumns.ColumnName).ToString = "" Or String.IsNullOrEmpty(CStr(looprows.Item(loopcolumns.ColumnName))) = True Then
                        '                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemMaterial", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                        '            Else

                        '                'If System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName)) = looprows.Item(loopcolumns.ColumnName) Then
                        '                'Else
                        '                '    ' MsgBox(itemrows.Item(loopcolumns.ColumnName).dataType.ToString)
                        '                '    ' MsgBox(dsimportitems.Tables("Item").Columns("Lighted").DataType.Name.ToString)
                        '                '    ' End If

                        '                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemMaterial", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber, importproposalnumber, importrev)
                        '                'End If
                        '            End If
                        '        End If
                        '    End If
                        'End If
                        'End If
                        '    End If

                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & "  " & sErrorcolumnname)


        End Try


        'dsChanges.Tables(0).Columns.Add("RowNumber", System.Type.GetType("System.Int32"))
        'dsChanges.Tables(0).Columns.Add("TableName", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("ColumnName", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("OldValue", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("NewValue", System.Type.GetType("System.String"))
    End Sub

    Public Sub compareitemAssortmantchanges(ByRef DtSet As DataSet, ByRef dsImport As DataSet, ByRef dsChanges As DataSet, ByRef dsimportitems As DataSet)
        ' compare the values in dtset item table to the item table in the dsimport. write out any changes to the dschanges.
        ' . then we can display them.
        'Dim rows() As DataRow = DataTable.Select("ColumnName1 = 'value3'")
        Dim importproposalnumber As Long
        Dim importrev As Long
        Dim looprows As DataRow, itemrows As DataRow
        Dim loopcolumns As DataColumn
        ' Dim columnname As String, mycolumnname As String
        Dim sErrorcolumnname As String = ""
        Dim columnnumber As String = ""
        Try
            ' CheckDataType(DtSet.Tables(0))
            ' CheckDataType(dsimportitems.Tables("ItemAssortments"))
            For Each looprows In DtSet.Tables(0).Rows
                If String.IsNullOrEmpty(looprows.Item("FirstRowFormatted").ToString) = True And String.IsNullOrEmpty(looprows.Item("SecondRowFormatted").ToString) = True And String.IsNullOrEmpty(looprows.Item("ThirdRowFormatted").ToString) = True Then

                    importproposalnumber = CLng(looprows.Item("ProposalNumber"))
                    importrev = CInt(looprows.Item("Rev"))
                    itemrows = dsimportitems.Tables("ItemAssortments").Select("ProposalNumber = " & importproposalnumber & " and Rev = " & importrev).FirstOrDefault()
                    For Each loopcolumns In DtSet.Tables(0).Columns
                        sErrorcolumnname = loopcolumns.ColumnName
                        If dsChanges.Tables(1).Columns.Contains(loopcolumns.ColumnName) Then
                            columnnumber = dsChanges.Tables(1).Rows(0).Item(sErrorcolumnname).ToString
                        Else
                            columnnumber = "0"
                        End If
                        If dsimportitems.Tables("ItemAssortments").Columns.Contains(loopcolumns.ColumnName) Then
                            If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                                'convert values to the same type boolean to 1-0 or yes no and all to three decimals.
                                If dsimportitems.Tables("ItemAssortments").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Byte" Then
                                    If looprows.Item(loopcolumns.ColumnName).ToString = "YES" Then
                                        looprows.Item(loopcolumns.ColumnName) = "1"
                                    Else
                                        looprows.Item(loopcolumns.ColumnName) = "0"
                                    End If
                                End If
                                If dsimportitems.Tables("ItemAssortments").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal" Then
                                    Try
                                        looprows.Item(loopcolumns.ColumnName) = System.Convert.ToDecimal(looprows.Item(loopcolumns.ColumnName).ToString)
                                    Catch exception As System.OverflowException
                                        MsgBox("Overflow in string-to-decimal conversion.")
                                    Catch exception As System.FormatException
                                        ' MsgBox("The string is not formatted as a decimal.")
                                        If Len(looprows.Item(loopcolumns.ColumnName).ToString) = 0 Then
                                            looprows.Item(loopcolumns.ColumnName) = 0
                                        End If
                                    Catch exception As System.ArgumentException
                                        MsgBox("The string is null.")
                                    End Try
                                End If
                                If loopcolumns.ColumnName = "Class" Then
                                    looprows.Item(loopcolumns.ColumnName) = looprows.Item(loopcolumns.ColumnName).ToString.PadLeft(2, CChar("0"))
                                    If looprows.Item(loopcolumns.ColumnName).ToString <> itemrows.Item(loopcolumns.ColumnName).ToString Then
                                        If dsimportitems.Tables("ItemAssortments").Columns(loopcolumns.ColumnName).DataType.Name.ToString = "Decimal" Then
                                            'MsgBox("helpme")
                                            If IsDBNull(looprows.Item(loopcolumns.ColumnName)) = True Then
                                                itemrows.Item(loopcolumns.ColumnName) = 0
                                                System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName)).ToString()
                                            Else
                                                'looprows.Item(loopcolumns.ColumnName) = 0
                                                ' System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName))
                                            End If
                                        Else
                                            'End If
                                            If itemrows.Item(loopcolumns.ColumnName).ToString = "" Or String.IsNullOrEmpty(CStr(looprows.Item(loopcolumns.ColumnName))) = True Then
                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "Item", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber)
                                            Else

                                                'If System.Convert.ToDecimal(itemrows.Item(loopcolumns.ColumnName)) = looprows.Item(loopcolumns.ColumnName) Then
                                                'Else
                                                '    ' MsgBox(itemrows.Item(loopcolumns.ColumnName).dataType.ToString)
                                                '    ' MsgBox(dsimportitems.Tables("Item").Columns("Lighted").DataType.Name.ToString)
                                                '    ' End If

                                                dsChanges.Tables(0).Rows.Add(looprows.Item("RowNumber").ToString, "ItemAssortments", loopcolumns.ColumnName, itemrows.Item(loopcolumns.ColumnName), looprows.Item(loopcolumns.ColumnName), columnnumber)
                                                'End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & "  " & sErrorcolumnname)


        End Try


        'dsChanges.Tables(0).Columns.Add("RowNumber", System.Type.GetType("System.Int32"))
        'dsChanges.Tables(0).Columns.Add("TableName", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("ColumnName", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("OldValue", System.Type.GetType("System.String"))
        'dsChanges.Tables(0).Columns.Add("NewValue", System.Type.GetType("System.String"))
    End Sub

    Private Function CheckDataType(ByRef incomingdatatype As String) As String
        ' find out what the datatype is and change it to the correctly matching datatypes
        ' Display the SqlType of each column. 
        ' Dim column As DataColumn
        ' Console.WriteLine("Data Types:")
        Dim returnvalue As String
        Try


            'For Each column In dt.Columns
            '    Console.WriteLine(" {0} -- {1}", column.ColumnName, column.DataType.UnderlyingSystemType)
            '    ToolStripStatusLabel1.Text = column.ColumnName.ToString & "  " & column.DataType.UnderlyingSystemType.ToString
            ' ToolStripStatusLabel1.Paint()
            Application.DoEvents()
            Select Case incomingdatatype
                Case "int"
                    returnvalue = "sytem.int32"
                Case "string"
                    returnvalue = "sytem.string"
                Case "DateTime"


                    'Case column.DataType.UnderlyingSystemType = "System.String"
                    'Case sytem.Double
                    ' Case System.Boolean
                    '  BooleanValue(Value)
                    'Case sytem.Decimal

                    'Case sytem.int32
                    'Case sytem.int16
                    'Case sytem.byte
                    'Case sytem.decimal
                    'Case sytem.boolean
                    'Case sytem.string

                Case Else



            End Select
            CheckDataType = returnvalue
            ' Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Private Sub CheckDataTypedatatable(ByRef dt As DataTable, ByRef column As Integer, ByRef dsDt As DataTable, ByRef strDatatype As String)
        ' find out what the datatype is and change it to the correctly matching datatypes
        ' Display the SqlType of each column. 
        'Dim column As DataColumn
        ' Console.WriteLine("Data Types:")

        Try


            'For Each column In dt.Columns
            '    Console.WriteLine(" {0} -- {1}", column.ColumnName, column.DataType.UnderlyingSystemType)
            '    ToolStripStatusLabel1.Text = column.ColumnName.ToString & "  " & column.DataType.UnderlyingSystemType.ToString
            ' ToolStripStatusLabel1.Paint()
            Application.DoEvents()
            Select Case strDatatype
                Case "int"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Int32"))
                Case "smallint"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Int16"))
                Case "bigint"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.int64"))


                Case "decimal"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Decimal"))
                Case "money"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Decimal"))
                Case "numeric"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Decimal"))
                Case "smallmoney"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Decimal"))
                Case "float"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Double"))
                Case "real"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Single"))

                Case "time"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Time"))
                Case "date"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Date"))
                Case "datetime"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.DateTime"))
                Case "datetime2"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.DateTime2"))
                Case "smalldatetime"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.DateTime"))

                Case "char"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))
                Case "text"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))
                Case "ntext"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))
                Case "nvarchar"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))
                Case "nchar"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))
                Case "ntext"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))
                Case "string"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))
                Case "varchar"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))

                Case "bit"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Boolean"))
                Case "tinyint"
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.Byte"))

                Case Else
                    dt.Columns.Add(dsDt.Columns(column).ColumnName, System.Type.GetType("System.String"))
                    MsgBox("here is the CONVERTED datatype to string " & strDatatype & " and columnname " & dsDt.Columns(column).ColumnName)

            End Select

            ' Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '    public static bool ChangeColumnDataType(DataTable table,string columnname,Type newtype)
    '{
    '    if (table.Columns.Contains(columnname) == false)
    '        return false;

    '    DataColumn column= table.Columns[columnname];
    '    if (column.DataType == newtype)
    '        return true;

    '    try
    '    {
    '        DataColumn newcolumn = new DataColumn("temperary", newtype);
    '        table.Columns.Add(newcolumn);
    '        foreach (DataRow row in table.Rows)
    '        {
    '            try
    '            {
    '                row["temperary"] = Convert.ChangeType(row[columnname], newtype);
    '            }
    '            catch
    '            {
    '            }
    '        }
    '        table.Columns.Remove(columnname);
    '        newcolumn.ColumnName = columnname;
    '    }
    '    catch (Exception)
    '    {
    '        return false;
    '    }

    '    return true;
    '}
    'Private Function SQLDataType(ByRef column As DataColumn)
    '    Dim datatype As Object
    '    datatype = column.DataType

    '    Select Case datatype
    '        Case "System.Int32"
    '            Return column.DataType = System.Type.GetType("System.Int32") 'SqlDbType.Int
    '        Case "System.String"
    '            Return SqlDbType.VarChar
    '        Case "System.Long"
    '            Return SqlDbType.BigInt
    '        Case "System.Boolean"
    '            Return SqlDbType.Bit
    '        Case "System.Decimal"
    '            Return SqlDbType.Decimal
    '        Case "System.DateTime"
    '            Return SqlDbType.DateTime
    '        Case Else
    '            Return SqlDbType.VarChar
    '    End Select

    'End Function

    Protected Function checkNullOrTextValue(ByVal m As Object) As Boolean
        If Convert.IsDBNull(m) And Convert.ToString(m) <> "someValue" Then
            Return False
        Else
            Return True
        End If


    End Function
    Private Shared Function FormatNumber(ByVal number As Object, ByVal type As Type) As String
        If number Is Nothing Then Throw New ArgumentNullException("number")
        If type Is Nothing Then Throw New ArgumentNullException("type")
        If type.Equals(GetType(System.Data.SqlTypes.SqlInt32)) Then
            Return CStr(Integer.Parse(CStr(number)))
        ElseIf type.Equals(GetType(System.Data.SqlTypes.SqlDecimal)) Then
            Return Decimal.Parse(number.ToString()).ToString("N")
        ElseIf type.Equals(GetType(System.Data.SqlTypes.SqlMoney)) Then
            Return Decimal.Parse(number.ToString()).ToString("C")
        End If
        Throw New ArgumentOutOfRangeException(String.Format("Unknown type specified : " & type.ToString()))
    End Function

    'Public Sub compareitemSpecchanges(ByRef DtSet As DataSet, ByRef dsImport As DataSet, ByRef dsChanges As DataSet)
    '    ' compare the values in dtset item table to the item table in the dsimport. write out any changes to the dschanges.
    '    ' . then we can display them.


    'End Sub
    'Public Sub compareitemAssormentchanges(ByRef DtSet As DataSet, ByRef dsImport As DataSet, ByRef dsChanges As DataSet)
    '    ' compare the values in dtset item table to the item table in the dsimport. write out any changes to the dschanges.
    '    ' . then we can display them.


    'End Sub
    'Public Sub compareitemMaterialchanges(ByRef DtSet As DataSet, ByRef dsImport As DataSet, ByRef dsChanges As DataSet)
    '    ' compare the values in dtset item table to the item table in the dsimport. write out any changes to the dschanges.
    '    ' . then we can display them.


    'End Sub

    Public Sub CreateRelationships()
        '' Define the relationship between the tables.
        'Dim data_relation As New DataRelation("Item_Tables", dsImportSplit.Tables("ImportItem").Columns("ContactID"), m_DataSet.Tables("TestScores").Columns("ContactID"))
        'm_DataSet.Relations.Add(data_relation)


        ' how to reorder columns
        'datatable.Columns["Col1"].SetOrdinal(1);
    End Sub
    ' Converts any given value into either True or False.
    'Public Function BooleanValue(ByVal Value As Object) As Boolean
    '    Select Case Value
    '        Case "False", "No", "0", False, 0
    '            BooleanValue = False
    '        Case Else
    '            BooleanValue = True
    '    End Select
    'End Function

    ' Echos a value, converting Null to the emtpy string
    '' This can be a useful filter function for values passed to other functions that require string parameters
    'Public Function NotNull(ByVal Value)
    '    If IsNull(Value) Then
    '        NotNull = ""
    '    Else
    '        NotNull = Value
    '    End If
    'End Function
    '' Returns true if the Value is empty, Null, or missing, or the empty string ""
    'Public Function IsBlank(ByVal Value) As Boolean  'Gary
    '    If IsNull(Value) Then
    '        IsBlank = True
    '    ElseIf IsEmpty(Value) Then
    '        IsBlank = True
    '    ElseIf IsMissing(Value) Then
    '        IsBlank = True
    '    ElseIf Value = "" Then
    '        IsBlank = True
    '    ElseIf Value = " " Then
    '        IsBlank = True
    '    ElseIf VarType(Value) = vbString Then
    '        If Left(Value, 1) = " " Then
    '            If Value = Space(Len(Value)) Then
    '                IsBlank = True
    '            End If
    '        End If
    '    End If
    'End Function

    '    Public Function NotBlank(ByVal Value)
    '        On Error GoTo ErrorHandler

    '        If IsBlank(Value) Then
    '            NotBlank = Null
    '        Else
    '            NotBlank = Value
    '        End If

    '        GoTo TheEnd
    'ErrorHandler:
    '        MsgBox(Err.Description, vbCritical, Err.Number)
    'TheEnd:
    '    End Function

    Private Sub ReadExcelFile(ByRef objdt As DataTable, ByVal StrFilePath As String)
        Dim ExcelCon As New OleDbConnection
        Dim ExcelAdp As OleDbDataAdapter
        Dim ExcelComm As OleDbCommand
        ' Dim Col1 As DataColumn
        Try
            ExcelCon.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & StrFilePath & " ;Extended Properties=""Excel 14.0;HDR=YES;IMEX=1"""
            ' ExcelCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;"  "Data Source= " & StrFilePath  ";Extended Properties=""Excel 8.0;"""
            ExcelCon.Open()
            Dim strsql As String
            strsql = "Select * From [Sheet1$]"
            ExcelComm = New OleDbCommand(strsql, ExcelCon)
            ExcelAdp = New OleDbDataAdapter(ExcelComm)
            objdt = New DataTable()
            ExcelAdp.Fill(objdt)

            ''--- Create Column With SRNo.
            'Col1 = New DataColumn
            'Col1.DefaultValue = 0
            'Col1.DataType = System.Type.GetType("System.Decimal")
            'Col1.Caption = "Sr No."
            'Col1.ColumnName = "SrNo"
            'objdt.Columns.Add(Col1)
            'Col1.SetOrdinal(1)

            ExcelCon.Close()
        Catch ex As Exception

        Finally
            ExcelCon = Nothing
            ExcelAdp = Nothing
            ExcelComm = Nothing
        End Try
    End Sub

    Public Sub ExportExcel(ByRef strFileName As String, ByRef dschanges As DataSet)


        'Dim xlApp As New MsExcel.Application
        'Dim xlApp As MsExcel.Application = Nothing
        'Dim xlWorkBooks As MsExcel.Workbooks = Nothing
        'Dim xlWorkBook As MsExcel.Workbook = Nothing
        'Dim xlWorkSheet As MsExcel.Worksheet = Nothing
        'Dim xlWorkSheets As MsExcel.Sheets = Nothing
        'Dim xlCells As MsExcel.Range = Nothing
        'Dim xlRange As MsExcel.Range = Nothing

        ''If IO.File.Exists(txtFile.Text) Then



        'xlApp = New MsExcel.Application
        'xlApp.DisplayAlerts = False
        '' '' xlWorkBooks = xlApp.Workbooks
        ' '' xlWorkBook = xlWorkBooks.Open(txtFile.Text, [ReadOnly]:=False, Editable:=True)
        ''xlWorkBook = xlApp.Workbooks.Open(txtFile.Text, ReadOnly:=False, Editable:=True)


        'xlWorkBooks = xlApp.Workbooks
        'xlWorkBook = xlWorkBooks.Open(txtFile.Text)
        ' ''xlApp.Visible = True
        ' ''xlWorkSheets = xlWorkBook.Sheets


        ''If xlWorkBook.ReadOnly = True Then
        ''    MsgBox(" IT IS READ ONLY right afert opening the file")
        ''End If
        ''xlApp.Visible = False

        ''xlWorkSheets = xlWorkBook.Sheets
        ''xlWorkSheet = xlWorkBook.Sheets(1)
        ' ''xlWorkSheet = CType(xlWorkSheets(1), Excel.Worksheet)
        ''Dim sb As New System.Text.StringBuilder






        'Try
        '    xlWorkSheet = CType(xlWorkBook.ActiveSheet, MsExcel.Worksheet)
        '    ' oRange = CType(xlworkbook.ActiveSheet.UsedRange, Excel.Range)
        '    'txtFile.Text
        '    'Dim xlStyles As MsExcel.Styles = xlWorkBook.Styles
        '    ' Dim xlStyle As MsExcel.Style = Nothing
        '    'Dim isstyleexists As Boolean = False
        '    'For Each xlStyle In xlStyles
        '    '    If xlStyle.Name = "DotShadeStyle" Then
        '    '        isstyleexists = True
        '    '        ' xlStyle.Delete()
        '    '        Exit For
        '    '    End If
        '    'Next
        '    'If (Not isstyleexists) Then
        '    '    xlStyles.Add("DotShadeStyle")
        '    '    xlStyle.Interior.Pattern = XlFillPattern.xlGray8
        '    '    xlWorkBook.Save()
        '    'End If

        '    Dim changesSortedDV As DataView = New DataView(dschanges.Tables(0))
        '    changesSortedDV.Sort = "RowNumber ASC"
        '    Dim i As Integer

        '    Dim colletter As String
        '    Dim columnnumber As String
        '    Dim colheadername As String
        '    Dim colrow As String
        '    xlWorkBook.Application.Visible = True
        '    For i = 0 To changesSortedDV.Count - 1
        '        ' find the column header name.
        '        colheadername = changesSortedDV(i).Item(2).ToString
        '        ' find row and column and update the value
        '        columnnumber = changesSortedDV(i).Item(5).ToString
        '        'change the number into a letter 
        '        colletter = GetExcelColumnName(CInt(columnnumber))
        '        colrow = colletter & (CInt(changesSortedDV(i).Item(0)) + 1)
        '        ' get the row number from the dataview, this is Rownumber
        '        With xlWorkSheet
        '            .Range(colrow).Value = changesSortedDV(i).Item(3).ToString  '' update spreadsheet with value from SQL
        '            ' .Range(colrow).Style = "DotShadeStyle"
        '            ' .Range(colrow).Interior.Pattern = XlFillPattern.xlGray8
        '            ' .Range(colrow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
        '            xlWorkBook.Save()
        '        End With
        '    Next
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'Finally

        '    xlWorkBook.Save()
        '    xlWorkBook.Close()
        '    xlApp.Quit()
        '    xlApp = Nothing
        '    '' Release object references.
        '    'oRange = Nothing
        '    'xlWorkSheet = Nothing
        '    'xlworkbook = Nothing
        '    '' xlApp.Quit()
        '    '' xlApp = Nothing
        'End Try


    End Sub

    Public Sub ImportExcel(ByRef strFileName As String, ByRef dschanges As DataSet)

        Dim Xls As ExcelFile = New XlsFile(True)
        Xls.Open(strFileName)
        Dim format As TFlxFormat = Xls.GetDefaultFormat
        format.FillPattern = New TFlxFillPattern() With {.Pattern = TFlxPatternStyle.CrissCross, .FgColor = Color.LightGray} ''3
        'Xls.AddFormat(format)

        'Dim xlApp As New MsExcel.Application
        'Dim xlApp As MsExcel.Application = Nothing
        'Dim xlWorkBooks As MsExcel.Workbooks = Nothing
        'Dim xlWorkBook As MsExcel.Workbook = Nothing
        'Dim xlWorkSheet As MsExcel.Worksheet = Nothing
        'Dim xlWorkSheets As MsExcel.Sheets = Nothing
        'Dim xlCells As MsExcel.Range = Nothing
        'Dim xlRange As MsExcel.Range = Nothing

        'xlApp = New MsExcel.Application
        'xlApp.DisplayAlerts = False
        'xlWorkBooks = xlApp.Workbooks
        'xlWorkBook = xlWorkBooks.Open(txtFile.Text)

        ' Dim workbook As New Workbook()
        ' workbook.LoadFromFile(strFileName)
        '  Dim worksheet As Worksheet = workbook.Worksheets(0)


        '  Dim document As New Spreadsheet()
        ' document.LoadFromFile(strFileName)

        '' testing viewing of file.
        'Dim newfilename As String
        'newfilename = Replace(strFileName, ".xlsx", "1.xlsx")

        ' Get the worksheet named "Sheet1" from excel file
        'Dim worksheet As Worksheet = spreadsheet.Workbook.Worksheets.ByName("Sheet1")
        '  Dim worksheet As Worksheet = document.Workbook.Worksheets(0)
        ''cell.FillPatternBackColor.ToString())

        'For i As Integer = 0 To spreadsheet.Workbook.Worksheets.Count - 1
        ' Dim worksheet As Worksheet = spreadsheet.Workbook.Worksheets(0)

        'Dim table As DataTable = DataSet.Tables.Add(worksheet.Name)

        'Dim fmt As TFlxFormat = arguments.Xls.GetCellVisibleFormatDef(SourceCell.Sheet1, SourceCell.Top, SourceCell.Left)
        'Dim SourceColor As Integer = fmt.FillPattern.FgColor.ToColor(arguments.Xls).ToArgb()
        'Dim Clr As TUIColor = Color.Empty
        Dim clr As TUIColor = Color.LightGray
        'Dim xf As Integer = TFlxFormat.FillPattern
        Dim cformat As TUIHatchStyle = TUIHatchStyle.DottedGrid
        Try
            'xlWorkSheet = CType(xlWorkBook.ActiveSheet, MsExcel.Worksheet)

            'Dim xlStyles As MsExcel.Styles = xlWorkBook.Styles
            'Dim xlStyle As MsExcel.Style = Nothing
            ' Dim isstyleexists As Boolean = False
            Dim SqlCommand As SqlCommand = New SqlCommand
            Dim changesSortedDV As DataView = New DataView(dschanges.Tables(0))
            ' find out if there are any changes
            Dim rownumber As Integer
            If dschanges.Tables(0) IsNot Nothing AndAlso dschanges.Tables(0).Rows.Count > 0 Then

                changesSortedDV.Sort = "RowNumber ASC"
                Dim i As Integer

                Dim colletter As String
                Dim columnnumber As String
                Dim colheadername As String
                Dim colrow As String
                Dim firstcol As String
                Using sqlConn = (SQLConnection())
                    Using SqlCommand
                        'xlWorkBook.Application.Visible = True
                        For i = 0 To changesSortedDV.Count - 1
                            ' find the column header name.
                            colheadername = changesSortedDV(i).Item(2).ToString

                            rownumber = CInt(changesSortedDV(i).Item(0))

                            ' find row and column and update the value
                            columnnumber = changesSortedDV(i).Item(5).ToString
                            'change the number into a letter 
                            colletter = GetExcelColumnName(CInt(columnnumber))
                            colrow = colletter & (CInt(changesSortedDV(i).Item(0)) + 1)
                            firstcol = "A" & (CInt(changesSortedDV(i).Item(0)) + 1)
                            '   worksheet.Range(firstcol).Value = ""
                            Xls.SetCellValue(rownumber + 1, 1, "")
                            Xls.SetCellFormat(rownumber + 1, columnnumber, Xls.AddFormat(format))
                            Xls.SetCellValue(rownumber + 1, columnnumber, changesSortedDV(i).Item(4).ToString)

                            ' worksheet.Range(colrow).Style.KnownColor = ExcelColors.Gray25Percent
                            ' worksheet.Range(colletter & rownumber).Style.FillPattern = ExcelPatternType.LightDownwardDiagonal

                            Try
                                SqlCommand.CommandText = "UPDATE " & changesSortedDV(i).Item(1).ToString & "  SET " & changesSortedDV(i).Item(2).ToString & " =  '" & changesSortedDV(i).Item(4).ToString.Replace("'", "") & "'  WHERE Proposalnumber  = " & changesSortedDV(i).Item(6).ToString & " and Rev = " & changesSortedDV(i).Item(7).ToString
                                SqlCommand.Connection = sqlConn
                                sqlConn.Open()

                                'here execute scalar will get firsr row first column value
                                Dim retValue As Integer = SqlCommand.ExecuteNonQuery()
                                If retValue > 0 Then

                                Else
                                    ' MsgBox("No record(s) inserted!") '' write out the log file
                                End If
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            Finally
                                If (SqlCommand.Connection.State = ConnectionState.Open) Then
                                    SqlCommand.Connection.Close()
                                End If
                            End Try

                        Next
                    End Using
                End Using
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            'spreadsheet.SaveAs(txtFile.Text)
            ' spreadsheet.Close()
            ' spreadsheet.Dispose()
        Finally
            '  workbook.SaveToFile(strFileName)
            ' workbook.SaveToFile(newfilename, ExcelVersion.Version2010)
            'System.Diagnostics.Process.Start(workbook.FileName)
            Xls.Save(strFileName)
            Me.btnOpenExcel.Enabled = True


        End Try


    End Sub

    Public Sub ValidateExcel(ByRef strFileName As String, ByRef dschanges As DataSet, ByRef dtImportedExcel As DataTable)


        ''Dim xlApp As New MsExcel.Application
        'Dim xlApp As MsExcel.Application = Nothing
        'Dim xlWorkBooks As MsExcel.Workbooks = Nothing
        'Dim xlWorkBook As MsExcel.Workbook = Nothing
        'Dim xlWorkSheet As MsExcel.Worksheet = Nothing
        'Dim xlWorkSheets As MsExcel.Sheets = Nothing
        'Dim xlCells As MsExcel.Range = Nothing
        'Dim xlRange As MsExcel.Range = Nothing
        ''If IO.File.Exists(txtFile.Text) Then

        'xlApp = New MsExcel.Application
        'xlApp.DisplayAlerts = False
        'xlWorkBooks = xlApp.Workbooks
        'xlWorkBook = xlWorkBooks.Open(txtFile.Text)

        'Try
        '    xlWorkSheet = CType(xlWorkBook.ActiveSheet, MsExcel.Worksheet)
        '    ' oRange = CType(xlworkbook.ActiveSheet.UsedRange, Excel.Range)
        '    'txtFile.Text
        '    Dim xlStyles As MsExcel.Styles = xlWorkBook.Styles
        '    Dim xlStyle As MsExcel.Style = Nothing
        '    Dim isstyleexists As Boolean = False
        '    Dim changesSortedDV As DataView = New DataView(dschanges.Tables(1))
        '    ' changesSortedDV.Sort = "RowNumber ASC"
        '    Dim i As Integer
        '    Dim colletter As String
        '    Dim columnnumber As String
        '    Dim colheadername As String
        '    Dim colrow As String

        '    xlWorkBook.Application.Visible = True
        '    For i = 0 To changesSortedDV.Count - 1
        '        ' find the column header name.
        '        colheadername = changesSortedDV(i).Item(2).ToString
        '        ' find row and column and update the value
        '        columnnumber = changesSortedDV(i).Item(5).ToString
        '        'change the number into a letter 
        '        colletter = GetExcelColumnName(CInt(columnnumber))
        '        colrow = colletter & (CInt(changesSortedDV(i).Item(0)) + 1)
        '        ' get the row number from the dataview, this is Rownumber
        '        With xlWorkSheet
        '            .Range(colrow).Value = changesSortedDV(i).Item(3).ToString  '' update spreadsheet with value from SQL
        '            ' .Range(colrow).Style = "DotShadeStyle"
        '            .Range(colrow).Interior.Pattern = XlFillPattern.xlGray8
        '            ' .Range(colrow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
        '            xlWorkBook.Save()
        '        End With
        '    Next
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'Finally

        '    xlWorkBook.Save()
        '    xlWorkBook.Close()
        '    xlApp.Quit()
        '    xlApp = Nothing
        '    '' Release object references.
        '    'oRange = Nothing
        '    'xlWorkSheet = Nothing
        '    'xlworkbook = Nothing
        '    '' xlApp.Quit()
        '    '' xlApp = Nothing
        'End Try


    End Sub

    '' Function GetExcelColumn - returns the column reference 
    '' from an integer representing a column in a datagrid or dataset
    'Function GetExcelColumn(ByVal col As Integer) As String
    '    Dim result As String
    '    Select Case col
    '        Case 0 ' first column
    '            result = "A"
    '        Case 1
    '            result = "B"
    '        Case 2
    '            result = "C"
    '        Case 3
    '            result = "D"
    '        Case 4
    '            result = "E"
    '        Case 5
    '            result = "F"
    '        Case 6
    '            result = "G"
    '        Case 7
    '            result = "H"
    '        Case 8
    '            result = "I"
    '        Case 9
    '            result = "J"
    '        Case 10
    '            result = "K"
    '        Case 11
    '            result = "L"
    '        Case 12
    '            result = "M"
    '        Case 13
    '            result = "N"
    '        Case 14
    '            result = "O"
    '        Case 15
    '            result = "P"
    '        Case 16
    '            result = "Q"
    '        Case 17
    '            result = "R"
    '        Case 18
    '            result = "S"
    '        Case 19
    '            result = "T"
    '        Case 20
    '            result = "U"
    '        Case 21
    '            result = "V"
    '        Case 22
    '            result = "W"
    '        Case 23
    '            result = "X"
    '        Case 24
    '            result = "Y"
    '        Case 25
    '            result = "Z"
    '    End Select
    '    Return result
    'End Function


    Private Function GetExcelColumnName(columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = [String].Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) \ 26)
        End While

        Return columnName
    End Function


    Private Sub releaseObject(ByVal obj As Object)
        Try
            Dim intRel As Integer = 0
            Do
                intRel = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            Loop While intRel > 0
            '  MsgBox("Final Released obj # " & intRel)
        Catch ex As Exception
            '  MsgBox("Error releasing object" & ex.ToString)
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Private Sub ToolStripStatusLabel2_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripStatusLabel2.ButtonClick
        MessageBox.Show("When an Excel image being displayed Excel is in memory, A green dot image, Excel is not in memory.")
    End Sub



    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'If ExcelInMemory() Then
        '    Me.ToolStripStatusLabel2.Image = Image.FromFile("C:\Users\rstruck\Documents\Visual Studio 2012\Projects\Import-Export-Excel\Import-Export-Excel\Excel1.png")
        'Else
        '    Me.ToolStripStatusLabel2.Image = Image.FromFile("C:\Users\rstruck\Documents\Visual Studio 2012\Projects\Import-Export-Excel\Import-Export-Excel\Excel2.png")
        'End If

        Me.ToolStripStatusLabel2.Invalidate()
        Application.DoEvents()

        'If IO.File.Exists(ExcelFileName) Then
        '    ' cmdOpenFile.Enabled = True
        '    ' cmdGetCellValue.Enabled = True
        'Else
        '    ' cmdOpenFile.Enabled = False
        '    ' cmdGetCellValue.Enabled = False
        'End If
    End Sub

    Private Sub btnValidate_Click(sender As Object, e As EventArgs) Handles btnValidate.Click
        ProcessExcel(3)
    End Sub

    Private Sub btnOpenExcel_Click(sender As Object, e As EventArgs) Handles btnOpenExcel.Click
        System.Diagnostics.Process.Start(txtFile.Text)
    End Sub
End Class
