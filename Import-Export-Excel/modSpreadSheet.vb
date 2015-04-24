



'Imports system
'Imports System.Configuration
'Imports system.Data
'Imports System.Data.SqlClient
'Imports Excel
'Imports Microsoft.Office.Interop.Excel
'Imports Microsoft.Office.Interop


    
'    ' Spreadsheet constants
'    'Public Const glMAX_Cols = 256                  '2010/11/15 - Office 2003
'    ' Used for the dynamic addition of assortments (Excel Export)
'    ' NOTE: Validation types are 7 characters to make parsing simpler
'    ' Import Spreadsheets are required to have the following fields in the following positions
'    ' the following used to suppress # decimals validation in bValidateField function
'    ' Function Codes
'    ' Miscellaneous Constants
'    ' Log file reporting for all imported rows
'    ' NOTE: This log file may be used by the Delete process & Update SpeedQuote process
'    '2012/01/04 - moved columns below up 1 so that we can move Row# to 1st column in logfile
'    '2012/01/17 - deleted log file column constants- were causing too many problems, were used in the wrong forms at times
'    ' Other log file constants
'    ' Import Special columns (require extra validations or updates)
'    ' Information about the spreadsheet columns (frmRefresh, frmImport)
'    'Public mlNbrSORTFields      As Long    '01/16/2009 - hn
'    'Public SORTFieldsArray()    As typSortField
'    ' Sort information (used by frmSelectItemCriteria)
'    'Public Type typSortField
''    sFieldValue     As String 
'    '    lFieldSortType  As Long
'    'End Type
'    ' Retrieved from UPCNumberPrefixes table - 6 digits used for generating UPC codes
'    ' ---------- Array of Original Import/Proposal Field values
'    ' Pre-read lookup tables for bValidateField - need to remember these over calls
'    ' ET 2012-12-11 - don't read the database for every lookup.  for the smaller ones, on the first reference
'    ' read and keep them in memory for later. do the short and easy ones for now.
'    ' commented out lookups may be too big to keep in memory - do them the old way for now.
'    Public Class modSpreadSheet
        
'        Private sItemStatus_ORIG             As  String           '2012/04/11
        
'    Private dtDataArray As [Object]
'    Friend gbCancelExport As Boolean
'    Friend gbValidationCancelled As Boolean
'    Friend gbCancelRefresh As Boolean
'    Friend glSHADING As Long
'    Friend gbFromIMPORT As Boolean
'    Friend gbFromREFRESH As Boolean
'    Friend gbFromEXPORT As Boolean
'    Friend gbFromSQSelect As Boolean
'    Friend gbFromParamItem As Boolean
'    Friend gbFromReportList As Boolean
'    Friend gbFromPOLSelect As Boolean
'    Friend bNewRevFromProposal As Boolean
'    Friend bItemMaterialError As Boolean  k            '01/15/2008        
'    Friend lngRecordNum As Long               'to return to same record in ProposalList        
'    Friend lngRevNum As Long
'    Friend lngProposalNum As Long
'    Friend msReportLabel As String
'    Friend gsItemNumber As String
'    Friend gsUPCNumber As String             'on import Spreadsheet before arrays are changed        
'    Friend gbProposal As sortmentsChanged
'    Private sItemSpecsArray_ORIG() As String              'see note where it's used!        
'    Private sItemArray_Orig() As String
'    Private s As sortmentArray_ORIG()
'    Friend gbFromSelectItem As Boolean             'christa - 09/03/2008       
'    Friend gbDEVITEM As Boolean              ' 2014/05/13 R        
'    'Private rsLookupItem As ADODB.Recordset  ' (316000. rows)        
'    'Private rsLookupItemSpecs As ADODB.Recordset = Nothing ' (316000. rows)        
'    'Private rsLookupItem_ As sortments
'    'Private rsLookupItemStatusCodes As ADODB.Recordset = Nothing ' (19 rows)        
'    'Private rsLookupCategory As ADODB.Recordset = Nothing ' (14 rows)        
'    'Private rsLookupCustomer As ADODB.Recordset = Nothing ' (151 rows)        
'    'Private rsLookupProgram As ADODB.Recordset = Nothing ' (177 rows) check if inactive?        
'    'Private rsLookupGrade As ADODB.Recordset = Nothing ' (4 rows)        
'    'Private rsLookupLicensor As ADODB.Recordset = Nothing ' (50 rows)        
'    'Private rsLookupSubProgram As ADODB.Recordset = Nothing ' (67 rows)        
'    'Private rsLookupFactory As ADODB.Recordset = Nothing ' (840 rows)        
'    'Private rsLookupVendor As ADODB.Recordset = Nothing ' (202 rows)        
'    'Private rsLookupSeason As ADODB.Recordset = Nothing ' (9 rows)        
'    'Private rsLookupFOBPoints As ADODB.Recordset = Nothing ' (26 rows)        
'    'Private rsLookupCountry As ADODB.Recordset = Nothing ' (9 rows)        
'    'Private rsLookupPackageTypes As ADODB.Recordset = Nothing ' (97 rows)        
'    'Private rsLookupEle_Connection As ADODB.Recordset = Nothing ' (41 rows)        
'    'Private rsLookupEle_IndoorOrIndoorOutdoor As ADODB.Recordset = Nothing ' (3 rows)        
'    'Private rsLookupEle_TrayPl As ticPaper

'    'Private rsLookupEle_FuseRating As ADODB.Recordset = Nothing ' (8 rows)        
'    'Private rsLookupEle_TrayOrBulk As ADODB.Recordset = Nothing ' (2 rows)        
'    'Private rsLookupEle_WireGauge As ADODB.Recordset = Nothing ' (9 rows)        
'    'Private rsLookupEle_LampB As eType

'    'Private rsLookupEle_LampBrightness As ADODB.Recordset = Nothing ' (5 rows)        
'    'Private rsLookupCertificationMark As ADODB.Recordset = Nothing ' (19 rows)        
'    'Private rsLookupCertificationType As ADODB.Recordset = Nothing ' (9 rows)        
'    'Private rsLookupBag_PaperType As ADODB.Recordset = Nothing ' (42 rows)        
'    'Private rsLookupBag_PrintType As ADODB.Recordset = Nothing ' (8 rows)        
'    'Private rsLookupBag_Finish As ADODB.Recordset = Nothing ' (4 rows)        
'    'Private rsLookupBag_HandleType As ADODB.Recordset = Nothing ' (12 rows)        
'    'Private rsLookupTree_Construction As ADODB.Recordset = Nothing ' (4 rows)        
'    'Private rsLookupTree_LightConstruction As ADODB.Recordset = Nothing ' (4 rows)        
'    'Private rsLookupSalesRep As ADODB.Recordset = Nothing ' (18 rows)        
'    'Private rsLookupLightType As ADODB.Recordset = Nothing ' (7 rows)        
'    'Private rsLookupBatteryType As ADODB.Recordset = Nothing ' (12 rows)        
'    'Private rsLookupEle_PlugTypeStackableStd As ADODB.Recordset = Nothing ' (7 rows)        
'    'Private rsLookupEle_ULWireType As ADODB.Recordset = Nothing ' (19 rows)        
'    'Private rsLookupLEDEpoxyType As ADODB.Recordset = Nothing ' (5 rows)        
'    'Private rsLookupPriceOption As ADODB.Recordset = Nothing ' (4 rows)        
'    'Private rsLookupFCAOrderPoint As ADODB.Recordset = Nothing ' (44 rows)        
'    'Private rsLookupTargetCertifiedPrinters As ADODB.Recordset = Nothing ' (37 rows)        
'    Private smessage As String
'    Private objEXCELName As String
'    Private _lColorIndex As Object
'    Private _sColHeader As String
'    Dim lNbrErrorsFieldLabels As Object
'    Dim sErrorMsgFieldLabels As String
'    Dim lColsOnSheet As Long
'    Dim glMAX_Cols As Integer
'    Dim lRevCOLPos As Integer
'    Dim lProposalNumberCOLPos As Integer
'    Dim lDB4FieldsFound As Integer
'    Dim lMax_Assortment_xx As Integer
'    Dim lMaxMaterialColumns As Integer
'    Dim msPRODDEV As Object
'    Dim msCREATIVE As Object
'    Dim msSHIP As Object
'    Dim msPHOTO As Object
'    Dim msHONGKONG As Object

'    Private Property lColorIndex(lColCounter As Long) As Object
'        Get
'            Return _lColorIndex
'        End Get
'        Set(value As Object)
'            _lColorIndex = value
'        End Set
'    End Property

'    Private Property sColHeader(lColCounter As Long) As String
'        Get
'            Return _sColHeader
'        End Get
'        Set(value As String)
'            _sColHeader = value
'        End Set
'    End Property

'    Private Property xlNone As Long

'    Private Property msUserGroup As String

'    Private Property msSALES As Object

'    Private Property msMKTGBASIC As String

'    Private Property msPRODUCTMGR As String

'    Friend Overridable Function bLoadSpreadsheetColumnArray(ByVal objExcel As Excel.Application) As Object
'        'REFRESH ONLY: lItemNumColPos added to put ItemNumber on Refresh log file
'        'REFRESH ONLY: lFunctionCodeColPos for refresh, for "END" in fc-col to indicate end of spreadsheet

'        ' for each column on spreadsheet checks info and reads Field table to get info, then saves in array
'        '10    On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sCellValue As String
'        Dim sCellAliasValue As String  ' Alias Value found between delimiters
'        Dim bFunctionCodeFound As Boolean
'        Dim bProposalNumberFound As Boolean
'        Dim bRevFound As Boolean

'        Dim bMarkColumn As Boolean
'        Dim lColCounter As Long
'        Dim lRowCounter As Long           '12/23/2008 - hn

'        Dim lRecordsetPos As Long
'        Dim sNewColHeading As String
'        Dim bFieldRequired As Boolean
'        Dim bXUnderscore As Boolean

'        Dim SQL As String
'        Dim lRequiredFields As Long
'        Dim lRequiredCounter As Long

'        Dim lRequiredFieldsFOUND As Long
'        Dim RequiredArray() As typColumn
'        Dim lCustomerCOLNumber As Long         '12/23/2008 - hn
'        Dim lMaterialX As Long
'        Dim sMaterialHeading As String

'        Dim sFunctioncode As String       '12/23/2008 - hn
'        Dim lColorIndex(0 To glMAX_Cols) As Long
'        Dim lColor As Long
'        Dim sClass As String
'        Dim lClass As Integer

'        bLoadSpreadsheetColumnArray = False
'        lColsOnSheet = 0
'        lDB4FieldsFound = 0
'        lProposalNumberCOLPos = 0 : lRevCOLPos = 0
'        lMax_Assortment_xx = 0
'        lRecordsetPos = 0
'        lMaxMaterialColumns = 0

'        ' save column headers and ColorIndexes
'        '2012/12/11 -hn- added test for 1st blank, nocolor column
'        For lColCounter = 1 To glMAX_Cols
'            sCellValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColCounter))
'            lColor = objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColCounter).Interior.ColorIndex
'            If Len(sCellValue) = 0 And lColor = xlNone Then
'                lColsOnSheet = lColCounter - 1
'                Exit For
'            End If
'            sColHeader(lColCounter) = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColCounter))
'            lColorIndex(lColCounter) = objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColCounter).Interior.ColorIndex

'            ' ET 2012-12-03 - check for obsolete field names and tell user to change them.
'            If sColHeader(lColCounter) = "Ele_UL" Then
'                lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                "Ele_UL - Obsolete Field, please change column header to ""CertificationMark"""
'            End If

'            If sColHeader(lColCounter) = "Ele_CSA" Then
'                lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                "Ele_CSA - Obsolete Field, please change column header to ""CertificationType"""
'            End If

'        Next lColCounter

'        If lColsOnSheet = 0 Then lColsOnSheet = glMAX_Cols
'        '2012/12/12 -hn- if more than 300 columns on spreadsheet, above code kept it at 0,
'        'if it did not find 1st blank column

'        If gbFromREFRESH = False And gbFromEXPORT = False Then
'            '12/23/2008 - hn - go thru all rows on the spreadsheet to find all the different CustomerNumbers
'            '                  and check whether the Required Columns are present
'            '                  the code will check the values of these fields later
'            '12/23/2008 - hn - added ImportRequired, ImportProductDevRequired, Restriction below
'            SQL = "SELECT FieldName, ImportRequired, ImportProductDevRequired, Restriction FROM Field " & vbCrLf & _
'                    "WHERE ImportRequired = 1 OR ImportProductDevRequired = 1 ORDER BY FieldName"
'            Dim rs As ADODB.Recordset
'            rs.Open SQL
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockReadOnly           '12/08/2008 - hn As Object 

'            If Not rs.EOF Then
'                lRequiredCounter = 1
'                lRequiredFields = rs.Recordcount
'                ReDim RequiredArray(0 To lRequiredFields)
'                Do Until rs.EOF
'                    RequiredArray(lRequiredCounter).sColumnName = rs!FieldName
'                    RequiredArray(lRequiredCounter).bRequired = rs!ImportRequired
'                    RequiredArray(lRequiredCounter).bRequiredPD = rs!ImportProductDevRequired
'                    If IsBlank(rs!Restriction) Then
'                        RequiredArray(lRequiredCounter).Restriction = ""
'                    Else
'                        '2010/03/16 - for HK1 - permission to see CostPricing columns
'                        If msUserGroup = "HK1" And RequiredArray(lRequiredCounter).Restriction = "CostPricing" Then
'                            RequiredArray(lRequiredCounter).Restriction = ""
'                        ElseIf msUserGroup <> "HK1" And RequiredArray(lRequiredCounter).Restriction = "RetailPricing" Then    '2010/03/18
'                            RequiredArray(lRequiredCounter).Restriction = ""
'                        Else
'                            RequiredArray(lRequiredCounter).Restriction = rs!Restriction
'                        End If
'                    End If
'                    RequiredArray(lRequiredCounter).bColumnRequiredFound = False    '12/23/2008 - hn - sets to false initially
'                    lRequiredCounter = lRequiredCounter + 1
'                    rs.MoveNext()
'                Loop
'            End If
'            '    End If

'            For lColCounter = 1 To glMAX_Cols

'                '            sCellValue = sGetCellValue(objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, lCOLCounter))
'                sCellValue = sColHeader(lColCounter)
'                '            lColorIndex = objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, lCOLCounter).Interior.ColorIndex
'                lColor = lColorIndex(lColCounter)
'                'first blank/no color column indicates last column on spreadsheet ...
'                If Len(sCellValue) = 0 And lColor = xlNone Then
'                    lColsOnSheet = lColCounter - 1
'                    Exit For
'                Else

'                    '----mark required columns are present
'                    For lRequiredCounter = 1 To lRequiredFields
'                        '06/25/2008 - hn - wasn't working for aliased fields
'                        sCellAliasValue = ""
'                        sCellAliasValue = sGetDelimitedValue(sCellValue, gsLEFT_DELIMITER, gsRIGHT_DELIMITER)
'                        If sCellValue = RequiredArray(lRequiredCounter).sColumnName Or _
'                            sCellAliasValue = RequiredArray(lRequiredCounter).sColumnName Then
'                            RequiredArray(lRequiredCounter).bColumnRequiredFound = True
'                            lRequiredFieldsFOUND = lRequiredFieldsFOUND + 1             '12/23/2008 - hn
'                            Exit For
'                        End If
'                    Next lRequiredCounter

'                    If sCellValue = "CustomerNumber" Then
'                        lCustomerCOLNumber = lColCounter
'                    End If
'                    '2014/05/06 RAS validation for class
'                    '2014/03/04 not validation class for now
'                    '2014/02/27 RAS setting variables for Class for further checking below
'                    If sCellValue = "Class" Then
'                        lClass = lColCounter
'                    End If
'                    '2014/05/13 RAS
'                    If sCellValue = "ItemStatus" Then
'                        llItemStatus = lColCounter
'                    End If
'                End If
'            Next lColCounter

'            If lRequiredFieldsFOUND <> lRequiredFields Then '12/23/2008 - hn - if equal then all columns are present irrespective of CustomerNumber
'                '12/23/2008 - hn - Find Distinct CustomerNumber to check that Required columns are present for that CustomerNumber
'                Dim sCustomerNumber As String
'                '            If sGetCellValue(objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, 1)) <> "FunctionCode" Then
'                If sColHeader(1) <> "FunctionCode" Then
'                    lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                    sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                    "FunctionCode Column is Required in Column 1!"
'                Else

'                    For lRowCounter = glDATA_START_ROW To glMAX_Rows
'                        If lCustomerCOLNumber < 1 Then                          '03/20/2009 - hn
'                            lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                            sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                            "CustomerNumber Column is Required! Please insert."
'                            '860                           GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        End If
'                        lColor = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, lCustomerCOLNumber).Interior.ColorIndex
'                        sCustomerNumber = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, lCustomerCOLNumber))
'                        If Len(sCustomerNumber) = 0 And lColor = xlNone Then
'                            Exit For 'last row for import
'                        End If

'                        sFunctioncode = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, 1))
'                        '2014/05/13 RAS
'                        If UCase(sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, llItemStatus))) = "DEV" Then
'                            gbDEVITEM = True
'                        Else
'                            gbDEVITEM = False
'                        End If
'                        If sFunctioncode <> "" Then
'                            Select Case sCustomerNumber
'                                '2014/05/29 RAS Adding 998 to 999 and 100
'                                Case "100", "999", "998" 'Product Dev
'                                    For lRequiredCounter = 1 To lRequiredFields
'                                        If RequiredArray(lRequiredCounter).bColumnRequiredFound = False And _
'                                                    RequiredArray(lRequiredCounter).bRequiredPD = True Then
'                                            If RequiredArray(lRequiredCounter).Restriction <> "" Then
'                                                Select Case msUserGroup
'                                                    Case msSALES, msHONGKONG, msPHOTO, msSHIP, msCREATIVE, msPRODDEV   '2010/03/16 - removed HK1
'                                                        'not an error, those users cant see these fields
'                                                        '                                        Case msHK1                                                          '2010/03/16 - HK1 can see CustomerPricing fields
'                                                        '                                            Application.DoEvents
'                                                    Case Else
'                                                        lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                                                        sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                                                            "Row: " & lRowCounter & "-'" & RequiredArray(lRequiredCounter).sColumnName & "' - this Spreadsheet Column is Required for CustomerNumber: " & sCustomerNumber
'                                                End Select
'                                            Else            '02/11/2009 - hn
'                                                lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                                                sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                                                "Row: " & lRowCounter & "-'" & RequiredArray(lRequiredCounter).sColumnName & "' - this Spreadsheet Column is Required for CustomerNumber: " & sCustomerNumber
'                                            End If
'                                        End If
'                                    Next lRequiredCounter

'                                Case Else   'all other customers - not Product Dev
'                                    '2014/08/15 RAS changing this to be just not DEV but not ORD
'                                    If gbDEVITEM = True Or (msUserGroup = msMKTGBASIC Or msUserGroup = msPRODUCTMGR) Then     '2014/05/13 Adding for DEV
'                                        For lRequiredCounter = 1 To lRequiredFields
'                                            If RequiredArray(lRequiredCounter).bColumnRequiredFound = False And _
'                                                                 RequiredArray(lRequiredCounter).bRequiredPD = True Then
'                                                If RequiredArray(lRequiredCounter).Restriction <> "" Then
'                                                    Select Case msUserGroup
'                                                        Case msSALES, msHONGKONG, msPHOTO, msSHIP, msCREATIVE, msPRODDEV   '2010/03/16 - removed HK1
'                                                            'not an error, those users cant see these fields
'                                                            '                                        Case msHK1                                                          '2010/03/16 - HK1 can see CustomerPricing fields
'                                                            '                                            Application.DoEvents
'                                                        Case Else
'                                                            lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                                                            sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                                                                     "Row: " & lRowCounter & "-'" & RequiredArray(lRequiredCounter).sColumnName & "' - this Spreadsheet Column is Required for CustomerNumber: " & sCustomerNumber
'                                                    End Select
'                                                Else            '02/11/2009 - hn
'                                                    lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                                                    sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                                                         "Row: " & lRowCounter & "-'" & RequiredArray(lRequiredCounter).sColumnName & "' - this Spreadsheet Column is Required for CustomerNumber: " & sCustomerNumber
'                                                End If
'                                            End If
'                                        Next lRequiredCounter
'                                    Else
'                                        For lRequiredCounter = 1 To lRequiredFields
'                                            If RequiredArray(lRequiredCounter).bColumnRequiredFound = False And _
'                                                             RequiredArray(lRequiredCounter).bRequired = True Then
'                                                If RequiredArray(lRequiredCounter).Restriction <> "" Then
'                                                    Select Case msUserGroup
'                                                        Case msSALES, msHONGKONG, msPHOTO, msSHIP, msCREATIVE, msPRODDEV    '2010/03/16 - removed HK1
'                                                            'not an error, those users cant see these fields
'                                                            '                                        Case msHK1                                                          '2010/03/16
'                                                            '                                            Application.DoEvents
'                                                        Case Else
'                                                            lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                                                            sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                                                            "Row: " & lRowCounter & "-'" & RequiredArray(lRequiredCounter).sColumnName & "' - this Spreadsheet Column is Required for CustomerNumber: " & sCustomerNumber
'                                                    End Select
'                                                Else            '02/11/2009 - hn
'                                                    lNbrErrorsFieldLabels = lNbrErrorsFieldLabels + 1
'                                                    sErrorMsgFieldLabels = sErrorMsgFieldLabels & vbCrLf & _
'                                                            "Row: " & lRowCounter & "-'" & RequiredArray(lRequiredCounter).sColumnName & "' - this Spreadsheet Column is Required for CustomerNumber: " & sCustomerNumber
'                                                End If
'                                            End If
'                                        Next lRequiredCounter
'                                    End If
'                            End Select
'                        End If
'                    Next lRowCounter
'                    ' End If
'                End If
'            End If

'            If rs.State <> 0 Then rs.Close() '2014/01/09 RAS adding statement based on SJM
'        End If
'        '-------------------------------------------------------
'        For lColCounter = 1 To lColsOnSheet 'glMAX_Cols '2012/12/11 - determined this value at start of this function
'            '1330          'If gbCancelRefresh = True Or gbValidationCancelled = True Or gbCancelExport = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            Call bUpdateStatusMessage(frmThis, "Retrieving Column " & CStr(lColCounter) & "...")

'            bMarkColumn = False
'            '1360         ' If gbCancelRefresh = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET

'            '        sCellValue = sGetCellValue(objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, lCOLCounter))
'            sCellValue = sColHeader(lColCounter)

'            '        lColorIndex = objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, lCOLCounter).Interior.ColorIndex
'            lColor = lColorIndex(lColCounter)
'            'first blank/no color column indicates last column on spreadsheet ...
'            If Len(sCellValue) = 0 And lColor = xlNone Then
'                lColsOnSheet = lColCounter - 1
'                Exit For
'            Else
'                'continue with saving columns found
'                SpreadsheetCOLArray(lColCounter).sColumnName = sCellValue
'                sCellAliasValue = sGetDelimitedValue(sCellValue, gsLEFT_DELIMITER, gsRIGHT_DELIMITER)
'                If Len(sCellAliasValue) > 0 Then
'                    sCellValue = sCellAliasValue
'                End If

'                If sCellValue = gsFunctionCode Then
'                    If Not bFunctionCodeFound Then
'                        ' Only mark the first FunctionCode if multiples are on the spreadsheet
'                        bFunctionCodeFound = True
'                        bMarkColumn = True
'                        lFunctionCodeCOLPos = lColCounter
'                    Else
'                        lFunctionCodeCOLPos = 0
'                    End If

'                ElseIf sCellValue = gsProposal Then
'                    If Not bProposalNumberFound Then
'                        ' Only track the first ProposalNumber if multiples are on the spreadsheet
'                        lProposalNumberCOLPos = lColCounter
'                        bProposalNumberFound = True
'                        bMarkColumn = True
'                    End If

'                ElseIf sCellValue = gsREV Then
'                    If Not bRevFound Then
'                        ' Only track the first Rev if multiples are on the spreadsheet
'                        lRevCOLPos = lColCounter
'                        bRevFound = True
'                        bMarkColumn = True
'                    End If

'                ElseIf sCellValue = gsCOL_ItemNumber Then              'Item # required on refresh logfile
'                    lItemNumCOLPos = lColCounter
'                ElseIf sCellValue = gsCOL_ProgramYear Then              '09/29/2008 - hn
'                    lProgYearCOLPos = lColCounter
'                ElseIf sCellValue = gsCOL_FactoryNumber Then               'new 10/15/2007
'                    lFactoryNumberCOLPos = lColCounter
'                ElseIf sCellValue = gsCOL_Lighted Then
'                    lLightedCOLPos = lColCounter

'                ElseIf sCellValue = gsCOL_ProductBatteriesIncluded Then     '03/19/2008
'                    lProductBatteriesCOLPos = lColCounter

'                ElseIf sCellValue = gsCol_Bag_SpecialEffects Then       '2009/11/18 - hn
'                    lBag_SpecialEffectsCOLPos = lColCounter

'                ElseIf sCellValue = gsCOL_CertifiedPrinterID Then      '2011/10/26
'                    lCertifiedPrinterIDCOLPos = lColCounter
'                ElseIf sCellValue = gsCOL_X_CertifiedPrinterName Then
'                    lX_CertifiedPrinterNameCOLPos = lColCounter

'                ElseIf sCellValue = gsCOL_Technologies Then
'                    lTechnologiesColPos = lColCounter
'                ElseIf sCellValue = gsCOL_X_Technologies Then
'                    lX_TechnologiesCOLPos = lColCounter
'                ElseIf sCellValue = gsCOL_X_UPCNumber Then
'                    lX_UPCCOLPos = lColCounter                           '2010/10/25
'                ElseIf sCellValue = gsCOL_X_PalletUPC Then              '2010/10/25
'                    lX_PalletUPCCOLPos = lColCounter

'                ElseIf Microsoft.VisualBasic.Left(sCellValue, 8) = "Material" And Len(sCellValue) > 8 And sCellValue <> "MaterialType" Then
'                    'store Materialx Column Positions
'                    sMaterialHeading = sCellValue
'                    sMaterialHeading = Replace(sMaterialHeading, "Material", "")
'                    If IsNumeric(sMaterialHeading) Then
'                        lMaxMaterialColumns = lMaxMaterialColumns + 1
'                        lMaterialX = CLng(sMaterialHeading)
'                        ReDim Preserve SpreadsheetMaterialColumnX(0 To lMaxMaterialColumns)
'                        If lMaterialX > 0 Then
'                            '                        'check for duplicate columns , in bValidateMaterialInfo
'                            SpreadsheetMaterialColumnX(lMaterialX) = lColCounter
'                        End If
'                    End If

'                    '            ElseIf sCellValue = "Photo" Then ' Code added 11/23/05 by Gary to support photos.
'                    '                SpreadsheetCOLArray(lCounter).sDB4Field = "CoreItemNumber"   ' Photo file names are based on the CoreItemNumber value
'                ElseIf sCellValue = msPHOTO Or sCellValue = "<Photo>" Then                        '2010/09/09
'                    lPhotoCOLPos = lColCounter
'                ElseIf sCellValue = "AlternatePhoto" Or sCellValue = "<AlternatePhoto>" Then     '2011/12/20 was missing last >
'                    lAlternatePhotoCOLPos = lColCounter

'                    ' new code for Refresh when adding Items with Assortments.....
'                ElseIf gbFromREFRESH = True Then                                     'check if assortments on spreadsheet
'                    lASSORTcounter = 0
'                    If Microsoft.VisualBasic.Left(sCellValue, 5) = cITEM_XX Then
'                        lASSORTcounter = Mid(sCellValue, 6, 3)
'                        sASSORTArray(lASSORTcounter).sItemInd = "1"                   '1 indicates that heading already on spreadsheet
'                        sASSORTArray(lASSORTcounter).lExcelItemColPos = lColCounter
'                    End If
'                    If Microsoft.VisualBasic.Left(sCellValue, 4) = cQTY_XX Then
'                        lASSORTcounter = Mid(sCellValue, 5, 3)
'                        sASSORTArray(lASSORTcounter).sQtyInd = "1"
'                        sASSORTArray(lASSORTcounter).lExcelQtyColPos = lColCounter
'                    End If
'                    If lMax_Assortment_xx < lASSORTcounter Then
'                        lMax_Assortment_xx = lASSORTcounter
'                    End If

'                End If

'                'obtains information on each column from Field table
'                If bLoadImportFieldData(sCellValue, lColCounter, SpreadsheetCOLArray(), _
'                                          lRecordsetPos, lDB4FieldsFound, _
'                                          bMarkColumn, bFieldRequired, bXUnderscore) = False Then
'                    ErrorHandler() 'TODO - GoTo Statements are redundant in .NET
'                End If

'            End If

'            Call bMarkDB4Column(objExcel, lColCounter, bMarkColumn, bFieldRequired, bXUnderscore, _
'                SpreadsheetCOLArray(lColCounter).sValidation, SpreadsheetCOLArray(lColCounter).bImport)
'        Next lColCounter


'        If lColCounter > glMAX_Cols Then
'            lColsOnSheet = glMAX_Cols
'        End If

'        ' Return to the first column
'        objExcel.Application.Workbooks(1).Worksheets(1).Columns(1).Select()
'        Application.DoEvents()
'        bLoadSpreadsheetColumnArray = True
'ExitRoutine:
'        'On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        '2008/12/16 -hn
'        '2013/07/24 -HN -expanded message below
'        MsgBox("Please check that you have the First Active worksheet open! " & vbCrLf & _
'            "Select 1st Worksheet in Spreadsheet and Save before proceeding, and close other open Spreadsheet(s)," & vbCrLf & _
'            Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bLoadSpreadsheetColumnArray")

'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bLoadSpreadsheetColumnArray, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        objExcel.Application.Workbooks(1).Close()
'        gbValidationCancelled = True                                '2013/07/24 -HN
'        Resume ExitRoutine
'        '    Resume Next 'for debug
'    End Function
'    Public Function bLoadImportFieldData(ByVal sFieldName As String,
'                                ByVal lColumn As Long, _
'                                ByRef SpreadsheetCOLArray() As typColumn, _
'                                ByRef lRecordsetPos As Long, _
'                                ByRef lDB4FieldsFound As Long, _
'                                ByRef bMarkColumn As Boolean, _
'                                ByRef bFieldRequired As Boolean, ByRef bXUnderscore As Boolean) As Boolean
'        'obtains information on each column from Field table
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rsField As ADODB.Recordset
'        bLoadImportFieldData = False
'        bFieldRequired = False
'        bXUnderscore = False

'        rsField = New ADODB.Recordset
'        If Microsoft.VisualBasic.Left(sFieldName, 8) = "Material" And sFieldName <> "MaterialType" Then
'            'the Field table no longer holds these fields;
'            'Material, breakdown, Cost ..changes the logic below
'            '        GoTo MaterialColumns'TODO - GoTo Statements are redundant in .NET
'        End If

'        sSQL = "SELECT TableName, Import, ImportValidation,ImportRequired, ImportProductDevRequired, " & _
'                    "ExcelColWidth, ImportDefaultValue, LookupField, ExcelShadeCell, Restriction " & _
'                "FROM Field WHERE FieldName = " & sAddQuotes(sFieldName)

'    End Function
        
'    Friend Overridable Function bLoadImportFieldData(ByVal sFieldName As String, ByVal lColumn As Long) As Object
'        'obtains information on each column from Field table
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rsField As ADODB.Recordset
'    bLoadImportFieldData = False		 As Object 
'        bFieldRequired = False
'        bXUnderscore = False

'        rsField = New ADODB.Recordset
'        If Microsoft.VisualBasic.Left(sFieldName, 8) = "Material" And sFieldName <> "MaterialType" Then
'            'the Field table no longer holds these fields;
'            'Material, breakdown, Cost ..changes the logic below
'            '        GoTo MaterialColumns'TODO - GoTo Statements are redundant in .NET
'        End If

'        sSQL = "SELECT TableName, Import, ImportValidation,ImportRequired, ImportProductDevRequired, " & _
'                    "ExcelColWidth, ImportDefaultValue, LookupField, ExcelShadeCell, Restriction " & _
'                "FROM Field WHERE FieldName = " & sAddQuotes(sFieldName)

'Dim     rsField.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic As Object

'        If Not rsField.EOF Then
'            SpreadsheetCOLArray(lColumn).sDB4Field = sFieldName
'            If Microsoft.VisualBasic.Left(sFieldName, 2) = "X_" Then
'                bXUnderscore = True
'            End If
'            SpreadsheetCOLArray(lColumn).bImport = CBool(rsField!Import)
'            SpreadsheetCOLArray(lColumn).sDB4Table = CStr(rsField!TableName)
'            If IsNull(rsField!ImportValidation) Then
'                SpreadsheetCOLArray(lColumn).sValidation = ""
'            Else
'                SpreadsheetCOLArray(lColumn).sValidation = CStr(rsField!ImportValidation)
'            End If
'            SpreadsheetCOLArray(lColumn).bRequired = CBool(rsField!ImportRequired)
'            bFieldRequired = SpreadsheetCOLArray(lColumn).bRequired
'            SpreadsheetCOLArray(lColumn).bRequiredPD = CBool(rsField!ImportProductDevRequired)
'            If IsNull(rsField!ExcelColWidth) Then
'                SpreadsheetCOLArray(lColumn).dblWidth = 10
'            Else
'                SpreadsheetCOLArray(lColumn).dblWidth = CDbl(rsField!ExcelColWidth)
'            End If
'            SpreadsheetCOLArray(lColumn).bExcelShadeCell = CBool(rsField!ExcelShadeCell)
'            If IsNull(rsField!Restriction) Then
'                SpreadsheetCOLArray(lColumn).Restriction = ""
'                SpreadsheetCOLArray(lColumn).bRestrictedCol = False
'            Else
'                If msUserGroup = "HK1" And rsField!Restriction = "CostPricing" Then
'                    SpreadsheetCOLArray(lColumn).Restriction = ""           '2010/03/16
'                    SpreadsheetCOLArray(lColumn).bRestrictedCol = False
'                ElseIf msUserGroup <> "HK1" And rsField!Restriction = "RetailPricing" Then 'new restriction for HK1 only
'                    SpreadsheetCOLArray(lColumn).Restriction = ""           '2010/03/18
'                    SpreadsheetCOLArray(lColumn).bRestrictedCol = False
'                Else
'                    SpreadsheetCOLArray(lColumn).Restriction = rsField!Restriction
'                    If bRestrictedGroup(SpreadsheetCOLArray(lColumn).Restriction) = True Then
'                        SpreadsheetCOLArray(lColumn).bRestrictedCol = True
'                    Else
'                        SpreadsheetCOLArray(lColumn).bRestrictedCol = False
'                    End If
'                End If
'            End If
'            If IsNull(rsField!ImportDefaultValue) Then
'                SpreadsheetCOLArray(lColumn).vDefault = ""
'            Else
'                SpreadsheetCOLArray(lColumn).vDefault = rsField!ImportDefaultValue
'            End If
'            If IsNull(rsField!LookupField) Then
'                SpreadsheetCOLArray(lColumn).sLookupField = ""
'            Else

'                SpreadsheetCOLArray(lColumn).sLookupField = rsField!LookupField
'            End If

'            If rsField!TableName <> gsNONE Then
'                SpreadsheetCOLArray(lColumn).lRecordsetPos = lRecordsetPos
'                lRecordsetPos = lRecordsetPos + 1
'                lDB4FieldsFound = lDB4FieldsFound + 1
'                bMarkColumn = True
'            Else
'                SpreadsheetCOLArray(lColumn).lRecordsetPos = glCOL_NOT_IN_DB4
'                ' lDB4FieldsFound = lDB4FieldsFound + 1 '2010/03/17 need the total number columns - for restricted cols when this was no there, did not find them even though it was set in the array
'                ' use lNoColsOnSheet instead
'            End If
'        Else
'MaterialColumns:
'            SpreadsheetCOLArray(lColumn).lRecordsetPos = glCOL_NOT_IN_DB4
'            SpreadsheetCOLArray(lColumn).sDB4Field = ""
'            SpreadsheetCOLArray(lColumn).sDB4Table = ""
'            SpreadsheetCOLArray(lColumn).bRequired = False
'            SpreadsheetCOLArray(lColumn).bRequiredPD = False
'            SpreadsheetCOLArray(lColumn).sLookupField = ""
'            SpreadsheetCOLArray(lColumn).sValidation = ""
'            SpreadsheetCOLArray(lColumn).vDefault = ""
'            SpreadsheetCOLArray(lColumn).bExcelShadeCell = False
'            SpreadsheetCOLArray(lColumn).Restriction = ""

'        End If
'    bLoadImportFieldData = True		 As Object 
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rsField.State <> 0 Then rsField.Close()
'        rsField = Nothing
'        Exit Function
'ErrorHandler:

'Dim     MsgBox Err.Description As Object 
'Dim  vbExclamation + vbMsgBoxSetForeground As Object 
' "modSpreadSheet-bLoadImportFieldData" As Object 'Translation Error-

'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bLoadImportFieldData As Object 'Translation Error-"
'Dim  Err Number " & Err.Number & "Error Description: " & Err.Description As Object 

'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function
'    Public Function bFindRestrictedColumns(SpreadsheetCOLArray() As typColumn, ByVal lColsOnSheet As Long, _
'                                        ByRef sRestrictedColumnNames As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lArrayCounter As Long

'        bFindRestrictedColumns = False
'        sRestrictedColumnNames = ""
'        For lArrayCounter = 1 To lColsOnSheet
'            If SpreadsheetCOLArray(lArrayCounter).bRestrictedCol = True Then
'                '2014/03/04 RAS commenting out class validation for now.
'                '2014/02/27 RAS only require Class for Target Customers
'                '            If SpreadsheetCOLArray(lArrayCounter).sDB4Field = "Class" Then
'                '            Else
'                '            End If
'                '            2010/01/13 - CUSTOMS now allowed FactoryNumber, FactoryName, VendorNumber, VendorName
'                If (SpreadsheetCOLArray(lArrayCounter).sDB4Field = "FactoryNumber" Or _
'                    SpreadsheetCOLArray(lArrayCounter).sDB4Field = "VendorNumber" Or _
'                    SpreadsheetCOLArray(lArrayCounter).sDB4Field = "X_FactoryName" Or _
'                    SpreadsheetCOLArray(lArrayCounter).sDB4Field = "X_VendorName") And (msUserGroup = "CUSTOMS" Or msUserGroup = "BASICPARTIAL") Then
'                    '2014/04/30 RAS Added the new group basic partial to skip the rule.
'                    Application.DoEvents()
'                Else
'                    sRestrictedColumnNames = sRestrictedColumnNames & SpreadsheetCOLArray(lArrayCounter).sDB4Field & " /  "
'                End If
'            End If
'        Next
'        Application.DoEvents()
'        bFindRestrictedColumns = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet: bFindRestrictedColumns")

'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFindRestrictedColumns, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function


'    Public Function bRestrictedColumnMessage(ByVal sRestrictedCOLNames As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        MsgBox("You do NOT have permission to REFRESH OR IMPORT these column(s): " & _
'            vbCrLf & vbCrLf & _
'            sRestrictedCOLNames & vbCrLf & vbCrLf & _
'            "Please DELETE these column(s) on Spreadsheet before Proceeding.", vbOKOnly + vbCritical + vbMsgBoxSetForeground, "DB4 - Delete Restricted Column(s)")
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet: bRestrictedColumnMessage")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "InbRestrictedColumnMessage, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function
        
'    Friend Overridable Function bValidateImportColumns(ByRef sErrorMsg As String, ByRef lNbrErrors As Long) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sDB4ColumnName As String
'        Dim lColCounter As Long
'        Dim sCheckCOLumnName As String
'        Dim lCheckCOLCounter As Long

'        Dim lCustomerNumberColPos As Long
'        Dim lProductBatteriesRequiredColPos As Long
'        Dim lLightTypeColPos As Long
'        Dim lTechnologiesColPos As Long '03/19/2008

'        Dim lCertifiedPrinterIDCOLPos As Long     '2011/10/26

'        bValidateImportColumns = False
'        sErrorMsg = ""
'        lNbrErrors = 0

'        ' The first three column positions are fixed for every Spreadsheet imported, error if not!
'        If SpreadsheetCOLArray(glFunctionCode_ColPos).sColumnName <> gsFunctionCode Then
'            sErrorMsg = sErrorMsg & gsFunctionCode & " not found in column " & glFunctionCode_ColPos & vbCrLf
'            lNbrErrors = lNbrErrors + 1
'        End If

'        If SpreadsheetCOLArray(glProposal_ColPos).sColumnName <> gsProposal Then
'            sErrorMsg = sErrorMsg & gsProposal & " not found in column " & glProposal_ColPos & vbCrLf
'            lNbrErrors = lNbrErrors + 1
'        End If

'        If SpreadsheetCOLArray(glREV_ColPos).sColumnName <> gsREV Then
'            sErrorMsg = sErrorMsg & gsREV & " not found in column " & glREV_ColPos & vbCrLf
'            lNbrErrors = lNbrErrors + 1
'        End If

'        ' Validate to ensure a column is not included > 1 on the same spreadsheet
'        For lColCounter = 1 To lColsOnSheet
'            If gbValidationCancelled = True Then Exit Function
'            sDB4ColumnName = SpreadsheetCOLArray(lColCounter).sDB4Field
'            If Len(sDB4ColumnName) > 0 Then
'                For lCheckCOLCounter = lColCounter To lColsOnSheet
'                    sCheckCOLumnName = SpreadsheetCOLArray(lCheckCOLCounter).sDB4Field

'                    If sDB4ColumnName = sCheckCOLumnName And lColCounter <> lCheckCOLCounter Then
'                        sErrorMsg = sErrorMsg & sDB4ColumnName & _
'                                    " is included on the spreadsheet more than once" & vbCrLf
'                        lNbrErrors = lNbrErrors + 1
'                        Exit For
'                    End If
'                Next
'            End If
'        Next

'        lCustomerNumberColPos = lGetSpreadsheetCOL(gsCOL_CustomerNumber, SpreadsheetCOLArray(), lColsOnSheet)
'        If lCustomerNumberColPos < 1 Then
'            sErrorMsg = sErrorMsg & gsCOL_CustomerNumber & " not found on the spreadsheet" & vbCrLf
'            lNbrErrors = lNbrErrors + 1
'        End If

'        lCertifiedPrinterIDCOLPos = lGetSpreadsheetCOL(gsCOL_CertifiedPrinterID, SpreadsheetCOLArray(), lColsOnSheet) '2011/10/26

'        lTechnologiesColPos = lGetSpreadsheetCOL(gsCOL_Technologies, SpreadsheetCOLArray(), lColsOnSheet) '05/22/2008 - hn
'        lLightTypeColPos = lGetSpreadsheetCOL("LightType", SpreadsheetCOLArray(), lColsOnSheet)
'        lProductBatteriesRequiredColPos = lGetSpreadsheetCOL(gsCOL_ProductBatteriesIncluded, SpreadsheetCOLArray(), lColsOnSheet)
'        If lTechnologiesColPos > 0 Or lLightTypeColPos > 0 Then
'            If lProductBatteriesRequiredColPos < 1 Then
'                sErrorMsg = sErrorMsg & "'" & gsCOL_ProductBatteriesIncluded & "' column not on spreadsheet if 'Technologies' OR 'LightType' is present. Please insert column." & vbCrLf
'                lNbrErrors = lNbrErrors + 1
'            End If
'        End If

'        bValidateImportColumns = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateImportColumns")

'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "bValidateImportColumns, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'Dim Public Function bValidateFunctionCodes(objExcel As Excel.Application
'Dim  ByRef lRowsOnSheet As Long
'Dim  _ As Object 

'    End Function
        
'    Friend Overridable Function bValidateFunctionCodes(ByVal objExcel As Excel.Application, ByRef lRowsOnSheet As Long, ByRef sErrorMsg As String, ByRef lNbrErrors As Long) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lRowsForImport As Long
'        Dim sFunctioncode As String
'        Dim sProposal As String
'        Dim sRev As String


'        'Dim lColorIndex  As Long      '2011/12/21 - need to check first 3 columns for no color! - no compile error??
'        Dim lRowCol1ColorIndex As Long
'        Dim lRowCol2ColorIndex As Long
'        Dim lRowCol3ColorIndex As Long
'        Dim lCounter As Long
'        Const sValidateMsg = "Validating Function Codes ... for row: "

'        bValidateFunctionCodes = False

'        ' The first row contains field names, start on row 2
'        For lCounter = glDATA_START_ROW To glMAX_Rows
'            If gbValidationCancelled = True Then Exit Function
'            Call bUpdateStatusMessage(Form_frmExcelImport, sValidateMsg & lCounter)
'            '2011/12/21 - below dsn't give a compile error??
'            '        lColorIndex = objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glFunctionCode_ColPos).Interior.ColorIndex
'            lRowCol1ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glFunctionCode_ColPos).Interior.ColorIndex
'            sFunctioncode = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glFunctionCode_ColPos))

'            sProposal = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glProposal_ColPos))
'            lRowCol2ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glProposal_ColPos).Interior.ColorIndex

'            sRev = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glREV_ColPos))
'            lRowCol3ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glREV_ColPos).Interior.ColorIndex

'            If sFunctioncode = gsENDofSpreadsheet Then
'                lRowsOnSheet = lCounter - glDATA_START_ROW + 1
'                Exit For

'            ElseIf Len(sFunctioncode) = 0 _
'                    And Len(sProposal) = 0 _
'                    And Len(sRev) = 0 _
'                    And (lRowCol1ColorIndex = xlNone And lRowCol2ColorIndex = xlNone And lRowCol3ColorIndex = xlNone) Then '2011/12/21
'                lRowsOnSheet = lCounter - glDATA_START_ROW + 1
'                Exit For

'            Else
'                If Len(sFunctioncode) > 0 Then
'                    lRowsForImport = lRowsForImport + 1

'                    Select Case sFunctioncode
'                        Case gsNEW_PROPOSAL
'                            Select Case msUserGroup
'                                Case msHONGKONG, msHK1, msCREATIVE, msPRODDEV
'                                    sErrorMsg = sErrorMsg & "ROW " & lCounter & ": " & _
'                                    " (FC = " & sFunctioncode & ") Permission Denied for Adding a New Proposal via Import Process." & vbCrLf
'                                    lNbrErrors = lNbrErrors + 1
'                                Case Else
'                                    If (Len(sProposal) = 0 And Len(sRev) <> 0) _
'                                        Or (Len(sProposal) <> 0 And Len(sRev) = 0) Then
'                                        sErrorMsg = sErrorMsg & "ROW " & lCounter & ": " & _
'                                            "When creating a new proposal, either both " & gsProposal & " " & _
'                                            "and " & gsREV & " must have values " & _
'                                            "or neither can have values" & vbCrLf
'                                        lNbrErrors = lNbrErrors + 1
'                                    End If
'                            End Select

'                        Case gsNEW_REVISION
'                            If Len(sProposal) = 0 Or Len(sRev) = 0 Then
'                                sErrorMsg = sErrorMsg & "ROW " & lCounter & ": " & _
'                                        "When creating a new revision, both " & gsProposal & " " & _
'                                        "and " & gsREV & " must have values" & vbCrLf
'                                lNbrErrors = lNbrErrors + 1
'                            End If

'                        Case gsEDIT_PROPOSAL
'                            Select Case msUserGroup
'                                Case msHONGKONG, msHK1, msCREATIVE, msPRODDEV
'                                    sErrorMsg = sErrorMsg & "ROW " & lCounter & ": " & _
'                                        " (FC = " & sFunctioncode & ") Permission Denied for changing records via Import Process." & vbCrLf
'                                    lNbrErrors = lNbrErrors + 1
'                                Case msHKBASIC
'                                    sErrorMsg = sErrorMsg & "ROW " & lCounter & ": " & _
'                                        " (FC = " & sFunctioncode & ") Permission Denied for changing records via Import Process." & vbCrLf
'                                    lNbrErrors = lNbrErrors + 1
'                                Case Else
'                                    If Len(sProposal) = 0 Or Len(sRev) = 0 Then
'                                        sErrorMsg = sErrorMsg & "ROW " & lCounter & ": " & _
'                                            "When editing(FC=A) a proposal, both " & gsProposal & " " & _
'                                            "and " & gsREV & " must have values" & vbCrLf
'                                        lNbrErrors = lNbrErrors + 1
'                                    End If
'                            End Select
'                        Case Else
'                            sErrorMsg = sErrorMsg & "ROW " & lCounter & ": " & _
'                                        "Undefined Function Code: " & sFunctioncode & vbCrLf
'                            lNbrErrors = lNbrErrors + 1

'                    End Select
'                Else
'                    Application.DoEvents()
'                End If

'                Application.DoEvents()
'            End If
'            Application.DoEvents()
'        Next

'        If lRowsForImport = 0 Then
'            sErrorMsg = sErrorMsg & "No rows available for import found on the spreadsheet.(Enter FunctionCode)" & vbCrLf '03/19/2008
'            lNbrErrors = lNbrErrors + 1
'        End If

'        bValidateFunctionCodes = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateFunctionCodes")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "bValidateFunctionCodes, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If

'        Resume ExitRoutine

'    End Function
        
'    Friend Overridable Function bValidateProposalReferences(ByVal objExcel As Excel.Application, ByRef sErrorMsg As String, ByRef lNbrErrors As Long) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim rsProposalRef As ADODB.Recordset
'        Dim sSQL As String
'        Dim sFunctioncode As String
'        Dim sProposal As String
'        Dim sRev As String

'        Dim lRowCounter As Long
'        Const sValidateMsg = "Validating Proposal References... for row: "

'        bValidateProposalReferences = False
'        rsProposalRef = New ADODB.Recordset

'        ' The first row contains field names, start on row 2
'        For lRowCounter = glDATA_START_ROW To lRowsOnSheet
'            If gbValidationCancelled = True Then Exit Function
'            Call bUpdateStatusMessage(Form_frmExcelImport, sValidateMsg & lRowCounter)
'            sFunctioncode = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, glFunctionCode_ColPos))
'            sProposal = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, glProposal_ColPos))
'            sRev = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, glREV_ColPos))

'            If Len(sFunctioncode) > 0 Then
'                If Len(sProposal) > 0 Or Len(sRev) > 0 Then
'                    If Not IsNumeric(sProposal) Or Not IsNumeric(sRev) Then
'                        sErrorMsg = sErrorMsg & "ROW " & lRowCounter & ": Both " & gsProposal & " and " & gsREV & " " & _
'                                "must be numeric if values are supplied" & vbCrLf
'                        lNbrErrors = lNbrErrors + 1
'                    Else
'                        sSQL = "SELECT ProposalNumber FROM " & gsItem_Table & " " & _
'                                    "WHERE ProposalNumber = " & sProposal & " AND Rev = " & sRev
'                        '2012/01/12 - run against Item table instead, is main table

'Dim                     rsProposalRef.Open sSQL As Object 
'                        Dim SSDataConn As Object
'                        Dim adOpenStatic As Object
'                        Dim adLockOptimistic As Object

'                        If rsProposalRef.EOF Then
'                            sErrorMsg = sErrorMsg & "ROW " & lRowCounter & ": ProposalNumber " & sProposal & ", " & _
'                                    "Rev " & sRev & " not in Item table" & vbCrLf
'                            lNbrErrors = lNbrErrors + 1
'                        End If

'                        rsProposalRef.Close()

'                    End If
'                End If
'            End If
'        Next lRowCounter
'        Application.DoEvents()
'        bValidateProposalReferences = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rsProposalRef.State <> 0 Then rsProposalRef.Close()
'        rsProposalRef = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateProposalReferences")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidateProposalReferences, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If

'        Resume ExitRoutine

'    End Function
        
'        Friend Overridable Function bLoadImportDataArray(ByVal objExcel As Excel.Application) As Object
''On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'Dim sFunctioncode       As String
'Dim  lProposalNumber As Long
'Dim  lRev As Long
'        Dim sItemNumber As String

'        Dim sCellValue As String
'Dim lRowCounter         As Long
'Dim  lColCounter As Long
'Dim  lMaterialCounter As Long


''read the worksheet into a local array in one big chunk
'Dim rng As Object 'Translation Error-
'Dim  TempArr As Object 


'    ' 0. load worksheet into local array
'    ' 1. load all rows with a valid Function Code into sImportDataArray
'    '    Reads function code, item number, from spreadsheet then loads all rows and columns
'    ' 2. Load column information for materials into SpreadsheetMaterialColumnX
    
'    ReDim sImportDataArray(1 To lRowsOnSheet, 1 To lColsOnSheet)
'    ReDim sX_Technology(1 To lRowsOnSheet)
'    ReDim sX_CertifiedPrinterID(1 To lRowsOnSheet)  '2011/10/26

'    If lMaxMaterialColumns > 0 Then
'        ReDim SpreadsheetMaterialValuesX(1 To lRowsOnSheet, 1 To lMaxMaterialColumns) As typMaterials
'    End If
    
'    With objExcel.Application.Workbooks(1).Worksheets(1)
'        Set rng = .Range(.Cells(1, 1), .Cells(lRowsOnSheet, lColsOnSheet))
''        Set rng = .range("A1:" & .Cells(lRowsOnSheet, lColsOnSheet).Address)  '  also works
'    End With
    
'    'Copy the values of the range to temporary array
'    TempArr = rng
    
'    'Confirm that an array was returned.
'    'Value will not be an array if the range is only 1 cell - this should never happen
'    If IsArray(TempArr) Then
''        do this later by using sGetCellValue
''        For row = 1 To lRowsOnSheet
''            For col = 1 To lColsOnSheet
''                'Make sure array value is not empty and is numeric
''                If Not IsEmpty(myArr(row, col)) And _
''                            IsNumeric(myArr(row, col)) Then
''                    'Replace numeric value with a string of the text.
''                    myArr(row, col) = range.Cells(row, col).Text
''                End If
''            Next
''        Next
'    Else
'        'Change TempArr into an array so you still return an array.
'        Dim TempArr1(1 To 1 As Object 'Translation Error-
'Dim  1 To 1) As Object 

'        TempArr1(1, 1) = TempArr
'        TempArr = TempArr1
'    End If

'    'loads all rows with a valid Function Code into sImportDataArray
'    bLoadImportDataArray = False		 As Object 
    
'    For lRowCounter = 1 To lRowsOnSheet
'        If gbValidationCancelled = True Then Exit Function
        
'        'sFunctionCode = sGetCellValue(objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, glFunctionCode_ColPos))
'        sFunctioncode = sGetCellValue(TempArr(lRowCounter, glFunctionCode_ColPos))
'        'sItemNumber = sGetCellValue(objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, lItemNumberCOLPos))
'        sItemNumber = sGetCellValue(TempArr(lRowCounter, lItemNumberCOLPos))
        
'        If Len(sFunctioncode) = 0 Then
'            Call bUpdateStatusMessage(Forms.frmExcelImport, "Skipping Spreadsheet Row " & lRowCounter & "...")
'        Else
'            If lRowCounter = 1 Then
'                Call bUpdateStatusMessage(Forms.frmExcelImport, "Loading Spreadsheet Headings Row " & lRowCounter & "...")
'            Else
'                Call bUpdateStatusMessage(Forms.frmExcelImport, "Loading Spreadsheet Row " & lRowCounter & " into memory ...")
'            End If
            
'            For lColCounter = 1 To lColsOnSheet
'                ' save two reads from the spreadsheet per row
'                Select Case lColCounter
'                    Case glFunctionCode_ColPos
'                        sImportDataArray(lRowCounter, lColCounter) = sFunctioncode
                    
'                    Case lItemNumberCOLPos
'                        sImportDataArray(lRowCounter, lColCounter) = sItemNumber
                    
'                    Case Else
'                        'sCellValue = sGetCellValue(objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, lCOLCounter))
'                        sCellValue = sGetCellValue(TempArr(lRowCounter, lColCounter))
'                        sImportDataArray(lRowCounter, lColCounter) = sCellValue
'                End Select
                
'                If lColCounter = lProposalNumberCOLPos And lRowCounter > 1 Then
'                    If IsNumeric(sCellValue) Then
'                        lProposalNumber = CDec(sCellValue)
'                    Else
'                        lProposalNumber = 0
'                    End If
'                End If
                
'                If lColCounter = lRevCOLPos And lRowCounter > 1 Then
'                    If IsNumeric(sCellValue) Then
'                        lRev = CDec(sCellValue)
'                    Else
'                        lRev = 0
'                    End If
'                End If
                
'                'store Material columns in SpreadsheetMaterialValuesX array
'                If lMaxMaterialColumns > 0 Then
'                    For lMaterialCounter = 1 To lMaxMaterialColumns
'                        If lColCounter = SpreadsheetMaterialColumnX(lMaterialCounter) Then
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sImportConcatenatedMaterial = Trim(sCellValue)
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialName = Trim(sCellValue)
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).lMaterialCOL = lColCounter
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).lProposalNumber = lProposalNumber
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).lRev = lRev
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sFunctioncode = sFunctioncode
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sItemNumber = sItemNumber
'                        End If
'                    Next lMaterialCounter
'                End If

'            Next lColCounter
'        End If
'    Next lRowCounter
    
'    Application.DoEvents
'    bLoadImportDataArray = True		 As Object 
'ExitRoutine:
'    Exit Function
'ErrorHandler:

''    If gbValidationCancelled = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'Dim     MsgBox Err.Description & " - Row: " & lRowCounter & " Column: " & lColCounter As Object 
'Dim  vbExclamation + vbMsgBoxSetForeground As Object 
' "modSpreadSheet-bLoadImportDataArray" As Object 'Translation Error-

    
'    If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'        smessage = "In bLoadImportDataArray As Object 'Translation Error-
'Dim  - Row: " & lRowCounter & " Column: " & lColCounter & " Err Number " & Err.Number & "Error Description: " & Err.Description As Object 

'        If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'        End If
'    End If
'    Resume ExitRoutine
'    End Function
'    Private Function bLoadSaveArray(ByRef vSaveArray As Object, ByRef dtValidationArray As typColumn, _
'                                    ByRef sImportDataArray() As String, _
'                                    ByVal lRowsOnSheet As Long, _
'                                    ByRef lNbrSaveFields As Long, _
'                                    ByVal sTableName As String) As Boolean
'        ' Loads the Save Array (to hold the records in memory) for each of ItemSpecs, Item, Item_Assortment tables
'        ' along with the parallel Validation Array

'        ' Loads the Save Array (to hold the records in memory) for each of ItemSpecs, Item, Item_Assortment tables
'        ' along with the parallel Validation Array

'        ' sImportDataArray holds mirror image of Import Spreadsheet columns OR Proposal fields
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim rsSaveFields As ADODB.Recordset
'        Dim rsProposalImage As ADODB.Recordset

'        Dim sSQL As String
'        Dim sSQLWhere As String

'        Dim sFieldName As String
'        Dim lFieldCounter As Long
'        Dim lNbrFields As Long

'        Dim vFieldData As Object
'        Dim sFieldData As String

'        Dim vDefaultValue As Object
'        Dim lDataRowCounter As Long

'        bLoadSaveArray = False
'        If gbValidationCancelled = True Then Exit Function

'        rsSaveFields = New ADODB.Recordset
'        rsProposalImage = New ADODB.Recordset

'        sSQL = "SELECT FieldName, Import," & _
'                "ImportValidation, " & _
'                "ImportRequired, " & _
'                "ImportProductDevRequired, " & _
'                "ImportDefaultValue " & _
'                "FROM Field WHERE TableName = " & sAddQuotes(sTableName) & " ORDER BY FieldName"

'Dim     rsSaveFields.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic As Object


'        If Not rsSaveFields.EOF Then
'            lNbrFields = rsSaveFields.Recordcount
'            lNbrSaveFields = lNbrFields
'        Else
'            lNbrFields = 0
'            lNbrSaveFields = 0
'        End If

'        If lNbrFields > 0 Then
'        ReDim vSaveArray(1 To lRowsOnSheet, 1 To lNbrFields + glFIXED_COLS)
'        ReDim dtValidationArray(1 To lNbrFields)
'        ReDim RowChangesARRAY(1 To lRowsOnSheet)           '2010/02/11 - don't want to add to vSaveArray in case we break some code somewhere..

'            vSaveArray(1, glFunctionCode_ColPos) = gsFunctionCode
'            vSaveArray(1, glProposal_ColPos) = gsProposal
'            vSaveArray(1, glREV_ColPos) = gsREV

'            sSQL = "SELECT "
'            For lFieldCounter = 1 To lNbrFields
'                If gbValidationCancelled = True Then Exit Function

'                ' Load the first row of the Save Array with the Field Names
'                sFieldName = rsSaveFields!FieldName
'                vSaveArray(1, lFieldCounter + glFIXED_COLS) = sFieldName

'                ' Load the parallel Validation Array for future use
'                dtValidationArray(lFieldCounter).sColumnName = sFieldName
'                dtValidationArray(lFieldCounter).sDB4Field = sFieldName
'                dtValidationArray(lFieldCounter).bImport = CBool(rsSaveFields!Import)
'                dtValidationArray(lFieldCounter).sDB4Table = sTableName
'                dtValidationArray(lFieldCounter).sValidation = CStr(rsSaveFields!ImportValidation)
'                dtValidationArray(lFieldCounter).bRequired = CBool(rsSaveFields!ImportRequired)
'                dtValidationArray(lFieldCounter).bRequiredPD = CBool(rsSaveFields!ImportProductDevRequired)
'                If IsNull(rsSaveFields!ImportDefaultValue) Then
'                    dtValidationArray(lFieldCounter).vDefault = ""
'                Else
'                    dtValidationArray(lFieldCounter).vDefault = rsSaveFields!ImportDefaultValue
'                End If

'                sSQL = sSQL & "[" & sFieldName & "], "

'                rsSaveFields.MoveNext()
'            Next lFieldCounter

'            sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 2) & " FROM " & sTableName & " "

'            For lDataRowCounter = glDATA_START_ROW To lRowsOnSheet
'                If gbValidationCancelled = True Then Exit Function

'                ' Populate the specific position columns first
'                vSaveArray(lDataRowCounter, glFunctionCode_ColPos) = sImportDataArray(lDataRowCounter, glFunctionCode_ColPos)

'                vSaveArray(lDataRowCounter, glProposal_ColPos) = sImportDataArray(lDataRowCounter, glProposal_ColPos)

'                vSaveArray(lDataRowCounter, glREV_ColPos) = sImportDataArray(lDataRowCounter, glREV_ColPos)

'                If sImportDataArray(lDataRowCounter, glProposal_ColPos) <> "" _
'                        And sImportDataArray(lDataRowCounter, glREV_ColPos) <> "" Then
'                    ' Populate the rest of the fields, so we have a 'memory image' of the record
'                    sSQLWhere = "WHERE ProposalNumber = " & sImportDataArray(lDataRowCounter, glProposal_ColPos) & " " & _
'                                        "AND Rev = " & sImportDataArray(lDataRowCounter, glREV_ColPos)

'Dim                 rsProposalImage.Open sSQL & sSQLWhere As Object 
'                    Dim SSDataConn As Object
'                    Dim adOpenStatic As Object
'                    Dim adLockOptimistic As Object


'                    If Not rsProposalImage.EOF Then
'                        ' Load the values from the From Proposal / From Rev in the database
'                        For lFieldCounter = 1 To lNbrFields
'                            If gbValidationCancelled = True Then Exit Function
'                            vFieldData = rsProposalImage(lFieldCounter - 1).Value
'                            sFieldData = sConvertVariantToString(vFieldData)
'                            vSaveArray(lDataRowCounter, lFieldCounter + glFIXED_COLS) = sFieldData
'                        Next lFieldCounter
'                    End If

'                    rsProposalImage.Close()
'                Else
'                    ' This is a brand-new Proposal / Rev - load the defaults
'                    For lFieldCounter = 1 To lNbrFields
'                        If gbValidationCancelled = True Then Exit Function
'                        vDefaultValue = dtValidationArray(lFieldCounter).vDefault
'                        vSaveArray(lDataRowCounter, lFieldCounter + glFIXED_COLS) = vDefaultValue
'                    Next lFieldCounter
'                End If
'            Next lDataRowCounter
'        End If

'        bLoadSaveArray = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rsSaveFields.State <> 0 Then rsSaveFields.Close()
'        rsSaveFields = Nothing
'        If rsProposalImage.State <> 0 Then rsProposalImage.Close()
'        rsProposalImage = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bLoadSaveArray")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bLoadSaveArray,  for table " & sTableName & " Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function
'    Private Function bMoveImportData(ByRef dtSpreadsheetCOLArray() As typColumn, _
'                            ByRef sImportDataArray() As String, _
'                            ByRef sItemSPECSArray() As String, ByVal lItemSpecsFields As Long, _
'                            ByRef sItemArray() As String, ByVal lItemFields As Long, _
'                            ByRef sAssortmentArray() As String, ByVal lAssortmentFields As Long, _
'                            ByVal lRowsOnSheet As Long, ByVal lColsOnSheet As Long) As Boolean

'        'Private Function bMoveImportData(ByRef dtSpreadsheetCOLArray( As Object) As  typColumn, _
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        ' Move the import data (non-fixed columns found in Field)
'        ' from the spreadsheet or data-entry form into the save array(s)
'        Dim lRowCounter As Long
'        Dim lColCounter As Long

'        Dim lSAVECOLCounter As Long
'        Dim sColumnName As String

'        Dim sStoreInTable As String
'        Dim sFieldData As String
'        Dim vSaveArray As Object
'        Dim lNbrSaveFields As Long


'        bMoveImportData = False

'        For lColCounter = glFIXED_COLS + 1 To lColsOnSheet
'            If gbValidationCancelled = True Then Exit Function

'            If dtSpreadsheetCOLArray(lColCounter).lRecordsetPos <> glCOL_NOT_IN_DB4 And _
'                dtSpreadsheetCOLArray(lColCounter).bImport = True Then

'                sColumnName = dtSpreadsheetCOLArray(lColCounter).sDB4Field
'                sStoreInTable = dtSpreadsheetCOLArray(lColCounter).sDB4Table

'                ' Determine the correct Save Array for moving the data
'                Select Case sStoreInTable
'                    Case gsItem_Table
'                        vSaveArray = sItemArray()
'                        lNbrSaveFields = lItemFields

'                    Case gsItemSpecs_Table
'                        If bPROPOSALFormIndicator = False Then
'                            vSaveArray = sItemSPECSArray()
'                            lNbrSaveFields = lItemSpecsFields
'                        End If

'                    Case gsItem_Assortments_Table
'                        vSaveArray = sAssortmentArray()
'                        lNbrSaveFields = lAssortmentFields

'                    Case Else
'                        MsgBox("Could not move data into " & sStoreInTable, vbExclamation + vbMsgBoxSetForeground, "bMoveImportData")
'                        '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End Select

'                ' Get the corresponding column position of the Save Array
'                For lSAVECOLCounter = glFIXED_COLS + 1 To glFIXED_COLS + lNbrSaveFields
'                    If vSaveArray(1, lSAVECOLCounter) = sColumnName Then
'                        Exit For
'                    End If
'                Next lSAVECOLCounter

'                For lRowCounter = glDATA_START_ROW To lRowsOnSheet
'                    If gbValidationCancelled = True Then Exit Function

'                    sFieldData = sImportDataArray(lRowCounter, lColCounter)
'                    '--for Spreadsheet DEFAULT values
'                    If (dtSpreadsheetCOLArray(lColCounter).vDefault <> "") Then
'                        If IsNumeric(sFieldData) = True Then
'                            If CDec(sFieldData) = 0 Then sFieldData = "0"
'                        End If
'                        If (IsNull(sFieldData) = True Or sFieldData = "" Or sFieldData = "0") Then      'SET DEFAULT FIELDS
'                            sFieldData = dtSpreadsheetCOLArray(lColCounter).vDefault
'                        End If
'                    End If

'                    Select Case sStoreInTable
'                        Case gsItem_Table
'                            sItemArray(lRowCounter, lSAVECOLCounter) = sFieldData
'                        Case gsItemSpecs_Table
'                            If bPROPOSALFormIndicator = False Then
'                                sItemSPECSArray(lRowCounter, lSAVECOLCounter) = sFieldData
'                            End If
'                        Case gsItem_Assortments_Table
'                            sAssortmentArray(lRowCounter, lSAVECOLCounter) = sFieldData
'                    End Select
'                Next lRowCounter
'            End If
'        Next lColCounter

'        bMoveImportData = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bMoveImportData")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bMoveImportData , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function


'    Private Function bSetYesNoNullValue(ByRef vFieldData As Object, ByRef vSaveArray As Object, _
'                                      ByVal lRow As Long, ByVal lColumn As Long) As Boolean
'        ' Translate certain spreadsheet values to yes/no/null equivalents
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        bSetYesNoNullValue = False

'        If UCase(vFieldData) = gsDENOTE_TRUE Or _
'                UCase(vFieldData) = "TRUE" Or _
'                UCase(vFieldData) = "YES" Or _
'                UCase(vFieldData) = "Y" Or _
'                vFieldData = "1" Then
'            vSaveArray(lRow, lColumn) = "1"
'            vFieldData = "YES"

'        ElseIf UCase(vFieldData) = "NO" Or _
'                UCase(vFieldData) = "FALSE" Or _
'                UCase(vFieldData) = "N" Or _
'                vFieldData = "0" Then
'            vSaveArray(lRow, lColumn) = "0"
'            vFieldData = "NO"
'        End If
'        bSetYesNoNullValue = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bSetYesNoNullValue")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bSetYesNoNullValue , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'        '2012/07/30 - added sVendorNumber
'    End Function
        
'    Private Function bValidateSaveArray(ByVal sFunctioncode As String, ByVal sTableName As String, ByVal lRow As Long, ByVal sProgramYR As String, ByVal lProgNumber As Long, ByVal sCustomerNumber As String, ByVal sVendorNumber As String, ByVal sFactoryNumber As String, ByVal bProductDevelopment As Boolean, ByVal vSaveArray As Object) As Object
'        'new 11/08/2007 added sCustomerNumber above
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lColCounter As Long
'        Dim lAssortCOLCounter As Long

'        Dim sProposalNum As String
'        Dim sRev As String

'        Dim sAssortQTYFieldName As String
'        Dim vAssortQTYFieldData As Object
'        Dim vFieldData As Object
'        Dim sFieldName As String
'        Dim sFieldValidation As String

'        Dim bFieldRequired As Boolean
'        Dim sValidationErrorMsg As String
'        Dim bImport As Boolean
'        Dim sRowErrorMsg As String        '11/24/2008 - hn
'        bValidateSaveArray = False
'        '2014/05/21 RAS Adding a varible for class index
'        Dim iClass As Integer
'        Dim sClass As Object
'        For lColCounter = 1 To UBound(vSaveArray, 2)
'            If vSaveArray(1, lColCounter) = "Class" Then
'                iClass = lColCounter
'            End If
'        Next lColCounter

'        For lColCounter = glFIXED_COLS + 1 To glFIXED_COLS + lNbrFields
'            '        If gbValidationCancelled = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            vFieldData = vSaveArray(lRow, lColCounter)

'            sProposalNum = vSaveArray(lRow, glProposal_ColPos)
'            sRev = vSaveArray(lRow, glREV_ColPos)

'            sFieldName = dtValidationArray(lColCounter - glFIXED_COLS).sDB4Field

'            If sFieldName = "CustomerNumber" Then           'new 10/23/2007
'                sCustomerNumber = vSaveArray(lRow, lColCounter)
'                '2014/05/20 RAS Validation check
'                If bValidateCustomerNumber(sCustomerNumber) = False Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " you do not have permission")
'                End If
'                sClass = vSaveArray(lRow, iClass)
'                If sClass <> "" Then
'                    If Len(sClass) < 2 And sClass < 10 Then
'                        vSaveArray(lRow, iClass) = "0" & sClass
'                        sClass = vSaveArray(lRow, iClass)
'                    End If
'                End If
'                If bValidateClass(CStr(sClass), sCustomerNumber) = False Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & ".Class  you do not have permission")
'                End If
'            End If
'            '        If sFieldName = "CertifiedPrinterID" Then       '2011/08/29 -2011/10/26 commented now multiselect
'            '            Dim sZeroPad As String 
'            '            If Not IsBlank(vSaveArray(lRow, lCOLCounter)) Then
'            '                sZeroPad = CharString(vSaveArray(lRow, lCOLCounter), 4, True)
'            '                sZeroPad = Replace(sZeroPad, " ", 0)
'            '                vSaveArray(lRow, lCOLCounter) = sZeroPad
'            '                vFieldData = vSaveArray(lRow, lCOLCounter)
'            '            End If
'            '        End If

'            '        If sFunctionCode = gsNEW_PROPOSAL And sFieldName = "SellPrice" Then
'            If sFunctioncode = gsNEW_PROPOSAL And sFieldName = gsCOL_FOBSellPrice Then      '2011/03/21
'                Select Case sCustomerNumber
'                    Case "100", "999"
'                        If lGetSpreadsheetCOL(sFieldName, dtSpreadsheetCOLArray(), lColsOnSheet) = 0 Then
'                            vFieldData = "" 'if column wasn't on Import spreadsheet
'                        End If
'                End Select
'            End If

'            bImport = dtValidationArray(lColCounter - glFIXED_COLS).bImport
'            sFieldValidation = dtValidationArray(lColCounter - glFIXED_COLS).sValidation

'            If bProductDevelopment Then
'                bFieldRequired = dtValidationArray(lColCounter - glFIXED_COLS).bRequiredPD
'            Else
'                bFieldRequired = dtValidationArray(lColCounter - glFIXED_COLS).bRequired
'            End If

'            If sFieldValidation = gsBOOLEAN Then
'                ' Translate certain spreadsheet values to boolean equivalents
'                Call bSetBooleanValue(vFieldData, vSaveArray, lRow, lColCounter)

'                '2013/05/15 -HN- Even if it's false on Spreadsheet it MUST be TRUE
'                'for CargoReceipt, ChildLabor, CommercialInvoice, GCC, PackingList, PrisonLaborDoc
'                If dtValidationArray(lColCounter - glFIXED_COLS).vDefault = "TRUE" Then
'                    '                If vSaveArray(lRow, lCOLCounter) <> "TRUE" Then
'                    '                   vSaveArray(lRow, lCOLCounter) = "TRUE"
'                    If vFieldData <> "TRUE" Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " should be = True!")
'                    End If
'                End If
'            End If

'            If sFieldValidation = gsYESNONULL Then
'                ' Translate certain spreadsheet values to yes/no/null equivalents
'                Call bSetYesNoNullValue(vFieldData, vSaveArray, lRow, lColCounter)         '03/19/2008
'            End If
'            '        If bFieldRequired = True And sFieldName <> "SubProgram" And sFieldName <> "Grade" Then
'            'new 11/07/2007
'            If bFieldRequired = True Then                                       'new 10/23/2007
'                'this checks if the actual column exists
'                If lGetSpreadsheetCOL(sFieldName, dtSpreadsheetCOLArray(), lColsOnSheet) = 0 Then
'                    If sFunctioncode = "A" Then
'                        '10/24/2007 - according to TW only required for FunctionCode=A (on import)for SellPrice, FactoryFOBCost
'                        '04/07/2008 next line new
'                        '06/25/2008 - hn - put values in quotes in next line
'                        If Not IsBlank(sCustomerNumber) And sCustomerNumber <> gs999PD_ACCOUNT And sCustomerNumber <> gs100_ACCOUNT And sFieldName <> "FOBSellPrice" Then   '2010/09/10
'                            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " is required, for FunctionCode=A, (not in spreadsheet or base record)")
'                        End If
'                    End If
'                End If
'            End If

'            '03/05/2009 - hn - new ------------------------------------------
'            If bPROPOSALFormIndicator = False Then               'Proposal Form handles this before saving
'                Select Case sFieldName
'                    '2014/02/26 RAS blanking out Class for non target
'                    Case "CustomerHSNumber", "CustomerDutyRate", "Class"
'                        If sFunctioncode = gsNEW_PROPOSAL Then
'                            Select Case sCustomerNumber
'                                Case 102, 103, 104, 206, 235, 248, 252, 253, 254, 255, 257, 888, 889 'Target
'                                    'don't blank out CustomerHSNumber and CustomerDutyRate
'                                Case 101, 998                       '03/09/2009 - hn
'                                    'don't blank out CustomerHSNumber and CustomerDutyRate
'                                Case Else
'                                    'set CustomerHSNumber, CustomerDutyRate = blank
'                                    vSaveArray(lRow, lColCounter) = ""
'                            End Select
'                        End If
'                End Select
'            End If
'            '-----------------------------------------------------------------

'            '11/19/2007 - check if SellPrice column exists for FC=A
'            '        If sFieldName = "SellPrice" Then 'bFieldRequired = false for account <> 999, 100
'            If sFieldName = gsCOL_FOBSellPrice Then '2011/03/21
'                If sFunctioncode = gsNEW_PROPOSAL Then
'                    If sCustomerNumber <> "" Then '11/26/2007 'if blank, error caught later
'                        If (sCustomerNumber <> 999 And sCustomerNumber <> 100 And sCustomerNumber <> 998 And gbDEVITEM = False) Then
'                            If lGetSpreadsheetCOL(sFieldName, dtSpreadsheetCOLArray(), lColsOnSheet) = 0 Then
'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " is required, for FunctionCode=A")
'                            End If
'                        End If
'                    End If
'                End If
'            End If

'            If sFieldName = gsCOL_REGLinePrice And vFieldData = "" Then            'new 11/08/2007
'                '2014/05/12 RAS DEV and 999,998 required fields this is not a required field
'                'If scustomernumber = gs999PD_ACCOUNT Or scustomernumber = gs100_ACCOUNT Then
'                If sCustomerNumber = gs100_ACCOUNT Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - value > 0  is required, for CustomerNumber: " & sCustomerNumber)
'                End If
'            End If

'            ' 11/19/2007
'            ' 2014/05/28 RAS for item status of DEV or 999,998,100 do not do the check
'            ' 2014/08/18 RAS adding mkt basic and product mgr to if statemetent
'            If sFieldName = gsCOL_FOBSellPrice Then
'                If sCustomerNumber <> "" And (sCustomerNumber = gs999PD_ACCOUNT Or sCustomerNumber = gs998PD_ACCOUNT Or sCustomerNumber = gs100_ACCOUNT Or gbDEVITEM = True Or msUserGroup = msMKTGBASIC Or msUserGroup = msPRODUCTMGR) Then        '11/26/2007
'                    '2014/05/28 do not do the check
'                    '                If gbDEVITEM = True Then
'                    '                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - value must be blank, for ItemStatus of DEV")
'                    '                Else
'                    '                    If vFieldData <> "" Then
'                    '                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - value must be blank, for CustomerNumber: " & scustomernumber)
'                    '                    End If
'                    '                End If
'                Else            '11/19/2007
'                    If vFieldData = "" Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - value > 0  is required, for CustomerNumber: " & sCustomerNumber)
'                    End If

'                End If
'            End If


'            ' for some programs SubProgram and Grade can be blank; validated later
'            '10/12/2007 SellPrice validation is caught later in bValidateCoreFields function
'            '        If bFieldRequired = True And vFieldData = "" And sFieldName <> "SellPrice" Then '11/19/2007 sellprice caught above
'            If bFieldRequired = True And vFieldData = "" And sFieldName <> gsCOL_FOBSellPrice Then '2011/03/21 sellprice caught above
'                '2014/08/15 RAS skipping validation for marketing and product mgr
'                '2014/09/17 RAS if item status <> DEV and Marketing or PROD mgr then skip
'                If (msUserGroup = msMKTGBASIC Or msUserGroup = msPRODUCTMGR) And gbDEVITEM = False Then
'                    ' skip validation
'                Else
'                    '10/29/2007 TW said to ignore for now:
'                    'Or sFieldName = "FactoryFCACost"
'                    If sFieldName = "FactoryFOBCost" Then      'new 10/26/2007
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - value > 0  is required")

'                    ElseIf sFieldName = gsCOL_FOBSellPrice Then
'                        '2014/08/15 RAS adding marketingbasic and product mgr to the skip
'                        If sCustomerNumber <> gs999PD_ACCOUNT And sCustomerNumber <> gs100_ACCOUNT Then
'                            If sFunctioncode = gsNEW_PROPOSAL And bPROPOSALFormIndicator = False Then '10/29/2007 can ignore from spreadsheet
'                            Else
'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - value > 0  is required, for CustomerNumber: " & sCustomerNumber & " FunctionCode=A")
'                            End If
'                        End If
'                    Else
'                        Select Case sFieldName      '12/08/2008 - hn
'                            '03/25/2009 - hn - added FCAOrderPoint below:
'                            Case "AltCost", "AltSellPrice", "FactoryFCACost", "FCASellPrice", "NetFirstCost", "NetFirstCostAlt", "RegularLinePrice", "SubProgram", "Grade", "FCAOrderPoint"      '12/12/2008 - hn added RegularLinePrice
'                                'can be blank, checked later
'                            Case Else
'                                '                '10/30/2007 - hn new
'                                '                If sFieldName = "SubProgram" Or sFieldName = "Grade" Then 'can be blank , checked later by ProgramYear
'                                '                Else
'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - value is required, for CustomerNumber: " & sCustomerNumber)
'                                '                End If
'                        End Select
'                    End If
'                End If
'            ElseIf vFieldData <> "" Then
'                If sFieldName = "Class" Then                                '11/20/2008 - hn
'                    If Len(vFieldData) < 2 And vFieldData < 10 Then
'                        vSaveArray(lRow, lColCounter) = "0" & vFieldData
'                    End If
'                End If
'                'added sProposalNum, sRev, sProgramYR, sFieldName as parameter to bValidateFieldName function
'                'new 10/23/2007 added sCustomerNumber below
'                If bValidateField(sFunctioncode, sProposalNum, sRev, sProgramYR, _
'                        sFieldName, vFieldData, sCustomerNumber, sFactoryNumber, bImport, _
'                        sFieldValidation, sValidationErrorMsg) = False Then
'                    '                       lNbrErrors = lNbrErrors + 1
'                    '                       sErrorMsg = sErrorMsg & "ROW " & lRow & ": " & sTableName & "." & sFieldName & "[" & vFieldData & "] " & sValidationErrorMsg
'                    '                       If lGetSpreadsheetCOL(sFieldName, dtSpreadsheetCOLArray(), lColsOnSheet) = 0 Then
'                    '                           sErrorMsg = sErrorMsg & " (there is a problem with the base record)" & vbCrLf
'                    '                       Else
'                    '                           sErrorMsg = sErrorMsg & vbCrLf
'                    '                       End If
'                    If lGetSpreadsheetCOL(sFieldName, dtSpreadsheetCOLArray(), lColsOnSheet) = 0 Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & "[" & vFieldData & "] " & sValidationErrorMsg & "(there is a problem with the base record)")
'                        sRowErrorMsg = sRowErrorMsg & " (there is a problem with the base record)" & vbCrLf
'                    Else
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & "[" & vFieldData & "] " & sValidationErrorMsg)
'                    End If
'                End If
'                '2014/03/04 RAS commenting out the class validation for now.
'                '2014/05/06 RAS class validation is back in.
'            ElseIf sFieldName = "Class" Then
'                '2014/02/28 RAS Class is now required for Target customers
'                '2014/05/20 RAS Adding 998 to list of customers that class is required
'                Select Case sCustomerNumber
'                    Case 101, 102, 103, 104, 206, 235, 248, 252, 253, 254, 255, 257, 888, 889, 998
'                        If Len(vFieldData) < 2 And vFieldData < 10 Then
'                            vSaveArray(lRow, lColCounter) = "0" & vFieldData
'                        Else
'                            If IsBlank(vFieldData) = True Then
'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - value cannot be blank, for CustomerNumber: " & sCustomerNumber)
'                            End If
'                        End If
'                    Case Else
'                        If Len(vFieldData) < 2 And vFieldData < 10 Then
'                            vSaveArray(lRow, lColCounter) = "0" & vFieldData
'                        End If
'                End Select
'            End If

'            '2014/07/08 RAS check for div / 0 error
'            If vFieldData = gsDENOTE_CELL_ERROR Then
'                If Microsoft.VisualBasic.Left(sFieldValidation, 7) = gsNUMERIC Or Microsoft.VisualBasic.Left(sFieldValidation, 7) = gsDECIMAL Then
'                    If dtValidationArray(lColCounter - glFIXED_COLS).sDB4Field <> "" Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sFieldName & " has to have a numeric/decimal value")
'                    End If
'                End If
'            End If

'            ' end new 11/08/2007

'            ' Check for Assortment ItemNumbers(Item_**) that Quantity(Qty_**) is not null
'            If sTableName = gsItem_Assortments_Table Then
'                If Microsoft.VisualBasic.Left(sFieldName, 5) = cITEM_XX Then
'                    For lAssortCOLCounter = glFIXED_COLS + 1 To glFIXED_COLS + lNbrFields
'                        vAssortQTYFieldData = vSaveArray(lRow, lAssortCOLCounter)
'                        sAssortQTYFieldName = dtValidationArray(lAssortCOLCounter - glFIXED_COLS).sDB4Field
'                        If Microsoft.VisualBasic.Left(sAssortQTYFieldName, 4) = "Qty_" And Mid(sAssortQTYFieldName, 5, 2) = Mid(sFieldName, 6, 2) Then
'                            If IsNull(vFieldData) And IsNull(vAssortQTYFieldData) Then
'                                Exit For
'                            ElseIf vFieldData > "" And vAssortQTYFieldData > "" Then
'                                Exit For
'                            ElseIf vFieldData = "" And vAssortQTYFieldData = "" Then
'                                Exit For
'                            Else
'                                '                            lNbrErrors = lNbrErrors + 1
'                                Dim sAssortErrMsg As String
'                                sAssortErrMsg = " For Assortments, " & sFieldName & ": " & vFieldData & _
'                                            ", " & sAssortQTYFieldName & ": " & vAssortQTYFieldData & ", both should have values." & vbCrLf

'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - " & sAssortErrMsg)
'                                Exit For
'                            End If
'                        End If
'                    Next lAssortCOLCounter '2012/07/30

'                    '               2012/07/30 - check that ComponentItem has same ProgramYear etc
'                    Dim sCompItemErrMsg As String
'                    If Microsoft.VisualBasic.Left(sFieldName, 5) = cITEM_XX Then
'                        If Not IsBlank(vFieldData) Then
'                            If bCheckParentAssortmentWithComponentItem(vSaveArray(lRow, 2), vSaveArray(lRow, 3), sProgramYR, sCustomerNumber, sVendorNumber, vFieldData, sCompItemErrMsg) = False Then
'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sTableName & "." & sFieldName & " - " & sCompItemErrMsg)
'                                '                            Exit For
'                            End If
'                        End If
'                    End If

'                End If
'            End If
'            'added 11/24/2008 - hn - to match ssdev
'            If sRowErrorMsg <> "" Then
'                '            If bCreateNEWComponentItems = True Then                     '07/14/2008 - hn
'                '                sRowErrorMsg = "For Proposal:" & sProposalNum & " Rev:" & sRev & vbCrLf & sRowErrorMsg
'                '            End If
'                sErrorMsg = sErrorMsg & sRowErrorMsg
'                sRowErrorMsg = ""
'            End If

'        Next lColCounter

'        bValidateSaveArray = True

'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateSaveArray")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidateSaveArray , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        ' Resume Next '' TODO TESTING ONLY
'        Resume ExitRoutine

'        'sDataArray from Import or Proposal process, mirrors Import SpreadSheet or
'        'Proposal Form(which equals one row) before changes

'Dim Public Function bValidateData(frmThis As Form _
'        ,sDataArray() As String _
'        ,dtCOLArray() As typColumn, _


'                    ByRef sItemArray() As String , ByRef lItemFields As Long, _
'                    ByRef sItemSPECSArray() As String , ByRef lItemSpecsFields As Long, _
'                    ByRef sAssortmentArray() As String , ByRef lAssortmentFields As Long, _
'                    ByRef sX_CertifiedPrinterNames() As String , ByRef sX_Technology() As String , _
'                    dtCOLPos As typSpecialCOLPos, _
'                    ByVal lRowsToValidate As Long, ByVal lNbrCOLsToValidate As Long, _
'                    ByRef sErrorMsgFieldContent As String , ByRef lErrorsFieldContent As Long, _
'                    ByRef sErrorMsgCoreFields As String , ByRef lErrorsCoreFields As Long, _
'                    ByRef sReturnErrormsg As String ) As Boolean

''On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'Dim dtValidateITEMArray()           As typColumn 'typColumn=info about the Column(FieldName) as obtained from Field table
'Dim dtValidateItemSPECSArray()      As typColumn
'Dim dtValidateAssortmentArray()     As typColumn
'Dim lcustomernumber                 As Long
'Dim  bProductDevelopment As Boolean

'        Dim sCustomerNumber As String
'Dim lProgNumber                     As Long
'        Dim sProgNumber As String

'Dim lRowCounter                     As Long
'        Dim sPROGYR As String
'Dim sFunctioncode                   As String
'        Dim sX_TechnologyDescr As String
'        Dim sX_CertifiedPrinterName As String

'    bValidateData = False
''    Debug.Print "In bValidateData" & Now()
'    Debug.Print "Loading Item table Data..." & Now()
    
'    ' Load arrays for saving and validation
'    Call bUpdateStatusMessage(frmThis, "Loading Item table Data...")
    
'    'Save Item table data into array
'    If bLoadSaveArray(sItemArray(), dtValidateITEMArray(), sDataArray(), _
'                        lRowsToValidate, lItemFields, gsItem_Table) = False Then
'        sReturnErrormsg = "Could not load the Item table data into memory"
''        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'    End If

'    If bPROPOSALFormIndicator = False Then
'        Debug.Print "Loading ItemSpecs table Data..." & Now()
        
'        Call bUpdateStatusMessage(frmThis, "Loading ItemSpecs table Data...")
        
'        'Save ItemSpecs table data in array
'        If bLoadSaveArray(sItemSPECSArray(), dtValidateItemSPECSArray(), sDataArray(), _
'                                lRowsToValidate, lItemSpecsFields, gsItemSpecs_Table) = False Then
'            sReturnErrormsg = "Could not load the ItemSpecs table data into memory"
''            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If
'    End If
''    If bFromProposalForm = False Then    ' saving Proposal Form Assortments differently...
'    'save Item_Assortments data in array
''    Debug.Print "Loading Assortment Data..." & Now()
    
'        Call bUpdateStatusMessage(frmThis, "Loading Assortment Data...")
'        If bLoadSaveArray(sAssortmentArray(), dtValidateAssortmentArray(), sDataArray(), _
'                                lRowsToValidate, lAssortmentFields, gsItem_Assortments_Table) = False Then
'            sReturnErrormsg = "Could not load the Assortment data into memory"
''            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If
        
'        'check that each Item_XX matches Assortment Proposal for Year, Vendor etc
'        ' If Assortment Dev had been rolled out this would not have been neccessary!! :-(
''    End If

'    'up to  here Arrays have the values from the tables before updating...
'    '----------- (to keep orig values) -------------------
'    If bPROPOSALFormIndicator = False Then sItemSpecsArray_ORIG() = sItemSPECSArray()
'    sItemArray_Orig() = sItemArray()
'    sAssortmentArray_ORIG() = sAssortmentArray()

'    '---- original values are overwritten in next routines by values from Excel spreadsheet or Proposal Form

'    Debug.Print "Moving Spreadsheet/Proposal Data " & Now()
    
'    'Overwrite data in Save Arrays with the supplied spreadsheet data/also true for Proposal Form
'    Call bUpdateStatusMessage(frmThis, "Moving Spreadsheet/Proposal Data(into arrays for Item, ItemSpecs, etc tables)...")
    
'    If bMoveImportData(dtCOLArray(), sDataArray(), _
'                            sItemSPECSArray(), lItemSpecsFields, _
'                            sItemArray(), lItemFields, _
'                            sAssortmentArray(), lAssortmentFields, _
'                            lRowsToValidate, lNbrCOLsToValidate) = False Then
'        sReturnErrormsg = "Could not move the data changes into memory"
''        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'    End If

'    Debug.Print "Retrieving Core Field Positions..." & Now()
    
'    ' Retrieve the Core Field column positions
'    Call bUpdateStatusMessage(frmThis, "Retrieving Core Field Positions...")
    
'    '                   ---added: sAssortmentArray(), lAssortmentFields) below ....
'    If bGetSpecialSaveArrayCOLPositions(dtCOLPos, dtCOLArray(), _
'                                    sItemSPECSArray(), lItemSpecsFields, _
'                                    sItemArray(), lItemFields, _
'                                    sAssortmentArray(), lAssortmentFields) = False Then
'        sReturnErrormsg = "Could not retrieve the Core Field column positions"
''        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'    End If

'    Debug.Print "Validating Rows " & Now()
    
'    ' Validate data in Save Arrays
'    sErrorMsgFieldContent = "FIELD CONTENT VALIDATION:" & vbCrLf & vbCrLf
'    For lRowCounter = glDATA_START_ROW To lRowsToValidate
'        If Len(sItemArray(lRowCounter, glFunctionCode_ColPos)) = 0 Then
'            Call bUpdateStatusMessage(frmThis, "Skipping Row " & CStr(lRowCounter) & "...")
'        Else
'            Call bUpdateStatusMessage(frmThis, "Validating Row " & CStr(lRowCounter) & "...")
'             'check errors have changed for each row to avoid putting extra vbcrlf's after each row
'            Dim lRowContentErrors As Long
'            lRowContentErrors = lErrorsFieldContent
'            sCustomerNumber = sItemArray(lRowCounter, dtCOLPos.lcustomernumber)
'            If IsNumeric(sCustomerNumber) = True Then
'                lcustomernumber = sCustomerNumber
'            Else
'                lcustomernumber = 0
'            End If
            
'            sProgNumber = sItemArray(lRowCounter, dtCOLPos.lProgramNumber)
'            If IsNumeric(sProgNumber) = True Then
'                lProgNumber = sProgNumber
'            Else
'                lProgNumber = 0
'            End If
'            Select Case lcustomernumber
'            '2014/05/07 RAS added "DEV" itemstatus accounts to development
'                Case gs999PD_ACCOUNT, gs998PD_ACCOUNT
'                    bProductDevelopment = True
'                Case Else
'                    bProductDevelopment = False
'            End Select
'            '2014/05/07 RAS added "DEV" itemstatus accounts to development
'            If UCase(sItemArray(lRowCounter, dtCOLPos.lItemStatus)) = "DEV" Then
'                bProductDevelopment = True
'            End If
'            '2014/05/23 RAS adding a check if the item is ORD and in the marketing then they cannot import this row.
'            If sItemArray(lRowCounter, dtCOLPos.lItemStatus) = "ORD" And (msUserGroup = "MKTGBASIC" Or msUserGroup = "PRODUCTMGR") Then
'                sErrorMsgFieldContent = sErrorMsgFieldContent & "Cannot import Ordered Item for row:" & lRowCounter & " "
'                lErrorsFieldContent = lErrorsFieldContent + 1
''                ''GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            End If
            
'            sFunctioncode = sItemArray(lRowCounter, 1)
'            sPROGYR = sItemArray(lRowCounter, dtCOLPos.lProgramYear)
'            If UCase(sItemArray(lRowCounter, dtCOLPos.lItemStatus)) = "DEV" Then
'               gbDEVITEM = True
'            Else
'                gbDEVITEM = False
'            End If
'            If Not IsNumeric(sPROGYR) Then sPROGYR = 0
        
'            If Not bValidateSaveArray(sFunctioncode, gsItem_Table, lRowCounter, sPROGYR, lProgNumber, sCustomerNumber, _
'                        sItemArray(lRowCounter, dtCOLPos.lVendorNumber), _
'                        sItemArray(lRowCounter, dtCOLPos.lFactoryNumber), bProductDevelopment, _
'                        sItemArray(), dtValidateITEMArray(), lItemFields, _
'                        sErrorMsgFieldContent, lErrorsFieldContent, _
'                        dtCOLArray(), lNbrCOLsToValidate) Then
'                sReturnErrormsg = "Could not validate the Item data in memory"
''                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            End If
            
'            If bPROPOSALFormIndicator = False Then
'                If Not bValidateSaveArray(sFunctioncode, gsItemSpecs_Table, lRowCounter, sPROGYR, lProgNumber, sCustomerNumber, _
'                        sItemArray(lRowCounter, dtCOLPos.lVendorNumber), _
'                        sItemArray(lRowCounter, dtCOLPos.lFactoryNumber), bProductDevelopment, _
'                        sItemSPECSArray(), dtValidateItemSPECSArray(), lItemSpecsFields, _
'                        sErrorMsgFieldContent, lErrorsFieldContent, _
'                        dtCOLArray(), lNbrCOLsToValidate) Then
'                    sReturnErrormsg = "Could not validate the ItemSpecs data in memory"
''                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If
'            End If
         
'            If Not bValidateSaveArray(sFunctioncode, gsItem_Assortments_Table, lRowCounter, sPROGYR, lProgNumber, sCustomerNumber, _
'                        sItemArray(lRowCounter, dtCOLPos.lVendorNumber), _
'                        sItemArray(lRowCounter, dtCOLPos.lFactoryNumber), bProductDevelopment, _
'                        sAssortmentArray(), dtValidateAssortmentArray(), lAssortmentFields, _
'                        sErrorMsgFieldContent, lErrorsFieldContent, _
'                        dtCOLArray(), lNbrCOLsToValidate) Then
'                sReturnErrormsg = "Could not validate the Assortment data in memory"
''                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            End If
'            If lMaxMaterialCols > 0 And bPROPOSALFormIndicator = False Then 'frmProposal Form has it's own validation
'                If bValidateMaterialInfo(lRowCounter, sErrorMsgFieldContent, lErrorsFieldContent) = False Then
'                    sReturnErrormsg = "Could not validate the ItemMaterial(Material1, Material2 etc) data in memory!"
''                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If
'            End If
'        End If
'        If lRowContentErrors < lErrorsFieldContent Then
'            sErrorMsgFieldContent = sErrorMsgFieldContent & vbCrLf
'        End If
'    Next lRowCounter

'    Debug.Print "Validating Core Fields on Spreadsheet Row " & Now()
        
'    ' Validate the Core Fields
'    sErrorMsgCoreFields = vbCrLf & "CORE FIELD VALIDATION:" & vbCrLf
'    For lRowCounter = glDATA_START_ROW To lRowsToValidate
'        'check if row errors have changed for a new row,, so as not to keep adding vbcrlf to end for rows that have no errors
'        Dim lRowCoreErrors As Long
'        lRowCoreErrors = lErrorsCoreFields
                
'        If Len(sItemArray(lRowCounter, glFunctionCode_ColPos)) = 0 Then
'            Call bUpdateStatusMessage(frmThis, "Skipping Row " & CStr(lRowCounter) & "...")
'        Else
'            If bPROPOSALFormIndicator = True Then
'                Call bUpdateStatusMessage(frmThis, "Validating Core Fields on Proposal Form ...")
'            Else
'                Call bUpdateStatusMessage(frmThis, "Validating Core Fields on Spreadsheet Row " & CStr(lRowCounter) & "...")
'            End If
       
'            If bValidateCoreFields(lRowCounter, lRowsToValidate, _
'                        sItemSPECSArray(), sItemArray(), sAssortmentArray(), lAssortmentFields, dtCOLPos, _
'                        sErrorMsgCoreFields, lErrorsCoreFields, sX_Technology(), sX_CertifiedPrinterNames()) = False Then
'                sReturnErrormsg = "Could not validate the Core Fields in memory"
''                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            End If
'        End If
'        If lRowCoreErrors < lErrorsCoreFields Then
'            sErrorMsgCoreFields = sErrorMsgCoreFields & vbCrLf
'        End If
'    Next
'    Application.DoEvents
'    bValidateData = True

'    Debug.Print "Leaving bValidateData " & Now()
    
'ExitRoutine:
''    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'    ValidationCleanup   ' drop all recordsets used by bValidateField (called by ValidateSaveArray)
'    Exit Function
'ErrorHandler:
    
'    MsgBox Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateData"
'    If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'        smessage = "In bValidateData , Err Number " & Err.Number & "Error Description: " & Err.Description
'        If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'        End If
'    End If
'    Resume ExitRoutine
'    End Function

'    Public Function bValidateMaterialInfo(ByVal lRow As Long _
'    , ByRef sErrorMsgFieldContent As String, ByRef lErrorsFieldContent As Long) As Boolean

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        'Check that percentages add up to 100%, and no duplicate material names, for each row
'        Dim lCommaPos As Long
'        Dim lDecimalPos As Long
'        Dim lLeftOfDecimal As Long
'        Dim lRightOfDecimal As Long

'        Dim lMaterialCounter As Long
'        Dim sMaterialValues As String

'        Dim dTotalMaterialPct As Double
'        Dim sMaterialPctX As String
'        Dim sMaterialPcts As String

'        Dim dTotalCostPct As Double
'        Dim sCostPctX As String
'        Dim sCostPcts As String

'        bValidateMaterialInfo = False
'        For lMaterialCounter = 1 To lMaxMaterialCols

'            sMaterialValues = Trim(SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sImportConcatenatedMaterial)
'            If sMaterialValues <> "" Then
'                lCommaPos = InStr(sMaterialValues, ",")
'                If lCommaPos > 0 Then
'                    sMaterialPctX = Trim(Left(sMaterialValues, lCommaPos - 1))
'                    If IsNumeric(sMaterialPctX) Then
'                        lDecimalPos = InStr(1, sMaterialPctX, ".")
'                        If lDecimalPos > 0 Then
'                            lLeftOfDecimal = Len(Left(sMaterialPctX, lDecimalPos - 1))
'                            lRightOfDecimal = Len(sMaterialPctX) - lLeftOfDecimal - 1
'                            If lRightOfDecimal > 2 Then
'                                sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": MaterialPct" & lMaterialCounter & " CANNOT have more than 2 decimals: " & sMaterialPctX & vbCrLf
'                                lErrorsFieldContent = lErrorsFieldContent + 1
'                            End If
'                        End If
'                        sMaterialPcts = sMaterialPcts & sMaterialPctX & "+"
'                        dTotalMaterialPct = dTotalMaterialPct + CDec(sMaterialPctX)
'                        SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialPct = sMaterialPctX
'                    Else
'                        If sMaterialPctX <> "" Then
'                            sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": MaterialPct" & lMaterialCounter & " MUST be numeric: " & sMaterialPctX & vbCrLf
'                            lErrorsFieldContent = lErrorsFieldContent + 1
'                        End If
'                    End If
'                End If
'                sMaterialValues = Microsoft.VisualBasic.Right(sMaterialValues, Len(sMaterialValues) - lCommaPos)
'                lCommaPos = InStr(sMaterialValues, ",")
'                If lCommaPos > 0 Then
'                    sCostPctX = Trim(Mid(sMaterialValues, lCommaPos + 1, Len(sMaterialValues) - lCommaPos))
'                    If IsNumeric(sCostPctX) Then
'                        lDecimalPos = InStr(1, sCostPctX, ".")
'                        If lDecimalPos > 0 Then
'                            lLeftOfDecimal = Len(Left(sCostPctX, lDecimalPos - 1))
'                            lRightOfDecimal = Len(sCostPctX) - lLeftOfDecimal - 1
'                            If lRightOfDecimal > 2 Then
'                                sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": CostPct" & lMaterialCounter & " CANNOT have more than 2 decimals: " & sCostPctX & vbCrLf
'                                lErrorsFieldContent = lErrorsFieldContent + 1
'                            End If
'                        End If
'                        sCostPcts = sCostPcts & sCostPctX & "+"
'                        dTotalCostPct = dTotalCostPct + CDec(sCostPctX)
'                        SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sCostPct = sCostPctX
'                    Else
'                        If sCostPctX <> "" Then
'                            sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": CostPct" & lMaterialCounter & " MUST be numeric: " & sCostPctX & vbCrLf
'                            lErrorsFieldContent = lErrorsFieldContent + 1
'                        End If
'                    End If
'                End If
'                If lCommaPos > 0 Then
'                    SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName = Microsoft.VisualBasic.Left(sMaterialValues, lCommaPos - 1)
'                    SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName = Trim(SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName)

'                    If Len(SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName) > 30 Then      '01/11/2008
'                        sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": MaterialName cannot be more than 30 characters[ " & _
'                                    SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName & "] -length:" & Len(SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName) & vbCrLf
'                        lErrorsFieldContent = lErrorsFieldContent + 1
'                    End If

'                    'check for duplicate material names
'                    Dim lMaterialX As Long
'                    For lMaterialX = 1 To lMaxMaterialCols
'                        If lMaterialX <> lMaterialCounter Then
'                            If SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName = SpreadsheetMaterialValuesX(lRow, lMaterialX).sMaterialName Then
'                                sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": CANNOT have duplicate MaterialName: " & _
'                                    SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName & vbCrLf
'                                lErrorsFieldContent = lErrorsFieldContent + 1
'                            End If
'                            If InStr(1, SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName, ",") = True Then
'                                sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": MaterialName can't contain commas: " & _
'                                    SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName & vbCrLf
'                                lErrorsFieldContent = lErrorsFieldContent + 1

'                            End If
'                        End If

'                    Next lMaterialX
'                Else
'                    SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sMaterialName = ""
'                    sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": MaterialName" & lMaterialCounter & " CANNOT be blank" & vbCrLf
'                    lErrorsFieldContent = lErrorsFieldContent + 1
'                End If
'            End If

'        Next lMaterialCounter
'        If sMaterialPcts = "" And sCostPcts = "" Then
'        Else
'            If dTotalMaterialPct <> 100 And dTotalMaterialPct <> 0 Then
'                If Len(sMaterialPcts) > 1 Then
'                    sMaterialPcts = Microsoft.VisualBasic.Left(sMaterialPcts, Len(sMaterialPcts) - 1)
'                End If
'                sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": Total MaterialPct (" & sMaterialPcts & "= " & dTotalMaterialPct & ") UNEQUAL TO 100%" & vbCrLf
'                lErrorsFieldContent = lErrorsFieldContent + 1
'            End If
'            If dTotalCostPct <> 100 And dTotalCostPct <> 0 Then
'                If Len(sCostPcts) > 1 Then
'                    sCostPcts = Microsoft.VisualBasic.Left(sCostPcts, Len(sCostPcts) - 1)
'                End If
'                sErrorMsgFieldContent = sErrorMsgFieldContent & "ROW " & lRow & ": Total CostPct (" & sCostPcts & "= " & dTotalCostPct & ") UNEQUAL TO 100%" & vbCrLf
'                lErrorsFieldContent = lErrorsFieldContent + 1
'            End If
'        End If
'        Application.DoEvents()
'        bValidateMaterialInfo = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateMaterialInfo")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidateMaterialInfo , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function
        
'    Public Function bValidateField(ByVal sFunctioncode As String, ByVal sFieldName As String, ByVal vFieldData As Object, ByVal sCustomerNumber As String, ByVal sFactoryNumber As String, ByVal bImport As Boolean, ByVal sFieldValidation As String, ByRef sValidationErrorMsg As String) As Object
'        'new 11/08/2007 added sCustomerNumber above
'        '  Added sFieldname as parameter to be able to suppress validating some fields for HONGKONG, HK1 ?
'        '  when editing and those Fields are Hidden on form,
'        '   and therefore cannot be corrected if too many decimals for instance, etc
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim lMaxStringLength As Long
'        Dim lMinNumber As Long
'        Dim lMaxNumber As Long

'        Dim dblMinNumber As Double
'        Dim dblMaxNumber As Double     'new 11/08/2007

'        Dim lPrecision As Long
'        Dim lScale As Long

'        Dim lDecimalPos As Long
'        Dim lLeftOfDecimal As Long
'        Dim lRightOfDecimal As Long

'        Dim sValidationType As String
'        Dim sFieldValidRange As String
'        Dim bIgnoreDecimalValidation As Boolean
'        Dim sActiveMsg As String
'        Dim sDept As String    '2012/07/11
'        Dim sClass As String    '2012/07/11
'        Dim sCustItemNum As String    '2012/07/11
'        Dim sLookupTable As String
'        Dim rsLookupTable As ADODB.Recordset
'        Dim sTemp1 As String

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally

'        bValidateField = True

'        sValidationErrorMsg = "" 'new 10/27/2007

'        ' If there is an error in the cell, don't bother validating the data against the type
'        If vFieldData = gsDENOTE_CELL_ERROR Then
'            sValidationErrorMsg = "Spreadsheet cell error"
'        Else

'            sValidationType = Microsoft.VisualBasic.Left(sFieldValidation, 7)

'            Select Case sValidationType
'                Case gsFORMATC
'                    Application.DoEvents()

'                Case gsVARCHAR
'                    lMaxStringLength = Mid(sFieldValidation, InStr(1, sFieldValidation, "(") + 1, _
'                            InStr(1, sFieldValidation, ")") - InStr(1, sFieldValidation, "(") - 1)

'                    If Len(vFieldData) > lMaxStringLength Then
'                        sValidationErrorMsg = vFieldData & " is more than " & lMaxStringLength & " characters"
'                        '                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If

'                    ' check for more than 3 linefeed characters in LongDescription, for Crystal Reports
'                    ' this is also so that labels are not printed incorrectly because of multiple line feeds
'                    If sValidationErrorMsg = "" And sFieldName = gsCOL_LONGDESC Then
'                        If LineCount(vFieldData) >= 4 Then
'                            sValidationErrorMsg = sFieldName & " More >= 4 lines, remove LineFeed character; need less for Crystal Reports."
'                            '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        End If
'                    End If

'                    '2012/07/11 - For CustomerNumbers: 101,102,235,206,103,888,248 -check that
'                    'CustomerItemNumber has length of 11, to calculate CustomerCartonUPC in the Import Process
'                    If sValidationErrorMsg = "" And sFieldName = "CustomerItemNumber" Then
'                        Select Case sCustomerNumber
'                            Case 101, 102, 235, 206, 103, 888, 248
'                                If Not IsBlank(vFieldData) Then
'                                    If Len(vFieldData) <> 11 Then
'                                        sValidationErrorMsg = sFieldName & ": Length is not 11; can't calculate CustomerCartonUPCNumber" & vbCrLf & _
'                                        "Enter as Format: XXX XX XXXX"
'                                        '                                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                    End If

'                                    If sValidationErrorMsg = "" Then
'                                        sDept = Microsoft.VisualBasic.Left(vFieldData, 3)
'                                        sClass = Mid(vFieldData, 5, 2)
'                                        sCustItemNum = Microsoft.VisualBasic.Right(vFieldData, 4)
'                                        If Not (IsNumeric(sDept) And IsNumeric(sClass) And IsNumeric(sCustItemNum)) Then
'                                            sValidationErrorMsg = sFieldName & ": not numeric, can't calculate CustomerCartonUPCNumber"
'                                            '                                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                        End If
'                                    End If
'                                End If
'                            Case Else   ' do nothing
'                                '2013/05/28 -HN-for FunctionCode = A , cant carry this CustomerCartonUPC number
'                        End Select
'                    End If

'                    'For CustomerNumber = 104, check that OtherText has length of 11,to calculate CustomerCartonUPC
'                    'in the Import Process
'                    If sValidationErrorMsg = "" And sFieldName = "OtherText" Then
'                        If sCustomerNumber = 104 Then
'                            If Not IsBlank(vFieldData) Then
'                                If Len(vFieldData) <> 11 Then
'                                    sValidationErrorMsg = sFieldName & ": Length is not 11; can't calculate CustomerCartonUPCNumber" & vbCrLf & _
'                                    "Enter as Format: XXX XX XXXX"
'                                    '                                'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                End If

'                                If sValidationErrorMsg = "" Then
'                                    sDept = Microsoft.VisualBasic.Left(vFieldData, 3)
'                                    sClass = Mid(vFieldData, 5, 2)
'                                    sCustItemNum = Microsoft.VisualBasic.Right(vFieldData, 4)
'                                    If Not (IsNumeric(sDept) And IsNumeric(sClass) And IsNumeric(sCustItemNum)) Then
'                                        sValidationErrorMsg = sFieldName & ": not numeric, can't calculate CustomerCartonUPCNumber"
'                                        '                                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                    End If
'                                End If
'                            End If
'                        End If
'                    End If
'                    '2012/07/11 --------------------------------


'                Case gsNUMERIC
'                    lMinNumber = Mid(sFieldValidation, InStr(1, sFieldValidation, "(") + 1, _
'                            InStr(1, sFieldValidation, "-") - InStr(1, sFieldValidation, "(") - 1)
'                    lMaxNumber = Mid(sFieldValidation, InStr(1, sFieldValidation, "-") + 1, _
'                            InStr(1, sFieldValidation, ")") - InStr(1, sFieldValidation, "-") - 1)

'                    If Not IsNumeric(vFieldData) Then
'                        sValidationErrorMsg = vFieldData & " should be numeric but is not"
'                        '                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If

'                    If sValidationErrorMsg = "" And (vFieldData < lMinNumber Or vFieldData > lMaxNumber) Then
'                        sValidationErrorMsg = vFieldData & " is not between " & lMinNumber & " and " & lMaxNumber
'                        '                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If

'                Case gsDECIMAL
'                    bIgnoreDecimalValidation = False
'                    Select Case msUserGroup
'                        Case msSALES, msHK1, msHONGKONG, msPHOTO, msSHIP, msCREATIVE, msPRODDEV
'                            Select Case sFieldName
'                                Case gsCOL_FactoryFOBCost, gsCOL_FactoryFCACost, gsCOL_FOBSellPrice, gsCOL_NetFirstCost, _
'                                    gsCOL_NetFirstCostALT, gsCOL_LEDLightCost, gsCOL_StoreCost, gsCOL_StoreCostALT, _
'                                    gsCOl_StorageStoreCost, gsCOL_REGLinePrice, gsCOL_LYRFOBCost, gsCOL_LYRFOBSellPrice, _
'                                    gsCOL_SSCLineAddOn, gsCOL_AltSellPrice, gsCOL_DomELC, gsCOL_AltCost, _
'                                    gsCOL_FCASellPrice, gsCOL_LYRFCACost, gsCOL_LYRFCASellPrice, gsCOL_LYRAltCost, _
'                                    gsCOL_LYRAltSellPrice                                     '2010/09/14
'                                    bIgnoreDecimalValidation = True
'                                Case Else
'                                    bIgnoreDecimalValidation = False
'                            End Select
'                    End Select

'                    'ignore decimal validation for hidden fields, user cannot correct them if error on Proposal Form
'                    If Not bIgnoreDecimalValidation Then
'                        ' ET 2012-12-11 - the ";" indicates the separation of the numeric fields and value range.
'                        ' why are we using the return as a Boolean?
'                        If InStr(sFieldValidation, ";") = False Then
'                            lPrecision = Mid(sFieldValidation, InStr(1, sFieldValidation, "(") + 1, _
'                                InStr(1, sFieldValidation, ",") - InStr(1, sFieldValidation, "(") - 1)
'                            lScale = Mid(sFieldValidation, InStr(1, sFieldValidation, ",") + 1, _
'                                InStr(1, sFieldValidation, ")") - InStr(1, sFieldValidation, ",") - 1)
'                        Else        '-----changes made to test for range in decimal values too .......
'                            lPrecision = Mid(sFieldValidation, InStr(1, sFieldValidation, "(") + 1, _
'                                InStr(1, sFieldValidation, ",") - InStr(1, sFieldValidation, "(") - 1)
'                            lScale = Mid(sFieldValidation, InStr(1, sFieldValidation, ",") + 1, _
'                                    InStr(1, sFieldValidation, ";") - (InStr(1, sFieldValidation, ",") + 1))

'                            sFieldValidRange = "(" & Microsoft.VisualBasic.Right(sFieldValidation, InStr(1, sFieldValidation, ")") - InStr(1, sFieldValidation, ";"))
'                            'new 11/07/2007 changed below to double
'                            dblMinNumber = Mid(sFieldValidRange, InStr(1, sFieldValidRange, "(") + 1, _
'                                InStr(1, sFieldValidRange, "-") - InStr(1, sFieldValidRange, "(") - 1)
'                            dblMaxNumber = Mid(sFieldValidRange, InStr(1, sFieldValidRange, "-") + 1, _
'                                InStr(1, sFieldValidRange, ")") - InStr(1, sFieldValidRange, "-") - 1)
'                            If vFieldData < dblMinNumber Or vFieldData > dblMaxNumber Then '11/08/2007 changed to double
'                                sValidationErrorMsg = vFieldData & " is not between " & dblMinNumber & " and " & dblMaxNumber
'                                '                            'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        End If

'                        If sValidationErrorMsg = "" Then
'                            If Not IsNumeric(vFieldData) Then
'                                sValidationErrorMsg = vFieldData & " is not numeric"
'                                '                            'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If
'                        End If

'                        If sValidationErrorMsg = "" Then
'                            lDecimalPos = InStr(1, vFieldData, ".")
'                            If lDecimalPos = 0 Then
'                                If Len(vFieldData) > (lPrecision - lScale) Then
'                                    sValidationErrorMsg = vFieldData & " has more than " & CStr(lPrecision - lScale) & " digits"
'                                    '                                'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                End If
'                            Else
'                                lLeftOfDecimal = Len(Left(vFieldData, lDecimalPos - 1))
'                                lRightOfDecimal = Len(vFieldData) - lLeftOfDecimal - 1

'                                If lLeftOfDecimal > (lPrecision - lScale) And lRightOfDecimal > lScale Then
'                                    sValidationErrorMsg = vFieldData & " has more than " & CStr(lPrecision - lScale) & " " & _
'                                            "digits to the left of the decimal point and more than " & CStr(lScale) & " " & _
'                                            "digits to the right of the decimal point"
'                                    '                                'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                ElseIf lLeftOfDecimal > (lPrecision - lScale) Then
'                                    sValidationErrorMsg = vFieldData & " has more than " & CStr(lPrecision - lScale) & " " & _
'                                            "digits to the left of the decimal point"
'                                    '                                'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                ElseIf lRightOfDecimal > lScale Then
'                                    sValidationErrorMsg = vFieldData & " has more than " & CStr(lScale) & " " & _
'                                            "digits to the right of the decimal point"
'                                    '                                'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                End If
'                            End If
'                        End If
'                    End If

'                Case gsBOOLEAN
'                    '                Select Case sFieldName          '2013/05/15
'                    '                    Case "CargoReceipt", "ChildLabor", "CommercialInvoice", "GCC", "PackingList", "PrisonLaborDoc"
'                    '                        vFieldData = "TRUE"
'                    '                End Select
'                    If UCase(vFieldData) <> "TRUE" And UCase(vFieldData) <> "FALSE" Then
'                        sValidationErrorMsg = vFieldData & " is not TRUE or FALSE"
'                        '                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If

'                    '03/19/2008
'                Case gsYESNONULL
'                    If UCase(vFieldData) <> "YES" And UCase(vFieldData) <> "NO" Then
'                        sValidationErrorMsg = vFieldData & " is not YES/NO/Blank"
'                        '                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If

'                Case gsDATETYM
'                    If Not IsDate(vFieldData) Then
'                        sValidationErrorMsg = vFieldData & " is not a valid datetime"
'                        '                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If

'                Case gsITEMNBR
'                    If vFieldData <> gsNEW_ITEM_NBR Then
'                        If Len(vFieldData) <> 10 Then
'                            sValidationErrorMsg = vFieldData & " is not a valid ItemNumber (not 10 characters in length)"
'                            '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        ElseIf Not IsNumeric(Left(vFieldData, 6)) Then
'                            sValidationErrorMsg = vFieldData & " is not a valid ItemNumber (left 6 characters not numeric)"
'                            '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        ElseIf Not IsNumeric(Right(vFieldData, 3)) Then
'                            sValidationErrorMsg = vFieldData & " is not a valid ItemNumber (right 3 characters not numeric)"
'                            '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        ElseIf Mid(vFieldData, 7, 1) <> "-" Then
'                            sValidationErrorMsg = vFieldData & " is not a valid ItemNumber (dash separator not found)"
'                            '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        End If
'                    End If

'                Case gsLOOKUPS
'                    sLookupTable = Mid(sFieldValidation, InStr(1, sFieldValidation, "(") + 1, _
'                            InStr(1, sFieldValidation, ")") - InStr(1, sFieldValidation, "(") - 1)

'                    sSQL = "SELECT * FROM " & sLookupTable & " "
'                    sActiveMsg = ""

'                    Select Case sLookupTable
'                        Case gsItem_Table
'                            sSQL = sSQL & "WHERE ItemNumber = " & sAddQuotes(vFieldData)

'                        Case "ItemStatusCodes"
'                            If rsLookupItemStatusCodes Is Nothing Then
'                                sSQL = "SELECT Upper(ItemStatusCode) as ItemStatusCode FROM " & sLookupTable
'                                rsLookupItemStatusCodes = New ADODB.Recordset
'                                rsLookupItemStatusCodes.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupItemStatusCodes.MoveFirst()
'                            rsLookupItemStatusCodes.Find "ItemStatusCode = " & sTemp1

'                            If rsLookupItemStatusCodes.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Category"
'                            If rsLookupCategory Is Nothing Then
'                                sSQL = "SELECT Upper(CategoryCode) as CategoryCode, " & _
'                                            "convert(varchar, InactiveProgramYear) as InactiveProgramYear FROM Category"
'                                rsLookupCategory = New ADODB.Recordset
'                                rsLookupCategory.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            rsLookupCategory.MoveFirst()
'                            rsLookupCategory.Find "CategoryCode = " & sAddQuotes(vFieldData)

'                            If rsLookupCategory.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                            'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            Else
'                                If Not IsNull(rsLookupCategory!InactiveProgramYear) _
'                                        And rsLookupCategory!InactiveProgramYear <= CInt(sProgramYR) Then
'                                    sActiveMsg = " OR " & sLookupTable & ".InactiveProgramYear: " & _
'                                                rsLookupCategory!InactiveProgramYear & " <= Item.ProgramYear:" & sProgramYR
'                                    '                            sValidationErrorMsg = vFieldData & " not in " & sLookupTable & " table" & sActiveMsg          '11/19/2008
'                                    sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & vbCrLf & "        " & sActiveMsg
'                                    '                                'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                End If
'                            End If

'                        Case "Customer"     ' use old way
'                            sSQL = sSQL & "WHERE convert(varchar,CustomerNumber)= " & sAddQuotes(vFieldData)

'                        Case "Program"      ' use old way
'                            If Not IsNumeric(vFieldData) Then
'                                sValidationErrorMsg = "Program: " & vFieldData & " is not numeric"
'                                '                            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            Else
'                                sSQL = sSQL & "WHERE ProgramNumber = " & vFieldData
'                            End If

'                        Case "Grade"
'                            If rsLookupGrade Is Nothing Then
'                                sSQL = "SELECT Grade FROM " & sLookupTable
'                                rsLookupGrade = New ADODB.Recordset
'                                rsLookupGrade.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            ' sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupGrade.MoveFirst()
'                            rsLookupGrade.Find "Grade = " & vFieldData

'                            If rsLookupGrade.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Licensor"     ' use old way
'                            sSQL = sSQL & "WHERE convert(varchar,Licensor) = " & sAddQuotes(vFieldData)    'new 10/25/2007 in case Spreadsheet value is not numeric

'                        Case "SubProgram"   ' use old way
'                            If Not IsNumeric(vFieldData) Then
'                                sValidationErrorMsg = "SubProgram: " & vFieldData & " is not numeric"
'                                '                            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            Else
'                                sSQL = sSQL & "WHERE SubProgram = " & vFieldData
'                            End If

'                            '                Case "Technology"  'needs different routine, values are concatenated;  check valid values.....
'                            '                    sSQL = sSQL & "WHERE Technology = " & vFieldData

'                        Case "Factory"      ' use old way
'                            sSQL = sSQL & "WHERE Upper(FactoryNumber) = " & UCase(sAddQuotes(vFieldData))

'                        Case "Vendor"       ' use old way
'                            sSQL = sSQL & "WHERE VendorNumber = " & vFieldData

'                        Case "Season"
'                            If rsLookupSeason Is Nothing Then
'                                sSQL = "SELECT upper(SeasonCode) as SeasonCode FROM " & sLookupTable
'                                rsLookupSeason = New ADODB.Recordset
'                                rsLookupSeason.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupSeason.MoveFirst()
'                            rsLookupSeason.Find "SeasonCode = " & sTemp1

'                            If rsLookupSeason.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "FOBPoints"
'                            If rsLookupFOBPoints Is Nothing Then
'                                sSQL = "SELECT Upper(FOBPoint) as FOBPoint FROM " & sLookupTable
'                                rsLookupFOBPoints = New ADODB.Recordset
'                                rsLookupFOBPoints.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupFOBPoints.MoveFirst()
'                            rsLookupFOBPoints.Find "FOBPoint = " & sTemp1

'                            If rsLookupFOBPoints.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Country"
'                            '                        sSQL = sSQL & "WHERE Upper(Country) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupCountry Is Nothing Then
'                                sSQL = "SELECT Upper(Country) as Country FROM " & sLookupTable
'                                rsLookupCountry = New ADODB.Recordset
'                                rsLookupCountry.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupCountry.MoveFirst()
'                            rsLookupCountry.Find "Country = " & sTemp1

'                            If rsLookupCountry.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "PackageTypes"     ' use old way
'                            sSQL = sSQL & "WHERE Upper(Description) = " & UCase(sAddQuotes(vFieldData))

'                        Case "Ele_Connection"
'                            '                        sSQL = sSQL & "Where Upper(Ele_Connection) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_Connection Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_Connection) as Ele_Connection FROM " & sLookupTable
'                                rsLookupEle_Connection = New ADODB.Recordset
'                                rsLookupEle_Connection.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_Connection.MoveFirst()
'                            rsLookupEle_Connection.Find "Ele_Connection = " & sTemp1

'                            If rsLookupEle_Connection.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_IndoorOrIndoorOutdoor"
'                            '                        sSQL = sSQL & "WHERE Upper(Ele_IndoororIndoorOutdoor) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_IndoorOrIndoorOutdoor Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_IndoororIndoorOutdoor) as Ele_IndoororIndoorOutdoor FROM " & sLookupTable
'                                rsLookupEle_IndoorOrIndoorOutdoor = New ADODB.Recordset
'                                rsLookupEle_IndoorOrIndoorOutdoor.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_IndoorOrIndoorOutdoor.MoveFirst()
'                            rsLookupEle_IndoorOrIndoorOutdoor.Find "Ele_IndoororIndoorOutdoor = " & sTemp1

'                            If rsLookupEle_IndoorOrIndoorOutdoor.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_TrayPlasticPaper"
'                            '                       sSQL = sSQL & "WHERE Upper(Ele_TrayPlasticPaper) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_TrayPlasticPaper Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_TrayPlasticPaper) as Ele_TrayPlasticPaper FROM " & sLookupTable
'                                rsLookupEle_TrayPlasticPaper = New ADODB.Recordset
'                                rsLookupEle_TrayPlasticPaper.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_TrayPlasticPaper.MoveFirst()
'                            rsLookupEle_TrayPlasticPaper.Find "Ele_TrayPlasticPaper = " & sTemp1

'                            If rsLookupEle_TrayPlasticPaper.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_FuseRating"
'                            '                        sSQL = sSQL & "WHERE Upper(Ele_FuseRating) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_FuseRating Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_FuseRating) as Ele_FuseRating FROM " & sLookupTable
'                                rsLookupEle_FuseRating = New ADODB.Recordset
'                                rsLookupEle_FuseRating.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_FuseRating.MoveFirst()
'                            rsLookupEle_FuseRating.Find "Ele_FuseRating = " & sTemp1

'                            If rsLookupEle_FuseRating.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_TrayOrBulk"
'                            '                        sSQL = sSQL & "WHERE Upper(Ele_TrayOrBulk) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_TrayOrBulk Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_TrayOrBulk) as Ele_TrayOrBulk FROM " & sLookupTable
'                                rsLookupEle_TrayOrBulk = New ADODB.Recordset
'                                rsLookupEle_TrayOrBulk.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_TrayOrBulk.MoveFirst()
'                            rsLookupEle_TrayOrBulk.Find "Ele_TrayOrBulk = " & sTemp1

'                            If rsLookupEle_TrayOrBulk.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_WireGauge"
'                            '                        sSQL = sSQL & "WHERE Upper(Ele_WireGauge) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_WireGauge Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_WireGauge) as Ele_WireGauge FROM " & sLookupTable
'                                rsLookupEle_WireGauge = New ADODB.Recordset
'                                rsLookupEle_WireGauge.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_WireGauge.MoveFirst()
'                            rsLookupEle_WireGauge.Find "Ele_WireGauge = " & sTemp1

'                            If rsLookupEle_WireGauge.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_LampBaseType"
'                            '                        sSQL = sSQL & "WHERE Upper(Ele_LampBaseType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_LampBaseType Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_LampBaseType) as Ele_LampBaseType FROM " & sLookupTable
'                                rsLookupEle_LampBaseType = New ADODB.Recordset
'                                rsLookupEle_LampBaseType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_LampBaseType.MoveFirst()
'                            rsLookupEle_LampBaseType.Find "Ele_LampBaseType = " & sTemp1

'                            If rsLookupEle_LampBaseType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_LampBrightness"
'                            '                        sSQL = sSQL & "WHERE Upper(Ele_LampBrightness) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_LampBrightness Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_LampBrightness) as Ele_LampBrightness FROM " & sLookupTable
'                                rsLookupEle_LampBrightness = New ADODB.Recordset
'                                rsLookupEle_LampBrightness.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_LampBrightness.MoveFirst()
'                            rsLookupEle_LampBrightness.Find "Ele_LampBrightness = " & sTemp1

'                            If rsLookupEle_LampBrightness.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "CertificationMark"
'                            '                        sSQL = sSQL & "WHERE Upper(CertificationMark) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupCertificationMark Is Nothing Then
'                                sSQL = "SELECT Upper(CertificationMark) as CertificationMark, InactiveProgramYear FROM " & sLookupTable
'                                rsLookupCertificationMark = New ADODB.Recordset
'                                rsLookupCertificationMark.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupCertificationMark.MoveFirst()
'                            rsLookupCertificationMark.Find "CertificationMark = " & sTemp1

'                            If rsLookupCertificationMark.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                            'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            Else
'                                If Not IsNull(rsLookupCertificationMark!InactiveProgramYear) _
'                                            And rsLookupCertificationMark!InactiveProgramYear <= CInt(sProgramYR) Then
'                                    sActiveMsg = " OR " & sLookupTable & ".InactiveProgramYear: " & _
'                                               rsLookupCertificationMark!InactiveProgramYear & " <= Item.ProgramYear:" & _
'                                               sProgramYR
'                                    sValidationErrorMsg = vFieldData & " not in " & sLookupTable & " table" & vbCrLf & _
'                                               "        " & sActiveMsg
'                                    '                                'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                Else
'                                    sActiveMsg = ""
'                                End If
'                            End If

'                        Case "CertificationType"
'                            '                        sSQL = sSQL & "WHERE Upper(CertificationType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupCertificationType Is Nothing Then
'                                sSQL = "SELECT Upper(CertificationType) as CertificationType, InactiveProgramYear FROM " & sLookupTable
'                                rsLookupCertificationType = New ADODB.Recordset
'                                rsLookupCertificationType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupCertificationType.MoveFirst()
'                            rsLookupCertificationType.Find "CertificationType = " & sTemp1

'                            If rsLookupCertificationType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                            'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            Else
'                                If Not IsNull(rsLookupCertificationType!InactiveProgramYear) _
'                                            And rsLookupCertificationType!InactiveProgramYear <= CInt(sProgramYR) Then
'                                    sActiveMsg = " OR " & sLookupTable & ".InactiveProgramYear: " & _
'                                               rsLookupCertificationType!InactiveProgramYear & " <= Item.ProgramYear:" & _
'                                               sProgramYR
'                                    sValidationErrorMsg = vFieldData & " not in " & sLookupTable & " table" & vbCrLf & _
'                                               "        " & sActiveMsg
'                                    '                                 'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                Else
'                                    sActiveMsg = ""
'                                End If
'                            End If

'                        Case "Bag_PaperType"
'                            '                        sSQL = sSQL & "WHERE Upper(Bag_PaperType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupBag_PaperType Is Nothing Then
'                                sSQL = "SELECT Upper(Bag_PaperType) as Bag_PaperType FROM " & sLookupTable
'                                rsLookupBag_PaperType = New ADODB.Recordset
'                                rsLookupBag_PaperType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupBag_PaperType.MoveFirst()
'                            rsLookupBag_PaperType.Find "Bag_PaperType = " & sTemp1

'                            If rsLookupBag_PaperType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Bag_PrintType"
'                            '                        sSQL = sSQL & "WHERE Upper(Bag_PrintType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupBag_PrintType Is Nothing Then
'                                sSQL = "SELECT Upper(Bag_PrintType) as Bag_PrintType FROM " & sLookupTable
'                                rsLookupBag_PrintType = New ADODB.Recordset
'                                rsLookupBag_PrintType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupBag_PrintType.MoveFirst()
'                            rsLookupBag_PrintType.Find "Bag_PrintType = " & sTemp1

'                            If rsLookupBag_PrintType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Bag_Finish"
'                            '                        sSQL = sSQL & "WHERE Upper(Bag_Finish) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupBag_Finish Is Nothing Then
'                                sSQL = "SELECT Upper(Bag_Finish) as Bag_Finish FROM " & sLookupTable
'                                rsLookupBag_Finish = New ADODB.Recordset
'                                rsLookupBag_Finish.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupBag_Finish.MoveFirst()
'                            rsLookupBag_Finish.Find "Bag_Finish = " & sTemp1

'                            If rsLookupBag_Finish.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Bag_HandleType"
'                            '                        sSQL = sSQL & "WHERE Upper(Bag_HandleType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupBag_HandleType Is Nothing Then
'                                sSQL = "SELECT Upper(Bag_HandleType) as Bag_HandleType FROM " & sLookupTable
'                                rsLookupBag_HandleType = New ADODB.Recordset
'                                rsLookupBag_HandleType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupBag_HandleType.MoveFirst()
'                            rsLookupBag_HandleType.Find "Bag_HandleType = " & sTemp1

'                            If rsLookupBag_HandleType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                            '2009/11/19 - hn - commented below, SpecialEffects now concatenated field
'                            '                Case "Bag_SpecialEffects"
'                            '                    SQL = SQL & "WHERE Upper(Bag_SpecialEffects) = " & UCase(sAddQuotes(vFieldData))

'                        Case "Tree_Construction"
'                            '                        sSQL = sSQL & "WHERE Upper(Tree_Construction) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupTree_Construction Is Nothing Then
'                                sSQL = "SELECT Upper(Tree_Construction) as Tree_Construction FROM " & sLookupTable
'                                rsLookupTree_Construction = New ADODB.Recordset
'                                rsLookupTree_Construction.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupTree_Construction.MoveFirst()
'                            rsLookupTree_Construction.Find "Tree_Construction = " & sTemp1

'                            If rsLookupTree_Construction.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Tree_LightConstruction"
'                            '                        sSQL = sSQL & "WHERE Upper(Tree_LightConstruction) = " & UCase(sAddQuotes(vFieldData))
'                            '                                If vFieldData = "FALSE" And sLookupTable = "Tree_LightConstruction" Then
'                            '                                    vFieldData = ""
'                            ' translate old boolean values on spreadsheet later
'                            If rsLookupTree_LightConstruction Is Nothing Then
'                                sSQL = "SELECT Upper(Tree_LightConstruction) as Tree_LightConstruction FROM " & sLookupTable
'                                rsLookupTree_LightConstruction = New ADODB.Recordset
'                                rsLookupTree_LightConstruction.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupTree_LightConstruction.MoveFirst()
'                            rsLookupTree_LightConstruction.Find "Tree_LightConstruction = " & sTemp1

'                            If rsLookupTree_LightConstruction.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "SalesRep"
'                            If Len(vFieldData) < 2 And IsNumeric(vFieldData) Then    'pack with leading zeros  ... change in SQL version
'                                vFieldData = "0" & vFieldData
'                            End If
'                            '                        sSQL = sSQL & "WHERE SalesRepNumber = " & vFieldData '& " AND Inactive = 0 "
'                            'The SalesRepNumber Inactive condition commented out 4/4/07 by Gary following
'                            'a discussion with Theresa where she said to suppress this validation for now.
'                            'She's going to look into enhancing the sales rep info to track inactive date
'                            'and then this validation may possibly computer InactiveDate to some other date
'                            'which Theresa will determine after talking to finance.

'                            If rsLookupSalesRep Is Nothing Then
'                                sSQL = "SELECT SalesRepNumber FROM " & sLookupTable
'                                rsLookupSalesRep = New ADODB.Recordset
'                                rsLookupSalesRep.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            ' sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupSalesRep.MoveFirst()
'                            rsLookupSalesRep.Find "SalesRepNumber = " & vFieldData ' & " AND Inactive = 0 "

'                            If rsLookupSalesRep.EOF Then
'                                sActiveMsg = " , OR INACTIVE!"
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "LightType"
'                            '                        sSQL = sSQL & "WHERE Upper(LightType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupLightType Is Nothing Then
'                                sSQL = "SELECT Upper(LightType) as LightType FROM " & sLookupTable
'                                rsLookupLightType = New ADODB.Recordset
'                                rsLookupLightType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupLightType.MoveFirst()
'                            rsLookupLightType.Find "LightType = " & sTemp1

'                            If rsLookupLightType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                            '03/13/2008 new lookup tables follow:
'                        Case "BatteryType"
'                            '                        sSQL = sSQL & "WHERE Upper(BatteryType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupBatteryType Is Nothing Then
'                                sSQL = "SELECT Upper(BatteryType) as BatteryType FROM " & sLookupTable
'                                rsLookupBatteryType = New ADODB.Recordset
'                                rsLookupBatteryType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupBatteryType.MoveFirst()
'                            rsLookupBatteryType.Find "BatteryType = " & sTemp1

'                            If rsLookupBatteryType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_PlugTypeStackableStd"
'                            '                        sSQL = sSQL & "WHERE Upper(Ele_PlugTypeStackableStd) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_PlugTypeStackableStd Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_PlugTypeStackableStd) as Ele_PlugTypeStackableStd FROM " & sLookupTable
'                                rsLookupEle_PlugTypeStackableStd = New ADODB.Recordset
'                                rsLookupEle_PlugTypeStackableStd.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_PlugTypeStackableStd.MoveFirst()
'                            rsLookupEle_PlugTypeStackableStd.Find "Ele_PlugTypeStackableStd = " & sTemp1

'                            If rsLookupEle_PlugTypeStackableStd.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "Ele_ULWireType"
'                            '                        sSQL = sSQL & "WHERE Upper(Ele_ULWireType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupEle_ULWireType Is Nothing Then
'                                sSQL = "SELECT Upper(Ele_ULWireType) as Ele_ULWireType FROM " & sLookupTable
'                                rsLookupEle_ULWireType = New ADODB.Recordset
'                                rsLookupEle_ULWireType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupEle_ULWireType.MoveFirst()
'                            rsLookupEle_ULWireType.Find "Ele_ULWireType = " & sTemp1

'                            If rsLookupEle_ULWireType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "LEDEpoxyType"
'                            '                        sSQL = sSQL & "WHERE Upper(LEDEpoxyType) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupLEDEpoxyType Is Nothing Then
'                                sSQL = "SELECT Upper(LEDEpoxyType) as LEDEpoxyType FROM " & sLookupTable
'                                rsLookupLEDEpoxyType = New ADODB.Recordset
'                                rsLookupLEDEpoxyType.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupLEDEpoxyType.MoveFirst()
'                            rsLookupLEDEpoxyType.Find "LEDEpoxyType = " & sTemp1

'                            If rsLookupLEDEpoxyType.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                            '03/18/2009 - hn - lookups for PriceOption, FCAOrderPoint - added
'                        Case "PriceOption"
'                            '                        sSQL = sSQL & "WHERE Upper(PriceOption) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupPriceOption Is Nothing Then
'                                sSQL = "SELECT Upper(PriceOption) as PriceOption FROM " & sLookupTable
'                                rsLookupPriceOption = New ADODB.Recordset
'                                rsLookupPriceOption.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupPriceOption.MoveFirst()
'                            rsLookupPriceOption.Find "PriceOption = " & sTemp1

'                            If rsLookupPriceOption.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "FCAOrderPoint"
'                            '                        sSQL = sSQL & "WHERE Upper(FCAOrderPoint) = " & UCase(sAddQuotes(vFieldData))
'                            If rsLookupFCAOrderPoint Is Nothing Then
'                                sSQL = "SELECT Upper(FCAOrderPoint) as FCAOrderPoint FROM " & sLookupTable
'                                rsLookupFCAOrderPoint = New ADODB.Recordset
'                                rsLookupFCAOrderPoint.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupFCAOrderPoint.MoveFirst()
'                            rsLookupFCAOrderPoint.Find "FCAOrderPoint = " & sTemp1

'                            If rsLookupFCAOrderPoint.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case "TargetCertifiedPrinters"                              '2011/08/15
'                            Dim sSpacePad As String
'                            If Not IsBlank(vFieldData) Then
'                                sSpacePad = CharString(vFieldData, 4, True)
'                                sSpacePad = Replace(sSpacePad, " ", 0)
'                            End If
'                            sSQL = sSQL & "WHERE CertifiedPrinterID = '" & sSpacePad & "'" '2011/08/29

'                            If rsLookupTargetCertifiedPrinters Is Nothing Then
'                                sSQL = "SELECT CertifiedPrinterID FROM " & sLookupTable
'                                rsLookupTargetCertifiedPrinters = New ADODB.Recordset
'                                rsLookupTargetCertifiedPrinters.Open(sSQL, SSDataConn, adUseClient, adLockReadOnly)
'                            End If

'                            ' use pre-read lookup table
'                            ' sTemp1 = UCase(sAddQuotes(vFieldData))

'                            rsLookupTargetCertifiedPrinters.MoveFirst()
'                            rsLookupTargetCertifiedPrinters.Find "CertifiedPrinterID = '" & sSpacePad & "'"

'                            If rsLookupTargetCertifiedPrinters.EOF Then
'                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                '                        'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        Case Else
'                            sValidationErrorMsg = "Unknown Lookup Table: " & sLookupTable
'                            '                        MsgBox "Unknown Lookup Table: " & sLookupTable, vbExclamation + vbMsgBoxSetForeground, "bValidateField"
'                            '                        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End Select

'                    If sValidationErrorMsg = "" Then
'                        Select Case sLookupTable
'                            Case "ItemStatusCodes", "Category", "Grade", "Season", "FOBPoints", "Country", _
'                                    "Ele_Connection", "Ele_IndoorOrIndoorOutdoor", "Ele_TrayPlasticPaper", _
'                                    "Ele_FuseRating", "Ele_TrayOrBulk", "Ele_WireGauge", "Ele_LampBaseType", _
'                                    "Ele_LampBrightness", "CertificationMark", "CertificationType", "Bag_PaperType", _
'                                    "Bag_PrintType", "Bag_Finish", "Bag_HandleType", "Tree_Construction", "SalesRep", _
'                                    "LightType", "BatteryType", "Ele_PlugTypeStackableStd", "Ele_ULWireType", _
'                                    "LEDEpoxyType", "PriceOption", "FCAOrderPoint", "TargetCertifiedPrinters"
'                                ' do nothing - already been validated above

'                            Case Else
'                                ' use old way
'                                rsLookupTable = New ADODB.Recordset

'Dim                             rsLookupTable.Open sSQL As Object 
'                                Dim SSDataConn As Object
'                                Dim adOpenStatic As Object
'                                Dim adLockReadOnly As Object

'                                If rsLookupTable.EOF Then
'                                    '                                If sLookupTable = "Program" Or sLookupTable = "SalesRep" Then
'                                    sActiveMsg = ""
'                                    If sLookupTable = "Program" Then
'                                        sActiveMsg = " , OR INACTIVE!"
'                                    End If

'                                    '                                If vFieldData = "FALSE" And sLookupTable = "Tree_LightConstruction" Then
'                                    '                                    vFieldData = ""
'                                    '                                ' translate old boolean values on spreadsheet later
'                                    '                                Else
'                                    '                    sValidationErrorMsg = vFieldData & " not in " & sLookupTable & " table" & sActiveMsg           '11/19/2008
'                                    sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & sActiveMsg
'                                    '                                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                    '                                End If
'                                Else
'                                    Select Case sLookupTable
'                                        ' ET 2012-12-14 - removed the ones that are validated above
'                                        ' ET 2012-12-04 - add CertificationMark and CertificationType
'                                        '11/19/2008 - added PackageTypes
'                                        '                                    Case "Program", "Licensor", "Category", "PackageTypes", "CertificationMark", "CertificationType"
'                                        Case "Program", "Licensor", "PackageTypes"
'                                            If Not IsNull(rsLookupTable!InactiveProgramYear) And rsLookupTable!InactiveProgramYear <= CInt(sProgramYR) Then
'                                                sActiveMsg = " OR " & sLookupTable & ".InactiveProgramYear: " & rsLookupTable!InactiveProgramYear & " <= Item.ProgramYear:" & sProgramYR
'                                                '                            sValidationErrorMsg = vFieldData & " not in " & sLookupTable & " table" & sActiveMsg          '11/19/2008
'                                                sValidationErrorMsg = vFieldData & "not in " & sLookupTable & " table" & vbCrLf & "        " & sActiveMsg
'                                                '                                            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                            Else
'                                                sActiveMsg = ""
'                                            End If
'                                    End Select
'                                End If

'                                If rsLookupTable.State <> 0 Then rsLookupTable.Close()
'                                rsLookupTable = Nothing
'                        End Select
'                    End If

'                Case Else
'                    If bImport = True Then
'                        sValidationErrorMsg = "Unknown Validation Type: " & sValidationType
'                        '                    MsgBox "Unknown Validation Type: " & sValidationType, vbExclamation + vbMsgBoxSetForeground, "bValidateField"
'                        '                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If
'            End Select

'            If sValidationErrorMsg = "" Then
'                If sFieldName = "ItemStatus" And vFieldData = "ORD" Then                '01/23/2008
'                    If bCheckFactoryProhibitedRegisteredForItemStatus(sFactoryNumber, sCustomerNumber, sProgramYR, sValidationErrorMsg) = False Then
'                        '                    GoTo ExitRoutine    ' go to pitstop'TODO - GoTo Statements are redundant in .NET
'                    End If
'                End If
'            End If

'            ' 01/31/2008- TW decided against this for now!
'            '----------------------------------------------
'            'no more than one TempItemNumber <> blank per Proposal
'            '    If sFunctionCode <> gsNEW_PROPOSAL And sFieldName = "TempItemNumber" And Not IsBlank(vFieldData) Then
'            '        Dim rsTempItem As ADODB.Recordset: Set rsTempItem = New ADODB.Recordset
'            '        sSQL = "SELECT DISTINCT TempItemNumber FROM Item " & _
'            '               "WHERE ProposalNumber = " & sProposalNum & _
'            '               " AND TempItemNumber <> '' AND TempItemNumber <> '" & vFieldData & "'"
'Dim     '        rsTempItem.Open sSQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockReadOnly As Object

'            '
'            '        If Not rsTempItem.EOF Then
'            '            If rsTempItem.RecordCount > 1 Then 'this enables the user to change the tempitemnumber
'            '                sValidationErrorMsg = rsTempItem.RecordCount & " extra TempItemNumber(s) found:"
'            '                rsTempItem.MoveFirst
'            '                Do Until rsTempItem.EOF
'            '                    sValidationErrorMsg = sValidationErrorMsg & "[" & rsTempItem!TempItemNumber & "],"
'            '                    rsTempItem.MoveNext
'            '                Loop
'            '                sValidationErrorMsg = sValidationErrorMsg & vbCrLf & " Can't have more than 1 TempItemNumber per Proposal!"
'            '                If rsTempItem.RecordCount > 1 Then
'            '                    sValidationErrorMsg = sValidationErrorMsg & vbCrLf & " CONTACT System Administrator to fix!"
'            '                End If
'            '                If rsTempItem.State <> 0 Then rsTempItem.Close
'            '                Set rsTempItem = Nothing
'            '    '                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            '            End If
'            '        End If
'            '        If rsTempItem.State <> 0 Then rsTempItem.Close
'            '        Set rsTempItem = Nothing
'            '    End If

'            If bPROPOSALFormIndicator = False Then
'                'only do Import Spreadsheets, Proposal Form has another message....
'                If sFunctioncode <> gsNEW_REVISION Then
'                    If IsNull(sProposalNum) = True Or IsNull(sRev) = True Or sProposalNum = "" Or sRev = "" Or _
'                        sFunctioncode = gsNEW_PROPOSAL Then
'                        Application.DoEvents() '2013/04/23 - HN
'                        ' do nothing
'                    Else
'                        Dim lProposalNum As Long
'                        Dim lRev As Long

'                        lProposalNum = sProposalNum
'                        lRev = sRev

'                        sSQL = ""

'                        Select Case sFieldName
'                            'check Status/TempItem#/Season/Category/Program/CustItem# Fields are changed only for the Latest Rev
'                            Case "CustomerItemNumber", "CategoryCode", "ProgramNumber", "SeasonCode", "ItemStatus", "TempItemNumber"
'                                sSQL = "SELECT Item." & sFieldName & " AS ImportValue, " & _
'                                        "Item.Rev, Item.ProgramYear FROM Item " & _
'                                        "WHERE Item.ProposalNumber = " & lProposalNum & " and Item.ProgramYear = " & _
'                                        sProgramYR & " ORDER BY Item.Rev DESC"
'                                '            Case
'                                '                sSQL = "SELECT ItemSpecs." & sFieldName & " AS ImportValue, " & _
'                                '                        "Item.Rev, Item.ProgramYear FROM Item " & _
'                                '                        " INNER JOIN ItemSpecs ON (Item.ProposalNumber = ItemSpecs.ProposalNumber " & _
'                                '                        " AND Item.Rev = ItemSpecs.Rev) " & _
'                                '                    " WHERE Item.ProposalNumber = " & lProposalnum & _
'                                '                    " AND Item.ProgramYear = " & sProgramYR & " ORDER BY Item.Rev DESC"

'                            Case Else
'                                If sValidationErrorMsg <> "" Then           'new 11/08/2007
'                                    bValidateField = False
'                                Else
'                                    bValidateField = True
'                                End If

'                                '                            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        End Select

'                        If sSQL <> "" Then
'                            'first check if this field value has changed?
'                            Dim sImportValue As String
'                            Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'Dim                         rs.Open sSQL As Object 
'Dim  CurrentProject.Connection As Object 
'                            Dim adOpenStatic As Object
'                            Dim adLockReadOnly As Object


'                            Do Until rs.EOF
'                                If rs![Rev] = lRev Then                         'compare Rev on Spreadsheet against RS!Rev
'                                    If IsBlank(rs![ImportValue]) Then
'                                        sImportValue = ""
'                                    Else
'                                        sImportValue = rs![ImportValue]
'                                    End If

'                                    'if Value on Spreadsheet has been changed from db value
'                                    If sImportValue <> vFieldData And sImportValue <> "" Then   '03/17/2008
'                                        rs.MoveFirst()                        ' move to Latest Rev
'                                        If IsBlank(rs![ImportValue]) Then
'                                            sImportValue = ""
'                                        Else
'                                            sImportValue = rs![ImportValue]
'                                        End If

'                                        If sImportValue <> vFieldData Then  'compare value on Spreadsheet to Latest Rev's value
'                                            If rs![Rev] <> lRev Then
'                                                If sImportValue = "" Then sImportValue = "NULL"
'                                                sValidationErrorMsg = "CAN'T change Spreadsheet value for Rev(" & lRev & _
'                                                            ")," & vbCrLf & _
'                                                            "       when Highest Rev's(" & rs![Rev] & ") value[" & _
'                                                            sImportValue & "] for ProgramYr[" & sProgramYR & _
'                                                            "] differs." & vbCrLf & _
'                                                            "       Please correct on Proposal[" & lProposalNum & "] Form!"
'                                                rs.Close()
'                                                rs = Nothing
'                                                '                                            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                                            Else
'                                                Exit Do
'                                            End If
'                                        End If
'                                    Else
'                                        Exit Do
'                                    End If
'                                    Exit Do 'nothing changed
'                                End If
'                                rs.MoveNext()
'                            Loop

'                            rs.Close()
'                            rs = Nothing
'                        End If

'                    End If
'                End If
'            End If

'            If sValidationErrorMsg <> "" Then           '2013/04/23 - HN
'                bValidateField = False
'            Else
'                bValidateField = True
'            End If
'            '        bValidateField = True
'        End If  ' gsDENOTE_CELL_ERROR

'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rsLookupTable.State <> 0 Then rsLookupTable.Close()
'        rsLookupTable = Nothing
'        Exit Function

'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateField")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidateField , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine


'    Public Sub ValidationCleanup()
'        '   call this sub to drop all the recordsets to prevent memory leaks.
'        '   ET 2012-12-17 - called from frmUPCGenerator

'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally

'        '    if rsLookupItem.State <> 0 Then rsLookupItem.Close
'        '    if rsLookupItemSpecs.State <> 0 Then rsLookupItemSpecs.Close
'        '    if rsLookupItem_Assortments.State <> 0 Then rsLookupItem_Assortments.Close
'        If rsLookupItemStatusCodes.State <> 0 Then rsLookupItemStatusCodes.Close()
'        If rsLookupCategory.State <> 0 Then rsLookupCategory.Close()
'        '    if rsLookupCustomer.State <> 0 Then rsLookupCustomer.Close
'        '    if rsLookupProgram.State <> 0 Then rsLookupProgram.Close
'        If rsLookupGrade.State <> 0 Then rsLookupGrade.Close()
'        '    if rsLookupLicensor.State <> 0 Then rsLookupLicensor.Close
'        '    if rsLookupSubProgram.State <> 0 Then rsLookupSubProgram.Close
'        '    if rsLookupFactory.State <> 0 Then rsLookupFactory.Close
'        '    if rsLookupVendor.State <> 0 Then rsLookupVendor.Close
'        If rsLookupSeason.State <> 0 Then rsLookupSeason.Close()
'        If rsLookupFOBPoints.State <> 0 Then rsLookupFOBPoints.Close()
'        If rsLookupCountry.State <> 0 Then rsLookupCountry.Close()
'        '    if rsLookupPackageTypes.State <> 0 Then rsLookupPackageTypes.Close
'        If rsLookupEle_Connection.State <> 0 Then rsLookupEle_Connection.Close()
'        If rsLookupEle_IndoorOrIndoorOutdoor.State <> 0 Then rsLookupEle_IndoorOrIndoorOutdoor.Close()
'        If rsLookupEle_TrayPlasticPaper.State <> 0 Then rsLookupEle_TrayPlasticPaper.Close()
'        If rsLookupEle_FuseRating.State <> 0 Then rsLookupEle_FuseRating.Close()
'        If rsLookupEle_TrayOrBulk.State <> 0 Then rsLookupEle_TrayOrBulk.Close()
'        If rsLookupEle_WireGauge.State <> 0 Then rsLookupEle_WireGauge.Close()
'        If rsLookupEle_LampBaseType.State <> 0 Then rsLookupEle_LampBaseType.Close()
'        If rsLookupEle_LampBrightness.State <> 0 Then rsLookupEle_LampBrightness.Close()
'        If rsLookupCertificationMark.State <> 0 Then rsLookupCertificationMark.Close()
'        If rsLookupCertificationType.State <> 0 Then rsLookupCertificationType.Close()
'        If rsLookupBag_PaperType.State <> 0 Then rsLookupBag_PaperType.Close()
'        If rsLookupBag_PrintType.State <> 0 Then rsLookupBag_PrintType.Close()
'        If rsLookupBag_Finish.State <> 0 Then rsLookupBag_Finish.Close()
'        If rsLookupBag_HandleType.State <> 0 Then rsLookupBag_HandleType.Close()
'        If rsLookupTree_Construction.State <> 0 Then rsLookupTree_Construction.Close()
'        If rsLookupTree_LightConstruction.State <> 0 Then rsLookupTree_LightConstruction.Close()
'        If rsLookupSalesRep.State <> 0 Then rsLookupSalesRep.Close()
'        If rsLookupLightType.State <> 0 Then rsLookupLightType.Close()
'        If rsLookupBatteryType.State <> 0 Then rsLookupBatteryType.Close()
'        If rsLookupEle_PlugTypeStackableStd.State <> 0 Then rsLookupEle_PlugTypeStackableStd.Close()
'        If rsLookupEle_ULWireType.State <> 0 Then rsLookupEle_ULWireType.Close()
'        If rsLookupLEDEpoxyType.State <> 0 Then rsLookupLEDEpoxyType.Close()
'        If rsLookupPriceOption.State <> 0 Then rsLookupPriceOption.Close()
'        If rsLookupFCAOrderPoint.State <> 0 Then rsLookupFCAOrderPoint.Close()
'        If rsLookupTargetCertifiedPrinters.State <> 0 Then rsLookupTargetCertifiedPrinters.Close()

'        '    Set rsLookupItem = Nothing
'        '    Set rsLookupItemSpecs = Nothing
'        '    Set rsLookupItem_Assortments = Nothing
'        rsLookupItemStatusCodes = Nothing
'        rsLookupCategory = Nothing
'        '    Set rsLookupCustomer = Nothing
'        '    Set rsLookupProgram = Nothing
'        rsLookupGrade = Nothing
'        '    Set rsLookupLicensor = Nothing
'        '    Set rsLookupSubProgram = Nothing
'        '    Set rsLookupFactory = Nothing
'        '    Set rsLookupVendor = Nothing
'        rsLookupSeason = Nothing
'        rsLookupFOBPoints = Nothing
'        rsLookupCountry = Nothing
'        '    Set rsLookupPackageTypes = Nothing
'        rsLookupEle_Connection = Nothing
'        rsLookupEle_IndoorOrIndoorOutdoor = Nothing
'        rsLookupEle_TrayPlasticPaper = Nothing
'        rsLookupEle_FuseRating = Nothing
'        rsLookupEle_TrayOrBulk = Nothing
'        rsLookupEle_WireGauge = Nothing
'        rsLookupEle_LampBaseType = Nothing
'        rsLookupEle_LampBrightness = Nothing
'        rsLookupCertificationMark = Nothing
'        rsLookupCertificationType = Nothing
'        rsLookupBag_PaperType = Nothing
'        rsLookupBag_PrintType = Nothing
'        rsLookupBag_Finish = Nothing
'        rsLookupBag_HandleType = Nothing
'        rsLookupTree_Construction = Nothing
'        rsLookupTree_LightConstruction = Nothing
'        rsLookupSalesRep = Nothing
'        rsLookupLightType = Nothing
'        rsLookupBatteryType = Nothing
'        rsLookupEle_PlugTypeStackableStd = Nothing
'        rsLookupEle_ULWireType = Nothing
'        rsLookupLEDEpoxyType = Nothing
'        rsLookupPriceOption = Nothing
'        rsLookupFCAOrderPoint = Nothing
'        rsLookupTargetCertifiedPrinters = Nothing
'    End Sub

        
'    Private Function bGetSpecialSpreadsheetCOLPositions(ByRef dtCOLPos As typSpecialCOLPos, _
'                        dtSpreadsheetCOLArray() As typColumn, ByVal lColsOnSheet As Long) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lImportCounter As Long
'        Dim sDB4TableName As String

'        bGetSpecialSpreadsheetCOLPositions = False

'        dtCOLPos.lGrade = lGetSpreadsheetCOL(gsCOL_Grade, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lSubProgram = lGetSpreadsheetCOL(gsCOL_SubProgram, dtSpreadsheetCOLArray, lColsOnSheet)

'        dtCOLPos.lCertifiedPrinterID = lGetSpreadsheetCOL(gsCOL_CertifiedPrinterID, dtSpreadsheetCOLArray(), lColsOnSheet)          '2011/10/26
'        dtCOLPos.lX_CertifiedPrinterName = lGetSpreadsheetCOL(gsCOL_X_CertifiedPrinterName, dtSpreadsheetCOLArray(), lColsOnSheet)  '2011/10/26
'        dtCOLPos.lTechnologies = lGetSpreadsheetCOL(gsCOL_Technologies, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lX_Technologies = lGetSpreadsheetCOL(gsCOL_X_Technologies, dtSpreadsheetCOLArray(), lColsOnSheet)

'        dtCOLPos.lItemNumber = lGetSpreadsheetCOL(gsCOL_ItemNumber, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lTempItemNumber = lGetSpreadsheetCOL(gsCOL_TempItemNumber, dtSpreadsheetCOLArray(), lColsOnSheet)
'        '   UPCNumber column can no longer be imported ----  can only uncomment if user CAN import!!
'        '   then also change Import column on Field table
'        '    udtColPos.lUPCNumber = lGetSpreadsheetCol(gsCOL_UPC_NBR, dtSpreadsheetColArray(), lCOLsOnSheet)

'        dtCOLPos.lAltCost = lGetSpreadsheetCOL(gsCOL_AltCost, dtSpreadsheetCOLArray(), lColsOnSheet)                '12/08/2008 - hn
'        dtCOLPos.lAltSellPrice = lGetSpreadsheetCOL(gsCOL_AltSellPrice, dtSpreadsheetCOLArray(), lColsOnSheet)      '12/08/2008 - hn
'        dtCOLPos.lBag_SpecialEffects = lGetSpreadsheetCOL(gsCol_Bag_SpecialEffects, dtSpreadsheetCOLArray(), lColsOnSheet)  '2009/11/19 - hn
'        dtCOLPos.lClass = lGetSpreadsheetCOL(gsCOL_CLASS, dtSpreadsheetCOLArray(), lColsOnSheet)                    '11/20/2008 - hn
'        dtCOLPos.lCustomerHSNumber = lGetSpreadsheetCOL("CustomerHSNumber", dtSpreadsheetCOLArray(), lColsOnSheet)  '03/05/2009 - hn
'        dtCOLPos.lCustomerCartonUPCNumber = lGetSpreadsheetCOL("CustomerCartonUPCNumber", dtSpreadsheetCOLArray(), lColsOnSheet)    '2012/05/01
'        dtCOLPos.lOtherText = lGetSpreadsheetCOL("OtherText", dtSpreadsheetCOLArray(), lColsOnSheet)                '2012/05/02
'        dtCOLPos.lCustomerDutyRate = lGetSpreadsheetCOL("CustomerDutyRate", dtSpreadsheetCOLArray(), lColsOnSheet)  '03/05/2009 - hn
'        dtCOLPos.lCoreItemNumber = lGetSpreadsheetCOL(gsCOL_CoreItemNumber, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lcustomernumber = lGetSpreadsheetCOL(gsCOL_CustomerNumber, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lFactoryFCACost = lGetSpreadsheetCOL(gsCOL_FactoryFCACost, dtSpreadsheetCOLArray(), lColsOnSheet)  '12/08/2008 - hn
'        dtCOLPos.lFactoryFOBCost = lGetSpreadsheetCOL(gsCOL_FactoryFOBCost, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lFactoryNumber = lGetSpreadsheetCOL(gsCOL_FactoryNumber, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lFCASellPrice = lGetSpreadsheetCOL(gsCOL_FCASellPrice, dtSpreadsheetCOLArray(), lColsOnSheet)      '12/08/2008 - hn
'        dtCOLPos.lSQ2 = lGetSpreadsheetCOL(gsCOL_SQ2, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lRevisedDate = lGetSpreadsheetCOL(gsCOL_RevisedDate, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lRevisedUserID = lGetSpreadsheetCOL(gsCOL_RevisedUserID, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lBaseProposalNumber = lGetSpreadsheetCOL(gsCOL_BaseProposalNumber, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lBaseRev = lGetSpreadsheetCOL(gsCOL_BaseRev, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lDutyPercent = lGetSpreadsheetCOL(gsCOL_DutyPercent, dtSpreadsheetCOLArray(), lColsOnSheet)
'        '    dtCOLPos.lFactoryFOBCost = lGetSpreadsheetCOL(gsCOL_FactoryFOBCost, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lHSNumber = lGetSpreadsheetCOL(gsCOL_HSNumber, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lLEDLightCost = lGetSpreadsheetCOL(gsCOL_LEDLightCost, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lREGLinePrice = lGetSpreadsheetCOL(gsCOL_REGLinePrice, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lSellPrice = lGetSpreadsheetCOL(gsCOL_FOBSellPrice, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lTradeMarkCopyRight = lGetSpreadsheetCOL(gsCOL_TrademarkCopyRight, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lLicensor = lGetSpreadsheetCOL(gsCOL_Licensor, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lRoyaltyPercent = lGetSpreadsheetCOL(gsCOL_RoyaltyPercent, dtSpreadsheetCOLArray(), lColsOnSheet)      '11/19/2008 - hn
'        dtCOLPos.lLighted = lGetSpreadsheetCOL(gsCOL_Lighted, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lTreeLightConstruction = lGetSpreadsheetCOL(gsCOL_TreeLightConstruction, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lProgramYear = lGetSpreadsheetCOL(gsCOL_ProgramYear, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lInnerPackLength = lGetSpreadsheetCOL(gsCOL_InnerPackLength, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lInnerPackHeight = lGetSpreadsheetCOL(gsCOL_InnerPackHeight, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lInnerPackWidth = lGetSpreadsheetCOL(gsCOL_InnerPackWidth, dtSpreadsheetCOLArray(), lColsOnSheet)
'        '    dtColPos.lInnerPackCube = lGetSpreadsheetCol(gsCOL_INNERPACKCUBE, dtSpreadsheetColArray(), lCOLsOnSheet)
'        dtCOLPos.lMasterPackLength = lGetSpreadsheetCOL(gsCOL_MasterPackLength, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lMasterPackHeight = lGetSpreadsheetCOL(gsCOL_MasterPackHeight, dtSpreadsheetCOLArray(), lColsOnSheet)
'        dtCOLPos.lMasterPackWidth = lGetSpreadsheetCOL(gsCOL_MasterPackWidth, dtSpreadsheetCOLArray(), lColsOnSheet)
'        '    dtColPos.lMasterPackCube = lGetSpreadsheetCol(gsCOL_MASTERPACKCUBE, SpreadsheetCOLArray(), lCOLsOnSheet)
'        dtCOLPos.lDEVComments = lGetSpreadsheetCOL("DevelopComments", dtSpreadsheetCOLArray(), lColsOnSheet)

'        dtCOLPos.lProductBatteriesIncluded = lGetSpreadsheetCOL("ProductBatteriesIncluded", dtSpreadsheetCOLArray(), lColsOnSheet) 'new Elec Specs field 03/14/2008
'        dtCOLPos.lLightType = lGetSpreadsheetCOL("LightType", dtSpreadsheetCOLArray(), lColsOnSheet)            '03/19/2008 - hn
'        '    dtCOLPos.lAssortments = lGetSpreadsheetCOL("Assortments", dtSpreadsheetCOLArray(), lColsOnSheet)    '04/16/2008 - hn
'        dtCOLPos.lLowLeadWholeProduct = lGetSpreadsheetCOL("LowLeadWholeProduct", dtSpreadsheetCOLArray(), lColsOnSheet)            '2014/10/03 RAS
'        dtCOLPos.lFlammability = lGetSpreadsheetCOL("Flammability", dtSpreadsheetCOLArray(), lColsOnSheet)            '2014/10/03 RAS
'        dtCOLPos.lSurfaceLeadPaintRequirement = lGetSpreadsheetCOL("SurfaceLeadPaintReq", dtSpreadsheetCOLArray(), lColsOnSheet)            '2014/10/03 RAS

'        ' .... saves a lot of hard-coding, checks Field ExcelShadeCell column .............
'        For lImportCounter = 1 To UBound(dtSpreadsheetCOLArray())
'            If dtSpreadsheetCOLArray(lImportCounter).bExcelShadeCell = True Then
'                dtSpreadsheetCOLArray(lImportCounter).lExcelCOLNum = _
'                    lGetSpreadsheetCOL(dtSpreadsheetCOLArray(lImportCounter).sColumnName, dtSpreadsheetCOLArray(), lColsOnSheet)
'            End If
'        Next lImportCounter

'        bGetSpecialSpreadsheetCOLPositions = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bGetSpecialSpreadsheetCOLPositions")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bGetSpecialSpreadsheetCOLPositions , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine


'    End Function
        
'    Private Function bGetSpecialSaveArrayCOLPositions(ByRef dtCOLPos As typSpecialCOLPos) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lImportCounter As Long
'        Dim sDB4TableName As String
'        Dim sDB4FieldName As String

'        bGetSpecialSaveArrayCOLPositions = False
'        ' get column number from either the ITEM or ItemSPECS array
'        dtCOLPos.lAltCost = lGetSaveArrayCOL(gsCOL_AltCost, sItemArray(), lItemFields)            '12/08/2008 - hn
'        dtCOLPos.lAltSellPrice = lGetSaveArrayCOL(gsCOL_AltSellPrice, sItemArray(), lItemFields)   '12/08/2008 - hn
'        dtCOLPos.lBag_SpecialEffects = lGetSaveArrayCOL(gsCol_Bag_SpecialEffects, sItemSPECSArray(), lItemSpecsFields) '2009/11/19 - hn
'        dtCOLPos.lBaseProposalNumber = lGetSaveArrayCOL(gsCOL_BaseProposalNumber, sItemArray(), lItemFields)
'        dtCOLPos.lBaseRev = lGetSaveArrayCOL(gsCOL_BaseRev, sItemArray(), lItemFields)
'        dtCOLPos.lCategoryCode = lGetSaveArrayCOL(gsCOL_CategoryCode, sItemArray(), lItemFields)

'        dtCOLPos.lCertifiedPrinterID = lGetSaveArrayCOL(gsCOL_CertifiedPrinterID, sItemArray(), lItemFields)         '2011/10/26
'        dtCOLPos.lX_CertifiedPrinterName = lGetSaveArrayCOL(gsCOL_X_CertifiedPrinterName, sItemArray(), lItemFields) '2011/10/26

'        dtCOLPos.lClass = lGetSaveArrayCOL(gsCOL_CLASS, sItemArray(), lItemFields)              '11/20/2008 - hn
'        dtCOLPos.lCoreItemNumber = lGetSaveArrayCOL(gsCOL_CoreItemNumber, sItemArray(), lItemFields)
'        dtCOLPos.lCustomerItemNumber = lGetSaveArrayCOL(gsCOL_CustomerItemNumber, sItemArray(), lItemFields)
'        dtCOLPos.lcustomernumber = lGetSaveArrayCOL(gsCOL_CustomerNumber, sItemArray(), lItemFields)
'        dtCOLPos.lCustomerCartonUPCNumber = lGetSaveArrayCOL("CustomerCartonUPCNumber", sItemArray(), lItemFields)  '2012/05/01
'        dtCOLPos.lOtherText = lGetSaveArrayCOL("OtherText", sItemArray(), lItemFields)                   '2012/05/02
'        dtCOLPos.lDutyPercent = lGetSaveArrayCOL(gsCOL_DutyPercent, sItemArray(), lItemFields)
'        dtCOLPos.lFactoryFOBCost = lGetSaveArrayCOL(gsCOL_FactoryFOBCost, sItemArray(), lItemFields)
'        dtCOLPos.lFactoryFCACost = lGetSaveArrayCOL(gsCOL_FactoryFCACost, sItemArray(), lItemFields)
'        dtCOLPos.lFactoryNumber = lGetSaveArrayCOL(gsCOL_FactoryNumber, sItemArray(), lItemFields)
'        dtCOLPos.lFCASellPrice = lGetSaveArrayCOL(gsCOL_FCASellPrice, sItemArray(), lItemFields)    '12/08/2008
'        dtCOLPos.lGrade = lGetSaveArrayCOL(gsCOL_Grade, sItemArray(), lItemFields)
'        dtCOLPos.lHSNumber = lGetSaveArrayCOL(gsCOL_HSNumber, sItemArray(), lItemFields)
'        dtCOLPos.lItemNumber = lGetSaveArrayCOL(gsCOL_ItemNumber, sItemArray(), lItemFields)
'        dtCOLPos.lItemStatus = lGetSaveArrayCOL(gsCOL_ITEMSTATUS, sItemArray(), lItemFields)
'        dtCOLPos.lLongDesc = lGetSaveArrayCOL(gsCOL_LONGDESC, sItemArray(), lItemFields)
'        dtCOLPos.lProgramNumber = lGetSaveArrayCOL(gsCOL_ProgramNumber, sItemArray(), lItemFields)
'        dtCOLPos.lProgramYear = lGetSaveArrayCOL(gsCOL_ProgramYear, sItemArray(), lItemFields)
'        dtCOLPos.lREGLinePrice = lGetSaveArrayCOL(gsCOL_REGLinePrice, sItemArray(), lItemFields)
'        dtCOLPos.lSalesRepNumber = lGetSaveArrayCOL(gsCOL_SalesRepNumber, sItemArray(), lItemFields)
'        dtCOLPos.lSeasonCode = lGetSaveArrayCOL(gsCOL_SeasonCode, sItemArray(), lItemFields)
'        dtCOLPos.lSellPrice = lGetSaveArrayCOL(gsCOL_FOBSellPrice, sItemArray(), lItemFields)
'        dtCOLPos.lShortDesc = lGetSaveArrayCOL(gsCOL_SHORTDESC, sItemArray(), lItemFields)
'        dtCOLPos.lSubProgram = lGetSaveArrayCOL(gsCOL_SubProgram, sItemArray, lItemFields)
'        dtCOLPos.lTechnologies = lGetSaveArrayCOL(gsCOL_Technologies, sItemSPECSArray(), lItemSpecsFields)
'        dtCOLPos.lTempItemNumber = lGetSaveArrayCOL(gsCOL_TempItemNumber, sItemArray(), lItemFields)
'        dtCOLPos.lX_Technologies = lGetSaveArrayCOL(gsCOL_X_Technologies, sItemSPECSArray(), lItemSpecsFields)

'        dtCOLPos.lVendorItemNumber = lGetSaveArrayCOL(gsCOL_VendorItemNumber, sItemArray(), lItemFields)
'        dtCOLPos.lVendorNumber = lGetSaveArrayCOL(gsCOL_VendorNumber, sItemArray(), lItemFields)
'        dtCOLPos.lUPCNumber = lGetSaveArrayCOL(gsCOL_UPC_NBR, sItemArray(), lItemFields)
'        dtCOLPos.lIPKUPC = lGetSaveArrayCOL(gsCOL_IPKUPC, sItemArray(), lItemFields)
'        dtCOLPos.lMPKUPC = lGetSaveArrayCOL(gsCOL_MPKUPC, sItemArray(), lItemFields)
'        '    dtCOLPos.lPalletUPC = lGetSaveArrayCOL(gsCOL_PalletUPC, sItemArray(), lItemFields)       '2010/10/25 - if we do import as well...

'        dtCOLPos.lSQ2 = lGetSaveArrayCOL(gsCOL_SQ2, sItemArray(), lItemFields)
'        dtCOLPos.lRevisedDate = lGetSaveArrayCOL(gsCOL_RevisedDate, sItemArray(), lItemFields)
'        dtCOLPos.lRevisedUserID = lGetSaveArrayCOL(gsCOL_RevisedUserID, sItemArray(), lItemFields)
'        dtCOLPos.lLEDLightCost = lGetSaveArrayCOL(gsCOL_LEDLightCost, sItemArray(), lItemFields)

'        dtCOLPos.lTradeMarkCopyRight = lGetSaveArrayCOL(gsCOL_TrademarkCopyRight, sItemArray(), lItemFields)
'        dtCOLPos.lLicensor = lGetSaveArrayCOL(gsCOL_Licensor, sItemArray(), lItemFields)
'        dtCOLPos.lRoyaltyPercent = lGetSaveArrayCOL(gsCOL_RoyaltyPercent, sItemArray(), lItemFields)        '11/19/2008 - hn
'        dtCOLPos.lLighted = lGetSaveArrayCOL(gsCOL_Lighted, sItemArray(), lItemFields)
'        dtCOLPos.lTreeLightConstruction = lGetSaveArrayCOL(gsCOL_TreeLightConstruction, sItemSPECSArray(), lItemSpecsFields)
'        dtCOLPos.lInnerPackHeight = lGetSaveArrayCOL(gsCOL_InnerPackHeight, sItemArray(), lItemFields)
'        dtCOLPos.lInnerPackLength = lGetSaveArrayCOL(gsCOL_InnerPackLength, sItemArray(), lItemFields)
'        dtCOLPos.lInnerPackWidth = lGetSaveArrayCOL(gsCOL_InnerPackWidth, sItemArray(), lItemFields)
'        dtCOLPos.lInnerPackCube = lGetSaveArrayCOL(gsCOL_InnerPackCube, sItemArray(), lItemFields)
'        dtCOLPos.lMasterPackHeight = lGetSaveArrayCOL(gsCOL_MasterPackHeight, sItemArray(), lItemFields)
'        dtCOLPos.lMasterPackLength = lGetSaveArrayCOL(gsCOL_MasterPackLength, sItemArray(), lItemFields)
'        dtCOLPos.lMasterPackWidth = lGetSaveArrayCOL(gsCOL_MasterPackWidth, sItemArray(), lItemFields)
'        dtCOLPos.lMasterPackCube = lGetSaveArrayCOL(gsCOL_MasterPackCube, sItemArray(), lItemFields)

'        dtCOLPos.lDEVComments = lGetSaveArrayCOL("DevelopComments", sItemArray(), lItemFields)
'        dtCOLPos.lProductBatteriesIncluded = lGetSaveArrayCOL("ProductBatteriesIncluded", sItemSPECSArray(), lItemSpecsFields) '03/14/2008 - hn
'        dtCOLPos.lLightType = lGetSaveArrayCOL("LightType", sItemSPECSArray(), lItemSpecsFields)    '03/19/2008 - hn
'        '    dtCOLPos.lAssortments = lGetSaveArrayCOL("Assortments", sItemArray, lItemFields)            '04/17/2008 - hn

'        If glSHADING = NOShading Or glSHADING = NoShadingX Then      '2012/01/04 ' if from Proposal Form ensure it's set = 0
'        Else
'            For lImportCounter = 1 To UBound(dtSpreadsheetCOLArray())
'                If (glSHADING = FULLShading) Or (glSHADING = STDShading And dtSpreadsheetCOLArray(lImportCounter).bExcelShadeCell = True) Then
'                    sDB4TableName = dtSpreadsheetCOLArray(lImportCounter).sDB4Table
'                    sDB4FieldName = dtSpreadsheetCOLArray(lImportCounter).sDB4Field
'                    Select Case sDB4TableName
'                        Case gsItemSpecs_Table
'                            dtSpreadsheetCOLArray(lImportCounter).lArrayCOLNum = _
'                                lGetSaveArrayCOL(sDB4FieldName, sItemSPECSArray(), lItemSpecsFields)

'                        Case gsItem_Table
'                            dtSpreadsheetCOLArray(lImportCounter).lArrayCOLNum = _
'                                lGetSaveArrayCOL(sDB4FieldName, sItemArray(), lItemFields)

'                        Case gsItem_Assortments_Table
'                            dtSpreadsheetCOLArray(lImportCounter).lArrayCOLNum = _
'                                lGetSaveArrayCOL(sDB4FieldName, sAssortmentArray(), lItemSpecsFields)
'                        Case Else
'                            dtSpreadsheetCOLArray(lImportCounter).lArrayCOLNum = 0
'                    End Select
'                End If
'            Next lImportCounter
'        End If
'        bGetSpecialSaveArrayCOLPositions = True

'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bGetSpecialSaveArrayCOLPositions")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bGetSpecialSaveArrayCOLPositions , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'Private Function bValidateCoreFields(ByVal lRow As Long, ByVal lRowsOnSheet As Long, _
'                            sItemSPECSArray() As String, _
'                            sItemArray() As String , _
'        Dim lCounter As Long
'        Dim sFunctioncode As String
'        Dim sProposalNumber As String
'        Dim sRev As String


'        Dim lProgramYear As Long
'        Dim lProgNumber As Long
'        Dim sProgNumber As String
'        Dim sCategory As String

'        Dim sCertificationMark As String
'        Dim sCertificationType As String

'        Dim sInactiveProgramYear As String
'        Dim bGradeRequired As Boolean

'        Dim sSubProgram As String
'        Dim sGrade As String

'        Dim sLicensor As String  'to allow null values
'        Dim sRoyaltyPercent As String

'        Dim sTechnology As String
'        Dim sX_TechDescr As String
'        Dim sInvalidTechnology As String

'        Dim sCertifiedPrinterID As String
'        Dim sX_CertPrintDescr As String
'        Dim sInvalidCertifiedPrinter As String


'        Dim sBag_SpecialEffects As String
'        Dim sInvalidBag_SpecialEffects As String
'        Dim sLighted As String
'        Dim sLightedOptions As String
'        Dim sTreeLightConstruction As String

'        Dim sLightType As String
'        Dim sProductBatteriesIncluded As String
'        Dim sVendorNumber As String
'        Dim sFactoryNumber As String

'        Dim sHSNumber As String
'        Dim sDutyPercent As String

'        Dim lcustomernumber As Long
'        Dim sCustomerNumber As String
'        Dim lExistCustNum As Long

'        Dim sVendorItemNumber As String
'        Dim sItemNum As String
'        Dim sItemNumber As String

'        Dim sFOBSellPrice As String
'        Dim sBaseProposalNumber As String
'        Dim sBaseRev As String

'        Dim vReturnValue As Object
'        Dim sSQL As String

'        Dim sReturnErrormsg As String
'        Dim sSpreadsheetRow As String
'        Dim sSpreadsheetVendor As String
'        Dim sDevelopComments As String
'        Dim sItemStatus As String

'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        Dim lTotProposals As Long
'        Dim sRowErrorMsg As String
'        Const sThousand = "1000"

'        bValidateCoreFields = False
'        sFunctioncode = sItemArray(lRow, glFunctionCode_ColPos)
'        sProposalNumber = sItemArray(lRow, glProposal_ColPos)
'        sRev = sItemArray(lRow, glREV_ColPos)
'        sSubProgram = sItemArray(lRow, dtCOLPos.lSubProgram)
'        sGrade = sItemArray(lRow, dtCOLPos.lGrade)

'        If bPROPOSALFormIndicator = False Then
'            sCertifiedPrinterID = sItemArray(lRow, dtCOLPos.lCertifiedPrinterID)
'            sTechnology = sItemSPECSArray(lRow, dtCOLPos.lTechnologies)
'            sBag_SpecialEffects = sItemSPECSArray(lRow, dtCOLPos.lBag_SpecialEffects)
'        End If
'        '    If bFromProposalForm = False And dtCOLPos.lX_Technologies > 0 Then   'if qryItemData has less than 255 fields returned,  this can be  added:'' AS X_Technologies
'        '        sX_Technologies = sITEMArray(lRow, dtCOLPos.lX_Technologies)
'        '    End If
'        sVendorNumber = sItemArray(lRow, dtCOLPos.lVendorNumber)
'        sFactoryNumber = sItemArray(lRow, dtCOLPos.lFactoryNumber)
'        sCustomerNumber = sItemArray(lRow, dtCOLPos.lcustomernumber)

'        If IsNumeric(sCustomerNumber) = True Then
'            lcustomernumber = sCustomerNumber
'        Else
'            lcustomernumber = 0
'        End If

'        sCategory = sItemArray(lRow, dtCOLPos.lCategoryCode)
'        sProgNumber = sItemArray(lRow, dtCOLPos.lProgramNumber)
'        If IsNumeric(sProgNumber) = True Then
'            lProgNumber = sProgNumber
'        Else
'            lProgNumber = 0
'        End If

'        If IsNumeric(sItemArray(lRow, dtCOLPos.lProgramYear)) Then
'            lProgramYear = sItemArray(lRow, dtCOLPos.lProgramYear)
'        Else
'            lProgramYear = 0
'        End If

'        sLicensor = sItemArray(lRow, dtCOLPos.lLicensor)
'        sRoyaltyPercent = sItemArray(lRow, dtCOLPos.lRoyaltyPercent)
'        sLighted = sItemArray(lRow, dtCOLPos.lLighted)

'        If bPROPOSALFormIndicator = False Then
'            'these are subfrmItemSpecs fields , can only do this this for Import Spreadsheet, frmProposal handles it differently
'            sTreeLightConstruction = sItemSPECSArray(lRow, dtCOLPos.lTreeLightConstruction)
'            If sItemSPECSArray(lRow, dtCOLPos.lTreeLightConstruction) = "False" Then
'                sItemSPECSArray(lRow, dtCOLPos.lTreeLightConstruction) = ""
'            End If
'            sProductBatteriesIncluded = sItemSPECSArray(lRow, dtCOLPos.lProductBatteriesIncluded)
'            sLightType = sItemSPECSArray(lRow, dtCOLPos.lLightType)
'        End If

'        sVendorItemNumber = sItemArray(lRow, dtCOLPos.lVendorItemNumber)
'        sItemNum = sItemArray(lRow, dtCOLPos.lItemNumber)
'        sFOBSellPrice = sItemArray(lRow, dtCOLPos.lSellPrice)
'        sDevelopComments = sItemArray(lRow, dtCOLPos.lDEVComments)
'        sItemStatus = sItemArray(lRow, dtCOLPos.lItemStatus)
'        sItemStatus_ORIG = sItemArray_Orig(lRow, dtCOLPos.lItemStatus)

'        '......  above makes the code more readable as we continue ........

'        If sProposalNumber <> "" And sFunctioncode <> "A" Then
'            If bMatchExistingCustomerNumber(sProposalNumber, lcustomernumber, lExistCustNum) = False Then
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "ROW:" & lRow & " Proposal:" & sProposalNumber & " Rev:" & sRev & _
'                            "CustomerNumber on grid, does not match CustomerNumber[" & lExistCustNum & "] in DB4.")
'            End If
'        End If

'        If sItemNum <> "A" And Microsoft.VisualBasic.Right(sItemNum, 3) <> lcustomernumber And sItemNum <> "" And gbADDNewProposal = False Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, " Proposal:" & sProposalNumber & " Rev:" & sRev & _
'                        "The last 3 digits of ItemNumber[" & _
'                        Microsoft.VisualBasic.Right(sItemNum, 3) & "] ,must match CustomerNumber[" & lcustomernumber & "]")
'        End If

'        'cannot change ItemStatus = 'ORD' if ItemNumber is blank
'        If sItemStatus = "ORD" And sItemNum = "" Then
'            Call bBuildErrorMsg(bCreateNEWComponentItems, lRow, lNbrErrors, sRowErrorMsg, " Proposal:" & sProposalNumber & " Rev:" & sRev & _
'                        " ItemStatus cannot be = 'ORD' if ItemNumber is blank!")
'        End If

'        Dim sItemErrMsg As String

'        'cannot change Itemstatus from 'ORD' if an Order is present
'        '2013/05/07 -HN- Copied ET's code
'        ' ET 2013-03-15 - per Theresa should only get an error message here when
'        ' changing the status of an item when it exists on another OrderDetail where the order status
'        ' is not "cancelled" (OrderDetail.CancelCode is NULL or empty string).
'        If sFunctioncode <> "A" Then        '2011/03/22 - for a new one can't be on an Order yet!
'            If sItemStatus_ORIG = "ORD" And sItemStatus <> "ORD" Then  'for FC = R , A then don't have to check...
'                If bValidItemStatusChange(sProposalNumber, lProgramYear, sItemNum, sItemErrMsg) = False Then
'                    Call bBuildErrorMsg(bCreateNEWComponentItems, lRow, lNbrErrors, sRowErrorMsg, sItemErrMsg)
'                End If
'            End If
'        End If

'        If bCheckExistingItemNumber(sFunctioncode, sProposalNumber, sItemNum, sItemErrMsg) = True Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sItemErrMsg)
'        End If

'        Dim sCoreItem As String
'        If bProposalNEWItemfromList = False Then   'no need for new item from scratch to test this
'            If sProposalNumber <> "" And sItemNum <> "A" Then
'                If bCheckExistingCoreItemNumber(sProposalNumber, sItemNum, sCoreItem) = False Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "First 6 digits of ItemNumber[" & sItemNum & _
'                        "] does not match CoreItemNumber[" & sCoreItem & "] in DB4 of BaseProposal:" & sProposalNumber)
'                End If
'            End If
'        End If

'        sReturnErrormsg = ""
'        If bFindMultipleProposalNumbersInDB(sFunctioncode, sItemNum, sCategory, _
'                                        sVendorNumber, sFactoryNumber, _
'                                        lcustomernumber, sVendorItemNumber, _
'                                        sReturnErrormsg) = False Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sReturnErrormsg)
'        End If

'        sReturnErrormsg = ""
'        If sFunctioncode <> "" Or sItemNum = gsNEW_ITEM_NBR Then
'            If bFindDuplicateProposalsSpreadsheet(lRow, lRowsOnSheet, dtCOLPos, sItemSPECSArray(), sItemArray(), sReturnErrormsg) Then
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sReturnErrormsg)
'            End If
'        End If

'        ' Ensure there aren't multiple rows on the spreadsheet
'        ' for the ItemNumber / VendorNumber / FactoryNumber / CustomerNumber combination
'        sReturnErrormsg = "" ' Initialize: This is passed in by reference
'        If bFindDuplicateItemsOnSpreadsheet(lRow, lRowsOnSheet, dtCOLPos, _
'                                            sItemSPECSArray(), sItemArray(), _
'                                            sReturnErrormsg) Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sReturnErrormsg)
'        End If
'        '2014/06/05 RAS do not check or validate the lighted , batteries included Technoligies fields for Developement,  999,998, itemstatus DEV
'        '2014/08/15 RAS adding marketing basic or product manger to skip the checks
'        If UCase(sItemArray(lRow, dtCOLPos.lItemStatus)) = "DEV" Or (lcustomernumber = gs999PD_ACCOUNT Or lcustomernumber = gs998PD_ACCOUNT) Or (msUserGroup = msMKTGBASIC Or msUserGroup = msPRODUCTMGR) Then
'            'if the column is on the spreadsheet change it from "YES" to 1 and ""NO" to 0
'            If sItemArray(lRow, dtCOLPos.lLighted) = "YES" Then
'                sItemArray(lRow, dtCOLPos.lLighted) = "1"
'            End If
'            If sItemArray(lRow, dtCOLPos.lLighted) = "NO" Then
'                sItemArray(lRow, dtCOLPos.lLighted) = "0"
'            End If
'        Else
'            'Check that Lighted column is Blank, Y , N
'            sSQL = "SELECT LightedOptions FROM Program WHERE ProgramNumber = " & lProgNumber
'         rs.Open sSQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockReadOnly As Object

'            If Not rs.EOF Then
'                If Not IsNull(rs!LightedOptions) Then
'                    sLightedOptions = rs!LightedOptions
'                Else
'                    sLightedOptions = "B"
'                End If
'            End If

'            rs.Close()
'            Select Case sLighted 'value on Import or Proposal frm
'                Case "True", "YES", "Y", "1"
'                    If sLightedOptions = "Y" Then
'                        sItemArray(lRow, dtCOLPos.lLighted) = "1"
'                    Else
'                        If lProgramYear <= 2007 Then
'                            sItemArray(lRow, dtCOLPos.lLighted) = "1"
'                        Else
'                            If sLightedOptions = "B" Then
'                                sItemArray(lRow, dtCOLPos.lLighted) = "1"
'                            Else
'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Lighted: 'YES' Is Not Permissible for Program: " & lProgNumber & "-Lighted:" & sLightedOptions)
'                            End If
'                        End If
'                    End If

'                Case "NO", "N", "0", "False"
'                    If sLightedOptions = "N" Then
'                        sItemArray(lRow, dtCOLPos.lLighted) = "0"
'                    Else
'                        If lProgramYear <= 2007 Then
'                            sItemArray(lRow, dtCOLPos.lLighted) = "0"
'                        Else
'                            If sLightedOptions = "B" Then
'                                sItemArray(lRow, dtCOLPos.lLighted) = "0"
'                            Else
'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Lighted: 'NO' Is Not Permissible for Program: " & lProgNumber & "-Lighted:" & sLightedOptions)
'                            End If
'                        End If
'                    End If

'                Case ""
'                    If sLightedOptions = "" Or sLightedOptions = "B" Then
'                        If lProgramYear <= 2007 Then
'                            sItemArray(lRow, dtCOLPos.lLighted) = ""
'                        Else
'                            sErrorMsg = sErrorMsg & "ROW:" & lRow & _
'                            " - Lighted must be 'Yes', 'No', if ProgramYear >= 2008. " & vbCrLf
'                            lNbrErrors = lNbrErrors + 1
'                        End If
'                    Else
'                        If lProgramYear <= 2007 Then
'                            sItemArray(lRow, dtCOLPos.lLighted) = ""
'                        Else
'                            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Lighted: 'Blank' Is Not Permissible for Program: " & lProgNumber & "-Lighted:" & sLightedOptions)
'                        End If
'                    End If

'                Case Else

'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Lighted must be 'Yes', 'No', (or 'Blank' for ProgramYear < 2008).")
'            End Select

'            'If Lighted = "YES" then LightType must be defined!
'            If bPROPOSALFormIndicator = False Then 'Proposal Form checks it individually
'                Select Case sLighted 'value on Import
'                    Case "True", "YES", "Y", "1"
'                        If sLightType = "" Then
'                            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "if 'Lighted'= 'YES', then 'LightType' column must be present with a value.")
'                        End If
'                    Case Else
'                End Select
'            Else
'                Select Case sLighted
'                    'value on frmProposal form(tested differently, because ItemSpecs fields are tied to TempItemSpecs table until saved
'                    Case "True", "YES", "Y", "1"
'                        Dim rsTempItemSpecs As ADODB.Recordset : rsTempItemSpecs = New ADODB.Recordset
'                        Dim SQL As String
'                        SQL = "SELECT LightType FROM TempItemSpecs WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sRev
'                        rsTempItemSpecs.Open SQL
'                        CurrentProject.Connection()
'                        Dim adOpenStatic As Object
'                        Dim adLockOptimistic As Object

'                        If Not rsTempItemSpecs.EOF Then
'                            If IsBlank(rsTempItemSpecs!LightType) Then
'                                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "if 'Lighted'= 'YES', then Electric Specifications 'LightType' must have a value.")
'                            End If
'                        End If
'                        If rsTempItemSpecs.State <> 0 Then rsTempItemSpecs.Close()
'                        rsTempItemSpecs = Nothing
'                End Select
'            End If
'        End If
'        'Find InactiveProgramYear, GradeRequired for this Program, if any
'        If bFindInactiveProgramYear(lProgNumber, sInactiveProgramYear, bGradeRequired) = True Then
'        End If

'        'Check that Category.InactiveProgramYear is valid
'        If bValidCategoryInactiveProgramYear(sCategory, lProgramYear, sInactiveProgramYear) = False Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "For CategoryCode-" & sCategory & " InactiveProgramYear(" & sInactiveProgramYear & ") >= This Item's Program year: " & lProgramYear)
'        End If

'        'Check that Category/Program combination is valid
'        If bValidCategoryProgramCombination(sCategory, lProgNumber, lProgramYear, sInactiveProgramYear) = False Then
'            If sInactiveProgramYear <> "" Then
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "CategoryCode[" & sCategory & "] ProgramNumber[" & lProgNumber & "] Combination Invalid" & vbCrLf & "          OR Program.InactiveProgramYear: " & lProgramYear & " <= Item.ProgramYear: " & sInactiveProgramYear)
'            Else
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "CategoryCode[" & sCategory & "] ProgramNumber[" & lProgNumber & "] Combination Invalid")
'            End If
'        End If

'        'Check that Only certain SubProgram/Grade's are permissible for certain Program Numbers
'        Dim sSubProgErrMsg As String
'        sSubProgErrMsg = ""
'        If bValidProgramSubProgram(lProgNumber, sSubProgram, lProgramYear, sInactiveProgramYear, sSubProgErrMsg) = False Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, sSubProgErrMsg & " for ProgramNumber[" & sProgNumber & "]")
'        End If

'        If bValidProgramGrade(lProgNumber, sGrade, lProgramYear, sInactiveProgramYear, bGradeRequired) = False Then
'            If sGrade = "" And bGradeRequired = True Then
'                If lProgramYear < 2008 Then 'can be blank for old Items prior to 2008
'                Else
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Grade IS REQUIRED for ProgramNumber[" & sProgNumber & "] for ProgramYear[" & lProgramYear & "]")
'                End If
'            ElseIf sGrade <> "" Or bGradeRequired = False Then
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Grade[" & sGrade & "] INVALID for ProgramNumber[" & sProgNumber & "] for ProgramYear[" & lProgramYear & "]")
'            Else
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Grade[" & sGrade & "] IS REQUIRED for ProgramNumber[" & sProgNumber & "] for ProgramYear[" & lProgramYear & "]")
'            End If
'        End If

'        'C H E C K   A G A I N - after the reclassification is rolled out, only true for ProgramYr =<2007
'        'set Licensor(was RoyaltyID 02/23/2007 hn) = ProgramNumber if ProgramNumber = 113, 501-555
'        If lProgramYear < 2008 Then
'            If lProgNumber = 113 Or (lProgNumber > 500 And lProgNumber < 556) Then
'                sItemArray(lRow, dtCOLPos.lLicensor) = lProgNumber
'                sLicensor = lProgNumber
'            Else
'                '            have to unset Licensor if previously set for programnumbers = 113, 501-555
'                If sItemArray(lRow, dtCOLPos.lLicensor) = "113" Or _
'                    (sItemArray(lRow, dtCOLPos.lLicensor) > "500" And _
'                     sItemArray(lRow, dtCOLPos.lLicensor) < "556") Then

'                    sItemArray(lRow, dtCOLPos.lLicensor) = ""
'                    sLicensor = ""
'                End If
'            End If

'            '        set TradeMarkCopyRight defaults to True if RoyaltyID(Licensor) = 113 Philips, 501 - 555
'            If sLicensor = "113" Or (sLicensor > "500" And sLicensor < "556") Then
'                If sItemArray(lRow, dtCOLPos.lTradeMarkCopyRight) <> "TRUE" Then
'                    sItemArray(lRow, dtCOLPos.lTradeMarkCopyRight) = "TRUE"
'                End If
'            End If

'        Else        '11/19/2008 - hn - new below...
'            If IsBlank(sLicensor) And Not IsBlank(sRoyaltyPercent) And sRoyaltyPercent <> "0" Then
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "RoyaltyPercent=" & sRoyaltyPercent & ", must be blank/zero if Licensor(RoyaltyID) is blank" & vbCrLf)
'            ElseIf Not IsBlank(sLicensor) And sLicensor <> "555" And IsBlank(sRoyaltyPercent) Then
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "RoyaltyPercent cannnot be blank if Licensor(RoyaltyID)=" & sLicensor & vbCrLf)
'            ElseIf sLicensor = "555" And Not IsBlank(sRoyaltyPercent) And sRoyaltyPercent <> "0" Then
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "RoyaltyPercent=" & sRoyaltyPercent & ", must be blank/zero if Licensor(RoyaltyID)=" & sLicensor & vbCrLf)
'            End If

'            'as above, forgot to do this for ProgYear >=2008        set TradeMarkCopyRight defaults to True if RoyaltyID(Licensor) = 113 Philips, 501 - 555
'            If sLicensor = "113" Or (sLicensor > "500" And sLicensor < "556") Then
'                If sItemArray(lRow, dtCOLPos.lTradeMarkCopyRight) <> "TRUE" Then
'                    sItemArray(lRow, dtCOLPos.lTradeMarkCopyRight) = "TRUE"
'                End If
'            End If
'        End If

'        If bPROPOSALFormIndicator = False Then
'            'new validation for concatenated Bag_SpecialEffects instead of Lookup table previously as was in the 'Field' table
'            If sBag_SpecialEffects <> "" Then
'                If bValidBag_SpecialEffects(lRow, sBag_SpecialEffects, sInvalidBag_SpecialEffects) = False Then
'                    Call bBuildErrorMsg(bCreateNEWComponentItems, lRow, lNbrErrors, sRowErrorMsg, "Invalid Bag_SpecialEffects(s)" & vbCrLf & sInvalidBag_SpecialEffects)
'                End If
'            End If

'            If sCertifiedPrinterID <> "" Then
'                If bValidCertifiedPrinterID(lRow, sCertifiedPrinterID, sX_CertPrintDescr, sInvalidCertifiedPrinter) = False Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Invalid Certified Printer ID(s)" & vbCrLf & sInvalidCertifiedPrinter)
'                Else
'                    sX_CertifiedPrinterNames(lRow) = sX_CertPrintDescr
'                End If
'            End If

'            If sTechnology <> "" Then
'                If bValidTechnology(lRow, sTechnology, sX_TechDescr, sInvalidTechnology, lProgramYear) = False Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Invalid Technology Code(s)" & vbCrLf & sInvalidTechnology)
'                Else
'                    sX_Technologies(lRow) = sX_TechDescr
'                End If
'            End If

'            'new validation If Technology 112 OR LightType = Battery Operated then
'            '             ProductBatteriesIncluded must be yes/no
'            If bValidateProductBatteriesIncluded(sTechnology, sLightType, sProductBatteriesIncluded) = False Then
'                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "ProductBatteriesIncluded(Yes/No) must be entered for Technology=112 OR LightType=BatteryOperated")
'            End If
'        End If
'        '2014/05/07 RAS added "DEV" itemstatus accounts to development they do not need all the cost/ price fields.
'        '2014/08/15 RAS added mktbasic and product mgr so to not check fields for these groups
'        If UCase(sItemArray(lRow, dtCOLPos.lItemStatus)) = "DEV" Or (msUserGroup = msMKTGBASIC Or msUserGroup = msPRODUCTMGR) Then
'        Else
'            Select Case lcustomernumber
'                Case gs999PD_ACCOUNT, gs100_ACCOUNT      ', 998_ACCOUNT    'SellPrice must be NULL/Blank for certain 'For Account's
'                    If IsNull(sFOBSellPrice) Or sFOBSellPrice = "" Then
'                    Else
'                        If sFunctioncode = gsNEW_PROPOSAL And (lcustomernumber = gs999PD_ACCOUNT) Then  'Or lCustomerNumber= gs998_ACCOUNT
'                            sItemArray(lRow, dtCOLPos.lSellPrice) = ""
'                        Else
'                            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "FOBSellPrice must be BLANK when 'CustomerNumber' = " & lcustomernumber & " For FC: " & sFunctioncode)
'                        End If
'                    End If
'                Case Else
'                    If IsNull(sFOBSellPrice) Or sFOBSellPrice = "" Or sFOBSellPrice = "0" Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "FOBSellPrice cannot be BLANK or Zero.")
'                    End If
'            End Select
'        End If

'        'CustomerNumber =998, 999 then DevelopComment can be present, else must be blank
'        '07/08/2009 - DevelopComments are now allowed for all Customers and  new Items
'        '    Select Case lCustomerNumber
'        '        Case gs999PD_ACCOUNT, gs998PD_ACCOUNT
'        '            'can be blank or not
'        '        Case Else
'        '            If IsBlank(sDevelopComments) Then
'        '                'it's correct
'        '            ElseIf sFunctionCode = gsNEW_PROPOSAL Then
'        '                sItemArray(lRow, dtCOLPos.lDEVComments) = ""
'        '                sDevelopComments = ""
'        '            ElseIf sFunctionCode <> gsNEW_PROPOSAL Then
'        '                Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "DevelopComments must be BLANK when 'CustomerNumber' = " & lCustomerNumber & ", FC = " & sFunctionCode)
'        '            End If
'        '    End Select

'        If sVendorNumber <> "" And sFactoryNumber <> "" Then
'            'Validate , if Vendor Number changed, that it is only from OR to'1000'
'            If sFunctioncode <> gsNEW_PROPOSAL Then
'                sSQL = "SELECT VendorNumber FROM " & gsItem_Table & " " & _
'                        "WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sRev
'                Call bGetFieldValue(sSQL, vReturnValue)
'                '            If vReturnValue = sThousand Then   'do NOT allow change to VendorNumber except for Function Code = A
'                '                gbVendorChangedFrom1000 = True
'                '            End If
'                If vReturnValue <> sVendorNumber Then
'                    '                If vReturnValue = sThousand Or sVendorNumber = sThousand Then
'                    '                    gbVendorChangedFrom1000 = True
'                    '                Else
'                    '                    If bValidateProposal = False Then
'                    '                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Spreadsheet.VendorNumber " & sVendorNumber & " " & _
'                    '                            "for Function Code:(" & sFunctionCode & _
'                    '                            ")  Change to VendorNumber can only be FROM or TO '" & sThousand & "', (it's = " & vReturnValue & " on Item), please create a New Proposal.(A)")
'                    '                    Else
'                    '                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Change to VendorNumber can only be FROM or TO '" & sThousand & "', (it's = " & vReturnValue & " on Item). Please correct.")
'                    '                    End If
'                    '                End If
'                    If bValidateProposal = False Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, _
'                        "Spreadsheet.VendorNumber " & sVendorNumber & " " & _
'                            "doesn't match Proposal Vendor: " & vReturnValue & _
'                            ". VendorNumber cannot be changed, please create a New Proposal if you want to change VendorNumber.(FunctionCode=A)")
'                    Else
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot change VendorNumber for existing Proposal. If this is a different Vendor create a new Proposal by clicking the 'New Vendor' option for creating an Item on the Proposal Form.")
'                    End If
'                End If
'            End If

'            '---Vendor and Factory must be associated with each other
'            '        If sVendorNumber <> sThousand And sFactoryNumber <> sThousand Then '12/28/2007
'            If bFindVendorFactoryInDB(sVendorNumber, sFactoryNumber) = False Then
'                If bValidateProposal = False Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "VendorNumber[" & sVendorNumber & "] " & _
'                        "is not associated with FactoryNumber[" & sFactoryNumber & "] in the database.")
'                Else
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "VendorNumber[" & sVendorNumber & "] " & _
'                        "is not associated with FactoryNumber[" & sFactoryNumber & "] in the database.")
'                End If
'                lNbrErrors = lNbrErrors + 1
'            End If

'            '--Vendor must be active, can be Inactive for ItemStatus = DS
'            If Not bVendorActive(sVendorNumber) Then
'                If sItemStatus <> "DS" Then
'                    If bValidateProposal = False Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "VendorNumber. Cannot Import/Save; Inactive Vendor Number:" & sVendorNumber)
'                    Else
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot Import/Save Inactive Vendor Number:" & sVendorNumber)
'                    End If
'                End If
'            End If

'            '--Factory must be active, can be Inactive for ItemStatus = DS
'            If Not bFactoryActive(sFactoryNumber) Then
'                If sItemStatus <> "DS" Then
'                    If bValidateProposal = False Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "FactoryNumber. Cannot Import/Save: Inactive Factory Number:" & sFactoryNumber)
'                    Else
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot Import/Save Inactive Factory Number:" & sFactoryNumber)
'                    End If
'                End If
'            End If

'            '--Vendor Factory relationship must be Active
'            If Not bVendorFactoryActiveRelationship(sVendorNumber, sFactoryNumber) Then
'                If sItemStatus <> "DS" Then
'                    If bValidateProposal = False Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot Import/Save Inactive Vendor/Factory Relationship. Vendor[" & sVendorNumber & "] Factory[" & sFactoryNumber & "]")
'                    Else
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot Import/Save Inactive Vendor/Factory Relationship. Vendor[" & sVendorNumber & "] Factory[" & sFactoryNumber & "]")
'                    End If
'                End If
'            End If

'            '-- Factory cannot be prohibited
'            If Not bFactoryProhibited(sFactoryNumber) Then
'                If sItemStatus <> "DS" Then
'                    If bValidateProposal = False Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot Import/Save Prohibited Factory: " & sFactoryNumber)
'                    Else
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot Import/Save Prohibited Factory: " & sFactoryNumber)
'                    End If
'                End If
'            End If

'            '-- Factory cannot be LightSourceOnly - new 10/22/2007
'            If Not bFactoryLightSourceOnly(sFactoryNumber) Then
'                If bValidateProposal = False Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot Import/Save Factory that is LightSourceOnly: " & sFactoryNumber)
'                Else
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Cannot Import/Save Factory that is LightSourceOnly: " & sFactoryNumber)
'                End If
'            End If

'        End If

'        ' If the VendorItemNumber is present Long Description must be present
'        If sVendorItemNumber <> "" Then
'            If sItemArray(lRow, dtCOLPos.lLongDesc) = "" Then
'                If bValidateProposal = False Then
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "LongDescription is required when VendorItemNumber is present.") '02/13/2009 - hn
'                Else
'                    Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "LongDescription is required when VendorItemNumber is present.")
'                End If
'            End If
'        End If

'        'Check when HSNumber is present that DutyPercent is not null
'        If bCheckHSNumberANDDutyPercent(lRow, lRowsOnSheet, dtCOLPos, sItemArray(), sHSNumber, sDutyPercent) = False Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "For HSNumber: " & sHSNumber & " DutyPercent must have a value! " & sDutyPercent)
'        End If

'        'Check that 'Vendor/Factory/VendorItemNumber' combination does not appear more than once on SPREADSHEET
'        ' to prevent more than one new ItemNumber from being created for same combination!
'        If bCheckSpreadsheetForVenFactItemNum(lRow, lRowsOnSheet, dtCOLPos, sItemArray(), _
'                    sItemArray(), sVendorNumber, sFactoryNumber, sVendorItemNumber, sSpreadsheetRow) = False Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "Vendor:" & sVendorNumber & "/Factory:" & sFactoryNumber & _
'                    "/VendorItemNumber:" & sVendorItemNumber & _
'                    ", Combination appears more than once on Spreadsheet. Please correct. Rows: " & sSpreadsheetRow)
'        End If

'        '  check ASSORTMENT ITEMNUMBERS ..... match on last 3 digits with 'For Account'
'        For lCounter = 4 To lAssortmentFields
'            If Microsoft.VisualBasic.Left(sAssortmentArray(1, lCounter), 5) = cITEM_XX Then
'                If sAssortmentArray(lRow, lCounter) <> "" Then
'                    If Microsoft.VisualBasic.Right(sAssortmentArray(lRow, lCounter), 3) <> lcustomernumber Then
'                        Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, " The last 3 digits of Item_Assortments.Item_ " & lCounter - 3 & _
'                                 ": " & Microsoft.VisualBasic.Right(sAssortmentArray(lRow, lCounter), 3) & _
'                                 " ,must match 'For Account': " & lcustomernumber)              '11/25/2008 - hn

'                        'HN THIS PART NOT EXPORTED TO LIVE 3/3/2006
'                        'check that Assortment Item's Vendor = original Item.Item's Vendor
'                        '                If bFindAssortmentVendor(sAssortmentArray(lRow, lCounter), sVendorNumber, sItemVendor) = False Then
'                        '                    sErrorMsg = sErrorMsg & "ROW " & lRow & ": Assortment: " & sAssortmentArray(1, lCounter) & _
'                        '                                ":" & sAssortmentArray(lRow, lCounter) & " Vendor: " & sVendorNumber & _
'                        '                                " NOT = " & "Item Main Item's Vendor: " & sItemVendor & _
'                        '                                " MUST create new Proposal" & vbCrLf
'                        '                    lNbrErrors = lNbrErrors + 1
'                        '                End If

'                        'check that Assortment_Item's Vendor is not with an ItemNumber on the SPREADSHEET where the vendor is being changed
'                        'otherwise the Import Process might change the Vendor and then the validation above is no longer valid
'                        'HN THIS PART NOT EXPORTED TO LIVE 2/27/2006
'                        '                If bSpreadsheetVendorForAssortItemNums(lRow, lRowsOnSheet, udtColPos, _
'                        '                            sITEMSpecsArray(), sITEMArray(), _
'                        '                            sAssortmentArray(lRow, lCounter), sVendorNumber, _
'                        '                            sSpreadsheetRow, sSpreadsheetVendor) = False Then
'                        '                     sErrorMsg = sErrorMsg & "ROW " & lRow & ": Assortment: " & sAssortmentArray(1, lCounter) & _
'                        '                                 ":" & sAssortmentArray(lRow, lCounter) & _
'                        '                                 " Vendor: " & sVendorNumber & " does not match, ROW(s): " & sSpreadsheetRow & "'s Main VendorNumber(s): " & _
'                        '                                 sSpreadsheetVendor & " MUST create new Proposal" & vbCrLf
'                        '                     lNbrErrors = lNbrErrors + 1
'                        '                End If
'                    End If
'                End If
'            End If
'        Next lCounter

'        ' Check to make sure other Item table data exists
'        If sItemArray(lRow, dtCOLPos.lShortDesc) = "" Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "ShortDescription not found when ItemNumber is present")
'        End If
'        If sItemArray(lRow, dtCOLPos.lSeasonCode) = "" Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "SeasonCode not found when ItemNumber is present")
'        End If
'        If sItemArray(lRow, dtCOLPos.lCategoryCode) = "" Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "CategoryCode not found when ItemNumber is present")
'        End If
'        If sItemArray(lRow, dtCOLPos.lProgramNumber) = "" Then
'            Call bBuildErrorMsg(False, lRow, lNbrErrors, sRowErrorMsg, "ProgramNumber not found when ItemNumber is present")
'        End If

'        If sRowErrorMsg <> "" Then
'            sErrorMsg = sErrorMsg & sRowErrorMsg
'            sRowErrorMsg = ""
'        End If

'        bValidateCoreFields = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateCoreFields")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidateCoreFields , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        '    Resume Next '2011/12/21 remove after fixing error
'        Resume ExitRoutine
'    End Function
'Public Function bSaveData(frmThis As Form, ByVal lRowsOnSheet As Long, ByVal lColsOnSheet As Long, _
'                        dtSpreadsheetCOLArray() As typColumn, _
'                        dtSaveArrayCOLPos As typSpecialCOLPos, _
'                        sItemSPECSArray() As String, ByVal lItemSpecsFields As Long, _
'                        sItemArray() As String, ByVal lItemFields As Long, _
'                        sAssortmentArray() As String, ByVal lAssortmentFields As Long, _
'                        ByRef sReturnErrormsg As String, ByRef lReturnErrorRow As Long, _
'                        ByVal bWriteToLog As Boolean, ByVal bUpdateSourceFile As Boolean, _
'                        Optional objExcelLogFile As Excel.Application, _
'                        Optional sExcelLogFileName As String, _
'                        Optional objEXCELImportFile As Excel.Application, Optional sExcelImportFileName As String, _
'                        Optional lNEWProposalNUM As Long) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sFromProcess As String
'        Dim sSQL As String
'        Dim sLatestItemStatus As String
'        Dim sItemStatus As String
'        Dim lRowCounter As Long

'        Dim rsItem As ADODB.Recordset  'these 3 tables MUST be updated simultaneously
'        Dim RSItemSpecs As ADODB.Recordset
'        Dim rsAssortments As ADODB.Recordset

'        Dim sORIGProposalNumber As String
'        Dim sORIGRevNumber As String

'        Dim sDescriptiveMsg As String
'        Dim sDescriptiveMsg1 As String

'        Dim dtSpreadsheetCOLPos As typSpecialCOLPos

'        Dim lOriginalRev As Long
'        Dim sFunctioncode As String
'        Dim sItemNumber As String

'        Dim sProposalNumber As String
'        Dim sRev As String
'        Dim sPrevRev As String

'        Dim sProgramYear As String
'        Dim sPrevProgYear As String
'        Dim sProgram As String
'        Dim sCategory As String
'        Dim scustomer As String
'        Dim sVendor As String
'        Dim sFactory As String
'        Dim sPrevFactory As String
'        Dim lRowChangesFound As Long
'        Dim sSaveRowErr As String
'        Dim smessage As String
'        '2014/03/04 RAS if the workbook name is nothing default a name.
'        If objEXCELImportFile Is Nothing Then
'            objEXCELName = "FromProposal"
'        Else
'            objEXCELName = objEXCELImportFile.ActiveWorkbook.Name  '- -sExcelImportFileName
'        End If
'        bSaveData = False
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            If bWritePrintToLogFile(False, objEXCELName & Space(2), Format(Now(), "yyyymmdd")) = False Then
'            End If
'            If bWritePrintToLogFile(False, objEXCELName & Space(2) & "Starting bSaveData", Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If

'        If bPROPOSALFormIndicator = True Then
'            sFromProcess = "P"
'        Else
'            sFromProcess = "I"
'        End If

'        rsItem = New ADODB.Recordset
'        RSItemSpecs = New ADODB.Recordset
'        rsAssortments = New ADODB.Recordset

'        If bUpdateSourceFile Then
'            ' Get the positions of the spreadsheet columns to update during the import process
'            If bGetSpecialSpreadsheetCOLPositions(dtSpreadsheetCOLPos, dtSpreadsheetCOLArray(), lColsOnSheet) = False Then
'                sReturnErrormsg = "Could not retrieve special column positions from the source file"
'                '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            End If
'        End If

'        ' Loop through the spreadsheet rows with a function code
'        For lRowCounter = glDATA_START_ROW To lRowsOnSheet
'            RowChangesARRAY(lRowCounter) = 0
'            lRowChangesFound = 0
'            'set background to green for msg below
'            If Len(sItemArray(lRowCounter, glFunctionCode_ColPos)) = 0 Then
'                '2013/04/25 -HN- need to get even a blank function code, to prevent extra Inspection records being created at end of this procedure
'                sFunctioncode = sItemArray(lRowCounter, glFunctionCode_ColPos)
'                Call bUpdateStatusMessage(frmThis, "Skipping Row " & CStr(lRowCounter) & " of  " & lRowsOnSheet & "...")
'            Else
'                Call bUpdateStatusMessage(frmThis, "Saving Row " & CStr(lRowCounter) & " of " & lRowsOnSheet & "...", True, vbGreen)

'                If bPrepareSaveArray(lRowCounter, dtSaveArrayCOLPos, _
'                            sItemSPECSArray(), sItemArray(), sAssortmentArray(), _
'                            sORIGProposalNumber, sORIGRevNumber, sDescriptiveMsg, lNEWProposalNUM) = False Then
'                    sReturnErrormsg = sDescriptiveMsg & ". Could not prepare to save record in memory"
'                    lReturnErrorRow = lRowCounter
'                    '                '2014/01/17 RAS changing the GOTO Exit reoutine to GOTO SaveError.  this will then have the code  goe in the validation label.'TODO - GoTo Statements are redundant in .NET
'                    '                GoTo SaveTableError'TODO - GoTo Statements are redundant in .NET
'                    '               ' GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET

'                End If

'                ' Because each record is bridged between three tables,
'                ' only update after the record passes bSaveRow for each table

'                '--- Item Table ----

'                '2013/05/30 - HN- added sSaveRowErr
'                If bSaveRow(lRowCounter, sItemArray(), lItemFields, gsItem_Table, rsItem, lRowChangesFound, sSaveRowErr) = False Then
'                    sReturnErrormsg = "Problem in saving Item table data: " & sSaveRowErr       '2013/05/30
'                    lReturnErrorRow = lRowCounter
'                    '                GoTo SaveTableError                                                         '2013/05/30'TODO - GoTo Statements are redundant in .NET
'                Else
'                    '2012/01/04 - write changes to ItemFieldHistory table
'                End If

'                sFunctioncode = sItemArray(lRowCounter, glFunctionCode_ColPos)
'                sProposalNumber = sItemArray(lRowCounter, glProposal_ColPos)
'                sRev = sItemArray(lRowCounter, glREV_ColPos)
'                sPrevRev = sItemArray_Orig(lRowCounter, glREV_ColPos)

'                sCategory = sItemArray(lRowCounter, dtSaveArrayCOLPos.lCategoryCode)
'                sProgram = sItemArray(lRowCounter, dtSaveArrayCOLPos.lProgramNumber)

'                sProgramYear = sItemArray(lRowCounter, dtSaveArrayCOLPos.lProgramYear)
'                sPrevProgYear = sItemArray_Orig(lRowCounter, dtSaveArrayCOLPos.lProgramYear)

'                scustomer = sItemArray(lRowCounter, dtSaveArrayCOLPos.lcustomernumber)
'                sVendor = sItemArray(lRowCounter, dtSaveArrayCOLPos.lVendorNumber)
'                sFactory = sItemArray(lRowCounter, dtSaveArrayCOLPos.lFactoryNumber)
'                sPrevFactory = sItemArray_Orig(lRowCounter, dtSaveArrayCOLPos.lFactoryNumber)   'if factory changes then inserts ItemInspection record
'                sItemStatus = sItemArray(lRowCounter, dtSaveArrayCOLPos.lItemStatus)

'                sItemStatus_ORIG = sItemArray_Orig(lRowCounter, dtSaveArrayCOLPos.lItemStatus)
'                '2013/07/25 -HN- because this is a global variable in modSpreadSheet
'                'it was keeping the value of the last row, obtained in bValidateCoreFields

'                If bPROPOSALFormIndicator = False Then
'                    sItemNumber = sItemArray_Orig(2, dtSaveArrayCOLPos.lItemNumber)
'                Else
'                    sItemNumber = sItemArray(2, dtSaveArrayCOLPos.lItemNumber)
'                End If

'                '--- ItemSpecs table ----
'                If bPROPOSALFormIndicator = False Then
'                    'saving values from Import Spreadsheet - as it has always been done
'                    If bSaveRow(lRowCounter, sItemSPECSArray(), lItemSpecsFields, gsItemSpecs_Table, RSItemSpecs, lRowChangesFound, sSaveRowErr) = False Then
'                        sReturnErrormsg = "Problem in saving ItemSpecs table data: " & sSaveRowErr        '2013/05/30
'                        lReturnErrorRow = lRowCounter
'                        '                    GoTo SaveTableError                                                         '2013/05/30'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        '                    write changes to ItemFieldHistory table
'                    End If
'                Else
'                    'saving changes to ItemSpecs table from Proposal form
'                    If sFunctioncode = gsNEW_ITEM_NBR And bPROPOSALFormIndicator = True Then
'                        '                    sPrevRev = 0
'                    End If

'                    If Form_frmProposal.SpecChangesInProgress = False And sFunctioncode <> "A" Then      '2013/06/04 -HN- check that there are no missing ItemSpecs records
'                        Dim rsMISSING As ADODB.Recordset
'                        rsMISSING = New ADODB.Recordset
'                        Dim SQL As String
'                        SQL = "SELECT * FROM TempItemSpecs WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sPrevRev
'                        rsMISSING.Open SQL
'                        CurrentProject.Connection()
'                        Dim adOpenStatic As Object
'                        Dim adLockReadOnly As Object

'                        If rsMISSING.EOF Then
'                            MsgBox("Missing ItemSpecs and/or Item_Assortment record(s) for this Proposal: " & sProposalNumber & " and Rev: " & sPrevRev, vbCritical, "See System Administrator!")
'                            rsMISSING = Nothing
'                            '                        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        Else
'                            rsMISSING = Nothing
'                        End If

'                        ' If rsMISSING.State <> 0 Then rsMISSING.Close  '2014/01/09 RAS adding statement based on SJM, she just closed I am checking first to see if it is open
'                    End If
'                    If Form_frmProposal.SpecChangesInProgress = True Or bNewRevFromProposal = True Then
'                        If bSaveChangedItemSpecsValues(sFunctioncode, RSItemSpecs, sProposalNumber, sItemArray_Orig(2, 2), sPrevRev, sRev, bNewRevFromProposal, sItemNumber, lRowChangesFound) = False Then
'                        End If
'                    End If
'                End If

'                'Item_Assortments table
'                If bPROPOSALFormIndicator = False Then
'                    'save Item_Assortment values from spreadsheet
'                    If bSaveRow(lRowCounter, sAssortmentArray(), lAssortmentFields, gsItem_Assortments_Table, rsAssortments, lRowChangesFound, sSaveRowErr) = False Then
'                        sReturnErrormsg = "Problem in saving Item_Assortments table data: " & sSaveRowErr       '2013/05/30"
'                        lReturnErrorRow = lRowCounter
'                        '                    GoTo SaveTableError                                                                     '2013/05/30'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        'write changes to ItemFieldHistory table
'                    End If
'                End If

'                If bNewRevFromProposal = True And gbADDNewProposal = False And bPROPOSALFormIndicator = True Then
'                    If bProposalNEWItemfromList = True Or bNewItemFromProposal = True Then
'                    Else
'                        Call bReload_ORIGAssortmentValues(sItemArray_Orig(lRowCounter, glProposal_ColPos), sItemArray_Orig(lRowCounter, glREV_ColPos))
'                        Application.DoEvents()
'                    End If
'                    Call bSAVENEWAssortments(rsAssortments, sItemArray(lRowCounter, glProposal_ColPos), sItemArray(lRowCounter, glREV_ColPos))
'                    bNewRevFromProposal = False
'                    Form_frmProposal.mlbAssortmentsChanged = False

'                ElseIf bPROPOSALFormIndicator = True Then
'                    If gbADDNewProposal = True Then
'                        Call bSAVENEWAssortments(rsAssortments, sItemArray(lRowCounter, glProposal_ColPos), sItemArray(lRowCounter, glREV_ColPos))
'                    Else
'                        If Form_frmProposal.mlbAssortmentsChanged = True Then
'                            lRowChangesFound = lRowChangesFound + 1
'                            Call bSAVENEWAssortments(rsAssortments, sItemArray_Orig(lRowCounter, glProposal_ColPos), sItemArray_Orig(lRowCounter, glREV_ColPos))
'                        End If
'                        Form_frmProposal.mlbAssortmentsChanged = False
'                    End If
'                End If

'                '2010/02/12 - if MaterialChanges made only then also have to update RevisedDate etc
'                If bPROPOSALFormIndicator = True Then
'                    If Form_frmProposal.MaterialChangesInProgress = True Then
'                        lRowChangesFound = lRowChangesFound + 1
'                    End If
'                End If

'                '           'update the 3 main tables - have to Update before ItemMaterial can be updated
'                '2013/04/30 - HN- do in bSaveRow

'                Select Case sFunctioncode                                   'keeping track of actual changes made, in case of error in import
'                    Case "C"
'                        lImportChangedProposals = lImportChangedProposals + 1
'                    Case "R"
'                        lImportRevisedProposals = lImportRevisedProposals + 1
'                    Case "A"
'                        lImportAddedProposals = lImportAddedProposals + 1
'                End Select


'                'write changes to ItemFieldHistory table per row, to avoid not doing it when Import fails on a row.
'                Call bUpdateStatusMessage(frmThis, "Saving Row(" & lRowCounter & "), changes to ItemFieldHistory table...", True, vbGreen)           '2010/11/19
'                '2014/01/14 RAS Adding info to Trace message
'                If glTraceFlag = True Then
'                    If bWritePrintToLogFile(False, objEXCELName & Space(6) & "Starting bCheckItemFieldsChangedPerRow, for Row: " & lRowCounter, Format(Now(), "yyyymmdd")) = False Then
'                    End If
'                End If

'                If bCheckItemFieldsChangedPerRow(lRowCounter, RowChangesARRAY(), sFromProcess, sItemSpecsArray_ORIG(), sItemSPECSArray(), lItemSpecsFields, _
'                            sItemArray_Orig(), sItemArray(), lItemFields, _
'                            sAssortmentArray_ORIG(), sAssortmentArray(), lAssortmentFields, _
'                            lRowsOnSheet, dtSaveArrayCOLPos) = False Then
'                    '                GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'                End If


'                If lMaxMaterialCols > 0 And bPROPOSALFormIndicator = False Then
'                    'frmProposal form has it own save routine from the TempItemMaterial table
'                    'write Material changes to ItemMaterials

'                    lOriginalRev = SpreadsheetMaterialValuesX(lRowCounter, 1).lRev

'                    ' if not a new revision then lNewOrChangedREV = original revision
'                    If bUpdateRowItemMaterial(lRowCounter, sFunctioncode, sProposalNumber, _
'                                        lOriginalRev, sRev, sItemNumber, lRowChangesFound) = False Then
'                        bItemMaterialError = True
'                        sDescriptiveMsg = "Error in updating ItemMaterial table, fix Material Errors, with FC=C, or go to Proposal Form"
'                        lReturnErrorRow = lRowCounter
'                        ''                    GoTo ExitRoutine               '01/15/2008'TODO - GoTo Statements are redundant in .NET
'                    End If

'                    '            for FC=R and no Material Columns on Import spreadsheet
'                    '            then didn't create these from import revs for new rev
'                    '            for FC = A User has to enter Material Values in Material cols
'                ElseIf lMaxMaterialCols = 0 And bPROPOSALFormIndicator = False Then
'                    If sFunctioncode = "R" Then
'                        If bInsertRowItemMaterial(lRowCounter, sFunctioncode, sProposalNumber, _
'                                        sPrevRev, sRev, sItemNumber) = False Then
'                            bItemMaterialError = True
'                            sDescriptiveMsg = "Error updating ItemMaterial table, for FC=R, fix on Proposal Form for this Proposal:" & sProposalNumber
'                            lReturnErrorRow = lRowCounter
'                        End If
'                    End If

'                End If
'                '2014/01/23 RAS moving this down below all the updates
'                '            If bWriteToLog Then
'                '             '2014/01/14 RAS Adding info to Trace message
'                '                If glTraceFlag = True Then
'                '                    If bWritePrintToLogFile(False, objEXCELName & Space(4) & "Starting to Write to Excel log, for Row: " & lRowCounter, Format(Now(), "yyyymmdd")) = False Then
'                '                    End If
'                '                End If
'                '
'                '                ' Write a record of the update to the Excel log file
'                '                If lRowChangesFound = 0 And sFunctionCode = "C" Then        'inform user if no changes were found on a row
'                '                    sDescriptiveMsg = "For FC=C; no changes found, nothing updated!" & sDescriptiveMsg
'                '                End If
'                '                If bWriteToExcelLog(lRowCounter, objExcelLogFile, sExcelLogFileName, _
'                '                                dtSaveArrayCOLPos, sItemSPECSArray(), sItemArray(), _
'                '                                sORIGProposalNumber, sORIGRevNumber, sDescriptiveMsg) = False Then
'                '                    sReturnErrorMsg = "Could not add data to the log file"
'                '                    lReturnErrorRow = lRowCounter
'                ''                '2014/01/17 RAS changing the GOTO Exit reoutine to GOTO SaveError.  this will then have the code  goe in the validation label.'TODO - GoTo Statements are redundant in .NET
'                ''                GoTo SaveTableError'TODO - GoTo Statements are redundant in .NET
'                ''                '   GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                '                End If
'                '            End If
'                '2014/01/23 RAS moving this down so it will write to the source file after all the updates for the row.
'                '            sDescriptiveMsg = ""
'                '
'                '            If bUpdateSourceFile Then
'                '                ' Update the following columns in the Source file (Import spreadsheet):
'                '                '   FunctionCode, ProposalNumber, Rev, ItemNumber, UPCNumber
'                '                '   Plus the Columns overwritten in bOverwriteDuplicateFields
'                '                '   Plus the Display Name (X_) columns
'                '
'                '                If glSHADING <> NoShadingX Then
'                '                    Call bUpdateStatusMessage(frmThis, "Updating and Shading Spreadsheet Row " & CStr(lRowCounter) & " of  " & lRowsOnSheet & " ...")  '2010/11/19
'                '                Else
'                '                    Call bUpdateStatusMessage(frmThis, "Updating Spreadsheet Row " & CStr(lRowCounter) & " of  " & lRowsOnSheet & " ...")  '2010/11/19
'                '                End If
'                '
'                '                RowChangesARRAY(lRowCounter) = lRowChangesFound
'                '                 '2014/01/14 RAS Adding info to Trace message
'                '                If glTraceFlag = True Then
'                '                    If bWritePrintToLogFile(False, objEXCELName & Space(4) & "Updating Source Excel File, for Row: " & lRowCounter, Format(Now(), "yyyymmdd")) = False Then
'                '                    End If
'                '                End If
'                '
'                '                If bUpdateExcelSourceFile(RowChangesARRAY(), lRowCounter, lRowsOnSheet, sExcelImportFileName, _
'                '                                        objEXCELImportFile, _
'                '                                        dtSpreadsheetCOLArray(), dtSpreadsheetCOLPos, _
'                '                                        lColsOnSheet, _
'                '                                        lItemSpecsFields, lItemFields, lAssortmentFields, _
'                '                                        sItemSPECSArray(), sItemArray(), _
'                '                                        dtSaveArrayCOLPos, frmThis) = False Then
'                '                    sReturnErrorMsg = "Could not update the source file"
'                '                    lReturnErrorRow = lRowCounter
'                ''                    '2014/01/17 RAS changing the GOTO Exit reoutine to GOTO SaveError.  this will then have the code  goe in the validation label.'TODO - GoTo Statements are redundant in .NET
'                ''                    GoTo SaveTableError'TODO - GoTo Statements are redundant in .NET
'                ''                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                '                End If
'                '            End If

'                'For FC=C,R, ItemNumber=A; update all Revs for Proposal with the new ItemNumber
'                If sFunctioncode <> "A" Then

'                    If bUpdateOtherRevsNewItemNumbers(sProposalNumber, sItemArray(lRowCounter, dtSaveArrayCOLPos.lItemNumber), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lUPCNumber), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lMPKUPC), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lIPKUPC)) = False Then
'                        sReturnErrormsg = "Unable to update other Revisions with New ItemNumber assigned for Proposal[" & sProposalNumber & "]"
'                        lReturnErrorRow = lRowCounter
'                        '                '2014/01/17 RAS changing the GOTO Exit reoutine to GOTO SaveError.  this will then have the code  goe in the validation label.'TODO - GoTo Statements are redundant in .NET
'                        '                GoTo SaveTableError'TODO - GoTo Statements are redundant in .NET
'                        '                 '   GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If
'                End If

'                '--------  Set Item Status = Latest Rev's Item Status for a ProgramYear ----------
'                '          Also set Category, Season, CustomeritemNumber & Program = Latest Rev's Values for a Prog Year
'                '          only 1 TempItem # allowed per proposal, set to latest Revs Value, if prev years are blank leave those

'                If bUpdateItemStatusETC(sItemArray, lRowCounter, _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lProgramYear), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lItemStatus), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lCustomerItemNumber), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lCustomerCartonUPCNumber), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lTempItemNumber), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lCategoryCode), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lFactoryNumber), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lSeasonCode), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lProgramNumber), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lSubProgram), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lGrade), _
'                        sItemArray(lRowCounter, dtSaveArrayCOLPos.lClass), _
'                        frmThis, sDescriptiveMsg1, sLatestItemStatus) = False Then
'                    sReturnErrormsg = "Could not update Latest Rev's Item Status/Season/Cat/Program/CustomerItemNumber for Previous Revs"
'                    lReturnErrorRow = lRowCounter
'                    '            '2014/01/17 RAS changing the GOTO Exit reoutine to GOTO SaveError.  this will then have the code  goe in the validation label.'TODO - GoTo Statements are redundant in .NET
'                    '                GoTo SaveTableError'TODO - GoTo Statements are redundant in .NET
'                    '               ' GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If
'                ' 2014/01/15 RAS moved this big if statement inside of another if statement.
'                'Create Inspection Records for an Item when ItemStatus is changed to = ORD
'                If Not IsBlank(sFunctioncode) Then '2013/07/25 -HN- code should probably be moved up before the End If just above
'                    If sItemStatus = "ORD" And sItemStatus_ORIG <> "ORD" And sFunctioncode = "C" Then   '2012/07/09 - add Inspection Records when ItemStatus is changed = 'ORD'
'                        CreateInspectionsAllPerItem(sProgramYear, scustomer, sProposalNumber, sRev, sCategory, sProgram, sVendor, sFactory, False)

'                    ElseIf sItemStatus = "ORD" And sItemStatus_ORIG = "ORD" And sFunctioncode = "R" And sFactory = sPrevFactory And sProgramYear = sPrevProgYear Then  '2013/07/25 -HN
'                        UpdateItemInspectionsFromProposal(sProgramYear, scustomer, sProposalNumber, sRev, sPrevRev, sCategory, sProgram, sVendor, sFactory)

'                    ElseIf sItemStatus = "ORD" And sItemStatus_ORIG = "ORD" And sFunctioncode = "R" And sFactory = sPrevFactory And sProgramYear <> sPrevProgYear Then  '2013/07/25 -HN
'                        CreateInspectionsAllPerItem(sProgramYear, scustomer, sProposalNumber, sRev, sCategory, sProgram, sVendor, sFactory, False)

'                    ElseIf sItemStatus = "ORD" And sItemStatus_ORIG <> "ORD" And sFunctioncode = "R" Then
'                        CreateInspectionsAllPerItem(sProgramYear, scustomer, sProposalNumber, sRev, sCategory, sProgram, sVendor, sFactory, False)

'                    ElseIf sItemStatus = "ORD" And sFunctioncode = "A" Then
'                        CreateInspectionsAllPerItem(sProgramYear, scustomer, sProposalNumber, sRev, sCategory, sProgram, sVendor, sFactory, False)

'                        '2013/04/25 -HN- added test for blank FunctionCode
'                    ElseIf Not IsBlank(sFunctioncode) And sItemStatus = "ORD" And sFactory <> sPrevFactory Then
'                        '2012/04/24 - update ItemInspection record if no Order is present
'                        '2013/07/17 -HN- insert new rec if new FactoryNumber?
'                        CreateInspectionsAllPerItem(sProgramYear, scustomer, sProposalNumber, sRev, sCategory, sProgram, sVendor, sFactory, False)
'                        '            UpdateInspectionFactory sProgramYear, sCustomer, sFactory, sProposalNumber, sRev

'                    ElseIf sItemStatus <> "ORD" And sItemStatus <> "" Then 'And sItemStatus_ORIG = "ORD" Then   '2012/07/18
'                        'at this point previous validation has checked that this Item is not on an order, therefore the Inspection can be deleted
'                        If Not IsBlank(sProposalNumber) And Not IsBlank(sRev) Then
'                            DeleteInspectionByProposalRev(sProposalNumber, sRev)              '2012/04/20 - delete when Itemstatus is changed back to <> 'ORD'
'                        End If
'                    Else
'                    End If
'                End If

'                '05/07/2009 - VendorNumber cannot be changed for an existing Proposal
'                'set Vendor Number to new Number changed from 1000
'                '            If gbVendorChangedFrom1000 = True Then
'                '                If bSaveNewVendorFrom1000(sItemArray, lROWCounter, dtSaveArrayCOLPos) = False Then
'                '                    sReturnErrorMsg = "Could not update changed Vendor(from 1000) for other Revs"
'                '                    lReturnErrorRow = lROWCounter
'                '                End If
'                '                gbVendorChangedFrom1000 = False
'                '            End If
'                '201/01/23 RAS Added this   to write to the log file after all the updates
'                If bWriteToLog Then
'                    '2014/01/14 RAS Adding info to Trace message
'                    If glTraceFlag = True Then
'                        If bWritePrintToLogFile(False, objEXCELName & Space(4) & "Starting to Write to Excel log, for Row: " & lRowCounter, Format(Now(), "yyyymmdd")) = False Then
'                        End If
'                    End If

'                    ' Write a record of the update to the Excel log file
'                    If lRowChangesFound = 0 And sFunctioncode = "C" Then        'inform user if no changes were found on a row
'                        sDescriptiveMsg = "For FC=C; no changes found, nothing updated!" & sDescriptiveMsg
'                    End If
'                    If bWriteToExcelLog(lRowCounter, objExcelLogFile, sExcelLogFileName, _
'                                    dtSaveArrayCOLPos, sItemSPECSArray(), sItemArray(), _
'                                    sORIGProposalNumber, sORIGRevNumber, sDescriptiveMsg) = False Then
'                        sReturnErrormsg = "Could not add data to the log file"
'                        lReturnErrorRow = lRowCounter
'                        '                '2014/01/17 RAS changing the GOTO Exit reoutine to GOTO SaveError.  this will then have the code  goe in the validation label.'TODO - GoTo Statements are redundant in .NET
'                        '                GoTo SaveTableError'TODO - GoTo Statements are redundant in .NET
'                        '                '   GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If
'                End If

'                '201/01/23 RAS Added this   to write to the source  file after all the updates
'                sDescriptiveMsg = ""

'                If bUpdateSourceFile Then
'                    ' Update the following columns in the Source file (Import spreadsheet):
'                    '   FunctionCode, ProposalNumber, Rev, ItemNumber, UPCNumber
'                    '   Plus the Columns overwritten in bOverwriteDuplicateFields
'                    '   Plus the Display Name (X_) columns

'                    If glSHADING <> NoShadingX Then
'                        Call bUpdateStatusMessage(frmThis, "Updating and Shading Spreadsheet Row " & CStr(lRowCounter) & " of  " & lRowsOnSheet & " ...")  '2010/11/19
'                    Else
'                        Call bUpdateStatusMessage(frmThis, "Updating Spreadsheet Row " & CStr(lRowCounter) & " of  " & lRowsOnSheet & " ...")  '2010/11/19
'                    End If

'                    RowChangesARRAY(lRowCounter) = lRowChangesFound
'                    '2014/01/14 RAS Adding info to Trace message
'                    If glTraceFlag = True Then
'                        If bWritePrintToLogFile(False, objEXCELName & Space(4) & "Updating Source Excel File, for Row: " & lRowCounter, Format(Now(), "yyyymmdd")) = False Then
'                        End If
'                    End If

'                    If bUpdateExcelSourceFile(RowChangesARRAY(), lRowCounter, lRowsOnSheet, sExcelImportFileName, _
'                                            objEXCELImportFile, _
'                                            dtSpreadsheetCOLArray(), dtSpreadsheetCOLPos, _
'                                            lColsOnSheet, _
'                                            lItemSpecsFields, lItemFields, lAssortmentFields, _
'                                            sItemSPECSArray(), sItemArray(), _
'                                            dtSaveArrayCOLPos, frmThis) = False Then
'                        sReturnErrormsg = "Could not update the source file"
'                        lReturnErrorRow = lRowCounter
'                        '                    '2014/01/17 RAS changing the GOTO Exit reoutine to GOTO SaveError.  this will then have the code  goe in the validation label.'TODO - GoTo Statements are redundant in .NET
'                        '                    GoTo SaveTableError'TODO - GoTo Statements are redundant in .NET
'                        '                    'GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If
'                End If
'                If sDescriptiveMsg <> "" Or sDescriptiveMsg1 <> "" Then
'                    If bWriteToLog Then
'                        ' Concatenate  record of the Item_Status update etc to the Excel log file
'                        sDescriptiveMsg = sDescriptiveMsg & sDescriptiveMsg1 & " for Prog Yr: " & sItemArray(lRowCounter, dtSaveArrayCOLPos.lProgramYear)
'                        If bWriteToExcelLog(lRowCounter, objExcelLogFile, sExcelLogFileName, _
'                            dtSaveArrayCOLPos, sItemSPECSArray(), sItemArray(), _
'                            sORIGProposalNumber, sORIGRevNumber, sDescriptiveMsg, True) = False Then
'                            sReturnErrormsg = "Could not add data to the log file; when updating ItemStatus"
'                            lReturnErrorRow = lRowCounter
'                            '                        '2014/01/17 RAS changing the GOTO Exit reoutine to GOTO SaveError.  this will then have the code  goe in the validation label.'TODO - GoTo Statements are redundant in .NET
'                            '                        GoTo SaveTableError'TODO - GoTo Statements are redundant in .NET
'                            '                        ' GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                        End If
'                    End If
'                End If
'                '2014/01/14 RAS Adding info to Trace message
'                If glTraceFlag = True Then
'                    If bWritePrintToLogFile(False, objEXCELName & Space(4) & "Completed bSaveData for Row: " & lRowCounter, Format(Now(), "yyyymmdd")) = False Then
'                    End If
'                    If bWritePrintToLogFile(False, objEXCELName & Space(4), Format(Now(), "yyyymmdd")) = False Then
'                    End If
'                End If
'            End If  'end of if statement for checking if FunctionCode is blank
'            '2014/01/15 RAS the inspection records if statement was here. moved it up to only check the row with  function code.
'        Next lRowCounter


'        If bUpdateSourceFile Then
'            '2014/01/14 RAS Adding info to Trace message
'            If glTraceFlag = True Then
'                If bWritePrintToLogFile(False, objEXCELName & Space(4) & "Updating Source Spreadsheet after looping thru all the rows.", Format(Now(), "yyyymmdd")) = False Then
'                End If
'            End If
'            Call bUpdateStatusMessage(frmThis, "Now updating spreadsheet. If Refresh Photo option was chosen, it may take several minutes to update all the photos for very large spreadsheets... ", True, vbGreen)
'            'Remove the "END" if successful
'            If bUpdateExcelSourceFile(RowChangesARRAY(), lRowsOnSheet + 1, lRowsOnSheet, sExcelImportFileName, _
'                            objEXCELImportFile, _
'                            dtSpreadsheetCOLArray(), dtSpreadsheetCOLPos, _
'                            lColsOnSheet, _
'                            lItemSpecsFields, lItemFields, lAssortmentFields, _
'                            sItemSPECSArray(), sItemArray(), _
'                            dtSaveArrayCOLPos, frmThis) = False Then
'            MsgBox ("Could not remove 'END' from the Import Spreadsheet!"), vbInformation + vbMsgBoxSetForeground
'            End If
'        End If

'ValidateProposalRev_Exist:
'        '  ' On Error GoTo ExitRoutine'TODO - On Error must be replaced with Try, Catch, Finally
'        'loop thru the log file and check those rows
'        '2014/01/07 RAS checking to see if the changes are really in the three table for the proposal and rev
'        ' throw up an error message.   in the future do not right to the history table or spreadsheet or roll those back.
'        If bPROPOSALFormIndicator = False Then
'            '2014/01/14 RAS Adding info to Trace message
'            If glTraceFlag = True Then
'                If bWritePrintToLogFile(False, objEXCELName & Space(5) & "Checking that the Proposals Revs exist in DB4 for the three main tables. Else remove them.", Format(Now(), "yyyymmdd")) = False Then
'                End If
'            End If
'            Call bUpdateStatusMessage(frmThis, "Checking imported rows.... ", True, vbGreen)
'            Dim iarray As Integer
'            Dim sSQLcheck As String
'            Dim missingrecordflag As Boolean
'            Dim errormessageflag As Boolean
'            Dim bMultipleErrors As Boolean
'            bMultipleErrors = False
'            '2014/01/17 RAS closing recordsets if still open.
'            If Not rsItem Is Nothing Then
'                If rsItem.State <> 0 Then rsItem.Close()
'            End If
'            If RSItemSpecs.State <> 0 Then RSItemSpecs.Close()
'            If rsAssortments.State <> 0 Then rsAssortments.Close()
'            For iarray = 1 To lRowsOnSheet - 1
'                Call bUpdateStatusMessage(frmThis, "Checking if Proposal Rev exists for imported row... " & iarray, True, vbGreen)
'                Dim myfunctioncode As String
'                myfunctioncode = sItemArray(iarray, 1)
'                missingrecordflag = False
'                If IsBlank(myfunctioncode) = False And Not myfunctioncode = "FunctionCode" Then

'                    Dim checkProposalNumber As Long
'                    Dim checkRevnumber As Long
'                    If sItemArray(iarray, 2) = "" Then
'                        checkProposalNumber = 0
'                    Else
'                        checkProposalNumber = sItemArray(iarray, 2)
'                    End If
'                    If sItemArray(iarray, 3) = "" Then
'                        checkRevnumber = 0
'                    Else
'                        checkRevnumber = sItemArray(iarray, 3)
'                    End If

'                    Debug.Print("Now processing row: " & iarray & " for Proposal " & checkProposalNumber & " And Rev = " & checkRevnumber)
'                    sSQLcheck = ""
'                    ' just check to see if the proposal rev is in the tables.
'                    ' this is for item table
'                    sSQLcheck = "SELECT ProposalNumber, Rev FROM ITEM With (nolock) WHERE ProposalNumber = " & checkProposalNumber & " And Rev = " & checkRevnumber

'                    rsItem.Open(sSQLcheck, SSDataConn, adOpenKeyset, adLockPessimistic)
'                    Application.DoEvents()
'                    If rsItem.EOF Then
'                        missingrecordflag = True  ' there is no propsal rev throw up an error message.
'                    Else
'                        ' the Proposal Rev exist
'                    End If
'                    Application.DoEvents()
'                    If rsItem.State <> 0 Then rsItem.Close()

'                    ' this is for itemspecs table
'                    sSQLcheck = ""
'                    sSQLcheck = "SELECT ProposalNumber, Rev FROM ITEMSPECS With (nolock) WHERE ProposalNumber = " & checkProposalNumber
'                    sSQLcheck = sSQLcheck & " AND Rev = " & checkRevnumber
'                    RSItemSpecs.Open(sSQLcheck, SSDataConn, adOpenKeyset, adLockPessimistic)
'                    Application.DoEvents()
'                    If RSItemSpecs.EOF Then
'                        missingrecordflag = True ' there is no propsal rev throw up an error message
'                    Else
'                        ' the Proposal Rev exist
'                    End If
'                    Application.DoEvents()
'                    If RSItemSpecs.State <> 0 Then RSItemSpecs.Close()

'                    ' this is for item assorments table
'                    sSQLcheck = ""
'                    sSQLcheck = "SELECT ProposalNumber, Rev FROM Item_Assortments With (nolock) WHERE ProposalNumber = " & checkProposalNumber
'                    sSQLcheck = sSQLcheck & " AND Rev = " & checkRevnumber
'                    rsAssortments.Open(sSQLcheck, SSDataConn, adOpenKeyset, adLockPessimistic)
'                    Application.DoEvents()
'                    If rsAssortments.EOF Then
'                        missingrecordflag = True  ' there is no propsal rev throw up an error message.
'                    Else
'                        ' the Proposal Rev exist
'                    End If
'                    Application.DoEvents()
'                    If rsAssortments.State <> 0 Then rsAssortments.Close()

'                End If
'                If missingrecordflag = True Then
'                    bSaveData = False
'                    errormessageflag = True
'                    Dim lcheckRecordsAffected As Long
'                    If UCase(myfunctioncode) <> "C" Then
'                        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'                            smessage = "In Rolling back Proposal Rev for row :" & iarray & " .One of the main tables did not update correctly."
'                            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'                            End If
'                        End If
'                        sSQLcheck = ""
'                        sSQLcheck = "DELETE FROM ItemMaterial WHERE ProposalNumber = " & checkProposalNumber & " AND Rev = " & checkRevnumber
'                        sSQLcheck = sSQLcheck & vbCrLf & "DELETE FROM Item_Assortments WHERE ProposalNumber = " & checkProposalNumber & " AND Rev = " & checkRevnumber
'                        sSQLcheck = sSQLcheck & vbCrLf & "DELETE FROM ItemSPECS WHERE ProposalNumber = " & checkProposalNumber & " AND Rev = " & checkRevnumber
'                        sSQLcheck = sSQLcheck & vbCrLf & "DELETE FROM Item WHERE ProposalNumber = " & checkProposalNumber & " AND Rev = " & checkRevnumber
'                        sSQLcheck = sSQLcheck & vbCrLf & "DELETE FROM ItemFieldHistory WHERE ProposalNumber = " & checkProposalNumber & " AND Rev = " & checkRevnumber
'                        SSDataConn.Execute(sSQLcheck, lcheckRecordsAffected)
'                        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'                            smessage = "In bSaveData-CheckingPropRev  Rolled back Rev  for row :" & lReturnErrorRow & "ProposalNumber = " & checkProposalNumber & " AND Rev = " & checkRevnumber
'                            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'                            End If
'                        End If
'                    End If
'                    ' Write a record of the update to the Excel log file
'                    sDescriptiveMsg = "Proposal/Rev was not updated. Please reimport!" & sDescriptiveMsg
'                    If bUpdateExcellogFile(iarray, sDescriptiveMsg, "", objExcelLogFile, sExcelLogFileName, 14) = False Then
'                        sReturnErrormsg = "Could not update data to the log file"
'                        lReturnErrorRow = iarray
'                        '2014/01/17 RAS changed this to go to the error handler from the exit routine
'                        '                         GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If
'                    If bRollBackExcelSourceFile(iarray, sItemArray_Orig(), sItemArray(), objEXCELImportFile, _
'                                             sExcelImportFileName, lColsOnSheet, dtSaveArrayCOLPos, frmThis, RowChangesARRAY()) = False Then
'                        sReturnErrormsg = "Could not update the source file"
'                        lReturnErrorRow = iarray
'                        '2014/01/17 RAS changed this to go to the error handler from the exit routine
'                        '                        GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'                        'End If
'                    End If
'                    Debug.Print("Now finishing processing row: " & iarray)
'                    If lReturnErrorRow >= iarray Or lReturnErrorRow = 0 Then
'                        lReturnErrorRow = iarray
'                    Else
'                        bMultipleErrors = True
'                    End If
'                End If
'                Debug.Print("Now finishing processing row: " & iarray)
'            Next
'            If errormessageflag = True Then
'                ' change back values to the original values?
'                MsgBox("Not all rows were able to be imported." & vbCrLf & _
'                " Check your log file and reimport rows or contact IT.", vbOKOnly)
'                Application.DoEvents()
'                Call bUpdateStatusMessage(frmThis, "Check your log file, not all rows imported.... ", True, vbRed)
'                sReturnErrormsg = "Check your log file, not all rows imported"
'                bSaveData = False
'                lReturnErrorRow = lReturnErrorRow
'                '2014/01/17 RAS changed this to go to the error handler from the exit routine
'                '                   ' GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'            End If
'            If lReturnErrorRow = 0 Then
'                Call bUpdateStatusMessage(frmThis, "Completed Proposal - Rev Validation ", True, vbGreen)
'                bSaveData = True
'            Else
'                If bMultipleErrors = True Then
'                    Call bUpdateStatusMessage(frmThis, "Multiple Errors in importing starting at row " & lReturnErrorRow, True, vbRed)
'                Else
'                    Call bUpdateStatusMessage(frmThis, "Errors in importing ", True, vbRed)
'                    '                   ' GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If
'            End If
'            'End If
'        Else
'            '2014/03/04 RAS setting the boolean to true if from a proposal form.  The Proposal Rev already exists.
'            bSaveData = True
'        End If
'        '2014/01/09 RAS end of my code
'        '    Application.DoEvents
'        '    bSaveData = True

'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Application.DoEvents()    '2013/04/28 -HN

'        '2013/04/30 -HN- now each table is saved in bSaveRow
'        '    If RSItem.State <> 0 Then RSItem.Close
'        rsItem = Nothing
'        '    If RSItemSpecs.State <> 0 Then RSItemSpecs.Close
'        RSItemSpecs = Nothing
'        '    If rsAssortments.State <> 0 Then rsAssortments.Close
'        rsAssortments = Nothing

'        Exit Function
'ErrorHandler:
'        'Resume Next ''' testing
'        If lReturnErrorRow < 1 Then         '2012/01/04
'            lReturnErrorRow = lRowCounter
'        End If

'        If Err.Number = 9 Then                  '2013/04/28 trouble shooting subscript out of range error
'            If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'                smessage = "In bSaveData  for row :" & lReturnErrorRow & " Error number: " & Err.Number & " ,Error Description: " & Err.Description
'                If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'                End If
'            End If
'            Application.DoEvents()
'            Resume Next                        '2014/01/20 RAS instead of continueing cancel processing.
'        End If
'        If Err.Number = 3219 Then               '2013/04/28 - operation not allowed in this context error
'            If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'                smessage = "In bSaveData  for row :" & lReturnErrorRow & " Error number: " & Err.Number & " ,Error Description: " & Err.Description
'                If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'                End If
'            End If
'            Resume Next                         '2014/01/20 RAS instead of continueing cancel processing.
'        End If
'        If Err.Number = -2147467259 Then
'            If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'                smessage = "In bSaveData  for row :" & lReturnErrorRow & " Error number: " & Err.Number & " ,Error Description: " & Err.Description
'                If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'                End If
'            End If
'            ' close the connection???
'            SSDataConn.Close()

'        End If
'        If Err.Number = -2147217885 Then        '2013/04/28 - cursor conflict error happens from time to time??
'            If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'                smessage = "In bSaveData  for row :" & lReturnErrorRow & " Error number: " & Err.Number & " ,Error Description: " & Err.Description
'                If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'                End If
'            End If
'            Resume Next                         '2014/01/20 RAS instead of continueing cancel processing.
'        Else
'            If Err.Number > 0 Then
'                MsgBox(Err.Number & "-" & Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bSaveData") '2012/01/12
'                '2012/09/06
'            End If
'        End If
'        Resume Next   '' 'testing
'        '2014/01/17 RAS prevalidation of pROPSAL AND REV
'        'Resume ExitRoutine
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bSaveData  for row :" & lReturnErrorRow & " Error number: " & Err.Number & " ,Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        bSaveData = False
'        '    GoTo ValidateProposalRev_Exist'TODO - GoTo Statements are redundant in .NET
'SaveTableError:
'        '2014/01/20 RAS checking what Row has errored and setting lReturnErrorRow  = lRow If lReturnErrorRow = 0
'        If lReturnErrorRow = 0 Then
'            lReturnErrorRow = lRowCounter
'        End If
'        '2013/05/30 -HN- delete all records if all 4 tables were not saved for either/or Item, ItemSpecs, Item_Assortments tables
'        '
'        '        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'        '            smessage = "Records deleted for Row: " & lRowCounter & vbCrLf & "Restart Import Process/Call System Administrator, Error Number:" & Err.Number & "Error Description: " & Err.Description
'        '            If bWritePrintToLogFile(False, objExcel.name & smessage, "ErrorMessageLog") = False Then
'        '            End If
'        '        End If

'        ' 2014/01/07 RAS changed one of the deletes from ITEM to ITEMSPECS
'        '2014/01/07 RAS replace * with & in the item assorment delete statement
'        '2014/01/09 RAS commenting this out recieved a Query time out error and then it deleted out all the records.
'        '    GoTo ValidateProposalRev_Exist'TODO - GoTo Statements are redundant in .NET
'        '    Dim lRecordsAffected As Long
'        '    sProposalNumber = sItemArray(lRowCounter, glProposal_ColPos)
'        '    sRev = sItemArray(lRowCounter, glREV_ColPos)
'        '    sSQL = "DELETE FROM ItemMaterial WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sRev
'        '    sSQL = sSQL & vbCrLf & "DELETE FROM Item_Assortments WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sRev
'        '    sSQL = sSQL & vbCrLf & "DELETE FROM ItemSPECS WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sRev
'        '    sSQL = sSQL & vbCrLf & "DELETE FROM Item WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sRev
'        '
'        '    SSDataConn.Execute sSQL, lRecordsAffected
'        '    Application.DoEvents
'        '    MsgBox "Records deleted for Row: " & lRowCounter & vbCrLf & "Restart Import Process/Call System Administrator", vbCritical, "Error in Import"
'        '
'        ''    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'    End Function
'Private Function bPrepareSaveArray(ByVal lRow As Long, _
'                                dtCOLPos As typSpecialCOLPos, _
'                                sItemSPECSArray() As String , sItemArray() As String , _
'                                sAssortmentArray() As String , _
'                                ByRef sORIGProposalNumber As String , ByRef sORIGRevNumber As String , _
'                                ByRef sDescriptiveMsg As String , _
'                                Optional ByRef lNEWProposalNUM As Long) As Boolean
'        ' Generate new ProposalNumbers, Revs, ItemNumbers, and UPC Codes (if necessary)
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally

'        ' Generate new ProposalNumbers, Revs, ItemNumbers, and UPC Codes (if necessary)
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sFunctioncode As String
'        Dim sBaseProposalNumber As String
'        Dim sBaseRevNumber As String

'        Dim sNewProposalNumber As String
'        Dim sNewRevNumber As String

'        Dim sNewItemNumber As String
'        Dim sItemNumber As String
'        Dim sNewUPCNumber As String

'        Dim sUPCNumber As String
'        Dim sUPCPrefix As String
'        Dim sIPKUPCNumber As String
'        Dim sMPKUPCNumber As String

'        'Dim sPalletUPC                      As String 
'        Dim lcustomernumber As Long
'        Dim sCustomerNumber As String

'        Dim sCustomerItemNumber As String
'        Dim sOtherText As String

'        Dim sCalculatedCustCartonUPC As String
'        Dim sDept As String
'        Dim sClass As String
'        Dim sCustItemNum As String

'        Dim sVendor As String
'        Dim sVendorItemNumber As String

'        Dim sErrMsg As String
'        '2013/07/17 -HN:
'        Const sNewRevMsg = "Change in FactoryFOBCost/FactoryFCACost/Sell Price/RegLinePrice/Prog_Year/(FactoryNumber For Status=ORD), OR -100 Account Requires NEW REVISION!"
'        '2014/01/24 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            smessage = Space(3) & "Starting bPrepareSaveArray for Row " & lRow
'            If bWritePrintToLogFile(False, objEXCELName & smessage, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        bPrepareSaveArray = False

'        ' Get the original Proposal/Rev for log file reporting purposes
'        sORIGProposalNumber = sItemArray(lRow, glProposal_ColPos)
'        sORIGRevNumber = sItemArray(lRow, glREV_ColPos)

'        ' Get the base Proposal/Rev
'        sBaseProposalNumber = sItemArray(lRow, dtCOLPos.lBaseProposalNumber)
'        sBaseRevNumber = sItemArray(lRow, dtCOLPos.lBaseRev)

'        ' Get the CustomerNumber
'        sCustomerNumber = sItemArray(lRow, dtCOLPos.lcustomernumber)
'        If IsNumeric(sCustomerNumber) = True Then
'            lcustomernumber = sCustomerNumber
'        Else
'            lcustomernumber = 0
'        End If

'        'Get the the CustomerItemNumber, OtherText
'        sCustomerItemNumber = sItemArray(lRow, dtCOLPos.lCustomerItemNumber)
'        sOtherText = sItemArray(lRow, dtCOLPos.lOtherText)
'        sVendor = sItemArray(lRow, dtCOLPos.lVendorNumber)
'        sVendorItemNumber = sItemArray(lRow, dtCOLPos.lVendorItemNumber)



'        ' changes to
'        ' AltCost/AltSellPrice/FactoryFCACost/FactoryFOBCost/FCASellPrice/SellPrice/Regular Line Price,
'        ' or, if -100 Account New Revision will be created ..........
'        If bFromProposalForm = False Then
'            sDescriptiveMsg = ""
'            '2014/05/27 RAS Changing this to be <> DEV instead of = Dev. the logic should be  go into if not 999, go into if not  998 , go into if itemstatus is not DEV
'            '2014/06/27 RAS rewriting this if, to make it simpler to read
'            'If (lCustomerNumber <> gs998PD_ACCOUNT And lCustomerNumber <> gs999PD_ACCOUNT) Or UCase(sItemArray(lRow, dtCOLPos.lItemStatus)) <> "DEV" Then  'new rev only for other accounts....
'            If (lcustomernumber <> gs998PD_ACCOUNT And lcustomernumber <> gs999PD_ACCOUNT And msUserGroup <> msMKTGBASIC And msUserGroup <> msPRODUCTMGR) Then
'                '2014/08/15 RAS added mktgbasic an productmgr so it does not check for these groups
'                If UCase(sItemArray(lRow, dtCOLPos.lItemStatus)) <> "DEV" Then
'                    'sometimes these columns are defined as 'text' on spreadsheet columns & need to convert value to decimal for comparision to DB
'                    '12/08/2008 - hn - changed IsNUll to IsBlank below ....
'                    If Not IsBlank(sItemArray(lRow, dtCOLPos.lAltCost)) And _
'                        sItemArray(lRow, dtCOLPos.lAltCost) <> "" And _
'                        sItemArray(lRow, dtCOLPos.lAltCost) <> " " Then
'                        sItemArray(lRow, dtCOLPos.lAltCost) = CDec(sItemArray(lRow, dtCOLPos.lAltCost))
'                    End If
'                    If Not IsBlank(sItemArray(lRow, dtCOLPos.lAltSellPrice)) And _
'                        sItemArray(lRow, dtCOLPos.lAltSellPrice) <> "" And _
'                        sItemArray(lRow, dtCOLPos.lAltSellPrice) <> " " Then
'                        sItemArray(lRow, dtCOLPos.lAltSellPrice) = CDec(sItemArray(lRow, dtCOLPos.lAltSellPrice))
'                    End If
'                    If Not IsBlank(sItemArray(lRow, dtCOLPos.lFactoryFCACost)) And _
'                        sItemArray(lRow, dtCOLPos.lFactoryFCACost) <> "" And _
'                        sItemArray(lRow, dtCOLPos.lFactoryFCACost) <> " " Then
'                        sItemArray(lRow, dtCOLPos.lFactoryFCACost) = CDec(sItemArray(lRow, dtCOLPos.lFactoryFCACost))
'                    End If
'                    If Not IsBlank(sItemArray(lRow, dtCOLPos.lFactoryFOBCost)) And _
'                        sItemArray(lRow, dtCOLPos.lFactoryFOBCost) <> "" And _
'                        sItemArray(lRow, dtCOLPos.lFactoryFOBCost) <> " " Then
'                        sItemArray(lRow, dtCOLPos.lFactoryFOBCost) = CDec(sItemArray(lRow, dtCOLPos.lFactoryFOBCost))
'                    End If
'                    If Not IsBlank(sItemArray(lRow, dtCOLPos.lFCASellPrice)) And _
'                        sItemArray(lRow, dtCOLPos.lFCASellPrice) <> "" And _
'                        sItemArray(lRow, dtCOLPos.lFCASellPrice) <> " " Then
'                        If Len(sItemArray(lRow, dtCOLPos.lFCASellPrice)) > 0 And IsNumeric(sItemArray(lRow, dtCOLPos.lFCASellPrice)) = False Then
'                            sDescriptiveMsg = "There is a problem with " & sItemArray(1, dtCOLPos.lFCASellPrice)
'                        End If
'                        sItemArray(lRow, dtCOLPos.lFCASellPrice) = CDec(sItemArray(lRow, dtCOLPos.lFCASellPrice))
'                    End If
'                    If Not IsBlank(sItemArray(lRow, dtCOLPos.lREGLinePrice)) And _
'                        sItemArray(lRow, dtCOLPos.lREGLinePrice) <> "" And _
'                        sItemArray(lRow, dtCOLPos.lREGLinePrice) <> " " Then
'                        sItemArray(lRow, dtCOLPos.lREGLinePrice) = CDec(sItemArray(lRow, dtCOLPos.lREGLinePrice))
'                    End If
'                    If Not IsBlank(sItemArray(lRow, dtCOLPos.lSellPrice)) And _
'                        sItemArray(lRow, dtCOLPos.lSellPrice) <> "" And _
'                        sItemArray(lRow, dtCOLPos.lSellPrice) <> " " Then
'                        sItemArray(lRow, dtCOLPos.lSellPrice) = CDec(sItemArray(lRow, dtCOLPos.lSellPrice))
'                    End If

'                    'added new fields that can cause a New Rev below...
'                    '2013/07/17 - For ItemStatus = ORD, if FactoryNumber changes then a new Rev will be created and also a new ItemInspection record.
'                    If sItemArray(lRow, dtCOLPos.lAltCost) <> sItemArray_Orig(lRow, dtCOLPos.lAltCost) Or _
'                        sItemArray(lRow, dtCOLPos.lAltSellPrice) <> sItemArray_Orig(lRow, dtCOLPos.lAltSellPrice) Or _
'                        sItemArray(lRow, dtCOLPos.lFactoryFCACost) <> sItemArray_Orig(lRow, dtCOLPos.lFactoryFCACost) Or _
'                        sItemArray(lRow, dtCOLPos.lFactoryFOBCost) <> sItemArray_Orig(lRow, dtCOLPos.lFactoryFOBCost) Or _
'                        sItemArray(lRow, dtCOLPos.lFCASellPrice) <> sItemArray_Orig(lRow, dtCOLPos.lFCASellPrice) Or _
'                        sItemArray(lRow, dtCOLPos.lREGLinePrice) <> sItemArray_Orig(lRow, dtCOLPos.lREGLinePrice) Or _
'                        sItemArray(lRow, dtCOLPos.lSellPrice) <> sItemArray_Orig(lRow, dtCOLPos.lSellPrice) Or _
'                        sItemArray(lRow, dtCOLPos.lProgramYear) <> sItemArray_Orig(lRow, dtCOLPos.lProgramYear) Or _
'                        lcustomernumber = gs100_ACCOUNT Or _
'                        (sItemArray(lRow, dtCOLPos.lFactoryNumber) <> sItemArray_Orig(lRow, dtCOLPos.lFactoryNumber) And sItemArray(lRow, dtCOLPos.lItemStatus) = "ORD") Then

'                        If sItemArray(lRow, glFunctionCode_ColPos) <> gsNEW_REVISION And _
'                           sItemArray(lRow, glFunctionCode_ColPos) <> gsNEW_PROPOSAL Then
'                            sItemArray(lRow, glFunctionCode_ColPos) = gsNEW_REVISION
'                            sDescriptiveMsg = sNewRevMsg
'                        End If
'                    End If
'                End If
'            Else
'                If sItemArray(lRow, dtCOLPos.lProgramYear) <> sItemArray_Orig(lRow, dtCOLPos.lProgramYear) Then
'                    If sItemArray(lRow, glFunctionCode_ColPos) <> gsNEW_REVISION And _
'                       sItemArray(lRow, glFunctionCode_ColPos) <> gsNEW_PROPOSAL Then
'                        sItemArray(lRow, glFunctionCode_ColPos) = gsNEW_REVISION
'                        sDescriptiveMsg = sNewRevMsg
'                    End If
'                End If
'            End If
'        End If

'        ' obtain FunctionCode
'        sFunctioncode = sItemArray(lRow, glFunctionCode_ColPos)

'        If sFunctioncode = gsNEW_PROPOSAL Then

'            If bGetNextProposalNumber(sNewProposalNumber) = False Then
'                MsgBox("Could not retrieve new Proposal Number from ProposalReservations() for Row " & lRow, _
'                        vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'                lNEWProposalNUM = 0
'                '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            Else
'                lNEWProposalNUM = sNewProposalNumber
'                ' New Proposal, Rev = 0
'                sNewRevNumber = "0"
'                sItemArray(lRow, glProposal_ColPos) = sNewProposalNumber
'                sItemArray(lRow, glREV_ColPos) = sNewRevNumber
'                If bPROPOSALFormIndicator = False Then
'                    sItemSPECSArray(lRow, glProposal_ColPos) = sNewProposalNumber
'                    sItemSPECSArray(lRow, glREV_ColPos) = sNewRevNumber
'                End If
'                sAssortmentArray(lRow, glProposal_ColPos) = sNewProposalNumber
'                sAssortmentArray(lRow, glREV_ColPos) = sNewRevNumber

'                ' Store the source ProposalNumber/Rev with the new record for PD items
'                '            If lCustomerNumber = gs999PD_ACCOUNT Or lCustomerNumber = gs998PD_ACCOUNT Then
'                If sORIGProposalNumber = "0" Then
'                    sItemArray(lRow, dtCOLPos.lBaseProposalNumber) = ""
'                Else
'                    sItemArray(lRow, dtCOLPos.lBaseProposalNumber) = sORIGProposalNumber
'                End If
'                If sORIGRevNumber = "0" And sORIGProposalNumber = "0" Then
'                    sItemArray(lRow, dtCOLPos.lBaseRev) = ""
'                Else
'                    sItemArray(lRow, dtCOLPos.lBaseRev) = sORIGRevNumber
'                End If
'                '            Else
'                '                sItemArray(lRow, dtCOLPos.lBaseProposalNumber) = ""
'                '                sItemArray(lRow, dtCOLPos.lBaseRev) = ""
'                '            End If
'            End If

'        ElseIf sFunctioncode = gsNEW_REVISION Then
'            If bGetNextRevNumber(sORIGProposalNumber, sNewRevNumber) = False Then
'                MsgBox("Could not retrieve next Rev for Proposal " & sORIGProposalNumber, _
'                        vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'                '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            Else
'                ' New revision for proposal
'                sItemArray(lRow, glREV_ColPos) = sNewRevNumber
'                If bPROPOSALFormIndicator = False Then sItemSPECSArray(lRow, glREV_ColPos) = sNewRevNumber
'                sAssortmentArray(lRow, glREV_ColPos) = sNewRevNumber

'                ' The base ProposalNumber/Rev doesn't change across PD revisions
'                '   because the purpose for the base proposal is rolling PD changes back into the base line
'                '            If lCustomerNumber = gs999PD_ACCOUNT Or lCustomerNumber = gs998PD_ACCOUNT Then
'                If sBaseProposalNumber = "0" Then
'                    sItemArray(lRow, dtCOLPos.lBaseProposalNumber) = ""
'                Else
'                    sItemArray(lRow, dtCOLPos.lBaseProposalNumber) = sBaseProposalNumber
'                End If
'                If sBaseRevNumber = "0" And sBaseProposalNumber = "0" Then
'                    sItemArray(lRow, dtCOLPos.lBaseRev) = ""
'                Else
'                    sItemArray(lRow, dtCOLPos.lBaseRev) = sBaseRevNumber
'                End If
'                '            Else
'                '                sItemArray(lRow, dtCOLPos.lBaseProposalNumber) = ""
'                '                sItemArray(lRow, dtCOLPos.lBaseRev) = ""
'                '            End If
'            End If

'        End If

'        ' ItemNumber & UPC
'        sItemNumber = sItemArray(lRow, dtCOLPos.lItemNumber)
'        gsItemNumber = sItemNumber 'saved for log file Orig values on import spreadsheet
'        '    sItemArray(lRow, dtCOLPos.lUPCNumber) = UPCReservations(0).UPCNumber
'        sUPCNumber = sItemArray(lRow, dtCOLPos.lUPCNumber)
'        sUPCPrefix = Microsoft.VisualBasic.Left(sUPCNumber, 6)
'        sIPKUPCNumber = sItemArray(lRow, dtCOLPos.lIPKUPC)
'        sMPKUPCNumber = sItemArray(lRow, dtCOLPos.lMPKUPC)
'        '    sPalletUPC = sItemArray(lRow, dtCOLPos.lPalletUPC)
'        gsUPCNumber = sUPCNumber

'        If bProposalGetNewItemNumber = True Or _
'            (sItemNumber = gsNEW_ITEM_NBR And bFromProposalForm = False) Then
'            '    If (sItemNumber = gsNEW_ITEM_NBR And bNewCustomerFromProposal = False) Or _
'            '        (sItemNumber <> "" And bNewItemFromProposal = True And bNewCustomerFromProposal = False) Or _
'            '            bProposalGetNewItemNumber = True Then
'            If bGetNextITEMNumber(sNewItemNumber, lcustomernumber, sVendor, sVendorItemNumber, sErrMsg) = False Then
'                If sErrMsg <> "" Then
'                    MsgBox("For Row: " & lRow & " " & sORIGProposalNumber & " " & sORIGRevNumber _
'                    & vbCrLf & sErrMsg, vbCritical + vbMsgBoxSetForeground, "New ItemNumber Prohibited!")
'                    sDescriptiveMsg = sErrMsg
'                Else
'                    MsgBox("Could not retrieve new ItemNumber from UPCReservations Array, for Row " & lRow, _
'                        vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'                End If
'                '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET

'            Else
'                ' New ItemNumber
'                sItemNumber = sNewItemNumber & "-" & sItemArray(lRow, dtCOLPos.lcustomernumber)
'                sItemArray(lRow, dtCOLPos.lItemNumber) = sItemNumber

'                '            If Not bCreateBarcode(gsNewItemUPCPrefix, Microsoft.VisualBasic.Right(sNewItemNumber, 5), sNewUPCNumber) Then
'                '                MsgBox "Could not generate barcode for new ItemNumber " & sItemNumber, _
'                '                        vbExclamation, "modSpreadSheet-bPrepareSaveArray"
'                ''                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                '            Else
'                ' New UPC
'                '                sITEMArray(lRow, dtCOLPos.lUPCNumber) = sNewUPCNumber
'                sItemArray(lRow, dtCOLPos.lUPCNumber) = UPCReservations(lNewItemCounter).UPCNumber
'                '                sUPCPrefix = Microsoft.VisualBasic.Left(sNewUPCNumber, 6)
'                sUPCPrefix = Microsoft.VisualBasic.Left(UPCReservations(lNewItemCounter).UPCNumber, 6)
'                '            End If

'                If bCreateIpkMpkBarcode("30" & sUPCPrefix, Microsoft.VisualBasic.Right(sNewItemNumber, 5), sNewUPCNumber) = False Then
'                    MsgBox("Could not generate IPK barcode for new ItemNumber " & sItemNumber, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'                    '                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                Else
'                    ' New IPK UPC
'                    sItemArray(lRow, dtCOLPos.lIPKUPC) = sNewUPCNumber
'                End If

'                If bCreateIpkMpkBarcode("50" & sUPCPrefix, Microsoft.VisualBasic.Right(sNewItemNumber, 5), sNewUPCNumber) = False Then
'                    MsgBox("Could not generate MPK barcode for new ItemNumber " & sItemNumber, _
'                            vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'                    '                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                Else
'                    ' New MPK UPC
'                    sItemArray(lRow, dtCOLPos.lMPKUPC) = sNewUPCNumber
'                End If
'            End If
'        Else
'            ' Existing ItemNumber(no A)
'            If Len(sItemNumber) >= 5 Then
'                If Len(sUPCNumber) = 0 Then
'                    If bGetUPCPrefix(Left(sItemNumber, 6), sUPCPrefix) = True Then
'                        gsSelectedUPCPrefix = sUPCPrefix
'                    End If
'                    If bCreateBarcode(gsSelectedUPCPrefix, Mid(sItemNumber, 2, 5), sNewUPCNumber) = False Then
'                        MsgBox("Could not generate barcode for ItemNumber " & sItemNumber, _
'                            vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'                        '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        ' Regenerate UPC
'                        sItemArray(lRow, dtCOLPos.lUPCNumber) = sNewUPCNumber
'                        '                    sUPCPrefix = Microsoft.VisualBasic.Left(sNewUPCNumber, 6)
'                    End If
'                End If
'                '            If Len(sIPKUPCNumber) = 0 Then
'                If bCreateIpkMpkBarcode("30" & sUPCPrefix, Mid(sItemNumber, 2, 5), sNewUPCNumber) = False Then
'                    MsgBox("Could not generate IPK barcode for ItemNumber " & sItemNumber, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'                    '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                Else
'                    ' Regenerate IPK UPC
'                    sItemArray(lRow, dtCOLPos.lIPKUPC) = sNewUPCNumber
'                End If
'                '            End If
'                '            If Len(sMPKUPCNumber) = 0 Then
'                If bCreateIpkMpkBarcode("50" & sUPCPrefix, Mid(sItemNumber, 2, 5), sNewUPCNumber) = False Then
'                    MsgBox("Could not generate MPK barcode for ItemNumber " & sItemNumber, _
'                        vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'                    '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                Else
'                    ' Regenerate IPK UPC
'                    sItemArray(lRow, dtCOLPos.lMPKUPC) = sNewUPCNumber
'                End If
'                '            End If

'            Else
'                ' should be caught in validation process
'                '            If Len(sUPCNumber) = 0 Then
'                '                sITEMArray(lRow, udtColPos.lUPCNumber) = ""
'                '            End If
'            End If
'        End If
'        'calculate CustomerCartonUPC Number for Target Customers
'        'put validation in bValidateField

'        Select Case lcustomernumber
'            Case 101, 102, 235, 206, 103, 888, 248
'                If Not IsBlank(sCustomerItemNumber) Then
'                    sDept = Microsoft.VisualBasic.Left(sCustomerItemNumber, 3)
'                    sClass = Mid(sCustomerItemNumber, 5, 2)
'                    sCustItemNum = Microsoft.VisualBasic.Right(sCustomerItemNumber, 4)

'                    If Not bCreateCustomerCartonUPC(7049, sDept, sClass, sCustItemNum, sCalculatedCustCartonUPC) Then
'                        sDescriptiveMsg = "Cannot generate CustomerCartonUPC; check CustomerItemNumber"
'                        sItemArray(lRow, dtCOLPos.lCustomerCartonUPCNumber) = ""
'                        '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        sItemArray(lRow, dtCOLPos.lCustomerCartonUPCNumber) = sCalculatedCustCartonUPC
'                    End If
'                Else
'                    sItemArray(lRow, dtCOLPos.lCustomerCartonUPCNumber) = "" '2013/05/28 -HN-
'                End If

'            Case 104
'                If Not IsBlank(sOtherText) Then
'                    sDept = Microsoft.VisualBasic.Left(sOtherText, 3)
'                    sClass = Mid(sOtherText, 5, 2)
'                    sCustItemNum = Microsoft.VisualBasic.Right(sOtherText, 4)
'                    If Not bCreateCustomerCartonUPC(7049, sDept, sClass, sCustItemNum, sCalculatedCustCartonUPC) Then
'                        sDescriptiveMsg = "Cannot generate CustomerCartonUPC; check OtherText"
'                        sItemArray(lRow, dtCOLPos.lCustomerCartonUPCNumber) = ""
'                        '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        sItemArray(lRow, dtCOLPos.lCustomerCartonUPCNumber) = sCalculatedCustCartonUPC
'                    End If
'                Else
'                    sItemArray(lRow, dtCOLPos.lCustomerCartonUPCNumber) = "" '2013/05/28 -HN-
'                End If
'            Case 235, 252, 253, 254, 255, 257, 889, 998
'                '2014/05/30 RAS check for all other targets
'            Case Else
'                '2013/05/28 -HN- CustomerCartonUPCNumber must be blank for non Target Customers for FunctionCode = A
'                If sFunctioncode = "A" And bPROPOSALFormIndicator = False Then
'                    sItemArray(lRow, dtCOLPos.lCustomerCartonUPCNumber) = ""
'                    '2014/03/04 RAS not rolling out blanking out Class at this time.
'                    '2014/02/27 RAS blanking out class for non target customers.
'                    '2014/05/06 RAS class validation is back in. Non Target customers are blanked out.
'                    sItemArray(lRow, dtCOLPos.lClass) = ""
'                End If

'        End Select

'        If bOverwriteDuplicateFields(lRow, dtCOLPos, sItemSPECSArray(), sItemArray()) = False Then
'            MsgBox("Could not overwrite the duplicate fields for Row " & lRow, _
'                    vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If

'        If bCalcCubes(lRow, dtCOLPos, sItemSPECSArray(), sItemArray()) = False Then
'            MsgBox("Could not Calculate MasterPackCube and/or InnerPackCube for Row " & lRow, _
'                    vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If
'        Application.DoEvents()
'        bPrepareSaveArray = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:
'        ' Resume Next  ' testing only
'        If sDescriptiveMsg <> "" Then
'            MsgBox(sDescriptiveMsg & " " & Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'        Else
'            MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bPrepareSaveArray")
'        End If
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bSaveData " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    '2011/12/29 - made function public As Object 
'    'so that Import process can write to it As Object 

'    '2011/12/29 - made function public, so that Import process can write to it
'    Public Function bWriteToExcelLog(ByVal lRow As Long, objExcel As Excel.Application, _
'                            ByVal sfilename As String, _
'                            dtCOLPos As typSpecialCOLPos, _
'                            sItemSPECSArray() As String, sItemArray() As String, _
'                            ByVal sORIGProposalNumber As String, ByVal sORIGRevNumber As String, _
'                            sDescriptiveMsg As String, Optional bFROMItemStatus As Boolean = False) As Boolean

'        Const lMaxLogFileCol = 14

'Dim sLogArr(1 To lMaxLogFileCol - 1) As String

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally

'        bWriteToExcelLog = False
'        If sItemArray(lRow, glFunctionCode_ColPos) <> "" Then     'skip reporting blank rows on log file
'            If lLogFileRow = 0 Then ' skip heading row
'                lLogFileRow = 2
'            Else
'                If bFROMItemStatus = False Then 'keep msg on the same logfile row
'                    lLogFileRow = lLogFileRow + 1
'                End If
'            End If

'            '        If bFROMItemStatus = True Then GoTo SKIPCELLS'TODO - GoTo Statements are redundant in .NET
'            ' Import Log File reporting

'            sLogArr(1) = CStr(lRow)
'            sLogArr(2) = sItemArray(lRow, glFunctionCode_ColPos)

'            sLogArr(3) = sItemArray(lRow, glProposal_ColPos)
'            sLogArr(4) = sItemArray(lRow, glREV_ColPos)
'            sLogArr(5) = sItemArray(lRow, dtCOLPos.lProgramYear)
'            sLogArr(6) = sItemArray(lRow, dtCOLPos.lItemNumber)
'            sLogArr(7) = "'" & sItemArray(lRow, dtCOLPos.lUPCNumber)
'            ' The single-quote prevents Excel from auto-formatting & dropping leading zeros
'            sLogArr(8) = sItemArray_Orig(lRow, glProposal_ColPos)
'            sLogArr(9) = sItemArray_Orig(lRow, glREV_ColPos)
'            sLogArr(10) = sItemArray_Orig(lRow, dtCOLPos.lItemNumber)
'            '        objEXCEL.ActiveWorkbook.Worksheets(1).Cells(lRow, glLog_ORIG_ITEM_ColPos).Value = gsItemNumber
'            sLogArr(11) = gsUPCNumber
'            sLogArr(12) = sItemArray(lRow, dtCOLPos.lLongDesc)
'            sLogArr(13) = sItemArray(lRow, dtCOLPos.lVendorItemNumber)

'            With objExcel.ActiveWorkbook.Worksheets(1)
'                .Range(.Cells(lLogFileRow, 1), .Cells(lLogFileRow, lMaxLogFileCol - 1)) = sLogArr
'            End With

'SKIPCELLS:
'            ' do the last row differently because need to concatenate and set color of the cell
'            objExcel.ActiveWorkbook.Worksheets(1).Cells(lLogFileRow, lMaxLogFileCol).Font.ColorIndex = 3
'            If bFROMItemStatus = True Then
'                objExcel.ActiveWorkbook.Worksheets(1).Cells(lLogFileRow, lMaxLogFileCol).Value = _
'                objExcel.ActiveWorkbook.Worksheets(1).Cells(lLogFileRow, lMaxLogFileCol).Value & sDescriptiveMsg
'            Else
'                objExcel.ActiveWorkbook.Worksheets(1).Cells(lLogFileRow, lMaxLogFileCol).Value = sDescriptiveMsg
'            End If

'            Application.DoEvents()
'            '    objEXCEL.Application.Workbooks(1).Save  ' a quirk with Excel 9, dsnt like to save as??
'            '    objEXCEL.Application.Workbooks(1).SaveAs sFileName, xlNormal   '01/16/2008 MS Office 2007 dsnt like?
'            '    objEXCEL.Application.Workbooks(1).Save                         '01/16/2008 try it this way again
'            Application.DoEvents()
'            If Version >= 12.0# Then
'                objExcel.ActiveWorkbook.Worksheets(1).SaveAs(sfilename, , , , , 0)
'                '        objEXCEL.Application.Workbooks(1).SaveAs sFilename, , , , , 0 '2010/05/11 to prevent .xlk backup file being created
'                Application.DoEvents()
'            Else
'                objExcel.Application.Workbooks(1).SaveAs(sfilename, xlNormal)
'            End If
'            Application.DoEvents()
'        End If                      '2012/01/04
'        bWriteToExcelLog = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:


'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bWriteToExcelLog-(lrow=" & lRow & "), Updating ItemStatus=" & bFROMItemStatus)           '2011/12/28 added lrow
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bWriteToExcelLog -(lrow=" & lRow & ") " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Private Function bUpdateExcelSourceFile(RowChangesARRAY() As String, ByVal lRow As Long, lRowsOnSheet As Long, ByVal sfilename As String, _
'                            objExcel As Excel.Application, _
'                            dtSpreadsheetCOLArray() As typColumn, dtSpreadsheetCOLPos As typSpecialCOLPos, _
'                            ByVal lColsOnSheet As Long, _
'                            ByVal lNbrItemSpecsCOLS As Long, ByVal lNbrItemCOLS As Long, _
'                            ByVal lNbrAssortmentCOLS As Long, _
'                            sItemSPECSArray() As String, sItemArray() As String, _
'                            dtSaveArrayCOLPos As typSpecialCOLPos, ByVal frmThis As Form) As Boolean

'        'Private Function bUpdateExcelSourceFile(ByVal RowChangesARRAY( As Object) As [String]
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sFunctioncode As String
'        Dim sProposalOnSheet As String
'        Dim sRevOnSheet As String

'        Dim sItemNumberOnSheet As String
'        Dim sUPCNumberOnSheet As String

'        Dim sTechnologies As String
'        Dim sX_Technologies As String
'        Dim sInvalidTechnologies As String

'        Dim lTechnologiesCol As Long
'        Dim lCertifiedPrinterIDCol As Long
'        Dim sCertifiedPrinterID As String
'        Dim sX_CertifiedPrinterName As String
'        Dim sInvalidCertifiedPrinterID As String     '2011/10/26

'        Dim lSpreadsheetCOLPos As Long
'        Dim sProgramYear As String
'        Dim lProgramYear As Long

'        Dim lColCounter As Long
'        Dim sLookupValue As String
'        Dim sSpreadsheetValue As String

'        Dim sDefaultValue As String
'        Dim sSpreadSheetColHeading As String
'        Dim sSpreadSheetColHeadingAlias As String

'        '-----------------for spreadsheet shading
'        Dim lUdtColumn As Long
'        Dim lArrayCol As Long
'        Dim sTableName As String
'        Dim sOLDCellValue As String
'        Dim i As Integer

'        bUpdateExcelSourceFile = False
'        sFunctioncode = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, glFunctionCode_ColPos).Value
'        '------------------ if 'END' is used to denote end of spreadsheet in FC column ---------------------------------------
'        If lRow = lRowsOnSheet + 1 Then
'            Dim ws As Excel.Worksheet ' Set a reference to this excel worksheet.
'            ws = objExcel.Application.Workbooks(1).Worksheets(1)
'            objExcel.Visible = False
'            If ws.Cells(lRow, glFunctionCode_ColPos).Value = gsENDofSpreadsheet Then
'                ws.Cells(lRow, glFunctionCode_ColPos).Value = ""
'                bUpdateExcelSourceFile = True
'                If bRefreshImportPhotos = True Then
'                    Screen.MousePointer = 11 ' Hourglass
'                    Call bUpdateStatusMessage(frmThis, "Deleting all photos  ...", True, vbGreen)
'                    '07/01/2008 - hn - commented below - don't want to lose all photos after blank row - oops!
'                    '                Do Until ws.Pictures.Count = 0  'this clears out ALL photos
'                    '                    ws.Pictures(1).Delete
'                    '                Loop
'                    '07/01/2008 - hn - copied from frmExcelRefresh
'                    For i = ws.Pictures.Count To 1 Step -1
'                        If ws.Pictures(i).TopLeftCell.Row <= lRowsOnSheet + 1 Then
'                            If ws.Pictures(i).TopLeftCell.Column < lColsOnSheet Then
'                                'This picture is to left of the first blank column; so can delete the photo, any other photos after blank column remain
'                                ws.Pictures(i).Delete()
'                            End If
'                        End If
'                    Next

'                    Call bUpdateStatusMessage(frmThis, "Refreshing all photos  ...", True, vbGreen)
'                    If bInsertPhotoInEachRow(objExcel, lRowsOnSheet, sOtherPhotoPath) = True Then
'                        If Version >= 12.0# Then
'                            objExcel.Application.Workbooks(1).SaveAs(sfilename, , , , , 0) 'to prevent .xlk backup file being created
'                            Application.DoEvents()
'                        Else
'                            objExcel.Application.Workbooks(1).SaveAs(sfilename, xlNormal)
'                        End If
'                    End If
'                End If
'                Screen.MousePointer = 0
'                '             Save the file
'                If Version >= 12.0# Then
'                    objExcel.Application.Workbooks(1).SaveAs(sfilename, , , , , 0) 'to prevent .xlk backup file being created
'                    Application.DoEvents()
'                Else
'                    objExcel.Application.Workbooks(1).SaveAs(sfilename, xlNormal)
'                End If

'                bUpdateExcelSourceFile = True
'            Else
'                If lRow > lRowsOnSheet Then
'                Else
'                    '                If RowChangesARRAY(lRow) = 0 And sFunctioncode = "C" Then GoTo RemoveFunctionCode'TODO - GoTo Statements are redundant in .NET
'                End If

'                If bRefreshImportPhotos = True Then
'                    Screen.MousePointer = 11 ' Hourglass
'                    Call bUpdateStatusMessage(frmThis, "Deleting all photos  ... (may take several minutes for large spreadsheets)", True, vbGreen)  '2010/11/29
'                    '07/01/2008 - hn - commented below - don't want to lose all photos after blank row - oops!
'                    '                Do Until ws.Pictures.Count = 0  'this clears out ALL photos
'                    '                    ws.Pictures(1).Delete
'                    '                Loop
'                    '07/01/2008 - hn - copied from frmExcelRefresh
'                    For i = ws.Pictures.Count To 1 Step -1
'                        If ws.Pictures(i).TopLeftCell.Row <= lRowsOnSheet + 1 Then
'                            If ws.Pictures(i).TopLeftCell.Column < lColsOnSheet Then
'                                'This picture is to left of the first blank column; so can delete the photo, any other photos after blank column remain
'                                ws.Pictures(i).Delete()
'                            End If
'                        End If
'                    Next
'                    Call bUpdateStatusMessage(frmThis, "Refreshing all photos  ... (may take several minutes for large spreadsheets)", True, vbGreen)  '2010/11/29
'                    Screen.MousePointer = 11 ' Hourglass
'                    If bInsertPhotoInEachRow(objExcel, lRowsOnSheet, sOtherPhotoPath) = True Then
'                        If Version >= 12.0# Then
'                            objExcel.Application.Workbooks(1).SaveAs(sfilename, , , , , 0)           'to prevent .xlk backup file being created
'                            Application.DoEvents()                                                                                              '2010/11/19
'                        Else
'                            objExcel.Application.Workbooks(1).SaveAs(sfilename, xlNormal)
'                        End If
'                    End If
'                End If
'                bUpdateExcelSourceFile = True
'                Screen.MousePointer = 0
'                Exit Function
'            End If
'        End If
'RemoveFunctionCode:
'        Screen.MousePointer = 11 ' Hourglass
'        '2011/12/13 - commented below trying to speed up Import by not displaying all messages to user
'        '    Call bUpdateStatusMessage(frmThis, "Updating spreadsheet  ...", True, vbGreen)
'        objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, glFunctionCode_ColPos).Value = ""


'        '  comment row below, if X_fields should be updated,if no other changes on row
'        '    If RowChangesARRAY(lRow) = 0 And sFunctioncode = "C" Then GoTo SaveExcelFile'TODO - GoTo Statements are redundant in .NET
'        ' Update ProposalNumber, Rev, ItemNumber, UPCNumber
'        sProposalOnSheet = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, glProposal_ColPos))
'        Call bUpdateSpreadsheetCell(sDefaultValue, sItemArray(lRow, glProposal_ColPos), sProposalOnSheet, _
'                                        objExcel, lRow, glProposal_ColPos, False)

'        sRevOnSheet = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, glREV_ColPos))
'        Call bUpdateSpreadsheetCell(sDefaultValue, sItemArray(lRow, glREV_ColPos), sRevOnSheet, _
'                                        objExcel, lRow, glREV_ColPos, False)

'        lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lItemNumber
'        If lSpreadsheetCOLPos > 0 Then
'            Call bUpdateSpreadsheetCell(sDefaultValue, sItemArray(lRow, dtSaveArrayCOLPos.lItemNumber), _
'                                        sItemArray_Orig(lRow, dtSaveArrayCOLPos.lItemNumber), _
'                                            objExcel, lRow, lSpreadsheetCOLPos, False)
'        End If

'        lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lClass
'        If lSpreadsheetCOLPos > 0 Then
'            Call bUpdateSpreadsheetCell(sDefaultValue, sItemArray(lRow, dtSaveArrayCOLPos.lClass), _
'                    sItemArray_Orig(lRow, dtSaveArrayCOLPos.lClass), objExcel, lRow, lSpreadsheetCOLPos, False)
'        End If

'        '----------- set SellPrice = blank on Spreadsheet, for FC = A and CustomerNumber = 998/999
'        '2014/05/07 RAS and now for ItemStatus "DEV"
'        If sFunctioncode = gsNEW_PROPOSAL And _
'            (sItemArray(lRow, dtSaveArrayCOLPos.lcustomernumber) = gs999PD_ACCOUNT Or UCase(sItemArray(lRow, dtSaveArrayCOLPos.lItemStatus)) = "DEV") Then
'            ' Or
'            'sItemSPECSArray(lRow, dtSaveArrayColPos.lVendorCustomerNumber) = gs998_ACCOUNT) Then
'            lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lSellPrice
'            If lSpreadsheetCOLPos > 0 Then
'                Call bUpdateSpreadsheetCell(sDefaultValue, _
'                        sItemArray(lRow, dtSaveArrayCOLPos.lSellPrice), _
'                        sItemArray_Orig(lRow, dtSaveArrayCOLPos.lSellPrice), _
'                        objExcel, lRow, lSpreadsheetCOLPos, False)  'new 10/15/2007
'            End If
'        End If

'        'set CustomerHSNumber and CustomerDutyRate to blank if Customer <> Target -------------
'        If sFunctioncode = gsNEW_PROPOSAL Then
'            Select Case sItemArray(lRow, dtSaveArrayCOLPos.lcustomernumber)
'                '2014/05/28 Adding all the other targets.
'                Case 101, 102, 103, 104, 206, 235, 248, 252, 253, 254, 255, 257, 888, 889, 998
'                    'Case 101, 102, 103, 104, 206, 235, 248 '2011/11/08 - added 248
'                Case Else
'                    'blank out CustomerHSNumber and CustomerDutyRate on Import spreadsheet
'                    lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lCustomerDutyRate
'                    If lSpreadsheetCOLPos > 0 Then
'                        Call bUpdateSpreadsheetCell(sDefaultValue, _
'                            "", "Blank", _
'                            objExcel, lRow, lSpreadsheetCOLPos, False)
'                    End If
'                    lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lCustomerHSNumber
'                    If lSpreadsheetCOLPos > 0 Then
'                        Call bUpdateSpreadsheetCell(sDefaultValue, _
'                            "", "Blank", _
'                            objExcel, lRow, lSpreadsheetCOLPos, False)

'                    End If
'                    '2014/05/06 RAS Class check is back in.
'                    '2014/03/04 RAS not blanking out class at this time.
'                    '2014/02/27 RAS Blanking out Class for non target customers.
'                    lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lClass
'                    If lSpreadsheetCOLPos > 0 Then
'                        Call bUpdateSpreadsheetCell(sDefaultValue, _
'                            "", "Blank", _
'                            objExcel, lRow, lSpreadsheetCOLPos, False)

'                    End If

'            End Select
'        End If
'        '--------------------------------------------------------------------------------------------------------

'        'HN 02/13/2007 RoyaltyID(Licensor) = ProgramNumber if =113, 501-555
'        'hn 04/11/2007 after reclassification rollout, only true for ProgramYear < 2008...
'        lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lProgramYear
'        sProgramYear = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lSpreadsheetCOLPos))
'        If sProgramYear <> "" Then
'            lProgramYear = CInt(sProgramYear)
'        Else
'            lProgramYear = 0
'        End If
'        If lProgramYear < 2008 Then
'            If sItemArray(lRow, dtSaveArrayCOLPos.lProgramNumber) = 113 Or _
'                (sItemArray(lRow, dtSaveArrayCOLPos.lProgramNumber) > 500 And _
'                sItemArray(lRow, dtSaveArrayCOLPos.lProgramNumber) < 556) Then
'                '        If sITEMArray(lRow, dtSaveArrayColPos.lLicensor) = _
'                '                sITEMArray(lRow, dtSaveArrayColPos.lProgramNumber) Then
'                lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lLicensor
'                If lSpreadsheetCOLPos > 0 Then
'                    sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lSpreadsheetCOLPos))
'                    If sSpreadsheetValue <> sItemArray(lRow, dtSaveArrayCOLPos.lProgramNumber) Then
'                        Call bUpdateSpreadsheetCell(sDefaultValue, _
'                            sItemArray(lRow, dtSaveArrayCOLPos.lLicensor), _
'                            sSpreadsheetValue, _
'                            objExcel, lRow, lSpreadsheetCOLPos, False)
'                    End If
'                End If
'                '        End If
'            Else
'                'if Program number not = 113, 501-555 then Licensor on spreadsheet cant be = to that
'                lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lLicensor
'                If lSpreadsheetCOLPos > 0 Then
'                    sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lSpreadsheetCOLPos))
'                    If sSpreadsheetValue = "113" Or (sSpreadsheetValue > "500" And sSpreadsheetValue < "556") Then
'                        Call bUpdateSpreadsheetCell(sDefaultValue, _
'                                "", sItemArray(lRow, dtSaveArrayCOLPos.lProgramNumber), _
'                                objExcel, lRow, lSpreadsheetCOLPos, False)  'new 10/15/2007
'                    End If
'                End If
'            End If
'            '    only for ProgramYear < 2008
'        End If

'        '2/12/2007 - hn - TrademarkCopyRight = True if RoyaltyID(Licensor) = 113, 501-555
'        'HN 2/12/2007: TrademarkCopyRight = True if RoyaltyID(Licensor) = 113, 501-555
'        If sItemArray(lRow, dtSaveArrayCOLPos.lLicensor) = "113" Or _
'            (sItemArray(lRow, dtSaveArrayCOLPos.lLicensor) > "500" And _
'            sItemArray(lRow, dtSaveArrayCOLPos.lLicensor) < "556") Then

'            lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lTradeMarkCopyRight
'            If lSpreadsheetCOLPos > 0 Then
'                sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lSpreadsheetCOLPos))
'                If sSpreadsheetValue <> "TRUE" Then
'                    Call bUpdateSpreadsheetCell(sDefaultValue, _
'                        sItemArray(lRow, dtSaveArrayCOLPos.lTradeMarkCopyRight), _
'                        sSpreadsheetValue, objExcel, lRow, lSpreadsheetCOLPos, False)
'                End If
'            End If

'        End If


'        ' --- If Lighted column value = True, Y, 1, replace with Yes, else if = False, 0, N replace with NO
'        lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lLighted
'        If lSpreadsheetCOLPos > 0 Then
'            sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lSpreadsheetCOLPos)) '12/11/2007 to prevent shading on Spreadsheet
'            Select Case sItemArray(lRow, dtSaveArrayCOLPos.lLighted)
'                Case "True", "Y", "1"
'                    Call bUpdateSpreadsheetCell(sDefaultValue, "YES", _
'                            sSpreadsheetValue, _
'                            objExcel, lRow, lSpreadsheetCOLPos, False)
'                Case "False", "N", "0"
'                    Call bUpdateSpreadsheetCell(sDefaultValue, "NO", _
'                            sSpreadsheetValue, _
'                            objExcel, lRow, lSpreadsheetCOLPos, False)
'            End Select
'        End If

'        'update spreadsheet col for ProductBatteriesIncluded
'        lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lProductBatteriesIncluded
'        If lSpreadsheetCOLPos > 0 Then
'            sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lSpreadsheetCOLPos))
'            Select Case sItemSPECSArray(lRow, dtSaveArrayCOLPos.lProductBatteriesIncluded)
'                Case "True", "Y", "1"
'                    Call bUpdateSpreadsheetCell(sDefaultValue, "YES", _
'                            sSpreadsheetValue, _
'                            objExcel, lRow, lSpreadsheetCOLPos, False)
'                Case "False", "N", "0"
'                    Call bUpdateSpreadsheetCell(sDefaultValue, "NO", _
'                            sSpreadsheetValue, _
'                            objExcel, lRow, lSpreadsheetCOLPos, False)
'            End Select
'        End If

'        ' Update spreadsheet TreeLightConstruction value set False = blank, leave others as they were
'        lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lTreeLightConstruction
'        If lSpreadsheetCOLPos > 0 Then
'            sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lSpreadsheetCOLPos))
'            If sSpreadsheetValue = "FALSE" Then
'                Call bUpdateSpreadsheetCell(sDefaultValue, "", sSpreadsheetValue, objExcel, lRow, lSpreadsheetCOLPos, False)
'            End If
'        End If

'        For lColCounter = 4 To lColsOnSheet
'            sDefaultValue = ""
'            Select Case lColCounter
'                Case dtSpreadsheetCOLPos.lLowLeadWholeProduct
'                    '2014/10/03 RAS these three will always be TRUE
'                    ' lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lLowLeadWholeProduct
'                    ' If lSpreadsheetCOLPos > 0 Then
'                    sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, dtSpreadsheetCOLPos.lLowLeadWholeProduct))
'                    If sSpreadsheetValue <> "TRUE" Then
'                        Call bUpdateSpreadsheetCell("TRUE", "TRUE", sSpreadsheetValue, objExcel, lRow, dtSpreadsheetCOLPos.lLowLeadWholeProduct, False)
'                    End If
'                    'End If
'                Case dtSpreadsheetCOLPos.lFlammability
'                    '2014/10/03 RAS these three will always be TRUE
'                    ' lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lFlammability
'                    'If lSpreadsheetCOLPos > 0 Then
'                    sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, dtSpreadsheetCOLPos.lFlammability))
'                    If sSpreadsheetValue <> "TRUE" Then
'                        Call bUpdateSpreadsheetCell("TRUE", "TRUE", sSpreadsheetValue, objExcel, lRow, dtSpreadsheetCOLPos.lFlammability, False)
'                    End If
'                    ' End If
'                Case dtSpreadsheetCOLPos.lSurfaceLeadPaintRequirement
'                    '2014/10/03 RAS these three will always be TRUE
'                    ' lSpreadsheetCOLPos = dtSpreadsheetCOLPos.lSurfaceLeadPaintRequirement
'                    'If lSpreadsheetCOLPos > 0 Then
'                    sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, dtSpreadsheetCOLPos.lSurfaceLeadPaintRequirement))
'                    If sSpreadsheetValue <> "TRUE" Then
'                        Call bUpdateSpreadsheetCell("TRUE", "TRUE", sSpreadsheetValue, objExcel, lRow, dtSpreadsheetCOLPos.lSurfaceLeadPaintRequirement, False)
'                    End If
'                    ' End If

'                Case dtSpreadsheetCOLPos.lItemNumber, dtSpreadsheetCOLPos.lUPCNumber, dtSpreadsheetCOLPos.lIPKUPC, dtSpreadsheetCOLPos.lMPKUPC
'                    'these cols cannot be changed
'                    Application.DoEvents()

'                    'update X_certifiedprintername  ----
'                Case dtSpreadsheetCOLPos.lX_CertifiedPrinterName
'                    'update X_CertifiedPrinterName column if CertifiedPrinterID col exists
'                    lCertifiedPrinterIDCol = dtSpreadsheetCOLPos.lCertifiedPrinterID
'                    sX_CertifiedPrinterName = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColCounter))
'                    If lCertifiedPrinterIDCol > 3 Then
'                        sCertifiedPrinterID = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lCertifiedPrinterIDCol))
'                        If sCertifiedPrinterID = "" Then
'                            If sX_CertifiedPrinterName <> "" Then
'                                Call bUpdateSpreadsheetCell(sDefaultValue, "", sX_CertifiedPrinterName, objExcel, lRow, lColCounter, False)
'                            End If
'                        Else
'                            'find X_CertifiedPrinterName value
'                            sX_CertifiedPrinterName = ""
'                            Call bValidCertifiedPrinterID(lRow, sCertifiedPrinterID, sX_CertifiedPrinterName, sInvalidCertifiedPrinterID)

'                            sX_CertifiedPrinterName = sX_CertifiedPrinterName & sInvalidCertifiedPrinterID
'                            Call bUpdateSpreadsheetCell(sDefaultValue, sX_CertifiedPrinterName, "", objExcel, lRow, lColCounter, False)
'                        End If
'                    Else
'                        If sX_CertifiedPrinterName <> "" Then
'                            Call bUpdateSpreadsheetCell(sDefaultValue, "", sX_CertifiedPrinterName, objExcel, lRow, lColCounter, False)
'                        End If
'                    End If
'                    '-------------------------------------------------

'                Case dtSpreadsheetCOLPos.lX_Technologies
'                    'update X_Technology column if Technology col exists
'                    lTechnologiesCol = dtSpreadsheetCOLPos.lTechnologies
'                    sX_Technologies = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColCounter))
'                    If lTechnologiesCol > 3 Then
'                        sTechnologies = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lTechnologiesCol))
'                        If sTechnologies = "" Then
'                            If sX_Technologies <> "" Then
'                                Call bUpdateSpreadsheetCell(sDefaultValue, "", sX_Technologies, objExcel, lRow, lColCounter, False)
'                            End If
'                        Else
'                            'find X_Technologies value
'                            sX_Technologies = ""
'                            Call bValidTechnology(lRow, sTechnologies, sX_Technologies, sInvalidTechnologies, lProgramYear)
'                            sX_Technologies = sX_Technologies & sInvalidTechnologies
'                            Call bUpdateSpreadsheetCell(sDefaultValue, sX_Technologies, "", objExcel, lRow, lColCounter, False)
'                        End If
'                    Else
'                        If sX_Technologies <> "" Then
'                            Call bUpdateSpreadsheetCell(sDefaultValue, "", sX_Technologies, objExcel, lRow, lColCounter, False)
'                        End If
'                    End If
'                Case Else
'                    If lMaxMaterialCols > 0 And glSHADING <> NOShading And glSHADING <> NoShadingX Then
'                        'check if material column's values have changed....
'                        Dim lMaterialCounter As Long
'                        For lMaterialCounter = 1 To lMaxMaterialCols
'                            If lColCounter = SpreadsheetMaterialValuesX(lRow, lMaterialCounter).lMaterialCOL Then
'                                sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColCounter))
'                                sDefaultValue = SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sImportConcatenatedMaterial
'                                sOLDCellValue = SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sItemMaterialOldValue
'                                If SpreadsheetMaterialValuesX(lRow, lMaterialCounter).sChangeIndicator = "Y" Then
'                                    'shade Material cell
'                                    Call bUpdateSpreadsheetCell(sDefaultValue, sSpreadsheetValue, sOLDCellValue, objExcel, lRow, lColCounter, False)
'                                Else
'                                    'unshade Material cell
'                                    Call bUpdateSpreadsheetCell(sDefaultValue, sSpreadsheetValue, sOLDCellValue, objExcel, lRow, lColCounter, False)
'                                End If

'                            End If

'                        Next lMaterialCounter

'                    End If
'                    '------ Update the Display Name columns (X_  )
'                    If Len(dtSpreadsheetCOLArray(lColCounter).sLookupField) > 0 Then
'                        sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColCounter))
'                        sLookupValue = sLookupField(CLng(sItemArray(lRow, glProposal_ColPos)), _
'                                                        CLng(sItemArray(lRow, glREV_ColPos)), _
'                                                        dtSpreadsheetCOLArray(lColCounter).sLookupField)
'                        Call bUpdateSpreadsheetCell(sDefaultValue, sLookupValue, sSpreadsheetValue, objExcel, lRow, lColCounter, False)
'                    Else
'                        '---otherwise update other changed cells ----------
'                        If glSHADING <> NOShading And glSHADING <> NoShadingX Then
'                            sSpreadSheetColHeading = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColCounter))
'                            If Microsoft.VisualBasic.Left(sSpreadSheetColHeading, 8) <> "Material" Then 'already done above
'                                sSpreadSheetColHeadingAlias = sGetDelimitedValue(sSpreadSheetColHeading, gsLEFT_DELIMITER, gsRIGHT_DELIMITER)
'                                If Len(sSpreadSheetColHeadingAlias) > 0 Then
'                                    sSpreadSheetColHeading = sSpreadSheetColHeadingAlias
'                                End If

'                                For lUdtColumn = 1 To UBound(dtSpreadsheetCOLArray())
'                                    If (glSHADING = FULLShading) Or (glSHADING = STDShading And dtSpreadsheetCOLArray(lUdtColumn).bExcelShadeCell = True) Then

'                                        If sSpreadSheetColHeading = dtSpreadsheetCOLArray(lUdtColumn).sDB4Field And _
'                                            dtSpreadsheetCOLArray(lUdtColumn).bImport = True Then
'                                            'don't shade columns that have bImport = false, for fields that are on the tables but are not changeable by user

'                                            lArrayCol = dtSpreadsheetCOLArray(lUdtColumn).lArrayCOLNum
'                                            sTableName = dtSpreadsheetCOLArray(lUdtColumn).sDB4Table
'                                            Select Case sTableName
'                                                Case gsItemSpecs_Table
'                                                    sOLDCellValue = sItemSpecsArray_ORIG(lRow, lArrayCol)
'                                                Case gsItem_Table
'                                                    sOLDCellValue = sItemArray_Orig(lRow, lArrayCol)
'                                                Case gsItem_Assortments_Table
'                                                    sOLDCellValue = sAssortmentArray_ORIG(lRow, lArrayCol)
'                                                Case Else
'                                                    Exit For
'                                            End Select
'                                            sSpreadsheetValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColCounter))
'                                            'For DEFAULT VALUES
'                                            If dtSpreadsheetCOLArray(lUdtColumn).vDefault <> "" Then
'                                                If IsNumeric(sSpreadsheetValue) = True Then
'                                                    If CDec(sSpreadsheetValue) = 0 Then sSpreadsheetValue = ""
'                                                End If
'                                                If (IsNull(sSpreadsheetValue) = True Or sSpreadsheetValue = "" Or sSpreadsheetValue = "0") Then
'                                                    sDefaultValue = dtSpreadsheetCOLArray(lUdtColumn).vDefault
'                                                End If
'                                            End If
'                                            Call bUpdateSpreadsheetCell(sDefaultValue, sSpreadsheetValue, sOLDCellValue, objExcel, lRow, lColCounter, False)    'new 10/15/2007
'                                            Exit For
'                                        End If
'                                    End If
'                                Next lUdtColumn

'                            End If
'                        End If
'                    End If
'            End Select
'        Next lColCounter              'for each column in spreadsheet row.....

'        ' Save the Excel file
'SaveExcelFile:
'        Application.DoEvents()                                                        '01/17/2008 office 2007 still causes an error at times?
'        '2011/12/13 - commented below trying to speed up Import by not displaying all messages to user
'        '    Call bUpdateStatusMessage(frmThis, "Saving spreadsheet  ...", True, vbGreen)
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If Version >= 12.0# Then
'            objExcel.Application.Workbooks(1).SaveAs(sfilename, , , , , 0) 'to prevent .xlk backup file being created
'            Application.DoEvents()
'        Else
'            objExcel.Application.Workbooks(1).SaveAs(sfilename, xlNormal)
'        End If
'        Application.DoEvents()
'        '    On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        bUpdateExcelSourceFile = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:
'        ' 01/09/2008 commented below to possibly avoid Office 2007 error on Excel Spreadsheets or get more info
'        '    If Err.Number = 1004 Then
'        '        MsgBox "Spreadsheet was saved as 'Read-only Recommended'" & vbCrLf & vbCrLf & _
'        '        "Open Spreadsheet, go to Tools/Options Security Tab, " & vbCrLf & vbCrLf & _
'        '        "Unclick 'Read-only Recommended', Save & re-import!", vbExclamation + vbOKOnly, _
'        '        "modSpreadSheet-bUpdateExcelSourceFile: Unclick 'Read_only Recommended' checkbox"
'        '    Else

'        MsgBox(Err.Number & ":" & Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bUpdateExcelSourceFile") '01/17/2008
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bUpdateExcelSourceFile , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        '    End If
'        '  Resume Next '' TODO TESTING
'        Resume ExitRoutine
'    End Function

'    'to update ItemNumbers/CoreItems for other Revisions for a Proposal, when a new one is assigned, where previously blank
'    Private Function bUpdateOtherRevsNewItemNumbers(ByVal sProposalNumber As String, ByVal sNewItemNumber As String, _
'                    ByVal sUPCNumber As String, ByVal sMPKUPC As String, ByVal sIPKUPC As String)
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim SQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        'End Function

'        '    Private Function bUpdateOtherRevsNewItemNumbers(ByVal sProposalNumber As String, ByVal sNewItemNumber As String, ByVal ByVal sUPCNumber As String, ByVal sMPKUPC As String, ByVal sIPKUPC As String) As Object
'        '    'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        '    Dim SQL As String 
'        '    Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        Dim Recordcount As Integer
'        Dim bupdatebatchflag As Boolean
'        bupdatebatchflag = False
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            If bWritePrintToLogFile(False, objEXCELName & Space(7) & "Updating  ItemNumbers/CoreItems for other Revisions for a Proposal, when a new one is assigned, where previously blank ,bUpdateOtherRevsNewItemNumbers, for Proposalnumber : " & sProposalNumber, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If

'        bUpdateOtherRevsNewItemNumbers = False
'        '2014/01/20 RAS get count of records
'        SQL = "SELECT Count(1) FROM Item WHERE ItemNumber is NULL and ProposalNumber = " & sProposalNumber
'        rs.Open(SQL, SSDataConn, adOpenDynamic, adLockPessimistic)
'        Recordcount = rs.Fields(0).Value
'        If rs.State <> 0 Then rs.Close()

'        SQL = "SELECT ItemNumber, CoreItemNumber, UPCNumber, MPKUPC, IPKUPC FROM Item WHERE ItemNumber is NULL and ProposalNumber = " & sProposalNumber
'        rs.Open(SQL, SSDataConn, adOpenDynamic, adLockPessimistic)
'        ' 2014/01/10 RAS Changed selection to only bring back the empty (Null) itemnumbers is should help on the query timeouts and speed up looping
'        '2014/01/20 RAS checking if there are any records to update. If there are none, do not do the update statements.
'        Debug.Print(SQL)
'        If Recordcount > 0 Then
'            Do Until rs.EOF
'                If IsBlank(rs!ItemNumber) Then
'                    If sNewItemNumber <> "" Then
'                        rs!ItemNumber = sNewItemNumber
'                        rs!CoreItemNumber = Microsoft.VisualBasic.Left(sNewItemNumber, 6)
'                        rs!UPCNumber = sUPCNumber
'                        rs!MPKUPC = sMPKUPC
'                        rs!IPKUPC = sIPKUPC
'                        rs.Update()
'                        Application.DoEvents()
'                        bupdatebatchflag = True
'                    End If
'                End If
'                rs.MoveNext()
'            Loop

'            If rs.State <> 0 Then
'                If bupdatebatchflag = True Then  ' 2014/01/20 RAS added flag to only update if there was at least one new record
'                    rs.UpdateBatch()
'                    Application.DoEvents()
'                End If
'                rs.Close()
'            End If
'        Else
'            If rs.State <> 0 Then rs.Close()
'        End If
'        rs = Nothing
'        bUpdateOtherRevsNewItemNumbers = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bUpdateOtherRevsNewItemNumbers")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bUpdateOtherRevsNewItemNumbers, Error Number:" & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        'Resume Next
'        Resume ExitRoutine
'    End Function
'    '    '05/07/2009 - VendorNumber cannot be changed for an existing Proposal; this function no longer neccessary - keep in case we need to update other revs again in future
'    'Dim 'Public Function bSaveNewVendorFrom1000(sItemArray As String
'    '    Dim lROWCounter As Long
'    'Dim  udtCOLPos As typSpecialCOLPos)

'    '    ''On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'    '    'Dim sSQL                    As String 
'    '    'Dim rsSaveItemVendor      As ADODB.Recordset
'    '    '
'    '    '    Set rsSaveItemVendor = New ADODB.Recordset
'    '    '    sSQL = "SELECT Item.VendorNumber, Item.FactoryNumber FROM Item " & _
'    '    '            " WHERE Item.ProposalNumber = " & sItemArray(lROWCounter, glProposal_ColPos) & _
'    '    '            " AND Item.Rev  <> " & sItemArray(lROWCounter, glREV_ColPos)
'    '    '
'    '    '    rsSaveItemVendor.Open sSQL, SSDataConn, adOpenDynamic, adLockPessimistic
'    '    End Function

'    Public Function sLookupField(ByVal lProposalNumber As Long, ByVal lRev As Long, ByVal sLookup As String) As String
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sLookupSQL As String
'        Dim sIDField As String
'        Dim sIDTable As String

'        Dim sIDValue As String
'        Dim sNameValue As String

'        Dim sSQL As String

'        sLookupField = ""
'        sLookupSQL = sGetDelimitedValue(sLookup, "{", "}")
'        sIDField = sGetDelimitedValue(sLookup, "<", ">")
'        sIDTable = Microsoft.VisualBasic.Left(sIDField, InStr(1, sIDField, ".") - 1)

'        sSQL = "SELECT " & sIDField & " FROM " & sIDTable & " " & _
'                "WHERE ProposalNumber = " & lProposalNumber & " AND Rev = " & lRev

'        ' Get the value from one of the 3 Item tables
'        Call bGetFieldValue(sSQL, sIDValue, False)

'        If Len(sIDValue) > 0 Then
'            If Len(sLookupSQL) > 0 Then
'                ' Find the associated name value in the lookup table
'                sSQL = sLookupSQL & sAddQuotes(sIDValue)
'                Call bGetFieldValue(sSQL, sNameValue, False)
'                sLookupField = sNameValue
'            Else
'                ' Return the value  as-is
'                sLookupField = sIDValue
'            End If
'        End If

'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-sLookupField")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In sLookupField , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine


'        '2013/05/30 -HN- return error message , for further troubleshooting
'    Private Function bSaveRow(lRow As Long, vSaveArray As Object, ByVal lNbrSaveFields As Long, _
'                ByVal sTableName As String, rsSaveTable As ADODB.Recordset, ByRef lRowChangesFound As Long, _
'                ByRef sErrorMsg As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sConcatenatedComments As String
'        Dim lColCounter As Long
'        Dim sSQL As String
'        Dim bNewRow As Boolean
'        Dim lRevisedUserIDCOL As Long
'        Dim lRevisedDateCOl As Long

'        Dim smessage As String   '2014/01/14 RAS
'        Dim rssavetablevalue As String
'        Dim bErrorCount As Integer   '2014/01/20 RAS
'        bErrorCount = 0
'        bSaveRow = False
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            smessage = Space(3) & "Starting bSaveRow for " & sTableName & " and Row " & lRow
'            If bWritePrintToLogFile(False, objEXCELName & smessage, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        Debug.Print("Proposalnumber = " & vSaveArray(lRow, glProposal_ColPos))  ''' 2014/01/16 RAS done for testing
'    End Function

'        Private Function bSaveRow(ByVal lRow As Long, ByVal  vSaveArray As ,  lNbrSaveFields As Long, ByVal ByVal sTableName As String, ByVal  rsSaveTable As ADODB.Recordset, ByRef lRowChangesFound As Long, ByVal ByRef sErrorMsg As String) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sConcatenatedComments As String
'        Dim lColCounter As Long
'        Dim sSQL As String
'        Dim bNewRow As Boolean
'        Dim lRevisedUserIDCOL As Long
'        Dim lRevisedDateCOl As Long

'        Dim smessage As String   '2014/01/14 RAS
'        Dim rssavetablevalue As String
'        Dim bErrorCount As Integer   '2014/01/20 RAS
'        bErrorCount = 0
'        bSaveRow = False
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            smessage = Space(3) & "Starting bSaveRow for " & sTableName & " and Row " & lRow
'            If bWritePrintToLogFile(False, objEXCELName & smessage, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        Debug.Print("Proposalnumber = " & vSaveArray(lRow, glProposal_ColPos))  ''' 2014/01/16 RAS done for testing
'        Debug.Print("Rev number =     " & vSaveArray(lRow, glREV_ColPos))  ''' 2014/01/16 RAS done for testing
'StartAllOver:
'        ' Generate a recordset from the Save Array, selecting the fieldname for each column
'        sSQL = "SELECT ProposalNumber, Rev, "
'        For lColCounter = glFIXED_COLS + 1 To glFIXED_COLS + lNbrSaveFields
'            sSQL = sSQL & vSaveArray(1, lColCounter) & ","
'        Next

'        sSQL = Microsoft.VisualBasic.Left(sSQL, Len(sSQL) - 1) 'remove comma
'        sSQL = sSQL & " FROM " & sTableName & " " & _
'                "WHERE ProposalNumber = " & vSaveArray(lRow, glProposal_ColPos) & " " & _
'                "AND Rev = " & vSaveArray(lRow, glREV_ColPos)
'    Debug.Print "the sql statement = "; sSQL    ''' 2014/01/16 RAS done for testing
'        rsSaveTable.Open(sSQL, SSDataConn, adOpenKeyset, adLockPessimistic)
'Dim '    rsSaveTable.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic       '2013/04/30 - hn As Object 

'        Application.DoEvents()    'Cursor conflict error?
'        Application.DoEvents()

'        If rsSaveTable.EOF Then
'            bNewRow = True
'            rsSaveTable.AddNew()
'            rsSaveTable(0) = vSaveArray(lRow, glProposal_ColPos)
'            rsSaveTable(1) = vSaveArray(lRow, glREV_ColPos)
'            Application.DoEvents()    '2013/04/28 -HN
'        Else
'            bNewRow = False
'        End If
'        '2014/01/20 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            smessage = Space(3) & " In bSaveRow , Open recordset to see if this is a new item or to update an existing item " & sTableName & " and Row " & lRow
'            If bWritePrintToLogFile(False, objEXCELName & smessage, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        For lColCounter = glFIXED_COLS + 1 To glFIXED_COLS + lNbrSaveFields
'            '        '2014/01/14 RAS Adding info to Trace message    --- put this back when the scheduler is running the ship grid update
'            '        If glTraceFlag = True Then
'            '            smessage = Space(5) & "looping thru columns in bSaveRow for " & sTableName & " and Column: " & vSaveArray(1, lCOLCounter)
'            '            if bWritePrintToLogFile(False, objExcel.Name & smessage, Format(Now(), "yyyymmdd")) = False Then
'            '            End If
'            '        End If
'            If vSaveArray(lRow, lColCounter) = "" Then
'                If bNewRow = False Then
'                    ' Set Text and Memo types to an empty string; Other types to Null
'                    If Not IsNull(rsSaveTable(lColCounter - glFIXED_COLS + 1)) Then
'                        If rsSaveTable(lColCounter - glFIXED_COLS + 1).Type = adVarWChar Or _
'                                rsSaveTable(lColCounter - glFIXED_COLS + 1).Type = adLongVarWChar Then
'                            If rsSaveTable(lColCounter - glFIXED_COLS + 1).Name = "SalesComments" Then      '2011/12/08 - dont save in uppercase
'                                rsSaveTable(lColCounter - glFIXED_COLS + 1) = vSaveArray(lRow, lColCounter)
'                            ElseIf rsSaveTable(lColCounter - glFIXED_COLS + 1).Name = "SpecComments" Then   '2011/12/19 - dont save in uppercase
'                                rsSaveTable(lColCounter - glFIXED_COLS + 1) = vSaveArray(lRow, lColCounter)
'                            Else
'                                rsSaveTable(lColCounter - glFIXED_COLS + 1) = UCase(vSaveArray(lRow, lColCounter))
'                                Application.DoEvents()    '2013/04/28 -HN
'                            End If
'                        Else
'                            rsSaveTable(lColCounter - glFIXED_COLS + 1) = Null
'                            Application.DoEvents()    '2013/04/28 -HN
'                        End If
'                    End If
'                End If
'            Else
'                '2010/02/11
'                If IsBlank(rsSaveTable(lColCounter - glFIXED_COLS + 1)) Then 'some fields have null values, can't compare them..
'                    If IsBlank(vSaveArray(lRow, lColCounter)) Then
'                    Else
'                        '                        GoTo FieldUnequal'TODO - GoTo Statements are redundant in .NET
'                    End If

'                Else

'                    If (CStr(rsSaveTable(lColCounter - glFIXED_COLS + 1)) <> CStr(vSaveArray(lRow, lColCounter))) Then
'FieldUnequal:
'                        '2014/01/14 RAS Adding info to Trace message
'                        If glTraceFlag = True Then
'                            If IsNull(rsSaveTable(lColCounter - glFIXED_COLS + 1)) = True Then rssavetablevalue = "" Else rssavetablevalue = CStr(rsSaveTable(lColCounter - glFIXED_COLS + 1))
'                            smessage = Space(6) & "In bSaveRow for " & sTableName & " and Row " & lRow & " Comparing column " & rsSaveTable(lColCounter - glFIXED_COLS + 1).Name & " for values " & rssavetablevalue & " and " & CStr(vSaveArray(lRow, lColCounter))
'                            If bWritePrintToLogFile(False, objEXCELName & smessage, Format(Now(), "yyyymmdd")) = False Then
'                            End If
'                        End If
'                        Select Case rsSaveTable(lColCounter - glFIXED_COLS + 1).Name
'                            '2010/02/11 - change found: update RevisedUserId & RevisedDate
'                            '- sometimes a user imports a spreadsheet
'                            'with Function code set = C and no changes made, dont update these records
'                            'Material & Component values have to be checked also!!
'                            Case "RevisedUserID"
'                                lRevisedUserIDCOL = lColCounter - glFIXED_COLS + 1

'                            Case "RevisedDate"
'                                lRevisedDateCOl = lColCounter - glFIXED_COLS + 1

'                                '2011/08/17 - check for caret ^ in 1st position, to determine if comments need to be concatenated
'                            Case "DevelopComments", "SeasonalComments", "SeasonalCommentsToVendor", "VendorComments", "SalesComments", "SpecComments" '2011/12/19 added SpecComments
'                                If Microsoft.VisualBasic.Left(CStr(vSaveArray(lRow, lColCounter)), 1) = "^" Then
'                                    '                                    sConcatenatedComments = Replace(vSaveArray(lRow, lCOLCounter), "^", "") & vbCrLf & rsSaveTable(lCOLCounter - glFIXED_COLS + 1)
'                                    'sItemArray_ORIG(lRow, lCOLCounter) holds the original value of these comments fields on the Item table, and we need these for FC=R otherwise the above code would work..

'                                    If IsBlank(sItemArray_Orig(lRow, lColCounter)) Then
'                                        sConcatenatedComments = Replace(vSaveArray(lRow, lColCounter), "^", "")
'                                    Else
'                                        sConcatenatedComments = Replace(vSaveArray(lRow, lColCounter), "^", "") & vbCrLf & sItemArray_Orig(lRow, lColCounter)
'                                    End If
'                                    If rsSaveTable(lColCounter - glFIXED_COLS + 1).Name = "SalesComments" Then '2011/12/08
'                                        vSaveArray(lRow, lColCounter) = sConcatenatedComments
'                                        rsSaveTable(lColCounter - glFIXED_COLS + 1) = sConcatenatedComments
'                                        Application.DoEvents()    '2013/04/28 -HN
'                                    ElseIf rsSaveTable(lColCounter - glFIXED_COLS + 1).Name = "SpecComments" Then '2011/12/19
'                                        vSaveArray(lRow, lColCounter) = sConcatenatedComments
'                                        rsSaveTable(lColCounter - glFIXED_COLS + 1) = sConcatenatedComments
'                                        Application.DoEvents()    '2013/04/28 -HN
'                                    Else
'                                        vSaveArray(lRow, lColCounter) = UCase(sConcatenatedComments)
'                                        Application.DoEvents()    '2013/04/28 -HN
'                                    End If

'                                    lRowChangesFound = lRowChangesFound + 1
'                                Else
'                                    'overwrites old comment
'                                    If rsSaveTable(lColCounter - glFIXED_COLS + 1).Name = "SalesComments" Then          '2011/12/08
'                                        rsSaveTable(lColCounter - glFIXED_COLS + 1) = vSaveArray(lRow, lColCounter)
'                                        Application.DoEvents()    '2013/04/28 -HN
'                                    ElseIf rsSaveTable(lColCounter - glFIXED_COLS + 1).Name = "SpecComments" Then      '2011/12/19
'                                        rsSaveTable(lColCounter - glFIXED_COLS + 1) = vSaveArray(lRow, lColCounter)
'                                        Application.DoEvents()    '2013/04/28 -HN
'                                    Else
'                                        rsSaveTable(lColCounter - glFIXED_COLS + 1) = UCase(vSaveArray(lRow, lColCounter))
'                                        Application.DoEvents()    '2013/04/28 -HN
'                                    End If
'                                    lRowChangesFound = lRowChangesFound + 1
'                                End If
'                                '2014/10/06 RAS Adding special case so these values are always true, later on we update the spreadsheet to True.
'                            Case "LowLeadWholeProduct", "Flammability", "SurfaceLeadPaintReq"
'                                If UCase(vSaveArray(lRow, lColCounter)) <> "TRUE" Then
'                                    If rsSaveTable(lColCounter - glFIXED_COLS + 1) = "FALSE" Then
'                                        rsSaveTable(lColCounter - glFIXED_COLS + 1) = "TRUE"
'                                        lRowChangesFound = lRowChangesFound + 1
'                                    Else
'                                        lRowChangesFound = lRowChangesFound + 1
'                                    End If
'                                End If

'                            Case Else
'                                rsSaveTable(lColCounter - glFIXED_COLS + 1) = UCase(vSaveArray(lRow, lColCounter))
'                                Application.DoEvents()    '2013/04/28 -HN
'                                lRowChangesFound = lRowChangesFound + 1
'                        End Select
'                    End If
'                End If
'                '            rsSaveTable.Update 'delete when error is fixed

'            End If
'        Next lColCounter
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            smessage = Space(5) & "Completed bSaveRow for " & sTableName & " done with saving row:" & lRow
'            If bWritePrintToLogFile(False, objEXCELName & smessage, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        '2013/04/30 -HN- update Item table with RevisedUserId and RevisedDate-previously done in bSavedata
'        If lRowChangesFound > 0 And sTableName = "Item" Then
'            If lRevisedUserIDCOL > 0 Then
'                rsSaveTable(lRevisedUserIDCOL) = gsUserID
'            End If
'            If lRevisedDateCOl > 0 Then
'                rsSaveTable(lRevisedDateCOl) = Now.ToShortDateString()
'            End If
'        End If

'        '    On Error GoTo UpdateTableError                      '2013/05/30 -HN'TODO - On Error must be replaced with Try, Catch, Finally
'        Application.DoEvents()                                            '2013/07/24 -HN- trying to avoid Cursor Conflict error
'        If rsSaveTable.State = 1 Then
'            Application.DoEvents()                                        '2013/07/24 -HN- trying to avoid Cursor Conflict error
'            rsSaveTable.Update() 'delete when error is fixed  '2013/04/30 -HN
'            Application.DoEvents()                                        '2013/07/24 -HN- trying to avoid Cursor Conflict error
'        End If
'        If rsSaveTable.State <> 0 Then
'            Application.DoEvents()                                        '2013/07/24 -HN- trying to avoid Cursor Conflict error
'            rsSaveTable.Close()
'        End If
'        '    Set rsSaveTable = Nothing

'        '2012/01/04 - save changes to ItemFieldHistory table per row, instead of at the end?

'        bSaveRow = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            smessage = Space(3) & "Exiting bSaveRow for " & sTableName & " and Row " & lRow
'            If bWritePrintToLogFile(False, objEXCELName & smessage, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        Exit Function
'ErrorHandler:

'        'If rsSaveTable.State <> 0 Then rsSaveTable.Close
'        'Resume Next ''' testing only
'        '    If bPROPOSALFormIndicator = False Then
'        '        If Err.Number = -2147467259 Then  '  2014/01/20 RAS Adding this to have it try the query again.
'        '           If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'        '                smessage = "In bSaveRow  Retrying ,for " & sTableName & " and Row " & lRow & " There has been and error. Error Number:" & Err.Number & "Error Description: " & Err.Description
'        '                If bWritePrintToLogFile(False, objExcel.name & smessage, "ErrorMessageLog") = False Then
'        '                End If
'        '            End If
'        '            bErrorCount = bErrorCount + 1
'        '            If bErrorCount < 6 Then
'        '                If rsSaveTable.State <> 0 Then
'        '                    rsSaveTable.CancelUpdate
'        '                    rsSaveTable.Close
'        '                End If
'        ''                GoTo StartAllOver    'TOP of the select again.'TODO - GoTo Statements are redundant in .NET
'        '            Else
'        '                If rsSaveTable.State <> 0 Then
'        '                    rsSaveTable.CancelUpdate
'        '                    rsSaveTable.Close
'        '                End If
'        '                Resume ExitRoutine
'        '            End If
'        '        End If
'        '    End If
'        If rsSaveTable.State <> 0 Then rsSaveTable.Close()
'        MsgBox(Err.Description & vbCrLf & vSaveArray(1, lColCounter) & ": " & vSaveArray(lRow, lColCounter) & _
'               vbCrLf & "SYS ADMIN: Check value of above Field and/or Field definition on table(s).", vbExclamation + vbMsgBoxSetForeground, "CONTACT SYS ADMIN: (modSpreadSheet-bSaveRow)")
'        ' Resume Next  '2013/05/30 -HN- only for testing here
'        '2014/09/10 RAS moved here so it would not lose the error message when writing to the log
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            smessage = "In bSaveRow for " & sTableName & " and Row " & lRow & " There has been and error. Error Number:" & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        Resume ExitRoutine
'UpdateTableError:  '2013/05/30 -HN- sometimes record was created but seems to have an error on Update or Close?

'        '    If bPROPOSALFormIndicator = False Then
'        '        If Err.Number = -2147467259 Then  '  2014/01/20 RAS Adding this to have it try the query again.
'        '            If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'        '                smessage = "In bSaveRow  Retrying ,for " & sTableName & " and Row " & lRow & " There has been and error. Error Number:" & Err.Number & "Error Description: " & Err.Description
'        '                If bWritePrintToLogFile(False, objExcel.name & smessage, "ErrorMessageLog") = False Then
'        '                End If
'        '            End If
'        '            bErrorCount = bErrorCount + 1
'        '            If bErrorCount < 6 Then
'        '                If rsSaveTable.State <> 0 Then
'        '                    rsSaveTable.CancelUpdate
'        '                    rsSaveTable.Close
'        '                End If
'        ''                GoTo StartAllOver    'TOP of the select again.'TODO - GoTo Statements are redundant in .NET
'        '            Else
'        '                If rsSaveTable.State <> 0 Then
'        '                    rsSaveTable.CancelUpdate
'        '                    rsSaveTable.Close
'        '                End If
'        '                Resume ExitRoutine
'        '            End If
'        '
'        '        End If
'        '    End If
'        MsgBox(Err.Description, vbCritical + vbMsgBoxSetForeground, "modSpreadsheet-bSaveRow- UpdateTableError!")
'        If glErrorMessageFlag = True Then
'            smessage = "In bSaveRow for " & sTableName & " and Row " & lRow & " There has been and error. Error Number:" & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume Next

'    End Function
'    Private Function bCheckExistingItemNumber(ByVal sFunctioncode As String, ByVal sProposalNumber As String, ByVal sItemNumber As String, ByRef sItemErrMsg As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim rs As ADODB.Recordset
'        Dim SQL As String
'        Dim sRSItemNumber As String
'        bCheckExistingItemNumber = False
'        sItemErrMsg = ""
'        rs = New ADODB.Recordset

'        If sFunctioncode <> gsNEW_PROPOSAL Then  'for C and R
'            If sProposalNumber <> "" And sItemNumber <> "" And sItemNumber <> gsNEW_ITEM_NBR Then
'                SQL = "SELECT DISTINCT ItemNumber FROM Item WHERE ProposalNumber = " & CLng(sProposalNumber) & _
'                    " AND ItemNumber IS NOT NULL AND ItemNumber <> ''"

'                rs.Open SQL
'                Dim SSDataConn As Object
'                Dim adOpenStatic As Object
'                Dim adLockReadOnly As Object

'                If rs.EOF Then
'                    '                GoTo ExitX'TODO - GoTo Statements are redundant in .NET
'                Else
'                    If rs.Recordcount > 1 Then
'                        sItemErrMsg = " - " & rs.Recordcount & " different ItemNumbers found in DB4, for Proposal."
'                        '                    GoTo ExitX'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        '11/14/2007
'                        If IsNull(rs!ItemNumber) Then
'                            sRSItemNumber = ""
'                        Else
'                            sRSItemNumber = rs!ItemNumber
'                        End If

'                        If sItemNumber <> sRSItemNumber Then
'                            sItemErrMsg = " - ItemNumber on grid does not match ItemNumber[" & rs!ItemNumber & "] in DB4 for BaseProposal:" & sProposalNumber & vbCrLf & _
'                           "           (If NEW Item, make NEW Proposal:FC=A,ItemNumber=A,change VendorItemNumber)"
'                            '                        GoTo ExitX'TODO - GoTo Statements are redundant in .NET
'                        Else
'                            '                        GoTo ExitX'TODO - GoTo Statements are redundant in .NET
'                        End If
'                    End If
'                End If
'            Else

'                If sProposalNumber <> "" And (sItemNumber = "" Or sItemNumber = gsNEW_ITEM_NBR) Then
'                    SQL = "SELECT DISTINCT ItemNumber FROM Item WHERE ProposalNumber = " & CLng(sProposalNumber) & " " & _
'                            "AND  (ItemNumber IS NOT NULL AND ItemNumber <> '')"
'                    rs.Open SQL
'                    Dim SSDataConn As Object
'                    Dim adOpenStatic As Object
'                    Dim adLockReadOnly As Object

'                    If Not rs.EOF Then
'                        If rs.Recordcount > 1 Then        'cant have more than one ItemNumber for a Proposal, except a blank one
'                            sItemErrMsg = " -  " & rs.Recordcount & " ItemNumbers found in DB4, for this Proposal."
'                            '                        GoTo ExitX'TODO - GoTo Statements are redundant in .NET
'                        Else
'                            If IsNull(rs!ItemNumber) Then
'                                sRSItemNumber = ""
'                            Else
'                                sRSItemNumber = rs!ItemNumber
'                            End If

'                            If sRSItemNumber <> sItemNumber Then
'                                If sItemNumber = "" Then
'                                    sItemErrMsg = " - Cannot delete existing ItemNumber[" & rs!ItemNumber & "] for this Proposal."

'                                ElseIf sItemNumber = "A" Then
'                                    sItemErrMsg = " - ItemNumber[" & rs!ItemNumber & "] already exists for this Proposal!"

'                                Else

'                                    sItemErrMsg = " - ItemNumber on grid, does not match ItemNumber[" & rs!ItemNumber & "] for this Proposal."
'                                End If
'                                '                            GoTo ExitX'TODO - GoTo Statements are redundant in .NET
'                            Else
'                                '                            GoTo ExitX'TODO - GoTo Statements are redundant in .NET
'                            End If

'                        End If
'                    Else
'                        '                   GoTo ExitX'TODO - GoTo Statements are redundant in .NET
'                    End If

'                End If
'            End If

'        Else
'            'for FunctionCode=A
'            '12/13/2007 commented because this should be caught in bFindMultipleProposalNumbersInDB
'            '        If sProposalNumber = "" And sItemNumber <> "" Then
'            '        If sItemNumber <> "" And sItemNumber <> "A" Then            '12/06/2007
'            '            SQL = "SELECT CoreItemNumber FROM Item WHERE ItemNumber = '" & sItemNumber & "'"
'            '            rs.Open SQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockReadOnly As Object

'            '            If Not rs.EOF Then
'            '                sItemErrMsg = " - ItemNumber[" & sItemNumber & "] exists in " & rs.RecordCount & " Proposals."
'            '            End If
'            '        End If
'        End If
'        '    End If
'ExitX:
'        If sItemErrMsg <> "" Then
'            bCheckExistingItemNumber = True
'        Else
'            bCheckExistingItemNumber = False
'        End If

'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rs.State <> 0 Then rs.Close()
'        rs = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bCheckExistingItemNumber")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCheckExistingItemNumbers , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitX
'    End Function
'    'Private Function bMatchExistingItemNumber(ByVal sProposalNumber As String, ByVal sItemNumber As String, _
'    '                                          Optional bFirst6Digits As Boolean) As Boolean
'    ''On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'    'Dim lNbrMatchProposal   As Long
'    ' Dim lNbrMatchItem As Long

'    'Dim rsMatchProposal     As ADODB.Recordset
'    ' Dim rsMatchItem As ADODB.Recordset

'    'Dim sSQL                As String 
'    '
'    '    bMatchExistingItemNumber = False
'    '    If sProposalNumber = "" Then
'    '        bMatchExistingItemNumber = True
'    '        Exit Function
'    '    End If
'    '    Set rsMatchProposal = New ADODB.Recordset
'    '    Set rsMatchItem = New ADODB.Recordset
'    '
'    '    '11/02/2007 commented below - hn
'    '    sSQL = "SELECT ItemNumber FROM " & gsItem_Table & " WHERE ProposalNumber = " & CLng(sProposalNumber) '& " " & _
'    '                '"AND (ItemNumber IS NOT NULL AND ItemNumber <> '')"
'    '
'    '    rsMatchProposal.Open sSQL As Object 
'    ' SSDataConn 
'    ' adOpenStatic
'    ' adLockOptimistic

'    '
'    '    If Not rsMatchProposal.EOF Then
'    '        lNbrMatchProposal = rsMatchProposal.RecordCount
'    '
'    '        If bFirst6Digits = True Then
'    '            sSQL = "SELECT Microsoft.VisualBasic.Left(ItemNumber, 6) FROM " & gsItem_Table & " WHERE ProposalNumber = " & CLng(sProposalNumber) & " " & _
'    '                    "AND Microsoft.VisualBasic.Left(ItemNumber ,6) = " & sAddQuotes(Left(sItemNumber, 6))
'    '            bFirst6Digits = False
'    '        Else
'    '
'    '            sSQL = "SELECT ItemNumber FROM " & gsItem_Table & " WHERE ProposalNumber = " & CLng(sProposalNumber) & " " & _
'    '                    "AND ItemNumber = " & sAddQuotes(sItemNumber)
'    '        End If
'    '        rsMatchItem.Open sSQL As Object 
'    ' SSDataConn
'    ' adOpenStatic
'    ' adLockOptimistic 

'    '        If Not rsMatchItem.EOF Then
'    '            lNbrMatchItem = rsMatchItem.RecordCount
'    '        End If
'    '        If IsNull(rsMatchProposal!ItemNumber) Or rsMatchProposal!ItemNumber = "" Then   'new 11/02/2007 - hn
'    '        Else
'    '            If lNbrMatchItem <> lNbrMatchProposal Then
'    ''                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'    '            End If
'    '        End If
'    '    End If
'    '    bMatchExistingItemNumber = True
'    'ExitRoutine:
'    ''    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'    '    If rsMatchProposal.State <> 0 Then rsMatchProposal.Close
'    '    Set rsMatchProposal = Nothing
'    '    If rsMatchItem.State <> 0 Then rsMatchItem.Close
'    '    Set rsMatchItem = Nothing
'    '    Exit Function
'    'ErrorHandler:
'    '    MsgBox Err.Description, vbExclamation, "modSpreadSheet-bMatchExistingItemNumber"
'    '    Resume ExitRoutine
'    '
'    'End Function

'    Private Function bCheckExistingCoreItemNumber(ByVal sProposalNumber As String, ByVal sItemNumber As String, ByRef sCoreItem As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim rs As ADODB.Recordset
'        Dim SQL As String
'        bCheckExistingCoreItemNumber = False
'        rs = New ADODB.Recordset

'        If sItemNumber = "A" Or sItemNumber = "" Then
'            SQL = "SELECT CoreItemNumber, ItemNumber FROM Item WHERE ProposalNumber = " & sProposalNumber
'Dim         rs.Open SQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockReadOnly As Object

'            If Not rs.EOF Then
'                'ItemNumber = A the following fields dont exist yet..
'                If (IsNull(rs!CoreItemNumber) Or rs!CoreItemNumber = "") And (IsNull(rs!ItemNumber) Or rs!ItemNumber = "") Then
'                    bCheckExistingCoreItemNumber = True
'                    '                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                Else
'                    sCoreItem = rs!CoreItemNumber
'                End If
'            Else
'                sCoreItem = "Unknown"
'            End If

'        Else

'            SQL = "SELECT DISTINCT CoreItemNumber FROM Item Where ProposalNumber = " & sProposalNumber & " AND CoreItemNumber IS NOT NULL"   '12/20/2007
'Dim          rs.Open SQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockReadOnly As Object

'            If Not rs.EOF Then
'                If IsBlank(rs!CoreItemNumber) Then
'                    sCoreItem = ""
'                Else
'                    sCoreItem = rs!CoreItemNumber
'                End If
'                If sCoreItem = Microsoft.VisualBasic.Left(sItemNumber, 6) Then
'                    bCheckExistingCoreItemNumber = True
'                Else
'                    bCheckExistingCoreItemNumber = False
'                End If
'            Else
'                sCoreItem = "Unknown"
'            End If

'        End If
'ExitRoutine:
'        If rs.State <> 0 Then rs.Close()
'        rs = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bCheckExistingCoreItemNumber")

'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCheckExistingCoreItemNumber , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'    Private Function bFindVendorFactoryInDB(ByVal sVendorNumber As String, ByVal sFactoryNumber As String) As Boolean
'        ' Verify that the Vendor is assocated with the Factory in the VendorFactories table
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        bFindVendorFactoryInDB = False

'        If sVendorNumber = 1000 Or sFactoryNumber = 1000 Then
'            bFindVendorFactoryInDB = True
'        Else
'            Dim RSVendorFactory As ADODB.Recordset
'            RSVendorFactory = New ADODB.Recordset
'            sSQL = "SELECT * FROM VendorFactory" & _
'                    " WHERE VendorNumber = " & sVendorNumber & " AND FactoryNumber = " & sFactoryNumber

'Dim         RSVendorFactory.Open sSQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockOptimistic As Object


'            If Not RSVendorFactory.EOF Then
'                bFindVendorFactoryInDB = True
'            End If
'            RSVendorFactory.Close()
'            RSVendorFactory = Nothing
'        End If

'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bFindVendorFactoryInDB")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFindVendorFactoryInDB , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Private Function bFindMultipleProposalNumbersInDB(ByVal sFunctioncode As String, ByVal sItemNumber As String, ByVal sCategoryCode As String, ByVal sVendorNumber As String, ByVal sFactoryNumber As String, ByVal lcustomernumber As Long, ByVal sVendorItemNumber As String, ByRef sErrorMsg As String) As Object

'        ' Ensure there aren't multiple proposal numbers in the database
'        ' for the ItemNumber / VendorNumber / FactoryNumber / CustomerNumber combination

'        'IMPORTANT:
'        '  the following will not be TRUE IF: Item.CategoryCode ='EL','ELLS','ELOTH','ELOUT' OR 'BG'
'        '  we forget this all the time on Imports!!

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim rs As ADODB.Recordset
'        Dim SQL As String
'        Dim sCustomerNumberSQL As String
'        Dim sCustomerNumberMsg As String

'        Const CatMsg = "(=EL,ELLS,ELOTH,ELOUT,BG)"

'        bFindMultipleProposalNumbersInDB = False
'        rs = New ADODB.Recordset

'        Select Case lcustomernumber
'            Case gs100_ACCOUNT, gs999PD_ACCOUNT, gs998PD_ACCOUNT
'                sCustomerNumberSQL = " AND Item.CustomerNumber  IN  (100, 998, 999) "
'                sCustomerNumberMsg = "CustomerNumber=[100/998/999]"
'            Case Else
'                sCustomerNumberSQL = " AND (Item.CustomerNumber = " & lcustomernumber & ") "
'                sCustomerNumberMsg = "CustomerNumber=[" & lcustomernumber & "]"
'        End Select

'        If sFunctioncode = gsNEW_PROPOSAL And IsBlank(sItemNumber) = False And sItemNumber <> gsNEW_ITEM_NBR Then
'            Select Case sCategoryCode  'this is categorycode on spreadsheet
'                Case "EL", "BG", "ELLS", "ELOTH", "ELOUT"
'                    'OPTION 1: FC=A, ItemNumber Exists, Cat='EL', 'BG', 'ELLS','ELOTH','ELOUT'
'                    '------------------------------------------------------------------------
'                    SQL = "SELECT DISTINCT Item.ProposalNumber, Item.Rev  FROM Item " & vbCrLf & _
'                          "WHERE Item.ItemNumber = " & sAddQuotes(sItemNumber) & _
'                          "AND Item.VendorNumber = " & sAddQuotes(sVendorNumber) & " " & _
'                          "AND Item.FactoryNumber = " & sAddQuotes(sFactoryNumber) & " " & vbCrLf & _
'                          sCustomerNumberSQL & vbCrLf & _
'                          "AND Item.CategoryCode IN ('EL', 'BG', 'ELLS','ELOTH','ELOUT') " & vbCrLf & _
'                          "ORDER BY Item.ProposalNumber, Item.Rev DESC"
'Dim                 rs.Open SQL As Object 
'                    Dim SSDataConn As Object
'                    Dim adOpenStatic As Object
'                    Dim adLockReadOnly As Object

'                    If Not rs.EOF Then
'                        sErrorMsg = sErrorMsg & " PROPOSAL ALREADY EXISTS FOR:" & vbCrLf & _
'                            "FunctionCode= " & sFunctioncode & _
'                            ", Item[" & sItemNumber & "]," & _
'                            "Cust[" & lcustomernumber & "]" & _
'                            "Vendor[" & sVendorNumber & "]," & _
'                            "Factory[" & sFactoryNumber & "]" & _
'                            "Cat[" & sCategoryCode & "] AND" & CatMsg & vbCrLf
'                        Do Until rs.EOF
'                            sErrorMsg = sErrorMsg & _
'                            "      Proposal:" & rs!ProposalNumber & " Rev:" & rs!Rev & " - already exists!" & vbCrLf

'                            rs.MoveNext()
'                        Loop
'                        '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        '                    GoTo ExitTrue'TODO - GoTo Statements are redundant in .NET
'                    End If

'                Case Else
'                    'OPTION 2: FC=A, ItemNumber Exists, Cat NOT ='EL', 'BG', 'ELLS','ELOTH','ELOUT'
'                    '------------------------------------------------------------------------------
'                    SQL = "SELECT DISTINCT Item.ProposalNumber, Item.Rev  FROM Item " & vbCrLf & _
'                          "WHERE Item.ItemNumber = " & sAddQuotes(sItemNumber) & _
'                          "AND Item.VendorNumber = " & sAddQuotes(sVendorNumber) & " " & _
'                          "AND Item.FactoryNumber = " & sAddQuotes(sFactoryNumber) & " " & vbCrLf & _
'                          "AND Item.VendorItemNumber = " & sAddQuotes(sVendorItemNumber) & " " & _
'                          "AND Item.VendorItemNumber <> '' " & _
'                          sCustomerNumberSQL & vbCrLf & _
'                          "AND Item.CategoryCode NOT IN ('EL', 'BG', 'ELLS','ELOTH','ELOUT') " & vbCrLf & _
'                            "ORDER BY Item.ProposalNumber, Item.Rev DESC"

'Dim                 rs.Open SQL As Object 
'                    Dim SSDataConn As Object
'                    Dim adOpenStatic As Object
'                    Dim adLockReadOnly As Object

'                    If Not rs.EOF Then
'                        sErrorMsg = sErrorMsg & " PROPOSAL ALREADY EXISTS FOR:" & vbCrLf & _
'                                    "FunctionCode= " & sFunctioncode & _
'                                     ", VendorItemNbr[" & sVendorItemNumber & "]," & _
'                                    "Item[" & sItemNumber & "]," & _
'                                    "Cust[" & lcustomernumber & "]" & _
'                                    "Vendor[" & sVendorNumber & "]," & _
'                                    "Factory[" & sFactoryNumber & "]" & _
'                                    "Cat[" & sCategoryCode & "] NOT" & CatMsg & vbCrLf
'                        Do Until rs.EOF
'                            sErrorMsg = sErrorMsg & _
'                            "     Proposal:" & rs!ProposalNumber & " Rev:" & rs!Rev & " - already exists!" & vbCrLf

'                            rs.MoveNext()
'                        Loop
'                        '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        '                    GoTo ExitTrue'TODO - GoTo Statements are redundant in .NET
'                    End If

'            End Select
'        End If

'        'OPTION 3: FC=A, ItemNumber Blank, Cat='EL', 'BG', 'ELLS','ELOTH','ELOUT'
'        '------------------------------------------------------------------------------
'        If sFunctioncode = gsNEW_PROPOSAL And IsBlank(sItemNumber) = True Then
'            Select Case sCategoryCode
'                Case "EL", "BG", "ELLS", "ELOTH", "ELOUT"
'                    '                GoTo ExitTrue'TODO - GoTo Statements are redundant in .NET
'                Case Else
'                    'OPTION 4: FC=A, ItemNumber Blank, Cat NOT ='EL', 'BG', 'ELLS','ELOTH','ELOUT'
'                    '------------------------------------------------------------------------------
'                    '10/16/2008 - hn - added test for blank item number below
'                    SQL = "SELECT DISTINCT Item.ProposalNumber, Item.Rev  FROM Item " & vbCrLf & _
'                          "WHERE Item.ItemNumber <> '' " & _
'                          "AND Item.VendorNumber = " & sAddQuotes(sVendorNumber) & " " & _
'                          "AND Item.FactoryNumber = " & sAddQuotes(sFactoryNumber) & " " & vbCrLf & _
'                          "AND Item.VendorItemNumber = " & sAddQuotes(sVendorItemNumber) & " " & _
'                          "AND Item.VendorItemNumber <> '' " & _
'                          sCustomerNumberSQL & vbCrLf & _
'                          "AND Item.CategoryCode NOT IN ('EL', 'BG', 'ELLS','ELOTH','ELOUT') " & vbCrLf & _
'                            "ORDER BY Item.ProposalNumber, Item.Rev DESC"

'Dim                 rs.Open SQL As Object 
'                    Dim SSDataConn As Object
'                    Dim adOpenStatic As Object
'                    Dim adLockReadOnly As Object

'                    If Not rs.EOF Then
'                        sErrorMsg = sErrorMsg & " PROPOSAL ALREADY EXISTS FOR:" & vbCrLf & _
'                                 "FunctionCode= " & sFunctioncode & _
'                                 ", VendorItemNbr[" & sVendorItemNumber & "]," & _
'                                 "Item[" & sItemNumber & "]," & _
'                                 "Cust[" & lcustomernumber & "]" & _
'                                 "Vendor[" & sVendorNumber & "]," & _
'                                 "Factory[" & sFactoryNumber & "]" & _
'                                 "Cat[" & sCategoryCode & "] NOT" & CatMsg & vbCrLf
'                        Do Until rs.EOF
'                            sErrorMsg = sErrorMsg & _
'                            "     Proposal:" & rs!ProposalNumber & " Rev:" & rs!Rev & " - already exists!" & vbCrLf
'                            rs.MoveNext()
'                        Loop
'                        '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        '                    GoTo ExitTrue'TODO - GoTo Statements are redundant in .NET
'                    End If
'            End Select
'        End If

'        'OPTION 5: FC=any function code, ItemNumber = A, Cat='EL', 'BG', 'ELLS','ELOTH','ELOUT'
'        '------------------------------------------------------------------------------
'        If sItemNumber = "A" Then
'            Select Case sCategoryCode
'                Case "EL", "BG", "ELLS", "ELOTH", "ELOUT"
'                    '                GoTo ExitTrue'TODO - GoTo Statements are redundant in .NET
'                Case Else
'                    'OPTION 6: FC=any function code, ItemNumber = A, Cat NOT ='EL', 'BG', 'ELLS','ELOTH','ELOUT'
'                    '--------------------------------------------------------------------------------------------
'                    '09/17/2008 -hn - added test for blank ItemNumber below
'                    SQL = "SELECT DISTINCT Item.ProposalNumber, Item.Rev, Item.ItemNumber  FROM Item " & vbCrLf & _
'                              "WHERE Item.ItemNumber <> '' " & vbCrLf & _
'                              "AND Item.VendorNumber = " & sAddQuotes(sVendorNumber) & " " & _
'                              "AND Item.FactoryNumber = " & sAddQuotes(sFactoryNumber) & " " & vbCrLf & _
'                              "AND Item.VendorItemNumber = " & sAddQuotes(sVendorItemNumber) & " " & _
'                              "AND Item.VendorItemNumber <> '' " & _
'                              sCustomerNumberSQL & vbCrLf & _
'                              "AND Item.CategoryCode NOT IN ('EL', 'BG', 'ELLS','ELOTH','ELOUT') " & vbCrLf & _
'                                "ORDER BY Item.ProposalNumber, Item.Rev DESC"

'Dim                 rs.Open SQL As Object 
'                    Dim SSDataConn As Object
'                    Dim adOpenStatic As Object
'                    Dim adLockReadOnly As Object

'                    If Not rs.EOF Then
'                        sErrorMsg = sErrorMsg & " PROPOSAL ALREADY EXISTS FOR:" & vbCrLf & _
'                                "FunctionCode= " & sFunctioncode & _
'                                ", VendorItemNbr[" & sVendorItemNumber & "]," & _
'                                "Item[" & sItemNumber & "]," & _
'                                "Cust[" & lcustomernumber & "]" & _
'                                "Vendor[" & sVendorNumber & "]," & _
'                                "Factory[" & sFactoryNumber & "]" & _
'                                "Cat[" & sCategoryCode & "] NOT" & CatMsg & vbCrLf
'                        Do Until rs.EOF
'                            '09/17/2008 - hn - added ItemNumber in error msg below
'                            sErrorMsg = sErrorMsg & _
'                            "      Proposal:" & rs!ProposalNumber & " Rev:" & rs!Rev & " Item[ " & rs!ItemNumber & " ] - already exists!" & vbCrLf

'                            rs.MoveNext()
'                        Loop
'                        '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    Else
'                        '                    GoTo ExitTrue'TODO - GoTo Statements are redundant in .NET
'                    End If
'            End Select
'        End If
'ExitTrue:
'        bFindMultipleProposalNumbersInDB = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rs.State <> 0 Then rs.Close()
'        rs = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bFindMultipleProposalNumbersInDB")

'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFindMultipleProposalNumbersInDB , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function
'Private Function bFindDuplicateItemsOnSpreadsheet(ByVal lCurrentRow As Long, _
'                                    ByVal lRowsOnSheet As Long, _
'                                    udtCOLPos As typSpecialCOLPos, _
'                                    sItemSPECSArray() As String, sItemArray() As String, _
'                                    ByRef sErrorMsg As String) As Boolean
'        ' Each ItemNumber/VendorNumber/FactoryNumber/CustomerNumber combination
'        ' should only be on the spreadsheet once
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lCounter As Long
'        Dim sDuplicateRowsList As String
'        Dim sItemNumber As String
'        Dim sCheckItemNumber As String

'        Dim sVendorNumber As String
'        Dim sCheckVendorNumber As String

'        Dim sFactoryNumber As String
'        Dim sCheckFactoryNumber As String

'        Dim sCustomerNumber As String
'        Dim sCheckCustomerNumber As String

'        Dim sFunctioncode As String
'        bFindDuplicateItemsOnSpreadsheet = False
'        sItemNumber = sItemArray(lCurrentRow, udtCOLPos.lItemNumber)
'        sVendorNumber = sItemArray(lCurrentRow, udtCOLPos.lVendorNumber)
'        sFactoryNumber = sItemArray(lCurrentRow, udtCOLPos.lFactoryNumber)
'        sCustomerNumber = sItemArray(lCurrentRow, udtCOLPos.lcustomernumber)

'        For lCounter = glDATA_START_ROW To lRowsOnSheet
'            If lCounter <> lCurrentRow Then
'                sCheckItemNumber = sItemArray(lCounter, udtCOLPos.lItemNumber)
'                sCheckVendorNumber = sItemArray(lCounter, udtCOLPos.lVendorNumber)
'                sCheckFactoryNumber = sItemArray(lCounter, udtCOLPos.lFactoryNumber)
'                sCheckCustomerNumber = sItemArray(lCounter, udtCOLPos.lcustomernumber)
'                '11/08/2007
'                If sItemNumber = sCheckItemNumber _
'                    And sItemNumber <> "A" And sItemNumber <> "" _
'                    And sVendorNumber = sCheckVendorNumber _
'                    And sFactoryNumber = sCheckFactoryNumber _
'                    And sCustomerNumber = sCheckCustomerNumber Then

'                    sDuplicateRowsList = sDuplicateRowsList & CStr(lCounter) & ", "
'                End If
'            End If
'        Next lCounter

'        If Len(sDuplicateRowsList) > 0 Then
'            sErrorMsg = "Duplicate rows found :" & _
'                            "ItemNumber[" & sItemNumber & "], " & _
'                            "Vendor[" & sVendorNumber & "], " & _
'                            "Factory[" & sFactoryNumber & "], " & _
'                            "CustomerNumber[" & sCustomerNumber & "], Duplicate Row: " & sDuplicateRowsList
'            sErrorMsg = Microsoft.VisualBasic.Left(sErrorMsg, Len(sErrorMsg) - 2)
'            bFindDuplicateItemsOnSpreadsheet = True
'        End If
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bFindDuplicateItemsOnSpreadsheet")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFindDuplicateItemsOnSpreadsheet , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function
'    Private Function bGetNextRevNumber(ByVal sProposalNumber As String, ByRef sRevNumber As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim rsNextRev As ADODB.Recordset
'        Dim sTempRevNumber As String
'        Dim sSQL As String

'        bGetNextRevNumber = False
'        rsNextRev = New ADODB.Recordset
'        sSQL = "SELECT MAX(Rev) FROM " & gsItem_Table & " WHERE ProposalNumber = " & CLng(sProposalNumber)
'Dim     rsNextRev.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic As Object

'        If rsNextRev.EOF Then
'            MsgBox("Could not obtain next Rev for Proposal " & sProposalNumber & vbCrLf & _
'                        "(Not in Item table)", vbExclamation + vbMsgBoxSetForeground, "bGetNextRevNumber")
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        Else
'            sTempRevNumber = rsNextRev(0)
'            sTempRevNumber = CStr(CLng(sTempRevNumber) + 1)
'            sRevNumber = sTempRevNumber
'        End If
'        bGetNextRevNumber = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rsNextRev.State <> 0 Then rsNextRev.Close()
'        rsNextRev = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bGetNextRevNumber")

'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bGetNextRevNumber , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'    Private Function bOverwriteDuplicateFields(ByVal lRow As Long, ByVal dtCOLPos As typSpecialCOLPos) As Object
'        ' This forces the fields that should be the same to be the same
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sItemNumber As String
'        Dim lcustomernumber As Long
'        Dim sVendorFactoryNumber As String
'        Dim sCustomerItemNumber As String

'        Dim dtNow As Date
'        Dim sSQL As String
'        Dim vReturnValue As Object


'        bOverwriteDuplicateFields = False
'        ' Overwrite CoreItem w/ Left 6 characters of Item.ItemNumber
'        sItemNumber = sItemArray(lRow, dtCOLPos.lItemNumber)
'        ' no longer required, removed from Item table 02/20/2007
'        '    sItemSPECSArray(lRow, dtColPos.lItemNumber) = sItemNumber
'        sItemArray(lRow, dtCOLPos.lCoreItemNumber) = Microsoft.VisualBasic.Left(sItemNumber, 6)

'        ' Overwrite Vendor FactoryNumber w/ Vendor Number (if Vendor FactoryNumber is blank)
'        sVendorFactoryNumber = sItemArray(lRow, dtCOLPos.lFactoryNumber)
'        If Len(sVendorFactoryNumber) = 0 Then
'            sItemArray(lRow, dtCOLPos.lFactoryNumber) = sItemArray(lRow, dtCOLPos.lVendorNumber)
'        End If

'        ' Overwrite SQ_2 w/ CustomerItemNumber in order to port the data to Excel Gold
'        sCustomerItemNumber = sItemArray(lRow, dtCOLPos.lCustomerItemNumber)
'        sItemArray(lRow, dtCOLPos.lSQ2) = sCustomerItemNumber

'        ' The Item.RevisedDate is set to the date/time of the update
'        dtNow = Now.ToShortDateString()
'        sItemArray(lRow, dtCOLPos.lRevisedDate) = CStr(dtNow)

'        ' Set the RevisedUserID field to the ID of the User who logged on to the system
'        sItemArray(lRow, dtCOLPos.lRevisedUserID) = gsUserID
'        bOverwriteDuplicateFields = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bOverwriteDuplicateFields")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bOverwriteDuplicateFields , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'Dim Public Function bCreateImportLogFile(objExcel As Excel.Application
'Dim  _ As Object 

'    End Function

'        Friend Overridable Function bCreateImportLogFile(ByVal objExcel As Excel.Application, ByVal ByVal sfilename As String, ByVal ByVal simportFileName As String, ByVal ByVal lNewProposals As Long, ByVal lNewRevisions As Long, ByVal ByVal lChangedProposals As Long, ByVal ByVal lReferenced100Accts As Long, ByVal ByVal lRollbackProposals As Long, ByVal dtNow As [Date], ByVal  bRefreshImportPhotos As Boolean, ByVal lAlternatePhotoCOLPos As Long) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sLastModified As String
'        Dim sFileSize As String

'        bCreateImportLogFile = False

'        If bGetFileInformation(simportFileName, sLastModified, sFileSize) = False Then
'            MsgBox("Could not get file information for:" & vbCrLf & simportFileName, _
'                    vbExclamation + vbMsgBoxSetForeground, "bCreateImportLogFile")
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If

'        ' Do not show the instance of Excel on the screen
'        objExcel.Visible = False
'        objExcel.Application.Workbooks.Add()

'        ' Excel includes 3 sheets by default - remove 2 of them
'        '2014/04/10 RAS find out how many sheets and delete them all but the first
'        Dim iws As Integer
'        Dim ii As Integer

'        iws = objExcel.Application.Workbooks(1).Sheets.Count
'        If iws > 1 Then
'            For ii = iws To 2 Step -1
'                objExcel.Application.Workbooks(1).Sheets(ii).Delete()
'                '    objEXCEL.Application.Workbooks(1).Sheets(2).Delete
'                '    objEXCEL.Application.Workbooks(1).Sheets(2).Delete
'            Next
'        End If

'        objExcel.Application.Workbooks(1).Activate()
'        objExcel.Application.Workbooks(1).Worksheets(1).Activate()

'        ' Add the column headers

'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 1).Value = "Row# Imported"

'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 1).Value = "Row# Imported"
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 3).Value = gsLOG_PROPOSAL_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 4).Value = gsLOG_REV_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 5).Value = gsLOG_ProgYr_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 6).Value = gsLOG_ItemNumber_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 7).Value = gsLOG_UPC_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 8).Value = glLog_ORIG_PROP_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 9).Value = glLog_ORIG_REV_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 10).Value = glLog_ORIG_ITEM_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 11).Value = glLog_ORIG_UPC_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 12).Value = gsLOG_LONGDESC_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 13).Value = gsLOG_VendorItemNumber_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 2).Value = gsLOG_FunctionCode_ColName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 14).Value = glLOG_MESSAGE_ColName

'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 15).Value = "Validate Start: " & dtValidateStart & " - End: " & dtValidateEnd

'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 16).Value = "Import File: " & simportFileName
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 17).Value = "Last Modified Date of ImportFile before Import: " & sLastModified
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 18).Value = "Import Start: " & dtImportStart

'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 20).Value = "Import File Size: " & sFileSize & " bytes"
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 21).Value = "# New Proposals to be created: " & lNewProposals
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 22).Value = "# New Revisions to be created: " & lNewRevisions
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 23).Value = "# Proposals to be changed: " & lChangedProposals
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 24).Value = "# Referenced '100' Accounts: " & lReferenced100Accts
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 25).Value = "# Rolled-Back '999' Proposals: " & lRollbackProposals
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 26).Value = "RevisedUserID: " & gsUserID

'        'add 3 new cols with info
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 27).Value = "Range of Data to Import(Rows): " & "1 to " & llastrow
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 28).Value = "Range of Data to Import(Cols): " & "1 to " & lLastColumn & "-'" & sLastColHeading & "'"
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 29).Value = "Photos(Col: " & lPhotoCOLPos & "), Refreshed/Added: " & UCase(bRefreshImportPhotos)


'        If lAlternatePhotoCOLPos <> 0 Then
'            objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 30).Value = "Alternate Photo: " & "YES"
'        Else
'            objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 30).Value = "Alternate Photo: " & "NO"
'        End If
'        Application.DoEvents()

'        With objExcel.Application.Workbooks(1).Worksheets(1)
'            .Rows("1:1").RowHeight = 45
'            .Rows("1:1").Select()
'        End With

'        objExcel.Application.Selection.HorizontalAlignment = xlGeneral
'        objExcel.Application.Selection.VerticalAlignment = xlTop
'        objExcel.Application.Selection.WrapText = True
'        objExcel.Application.Selection.Orientation = 0
'        objExcel.Application.Selection.AddIndent = False
'        objExcel.Application.Selection.IndentLevel = 0
'        objExcel.Application.Selection.ShrinkToFit = False
'        objExcel.Application.Selection.ReadingOrder = xlContext
'        objExcel.Application.Selection.MergeCells = False

'        ' Save the file
'        If Version >= 12.0# Then
'            objExcel.Application.Workbooks(1).SaveAs(sfilename, , , , , 0) 'to prevent .xlk backup file being created
'        Else
'            objExcel.Application.Workbooks(1).SaveAs(sfilename, xlNormal)
'        End If

'        bCreateImportLogFile = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bCreateImportLogFile")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCreateImportLogFile , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine


'Dim Public Function bValidateLogFileColumns(objExcel As Excel.Application
'Dim  _ As Object 

'                ByVal bImportLogFile As Boolean, _
'                ByRef sErrorMsg As String , ByRef lNbrErrors As Long, _
'                bFROMDelete As Boolean, Optional bFROMSpeedQuote As Boolean, _
'                Optional lCustomerNbrColPos As Long, _
'                Optional lItemNumberCOLPos As Long) As Boolean
'    End Function

'        Friend Overridable Function bValidateLogFileColumns(ByVal objExcel As Excel.Application, ByVal ByVal bImportLogFile As Boolean, ByVal ByRef sErrorMsg As String, ByRef lNbrErrors As Long, ByVal bFROMDelete As Boolean, ByVal  Optional bFROMSpeedQuote As Boolean, ByVal Optional lCustomerNbrColPos As Long, ByVal Optional lItemNumberCOLPos As Long) As Object

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sCellValue As String
'        Dim ProposalCOl As Long
'        Dim RevCOL As Long

'        bValidateLogFileColumns = False
'        sErrorMsg = ""
'        lNbrErrors = 0

'        ' The first three column positions are hard-coded for every XLS Log file
'        '2012/01/17 - deleting from spreadsheet also uses these column positions
'        '2013/01/02 - More Useful would probably be to find the column position, so that Items can be deleted from any file.
'        '2012/01/17 - deleting from spreadsheet also uses these column positions
'        If bImportLogFile = True Then
'            ProposalCOl = 3
'            RevCOL = 4
'        Else
'            ProposalCOl = 2
'            RevCOL = 3
'        End If
'        sCellValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, ProposalCOl))
'        If sCellValue <> gsLOG_PROPOSAL_ColName Then
'            sErrorMsg = sErrorMsg & gsLOG_PROPOSAL_ColName & " not found in column: " & ProposalCOl & vbCrLf
'            lNbrErrors = lNbrErrors + 1
'        End If

'        sCellValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, RevCOL))
'        If sCellValue <> gsLOG_REV_ColName Then
'            sErrorMsg = sErrorMsg & gsLOG_REV_ColName & " not found in column: " & RevCOL & vbCrLf
'            lNbrErrors = lNbrErrors + 1
'        End If

'        If bFROMSpeedQuote = True Or bFROMDelete = True Then
'            sCellValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lItemNumberCOLPos))

'        Else
'            sCellValue = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, 6))
'        End If
'        If sCellValue <> gsLOG_ItemNumber_ColName Then
'            sErrorMsg = sErrorMsg & gsLOG_ItemNumber_ColName & " not found in column 6" & vbCrLf
'        End If

'        bValidateLogFileColumns = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateLogFileColumns")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidateLogFileColumns , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'        '2012/01/17 - added bFromDelete below
'        '2013/01/02- hn- added bImportLogFile
'        '2013/05/07 -HN- added lProposalCOLPos,lRevCOLPos as Optional parameters to be passed in. SpeedQuote also calls this function...
'Dim Public Function bValidateLogFileProposalReferences(objExcel As Excel.Application
'        Dim frmThis As Form
'Dim  _ As Object 

'                ByVal bImportLogFile As Boolean, ByRef sErrorMsg As String , ByRef lNbrErrors As Long, _
'        End Function

'        Friend Overridable Function bValidateLogFileProposalReferences(ByVal objExcel As Excel.Application, ByVal  frmThis As Form, ByVal ByVal bImportLogFile As Boolean, ByRef sErrorMsg As String, ByRef lNbrErrors As Long, ByVal ByRef lRowsOnSheet As Long, ByVal bFromSQ As Boolean, ByVal Optional lCustomerNbrColPos As Long, ByVal Optional lItemNumberCOLPos As Long, ByVal Optional lProposalCOLPos As Long, ByVal Optional lRevCOLPos As Long) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim rsProposalRef As ADODB.Recordset
'        Dim SQL As String
'        Dim sProposalNumber As String
'        Dim sRev As String
'        Dim sItemNumber As Object

'        Dim lProposal As Long
'        Dim lRev As Long

'        Dim lcustomernumber As Long
'        Dim sCustomerNumber As String

'        Dim lCounter As Long
'        Dim lRowCol2ColorIndex As Long
'        Dim lRowCol3ColorIndex As Long
'        Dim ProposalCOl As Long
'        Dim RevCOL As Long

'        bValidateLogFileProposalReferences = False

'        rsProposalRef = New ADODB.Recordset

'        ' The first row contains field names, .. start on row 2
'        For lCounter = glLOG_DATA_START_ROW To glMAX_Rows
'            Call bUpdateStatusMessage(frmThis, "Validating Proposal Reference for Row " & CStr(lCounter) & "...")
'            If bFromSQ = True Or bFROMDelete = True Then
'                If bImportLogFile = True Then
'                    ProposalCOl = lProposalCOLPos   '3          '2013/05/02 -HN
'                    RevCOL = lRevCOLPos '4
'                Else
'                    ProposalCOl = lProposalCOLPos '2             '2013/05/07 -HN-
'                    RevCOL = lRevCOLPos '3
'                End If
'                sProposalNumber = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, ProposalCOl))
'                lRowCol2ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, ProposalCOl).Interior.ColorIndex
'                sRev = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, RevCOL))
'                lRowCol3ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, RevCOL).Interior.ColorIndex

'            Else
'                If bImportLogFile = True Then
'                    ProposalCOl = 3
'                    RevCOL = lRevCOLPos '4
'                Else
'                    ProposalCOl = 2
'                    RevCOL = lRevCOLPos '3
'                End If
'                sProposalNumber = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, ProposalCOl))
'                lRowCol2ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, ProposalCOl).Interior.ColorIndex
'                sRev = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, RevCOL))
'                lRowCol3ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, RevCOL).Interior.ColorIndex
'                '        lColorIndex = objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glLOG_PROPOSAL_ColPos).Interior.ColorIndex
'            End If

'            If Len(sProposalNumber) = 0 _
'                And Len(sRev) = 0 _
'                And lRowCol2ColorIndex = xlNone And lRowCol3ColorIndex = xlNone Then   '2011/12/21
'                '            And lColorIndex = xlNone Then
'                lRowsOnSheet = lCounter - glLOG_DATA_START_ROW + 1
'                Exit For
'            End If

'            If Not IsNumeric(sProposalNumber) Or Not IsNumeric(sRev) Then
'                '            If lColorIndex <> xlNone Then
'                If lRowCol2ColorIndex <> xlNone Then
'                    If sProposalNumber = "" And sRev = "" Then

'                    Else
'                        sErrorMsg = sErrorMsg & "ROW " & lCounter & ": Both Proposal Number:[" & sProposalNumber & "] and Rev:[" & _
'                        sRev & "] must be numeric, and MUST be in Column 2 & 3." & vbCrLf
'                        lNbrErrors = lNbrErrors + 1
'                    End If
'                Else
'                    '                Exit For
'                End If

'            End If

'            If bFromSQ = False And bFROMDelete = True Then
'                If lCustomerNbrColPos > 0 Then
'                    sCustomerNumber = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, lCustomerNbrColPos))
'                    If IsNumeric(sCustomerNumber) = True Then
'                        lcustomernumber = sCustomerNumber
'                    Else
'                        lcustomernumber = 0
'                    End If
'                End If
'                If lItemNumberCOLPos > 0 Then
'                    sItemNumber = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, lItemNumberCOLPos))
'                End If
'            Else  'this is for SpeedQuote .....
'                sItemNumber = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glSQItemNumCOLUMNPOS))
'            End If

'            ' msMKTGBASIC, msPRODUCTMGR, msAdmin, PowerUser can only delete for account = 998, 999  -----
'            Select Case lcustomernumber
'                Case gs998PD_ACCOUNT, gs999PD_ACCOUNT
'                    Select Case msUserGroup
'                        '2014/10/10 RAS Added msSupvrUser
'                        Case msMKTGBASIC, msPRODUCTMGR, msADMIN, msPowerUser, msSUPVRUSER
'                            'ok to delete
'                        Case Else
'                            sErrorMsg = sErrorMsg & "ROW " & lCounter & ": You do NOT have " & _
'                                        "Permission to DELETE when 'CustomerNumber' = " & lcustomernumber & " !" & vbCrLf
'                            lNbrErrors = lNbrErrors + 1
'                    End Select
'                Case Else
'                    'ok to delete
'            End Select

'            If Not IsNumeric(sProposalNumber) Or Not IsNumeric(sRev) Then
'                '            sErrorMsg = sErrorMsg & "ROW " & lCounter & ": Both Proposal Number and Rev " & _
'                '                    "must be numeric" & vbCrLf
'                '            lNbrErrors = lNbrErrors + 1
'            Else
'                lRev = sRev
'                lProposal = sProposalNumber
'                If IsBlank(sItemNumber) Then
'                    SQL = "SELECT ProposalNumber FROM Item WHERE ProposalNumber = " & lProposal & _
'                            " AND Rev = " & lRev & " AND ItemNumber IS NULL "
'                Else
'                    SQL = "SELECT ProposalNumber FROM Item WHERE ProposalNumber = " & lProposal & _
'                                " AND Rev = " & lRev & " AND ItemNumber = " & sAddQuotes(sItemNumber)
'                End If
'Dim             rsProposalRef.Open SQL As Object 
'                Dim SSDataConn As Object
'                Dim adOpenStatic As Object
'                Dim adLockReadOnly As Object

'                If rsProposalRef.EOF Then
'                    sErrorMsg = sErrorMsg & "ROW " & lCounter & ": ProposalNumber " & sProposalNumber & _
'                            ", Rev " & sRev & ", ItemNumber [" & sItemNumber & "] not in the database" & vbCrLf
'                    lNbrErrors = lNbrErrors + 1
'                End If

'                rsProposalRef.Close()
'            End If

'        Next

'        If lRowsOnSheet = 1 Then
'            sErrorMsg = sErrorMsg & "No data rows found on the spreadsheet." & vbCrLf
'            lNbrErrors = lNbrErrors + 1
'        End If
'        bValidateLogFileProposalReferences = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        rsProposalRef.Close()
'        rsProposalRef = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateLogFileProposalReferences")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidateLogFileProposalReferences , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine


'Dim Public Function bDeleteData(frmThis As Form
'        Dim sDataArray() As String
'            ByRef bImportLogFile As Boolean, _
'            ByVal lRowsOnLogSheet As Long, _
'            ByVal sDeletionLogFileName As String , _
'            ByVal bWriteToLog As Boolean, _
'            ByRef sReturnMsg As String , ByRef lReturnMsgRow As Long, _
'            ByVal sCategory As String , ByVal sSeason As String , ByVal sProgram As String , ByVal bFromDeleteForm As Boolean, _
'            ByVal sDeleteFileName As String , _
'        'End Function

'        'Friend Overridable Function bDeleteData(ByVal frmThis As Form) As  String, _

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lRowCounter As Long
'        Dim sLogMsg As String
'        Dim lItemNumCOLPos As Long
'        Dim rsDeleteItemSpecs As ADODB.Recordset
'        Dim rsDeleteItem As ADODB.Recordset
'        Dim rsDeleteAssortments As ADODB.Recordset

'        Dim rsDeleteItemMaterial As ADODB.Recordset
'        Dim sProposalNumber As String
'        Dim sRev As String
'        Dim sItemNumber As String

'        Dim sProgramYear As String
'        Dim sSQL As String
'        Dim sParentProposalNumber As String
'        Dim sDelCategory As String
'        Dim sDelSeason As String
'        Dim sDelProgram As String

'        Dim ProposalCOl As Long
'        Dim RevCOL As Long

'        Screen.MousePointer = 11 ' Hourglass
'        bDeleteData = False
'        lItemNumCOLPos = 6    'The log array only has 4 cols, so  stays this way for now
'        rsDeleteItemMaterial = New ADODB.Recordset
'        rsDeleteAssortments = New ADODB.Recordset
'        rsDeleteItemSpecs = New ADODB.Recordset
'        rsDeleteItem = New ADODB.Recordset

'        If bImportLogFile = True Then
'            ProposalCOl = lProposalCOLPos '3  '2013/05/07 -HN
'            RevCOL = lRevCOLPos '4
'        Else
'            ProposalCOl = 2                   '2013/05/07 -HN- obtained from DataArray behind ProposalList
'            RevCOL = 3
'        End If

'        For lRowCounter = glLOG_DATA_START_ROW To lRowsOnLogSheet
'            Call bUpdateStatusMessage(frmThis, "Deleting Row: " & lRowCounter & ".....")
'            '        If sDataArray(lRowCounter, glLOG_PROPOSAL_ColPos) = "" And sDataArray(lRowCounter, glLOG_REV_ColPos) = "" Then
'            If sDataArray(lRowCounter, 2) = "" And sDataArray(lRowCounter, 3) = "" Then

'            Else

'                sProposalNumber = sDataArray(lRowCounter, ProposalCOl)
'                sRev = sDataArray(lRowCounter, RevCOL)
'                '        sItemNumber = sDataArray(lRowCounter, 5)
'                If bImportLogFile = True Then
'                    sItemNumber = sDataArray(lRowCounter, RevCOL + 2)
'                Else
'                    If frmThis.Name = "frmProposalList" Then
'                        sItemNumber = sDataArray(lRowCounter, 5)
'                    Else
'                        sItemNumber = sDataArray(lRowCounter, RevCOL + 3)
'                    End If
'                End If

'                '        sProgramYear = sDataArray(lRowCounter, glLog_ProgYr_ColPos)
'                '        sProgramYear = sDataArray(lRowCounter, 4)
'                If frmThis.Name = "frmProposalList" Then
'                    sProgramYear = sDataArray(lRowCounter, 4)
'                Else
'                    sProgramYear = sDataArray(lRowCounter, RevCOL + 1)
'                End If

'                'cannot delete when Item is on an Order
'                Call bUpdateStatusMessage(frmThis, "Checking row:" & lRowCounter & ", for Purchase Order..", True, vbYellow)
'                If bItemInOrder(sProposalNumber, sRev, sDataArray(lRowCounter, 4), sReturnMsg) = True Then  '2013/05/04 -HN- removed 1st parameter
'                    sReturnMsg = sReturnMsg & vbCrLf & "     DELETION DENIED!!"
'                    lReturnMsgRow = lRowCounter
'                    '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If

'                Call bUpdateStatusMessage(frmThis, "Checking row:" & lRowCounter & ", if Assortment Item..")
'                If Not IsBlank(sItemNumber) Then    '2012/09/17
'                    If bAssortmentItemExists(sProposalNumber, sRev, sItemNumber, sProgramYear, sParentProposalNumber) = True Then                     '2010/11/23
'                        sReturnMsg = sReturnMsg & vbCrLf & sParentProposalNumber & vbCrLf & "     DELETION DENIED!!"
'                        lReturnMsgRow = lRowCounter
'                        '                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                    End If
'                End If

'                '        If bFromDeleteForm = True Then
'                '         sSQL = "SELECT CategoryCode AS CATEGORY, SeasonCode as SEASON, ProgramNumber as PROGRAM " & _
'                '           "FROM Item WHERE Item.ProposalNumber = "
'                '
'                '        End If
'                If bWriteToLog Then
'                    '            2012/01/17 - because import cols were changed had to change these
'                    sLogMsg = "Deleting row:" & lRowCounter & ", for ProposalNumber " & _
'                                sDataArray(lRowCounter, 2) & ", " & _
'                                "Rev " & sDataArray(lRowCounter, 3) & ", " & _
'                                "ItemNumber:" & sDataArray(lRowCounter, 5)      '....depending on where we come from!
'                    Call bSaveDataToFile(sDeletionLogFileName, sLogMsg, True)
'                End If

'                ' Because each record is bridged between 4 tables,
'                ' only update after the record passes bDeleteRow for each table
'                ' NOTE: Delete in reverse order of save due to referential integrity

'                Call bUpdateStatusMessage(frmThis, "Deleting data for row:" & lRowCounter & ", ..", True, vbGreen)

'                If bDeleteRow(sDataArray(), ProposalCOl, RevCOL, lRowCounter, "ItemMaterial", _
'                            rsDeleteItemMaterial, sDelCategory, sDelSeason, sDelProgram) = False Then
'                    sReturnMsg = "Could not delete the ItemMaterial data"
'                    lReturnMsgRow = lRowCounter
'                    '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If

'                If bDeleteRow(sDataArray(), ProposalCOl, RevCOL, lRowCounter, gsItem_Assortments_Table, _
'                            rsDeleteAssortments, sDelCategory, sDelSeason, sDelProgram) = False Then
'                    sReturnMsg = "Could not delete the Item_Assortments data"
'                    lReturnMsgRow = lRowCounter
'                    '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If
'                '        '2013/01/03 - this part below only for Assort Development, where Item_Assortment table is now replaced by ComponentItem table
'                '        If bDeleteRow(sDataArray(), ProposalCOl, RevCOL, lRowCounter, "ComponentItem", _
'                '                        rsDeleteComponentItem, sDelCategory, sDelSeason, sDelProgram) = False Then
'                '            sReturnMsg = "Could not delete the Assortment Component data"
'                '            lReturnMsgRow = lRowCounter
'                ''            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                '        End If

'                If bDeleteRow(sDataArray(), ProposalCOl, RevCOL, lRowCounter, gsItemSpecs_Table, _
'                                rsDeleteItemSpecs, sDelCategory, sDelSeason, sDelProgram) = False Then
'                    sReturnMsg = "Could not delete the ItemSpecs data"
'                    lReturnMsgRow = lRowCounter
'                    '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If

'                If bDeleteRow(sDataArray(), ProposalCOl, RevCOL, lRowCounter, gsItem_Table, _
'                                    rsDeleteItem, sDelCategory, sDelSeason, sDelProgram) = False Then
'                    sReturnMsg = "Could not delete the Item table data"
'                    lReturnMsgRow = lRowCounter
'                    '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                Else
'                    sCategory = sDelCategory 'need these fields for ItemFieldHistory table
'                    sSeason = sDelSeason
'                    sProgram = sDelProgram
'                End If

'                ' Update the four tables and close the recordsets
'                rsDeleteItemMaterial.UpdateBatch()
'                If bPauseTime(frmThis, 4) = False Then
'                End If

'                rsDeleteAssortments.UpdateBatch()
'                If bPauseTime(frmThis, 4) = False Then
'                    Application.DoEvents()
'                End If

'                rsDeleteItemSpecs.UpdateBatch()
'                If bPauseTime(frmThis, 4) = False Then
'                    Application.DoEvents()
'                End If

'                rsDeleteItem.UpdateBatch()
'                If bPauseTime(frmThis, 4) = False Then
'                    Application.DoEvents()
'                End If

'                'delete ItemInspection record if no PO yet?     '2012/06/05
'                '        DeleteInspectionByProposalRev sProposalNumber, sRev        '2012/06/07 not yet wait until later decision

'                '        Write Info to ItemFieldHistory table
'                Call bUpdateStatusMessage(frmThis, "Updating ItemFieldHistory table for deleted Item..", True, vbGreen)
'                Call bWriteDeletedItemFieldHistory(sProposalNumber, sRev, sItemNumber, sCategory, sSeason, sProgram, bFromDeleteForm, sDeleteFileName)

'                rsDeleteAssortments.Close()
'                rsDeleteItem.Close()
'                rsDeleteItemSpecs.Close()
'                rsDeleteItemMaterial.Close()

'                If bInsertDeleteDateIntoUPC(Left(sItemNumber, 6)) = True Then           'hn 10/09/2007
'                    sReturnMsg = "Deletion successful, CoreItemNumber: " & Microsoft.VisualBasic.Left(sItemNumber, 6) & " available again!"
'                Else
'                    sReturnMsg = "Deletion successful, CoreItemNumber: " & Microsoft.VisualBasic.Left(sItemNumber, 6) & " in use by other Proposals!"
'                End If
'            End If

'        Next lRowCounter

'        bDeleteData = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rsDeleteAssortments.State <> 0 Then
'            rsDeleteAssortments.CancelBatch()
'            rsDeleteAssortments.Close()
'        End If
'        rsDeleteAssortments = Nothing

'        If rsDeleteItem.State <> 0 Then
'            rsDeleteItem.CancelBatch()
'            rsDeleteItem.Close()
'        End If
'        rsDeleteItem = Nothing

'        If rsDeleteItemSpecs.State <> 0 Then
'            rsDeleteItemSpecs.CancelBatch()
'            rsDeleteItemSpecs.Close()
'        End If
'        rsDeleteItemSpecs = Nothing
'        Screen.MousePointer = 0
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bDeleteData")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bDeleteData , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function
'    Private Function bDeleteRow(sDataArray As String, ByVal ProposalCOl As Long, ByVal RevCOL As Long, ByVal lRow As Long, _
'                    ByVal sTableName As String, _
'                    rsDeleteTable As ADODB.Recordset, _
'                    ByRef sDelCategory As String, ByRef sDelSeason As String, ByRef sDelProgram As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        '    Dim sSQL As String 
'        '    bDeleteRow = False
'        '    If sTableName = gsItem_Table Then
'        '        sSQL = "SELECT ProposalNumber, CategoryCode, SeasonCode, ProgramNumber FROM "
'        'End Function

'        '    Private Function bDeleteRow(ByVal sDataArray( As Object) As [String]
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        bDeleteRow = False
'        If sTableName = gsItem_Table Then
'            sSQL = "SELECT ProposalNumber, CategoryCode, SeasonCode, ProgramNumber FROM "
'        Else
'            sSQL = "SELECT * FROM "

'        End If
'        '2012/01/17 - set proposal col = 2 because of changes to Import log file where constants were used
'        '2013/01/02 - Proposal COl and Rev col are different value for regular log files, and Import log fi;e becuase of change made above.
'        sSQL = sSQL & sTableName & " " & "WHERE ProposalNumber = " & sDataArray(lRow, ProposalCOl) & " " & _
'                     "AND Rev = " & sDataArray(lRow, RevCOL)

'        rsDeleteTable.Open(sSQL, SSDataConn, adOpenKeyset, adLockBatchOptimistic)
'        If Not rsDeleteTable.EOF Then
'            If sTableName = gsItem_Table Then
'                sDelCategory = rsDeleteTable!CategoryCode
'                sDelSeason = rsDeleteTable!SeasonCode
'                sDelProgram = rsDeleteTable!ProgramNumber
'            End If

'            Do Until rsDeleteTable.EOF
'                If Len(rsDeleteTable!ProposalNumber) <> 0 Then
'                    rsDeleteTable.Delete()
'                    rsDeleteTable.MoveNext()
'                End If
'            Loop

'        End If
'        bDeleteRow = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        '2014/03/07 RAS We need to have these recordsets stay open until after the .UpdateBatch in bDeleteData
'        ' If rsDeleteTable.State <> 0 Then rsDeleteTable.Close  '2014/01/09 RAS adding statement based on SJM, she just closed I am checking first to see if it is open

'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bDeleteRow")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bDeleteRow , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function
'    Public Function bWriteDeletedItemFieldHistory(ByVal sProposalNumber As String _
'    , ByVal sRev As String _
'    , ByVal sItemNumber As String, _
'                    ByVal sCategory As String, ByVal sSeason As String, ByVal sProgram As String, ByVal bFromDeleteForm As Boolean, _
'                    ByVal sDeleteFileName As String) As Boolean
'        'End Function

'        '    Friend Overridable Function bWriteDeletedItemFieldHistory(ByVal sProposalNumber As String, ByVal sRev As String, ByVal sItemNumber As String, ByVal ByVal sCategory As String, ByVal sSeason As String, ByVal sProgram As String, ByVal bFromDeleteForm As Boolean, ByVal ByVal sDeleteFileName As String) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim sCatSeasonProg As String
'        Dim sFromWhereDeleted As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset

'        If bFromDeleteForm = True Then
'            sFromWhereDeleted = "All Fields Deleted(from Spreadsheet): " & sDeleteFileName
'        Else
'            sFromWhereDeleted = "All Fields Deleted(from Proposal List)"
'        End If
'        sSQL = "SELECT * FROM ItemFieldHistory WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sRev
'        sCatSeasonProg = "Cat/Season/Program: " & sCategory & "/" & sSeason & "/" & sProgram
'        rs.Open(sSQL, SSDataConn, adOpenDynamic, adLockPessimistic)

'        '2010/02/11 - added 0 as 1st parameter
'        If bSAVEItemFieldsChanged(0, rs, sProposalNumber, sRev, sRev, sItemNumber, "D", "D", "Proposal/Rev Deleted", _
'                                 sCatSeasonProg, sFromWhereDeleted, "", gsUserID) = False Then
'            '        GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'        End If

'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rs.State <> 0 Then rs.Close() '2014/01/09 RAS adding statement based on SJM, she just closed I am checking first to see if it is open

'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bWriteDeletedItemFieldHistory")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bWriteDeletedItemFieldHistory , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function
'    Public Function bInsertDeleteDateIntoUPC(ByVal sCoreItemNumber As String) As Boolean

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        'this sets the ItemDeleteDate on the UPC table only IF no other Proposals exists for this CoreItemNumber
'        ' to eventually free up these numbers again
'        Dim SQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        bInsertDeleteDateIntoUPC = False
'        SQL = "SELECT DISTINCT ProposalNumber, Rev FROM Item WHERE CoreItemNumber = '" & sCoreItemNumber & "'"
'        rs.Open SQL
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object

'        If rs.EOF Then  ' ie there are no other proposals with this core item number
'            rs.Close()
'            rs = New ADODB.Recordset
'            SQL = "SELECT * FROM UPC WHERE CoreItemNumber = '" & sCoreItemNumber & "' AND ItemDeleteDate IS NULL"
'            rs.Open SQL
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockBatchOptimistic As Object

'            If Not rs.EOF Then
'                rs!ItemDeleteDate = Now.ToShortDateString()
'                rs.UpdateBatch()
'                bInsertDeleteDateIntoUPC = True
'            End If
'            rs.Close()

'        End If
'        'Set RS = Nothing  '2014/01/09 RAS Setting this to nothing in the Exit routine.
'ExitRoutine:
'        If rs.State <> 0 Then   '2014/01/09 RAS adding statement based on SJM, she just closed I am checking first to see if it is open
'            rs.Close()
'            rs = Nothing
'        End If
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bInsertDeleteDateIntoUPC")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bInsertDeleteDataIntoUPC , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Friend Overridable Function bGetAssortmentSelect(ByVal lMaxNbrItems As Long, ByVal lType As Long) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        ' Builds an addition to the select clause of the temp query
'        ' for the maximum number of assortment items found in the result set
'        Dim sSelect As String
'        Dim sTemp As String

'        Dim lCounter As Long

'        bGetAssortmentSelect = False
'        If gbCancelExport = True Then Exit Function
'        For lCounter = 1 To lMaxNbrItems
'            '        If lCounter > 24 Then
'            '            'then too many columns for Access Query'
'            '            Exit For
'            '        Else
'            If lCounter < 10 Then
'                sTemp = "0" & CStr(lCounter)
'            Else
'                sTemp = CStr(lCounter)
'            End If

'            Select Case lType           'Removing [ & ] brackets prevents SQL statement from becoming too long & erroring ...
'                Case glITEM_ONLY
'                    sSelect = sSelect & gsItem_Assortments_Table & ".Item_" & sTemp & ", "
'                Case glITEM_AND_QTY
'                    sSelect = sSelect & gsItem_Assortments_Table & ".Item_" & sTemp & ", " & gsItem_Assortments_Table & ".Qty_" & sTemp & ", "
'            End Select
'            '        End If
'        Next

'        sAssortmentSelect = Microsoft.VisualBasic.Left(sSelect, Len(sSelect) - 2)
'        bGetAssortmentSelect = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bGetAssortmentSelect")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bGetAssortmentSelect , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'        Private Function bMarkDB4Column(ByVal objExcel As Excel.Application, ByVal lColumn As Long, ByVal ByVal bIsDB4Column As Boolean, ByVal bFieldRequired As Boolean, ByVal bXUnderscore As Boolean, ByVal ByVal sValidation As String, ByVal ByVal bImport As Boolean) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lLineStyleBottom As Long
'        Dim lLineStyleRight As Long

'        Dim bMaterialCOL As Boolean
'        bMarkDB4Column = False
'        'Materialx columns dont exist in Itemmaterial table,
'        'only on Spreadsheets, where we need a double border around heading
'        If Microsoft.VisualBasic.Left(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn), 8) = "Material" Then
'            bMaterialCOL = True
'        Else
'            bMaterialCOL = False
'        End If

'        If (bIsDB4Column And bImport = True) Or bMaterialCOL = True Then
'            objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Borders(xlEdgeBottom).LineStyle = xlDouble
'            objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Borders(xlEdgeLeft).LineStyle = xlDouble
'            objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Borders(xlEdgeRight).LineStyle = xlDouble

'            If Microsoft.VisualBasic.Left(sValidation, 7) = gsVARCHAR Then        'this in Field table: first 7 characters of ImportValidation
'                '2010/11/18 - was causing ##### values in these columns
'                '@ means setting format to text
'                Select Case objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn)
'                    Case "DevelopComments", "SeasonalComments", "SeasonalCommentsToVendor", "VendorComments", "SalesComments", "SpecComments"   '2011/12/19
'                        If objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat <> "General" Or _
'                          IsBlank(objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat) Then  '2011/10/04 -format was null so fell through this!
'                            objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat = "General"
'                            Application.DoEvents()
'                        End If

'                    Case "CertifiedPrinterID"       '2011/10/28 - this field must be formatted as text, to avoid losing leading zeros in the spreadsheet
'                        If objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat <> "@" Or _
'                            IsNull(objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat) Then
'                            objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat = "@"  'Text
'                        End If

'                    Case Else
'                        If objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat <> "General" And _
'                            objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat <> "@" Then  '2010/11/22 - dont unset if it's general
'                            objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat = "@"  'Text
'                        End If
'                End Select

'            End If
'        Else
'            If Microsoft.VisualBasic.Left(sValidation, 7) = gsFORMATC Then
'                '           objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat = "\ 000\ 00000000\ 0"
'                ' THE ABOVE DOES NOT WORK WITH TEXT FIELDS, because it right justifies it
'                objExcel.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat = "@"
'            End If
'            lLineStyleBottom = objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Borders(xlEdgeBottom).LineStyle
'            lLineStyleRight = objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Borders(xlEdgeRight).LineStyle

'            If lLineStyleBottom = xlDouble Then
'                objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Borders(xlEdgeBottom).LineStyle = xlNone
'            End If

'            If lLineStyleRight = xlDouble Then
'                objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Borders(xlEdgeRight).LineStyle = xlNone
'            End If
'        End If
'        If bFieldRequired = True Then
'            'required field headings backcolor =ColorTranslator.FromOle( lite green)
'            objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Interior.ColorIndex = 35
'        End If
'        If bXUnderscore = True Then
'            'X_ headings backcolor =ColorTranslator.FromOle( lite yellow)
'            objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Interior.ColorIndex = 36
'        End If
'        bMarkDB4Column = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        If Err.Number = 13 Then '11/02/2009 - hn - to take care of #Value heading in spreadsheet

'            If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'                smessage = "In bMarkDB4Column , Err Number " & Err.Number & "Error Description: " & Err.Description
'                If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'                End If
'            End If
'            Resume Next
'        Else
'            MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bMarkDB4Column")
'            If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'                smessage = "In bMarkDB4Column , Err Number " & Err.Number & "Error Description: " & Err.Description
'                If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'                End If
'            End If
'        End If
'        Resume ExitRoutine


'    Public Function sGetCellValue(vSpreadsheetCell As Object) As String
'        On Error GoTo ErrorHandler
'        If IsError(vSpreadsheetCell) Then
'            sGetCellValue = gsDENOTE_CELL_ERROR
'        Else
'            sGetCellValue = sConvertVariantToString(vSpreadsheetCell)
'        End If
'ExitRoutine:
'        Exit Function
'ErrorHandler:
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In sGetCellValue , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-sGetCellValue")
'        Resume ExitRoutine
'    End Function

'    Public Function bUpdateSpreadsheetCell(ByVal sDEFAULTVAL As String, ByVal sNewValue As String, ByVal sOldValue As String, _
'                          objExcel As Excel.Application, ByVal lRow As Long, ByVal lColumn As Long, ByVal bProhibited As Boolean) As Boolean
'        On Error GoTo ErrorHandler
'        Dim lCellPattern As Long
'        Dim sCellFormula As String '01/10/2008
'        Dim sHeaderColName As String

'        bUpdateSpreadsheetCell = False
'        sHeaderColName = objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Value

'        '    Select Case objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Value             '2011/08/17
'        Select Case sHeaderColName             '2011/08/17
'            Case "DevelopComments", "SeasonalComments", "SeasonalCommentsToVendor", "VendorComments", "SalesComments", "SpecComments"   '2011/12/19
'                If Left(sNewValue, 1) = "^" Then
'                    sNewValue = Replace(sNewValue, "^", "") & vbCrLf & sOldValue
'                End If
'            Case Else
'        End Select

'        If sNewValue = "" And sDEFAULTVAL <> "" Then              '03/20/2009 - hn
'            sNewValue = sDEFAULTVAL
'        End If
'        If sNewValue <> sOldValue Then
'            '        If sNewValue = "" Then
'            '            sNewValue = sDEFAULTVAL
'            '        End If
'            'when updating the value of a cell with a new value then the formula of that cell is deleted '01/10/2008 yikes!
'            If objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).HasFormula = True Then
'                'if cell has no formula it sets the formula to the value
'                sCellFormula = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).Formula
'            Else
'                sCellFormula = "NoFormula"
'            End If

'            objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).Value = sNewValue

'            If sOldValue <> gsENDofSpreadsheet Then
'                If glSHADING <> NoShadingX Then                     '2012/01/04
'                    objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).Interior.Pattern = xlGray8
'                Else
'                    objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).Interior.Pattern = xlNone
'                End If
'            End If
'            If sCellFormula <> "NoFormula" Then     '01/10/2008
'                objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).Formula = sCellFormula
'            End If
'        Else
'            lCellPattern = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).Interior.Pattern
'            If lCellPattern = xlGray8 Then
'                objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).Interior.Pattern = xlNone
'            End If

'        End If
'        'new 10/15/2007
'        '    If objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Value = "FactoryNumber" Or _
'        '        objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Value = "X_FactoryName" Then
'        If sHeaderColName = "FactoryNumber" Or _
'            sHeaderColName = "X_FactoryName" Then
'            If bProhibited = True Then
'                objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, lColumn).Font.ColorIndex = 3
'            End If
'        End If
'        ' this now done in function bMarkDB4Column
'        '    2010/11/18 - these columns are defined as Text on the Item table, for some reason the NumberFormat of these are changed from General to Text, filling some of the excel cells with #####' s
'        '                       even redefining the fields as varchar on the table, still got the same results  ?
'        '    Select Case objEXCEL.Application.Workbooks(1).Worksheets(1).Cells(1, lColumn).Value
'        '        Case "DevelopComments", "SeasonalComments", "SeasonalCommentsToVendor", "VendorComments"
'        '            objEXCEL.Application.Workbooks(1).Worksheets(1).Columns(lColumn).NumberFormat = "General"
'        '            DoEvents
'        '        Case Else
'        '    End Select
'        bUpdateSpreadsheetCell = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:
'        'Resume Next    ''' testing only
'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bUpdateSpreadsheetCell")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bUpdateSpreadsheetCell , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        ' Resume Next '' testing
'        Resume ExitRoutine
'    End Function

'    Public Function bCalcCubes(ByVal lRow As Long, ByVal dtCOLPos As typSpecialCOLPos) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        bCalcCubes = False
'        Dim dblUnitLength As Double
'        Dim dblUnitWidth As Double
'        Dim dblUnitHeight As Double

'        Dim dblCube As Double

'        ' Calculate MasterPackCube
'        If IsNull(sItemArray(lRow, dtCOLPos.lMasterPackLength)) Or _
'                sItemArray(lRow, dtCOLPos.lMasterPackLength) = "" Then
'            dblUnitLength = 0
'        Else
'            dblUnitLength = sItemArray(lRow, dtCOLPos.lMasterPackLength)
'        End If
'        If IsNull(sItemArray(lRow, dtCOLPos.lMasterPackWidth)) Or _
'                sItemArray(lRow, dtCOLPos.lMasterPackWidth) = "" Then
'            dblUnitWidth = 0
'        Else
'            dblUnitWidth = sItemArray(lRow, dtCOLPos.lMasterPackWidth)
'        End If
'        If IsNull(sItemArray(lRow, dtCOLPos.lMasterPackHeight)) Or _
'                sItemArray(lRow, dtCOLPos.lMasterPackHeight) = "" Then
'            dblUnitHeight = 0
'        Else
'            dblUnitHeight = sItemArray(lRow, dtCOLPos.lMasterPackHeight)
'        End If

'        dblCube = (dblUnitLength * dblUnitWidth * dblUnitHeight) / 1728
'        If dblCube <> 0 Then
'            sItemArray(lRow, dtCOLPos.lMasterPackCube) = sRound(dblCube, 4)
'        Else
'            sItemArray(lRow, dtCOLPos.lMasterPackCube) = ""
'        End If

'        ' Calculate InnerPackCube
'        If IsNull(sItemArray(lRow, dtCOLPos.lInnerPackLength)) Or _
'                sItemArray(lRow, dtCOLPos.lInnerPackLength) = "" Then
'            dblUnitLength = 0
'        Else
'            dblUnitLength = sItemArray(lRow, dtCOLPos.lInnerPackLength)
'        End If
'        If IsNull(sItemArray(lRow, dtCOLPos.lInnerPackWidth)) Or _
'                sItemArray(lRow, dtCOLPos.lInnerPackWidth) = "" Then
'            dblUnitWidth = 0
'        Else
'            dblUnitWidth = sItemArray(lRow, dtCOLPos.lInnerPackWidth)
'        End If
'        If IsNull(sItemArray(lRow, dtCOLPos.lInnerPackHeight)) Or _
'                sItemArray(lRow, dtCOLPos.lInnerPackHeight) = "" Then
'            dblUnitHeight = 0
'        Else
'            dblUnitHeight = sItemArray(lRow, dtCOLPos.lInnerPackHeight)
'        End If

'        dblCube = (dblUnitLength * dblUnitWidth * dblUnitHeight) / 1728
'        If dblCube <> 0 Then
'            sItemArray(lRow, dtCOLPos.lInnerPackCube) = sRound(dblCube, 4)
'        Else
'            sItemArray(lRow, dtCOLPos.lInnerPackCube) = ""
'        End If

'        bCalcCubes = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bCalcCubes")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCalcCubes , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'    Private Function bVendorActive(ByVal sVendorNumber As Object) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rsVendorInactive As ADODB.Recordset

'        bVendorActive = False
'        rsVendorInactive = New ADODB.Recordset
'        sSQL = "SELECT Inactive FROM Vendor WHERE VendorNumber = " & sVendorNumber
'Dim     rsVendorInactive.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object

'        If rsVendorInactive.EOF Then
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        Else
'            '        If rsVendorInactive!InActive = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If
'        bVendorActive = True
'ExitRoutine:
'        If rsVendorInactive.State = 1 Then rsVendorInactive.Close()
'        rsVendorInactive = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bVendorActive Vendor:" & sVendorNumber)
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bVendorActive , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Private Function bVendorFactoryActiveRelationship(ByVal sVendorNumber As Object, ByVal sFactoryNumber As Object) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        bVendorFactoryActiveRelationship = False           '12/28/2007
'        If sVendorNumber = 1000 Then
'            bVendorFactoryActiveRelationship = True
'        Else
'            Dim rsVFREL As ADODB.Recordset
'            rsVFREL = New ADODB.Recordset
'            sSQL = "SELECT InactiveRelation FROM VendorFactory WHERE VendorNumber = " & sVendorNumber & _
'                    " AND FactoryNumber = " & sFactoryNumber
'Dim         rsVFREL.Open sSQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockReadOnly As Object

'            If Not rsVFREL.EOF Then
'                '            If rsVFREL!InactiveRelation = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            End If
'            If rsVFREL.State = 1 Then rsVFREL.Close()
'            rsVFREL = Nothing
'            bVendorFactoryActiveRelationship = True
'        End If
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet: bVendorFactoryActiveRelationship")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bVendorFactoryActiveRelationship, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Private Function bFactoryProhibited(ByVal sFactoryNumber As Object) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rsProhibit As ADODB.Recordset
'        bFactoryProhibited = False
'        rsProhibit = New ADODB.Recordset

'        sSQL = "SELECT Prohibitive FROM Factory WHERE FactoryNumber = " & sFactoryNumber & _
'               " AND Prohibitive <> 0"

'Dim     rsProhibit.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object

'        If Not rsProhibit.EOF Then
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If
'        bFactoryProhibited = True
'ExitRoutine:
'        If rsProhibit.State <> 0 Then rsProhibit.Close()
'        rsProhibit = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bFactoryProhibited - Factory: " & sFactoryNumber)
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFactoryProhibited , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Private Function bFactoryLightSourceOnly(ByVal sFactoryNumber As Object) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rs As ADODB.Recordset
'        bFactoryLightSourceOnly = False
'        rs = New ADODB.Recordset

'        sSQL = "SELECT LightSourceOnly FROM Factory WHERE FactoryNumber = " & sFactoryNumber

'Dim     rs.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object

'        '    If rs.EOF Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        '    If rs!LightSourceOnly = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        bFactoryLightSourceOnly = True
'ExitRoutine:
'        If rs.State <> 0 Then rs.Close()
'        rs = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bFactoryLightSourceOnly - Factory: " & sFactoryNumber)
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFactoryLightSourceOnly, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'    Private Function bFindDuplicateProposalsSpreadsheet(ByVal lCurrentRow As Long, ByVal lRowsOnSheet As Long, ByVal dtCOLPos As typSpecialCOLPos) As Object
'        ' for Proposals can be on the spreadsheet once
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lCounter As Long
'        Dim sDuplicateRowsList As String
'        Dim sFunctioncode As String
'        Dim sCheckFunctionCode As String

'        Dim sProposalNumber As String
'        Dim sCheckProposalNumber As String

'        Dim sItemNumber As String
'        Dim sCheckItemNumber As String


'        bFindDuplicateProposalsSpreadsheet = False
'        sFunctioncode = sItemArray(lCurrentRow, 1)
'        sProposalNumber = sItemArray(lCurrentRow, 2)
'        sItemNumber = sItemArray(lCurrentRow, dtCOLPos.lItemNumber)
'        For lCounter = glDATA_START_ROW To lRowsOnSheet
'            If lCounter <> lCurrentRow Then
'                sCheckFunctionCode = sItemArray(lCounter, 1)
'                sCheckProposalNumber = sItemArray(lCounter, 2)
'                sCheckItemNumber = sItemArray(lCounter, dtCOLPos.lItemNumber)

'                If sProposalNumber = sCheckProposalNumber And sProposalNumber <> "" Then
'                    sDuplicateRowsList = sDuplicateRowsList & CStr(lCounter) & ", "
'                End If

'            End If
'        Next lCounter

'        If Len(sDuplicateRowsList) > 0 Then
'            sErrorMsg = "Duplicate Proposal on spreadsheet for FC = " & sFunctioncode & _
'                            " ProposalNumber: " & sProposalNumber & ", " & _
'                            " ItemNumber: " & sItemNumber & _
'                            " Duplicate Row: " & sDuplicateRowsList
'            sErrorMsg = Microsoft.VisualBasic.Left(sErrorMsg, Len(sErrorMsg) - 2)
'            bFindDuplicateProposalsSpreadsheet = True
'        End If
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bFindDuplicateProposalsSpreadsheet")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFindDuplicateProposalsSpreadsheet , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'Private Function bCheckHSNumberANDDutyPercent(ByVal lCurrentRow As Long, ByVal lRowsOnSheet As Long, _
'                dtCOLPos As typSpecialCOLPos, sItemArray() As String , _
'        End Function

'    Private Function bCheckHSNumberANDDutyPercent(ByVal lCurrentRow As Long, ByVal lRowsOnSheet As Long, ByVal dtCOLPos As typSpecialCOLPos) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        bCheckHSNumberANDDutyPercent = False
'        'if HSNumber present, DuryPercent cant be null, 0 is ok
'        sHSNumber = sItemArray(lCurrentRow, dtCOLPos.lHSNumber)
'        sDutyPercent = sItemArray(lCurrentRow, dtCOLPos.lDutyPercent)
'        '    If IsNull(sHSNumber) Or sHSNumber = "" Then GoTo ExitOK'TODO - GoTo Statements are redundant in .NET

'        '    If IsNull(sDutyPercent) Or sDutyPercent = "" Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET

'ExitOK:
'        bCheckHSNumberANDDutyPercent = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bCheckHSNumberANDDutyPercent")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCheckHSNumberANDDutyPercent , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'Private Function bCheckSpreadsheetForVenFactItemNum(ByVal lCurrentRow As Long, ByVal lRowsOnSheet As Long, _
'        End Function

'    Private Function bCheckSpreadsheetForVenFactItemNum(ByVal lCurrentRow As Long, ByVal lRowsOnSheet As Long, ByVal dtCOLPos As typSpecialCOLPos) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lRowCounter As Long
'        Dim sCheckFunctionCode As String   'holds value of spreadsheet row
'        Dim sCheckVendorNumber As String
'        Dim sCheckFactoryNumber As String
'        Dim sCheckVendorItemNumber As String

'        Dim sCheckCategory As String

'        sSpreadsheetRow = ""
'        bCheckSpreadsheetForVenFactItemNum = False
'        For lRowCounter = glDATA_START_ROW To lRowsOnSheet
'            If lRowCounter <> lCurrentRow Then

'                sCheckFunctionCode = sItemSPECSArray(lRowCounter, 1)
'                If sCheckFunctionCode <> "" Then
'                    sCheckCategory = sItemArray(lRowCounter, dtCOLPos.lCategoryCode)
'                    Select Case sCheckCategory
'                        Case "EL", "BG", "ELLS", "ELOUT", "ELOTH"
'                            sCheckVendorNumber = sItemArray(lRowCounter, dtCOLPos.lVendorNumber)
'                            sCheckFactoryNumber = sItemArray(lRowCounter, dtCOLPos.lFactoryNumber)
'                            sCheckVendorItemNumber = sItemArray(lRowCounter, dtCOLPos.lVendorItemNumber)
'                            'check against spreadsheet combination values
'                            If sCheckVendorNumber = sVendorNumber And _
'                                sCheckFactoryNumber = sFactoryNumber And _
'                                sCheckVendorItemNumber = sVendorItemNumber And _
'                                Not IsNull(sCheckVendorItemNumber) And _
'                                sCheckVendorItemNumber <> "" Then
'                                sSpreadsheetRow = sSpreadsheetRow & lRowCounter & ", "
'                            End If

'                    End Select

'                End If
'            End If
'        Next
'        If sSpreadsheetRow = "" Then
'            bCheckSpreadsheetForVenFactItemNum = True
'        End If
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bCheckSpreadsheetForVenFactItemNum")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCheckSpreadsheetForVenFactItemNum , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'        '2010/02/11 - added lRowChangesFound
'        '2012/01/04 - added lRowCounter
'        'Private Function bCheckItemFieldsChanged(RowChangesARRAY As String, sFromProcess As String, sItemSpecsArray_ORIG As String, sItemSPECSArray As String, lItemSpecsFields As Long, _
'        '                    sItemArray_ORIG() As String , sItemArray() As String , lItemFields As Long, _
'        '                    sAssortmentArray_ORIG() As String , sAssortmentArray() As String , lAssortmentFields As Long, _
'        '                    lMaxRows As Long, dtSaveArrayCOLPos As typSpecialCOLPos) As Boolean
'        ''On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        'Dim lCOLCounter     As Long
'        Dim lRowCounter As Long

'        'Dim sSQL            As String 
'        'Dim sProposalNumber As String 
'        Dim sRev As String
'        Dim sPrevRev As String

'        'Dim sItemNumber     As String 
'        'Dim sFunctionCode   As String
'        'Dim dRevisionDate   As Date
'        'Dim sArray_ORIG()   As String 
'        Dim sArray() As String

'        'Dim RS As ADODB.Recordset: Set RS = New ADODB.Recordset
'        'Dim sCounter        As String  '* 2
'        'Dim lRowChangeFound As Long             '2010/02/11
'        '
'        '    bCheckItemFieldsChanged = False
'        '    dRevisionDate = Now.ToShortDateString() 'not used, because Access does not give us the time in milliseconds
'        '   ' define default value on ItemFieldHistory table as getdate() which gives us milliseconds
'        '
'        '    sSQL = "SELECT * FROM ItemFieldHistory WHERE RevisedDate > '" & dRevisionDate & "'"
'        '    RS.Open sSQL, SSDataConn, adOpenDynamic, adLockPessimistic
'        '
'        '    If bPROPOSALFormIndicator = False Then
'        '        sArray_ORIG() = sItemSpecsArray_ORIG()          '2010/02/12 - first check ItemSpecs fields
'        '        sArray() = sItemSPECSArray()
'        '
'        '        For lRowCounter = 2 To lMaxRows
'        '            sFunctionCode = sArray(lRowCounter, 1)
'        '            sProposalNumber = sArray(lRowCounter, 2)
'        '            sRev = sArray(lRowCounter, 3)
'        '            sPrevRev = sArray_ORIG(lRowCounter, 3)
'        '            sItemNumber = sItemArray_ORIG(lRowCounter, dtSaveArrayCOLPos.lItemNumber)
'        '            If sFunctionCode = "" Or sFunctionCode = gsNEW_ITEM_NBR Then
'        '            Else
'        '                For lCOLCounter = 4 To lItemSpecsFields + glFIXED_COLS  '2011/02/12 was using lItemFields
'        '                    If Not sArray_ORIG(lRowCounter, lCOLCounter) = sArray(lRowCounter, lCOLCounter) Then
'        '                        lRowChangeFound = RowChangesARRAY(lRowCounter)  '2010/02/11
'        '                        'need ItemNumber for the ItemFieldHistory table
'        '                        sItemNumber = sItemArray_ORIG(lRowCounter, dtSaveArrayCOLPos.lItemNumber)
'        '                        If bSAVEItemFieldsChanged(lRowChangeFound, RS, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctionCode, sArray_ORIG(1, lCOLCounter), _
'        '                                 sArray_ORIG(lRowCounter, lCOLCounter), sArray(lRowCounter, lCOLCounter), _
'        '                                dRevisionDate, gsUserID) = False Then
'        '                        End If
'        '                    End If
'        '                Next lCOLCounter
'        '            End If
'        '
'        '        Next lRowCounter
'        '    End If
'        '
'        '    sArray_ORIG() = sItemArray_ORIG()           '2010/02/12 -then check Item fields
'        ''    sArray() = sItemSPECSArray()               '2010/02/12 ? why here?
'        '    sArray() = sItemArray()
'        '
'        '    For lRowCounter = 2 To lMaxRows
'        '        sFunctionCode = sArray(lRowCounter, 1)
'        '        sProposalNumber = sArray(lRowCounter, 2)
'        '        sRev = sArray(lRowCounter, 3)
'        '        sPrevRev = sArray_ORIG(lRowCounter, 3)
'        '        sItemNumber = sItemArray_ORIG(lRowCounter, dtSaveArrayCOLPos.lItemNumber)
'        '        If sFunctionCode = "" Or sFunctionCode = gsNEW_ITEM_NBR Then
'        '        Else
'        '            For lCOLCounter = 3 To lItemFields + glFIXED_COLS           '2010/02/11 was 4 trying to show rev change, if a new rev without any other changes
'        '                If sArray_ORIG(1, lCOLCounter) = "RevisedUserID" Or _
'        '                    sArray_ORIG(1, lCOLCounter) = "RevisedDate" Then
'        '                    Application.DoEvents
'        '                Else
'        '                    If Not sArray_ORIG(lRowCounter, lCOLCounter) = sArray(lRowCounter, lCOLCounter) Then
'        '                        lRowChangeFound = RowChangesARRAY(lRowCounter)  '2010/02/11
'        '                        If bSAVEItemFieldsChanged(lRowChangeFound, RS, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctionCode, sArray_ORIG(1, lCOLCounter), _
'        '                                 sArray_ORIG(lRowCounter, lCOLCounter), sArray(lRowCounter, lCOLCounter), _
'        '                                dRevisionDate, gsUserID) = False Then
'        ''                            GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'        '                        End If
'        '                    End If
'        '                End If
'        '            Next lCOLCounter
'        '        End If
'        '    Next lRowCounter
'        '
'        '    If bPROPOSALFormIndicator = False Then 'do the following for Import Spreadsheet
'        '        sArray_ORIG() = sAssortmentArray_ORIG()     '2010/02/12 Item_Assortment fields
'        '        sArray() = sAssortmentArray()
'        '
'        '        For lRowCounter = 2 To lMaxRows
'        '            sFunctionCode = sArray(lRowCounter, 1)
'        '            sProposalNumber = sArray(lRowCounter, 2)
'        '            sRev = sArray(lRowCounter, 3)
'        '            sPrevRev = sArray_ORIG(lRowCounter, 3)
'        '            sItemNumber = sItemArray_ORIG(lRowCounter, dtSaveArrayCOLPos.lItemNumber)
'        '
'        '            If sFunctionCode = "" Or sFunctionCode = gsNEW_ITEM_NBR Then
'        '            Else
'        '                For lCOLCounter = 4 To lAssortmentFields + glFIXED_COLS           '2010/02/15 - dont start at 3 otherwise it writes a rev 2X- was going against lItemFields!'2010/02/11 was 4 trying to show rev change, if a new rev without any other changes
'        '                    If sArray_ORIG(1, lCOLCounter) = "RevisedUserID" Or _
'        '                        sArray_ORIG(1, lCOLCounter) = "RevisedDate" Then
'        '                    Else
'        '                        If sArray_ORIG(1, lCOLCounter) = "ASSORTMENTS" Then
'        '                        Application.DoEvents
'        '                        End If
'        '                    End If
'        '                    If Not sArray_ORIG(lRowCounter, lCOLCounter) = sArray(lRowCounter, lCOLCounter) Then
'        '                        lRowChangeFound = RowChangesARRAY(lRowCounter)  '2010/02/11
'        '                        If bSAVEItemFieldsChanged(lRowChangeFound, RS, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctionCode, sArray_ORIG(1, lCOLCounter), _
'        '                                 sArray_ORIG(lRowCounter, lCOLCounter), sArray(lRowCounter, lCOLCounter), _
'        '                                dRevisionDate, gsUserID) = False Then
'        ''                                GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'        '                        End If
'        '                    End If
'        '                Next lCOLCounter
'        '            End If
'        '        Next lRowCounter
'        '    Else
'        '    'do the following for Proposal Form Assortment changes
'        '        If gbProposalAssortmentsChanged = True Then
'        '            If sFunctionCode = "" Or sFunctionCode = gsNEW_ITEM_NBR Then
'        '        Else
'        '            For lCOLCounter = 0 To 40
'        '                If lCOLCounter + 1 < 10 Then
'        '                    sCounter = "0" & lCOLCounter + 1
'        '                Else
'        '                    sCounter = lCOLCounter + 1
'        '                End If
'        '
'        '                '2010/02/12 - for ProposalForm lRowCounter ALWAYS = 2; like it's a spreadsheet with 1 row only
'        '                lRowCounter = 2
'        '                                    'Assortment ItemNumbers
'        '                If Not sORIGAssortmentArray(lCOLCounter, 1) = sNEWAssortmentArray(lCOLCounter, 1) Then
'        '                    RowChangesARRAY(lRowCounter) = RowChangesARRAY(lRowCounter) + 1
'        '                    lRowChangeFound = RowChangesARRAY(lRowCounter)           '2010/02/11
'        '                    If bSAVEItemFieldsChanged(lRowChangeFound, RS, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctionCode, _
'        '                           cITEM_XX & sCounter, _
'        '                           sORIGAssortmentArray(lCOLCounter, 1), sNEWAssortmentArray(lCOLCounter, 1), _
'        '                           dRevisionDate, gsUserID) = False Then
'        ''                           GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'        '                   End If
'        '                End If
'        '                                           'Assortment Quantity's
'        '                If Not sORIGAssortmentArray(lCOLCounter, 2) = sNEWAssortmentArray(lCOLCounter, 2) Then
'        '                    RowChangesARRAY(lRowCounter) = RowChangesARRAY(lRowCounter) + 1
'        '                    lRowChangeFound = RowChangesARRAY(lRowCounter)           '2010/02/11
'        ''                    lRowChangeFound = RowChangesARRAY(lROWCounter)           '2010/02/11
'        '                    If bSAVEItemFieldsChanged(lRowChangeFound, RS, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctionCode, _
'        '                            cQTY_XX & sCounter, _
'        '                           sORIGAssortmentArray(lCOLCounter, 2), sNEWAssortmentArray(lCOLCounter, 2), _
'        '                           dRevisionDate, gsUserID) = False Then
'        ''                           GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'        '                   End If
'        '                End If
'    End Function

'    Friend Overridable Function bSAVEItemFieldsChanged(ByVal lRowChangesFound As Long, ByVal rs As ADODB.Recordset) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally

'        bSAVEItemFieldsChanged = False
'        Dim sSQL As String                       '2013/04/23 -HN
'        Select Case sFieldName
'            Case "BaseProposalNumber", "BaseRev"        'dont bother to follow these ancient fields
'            Case Else
'                '2013/04/23 -HN- added ET's cursor conflict error code
'                'following has error when field contains '(quotes), easier to do ADODB way
'                '    sSQL = "INSERT INTO ItemFieldHistory(ProposalNumber, Rev, FunctionCode, FieldName,OldValue, NewValue, RevisedDate, RevisedUserID)" & _
'                '                   " VALUES( '" & sProposalNumber & "', '" & sRev & "', '" & _
'                '                    sFunctionCode & "', '" & sFieldName & "', '" & sOldValue & "', '" & _
'                '                    sNewValue & "', '" & sRevisedDate & "' ,'" & sRevisedUserId & "')"
'                '            SSDataConn.Execute sSQL

'                ' ET 2013-03-25 - restore SQL command method in place of recordset update in an
'                ' attempt to prevent "cursor operation conflict" errors
'                'following has error when field contains '(quotes), easier to do ADODB way
'                '2014/01/14 RAS Adding info to Trace message
'                If glTraceFlag = True Then
'                    If bWritePrintToLogFile(False, objEXCELName & Space(6) & "Inserting into ItemFieldHistory ,bSAVEItemFieldsChanged for ProposalNumber: " & sProposalNumber & " ,Rev: " & sRev, Format(Now(), "yyyymmdd")) = False Then
'                    End If
'                End If
'                If IsBlank(sOldValue) Then sOldValue = "" '2013/06/05 -HN- was causing an error below!
'                If IsBlank(sNewValue) Then sNewValue = ""
'                sSQL = "INSERT INTO ItemFieldHistory(ProposalNumber, Rev, PrevRev, ItemNumber, FunctionCode, Process, " & _
'                    "FieldName, OldValue, NewValue, RevisedUserID)" & _
'                    " VALUES( '" & sProposalNumber & "', '" & sRev & "', '" & sPrevRev & "', '" & sItemNumber & "', '" & _
'                    sFunctioncode & "', '" & sFromProcess & "', '" & sFieldName & "', '" & Replace(sOldValue, "'", "''") & "', '" & _
'                    Replace(sNewValue, "'", "''") & "', '" & sRevisedUserID & "')"
'                SSDataConn.Execute sSQL

'                '            RS.AddNew
'                '            RS![ProposalNumber] = sProposalNumber
'                '            RS![Rev] = sRev
'                '            RS![PrevRev] = sPrevRev
'                '            RS![ItemNumber] = sItemNumber
'                '            RS![FunctionCode] = sFunctionCode
'                '            RS![Process] = sFromProcess
'                '            RS![FieldName] = sFieldName
'                '            RS![OldValue] = sOldValue
'                '            RS![NewValue] = sNewValue
'                ''            rs![RevisedDate] = sRevisedDate 'use Getdate() as default on table to get the time in milliseconds
'                '            RS![RevisedUserID] = sRevisedUserID
'                '            RS.Update
'        End Select

'        bSAVEItemFieldsChanged = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & vbCrLf & " Proposal:" & sProposalNumber & " Rev:" & sRev & " for Field:" & sFieldName & vbCrLf & vbCrLf & _
'            "OLD Value: " & sOldValue & vbCrLf & vbCrLf & _
'            "NEW Value: " & sNewValue, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bSaveItemFieldsChanged-ItemFieldHistory-error")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bSaveItemFieldsChanged-ItemFieldHistory-error,Proposal:" & sProposalNumber & " Rev:" & sRev & " for Field:" & sFieldName & "Error Number:" & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'        ' Resume Next ' 2014/01/15 RAS Commented this out.  it will never hit it because of the Resume ExitRoutine call

'Dim Public Function bFormatSpreadsheet(objExcel As Excel.Application
'Dim  frmThis As Form) As Boolean

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'Dim SpreadsheetCOLArray(1 To glMAX_Cols) As typColumn
'Dim sColHeader(1 To glMAX_Cols) As String 
'Dim lColorIndex(1 To glMAX_Cols) As Long
'        Dim lColsOnSheet As Long
'        Dim lDB4FieldsFound As Long
'        Dim lProposalNumberCOLPos As Long
'        Dim lRevCOLPos As Long

'        Dim lProgYearCOLPos As Long                                                         '09/29/2008 - hn
'        Dim lFactoryNumberCOLPos As Long                                                         'new 10/15/2007
'        Dim lProgNumColPos As Long
'        Dim dblColumnWidth As Double
'        Dim lCounter As Long
'        Dim lLongDescColumn As Long
'        Dim lItemNumCOLPos As Long
'        Dim lFunctionCodeCOLPos As Long
'        Dim lPhotoCOLPos As Long

'        Dim lAlternatePhotoCOLPos As Long                                                         '2010/09/09
'        Dim lBag_SpecialEffectsCOLPos As Long                                                         '2009/11/18 - hn
'        Dim lTechnologiesColPos As Long
'        Dim lX_TechnologiesCOLPos As Long
'        Dim lLightedCOLPos As Long

'        Dim lCertifiedPrinterIDCOLPos As Long
'        Dim lX_CertifiedPrinterNameCOLPos As Long                  '2011/10/26

'        Dim lX_PalletUPCCOLPos As Long                                                         '2010/10/25
'        Dim lX_UPCCOLPos As Long                                                         '2010/10/25
'        Dim lProductBatteriesCOLPos As Long                                                         '03/19/2008
'    End Function

'    Friend Overridable Function bFormatSpreadsheet(ByVal objExcel As Excel.Application, ByVal frmThis As Form) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'Dim SpreadsheetCOLArray(1 To glMAX_Cols) As typColumn
'Dim sColHeader(1 To glMAX_Cols) As String 
'Dim lColorIndex(1 To glMAX_Cols) As Long
'        Dim lColsOnSheet As Long
'        Dim lDB4FieldsFound As Long
'        Dim lProposalNumberCOLPos As Long
'        Dim lRevCOLPos As Long

'        Dim lProgYearCOLPos As Long                                                         '09/29/2008 - hn
'        Dim lFactoryNumberCOLPos As Long                                                         'new 10/15/2007
'        Dim lProgNumColPos As Long
'        Dim dblColumnWidth As Double
'        Dim lCounter As Long
'        Dim lLongDescColumn As Long
'        Dim lItemNumCOLPos As Long
'        Dim lFunctionCodeCOLPos As Long
'        Dim lPhotoCOLPos As Long

'        Dim lAlternatePhotoCOLPos As Long                                                         '2010/09/09
'        Dim lBag_SpecialEffectsCOLPos As Long                                                         '2009/11/18 - hn
'        Dim lTechnologiesColPos As Long
'        Dim lX_TechnologiesCOLPos As Long
'        Dim lLightedCOLPos As Long

'        Dim lCertifiedPrinterIDCOLPos As Long
'        Dim lX_CertifiedPrinterNameCOLPos As Long                  '2011/10/26

'        Dim lX_PalletUPCCOLPos As Long                                                         '2010/10/25
'        Dim lX_UPCCOLPos As Long                                                         '2010/10/25
'        Dim lProductBatteriesCOLPos As Long                                                         '03/19/2008
'        Dim lAssortmentsCOLPos As Long                                                         '04/16/2008 - hn

'        bFormatSpreadsheet = False

'        objExcel.Application.Workbooks(1).Worksheets(1).Cells.Select()
'        objExcel.Application.Selection.WrapText = True

'        ' Column Header row
'        objExcel.Application.Workbooks(1).Worksheets(1).Rows("1:1").Select()
'        objExcel.Application.Selection.RowHeight = 33

'        '03/19/2008 addded lProductBatteriesCOLPos
'        '04/16/2008 - hn - added lX_AssortmentsColPos:
'        '09/29/2008 - hn - added lProgYearCOLPos for Tech Code checking
'        '2009/11/18 - hn - added: lBag_SpecialEffectsCOLPos
'        '2010/10/25 - added lX_PalletUPCColPos, lUPCColPos,
'        '2011/10/26 - added CertifiedPrinterID's col pos
'        '2012/11/12 - added sColHeader
'        If bLoadSpreadsheetColumnArray(objExcel, SpreadsheetCOLArray(), SpreadsheetMaterialColumnX(), lColsOnSheet, _
'                    lMaxMaterialCols, lDB4FieldsFound, lProposalNumberCOLPos, lRevCOLPos, lItemNumCOLPos, _
'                    lProgYearCOLPos, lFactoryNumberCOLPos, lFunctionCodeCOLPos, lPhotoCOLPos, lAlternatePhotoCOLPos, _
'                    lLightedCOLPos, lProductBatteriesCOLPos, lBag_SpecialEffectsCOLPos, lCertifiedPrinterIDCOLPos, _
'                    lX_CertifiedPrinterNameCOLPos, lTechnologiesColPos, lX_TechnologiesCOLPos, lX_UPCCOLPos, _
'                    lX_PalletUPCCOLPos, lAssortmentsCOLPos, sColHeader(), frmThis) = False Then
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        Else
'            If gbCancelExport = True Then Exit Function
'            For lCounter = 1 To lColsOnSheet
'                ' Size each column individually
'                Call bUpdateStatusMessage(frmThis, "Sizing Column " & CStr(lCounter) & "...")

'                dblColumnWidth = SpreadsheetCOLArray(lCounter).dblWidth
'                If dblColumnWidth > 0 Then
'                    objExcel.Application.Workbooks(1).Worksheets(1).Columns(lCounter).Select()
'                    objExcel.Application.Selection.ColumnWidth = dblColumnWidth
'                End If
'            Next
'        End If

'        Call bUpdateStatusMessage(frmThis, "Sizing Rows...")
'        lLongDescColumn = lGetSpreadsheetCOL(gsCOL_LONGDESC, SpreadsheetCOLArray(), lColsOnSheet)

'        If lLongDescColumn > 0 Then
'            objExcel.Application.Workbooks(1).Worksheets(1).Columns(lLongDescColumn).Select()
'            objExcel.Application.Selection.Rows.AutoFit()
'        End If

'        ' Return to the first column in the worksheet
'        objExcel.Application.Workbooks(1).Worksheets(1).Columns(1).Select()

'        bFormatSpreadsheet = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bFormatSpreadsheet")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFormatSpreadsheet , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'        'commented 2011/08/31
'Dim 'Public Function bTranslateOldColumnNames(ByVal sOldColHeading As String
'Dim  ByRef sNewColumnName As String ) As Boolean

'        ''On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        '    bTranslateOldColumnNames = False
'        '    sNewColumnName = ""
'Dim '    'do it this way As Object 
'Dim  in case the column name is aliased between < > As Object 
' eg: Cust<Dimension_Customer> As Object 'Translation Error-

'        '
'        '    '03/14/2008 - changed Electrics Specs columns
'        '    If InStr(1, sOldColHeading, "LEDCoverType") Then
'        '        sNewColumnName = Replace(sOldColHeading, "LEDCoverType", "LEDCoverSize")
'        '
'        '    ElseIf InStr(1, sOldColHeading, "RechargeCloudyDayHrs") Then
'        '        If sOldColHeading = "RechargeCloudyDayHrs" Then
'        '            sNewColumnName = Replace(sOldColHeading, "RechargeCloudyDayHrs", "SolarRechargeCloudyDayHrs")
'        '        ElseIf InStr(1, sOldColHeading, "<RechargeCloudyDayHrs>") Then
'        '            sNewColumnName = Replace(sOldColHeading, "RechargeCloudyDayHrs", "SolarRechargeCloudyDayHrs")
'        '        End If
'        '
'        '    ElseIf InStr(1, sOldColHeading, "RechargeFullSunHrs") Then
'        '        If sOldColHeading = "RechargeFullSunHrs" Then
'        '            sNewColumnName = Replace(sOldColHeading, "RechargeFullSunHrs", "SolarRechargeFullSunHrs")
'        '        ElseIf InStr(1, sOldColHeading, "<RechargeFullSunHrs>") Then
'        '            sNewColumnName = Replace(sOldColHeading, "RechargeFullSunHrs", "SolarRechargeFullSunHrs")
'        '        End If
'        '    ElseIf InStr(1, sOldColHeading, "RunTimeAt0CHrs") Then
'        '        If sOldColHeading = "RunTimeAt0CHrs" Then
'        '            sNewColumnName = Replace(sOldColHeading, "RunTimeAt0CHrs", "SolarRunTimeAt0CHrs")
'        '        ElseIf InStr(1, sOldColHeading, "<RunTimeAt0CHrs>") Then
'        '            sNewColumnName = Replace(sOldColHeading, "RunTimeAt0CHrs", "SolarRunTimeAt0CHrs")
'        '        End If
'        '
'Dim '    ElseIf InStr(1 As Object 
'        Dim sOldColHeading As Object
' "Dimension_Category") Then As Object 'Translation Error-

'        '    '-------------
'Dim '        sNewColumnName = Replace(sOldColHeading As Object 
' "Dimension_Category" As Object 'Translation Error-
'Dim  "CategoryCode") As Object 

'Dim '    ElseIf InStr(1 As Object 
'        Dim sOldColHeading As Object
' "Dimension_Customer") Then As Object 'Translation Error-

'Dim '        sNewColumnName = Replace(sOldColHeading As Object 
' "Dimension_Customer" As Object 'Translation Error-
'Dim  "CustomerNumber") As Object 

'Dim '    ElseIf InStr(1 As Object 
'        Dim sOldColHeading As Object
' "Dimension_Season") Then As Object 'Translation Error-


'Dim '        sNewColumnName = Replace(sOldColHeading As Object 
' "Dimension_Season" As Object 'Translation Error-

'Dim  "SeasonCode") As Object 

'Dim '    ElseIf InStr(1 As Object 
'        Dim sOldColHeading As Object
' "Dimension_Program") Then As Object 'Translation Error-

'Dim '        sNewColumnName = Replace(sOldColHeading As Object 
' "Dimension_Program" As Object 'Translation Error-
'Dim  "ProgramNumber") As Object 

'        '    ElseIf InStr(1, sOldColHeading, "ForAccount") Then
'        '        sNewColumnName = Replace(sOldColHeading, "ForAccount", "CustomerNumber")
'        '    ElseIf InStr(1, sOldColHeading, "Program_Year") Then
'        '        sNewColumnName = Replace(sOldColHeading, "Program_Year", "ProgramYear")
'        '    ElseIf InStr(1, sOldColHeading, "RoyaltyID") Then
'        '        sNewColumnName = Replace(sOldColHeading, "RoyaltyID", "Licensor")
'        '    ElseIf InStr(1, sOldColHeading, "SalesRep") Then
'        '        If sOldColHeading = "SalesRep" Then
'        '            sNewColumnName = Replace(sOldColHeading, "SalesRep", "SalesRepNumber")
'        '        ElseIf InStr(1, sOldColHeading, "<SalesRep>") Then
'        '            sNewColumnName = Replace(sOldColHeading, "SalesRep", "SalesRepNumber")
'        '        End If
'        '    ElseIf InStr(1, sOldColHeading, "SeasonalRevisionTime") Then
'        '        sNewColumnName = Replace(sOldColHeading, "SeasonalRevisionTime", "RevisedDate")
'        '    ElseIf InStr(1, sOldColHeading, "SubProgramNumber") Then
'        '        sNewColumnName = Replace(sOldColHeading, "SubProgramNumber", "SubProgram")
'        '    ElseIf InStr(1, sOldColHeading, "X_SubProgramName") Then
'        '        sNewColumnName = Replace(sOldColHeading, "X_SubProgramName", "X_SubProgram")
'        '    ElseIf InStr(1, sOldColHeading, "Tree_SingleWireConstruction") Then
'        '        sNewColumnName = Replace(sOldColHeading, "Tree_SingleWireConstruction", "Tree_LightConstruction")
'        '    ElseIf InStr(1, sOldColHeading, "TemporaryItemNumber") Then
'        '        sNewColumnName = Replace(sOldColHeading, "TemporaryItemNumber", "TempItemNumber")
'Dim '    ElseIf InStr(1 As Object 
'        Dim sOldColHeading As Object
' "Dimension_Item") Then As Object 'Translation Error-

'Dim '        sNewColumnName = Replace(sOldColHeading As Object 
' "Dimension_Item" As Object 'Translation Error-
'Dim  "CoreItemNumber") As Object 

'        '    ElseIf InStr(1, sOldColHeading, "RevisedByUserID") Then
'        '        sNewColumnName = Replace(sOldColHeading, "RevisedByUserID", "RevisedUserID")
'        '    ElseIf InStr(1, sOldColHeading, "X_SeasonalRevisionTime") Then
'        '        sNewColumnName = Replace(sOldColHeading, "X_SeasonalRevisionTime", "X_RevisedDate")
'Dim '    ElseIf InStr(1 As Object 
'        Dim sOldColHeading As Object
' "X_Dimension_Item") Then As Object 'Translation Error-

'Dim '        sNewColumnName = Replace(sOldColHeading As Object 
' "X_Dimension_Item" As Object 'Translation Error-
'Dim  "X_CoreItemNumber") As Object 

'        '    ElseIf InStr(1, sOldColHeading, "X_DateAdded") Then
'        '        sNewColumnName = Replace(sOldColHeading, "X_DateAdded", "X_CreatedDate")
'        '    ElseIf InStr(1, sOldColHeading, "PkgItemRef") Then
'        '        sNewColumnName = Replace(sOldColHeading, "PkgItemRef", "PackageItemRef")
'        '
'        '    ElseIf InStr(1, sOldColHeading, "AltQtdPrice") Then                                     '11/25/2008 - hn
'        '        sNewColumnName = Replace(sOldColHeading, "AltQtdPrice", "AltSellPrice")
'        '
'        ''    ElseIf InStr(1, sOldColHeading, "Assortments") Then                                     '01/19/2009 - hn
'        ''        If sOldColHeading = "Assortments" Then
'        ''            sNewColumnName = Replace(sOldColHeading, "Assortments", "X_Assortments")        '04/16/2006 - hn - now calculated on the fly
'        ''        End If
'        '    End If
'        '    'done in ConcatenateOldMaterialcolumns for Import
'        ''    If gbFromREFRESH = True Then
'        ''        If Microsoft.VisualBasic.Left(sOldColHeading, 9) = "Breakdown" And Len(sOldColHeading) = 10 Then
'        ''            sNewColumnName = Replace(sOldColHeading, "Breakdown", "NotUsed_Breakdown")
'        ''        ElseIf Microsoft.VisualBasic.Left(sOldColHeading, 4) = "Cost" And Len(sOldColHeading) = 5 Then
'        ''            sNewColumnName = Replace(sOldColHeading, "Cost", "NotUsed_Cost")
'        ''        End If
'        ''    End If
'        '
'        '    bTranslateOldColumnNames = True
'        'ExitRoutine:
'        '    Exit Function
'        'ErrorHandler:
'        '    MsgBox Err.Description & ": Old Column Heading: " & sOldColHeading, vbExclamation, "modSpreadSheet-bTranslateOldColumnNames"
'        '    Resume ExitRoutine
'    End Function

'    Private Function bFindInactiveProgramYear(ByVal lProgramNumber As Long, ByRef bGradeRequired As Boolean) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        bFindInactiveProgramYear = False

'        sSQL = "SELECT GradeRequired, InactiveProgramYear FROM Program WHERE ProgramNumber = " & lProgramNumber
'Dim     rs.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly      '11/09/2007 As Object 


'        If rs.EOF = True Then
'            sInactiveProgramYear = ""
'            bGradeRequired = False
'        Else
'            If Not IsNull(rs!InactiveProgramYear) Then
'                sInactiveProgramYear = rs!InactiveProgramYear
'                '            bGradeRequired = RS!GradeRequired
'            End If
'            bGradeRequired = rs!GradeRequired    'new 11/09/2007 removed from If statement above
'        End If

'        bFindInactiveProgramYear = True
'ExitRoutine:
'        If rs.State <> 0 Then rs.Close()
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bFindProgramInactiveProgramYear")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bFindProgramInactiveProgramYear , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'        Private Function bValidCategoryProgramCombination(ByVal sCategory As String, ByVal lProgNumber As Long, ByVal ByVal lProgramYear As Long, ByVal sInactiveProgramYear As String) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        bValidCategoryProgramCombination = False
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset

'        sSQL = "SELECT CategoryCode, ProgramNumber, ProgramName, InactiveProgramYear FROM " & _
'                "Program WHERE ProgramNumber = " & lProgNumber
'        'CategoryCode = '" & sCategory & "' AND
'Dim     rs.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic As Object


'        If Not rs.EOF Then
'            If IsNull(rs!CategoryCode) Or rs!CategoryCode <> sCategory Then
'                If sInactiveProgramYear <> "" Then
'                    If lProgramYear < CInt(sInactiveProgramYear) Then
'                        bValidCategoryProgramCombination = True
'                    Else
'                        bValidCategoryProgramCombination = False
'                    End If
'                Else
'                    If lProgramYear <= 2007 Then
'                        bValidCategoryProgramCombination = True
'                    Else
'                        bValidCategoryProgramCombination = False
'                    End If
'                End If

'            Else
'                If sInactiveProgramYear = "" Then
'                    bValidCategoryProgramCombination = True
'                Else
'                    If lProgramYear < CInt(sInactiveProgramYear) Then
'                        bValidCategoryProgramCombination = True
'                    Else
'                        bValidCategoryProgramCombination = False
'                    End If
'                End If
'            End If
'        Else
'            If sInactiveProgramYear = "" Then
'                bValidCategoryProgramCombination = True
'            Else
'                If lProgramYear < CInt(sInactiveProgramYear) Then
'                    bValidCategoryProgramCombination = True
'                Else
'                    bValidCategoryProgramCombination = False
'                End If
'            End If
'        End If

'ExitRoutine:
'        If rs.State <> 0 Then rs.Close()
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidCategoryProgramCombination")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidCategoryProgramCombination , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'    Private Function bValidProgramSubProgram(ByVal lSpreadsheetProgNumber As Long, ByVal sInactiveProgramYear As String, ByRef sSubProgErrMsg As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        Dim RS2 As ADODB.Recordset : RS2 = New ADODB.Recordset

'        'SubProgramRequired removed from Program table but is always true for each program, except where InactiveProgramYear is null
'        bValidProgramSubProgram = False

'        sSQL = "SELECT SubProgram, InactiveProgramYear AS InactiveSubProgramYear FROM ProgramSubProgram  " & _
'                    "WHERE ProgramNumber = " & lSpreadsheetProgNumber
'        sSQL = sSQL & " AND convert(varchar,SubProgram) = " & sAddQuotes(sSpreadsheetSubProgram)

'Dim     rs.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object

'        If rs.EOF Then
'            'that SubProgram does NOT exist for that Program on ProgramSubProgram Table
'            If sSpreadsheetSubProgram = "" Then
'                sSQL = "SELECT SubProgram, InactiveProgramYear AS InactiveSubProgramYear FROM ProgramSubProgram  " & _
'                    "WHERE ProgramNumber = " & lSpreadsheetProgNumber
'Dim              RS2.Open sSQL As Object 
'                Dim SSDataConn As Object
'                Dim adOpenStatic As Object
'                Dim adLockReadOnly As Object

'                If RS2.EOF Then
'                    bValidProgramSubProgram = True
'                Else
'                    If lSpreadsheetProgramYear < 2008 Then
'                        bValidProgramSubProgram = True
'                    Else
'                        bValidProgramSubProgram = False
'                        sSubProgErrMsg = "SubProgram[    ] invalid"
'                    End If
'                End If
'            Else
'                If sInactiveProgramYear <> "" Or lSpreadsheetProgramYear < sInactiveProgramYear Then
'                    bValidProgramSubProgram = True
'                Else
'                    If lSpreadsheetProgramYear < 2008 Then
'                        bValidProgramSubProgram = True
'                    Else
'                        bValidProgramSubProgram = False
'                        sSubProgErrMsg = "SubProgram[" & sSpreadsheetSubProgram & "] invalid for ProgramYear: " & lSpreadsheetProgramYear
'                    End If
'                End If
'            End If
'        Else

'            'That SubProgram EXISTS for that Program on ProgramSubProgram table
'            If IsBlank(rs!InactiveSubProgramYear) = True Or lSpreadsheetProgramYear < rs!InactiveSubProgramYear Then
'                bValidProgramSubProgram = True
'            Else
'                sSubProgErrMsg = "SubProgram[" & sSpreadsheetSubProgram & "] Invalid for ProgramYear: " & rs!InactiveSubProgramYear
'                bValidProgramSubProgram = False
'            End If

'        End If

'ExitCloseRS:
'        If rs.State <> 0 Then rs.Close()
'        rs = Nothing
'        If RS2.State <> 0 Then RS2.Close()
'        RS2 = Nothing
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidProgramSubProgram")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidProgramSubProgram , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Private Function bValidProgramGrade(ByVal lProgNumber As Long, ByVal sGrade As String, ByVal sInactiveProgramYear As String, ByVal bGradeRequired As Boolean) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        bValidProgramGrade = False
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset

'        sSQL = "SELECT Grade, InactiveProgramYear AS InactiveGradeYear FROM ProgramGrade WHERE ProgramGrade.ProgramNumber = " & lProgNumber
'        If IsNull(sGrade) Or sGrade = "" Then
'            If bGradeRequired = True Then
'                If lProgramYear < sInactiveProgramYear Then
'                    bValidProgramGrade = True
'                Else
'                    bValidProgramGrade = False
'                End If
'            Else
'                bValidProgramGrade = True
'            End If
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        Else
'            sSQL = sSQL & " AND convert(varchar,ProgramGrade.Grade) = " & sAddQuotes(sGrade)       'new 10/26/2007
'        End If

'Dim         rs.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic As Object

'        If rs.EOF Then
'            If bGradeRequired = True Then
'                bValidProgramGrade = False
'            Else
'                If sInactiveProgramYear <> "" Then
'                    If lProgramYear < CInt(sInactiveProgramYear) Then
'                        bValidProgramGrade = True
'                    Else
'                        bValidProgramGrade = False
'                    End If
'                Else
'                    bValidProgramGrade = False 'new 10/23/2007  ??
'                End If
'            End If

'        Else
'            '            If lProgramYear < sInactiveProgramYear Then
'            If IsNull(rs!InactiveGradeYear) Or lProgramYear < rs!InactiveGradeYear Then
'                bValidProgramGrade = True
'            Else
'                bValidProgramGrade = False
'            End If
'            '            Else
'            '                bValidProgramGrade = False
'            '            End If
'        End If
'        '    End If

'ExitCloseRS:
'        rs.Close()
'        rs = Nothing
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidProgramGrade")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidProgramGrade , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Friend Overridable Function bValidTechnology(ByVal lRow As Long, ByRef sTechnologies As String, ByRef sX_TechDescr As String, ByRef sInvalidTechnologies As String, ByRef lProgramYear As Long) As Boolean        '09/29/2008 - hn - added ProgramYear
'        Dim i As Long
'        Dim sTechCode As String
'        Dim sTechCodes() As String    ' for now up to 8 technology codes

'        Dim SQL As String
'        Dim lDupCounter As Long
'        Dim sDuplicates As String

'        bValidTechnology = False
'        sInvalidTechnologies = ""

'        sTechnologies = Replace(sTechnologies, ",", "")
'        sTechCodes() = Split(sTechnologies, " ")
'        For i = LBound(sTechCodes) To UBound(sTechCodes)
'            Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'            SQL = "SELECT Technology, TechnologyName, InactiveProgramYear FROM Technology WHERE convert(varchar,Technology) = " & sAddQuotes(Trim(sTechCodes(i)))   '11/08/2007
'Dim         rs.Open SQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockOptimistic As Object

'            If rs.EOF Then
'                sInvalidTechnologies = sInvalidTechnologies & Trim(sTechCodes(i)) & " "
'            Else
'                If Not IsBlank(rs!InactiveProgramYear) Then                     '09/29/2008 - hn
'                    If lProgramYear < rs!InactiveProgramYear Then
'                        sX_TechDescr = sX_TechDescr & rs!TechnologyName & ", "
'                    Else
'                        sInvalidTechnologies = sInvalidTechnologies & " (" & rs!Technology & ") Invalid for ProgramYear:" & lProgramYear & vbCrLf  '09/29/2008 - hn
'                        '                    sTechnologies = Replace(sTechnologies, sTechCodes(i), "")          '09/29/2008 - hn
'                    End If
'                Else
'                    sX_TechDescr = sX_TechDescr & rs!TechnologyName & ", "
'                End If
'            End If
'            rs.Close()
'            rs = Nothing

'            'this could happen on importing a sheet:
'            For lDupCounter = 0 To i
'                If Trim(sTechCodes(i)) = Trim(sTechCodes(lDupCounter)) And i <> lDupCounter Then
'                    sDuplicates = sDuplicates & " " & Trim(sTechCodes(i))
'                End If
'            Next lDupCounter

'        Next i

'        If Len(sX_TechDescr) > 2 Then
'            sX_TechDescr = Microsoft.VisualBasic.Left(sX_TechDescr, Len(sX_TechDescr) - 2)
'        End If
'        If sDuplicates <> "" Then
'            sInvalidTechnologies = sInvalidTechnologies & " Duplicate Technology Code Found: " & sDuplicates
'        End If
'        '    If sInvalidTechnologies <> "" Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        bValidTechnology = True
'ExitRoutine:
'        sTechnologies = Replace(Trim(sTechnologies), " ", ", ")
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidTechnology")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidTechnology , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'        Friend Overridable Function bValidCertificationMarkInactiveProgramYear(ByVal sCertificationMark As String, ByVal ByVal lProgramYear As Long, ByRef sCertificationMarkInactiveProgramYear As String) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rs As ADODB.Recordset

'        bValidCertificationMarkInactiveProgramYear = False
'        rs = New ADODB.Recordset
'        sSQL = "SELECT InactiveProgramYear FROM CertificationMark WHERE CertificationMark = '" & sCertificationMark & "'"
'Dim     rs.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic As Object

'        If rs.EOF Then
'            If sCertificationMark = "" Then
'                bValidCertificationMarkInactiveProgramYear = True
'            End If
'        Else
'            If rs!InactiveProgramYear <> "" Then
'                If lProgramYear < rs!InactiveProgramYear Then
'                    bValidCertificationMarkInactiveProgramYear = True
'                Else
'                    bValidCertificationMarkInactiveProgramYear = False
'                    sCertificationMarkInactiveProgramYear = rs!InactiveProgramYear
'                End If
'            Else
'                bValidCertificationMarkInactiveProgramYear = True
'            End If
'        End If

'ExitRoutine:
'        If rs.State <> 0 Then rs.Close()
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidCertificationMarkInactiveProgramYear")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidCertificationMarkInactiveProgramYear , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine


'    End Function

'        Friend Overridable Function bValidCertificationTypeInactiveProgramYear(ByVal sCertificationType As String, ByVal ByVal lProgramYear As Long, ByRef sCertificationTypeInactiveProgramYear As String) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sSQL As String
'        Dim rs As ADODB.Recordset

'        bValidCertificationTypeInactiveProgramYear = False
'        rs = New ADODB.Recordset
'        sSQL = "SELECT InactiveProgramYear FROM CertificationType WHERE CertificationType = '" & sCertificationType & "'"
'Dim     rs.Open sSQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic As Object

'        If rs.EOF Then
'            If sCertificationType = "" Then
'                bValidCertificationTypeInactiveProgramYear = True
'            End If
'        Else
'            If rs!InactiveProgramYear <> "" Then
'                If lProgramYear < rs!InactiveProgramYear Then
'                    bValidCertificationTypeInactiveProgramYear = True
'                Else
'                    bValidCertificationTypeInactiveProgramYear = False
'                    sCertificationTypeInactiveProgramYear = rs!InactiveProgramYear
'                End If
'            Else
'                bValidCertificationTypeInactiveProgramYear = True
'            End If
'        End If

'ExitRoutine:
'        If rs.State <> 0 Then rs.Close()
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidCertificationTypeInactiveProgramYear")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidCertificationTypeInactiveProgramYear , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine


'    End Function

'    Friend Overridable Function bInsertX_Technology(ByVal objExcel As Excel.Application, ByVal lMaxRows As Long) As Boolean
'        Dim sCellValue As String
'        Dim lColCounter As Long
'        Dim lRowCounter As Long

'        Dim lTechnologiesColPos As Long
'        Dim lX_TechnologiesCOLPos As Long

'        Dim lProgramYearCOLPos As Long
'        Dim lProgramYear As Long                         '09/29/2008 - hn

'        Dim sTechnologies As String
'        Dim sX_Technologies As String
'        Dim sInvalidTechnologies As String

'        Dim ws As Excel.Worksheet ' Set a reference to this excel worksheet.

'        bInsertX_Technology = False
'        ws = objExcel.Application.Workbooks(1).Worksheets(1)
'        'first determine the column positions needed
'        For lColCounter = 1 To glMAX_Cols
'            '        If gbCancelExport = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            sCellValue = sGetCellValue(ws.Cells(1, lColCounter))
'            If sCellValue = gsCOL_Technologies Then lTechnologiesColPos = lColCounter
'            If sCellValue = gsCOL_X_Technologies Then lX_TechnologiesCOLPos = lColCounter
'            If sCellValue = gsCOL_ProgramYear Then lProgramYearCOLPos = lColCounter '09/29/2008 - hn

'            If lTechnologiesColPos > 0 And lX_TechnologiesCOLPos > 0 Then
'                Exit For
'            End If
'        Next lColCounter

'        If lTechnologiesColPos = 0 Or lX_TechnologiesCOLPos = 0 Then
'        Else
'            ' now insert X_Technologies for each row
'            For lRowCounter = glDATA_START_ROW To lMaxRows
'                sTechnologies = sGetCellValue(ws.Cells(lRowCounter, lTechnologiesColPos))
'                If IsBlank(sGetCellValue(ws.Cells(lRowCounter, lProgramYearCOLPos))) Then
'                    '03/31/2009 - hn -could indicate last row if export query is HK Vendor Grid and not all selected Items are returned for a Category!
'                Else
'                    lProgramYear = sGetCellValue(ws.Cells(lRowCounter, lProgramYearCOLPos))          '09/29/2008 - hn
'                    sX_Technologies = ""
'                    If sTechnologies <> "" Then
'                        If bValidTechnology(1, sTechnologies, sX_Technologies, sInvalidTechnologies, lProgramYear) = True Then   '09/29/2008 - hn
'                            ws.Cells(lRowCounter, lX_TechnologiesCOLPos) = sX_Technologies
'                        Else
'                            ws.Cells(lRowCounter, lX_TechnologiesCOLPos) = sX_Technologies & "Invalid Code(s): " & sInvalidTechnologies
'                        End If
'                    Else
'                        ws.Cells(lRowCounter, lX_TechnologiesCOLPos) = ""
'                    End If
'                End If
'            Next lRowCounter
'            bInsertX_Technology = True
'        End If

'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bInsertX_Technology")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bInsertX_Technology , Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume Next

'Dim Public Function bInsertX_Material(objExcel As Excel.Application
'Dim  ByVal lMaxRows As Long
'Dim  ByVal MaxMaterialColumns As Long) As Boolean

'    End Function

'    Friend Overridable Function bInsertX_Material(ByVal objExcel As Excel.Application, ByVal lMaxRows As Long, ByVal MaxMaterialColumns As Long) As Boolean
'        Dim sCellValue As String
'        Dim lColCounter As Long
'        Dim lRowCounter As Long

'        Dim lProposalCOLPos As Long
'        Dim lProposalNumber As Long

'        Dim lRevCOLPos As Long
'        Dim lRev As Long

'        Dim lMaterialXCOLPos() As Long
'        Dim lMaterialX As Long
'        Dim sMaterial As String
'        Dim sMaterialHeading As String

'        Dim ws As Excel.Worksheet ' Set a reference to this excel worksheet.
'        Dim SQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        bInsertX_Material = False
'    ReDim lMaterialXCOLPos(1 To MaxMaterialColumns)
'        ws = objExcel.Application.Workbooks(1).Worksheets(1)
'        'first determine the column positions needed
'        For lColCounter = 1 To glMAX_Cols
'            '        If gbCancelExport = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            sCellValue = sGetCellValue(ws.Cells(1, lColCounter))
'            If sCellValue = "ProposalNumber" Then lProposalCOLPos = lColCounter
'            If sCellValue = "Rev" Then lRevCOLPos = lColCounter

'            If Microsoft.VisualBasic.Left(sCellValue, 8) = "Material" And Len(sCellValue) > 8 And sCellValue <> "MaterialType" Then
'                sMaterialHeading = sCellValue
'                sMaterialHeading = Replace(sMaterialHeading, "Material", "")
'                If IsNumeric(sMaterialHeading) Then
'                    lMaxMaterialCols = lMaxMaterialCols + 1
'                    lMaterialX = CLng(sMaterialHeading)
'                    lMaterialXCOLPos(lMaterialX) = lColCounter
'                End If
'            End If
'        Next lColCounter

'        ' now insert concatenated Material, breakdown, cost for each row
'        For lRowCounter = glDATA_START_ROW To lMaxRows
'            If IsBlank(sGetCellValue(ws.Cells(lRowCounter, lProposalCOLPos))) = False Then '2013/05/01 -HN
'                lProposalNumber = sGetCellValue(ws.Cells(lRowCounter, lProposalCOLPos))
'                lRev = sGetCellValue(ws.Cells(lRowCounter, lRevCOLPos))
'                sMaterial = ""
'                ' first line below will still return all values even if one is null
'                SQL = "SET CONCAT_NULL_YIELDS_NULL OFF " & _
'                      "SELECT  Convert(varchar,MaterialPct) + ', ' +  MaterialName + ', ' + convert(varchar,CostPct) AS MaterialX  " & _
'                      "FROM ItemMaterial " & _
'                      "WHERE ProposalNumber = " & lProposalNumber & " AND Rev = " & lRev & _
'                      " ORDER BY MaterialPct DESC, MaterialName"
'Dim        rs.Open SQL As Object 
'                Dim SSDataConn As Object
'                Dim adOpenStatic As Object
'                Dim adLockReadOnly As Object

'                lMaterialX = 1
'                Do Until rs.EOF
'                    ws.Cells(lRowCounter, lMaterialXCOLPos(lMaterialX)) = rs!MaterialX
'                    lMaterialX = lMaterialX + 1
'                    rs.MoveNext()
'                Loop
'                rs.Close()
'            End If '2013/05/01 -HN
'        Next lRowCounter

'        bInsertX_Material = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bInsertX_Material: Error on Row: " & lRowCounter)   '2013/05/01
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bInsertX_Material: Error on Row: " & lRowCounter & ", Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume Next

'    End Function

'        Private Function bInsertRowItemMaterial(ByVal lRowCounter As Long, ByVal sFunctioncode As String, ByVal ByVal lProposalNumber As Long, ByVal lPreviousRev As Long, ByVal lNewRev As Long, ByVal ByVal sItemNumber As String) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            If bWritePrintToLogFile(False, objEXCELName & Space(8) & "Inserting Material for, for Row: " & lRowCounter & " and Proposalnumber: " & lProposalNumber & " and REV:" & lNewRev, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If

'        Dim SQL As String
'        bInsertRowItemMaterial = False
'        SQL = "INSERT INTO ItemMaterial(ProposalNumber, Rev, MaterialName, MaterialPct, CostPct) " & _
'              "SELECT  ProposalNumber, " & _
'              lNewRev & _
'              ", MaterialName, MaterialPct, CostPct FROM ItemMaterial " & _
'              "WHERE ProposalNumber = " & lProposalNumber & " AND  Rev = " & lPreviousRev

'        SSDataConn.Execute SQL
'        Application.DoEvents()
'        bInsertRowItemMaterial = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bInsertRowItemMaterial")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bInsertRowItemMaterial ,Error Number:" & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine


'        '2010/02/11 - added lRowChangesFound
'Private Function bUpdateRowItemMaterial(ByVal lRowCounter As Long, ByVal sFunctioncode As String, _
'                            ByVal lProposalNumber As Long, ByVal lOriginalRev As Long, ByVal lNewOrChangedREV As Long, _
'        End Function

'        Private Function bUpdateRowItemMaterial(ByVal lRowCounter As Long, ByVal sFunctioncode As String, ByVal ByVal lProposalNumber As Long, ByVal lOriginalRev As Long, ByVal lNewOrChangedREV As Long, ByVal ByVal sItemNumber As String, ByRef lRowChangesFound As Long) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lMaterialCounter As Long
'        Dim rs As ADODB.Recordset
'        Dim RSMaterialName As String
'        Dim RSMaterialPct As String
'        Dim RSCostPct As String

'        Dim sSpreadsheetMatPct As Object
'        Dim sSpreadsheetCostPct As Object

'        Dim SQL As String
'        Dim deletesql As String

'        Dim sOldMaterialValues As String
'        Dim sNewMaterialValues As String

'        bUpdateRowItemMaterial = False
'        'first update existing ItemMaterial records with same Materialnames
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            If bWritePrintToLogFile(False, objEXCELName & Space(6) & "Starting bUpdateRowItemMaterial, for Row: " & lRowCounter & " , FunctionCode: " & sFunctioncode & ", ProposalNumber: " & lProposalNumber, Format(Now, "yyyymmdd")) = False Then
'            End If
'        End If
'        For lMaterialCounter = 1 To lMaxMaterialCols
'            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sChangeIndicator = ""

'            sSpreadsheetMatPct = SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialPct
'            sSpreadsheetCostPct = SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sCostPct
'            If sSpreadsheetMatPct <> "" Then
'                sSpreadsheetMatPct = CDec(sSpreadsheetMatPct)
'            Else
'                sSpreadsheetMatPct = "0"
'            End If
'            If sSpreadsheetCostPct <> "" Then
'                sSpreadsheetCostPct = CDec(sSpreadsheetCostPct)
'            Else
'                sSpreadsheetCostPct = "0"
'            End If

'            rs = New ADODB.Recordset
'            If SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialName <> "" Then
'                SQL = "SELECT * FROM ItemMaterial " & _
'                        "WHERE ProposalNumber = " & lProposalNumber & _
'                        " AND Rev = " & lNewOrChangedREV & _
'                        " AND MaterialName = '" & Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialName) & "'"

'Dim             rs.Open SQL As Object 
'                Dim SSDataConn As Object
'                Dim adOpenStatic As Object
'                Dim adLockOptimistic As Object

'                Application.DoEvents()                                                    '03/10/2009 - hn - cursor conflict?
'                If Not rs.EOF Then
'                    If IsNull(rs!MaterialName) Then 'to take care of old bad data
'                        RSMaterialName = ""
'                    Else
'                        RSMaterialName = rs!MaterialName
'                    End If
'                    If IsNull(rs!MaterialPct) Then
'                        RSMaterialPct = ""
'                    Else
'                        RSMaterialPct = rs!MaterialPct
'                    End If
'                    If IsNull(rs!CostPct) Then
'                        RSCostPct = ""
'                    Else
'                        RSCostPct = rs!CostPct
'                    End If
'                    If Trim(RSMaterialName) = Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialName) Then

'                        If RSMaterialPct = sSpreadsheetMatPct And RSCostPct = sSpreadsheetCostPct Then
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sChangeIndicator = ""
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sItemMaterialOldValue = CDec(RSMaterialPct) & ", " & _
'                                Trim(rs!MaterialName) & ", " & CDec(RSCostPct)
'                        Else
'                            sOldMaterialValues = sOldMaterialValues & Trim(RSMaterialPct) & ", " & Trim(RSMaterialName) & ", " & Trim(RSCostPct) & "; "
'                            sNewMaterialValues = sNewMaterialValues & Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialPct) & ", " & _
'                                                Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialName) & ", " & _
'                                                Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sCostPct) & "; "
'                            rs!MaterialPct = sSpreadsheetMatPct
'                            rs!CostPct = sSpreadsheetCostPct
'                            rs.Update()
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sChangeIndicator = "Y"
'                            SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sItemMaterialOldValue = Trim(RSMaterialPct) & ", " & _
'                                Trim(RSMaterialName) & ", " & Trim(RSCostPct)
'                        End If
'                    End If
'                Else
'                    '2014/01/14 RAS Adding info to Trace message
'                    If glTraceFlag = True Then
'                        If bWritePrintToLogFile(False, objEXCELName & Space(10) & "Inserting into ITEMMATERIAL in bUpdateRowItemMaterial, for Row: " & lRowCounter & ", ProposalNumber: " & lProposalNumber & " ,Rev: " & lNewOrChangedREV, Format(Now(), "yyyymmdd")) = False Then
'                        End If
'                    End If
'                    'insert new record for that MaterialName
'                    If rs.State <> 0 Then rs.Close()
'                    SQL = "INSERT INTO ItemMaterial(ProposalNumber, Rev, MaterialName, MaterialPct, CostPct) " & _
'                        "VALUES(" & lProposalNumber & _
'                        ", " & lNewOrChangedREV & _
'                        ", '" & Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialName) & "'" & _
'                        ", " & sSpreadsheetMatPct & _
'                        ", " & sSpreadsheetCostPct & ")"
'Dim                 rs.Open SQL As Object 
'                    Dim SSDataConn As Object
'                    Dim adOpenStatic As Object
'                    Dim adLockOptimistic As Object

'                    Application.DoEvents()                                                    '03/10/2009 - hn - cursor conflict?
'                    sNewMaterialValues = sNewMaterialValues & Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialPct) & ", " & _
'                                   Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialName) & ", " & _
'                                   Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sCostPct) & "; "
'                    SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sChangeIndicator = "Y"
'                    SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sItemMaterialOldValue = "NewMaterial" & lMaterialCounter
'                End If
'                If rs.State <> 0 Then
'                    rs.Close()
'                End If

'            End If
'        Next lMaterialCounter
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            If bWritePrintToLogFile(False, objEXCELName & Space(6) & "Delete ItemMaterial records which dont have same MaterialNames as on spreadsheet, for Row: " & lRowCounter & ", ProposalNumber: " & lProposalNumber & " ,Rev: " & lNewOrChangedREV, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If

'        'now delete ItemMaterial records which dont have same MaterialNames as on spreadsheet
'        deletesql = " AND MaterialName NOT IN ("

'        SQL = "SELECT * FROM ItemMaterial WHERE ProposalNumber = " & lProposalNumber & _
'             " AND Rev = " & lNewOrChangedREV
'        For lMaterialCounter = 1 To lMaxMaterialCols
'            '        If SpreadsheetMaterialValuesX(lROWCounter, lMaterialCounter).sMaterialName <> "" Then
'            deletesql = deletesql & "'" & Trim(SpreadsheetMaterialValuesX(lRowCounter, lMaterialCounter).sMaterialName) & "',"
'            '        End If
'        Next lMaterialCounter

'        If deletesql <> "" Then
'            deletesql = Microsoft.VisualBasic.Left(deletesql, Len(deletesql) - 1) & ") "  'remove last comma, add closing parenthesis
'            SQL = SQL & deletesql
'Dim         rs.Open SQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockOptimistic As Object

'            If Not rs.EOF Then
'                Do Until rs.EOF
'                    If IsNull(rs!MaterialName) Then 'to take care of old bad data
'                        RSMaterialName = ""
'                    Else
'                        RSMaterialName = rs!MaterialName
'                    End If
'                    If IsNull(rs!MaterialPct) Then
'                        RSMaterialPct = ""
'                    Else
'                        RSMaterialPct = rs!MaterialPct
'                    End If
'                    If IsNull(rs!CostPct) Then
'                        RSCostPct = ""
'                    Else
'                        RSCostPct = rs!CostPct
'                    End If
'                    sOldMaterialValues = sOldMaterialValues & Trim(RSMaterialPct) & ", " & Trim(RSMaterialName) & ", " & Trim(RSCostPct) & "; "

'                    rs.MoveNext()

'                Loop
'                '            RS.MoveFirst' just dsnt work
'                '            RS.Delete
'                '            RS.UpdateBatch
'                SQL = Replace(SQL, "SELECT *", "DELETE")
'                SSDataConn.Execute SQL
'                Application.DoEvents()
'            End If

'            If rs.State <> 0 Then
'                rs.Close()
'            End If
'        End If
'        If Len(sOldMaterialValues) > 0 Then
'            sOldMaterialValues = Microsoft.VisualBasic.Left(sOldMaterialValues, Len(sOldMaterialValues) - 2)
'        Else
'            sOldMaterialValues = ""
'        End If
'        If Len(sNewMaterialValues) > 0 Then
'            sNewMaterialValues = Microsoft.VisualBasic.Left(sNewMaterialValues, Len(sNewMaterialValues) - 2)
'        Else
'            sNewMaterialValues = ""
'        End If
'        If sFunctioncode <> gsNEW_ITEM_NBR Then
'            If sOldMaterialValues <> sNewMaterialValues Then
'                '2014/01/14 RAS Adding info to Trace message
'                If glTraceFlag = True Then
'                    If bWritePrintToLogFile(False, objEXCELName & Space(12) & "update Item.RevisedDate, Item.RevisedUserid if not yet done, for Row: " & lRowCounter & ", ProposalNumber: " & lProposalNumber & " ,Rev: " & lNewOrChangedREV, Format(Now(), "yyyymmdd")) = False Then
'                    End If
'                End If
'                '2010/02/11 - update Item.RevisedDate, Item.RevisedUserid if not yet done
'                If lRowChangesFound < 1 Then
'                    SQL = "UPDATE Item SET RevisedUserId = '" & gsUserID & "' , RevisedDate = GetDate()" & vbCrLf & _
'                          "WHERE ProposalNumber = " & lProposalNumber & " And Rev = " & lNewOrChangedREV
'                    SSDataConn.Execute SQL
'                End If
'                lRowChangesFound = lRowChangesFound + 1
'                If bUpdateItemMaterialHistory(lProposalNumber, lOriginalRev, lNewOrChangedREV, _
'                            sItemNumber, "I", _
'                            sOldMaterialValues, sNewMaterialValues) = False Then
'                    '                GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If
'            End If
'        End If
'        bUpdateRowItemMaterial = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & "Error Updating ItemMaterial table for Row: " & lRowCounter & vbCrLf _
'        & " Check log file after Import completed!", vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bUpdateRowItemMaterial")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bUpdateRowItemMaterial, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'Dim Public Function bUpdateItemMaterialHistory(ProposalNumber As Object 
'        Dim ChangedRev As Object
'        Dim NewRev As Object
'        Dim ItemNumber As Object
'Dim  _ As Object 

'                            sFromProcess, sOldMaterialValues, sNewMaterialValues) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'    End Function

'    Friend Overridable Function bUpdateItemMaterialHistory(ByVal ProposalNumber As Object, ByVal ChangedRev As Object, ByVal NewRev As Object, ByVal ItemNumber As Object, ByVal sFromProcess As Object, ByVal sOldMaterialValues As Object, ByVal sNewMaterialValues As Object) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        '11/26/2007 - no need to write for new items
'        If bNewItemFromProposal = True Or bNewCustomerFromProposal = True Or bProposalNEWItemfromList = True Then
'            bUpdateItemMaterialHistory = True
'            '            GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            If bWritePrintToLogFile(False, objEXCELName & Space(14) & " bUpdateItemMaterialHistory for ProposalNumber: " & ProposalNumber & " ,Rev: " & NewRev, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        bUpdateItemMaterialHistory = False
'        If sOldMaterialValues <> sNewMaterialValues Then
'            Dim SQL As String
'            Dim sProcess As String


'            If ChangedRev <> NewRev Then
'                sProcess = "R"
'            Else
'                sProcess = "C"
'            End If
'            Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'            SQL = "SELECT * FROM ItemFieldHistory WHERE ProposalNumber = " & ProposalNumber & " AND Rev  = " & NewRev
'            rs.Open(SQL, SSDataConn, adOpenDynamic, adLockPessimistic)
'            '2010/02/11 - added lRowChangesFound=0
'            If bSAVEItemFieldsChanged(0, rs, ProposalNumber, NewRev, ChangedRev, ItemNumber, sFromProcess, _
'                                        sProcess, "MaterialValues", sOldMaterialValues, sNewMaterialValues, _
'                                        Now(), gsUserID) = False Then
'                '            GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'            End If
'            rs.Close()
'            rs = Nothing
'        End If
'        bUpdateItemMaterialHistory = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & vbExclamation + vbMsgBoxSetForeground, "modSpreadhseet: UpdateItemMaterialHistory")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In UpdateItemMaterialHistory, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'    Private Function bCheckFactoryProhibitedRegisteredForItemStatus(ByVal sFactoryNumber As Object, ByVal sCustomerNumber As Object, ByVal sProgramYR As Object, ByVal sValidationErrorMsg As Object) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim SQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        bCheckFactoryProhibitedRegisteredForItemStatus = False
'        sValidationErrorMsg = ""
'        If sFactoryNumber = "" Then
'            bCheckFactoryProhibitedRegisteredForItemStatus = True
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If

'        SQL = "SELECT * FROM Factory WHERE FactoryNumber = " & sFactoryNumber
'     rs.Open SQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object


'        If Not rs.EOF Then
'            If rs!InActive = True Or rs!Prohibitive = True Then
'                sValidationErrorMsg = "-Invalid if Factory[" & sFactoryNumber & "] is Inactive/Prohibited!"
'            End If
'        End If
'        If rs.State <> 0 Then rs.Close()

'        '2010/11/23 - run against view now, to simplify
'        SQL = "SELECT top 1 * FROM vAllFactoryCustomerRegistrations" & vbCrLf & _
'                "WHERE CustomerNumber = 100 AND FactoryNumber = " & sFactoryNumber & _
'                " AND ProgramYear = " & sProgramYR & vbCrLf & _
'                "ORDER BY RevisedDate DESC"         '2012/11/13 -hn- there were 2 history recs one Approved, one Pending

'        If CreateSSDataPassthroughQuery("pqFactoryCustomerRegistrations", SQL) = False Then
'            '            GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'        End If

'        '    Set RS = New ADODB.Recordset
'        SQL = "SELECT * FROM pqFactoryCustomerRegistrations"


'        Dim RSReg As DAO.Recordset
'        myDSRSReg = New DataSet()
'        myDSRSReg = GetDataSet(SQL, "myTable")
'        '    RS.Open SQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object

'        '    If RSReg.EOF Then GoTo FactRegError'TODO - GoTo Statements are redundant in .NET
'        '    If myDRRSReg("RegStatus") = "Approved" Then GoTo FactRegOK'TODO - GoTo Statements are redundant in .NET

'FactRegError:
'        If sValidationErrorMsg <> "" Then
'            sValidationErrorMsg = sValidationErrorMsg & vbCrLf
'        End If

'        ' Changed error message per Theresa ~Christa C 4/18/08
'        ' sValidationErrorMsg = sValidationErrorMsg & "-Invalid if Factory[" & sFactoryNumber & "] NOT registered for Customer=100,ProgramYear=" & sProgramYR & ",RegStatus='Approved'"
'        sValidationErrorMsg = sValidationErrorMsg & "-Factory[" & sFactoryNumber & "] was not  registered by Seasonal for ProgramYear " & sProgramYR & ", and RegStatus='Approved' (also not found in History Recs)"        '2010/11/22
'        '        End If          'no error

'        '    End If

'FactRegOK:
'        '    If sValidationErrorMsg <> "" Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET

'        bCheckFactoryProhibitedRegisteredForItemStatus = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rs.State <> 0 Then rs.Close()
'        rs = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bCheckFactoryProhibitedRegisteredForItemStatus")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCheckFactoryProhibitedRegisteredForItemStatus, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Public Function bUpdateItemStatusETC(sItemArray() As String, lRowCounter As Long, _
'                        sProgYear As String, sItemStatus As String, sCustomerItemNum As String, sCustomerCartonUPCNumber As String, _
'                        sTempItemNum As String, ByVal sCategory As String, ByVal sFactory As String, _
'                        ByVal sSeason As String, ByVal sProgram As String, ByVal sSubProgram As String, ByVal sGrade As String, ByVal sClass As String, frmThis As Form, _
'                        ByRef sDescriptiveMsg1 As String, _
'                        ByRef sLatestItemStatus As String) As Boolean
'        '11/24/2008 - hn - added sClass above
'        'On Error GoTo ErrorHandler:'TODO - On Error must be replaced with Try, Catch, Finally

'        Dim rs As ADODB.Recordset
'        Dim sLatestSeason As String
'        Dim sLatestCategory As String
'        Dim sLatestProgram As String
'        Dim sLatestTempItemNum As String

'        Dim sLatestCustItemNum As String
'        Dim sLatestSubProgram As String
'        Dim sLatestGrade As String
'        Dim sLatestFactory As String

'        Dim sLatestClass As String                    '11/24/2008 - hn - new
'        Dim sLatestCustomerCartonUPCNumber As String            '2012/05/02
'        Dim SQL As String
'        Dim sSpreadsheetProposal As String
'        Dim sSpreadsheetRev As String

'        Dim lMsgResponse As Long
'        Dim bSetToLatestRev As Boolean
'        Dim bErrorCount As Integer  '2014/01/21 RAS adding counting integer


'        bUpdateItemStatusETC = False
'        sSpreadsheetProposal = sItemArray(lRowCounter, glProposal_ColPos)
'        sSpreadsheetRev = sItemArray(lRowCounter, glREV_ColPos)
'        sDescriptiveMsg1 = ""

'        If sItemArray(lRowCounter, glFunctionCode_ColPos) = gsNEW_PROPOSAL Then
'            ' for fc = a , already went thru validation, will be only 1 record, so we dont have to proceed
'            bUpdateItemStatusETC = True
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If
'StartAllOver:
'        rs = New ADODB.Recordset
'        '2014/01/14 RAS Adding info to Trace message
'        If glTraceFlag = True Then
'            If bWritePrintToLogFile(False, objEXCELName & Space(7) & "Updating Item Status ETC ,bUpdateItemStatusETC, for Proposalnumber : " & sSpreadsheetProposal & " and Programyear: " & sProgYear, Format(Now(), "yyyymmdd")) = False Then
'            End If
'        End If
'        '11/24/2008 - hn - added Item.Class below..
'        '2012/05/02 - added CustomerCartonUPCNumber
'        SQL = "SELECT Item.ProposalNumber, Item.Rev, " & _
'                "Item.TempItemNumber, Item.ItemStatus, " & _
'                "Item.CustomerItemNumber, Item.CustomerCartonUPCNumber, Item.ProgramYear, " & _
'                "Item.SeasonCode, Item.CategoryCode, Item.ProgramNumber, " & _
'                "Item.SubProgram, Item.Grade, Item.Class, Item.FactoryNumber " & _
'                "FROM Item " & _
'                "WHERE Item.ProposalNumber = " & sSpreadsheetProposal & _
'                " AND Item.ProgramYear = " & sProgYear & _
'                " ORDER BY Item.Rev DESC"
'Dim     rs.Open SQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object


'        If rs.EOF = True Then
'            bUpdateItemStatusETC = True
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If

'        If rs.Recordcount < 2 Then
'            'no need to update revs if there is only one for that programyr
'            bUpdateItemStatusETC = True
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If

'        rs.MoveFirst()
'        If Not IsBlank(rs![TempItemNumber]) Then
'            sLatestTempItemNum = rs![TempItemNumber]
'        End If
'        If Not IsBlank(rs![ItemStatus]) Then
'            sLatestItemStatus = rs![ItemStatus]
'        End If
'        If Not IsBlank(rs![CustomerItemNumber]) Then
'            sLatestCustItemNum = rs![CustomerItemNumber]
'        End If
'        If Not IsBlank(rs![CustomerCartonUPCNumber]) Then            '2012/05/02
'            sLatestCustomerCartonUPCNumber = rs![CustomerCartonUPCNumber]
'        End If
'        If Not IsBlank(rs![SeasonCode]) Then
'            sLatestSeason = rs![SeasonCode]
'        End If
'        If Not IsBlank(rs![CategoryCode]) Then
'            sLatestCategory = rs![CategoryCode]
'        End If
'        If Not IsBlank(rs![ProgramNumber]) Then
'            sLatestProgram = rs![ProgramNumber]
'        End If
'        If Not IsBlank(rs![SubProgram]) Then
'            sLatestSubProgram = rs![SubProgram]
'        Else
'            sLatestSubProgram = ""                      '11/21/2008 - hn
'        End If
'        If Not IsBlank(rs![Grade]) Then
'            sLatestGrade = rs![Grade]
'        Else
'            sLatestGrade = ""
'        End If
'        If Not IsBlank(rs![Class]) Then                 '11/24/2008 - hn
'            sLatestClass = rs![Class]
'        Else
'            sLatestClass = ""
'        End If
'        If Not IsBlank(rs![FactoryNumber]) Then
'            sLatestFactory = rs![FactoryNumber]
'        End If


'        If sSpreadsheetRev = rs![Rev] Then
'            'UPDATE Previous REVS WITH LATEST REV'S VALUES if different
'            '11/24/2008 - hn - added sClass below..
'            '2012/05/02 - added sLatestCustomerCartonUPCNumber
'            If bUpdateRevValues(sSpreadsheetProposal, sSpreadsheetRev, sLatestCustItemNum, sLatestCustomerCartonUPCNumber, sLatestSeason, sLatestCategory, _
'                sLatestProgram, sLatestSubProgram, sLatestGrade, sProgYear, sLatestClass, sDescriptiveMsg1) = False Then
'                'do error msg
'                Application.DoEvents()
'            End If
'            If bUpdateItemStatus(sSpreadsheetProposal, sSpreadsheetRev, sLatestItemStatus, sProgYear, sLatestFactory, sDescriptiveMsg1) = False Then
'                'do error msg
'            End If
'            If sLatestTempItemNum <> "" Then
'                If bUpdateTempItem(sSpreadsheetProposal, sSpreadsheetRev, sProgYear, sLatestTempItemNum, sDescriptiveMsg1) = False Then
'                    'do error msg
'                End If
'            End If
'            bUpdateItemStatusETC = True
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If

'        'Check for differences with earlier revs
'        'Check for differences with earlier revs
'        '2012/05/02 - added sLatestCustomerCartonUPCNumber
'        If sItemStatus <> sLatestItemStatus Or _
'            sCustomerItemNum <> sLatestCustItemNum Or _
'            sCustomerCartonUPCNumber <> sLatestCustomerCartonUPCNumber Or _
'            sTempItemNum <> sLatestTempItemNum Or _
'            sCategory <> sLatestCategory Or _
'            sSeason <> sLatestSeason Or _
'            sProgram <> sLatestProgram Or _
'            sClass <> sLatestClass Then                     '11/24/2008 - hn -added class comparison
'            If bPROPOSALFormIndicator = True Then
'                '11/24/2008 - hn - added Class below..
'                lMsgResponse = MsgBox("Do you want to change the Latest  Rev's(" & rs![Rev] & "), and earlier Rev's values for " & vbCrLf & vbCrLf & _
'                       "ItemStatus/CustomerItem#/Category/Season/Program/SubProgram/Grade/Class " & vbCrLf & vbCrLf & _
'                        " for PROGRAM YEAR: " & sProgYear & vbCrLf & vbCrLf & _
'                       " to the values of this Revision(" & sSpreadsheetRev & ") ?" & vbCrLf & vbCrLf & _
'                       "'Yes' will set all Rev's values to this Rev's values" & vbCrLf & vbCrLf & _
'                       "'No' will set all Rev's values to the Latest Rev's values, for THIS PROGRAM YEAR.", vbQuestion + vbYesNo + vbMsgBoxSetForeground, _
'                       "Change ItemStatus/CustItem#/Category/Season/Program values for All Revs?")
'                '11/13/2008 - hn
'                If lMsgResponse = vbYes Then
'                    bSetToLatestRev = False
'                ElseIf lMsgResponse = vbNo Then
'                    bSetToLatestRev = True
'                Else
'                    bUpdateItemStatusETC = True
'                    '                    GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'                End If
'            Else
'                sDescriptiveMsg1 = sDescriptiveMsg1 & "This Rev has later Rev's with different values for ItemStatus/CustomerItem#/TempItem#/Category/Season/Program/SubProgram/Grade/Class" '11/24/2008 - hn - added class
'                bSetToLatestRev = False
'            End If
'        Else
'            bSetToLatestRev = True
'        End If

'        If bSetToLatestRev = False Then
'            'set to values of earlier rev(sSpreadsheetRev) for ProgramYear that has changed

'            '11/24/2008 - hn - added sClass below..
'            '2012/05/02 - added sLatestCustomerCartonUPCNumber
'            If bUpdateRevValues(sSpreadsheetProposal, sSpreadsheetRev, sCustomerItemNum, sLatestCustomerCartonUPCNumber, sSeason, sCategory, _
'                sProgram, sSubProgram, sGrade, sProgYear, sClass, sDescriptiveMsg1) = False Then
'                'do error msg
'            End If
'            If bUpdateItemStatus(sSpreadsheetProposal, sSpreadsheetRev, sItemStatus, sProgYear, sFactory, sDescriptiveMsg1) = False Then
'                'do error msg
'            End If
'            If sTempItemNum <> "" Then
'                If bUpdateTempItem(sSpreadsheetProposal, sSpreadsheetRev, sProgYear, sTempItemNum, sDescriptiveMsg1) = False Then
'                    'do error msg
'                End If
'            End If
'            bUpdateItemStatusETC = True
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET

'        Else
'            'set to values of latest rev (= rs!rev) for ProgramYear
'            '11/24/2008 - hn - added sClass below..
'            '2012/05/02 - added sLatestCustomerCartonUPCNumber
'            If bUpdateRevValues(sSpreadsheetProposal, sSpreadsheetRev, sCustomerItemNum, sLatestCustomerCartonUPCNumber, sSeason, sCategory, _
'                sProgram, sSubProgram, sGrade, sProgYear, sClass, sDescriptiveMsg1) = False Then
'                'do error msg
'            End If
'            If bUpdateItemStatus(sSpreadsheetProposal, sSpreadsheetRev, sLatestItemStatus, sProgYear, sLatestFactory, sDescriptiveMsg1) = False Then
'                'do error msg
'            End If
'            If sLatestTempItemNum <> "" Then
'                If bUpdateTempItem(sSpreadsheetProposal, sSpreadsheetRev, sProgYear, sLatestTempItemNum, sDescriptiveMsg1) = False Then
'                    'do error msg
'                End If
'            End If
'            bUpdateItemStatusETC = True
'            '        GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        End If

'        bUpdateItemStatusETC = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        If rs.State <> 0 Then rs.Close()
'        rs = Nothing
'        Exit Function
'ErrorHandler:
'        'Resume Next   '' testing only
'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bUpdateItemStatusETC")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bUpdateItemStatusETC, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function
'    Private Function bUpdateRevValues(ByVal lProposal As Long, ByVal lRev As Long, _
'                    ByVal sCustItemNum As String, ByVal sCustomerCartonUPCNumber As String, ByVal sSeason As String, ByVal sCategory As String, _
'                    ByVal sProgram As String, ByVal sSubProgram As String, ByVal sGrade As String, _
'                    ByVal sProgYear As String, ByVal sClass As String, _
'                    ByRef sMsg As String) As Boolean

'        '11/24/2008 - added sClass
'        '2012/05/02 - added sCustomerCartonUPCNumber
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim SQL As String
'        Dim sSubProgramSQL As String
'        Dim sGradeSQL As String

'        Dim sClassSQL As String
'        Dim SClassWhereSQL As String                                          '11/24/2008 - hn

'        Dim lRecordsAffected As Long                                                                     '11/24/2008 - hn
'        Dim bErrorCount As Integer

'StartAllOver:
'        bUpdateRevValues = False
'        ' to take care of SubProgram and Grade values before 2008, when they could be blank
'        ' and to prevent SQL error ; foreign key constraint on Grade tabel - FK_Item_Grade
'        If sSubProgram = "" Then
'            sSubProgramSQL = ", SubProgram = NULL"
'        Else
'            sSubProgramSQL = ", SubProgram = '" & sSubProgram & "'"
'        End If
'        If sGrade = "" Then
'            sGradeSQL = ", Grade = NULL"
'        Else
'            sGradeSQL = ", Grade = '" & sGrade & "'"
'        End If

'        If sClass = "" Then                                                                         '11/24/2008 - hn
'            sClassSQL = ", Class = NULL"
'            SClassWhereSQL = ""
'        Else
'            sClassSQL = ", Class = '" & sClass & "'"
'            SClassWhereSQL = " OR Class IS NULL "
'        End If

'        ' set prev CustomerItemNumber, SeasonCode, CategoryCode, ProgramNumber, SubProgram, Grade
'        ' values for ProgramYear   =  values of latest Rev for ProgramYear,
'        ' OR = values of lower rev for program year
'        ' depending on where this function was called from
'        '2012/05/02 - added sCustomerCartonUPCNumber
'        SQL = "UPDATE Item " & _
'                "SET CustomerItemNumber = '" & sCustItemNum & "'" & _
'                ", CustomerCartonUPCNumber = '" & sCustomerCartonUPCNumber & "'" & _
'                ", SeasonCode = '" & sSeason & "'" & _
'                ", CategoryCode = '" & sCategory & "'" & _
'                ", ProgramNumber = '" & sProgram & "'" & _
'                sSubProgramSQL & _
'                sGradeSQL & vbCrLf & _
'                sClassSQL & vbCrLf & _
'                " WHERE ProposalNumber = " & lProposal & _
'                " AND Rev <> " & lRev & _
'                " AND ProgramYear = " & sProgYear & _
'                " AND (CustomerItemNumber <> '" & sCustItemNum & "' OR SeasonCode <> '" & sSeason & "'" & _
'                " OR CategoryCode <> '" & sCategory & "'" & _
'                " OR ProgramNumber <> " & sProgram & _
'                " OR SubProgram <> '" & sSubProgram & "'" & _
'                " OR Grade <> '" & sGrade & "'" & _
'                " OR Class <> '" & sClass & "'" & _
'                SClassWhereSQL & ")"                '11/24/2008 - hn added last 2 lines

'        SSDataConn.Execute(SQL, lRecordsAffected)     '11/21/2008 - hn
'        If lRecordsAffected > 0 Then
'            '11/24/2008 - hn - added Class, removed TempItemNumber
'            sMsg = sMsg & "For ProgramYear=" & sProgYear & "; updated Season/Cat/Program/SubProgram/Grade/Class/CustomerItemNumber for Revs <> " & lRev
'        End If

'        If bPROPOSALFormIndicator = True Then
'            'also Update local Access table: TempItem (so that changed values show when scrolling thru revs)
'            Dim RSTemp As ADODB.Recordset : RSTemp = New ADODB.Recordset
'            '        SQL = "UPDATE TempItem " & _
'            '                "SET CustomerItemNumber = '" & sCustItemNum & _
'            '                "', SeasonCode = '" & sSeason & _
'            '                "', CategoryCode = '" & sCategory & _
'            '                "', ProgramNumber = '" & sProgram & _
'            '                sSubProgramSQL & _
'            '                sGradeSQL & _
'            '                "' WHERE ProposalNumber = " & lProposal & _
'            '                " AND Rev <> " & lRev & _
'            '                " AND ProgramYear = " & sProgYear & _
'            '                " AND (SeasonCode <> '" & sSeason & _
'            '                "' OR CategoryCode <> ' " & sCategory & _
'            '                "' OR ProgramNumber <> '" & sProgram & _
'            '                "' OR SubProgram <> '" & sSubProgram & _
'            '                "' OR  Grade <> '" & sGrade & "')"

'            '11/21/2008 - hn - ensure that TempItem SubProgram and Grade Fields are defined as Text(4) to allow Null values
'            SQL = "UPDATE TempItem " & _
'                "SET CustomerItemNumber = '" & sCustItemNum & "'" & _
'                ", SeasonCode = '" & sSeason & "'" & _
'                ", CategoryCode = '" & sCategory & "'" & _
'                ", ProgramNumber = '" & sProgram & "'" & _
'                sSubProgramSQL & _
'                sGradeSQL & vbCrLf & _
'                sClassSQL & vbCrLf & _
'                " WHERE ProposalNumber = " & lProposal & _
'                " AND Rev <> " & lRev & _
'                " AND ProgramYear = " & sProgYear & _
'                " AND (CustomerItemNumber <> '" & sCustItemNum & "' OR SeasonCode <> '" & sSeason & "'" & _
'                " OR CategoryCode <> '" & sCategory & "'" & _
'                " OR ProgramNumber <> " & sProgram & _
'                " OR SubProgram <> '" & sSubProgram & "'" & _
'                " OR Grade <> '" & sGrade & "'" & _
'                " OR Class <> '" & sClass & "'" & _
'                SClassWhereSQL & ")"

'Dim         RSTemp.Open SQL As Object 
'Dim  CurrentProject.Connection As Object 
'            Dim adOpenStatic As Object
'            Dim adLockOptimistic As Object

'            Application.DoEvents()
'            Form_frmProposal.Requery()
'            Application.DoEvents()
'        End If

'        bUpdateRevValues = True

'ExitRoutine:
'        Exit Function
'ErrorHandler:
'        '    If bPROPOSALFormIndicator = False Then
'        '        If Err.Number = -2147217871 Then  '  2014/01/20 RAS Adding this to have it try the query again.
'        '               If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'        '                smessage = "In bUpdateRevValues, Retrying update, Err Number " & Err.Number & "Error Description: " & Err.Description
'        '                If bWritePrintToLogFile(False, objExcel.name & smessage, "ErrorMessageLog") = False Then
'        '                End If
'        '            End If
'        '            bErrorCount = bErrorCount + 1
'        '            If bErrorCount < 6 Then
'        ''                GoTo StartAllOver    'TOP of the select again.'TODO - GoTo Statements are redundant in .NET
'        '            End If
'        '        End If
'        '    End If
'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bUpdateRevValues")
'        ' Regenerate original error.
'        Dim intErrNum As Long
'        intErrNum = Err()
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bUpdateRevValues, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        ' Resume Next '' this is for testing
'        '2014/01/17 RAS Raising the error so the calling subroutine can handle it.

'        Err.Clear()
'        Err.Raise(intErrNum)
'        Resume ExitRoutine

'    Private Function bUpdateItemStatus(ByVal lProposal As Long, ByVal lRev As Long, _
'                ByVal sItemStatus As String, ByVal sProgYear As String, ByVal sFactory As String, ByRef sMsg As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'    End Function

'    Private Function bUpdateTempItem(ByVal lProposal As Long, ByVal lRev As Long, ByVal sProgYear As String, ByVal sTempItemNum As String, ByRef sMsg As String) As Boolean
'        '01/31/2008 TW decided against this for now- will allow more than 1 TempItemNumber per Proposal
'        ''On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        'Dim SQL As String 
'        'Dim rsTemp As ADODB.Recordset: Set rsTemp = New ADODB.Recordset
'        '    SQL = "Update Item " & _
'        '            " SET TempItemNumber = '" & UCase(sTempItemNum) & _
'        '            "' WHERE ProposalNumber = " & lProposal & _
'        '            " AND Rev <> " & lRev & _
'        '            " AND ((ProgramYear = " & sProgYear & ") " & _
'        '            " OR ( ProgramYear <> " & sProgYear & " AND TempItemNumber <> ''))"
'        '        SSDataConn.Execute SQL
'        '        Application.DoEvents
'        '        sMsg = sMsg & " Updated TempItemNumber for all Revs where TempItemNumber is not blank"
'        '    If bPROPOSALFormIndicator = True Then
'        '        SQL = Replace(SQL, "Update Item", "Update TempItem")
'Dim '        rsTemp.Open SQL As Object 
'Dim  CurrentProject.Connection As Object 
'        Dim adOpenStatic As Object
'        Dim adLockOptimistic As Object

'        '        Application.DoEvents
'        '        Form_frmProposal.Requery
'        '    End If
'        'ExitRoutine:
'        '    Exit Function
'        'ErrorHandler:
'        '    MsgBox Err.Description, vbExclamation, "modSpreadSheet-bUpdateTempItem"
'        '    Resume ExitRoutine
'    End Function

'    Public Function ExcelColumnName(ByVal CellValue As Object) As Object '02/27/2008 gl
'        ' Returns the column name based on the value in an excel cell.
'        ' This normally is the same as the cell value,
'        '  except that if the value contains <> characters, then it is the alias value between the < and > characters.

'        Dim ValueBetweenSpecialChars As Object
'        ValueBetweenSpecialChars = ParseToken(CellValue, "<", ">")
'        ExcelColumnName = Nz(ValueBetweenSpecialChars, CellValue)

'    End Function

'    Private Function bValidateProductBatteriesIncluded(ByVal sTechnology As Object, ByVal sLighted As Object, ByVal sProductBatteriesIncluded As Object) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        bValidateProductBatteriesIncluded = False
'        If InStr(1, sTechnology, "112") Or sLighted = "Battery Operated" Then
'            '        If IsBlank(sProductBatteriesIncluded) Then GoTo ExitRoutine             '03/19/2008'TODO - GoTo Statements are redundant in .NET
'        End If

'        bValidateProductBatteriesIncluded = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateProductBatteriesIncluded")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidateProductBatteriesIncluded, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'    Private Function bBuildErrorMsg(ByVal bCreateNEWComponentItems As Boolean, ByVal lRow As Long, ByRef lNbrErrors As Long, ByRef sRowErrorMsg As String, ByVal sErrorFound As String) As Boolean
'        'On Error GoTo ErrorHandler:'TODO - On Error must be replaced with Try, Catch, Finally
'        lNbrErrors = lNbrErrors + 1
'        If bCreateNEWComponentItems = True Or bPROPOSALFormIndicator = True Then
'            sRowErrorMsg = sRowErrorMsg & " - " & sErrorFound & vbCrLf
'        Else
'            sRowErrorMsg = sRowErrorMsg & "Row: " & lRow & " - " & sErrorFound & vbCrLf  'this is for spreadsheets
'        End If
'        Application.DoEvents()
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bBuildErrorMsg")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bBuildErrorMsg, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine
'    End Function

'    Friend Overridable Function bCheckCommentsLength(ByVal objExcel As Excel.Application, ByVal lMaxRows As Long, ByVal sCommentsType As String) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim sCellValue As String
'        Dim lColCounter As Long
'        Dim lRowCounter As Long

'        Dim lProposalCOLPos As Long
'        Dim lProposalNumber As Long

'        Dim lRevCOLPos As Long
'        Dim lRev As Long

'        Dim lCommentsColumn As Long
'        Dim sComments As String

'        Dim ws As Excel.Worksheet ' Set a reference to this excel worksheet.
'        Dim i As Integer
'        Dim SQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        bCheckCommentsLength = False
'        ws = objExcel.Application.Workbooks(1).Worksheets(1)
'        'first determine the column positions needed
'        For lColCounter = 1 To glMAX_Cols
'            '        If gbCancelExport = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            sCellValue = sGetCellValue(ws.Cells(1, lColCounter))
'            If sCellValue = "ProposalNumber" Then lProposalCOLPos = lColCounter
'            If sCellValue = "Rev" Then lRevCOLPos = lColCounter
'            If sCellValue = sCommentsType Then
'                lCommentsColumn = lColCounter
'                If lProposalCOLPos > 0 And lRevCOLPos > 0 Then
'                    Exit For
'                End If
'            End If
'        Next lColCounter

'        ' now breakdown comments to strings of 256 - excel cannot import/export more
'        For lRowCounter = glDATA_START_ROW To lMaxRows
'            If IsBlank(sGetCellValue(ws.Cells(lRowCounter, lProposalCOLPos))) = False Then    '2013/05/01 -HN
'                lProposalNumber = sGetCellValue(ws.Cells(lRowCounter, lProposalCOLPos))
'                lRev = sGetCellValue(ws.Cells(lRowCounter, lRevCOLPos))

'                SQL = "SELECT  " & sCommentsType & " AS Comments " & _
'                      "FROM Item " & _
'                      "WHERE ProposalNumber = " & lProposalNumber & " AND Rev = " & lRev
'Dim        rs.Open SQL As Object 
'                Dim SSDataConn As Object
'                Dim adOpenStatic As Object
'                Dim adLockReadOnly As Object


'                If Not rs.EOF Then
'                    If Not IsBlank(rs!Comments) Then
'                        sComments = rs!Comments
'                        If Len(sComments) > 255 Then
'                            With ws.Cells(lRowCounter, lCommentsColumn).Select
'                                ws.Cells(lRowCounter, lCommentsColumn) = ""
'                                For i = 0 To Int(Len(sComments) / 255)
'                                    ws.Cells(lRowCounter, lCommentsColumn) = ws.Cells(lRowCounter, lCommentsColumn) & Mid(sComments, (i * 255) + 1, 255)
'                                Next
'                            End With
'                        End If
'                    End If

'                End If
'                rs.Close()
'            End If                                                                             '2013/05/01 -HN
'        Next lRowCounter

'        bCheckCommentsLength = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bCheckCommentsLength")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCheckCommentsLength, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume Next

'    End Function

'    Private Function bItemInOrder(ByRef sItemErrMsg As String) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        '2013/05/07 -HN- added ET's code below
'        ' ET 2013-03-15 - Separated this method into the one following - bValidItemStatusChange
'        ' to separate the two functions of checking for an item in an order or checking for an
'        ' invalid status change on an item.
'        ' Called from bDeleteData to check if item is in an order.  bDeleteData is called from
'        ' frmExcelDelete.cmdDeleteFromFile (on Maintenance menu) when deleting proposals and
'        ' revs using a list in a file and from frmProposalList.cmdDelete when deleting a single
'        ' proposal and rev.

'        Dim SQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        Dim RS1 As ADODB.Recordset : RS1 = New ADODB.Recordset

'        bItemInOrder = False
'        sItemErrMsg = ""
'        '2010/02/09 - took out compare to Rev; because the Order might exist for an earlier Rev,
'        'and the status is changed for ALL Revs for a ProgramYear if there are no errors
'        '2010/10/26 - for Deletes from ProposalList still need to delete by each rev
'        SQL = "SELECT * FROM vItemInOrder WHERE ProposalNumber = " & sProposalNumber & " AND Rev = " & sRev '2013/05/07 -HN- check only ProposalNumber and Rev
'        '              " AND ProgramYear = " & lProgramYear

'Dim     rs.Open SQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object


'        Do Until rs.EOF
'            If rs!ItemInOrder = "TRUE" Then

'                SQL = "SELECT OrderNumber, PONumber, ItemNumber" & vbCrLf & _
'                      "FROM vDB4Order" & vbCrLf & _
'                      "WHERE ProposalNumber = " & rs!ProposalNumber & " AND Rev = " & rs!Rev & _
'                                " AND ProgramYear = " & lProgramYear & vbCrLf & _
'                      " ORDER BY OrderNumber, PONumber"

'Dim             RS1.Open SQL As Object 
'                Dim SSDataConn As Object
'                Dim adOpenStatic As Object
'                Dim adLockReadOnly As Object


'                sItemErrMsg = "Cannot Delete a Proposal/Rev if Customer Orders/PONumbers exist:"

'                Do Until RS1.EOF
'                    sItemErrMsg = sItemErrMsg & vbCrLf & "OrderNumber: " & RS1!OrderNumber & ", PONumber: " & _
'                                RS1!PONumber & " ItemNumber: " & RS1!ItemNumber & vbCrLf
'                    RS1.MoveNext()
'                Loop
'                bItemInOrder = True
'                RS1.Close()
'            End If

'            rs.MoveNext()
'        Loop
'ExitRoutine:
'        If rs.State <> 0 Then rs.Close()
'        If RS1.State <> 0 Then RS1.Close()
'        rs = Nothing
'        RS1 = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bItemInOrder")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bItemInOrder, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume Next
'    End Function
'    '2013/05/07 -HN- copied ET's code below from SSApp_TEST.accdb
'Private Function bValidItemStatusChange(ByVal sProposalNumber As String, _
'        End Function

'        Private Function bValidItemStatusChange(ByVal sProposalNumber As String, ByVal ByVal lProgramYear As Long, ByVal ByVal sItemNumber As String, ByVal ByRef sItemErrMsg As String) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally

'        ' Called from bValidateCoreFields to check for an invalid status change for an item.

'        Dim SQL As String
'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset

'        bValidItemStatusChange = True
'        sItemErrMsg = ""

'        '2010/02/09 - took out compare to Rev; because the Order might exist for an earlier Rev,
'        ' and the status is changed for ALL Revs for a ProgramYear if there are no errors

'        SQL = "SELECT OrderNumber, PONumber FROM vDB4OrderNOCancel " & _
'               "WHERE ProposalNumber = " & sProposalNumber & _
'               " AND ProgramYear = " & lProgramYear & _
'               " AND ItemNumber = '" & sItemNumber & "'"
'Dim     rs.Open SQL As Object 
'        Dim SSDataConn As Object
'        Dim adOpenStatic As Object
'        Dim adLockReadOnly As Object


'        If rs.Recordcount = 0 Then
'            bValidItemStatusChange = True       ' if no records returned, it's OK to change the status of the item
'        Else

'            ' ET 2013-03-15 - per Theresa should only get an error message here when
'            ' changing the status of an item when it exists on another orderdetail where the order status
'            ' is not "cancelled" (OrderDetail.CancelCode is NULL or empty string). That's why we are
'            ' using vDB4OrderNOCancel above. It does not return any orders where the item is cancelled.

'            sItemErrMsg = sItemErrMsg & "For ProgramYear: " & lProgramYear & _
'                ", cannot change ItemStatus FROM 'ORD', Item exists in " & _
'                "OrderDetail that is not cancelled(" & _
'                sProposalNumber & "/" & sItemNumber & ")" & vbCrLf

'            bValidItemStatusChange = False

'            ' put a list of the other orders where this item appears in the error log (C:\DB4)
'            Do Until rs.EOF
'                sItemErrMsg = sItemErrMsg & vbCrLf & "OrderNumber: " & rs!OrderNumber & ", PONumber: " & _
'                            rs!PONumber & " ItemNumber: " & sItemNumber & vbCrLf
'                rs.MoveNext()
'            Loop

'            rs.Close()
'        End If

'ExitRoutine:
'        If rs.State <> 0 Then rs.Close()
'        rs = Nothing
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bValidItemStatusChange")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidItemStatusChange, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume Next

'        '2009/11/19 - hn - new - Bag_SpecialEffects now has concatenated values as from the Bag_SpecialEffects table
'Dim Public Function bValidBag_SpecialEffects(ByVal lRow As Long
'Dim  ByRef sBag_SpecialEffects As String 
'Dim  ByRef sInvalidBag_SpecialEffects As String ) As Boolean

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'    End Function

'    Friend Overridable Function bValidBag_SpecialEffects(ByVal lRow As Long, ByRef sBag_SpecialEffects As String, ByRef sInvalidBag_SpecialEffects As String) As Boolean
'        Dim i As Long
'        Dim sSpecialEffect As String
'        Dim sSpecialEffectsDesVbCr As String        'up to 12 Bag_SpecialEffects

'        Dim SQL As String
'        Dim lDupCounter As Long
'        Dim sDuplicates As String

'        bValidBag_SpecialEffects = False
'        sInvalidBag_SpecialEffects = ""
'        sSpecialEffectsDesVbCr = Split(sBag_SpecialEffects, ",")
'        For i = LBound(sSpecialEffectsDescr) To UBound(sSpecialEffectsDescr)
'            Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'            SQL = "SELECT Bag_SpecialEffects FROM Bag_SpecialEffects WHERE Bag_SpecialEffects = " & sAddQuotes(Trim(sSpecialEffectsDescr(i)))
'Dim         rs.Open SQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockOptimistic As Object

'            If rs.EOF Then
'                sInvalidBag_SpecialEffects = sInvalidBag_SpecialEffects & Trim(sSpecialEffectsDescr(i)) & " "
'            End If
'            rs.Close()
'            rs = Nothing

'            'this could happen on importing a sheet:
'            For lDupCounter = 0 To i
'                If Trim(sSpecialEffectsDescr(i)) = Trim(sSpecialEffectsDescr(lDupCounter)) And i <> lDupCounter Then
'                    sDuplicates = sDuplicates & " " & Trim(sSpecialEffectsDescr(i))
'                End If
'            Next lDupCounter

'        Next i

'        If sDuplicates <> "" Then
'            sInvalidBag_SpecialEffects = sInvalidBag_SpecialEffects & vbCrLf & " Duplicate Bag_SpecialEffects Found: " & sDuplicates
'        End If
'        '    If sInvalidBag_SpecialEffects <> "" Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        bValidBag_SpecialEffects = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidBag_SpecialEffects")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidBag_SpecialEffects, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'        '2010/11/17 - new
'    End Function

'    Public Function bFindLastColumnAndRow(ByVal objExcel As Excel.Application, ByRef lLastColumn As Long, ByRef llastrow As Long, ByRef sLastColCellHeading As Object) As Boolean
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lResponse As Long
'        Dim lColCounter As Long
'        Dim lRowCounter As Long

'        Dim ws As Excel.Worksheet ' Set a reference to this excel worksheet.
'        Dim sColCellValue As String
'        Dim lColColorIndex As Long
'        Dim sRowCol1Cell As String       'if the first 3 cells on a row are blank
'Dim  no color As Object 
'Dim  then that indicates last row As Object 

'        Dim sRowCol2Cell As String
'        Dim sRowCol3Cell As String
'        Dim lRowCol1ColorIndex As Long
'        Dim lRowCol2ColorIndex As Long
'        Dim lRowCol3ColorIndex As Long

'        bFindLastColumnAndRow = False
'        bMaterialFound = False : bBreakdownFound = False : bCostFound = False
'        ws = objExcel.Application.Workbooks(1).Worksheets(1)

'        lLastColumn = glMAX_Cols ' in case blank column is not found
'        sLastColCellHeading = ""
'        'determine last column if < 301
'        For lColCounter = 1 To glMAX_Cols
'            '        If gbCancelRefresh = True Or gbValidationCancelled = True Or gbCancelExport = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET

'            sColCellValue = sGetCellValue(ws.Cells(1, lColCounter))
'            'first blank/no color column indicates last column on spreadsheet ...
'            lColColorIndex = ws.Cells(1, lColCounter).Interior.ColorIndex
'            If Len(sColCellValue) = 0 And lColColorIndex = xlNone Then
'                sLastColCellHeading = sGetCellValue(ws.Cells(1, lColCounter - 1))
'                lLastColumn = lColCounter - 1
'                Exit For
'            End If
'        Next lColCounter

'        If sLastColCellHeading = "" Then
'            sLastColCellHeading = sGetCellValue(ws.Cells(1, lColCounter - 1))
'        End If '2011/12/20

'        'Determine last row - first 3 columns blank, no color is end of spreadsheet
'        llastrow = glMAX_Rows  'if no blank row found

'        For lRowCounter = glDATA_START_ROW To glMAX_Rows
'            sRowCol1Cell = sGetCellValue(ws.Cells(lRowCounter, 1))
'            sRowCol2Cell = sGetCellValue(ws.Cells(lRowCounter, 2))
'            sRowCol3Cell = sGetCellValue(ws.Cells(lRowCounter, 3))
'            lRowCol1ColorIndex = ws.Cells(lRowCounter, 1).Interior.ColorIndex
'            lRowCol2ColorIndex = ws.Cells(lRowCounter, 2).Interior.ColorIndex
'            lRowCol3ColorIndex = ws.Cells(lRowCounter, 3).Interior.ColorIndex

'            If Len(sRowCol1Cell) = 0 And lRowCol1ColorIndex = xlNone And _
'                Len(sRowCol2Cell) = 0 And lRowCol2ColorIndex = xlNone And _
'                Len(sRowCol3Cell) = 0 And lRowCol3ColorIndex = xlNone Then
'                llastrow = lRowCounter - glDATA_START_ROW + 1                   '2010/12/07
'                Exit For
'            End If
'        Next lRowCounter

'        bFindLastColumnAndRow = True

'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        If Err.Number = 1004 Then                   '2010/11/16
'            MsgBox("There is an error with the column: '" & lColCounter & "  heading' on the Spreadsheet;" & vbCrLf & vbCrLf & _
'            "   please check that your LAST COLUMN heading(to indicate end of spreadsheet columns)," & vbCrLf & vbCrLf & _
'            "   has NO formatting, color and is blank, then retry!", vbOKOnly + vbMsgBoxSetForeground, "Error on spreadsheet column on heading row!")
'            ws.Columns(lColCounter - 1).Select()
'            '        ws.Cells(1, lCOLCounter - 1).Select
'        Else
'            MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-ConcatenateOldMaterialColumns")  '2010/11/15
'            '    Resume Next
'        End If
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In ConcatenateOldMaterialColumns, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'    End Function

'    Friend Overridable Function bValidCertifiedPrinterID(ByVal lRow As Long, ByRef sCertifiedPrinterID As String, ByRef sXCertifiedPrinterName As String, ByRef sInvalidCertifiedPrinterIDs As String) As Object
'        Dim i As Long
'        Dim sCertPrintID As String
'        Dim sCertPrintIDs() As String    ' for now up to 6 Certified Printer IDs

'        Dim SQL As String
'        Dim lDupCounter As Long
'        Dim sDuplicates As String

'        bValidCertifiedPrinterID = False
'        sInvalidCertifiedPrinterIDs = ""
'        sCertPrintID = Replace(sCertifiedPrinterID, ",", " ")        '2011/10/27
'        sCertPrintID = Replace(sCertPrintID, "  ", " ")
'        sCertPrintID = Replace(sCertPrintID, "  ", " ")       '2011/10/27
'        sCertPrintIDs() = Split(sCertPrintID, " ")

'        For i = LBound(sCertPrintIDs) To UBound(sCertPrintIDs)
'            Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'            SQL = "SELECT CertifiedPrinterID, CertifiedPrintername FROM TargetCertifiedPrinters WHERE convert(varchar(4),CertifiedPrinterID) = " & sAddQuotes(Trim(sCertPrintIDs(i)))
'Dim         rs.Open SQL As Object 
'            Dim SSDataConn As Object
'            Dim adOpenStatic As Object
'            Dim adLockOptimistic As Object

'            If rs.EOF Then
'                sInvalidCertifiedPrinterIDs = sInvalidCertifiedPrinterIDs & Trim(sCertPrintIDs(i)) & " "
'            Else
'                sX_CertifiedPrinterName = sX_CertifiedPrinterName & rs!CertifiedPrinterName & "," & vbCrLf
'            End If

'            rs.Close()
'            rs = Nothing

'            'this could happen on importing a sheet:
'            For lDupCounter = 0 To i
'                If Trim(sCertPrintIDs(i)) = Trim(sCertPrintIDs(lDupCounter)) And i <> lDupCounter Then
'                    sDuplicates = sDuplicates & " " & Trim(sCertPrintIDs(i))
'                End If
'            Next lDupCounter
'        Next i

'        If Len(sX_CertifiedPrinterName) > 2 Then
'            sX_CertifiedPrinterName = Microsoft.VisualBasic.Left(sX_CertifiedPrinterName, Len(sX_CertifiedPrinterName) - 2)
'        End If
'        If sDuplicates <> "" Then
'            sInvalidCertifiedPrinterIDs = sInvalidCertifiedPrinterIDs & " Duplicate Certified Printer ID's Found: " & sDuplicates
'        End If
'        '    If sInvalidCertifiedPrinterIDs <> "" Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'        bValidCertifiedPrinterID = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidCertifiedPrinterID")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bValidCertifiedPrinterID, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'        '2011/10/26 - new
'Dim Public Function bInsertX_CertifiedPrinterName(objExcel As Excel.Application
'Dim  ByVal lMaxRows As Long) As Boolean

'    End Function

'    Friend Overridable Function bInsertX_CertifiedPrinterName(ByVal objExcel As Excel.Application, ByVal lMaxRows As Long) As Boolean
'        Dim sCellValue As String
'        Dim lColCounter As Long
'        Dim lRowCounter As Long

'        Dim lCertifiedPrinterIDCOLPos As Long
'        Dim lX_CertifiedPrinterNameCOLPos As Long

'        Dim sCertifiedPrinterIDs As String
'        Dim sX_CertifiedPrinterNames As String
'        Dim sInvalidCertifiedPrinterIDs As String

'        Dim ws As Excel.Worksheet ' Set a reference to this excel worksheet.

'        bInsertX_CertifiedPrinterName = False
'        ws = objExcel.Application.Workbooks(1).Worksheets(1)
'        'first determine the column positions needed
'        For lColCounter = 1 To glMAX_Cols
'            '        If gbCancelExport = True Then GoTo ExitRoutine'TODO - GoTo Statements are redundant in .NET
'            sCellValue = sGetCellValue(ws.Cells(1, lColCounter))
'            If sCellValue = gsCOL_CertifiedPrinterID Then lCertifiedPrinterIDCOLPos = lColCounter
'            If sCellValue = gsCOL_X_CertifiedPrinterName Then lX_CertifiedPrinterNameCOLPos = lColCounter

'            If lCertifiedPrinterIDCOLPos > 0 And lX_CertifiedPrinterNameCOLPos > 0 Then
'                Exit For
'            End If
'        Next lColCounter

'        If lCertifiedPrinterIDCOLPos = 0 Or lX_CertifiedPrinterNameCOLPos = 0 Then
'        Else
'            ' now insert X_CertifiedPrinterName for each row
'            For lRowCounter = glDATA_START_ROW To lMaxRows
'                sCertifiedPrinterIDs = sGetCellValue(ws.Cells(lRowCounter, lCertifiedPrinterIDCOLPos))

'                sX_CertifiedPrinterNames = ""
'                If sCertifiedPrinterIDs <> "" Then
'                    If bValidCertifiedPrinterID(1, sCertifiedPrinterIDs, sX_CertifiedPrinterNames, sInvalidCertifiedPrinterIDs) = True Then
'                        ws.Cells(lRowCounter, lX_CertifiedPrinterNameCOLPos) = sX_CertifiedPrinterNames
'                    Else
'                        ws.Cells(lRowCounter, lX_CertifiedPrinterNameCOLPos) = sX_CertifiedPrinterNames & "Invalid Certified Printer ID(s): " & sInvalidCertifiedPrinterIDs
'                    End If
'                Else
'                    ws.Cells(lRowCounter, lX_CertifiedPrinterNameCOLPos) = ""
'                End If

'            Next lRowCounter
'            bInsertX_CertifiedPrinterName = True
'        End If

'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description & vbExclamation + vbMsgBoxSetForeground, "modSpreadsheet: bInsertX_CertifiedPrinterName")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bInsertX_CertifiedPrinterName, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume Next

'        '2012/01/04 - new
'    End Function

'        Friend Overridable Function bWriteToExcelLogFileIfError(ByVal objExcel As Excel.Application, ByVal ByVal slogfilename As String, ByVal lErrorRow As Long, ByVal ByVal lImportChangedProposals As Long, ByVal lImportRevisedProposals As Long, ByVal lImportProposalsAdded As Long) As Object
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        bWriteToExcelLogFileIfError = False
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 19).Value = "Import ended: " & dtImportEnd        '2012/01/12
'        If lErrorRow > 0 Then
'            objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 31).Font.ColorIndex = 3 'red                  '2012/01/17
'            objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 31).Value = "Import Row Errored: " & lErrorRow - 1   '2012/01/17
'        Else
'            objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 31).Value = "No ERRORS! "                     '2012/01/17
'        End If
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 32).Value = "# Proposals changed: " & lImportChangedProposals
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 33).Value = "# Proposals revised: " & lImportRevisedProposals
'        objExcel.ActiveWorkbook.Worksheets(1).Cells(1, 34).Value = "# Proposals added: " & lImportProposalsAdded


'        Application.DoEvents()
'        If Version >= 12.0# Then
'            objExcel.ActiveWorkbook.Worksheets(1).SaveAs(slogfilename, , , , , 0)    '2010/05/11 to prevent .xlk backup file being created
'            Application.DoEvents()
'        Else
'            objExcel.Application.Workbooks(1).SaveAs(slogfilename, xlNormal)
'        End If
'        Application.DoEvents()
'        bWriteToExcelLogFileIfError = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bWriteToExcelLogFileIfError")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bWriteToExcelLogFileIfError, Err Number " & Err.Number & " ,Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        Resume ExitRoutine

'        '2012/01/04 - added lRowCounter
'    Private Function bCheckItemFieldsChangedPerRow(lRowCounter As Long, RowChangesARRAY() As String, sFromProcess As String, sItemSpecsArray_ORIG() As String, sItemSPECSArray() As String, lItemSpecsFields As Long, _
'                        sItemArray_Orig() As String, sItemArray() As String, lItemFields As Long, _
'                        sAssortmentArray_ORIG() As String, sAssortmentArray() As String, lAssortmentFields As Long, _
'                        lMaxRows As Long, dtSaveArrayCOLPos As typSpecialCOLPos) As Boolean

'        'Private Function bCheckItemFieldsChangedPerRow(ByVal lRowCounter As Object, ByVal  RowChangesARRAY( As Object) As String, sFromProcess As String
'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lColCounter As Long '
'        Dim lRowCounter As Long

'        Dim sSQL As String
'        Dim sProposalNumber As String
'        Dim sRev As String
'        Dim sPrevRev As String

'        Dim sItemNumber As String
'        Dim sFunctioncode As String
'        Dim dRevisionDate As Date
'        Dim sArray_ORIG() As String
'        Dim sArray() As String

'        Dim rs As ADODB.Recordset : rs = New ADODB.Recordset
'        Dim sCounter As String  '* 2
'        Dim lRowChangeFound As Long             '2010/02/11

'        bCheckItemFieldsChangedPerRow = False
'        dRevisionDate = Now.ToShortDateString() 'not used, because Access does not give us the time in milliseconds
'        ' define default value on ItemFieldHistory table as getdate() which gives us milliseconds

'        sSQL = "SELECT * FROM ItemFieldHistory WHERE RevisedDate > '" & dRevisionDate & "'"
'        rs.Open(sSQL, SSDataConn, adOpenDynamic, adLockPessimistic)

'        If bPROPOSALFormIndicator = False Then
'            sArray_ORIG() = sItemSpecsArray_ORIG()          '2010/02/12 - first check ItemSpecs fields
'            sArray() = sItemSPECSArray()

'            '        For lRowCounter = 2 To lMaxRows
'            sFunctioncode = sArray(lRowCounter, 1)
'            sProposalNumber = sArray(lRowCounter, 2)
'            sRev = sArray(lRowCounter, 3)
'            sPrevRev = sArray_ORIG(lRowCounter, 3)
'            sItemNumber = sItemArray_Orig(lRowCounter, dtSaveArrayCOLPos.lItemNumber)
'            If sFunctioncode = "" Or sFunctioncode = gsNEW_ITEM_NBR Then
'            Else
'                For lColCounter = 4 To lItemSpecsFields + glFIXED_COLS  '2011/02/12 was using lItemFields
'                    If Not sArray_ORIG(lRowCounter, lColCounter) = sArray(lRowCounter, lColCounter) Then
'                        lRowChangeFound = RowChangesARRAY(lRowCounter)  '2010/02/11
'                        'need ItemNumber for the ItemFieldHistory table
'                        sItemNumber = sItemArray_Orig(lRowCounter, dtSaveArrayCOLPos.lItemNumber)
'                        If bSAVEItemFieldsChanged(lRowChangeFound, rs, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctioncode, sArray_ORIG(1, lColCounter), _
'                                 sArray_ORIG(lRowCounter, lColCounter), sArray(lRowCounter, lColCounter), _
'                                dRevisionDate, gsUserID) = False Then
'                        End If
'                    End If
'                Next lColCounter
'            End If

'            '        Next lRowCounter
'        End If

'        sArray_ORIG() = sItemArray_Orig()           '2010/02/12 -then check Item fields
'        '    sArray() = sItemSPECSArray()               '2010/02/12 ? why here?
'        sArray() = sItemArray()

'        '    For lRowCounter = 2 To lMaxRows
'        sFunctioncode = sArray(lRowCounter, 1)
'        sProposalNumber = sArray(lRowCounter, 2)
'        sRev = sArray(lRowCounter, 3)
'        sPrevRev = sArray_ORIG(lRowCounter, 3)
'        sItemNumber = sItemArray_Orig(lRowCounter, dtSaveArrayCOLPos.lItemNumber)
'        If sFunctioncode = "" Or sFunctioncode = gsNEW_ITEM_NBR Then
'        Else
'            For lColCounter = 3 To lItemFields + glFIXED_COLS           '2010/02/11 was 4 trying to show rev change, if a new rev without any other changes
'                If sArray_ORIG(1, lColCounter) = "RevisedUserID" Or _
'                    sArray_ORIG(1, lColCounter) = "RevisedDate" Then
'                    Application.DoEvents()
'                Else
'                    If Not sArray_ORIG(lRowCounter, lColCounter) = sArray(lRowCounter, lColCounter) Then
'                        lRowChangeFound = RowChangesARRAY(lRowCounter)  '2010/02/11
'                        If bSAVEItemFieldsChanged(lRowChangeFound, rs, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctioncode, sArray_ORIG(1, lColCounter), _
'                                 sArray_ORIG(lRowCounter, lColCounter), sArray(lRowCounter, lColCounter), _
'                                dRevisionDate, gsUserID) = False Then
'                            '                            GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'                        End If
'                    End If
'                End If
'            Next lColCounter
'        End If
'        '    Next lRowCounter

'        If bPROPOSALFormIndicator = False Then 'do the following for Import Spreadsheet
'            sArray_ORIG() = sAssortmentArray_ORIG()     '2010/02/12 Item_Assortment fields
'            sArray() = sAssortmentArray()

'            '        For lRowCounter = 2 To lMaxRows
'            sFunctioncode = sArray(lRowCounter, 1)
'            sProposalNumber = sArray(lRowCounter, 2)
'            sRev = sArray(lRowCounter, 3)
'            sPrevRev = sArray_ORIG(lRowCounter, 3)
'            sItemNumber = sItemArray_Orig(lRowCounter, dtSaveArrayCOLPos.lItemNumber)

'            If sFunctioncode = "" Or sFunctioncode = gsNEW_ITEM_NBR Then
'            Else
'                For lColCounter = 4 To lAssortmentFields + glFIXED_COLS           '2010/02/15 - dont start at 3 otherwise it writes a rev 2X- was going against lItemFields!'2010/02/11 was 4 trying to show rev change, if a new rev without any other changes
'                    If sArray_ORIG(1, lColCounter) = "RevisedUserID" Or _
'                        sArray_ORIG(1, lColCounter) = "RevisedDate" Then
'                    Else
'                        If sArray_ORIG(1, lColCounter) = "ASSORTMENTS" Then
'                            Application.DoEvents()
'                        End If
'                    End If
'                    If Not sArray_ORIG(lRowCounter, lColCounter) = sArray(lRowCounter, lColCounter) Then
'                        lRowChangeFound = RowChangesARRAY(lRowCounter)  '2010/02/11
'                        If bSAVEItemFieldsChanged(lRowChangeFound, rs, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctioncode, sArray_ORIG(1, lColCounter), _
'                                 sArray_ORIG(lRowCounter, lColCounter), sArray(lRowCounter, lColCounter), _
'                                dRevisionDate, gsUserID) = False Then
'                            '                                GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'                        End If
'                    End If
'                Next lColCounter
'            End If
'            '        Next lRowCounter
'        Else
'            'do the following for Proposal Form Assortment changes
'            If gbProposalAssortmentsChanged = True Then
'                If sFunctioncode = "" Or sFunctioncode = gsNEW_ITEM_NBR Then
'                Else
'                    For lColCounter = 0 To 40
'                        If lColCounter + 1 < 10 Then
'                            sCounter = "0" & lColCounter + 1
'                        Else
'                            sCounter = lColCounter + 1
'                        End If

'                        '2010/02/12 - for ProposalForm lRowCounter ALWAYS = 2; like it's a spreadsheet with 1 row only
'                        lRowCounter = 2
'                        'Assortment ItemNumbers
'                        If Not sORIGAssortmentArray(lColCounter, 1) = sNEWAssortmentArray(lColCounter, 1) Then
'                            RowChangesARRAY(lRowCounter) = RowChangesARRAY(lRowCounter) + 1
'                            lRowChangeFound = RowChangesARRAY(lRowCounter)           '2010/02/11
'                            If bSAVEItemFieldsChanged(lRowChangeFound, rs, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctioncode, _
'                                   cITEM_XX & sCounter, _
'                                   sORIGAssortmentArray(lColCounter, 1), sNEWAssortmentArray(lColCounter, 1), _
'                                   dRevisionDate, gsUserID) = False Then
'                                '                           GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'                            End If
'                        End If
'                        'Assortment Quantity's
'                        If Not sORIGAssortmentArray(lColCounter, 2) = sNEWAssortmentArray(lColCounter, 2) Then
'                            RowChangesARRAY(lRowCounter) = RowChangesARRAY(lRowCounter) + 1
'                            lRowChangeFound = RowChangesARRAY(lRowCounter)           '2010/02/11
'                            '                    lRowChangeFound = RowChangesARRAY(lROWCounter)           '2010/02/11
'                            If bSAVEItemFieldsChanged(lRowChangeFound, rs, sProposalNumber, sRev, sPrevRev, sItemNumber, sFromProcess, sFunctioncode, _
'                                    cQTY_XX & sCounter, _
'                                   sORIGAssortmentArray(lColCounter, 2), sNEWAssortmentArray(lColCounter, 2), _
'                                   dRevisionDate, gsUserID) = False Then
'                                '                           GoTo ErrorHandler'TODO - GoTo Statements are redundant in .NET
'                            End If
'                        End If


'                    Next lColCounter
'                End If
'            End If
'        End If

'        rs.Close()
'        rs = Nothing
'        bCheckItemFieldsChangedPerRow = True
'ExitRoutine:
'        '    On Error Resume Next'TODO - On Error must be replaced with Try, Catch, Finally
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bCheckItemFieldsChangedPerRow")
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bCheckItemFieldsChangedPerRow, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        '    Resume Next '2010/02/12
'        Resume ExitRoutine
'    End Function

'    Private Function CharString(ByVal Value As Object, ByVal Characters As Integer, Optional ByVal PadBefore As Boolean = False)
'        ' This converts a value into a Char string with a given number of characters.
'        ' ET 2013-01-21 - moved here from modL-a-w-s-o-n because this is where all of the references are.

'        If IsNull(Value) Then
'            CharString = Space(Characters)
'        Else
'            ' Convert Value to a string and pad with any remaining spaces.
'            Value = Trim(Value)
'            Dim ValueLength As Integer
'            ValueLength = Len(CStr(Value))

'            ' This assumes that the length of the string does not exceed the available # of characters.
'            Dim SpacePadding As String
'            SpacePadding = Space(Characters - ValueLength)

'            If PadBefore Then
'                CharString = SpacePadding & CStr(Value)
'            Else
'                CharString = CStr(Value) & SpacePadding
'            End If
'        End If

'    End Function
'    Private Function bRollBackExcelSourceFile(ByVal lRow As Long, sItemArray_Orig() As String, sItemArray() As String, objExcel As Excel.Application, _
'                                                ByVal sfilename As String, ByVal lColsOnSheet As Long, dtSaveArrayCOLPos As typSpecialCOLPos, _
'                                                ByVal frmThis As Form, RowChangesARRAY() As String) As Boolean

'        '2014/01/07 RAS This is to roll back the updated spreadsheet to the original values and hopefully the original shading.
'        '
'        Dim sDefaultValue As String
'        'On Error GoTo ErrorHandler:'TODO - On Error must be replaced with Try, Catch, Finally
'        If UCase(sItemArray(lRow, 1)) = "A" Then
'            sItemArray(lRow, 1) = ""
'        Else
'            sDefaultValue = ""
'        End If
'        If UCase(sItemArray_Orig(lRow, 1)) = "R" Then
'            sItemArray(lRow, 1) = "T"
'        End If


'        ' 2014/01/13 RAS Added Proposalnumber to be put back also.
'        'For lCOLCounter = 1 To lColsOnSheet
'        Call bUpdateSpreadsheetCell(sDefaultValue, sItemArray_Orig(lRow, 3), sItemArray(lRow, 3), objExcel, lRow, 3, False)
'        objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, 3).Interior.Pattern = xlNone
'        Call bUpdateSpreadsheetCell(sDefaultValue, sItemArray_Orig(lRow, 1), sItemArray(lRow, 1), objExcel, lRow, 1, False)
'        objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, 1).Interior.Pattern = xlNone
'        Call bUpdateSpreadsheetCell(sDefaultValue, sItemArray_Orig(lRow, 2), sItemArray(lRow, 2), objExcel, lRow, 2, False)
'        objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRow, 2).Interior.Pattern = xlNone

'        'Next
'ExitFunction:

'        bRollBackExcelSourceFile = True
'        Exit Function
'ErrorHandler:
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bRollBackExcelSourceFile, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        bRollBackExcelSourceFile = False



'    End Function

'    Private Function bUpdateExcellogFile(ByVal lRow As Long, ByVal NewValue As String, ByVal OldValue As String, ByVal objExcel As Excel.Application, ByVal sfilename As String, ByVal lColsOnSheet As Long) As Object

'        '2014/01/07 RAS This is to update spreadsheet to the with the new message
'        '
'        'Dim lCOLCounter As Long
'        Dim sDefaultValue As String
'        Dim sRownumber As String
'        Dim lUpdateRowNumber As Long
'        Dim lRowCounter As Long
'        'On Error GoTo ErrorHandler:'TODO - On Error must be replaced with Try, Catch, Finally
'        sDefaultValue = ""
'        'this it to update the Excel Log file
'        ' the message cell on the log file is "N" 14th place
'        'need to search the log file in the first columm the the sourcefile row number and get that row number
'        ' Call bUpdateSpreadsheetCell(sDefaultValue, newvalue, oldvalue, objEXCEL, lRow, column, False)
'        For lRowCounter = 1 To glMAX_Rows
'            sRownumber = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lRowCounter, 1))
'            If sRownumber = CStr(lRow) Then
'                lUpdateRowNumber = lRowCounter
'                Exit For
'            End If
'            If sRownumber = "" Then Exit For
'        Next lRowCounter
'        If lUpdateRowNumber > 0 Then
'            Call bUpdateSpreadsheetCell(sDefaultValue, NewValue, OldValue, objExcel, lUpdateRowNumber, lColsOnSheet, False)
'        End If

'ExitFunction:
'        bUpdateExcellogFile = True
'        Exit Function
'ErrorHandler:
'        If glErrorMessageFlag = True Then  ''2014/01/15 RAS adding logging for the error messages
'            smessage = "In bUpdateExcellogFile, Err Number " & Err.Number & " ,Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If
'        bUpdateExcellogFile = False


'    End Function

'    Friend Overridable Function bValidateItemStatus(ByVal objExcel As Excel.Application, ByRef sErrorMsg As String, ByRef lNbrErrors As Long) As Boolean

'        'On Error GoTo ErrorHandler'TODO - On Error must be replaced with Try, Catch, Finally
'        Dim lRowsForImport As Long
'        Dim sFunctioncode As String
'        Dim sProposal As String
'        Dim sRev As String


'        'Dim lColorIndex  As Long      '2011/12/21 - need to check first 3 columns for no color! - no compile error??
'        Dim lRowCol1ColorIndex As Long
'        Dim lRowCol2ColorIndex As Long
'        Dim lRowCol3ColorIndex As Long
'        Dim lCounter As Long
'        Dim lColCounter As Long

'        Dim lItemStatusPOS As Long
'        Dim sItemStatus As String
'        Dim sColHeader As Object



'        Const sValidateMsg = "Validating Item Status Codes ... for row: "

'        bValidateItemStatus = False


'        For lColCounter = 1 To glMAX_Cols

'            sColHeader = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(1, lColCounter))
'            If sColHeader = "ItemStatus" Then
'                lItemStatusPOS = lColCounter
'                Exit For
'            End If
'        Next lColCounter

'        ' The first row contains field names, start on row 2
'        For lCounter = glDATA_START_ROW To lRowsOnSheet

'            If gbValidationCancelled = True Then Exit Function

'            lRowCol1ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glFunctionCode_ColPos).Interior.ColorIndex
'            sFunctioncode = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glFunctionCode_ColPos))
'            sItemStatus = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, lItemStatusPOS))
'            sProposal = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glProposal_ColPos))
'            lRowCol2ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glProposal_ColPos).Interior.ColorIndex

'            sRev = sGetCellValue(objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glREV_ColPos))
'            lRowCol3ColorIndex = objExcel.Application.Workbooks(1).Worksheets(1).Cells(lCounter, glREV_ColPos).Interior.ColorIndex

'            If sFunctioncode = gsENDofSpreadsheet Then
'                lRowsOnSheet = lCounter - glDATA_START_ROW + 1
'                Exit For

'            ElseIf Len(sFunctioncode) = 0 _
'                    And Len(sProposal) = 0 _
'                    And Len(sRev) = 0 _
'                    And (lRowCol1ColorIndex = xlNone And lRowCol2ColorIndex = xlNone And lRowCol3ColorIndex = xlNone) Then '2011/12/21
'                lRowsOnSheet = lCounter - glDATA_START_ROW + 1
'                Exit For

'            Else
'                If Len(sFunctioncode) > 0 Then
'                    Select Case sItemStatus
'                        Case "ORD"
'                            sErrorMsg = sErrorMsg & "ROW " & lCounter & ": (ItemStaus = ORD ) Permission Denied for changing records via Import Process." & vbCrLf
'                            lNbrErrors = lNbrErrors + 1
'                        Case Else
'                            ' not an orderd item so it is good
'                    End Select
'                Else
'                    Application.DoEvents()
'                End If

'                Application.DoEvents()
'            End If
'        Next lCounter

'        bValidateItemStatus = True
'ExitRoutine:
'        Exit Function
'ErrorHandler:

'        MsgBox(Err.Description, vbExclamation + vbMsgBoxSetForeground, "modSpreadSheet-bValidateItemStatus")
'        If glErrorMessageFlag = True Then
'            smessage = "bValidateItemStatus, Err Number " & Err.Number & "Error Description: " & Err.Description
'            If bWritePrintToLogFile(False, objEXCELName & smessage, "ErrorMessageLog") = False Then
'            End If
'        End If

'        Resume ExitRoutine









'    End Function

'    Private Function IsBlank(p1 As Object) As Boolean
'        Throw New NotImplementedException
'    End Function

'End Class


'Class typSpecialCOLPos

'End Class
