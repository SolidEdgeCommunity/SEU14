Imports System.Runtime.InteropServices
Module Main
    Public blnExcelAlreadyRunning As Boolean = False

    Public Sub Main()
        ' Get the type from the Solid Edge ProgID
        objSEType = Type.GetTypeFromProgID("SolidEdge.Application")
        

        OleMessageFilter.Register()

        'Add your code here!
        Dim objDraftDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim objPartsLists As SolidEdgeDraft.PartsLists = Nothing
        Dim objPartsList As SolidEdgeDraft.PartsList = Nothing
        Dim objActiveSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objSheets As SolidEdgeDraft.Sheets = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objExcelWorkBook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim objActiveExcelSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim objTableColumns As SolidEdgeDraft.TableColumns = Nothing
        Dim objTableColumn As SolidEdgeDraft.TableColumn = Nothing
        Dim oCell As SolidEdgeDraft.TableCell = Nothing
        Dim objTableCell As SolidEdgeDraft.TableCell = Nothing


        Dim blnUseCopyToClipboardAPI As Boolean = False
        Dim blnUseCopyCellByCellAPIs As Boolean = False

        Dim blnShowExcelDuringProcessing As Boolean = True
        Dim strExcelPath As String = String.Empty
        Dim strExcelFileName As String = String.Empty


        'this example code uses 2 different methods to extract the SE partslist
        'set the following booleans to control which method is run!
        blnUseCopyToClipboardAPI = False  ' simply copies the partslist object to the windows clipboard
        blnUseCopyCellByCellAPIs = True  'copies out each individual partslist cell value cell by cell for the entire partslist table
        blnShowExcelDuringProcessing = False




        


        'call function to connect to or start solid edge..  turn on visiblility and turn off display alerts
        If oConnectToSolidEdge(True, False) = True Then
            If objSEApp.ActiveDocumentType <> SolidEdgeFramework.DocumentTypeConstants.igDraftDocument Then
                MessageBox.Show("You must have a Solid Edge draft document opened", "Export SE Parts List", MessageBoxButtons.OK)
                Exit Sub
            End If
            Try
                'get the active document
                objDraftDoc = objSEApp.ActiveDocument

                'determine the excel filename and path for the excel
                'for this sample at this point in time use the same path as the opened draft file
                If objDraftDoc.Path <> String.Empty Then
                    strExcelPath = objDraftDoc.Path
                Else
                    strExcelPath = System.IO.Path.GetTempPath
                End If


                strExcelFileName = objDraftDoc.Name
                strExcelFileName = oGetFileNameWithoutExtension(strExcelFileName)
                strExcelFileName = strExcelPath + "\" + strExcelFileName

                'check to see if this spreadsheet already exists
                If oIsValidFileName(strExcelFileName + ".xls") = True Then
                    Try
                        System.IO.File.Delete(strExcelFileName + ".xls")
                    Catch ex As Exception
                        MessageBox.Show("Can not delete file: " + strExcelFileName, "Export SE Parts List", MessageBoxButtons.OK)
                        Exit Sub
                    End Try
                End If

                'check to see if this spreadsheet already exists
                If oIsValidFileName(strExcelFileName + ".xlsx") = True Then
                    Try
                        System.IO.File.Delete(strExcelFileName + ".xlsx")
                    Catch ex As Exception
                        MessageBox.Show("Can not delete file: " + strExcelFileName, "Export SE Parts List", MessageBoxButtons.OK)
                        Exit Sub
                    End Try
                End If

                'check to see if this spreadsheet already exists
                If oIsValidFileName(strExcelFileName + ".xlsm") = True Then
                    Try
                        System.IO.File.Delete(strExcelFileName + ".xlsm")
                    Catch ex As Exception
                        MessageBox.Show("Can not delete file: " + strExcelFileName, "Export SE Parts List", MessageBoxButtons.OK)
                        Exit Sub
                    End Try
                End If

                'find the first sheet containing a parts list 
                objPartsLists = objDraftDoc.PartsLists
                If objPartsLists.Count = 0 Then
                    MessageBox.Show("No partslists found in the document")
                    GoTo wrapup
                End If


                ' as currently written...  really assumes only one partslist per drawing.  If you need to support more than one partslist per draft
                ' you might want to add code to create a unique sheet in the xls for each of the partslist object found in the draft file


                For Each objPartsList In objPartsLists
                    'check to see if the partslist found is up to date...  if not update it
                    If objPartsList.IsUpToDate = False Then
                        objPartsList.Update()
                    End If

                    If blnUseCopyToClipboardAPI = True Then
                        'one method is to simply copy it to the clipboard and paste it as is into a spreadsheet
                        objPartsList.CopyToClipboard()
                        'create new excel doc
                        If oConnectToExcel(blnShowExcelDuringProcessing) = True Then
                            'add a new excel doc.
                            objExcelWorkBook = objExcel.Workbooks.Add()
                            objExcelWorkBook.ActiveSheet.paste()
                            'the excel file is left open....
                            GoTo wrapup
                        End If
                    End If


                    If blnUseCopyCellByCellAPIs = True Then  ' write out cell by cell example
                        If oConnectToExcel(blnShowExcelDuringProcessing) = True Then  ' this example uses an excel file....  could write to text file or anything for that matter
                            'add a new excel doc.
                            objExcelWorkBook = objExcel.Workbooks.Add()

                            ' set the active sheet in excel
                            objActiveExcelSheet = objExcelWorkBook.ActiveSheet

                            Dim objNumberOfPartsListColumns As Integer = 0
                            Dim objNumberOfPartsListRows As Integer = 0
                            Dim ii As Integer = 0
                            Dim jj As Integer = 0

                            'get the number of columns and rows in the partslist table object
                            objNumberOfPartsListColumns = objPartsList.Columns.Count
                            objNumberOfPartsListRows = objPartsList.Rows.Count

                            Dim strHeaderString As String = String.Empty
                            Dim ctr As Integer = 0

                            'take care of writing the header row.
                            For ii = 1 To objNumberOfPartsListColumns
                                objTableColumn = objPartsList.Columns.Item(ii)
                                'only write if the column is displayed in the partslist on the draft sheet
                                If objTableColumn.Show = True Then
                                    ctr = ctr + 1
                                    Try
                                        'get the header shown in the partslist on the draft sheet
                                        strHeaderString = objTableColumn.HeaderRowValue.ToString
                                        'write the header values to the 1st row in the spreadsheet
                                        objActiveExcelSheet.Cells(1, ctr).value = strHeaderString
                                    Catch ex As Exception
                                        'raise an exception
                                    End Try
                                End If
                            Next

                            ctr = 0  'reset

                            'loop through each row in the partslist
                            For ii = 1 To objNumberOfPartsListRows
                                'loop through each column shown in the partslist
                                For jj = 1 To objNumberOfPartsListColumns
                                    objTableColumn = objPartsList.Columns.Item(jj)
                                    'need to check to make sure the column is actually shown
                                    If objTableColumn.Show = True Then
                                        ctr = ctr + 1
                                        'hook up to each cell in the partslist (based on row,column)
                                        oCell = objPartsList.Cell(ii, jj)
                                        'let's write it to an excel sheet
                                        objActiveExcelSheet.Cells(ii + 1, ctr).value = oCell.value.ToString
                                    End If
                                Next
                                ctr = 0  ' reset
                            Next
                        End If
                    End If
                Next



                'now deal the excel 
                If blnExcelAlreadyRunning = False Then
                    'excel was not already running close the doc created and quit excel so it is not left running
                    objExcelWorkBook.SaveAs(strExcelFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal)
                    objExcelWorkBook.Close()
                    objExcel.Quit()
                ElseIf blnExcelAlreadyRunning = True Then
                    'excel was already running just close the doc created and leave excel running
                    objExcelWorkBook.SaveAs(strExcelFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal)
                    objExcelWorkBook.Close()

                End If

            Catch ex As Exception
                'raise exception
            End Try
        End If



        'reset the the display alerts status
        objSEApp.DisplayAlerts = True


wrapup:
        oReleaseObject(objPartsLists)
        oReleaseObject(objPartsList)
        oReleaseObject(objDraftDoc)
        oReleaseObject(objTableColumn)
        oReleaseObject(objTableCell)
        oReleaseObject(oCell)
        oReleaseObject(objSEApp)
        oReleaseObject(objExcelWorkBook)
        oReleaseObject(objActiveExcelSheet)
        oReleaseObject(objExcel)

        ' because of .NET
        OleMessageFilter.Revoke()

        End

    End Sub
End Module
