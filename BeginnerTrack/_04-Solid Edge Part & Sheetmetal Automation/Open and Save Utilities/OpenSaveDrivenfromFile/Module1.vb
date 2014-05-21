Imports System.Runtime.InteropServices
Imports System.IO

Module Module1
    Public FSO As Scripting.FileSystemObject = Nothing
    Public ObjSEApp As SolidEdgeFramework.Application = Nothing
    Public ObjSEAppType As Type
    Public ObjSEFilePropAPPType As Type
    Public intMaxNumberThenRestart As Integer = 0
    Public intSEECCount As Integer = 0
    Public StrFileNames As System.Collections.ArrayList
    Public intCtr As Integer = 0
    Public FSOLog As Scripting.FileSystemObject = Nothing
    Public txtStreamReportStatus As Scripting.TextStream = Nothing
    Public strSEInstalledPath As String = ""
    Public strSEVersion As String = ""
    Public SwitchExists As String = ""
    Public OpenSaveForm As Form
    Public blnRunFromCommandLine As Boolean = False
    Public blnCreatePDFsFromDrafts As Boolean = False
    Public strInputFilename As String = ""
    Public strLogFilename As String = ""
    Public strCommandLineCreatePDFs As String = ""
    Public strCommandLineInputFilename As String = ""
    Public strCommandLineInputUpdateDrawingViews As String = ""
    Public strCommandLineInputFitAndShade As String = ""
    Public strCommandLineInputRecompute As String = ""
    Public strCommandLineInputResetStatus As String = ""
    Public strCommandLineInputCheckAssembliesForCorruptLinks As String = ""
    Public DummySpreadsheetList As System.Collections.ArrayList
    Public blnRecompute As Boolean = False
    Public blnFitAndShade As Boolean = False
    Public strMacroPath As String = ""
    Public strMacroName As String = "DeleteInvalidSites.exe"
    Public strPathCMDLineOption As String = ""
    Public blnReadFilesFromFile As Boolean = False
    Public blnReadFilesFromFolder As Boolean = False
    Public blnUpdateDrawingViews As Boolean = False
    Public blnCheckAssemblyForCorruptLinks As Boolean = False
    Public blnResetStatus As Boolean = False
    Public blnTurnOffDisplayNextOccProp As Boolean = False
    Public blnCopyStyles As Boolean = False
    Public blnUpdateLinks As Boolean = False
    Public blnExtractPreviewBMP As Boolean = False
    Public blnRemoveProperties As Boolean = False
    Public strPropsToRemove As String = ""
    Public CustomPropertiesToDelete As System.Collections.ArrayList
    Public strCommandTurnDisplayOffNextHighest As String = ""
    Public oArrayOfOccsVisibleInFrontView As ArrayList = Nothing
    Public oArrayOfOccsVisibleInBackView As ArrayList = Nothing
    Public oArrayOfOccsVisibleInTopView As ArrayList = Nothing
    Public oArrayOfOccsVisibleInBottomView As ArrayList = Nothing
    Public oArrayOfOccsVisibleInLeftView As ArrayList = Nothing
    Public oArrayOfOccsVisibleInRightView As ArrayList = Nothing
    Public oArrayofTotallyHiddenOccs As ArrayList = Nothing
    Public oArrayofVisibleOccs As ArrayList = Nothing
    Public objSelectSet As SolidEdgeFramework.SelectSet = Nothing
    Public blnTurnOffCss As Boolean = False
    Public strAppVersion As String = String.Empty
    Public blnReplaceCharacters As Boolean = False
    Public strOldChar As String = String.Empty
    Public strNewChar As String = String.Empty
    Public blnFindReplaceCalloutBOM As Boolean = False
    Public arrDataLinkedPropertyText As ArrayList
    Public arrSearchForLinkedPropertyText As ArrayList
    Public arrReplaceWithLinkedPropertyText As ArrayList
    Public blnProcessASMs As Boolean = False
    Public blnProcessPARs As Boolean = False
    Public blnProcessPSMs As Boolean = False
    Public blnProcessDFTs As Boolean = False
    Public blnProcessPWDs As Boolean = False
    Public blnTurnOffGradientBackground As Boolean = False
    Public blnCheckIfAlreadyImported As Boolean = False
    Public arrayOfImportedFiles As ArrayList = Nothing
    Public arrayOfNotImportedFiles As ArrayList = Nothing
    Public blnCheckHardwareOption As Boolean = False
    Public strHardwarePropertyNameToCheck As String = String.Empty
    Public strHardwarePropertyValue As String = String.Empty
    Public stringArray() As String
    Public IntProcessHardwareCTR As Integer = 0
    Public blnCreatePreview As Boolean = False
    Public blnResetBodyStyle As Boolean = False




    Sub Garbage_Collect(ByVal obj As Object)
        Try
            If Not (obj Is Nothing) Then
                'Marshal.ReleaseComObject(obj)  ' this leaves a file lock when using SolidEdgeFileProperties API
                Do While (Marshal.ReleaseComObject(obj) > 0)
                Loop
                obj = Nothing
            End If
            GC.Collect(GC.MaxGeneration)
            GC.WaitForPendingFinalizers()
        Catch ex As Exception
            obj = Nothing
            GC.Collect(GC.MaxGeneration)
            GC.WaitForPendingFinalizers()
        End Try

    End Sub
    Sub Garbage_CollectFinal(ByVal obj As Object)
        Try
            '******* Added because of .NET
            If Not (obj Is Nothing) Then
                'Marshal.ReleaseComObject(obj)
                Marshal.FinalReleaseComObject(obj)
            End If
            GC.Collect(GC.MaxGeneration)
            GC.WaitForPendingFinalizers()
            GC.Collect(GC.MaxGeneration)
            GC.WaitForPendingFinalizers()
            '******* Added because of .NET
        Catch ex As Exception
            obj = Nothing
            GC.Collect(GC.MaxGeneration)
            GC.WaitForPendingFinalizers()
            GC.Collect(GC.MaxGeneration)
            GC.WaitForPendingFinalizers()
        End Try

    End Sub
    Public Sub Main(ByVal args As String())

        ObjSEAppType = Type.GetTypeFromProgID("SolidEdge.Application")

        ObjSEFilePropAPPType = Type.GetTypeFromProgID("SolidEdge.FileProperties")
        '"C:\_Work\VB\VB Programs\OpenSaveFromTextfile\Batch open.txt"





        If args.Length = 0 Then   ' then NOT running from commandline
            blnRunFromCommandLine = False
            StrFileNames = New ArrayList
            'must not be command line driven so show the form
            If OpenSaveForm Is Nothing Then
                OpenSaveForm = New Form1
                OpenSaveForm.ShowDialog()
                OpenSaveForm.TopMost = True
            Else
                OpenSaveForm.ShowDialog()
                OpenSaveForm.TopMost = True
            End If
            Exit Sub
        End If

        blnRunFromCommandLine = True

        blnReadFilesFromFile = True 'if run via commandline must use this option
        blnReadFilesFromFolder = False  'if run via commandline must use this option


        Try
            For Each arg As String In args
                ParseCommandlineInput(arg.ToString)
            Next arg
            strInputFilename = strCommandLineInputFilename
            strLogFilename = GetFilePath(strInputFilename) + "\OpenSaveLog.txt"


            '***** below gathers the input arguments for the supported options

            If strCommandLineCreatePDFs.ToUpper = "TRUE" Then
                blnCreatePDFsFromDrafts = True
            End If
            If strCommandLineCreatePDFs.ToUpper = "FALSE" Then
                blnCreatePDFsFromDrafts = False
            End If



            If strCommandLineInputUpdateDrawingViews.ToUpper = "TRUE" Then
                blnUpdateDrawingViews = True
            End If
            If strCommandLineInputUpdateDrawingViews.ToUpper = "FALSE" Then
                blnUpdateDrawingViews = False
            End If


            If strCommandLineInputFitAndShade.ToUpper = "TRUE" Then
                blnFitAndShade = True
            End If

            If strCommandLineInputFitAndShade.ToUpper = "FALSE" Then
                blnFitAndShade = False
            End If



            If strCommandLineInputRecompute.ToUpper = "TRUE" Then
                blnRecompute = True
            End If
            If strCommandLineInputRecompute.ToUpper = "FALSE" Then
                blnRecompute = False
            End If


            If strCommandLineInputCheckAssembliesForCorruptLinks.ToUpper = "TRUE" Then
                blnCheckAssemblyForCorruptLinks = True
            End If
            If strCommandLineInputCheckAssembliesForCorruptLinks.ToUpper = "FALSE" Then
                blnCheckAssemblyForCorruptLinks = False
            End If


            If strCommandLineInputResetStatus.ToUpper = "TRUE" Then
                blnResetStatus = True
            End If
            If strCommandLineInputResetStatus.ToUpper = "FALSE" Then
                blnResetStatus = False
            End If



        Catch ex As Exception

        End Try


        Call Process()



    End Sub


    Public Sub WriteToLogFile(ByVal oLogFileName As String, ByVal oStringToWrite As String)
        System.IO.File.AppendAllText(oLogFileName, oStringToWrite + Environment.NewLine)
    End Sub


    Public Sub Process()

        Dim ii As Long = 0
        Dim pp As Integer
        Dim objDoc As Object = Nothing
        Dim obj3dWindow As SolidEdgeFramework.Window = Nothing
        Dim obj2dWindow As SolidEdgeDraft.SheetWindow = Nothing

        Dim strtempFilename As String = ""
        'Dim objView As SolidEdgeFramework.View = Nothing
        Dim vString As String




        Dim objModels As SolidEdgePart.Models = Nothing
        Dim objModel As SolidEdgePart.Model = Nothing
        Dim objFlatPatternModels As SolidEdgePart.FlatPatternModels = Nothing
        Dim objFlatPatternModel As SolidEdgePart.FlatPatternModel = Nothing


        Dim blnCorruptLinkFileProcessed As Boolean = False

        '******* Added because of .NET
        Try
            OleMessageFilter.Register()
        Catch ex As Exception
            'PrintLine("Error registering message filter.")
        End Try
        '******* Added because of .NET


        GetSolidEdgePath()



        If createLogFile() = False Then
            'MessageBox.Show("Error creating the log file " + Form1.TxtLogFileName.Text, "SE Utilities", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        intMaxNumberThenRestart = 200

        If blnReadFilesFromFile = True Then
            If UCase(strInputFilename).EndsWith("TXT") Then
                ReadTextFile()
            End If

            If UCase(strInputFilename).EndsWith("CSV") Then
                DummySpreadsheetList = New ArrayList
                '   add function to red filenames plus path from the xml ordered list ReadTextFile()
                ReadFilenamesFromCSVFile(strInputFilename)
            End If
        End If


        'already got a list of files from recursing the folder specified in the dialog....  since running the UI (not via commandline input)

        'make sure Solid Edge is not running since we will be setting registry stuff
        Try
            ObjSEApp = Marshal.GetActiveObject("SolidEdge.Application")
            Do While ObjSEApp.Documents.Count <> 0
                ObjSEApp.DisplayAlerts = False
                ObjSEApp.ActiveDocument.close()
            Loop
            ObjSEApp.Quit()
            Garbage_Collect(ObjSEApp)
        Catch ex As Exception

        End Try


        '****************set the necessary registry stuff*************************************

        ' Set the assembly mode to open assemblies with Occurrences inactive
        SetAssemblyMode(1)

        ' Set the assembly mode to open assemblies with all Occurrences active.
        'SetAssemblyMode(2)

        'check to see if registry location exists
        SwitchExists = QueryValue(HKEY_CURRENT_USER, "Software\Unigraphics Solutions\Solid Edge\Version " & strSEVersion & "\Debug", "DocMgmt_OverrideStatusCheckForFileAccess")

        If SwitchExists = "" Then
            '-- Create key in CURRENT USER and set it
            SetKeyValue(HKEY_CURRENT_USER, "Software\Unigraphics Solutions\Solid Edge\Version " & strSEVersion & "\Debug", "DocMgmt_OverrideStatusCheckForFileAccess", "1", REG_DWORD)
        End If

        If SwitchExists = "0" Then
            '-- Create key in CURRENT USER and set it
            SetKeyValue(HKEY_CURRENT_USER, "Software\Unigraphics Solutions\Solid Edge\Version " & strSEVersion & "\Debug", "DocMgmt_OverrideStatusCheckForFileAccess", "1", REG_DWORD)
        End If

        If blnCheckAssemblyForCorruptLinks = True Then
            'need to set the registry switch to enable this....
            'check to see if registry location exists
            SwitchExists = QueryValue(HKEY_CURRENT_USER, "Software\Unigraphics Solutions\Solid Edge\Version " & strSEVersion & "\Debug", "WRITE_BADLINKS_TO_FILE")
            If SwitchExists = "" Then
                '-- Create key in CURRENT USER and set it
                SetKeyValue(HKEY_CURRENT_USER, "Software\Unigraphics Solutions\Solid Edge\Version " & strSEVersion & "\Debug", "WRITE_BADLINKS_TO_FILE", "1", REG_DWORD)
            End If
            If SwitchExists = "0" Then
                '-- Create key in CURRENT USER and set it
                SetKeyValue(HKEY_CURRENT_USER, "Software\Unigraphics Solutions\Solid Edge\Version " & strSEVersion & "\Debug", "WRITE_BADLINKS_TO_FILE", "1", REG_DWORD)
            End If
        End If

      


        '********************here is the big loop to process the files collected either from the dialog input( folder or textfile)
        Try
            For ii = 0 To StrFileNames.Count - 1

                If System.IO.File.Exists(StrFileNames(ii)) = False Then
                    GoTo skip
                End If
                
                If oIsSolidEdgeAssemblylFile(StrFileNames.Item(ii)) Then
                    If blnProcessASMs = False Then
                        GoTo skip
                    End If
                End If

                If oIsSolidEdgePartFile(StrFileNames.Item(ii)) Then
                    If blnProcessPARs = False Then
                        GoTo skip
                    End If
                End If

                If oIsSolidEdgeSheetMetalFile(StrFileNames.Item(ii)) Then
                    If blnProcessPSMs = False Then
                        GoTo skip
                    End If
                End If

                If oIsSolidEdgeDraftFile(StrFileNames.Item(ii)) Then
                    If blnProcessDFTs = False Then
                        GoTo skip
                    End If
                End If

                If oIsSolidEdgeWeldmentFile(StrFileNames.Item(ii)) Then
                    If blnProcessPWDs = False Then
                        GoTo skip
                    End If
                End If

                If blnRunFromCommandLine = False Then
                    If ii = 0 Then
                        OpenSaveForm.Controls.Item("Label2").Text = "processing 1 of " + StrFileNames.Count.ToString + " " + StrFileNames.Item(ii)
                        OpenSaveForm.Controls.Item("Label2").Refresh()
                    Else
                        OpenSaveForm.Controls.Item("Label2").Text = "processing " + ii.ToString + " of " + StrFileNames.Count.ToString + " " + StrFileNames.Item(ii)
                        OpenSaveForm.Controls.Item("Label2").Refresh()
                    End If
                End If

                If blnRunFromCommandLine = True Then
                    'Console.WriteLine("processing " + ii.ToString + " of " + StrFileNames.Count.ToString + " " + StrFileNames.Item(ii))
                End If

                If (GetAttr(StrFileNames(ii)) And FileAttribute.ReadOnly) <> 0 Then
                    txtStreamReportStatus.WriteLine("ERROR->ReadOnlyFile:" + StrFileNames.Item(ii))
                    GoTo skip
                End If


                If blnCheckHardwareOption = True Then
                    If strHardwarePropertyNameToCheck <> "" Then
                        'call function to do the work....
                        ProcessHardwareCheck(StrFileNames.Item(ii))
                        GoTo skip  'skip to the next file

                    Else
                        MessageBox.Show("You must specify the Solid Edge property name and value to check if hardware")
                        End
                    End If
                End If


                If blnResetStatus = True Then
                    If ResetSEFileStatusToAvailable(StrFileNames.Item(ii)) = False Then
                        GoTo skip
                    End If
                End If

                ' '' ''If UCase(StrFileNames(ii)).EndsWith("ASM") Then
                ' '' ''    ' To save time, rather than processing all assemblies, when
                ' '' ''    ' OpenSave is running, Assembly will save all subassemblies
                ' '' ''    ' when it saves a top-level assembly  0 is normal behavior 1 forces a save on subs
                ' '' ''    SetOpenSaveMacroFlag(0)
                ' '' ''End If


                If blnCheckIfAlreadyImported = True Then


                    If ProcessFilePropertiesToCheckIfImported(StrFileNames.Item(ii)) = True Then
                        arrayOfImportedFiles.Add(StrFileNames.Item(ii))
                        'txtStreamReportStatus.WriteLine(StrFileNames.Item(ii))
                    Else
                        arrayOfNotImportedFiles.Add(StrFileNames.Item(ii))
                    End If

                    GoTo skip
                End If


                If blnFindReplaceCalloutBOM = True Then  'read the linked property text for find/replace
                    Dim kk As Integer = 0
                    For kk = 0 To arrDataLinkedPropertyText.Count - 1
                        Dim strData As String = arrDataLinkedPropertyText(kk)
                        Dim arrStrDataSplit As String() = strData.Split(",")
                        If arrStrDataSplit(0) <> "" And arrStrDataSplit(1) <> "" Then
                            arrSearchForLinkedPropertyText.Add(arrStrDataSplit(0))
                            arrReplaceWithLinkedPropertyText.Add(arrStrDataSplit(1))
                        End If
                    Next
                End If


                If blnReplaceCharacters = True Then
                    Try
                        Dim origFName As String = StrFileNames.Item(ii)
                        Dim strFileNameOnly As String = System.IO.Path.GetFileNameWithoutExtension(origFName)

                        If strFileNameOnly.Contains(strOldChar) = False Then
                            GoTo skip
                        End If

                        Dim strExtensionOnly As String = System.IO.Path.GetExtension(origFName)
                        Dim strPathOnly As String = System.IO.Path.GetDirectoryName(origFName)
                        Dim strOrigCFGFilename As String = String.Empty

                        If strExtensionOnly.ToLower = ".asm" Then
                            strOrigCFGFilename = strPathOnly + "\" + strFileNameOnly + ".cfg"
                        End If

                        If strFileNameOnly.StartsWith("..") Then
                            strFileNameOnly = strFileNameOnly.Replace("..", "")
                            GoTo skip2
                        End If

                        If strFileNameOnly.EndsWith("..") Then
                            strFileNameOnly = strFileNameOnly.Replace("..", "")
                            GoTo skip2
                        End If

                        If strFileNameOnly.StartsWith(".") Then
                            strFileNameOnly = strFileNameOnly.Replace(".", "")
                            GoTo skip2
                        End If

                        If strFileNameOnly.EndsWith(".") Then
                            strFileNameOnly = strFileNameOnly.Replace(".", "")
                            GoTo skip2
                        End If


                        If strFileNameOnly.StartsWith("_") Then
                            strFileNameOnly = strFileNameOnly.Replace("_", "")
                            GoTo skip2
                        End If
                        strFileNameOnly = strFileNameOnly.Replace(strOldChar, strNewChar)
skip2:
                        System.IO.File.Move(origFName, strPathOnly + "\" + strFileNameOnly + strExtensionOnly)

                        txtStreamReportStatus.WriteLine("Renamed file from  : " + StrFileNames.Item(ii) + "  to " + strPathOnly + "\" + strFileNameOnly + strExtensionOnly)

                        If strExtensionOnly.ToLower = ".asm" Then
                            'check to make sure cfg exists first
                            If System.IO.File.Exists(strOrigCFGFilename) = True Then
                                System.IO.File.Move(strOrigCFGFilename, strPathOnly + "\" + strFileNameOnly + ".cfg")
                                txtStreamReportStatus.WriteLine("Renamed file from  : " + strOrigCFGFilename + "  to " + strPathOnly + "\" + strFileNameOnly + ".cfg")
                            Else
                                txtStreamReportStatus.WriteLine("Assembly cfg not found for " + StrFileNames.Item(ii))
                            End If
                        End If

                        GoTo skip
                    Catch ex As Exception
                        txtStreamReportStatus.WriteLine("error processing file " + StrFileNames.Item(ii) + " -> " + ex.Message)
                    End Try

                End If

                Dim strExtension As String = ParseExtension(StrFileNames.Item(ii))
                If IsValdidFileToProcess(strExtension) = False Then
                    GoTo skip
                End If






                If blnExtractPreviewBMP = True Then
                    Try
                        Dim junk As Object = Nothing

                        Dim oSEThumb As SeThumbnailLib.SeThumbnailExtractor
                        oSEThumb = CreateObject("SeThumbnail.SeThumbnailExtractor")
                        Dim hndlImage As Integer
                        oSEThumb.GetThumbnail(StrFileNames(ii), hndlImage, junk)

                        'got the handle now what to do with it!!!!!

                        Dim objBMP As System.Drawing.Bitmap = System.Drawing.Image.FromHbitmap(hndlImage)

                        Dim strtmpFname As String = StrFileNames(ii)
                        strtmpFname = System.IO.Path.ChangeExtension(strtmpFname, ".bmp")
                        objBMP.Save(filename:=strtmpFname)

                        Module1.Garbage_Collect(oSEThumb)
                        Module1.Garbage_Collect(objBMP)
                    Catch ex As Exception
                        MessageBox.Show("error getting bmp preview" + ex.Message)
                        End
                    End Try
                    GoTo skip
                End If  'end of extract BMP



                If blnRemoveProperties = True Then
                    ProcessFileProperties(StrFileNames(ii))
                End If  ' end of Remove Properties


                If FileInUse(StrFileNames(ii)) = True Then
                    'Beep()
                End If



                Try
                    If ConnectToSolidEdge(intCtr) = True Then
                        ObjSEApp.DoIdle()
                        ObjSEApp.WindowState = 2  'maximize the app window

                        'open the current document in the loop
                        objDoc = ObjSEApp.Documents.Open(StrFileNames.Item(ii))

                        If objDoc.readonly = True Then
                            txtStreamReportStatus.WriteLine("Warning Read-Only : File: " + StrFileNames.Item(ii))
                            GoTo SkipToReadONly
                        End If

                        If objDoc.type = SolidEdgeFramework.DocumentTypeConstants.igDraftDocument Then
                            obj2dWindow = ObjSEApp.ActiveWindow
                            If blnFitAndShade = True Then
                                obj2dWindow.FitEx(SolidEdgeDraft.SheetFitConstants.igFitSheet)
                            End If
                            objDoc.UpdatePropertyTextDisplay()
                            If blnCreatePDFsFromDrafts = True Then
                                ' then create the PDF of draft files ONLY
                                strtempFilename = Mid(StrFileNames.Item(ii), 1, InStrRev(StrFileNames.Item(ii), ".", , CompareMethod.Text))
                                ObjSEApp.ActiveDocument.saveas(strtempFilename + "pdf")
                            End If
                            If blnUpdateDrawingViews = True Then
                                ObjSEApp.StartCommand(SolidEdgeConstants.DetailCommandConstants.DetailDrawingViewsUpdateViews)
                            End If


                            


                            If blnCheckAssemblyForCorruptLinks = True Then
                                blnCorruptLinkFileProcessed = False
                                Try
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")
                                    End If
                                    ObjSEApp.StartCommand(64000)
                                    ObjSEApp.DoIdle()   'during my testing seemed to be a timing thing... sometime did not report an issue  this API or the sleep below seemed to make it work... need to check into this

                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        blnCorruptLinkFileProcessed = True
                                        Dim strInvalidLinks As String = ""
                                        strInvalidLinks = ReadInvalidLinksFile(strTempFolder + "InvalidLinks.txt")
                                        txtStreamReportStatus.WriteLine("CORRUPT LINKS FOUND IN FILE : " + StrFileNames.Item(ii) + " invalid links are->" + strInvalidLinks)
                                    End If
                                    System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")  'remove it
                                Catch ex As Exception
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    txtStreamReportStatus.WriteLine("Error could not delete the file : " + strTempFolder + "InvalidLinks.txt")
                                    GoTo skip
                                End Try
                            End If


                            If blnFindReplaceCalloutBOM = True Then
                                Dim objPartsLists As SolidEdgeDraft.PartsLists = Nothing
                                Dim objPartsList As SolidEdgeDraft.PartsList = Nothing
                                Try
                                    objPartsLists = objDoc.PartsLists
                                    Dim objPartsListColumns As SolidEdgeDraft.TableColumns = Nothing
                                    Dim objPartsListColumn As SolidEdgeDraft.TableColumn = Nothing
                                    For Each objPartsList In objPartsLists
                                        objPartsListColumns = objPartsList.Columns
                                        Dim ll As Integer = 0
                                        For Each objPartsListColumn In objPartsListColumns
                                            For ll = 0 To arrReplaceWithLinkedPropertyText.Count - 1


                                                Dim strLinkedPropertyText As String = objPartsListColumn.PropertyText.ToString
                                                Dim strSearchForLinkedPropText As String = arrSearchForLinkedPropertyText(ll)


                                                If strLinkedPropertyText.Contains("pset=0pid=2") And strSearchForLinkedPropText.Contains("Title") Then  'means Title
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Title", "pset=0pid=2")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Title", "pset=0pid=2")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=0pid=4") And strSearchForLinkedPropText.Contains("Author") Then  'means Author
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Author", "pset=0pid=4")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Author", "pset=0pid=4")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=0pid=3") And strSearchForLinkedPropText.Contains("Subject") Then  'means Subject
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Subject", "pset=0pid=3")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Subject", "pset=0pid=3")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=0pid=5") And strSearchForLinkedPropText.Contains("Keywords") Then  'means keywords
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Keywords", "pset=0pid=5")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Keywords", "pset=0pid=5")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=0pid=6") And strSearchForLinkedPropText.Contains("Comments") Then  'means Comments
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Comments", "pset=0pid=6")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Comments", "pset=0pid=6")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=2pid=14") And strSearchForLinkedPropText.Contains("Manager") Then  'means Manager
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Manager", "pset=2pid=14")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Manager", "pset=2pid=14")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=1pid=15") And strSearchForLinkedPropText.Contains("Company") Then  'means Company
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Company", "pset=1pid=15")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Company", "pset=1pid=15")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=2pid=2") And strSearchForLinkedPropText.Contains("Category") Then  'means Category
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Category", "pset=2pid=2")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Category", "pset=2pid=2")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=1pid=1001") And strSearchForLinkedPropText.Contains("Status") Then  'means Status
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Status", "pset=1pid=1001")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Status", "pset=1pid=1001")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=3pid=12") And strSearchForLinkedPropText.Contains("Document Number") Then  'means Document Number
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Document Number", "pset=3pid=12")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Document Number", "pset=3pid=12")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=3pid=3") And strSearchForLinkedPropText.Contains("Revision Number") Then  'means Revision Number
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Revision Number", "pset=3pid=3")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Revision Number", "pset=3pid=3")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf strLinkedPropertyText.Contains("pset=3pid=4") And strSearchForLinkedPropText.Contains("Project Name") Then  'means Project Name
                                                    strLinkedPropertyText = strLinkedPropertyText.Replace("Project Name", "pset=3pid=4")
                                                    strSearchForLinkedPropText = strLinkedPropertyText.Replace("Project Name", "pset=3pid=4")

                                                    If strLinkedPropertyText = strSearchForLinkedPropText Then
                                                        objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                    End If

                                                ElseIf objPartsListColumn.PropertyText = arrSearchForLinkedPropertyText(ll) Then
                                                    objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                End If
                                                '' ''If objPartsListColumn.PropertyText = arrSearchForLinkedPropertyText(ll) Then
                                                '' ''    objPartsListColumn.PropertyText = arrReplaceWithLinkedPropertyText(ll)
                                                '' ''End If
                                            Next
                                        Next
                                        objPartsList.Update()
                                        Garbage_Collect(objPartsListColumn)
                                        Garbage_Collect(objPartsListColumns)
                                    Next
                                    Garbage_Collect(objPartsList)
                                    Garbage_Collect(objPartsLists)
                                Catch ex As Exception
                                    txtStreamReportStatus.WriteLine("Error replacing BOM column : " + StrFileNames.Item(ii) + "err->" + ex.Message)
                                    'GoTo skip
                                End Try

                                'code the fix callouts start here!
                                Dim oSheets As SolidEdgeDraft.Sheets = Nothing
                                Dim oSheet As SolidEdgeDraft.Sheet = Nothing
                                Dim objActiveSheet As SolidEdgeDraft.Sheet = Nothing
                                Dim objCallouts As SolidEdgeFrameworkSupport.Balloons = Nothing
                                Dim objCallout As SolidEdgeFrameworkSupport.Balloon = Nothing
                                Try
                                    objActiveSheet = ObjSEApp.ActiveDocument.activesheet
                                    oSheets = ObjSEApp.ActiveDocument.sheets
                                    For Each oSheet In oSheets
                                        oSheet.Activate()
                                        objCallouts = oSheet.Balloons
                                        Dim kk As Integer = 0
                                        For Each objCallout In objCallouts
                                            If objCallout.Callout = 1 Then
                                                Dim oStringDisplayedInSECallout As String = objCallout.BalloonDisplayedText
                                                Dim strLinkedPropertyText As String = objCallout.BalloonText
                                                For kk = 0 To arrReplaceWithLinkedPropertyText.Count - 1
                                                    If strLinkedPropertyText = arrSearchForLinkedPropertyText(kk) Then
                                                        objCallout.BalloonText = arrReplaceWithLinkedPropertyText(kk)
                                                    End If
                                                Next
                                            End If
                                            ''If oSheet.Name <> objActiveSheet.Name Then
                                            ''    oSheet.Visible = False
                                            ''End If

                                        Next
                                        Garbage_Collect(objCallout)
                                        Garbage_Collect(objCallouts)
                                    Next
                                    Garbage_Collect(oSheet)
                                    Garbage_Collect(oSheets)


                                    objActiveSheet.Activate()

                                    Garbage_Collect(objActiveSheet)

                                    ' Dim objDraftDoc As SolidEdgeDraft.DraftDocument = Nothing
                                    'objDraftDoc = objDoc
                                    Dim oSections As SolidEdgeDraft.Sections = objDoc.Sections
                                    Dim oSection As SolidEdgeDraft.Section = Nothing

                                    For kk = 1 To oSections.Count
                                        oSection = oSections.Item(kk)
                                        If oSection.Type = SolidEdgeDraft.SheetSectionTypeConstants.igWorkingSection Then

                                        Else
                                            oSection.Deactivate()
                                        End If
                                    Next

                                    'Garbage_Collect(objDraftDoc)
                                    Garbage_Collect(oSection)
                                    Garbage_Collect(oSections)

                                Catch ex As Exception
                                    txtStreamReportStatus.WriteLine("Error replacing Callout Linked Text in file: " + StrFileNames.Item(ii) + "err->" + ex.Message)
                                End Try

                            End If

                            If blnCreatePreview = True Then
                                Try
                                    objDoc.createpreview()
                                Catch ex As Exception

                                End Try
                            End If

                        End If  ' end of if draft document

                        vString = CStr(objDoc.lastsavedversion)
                        If objDoc.type = SolidEdgeConstants.DocumentTypeConstants.igPartDocument Then

                            Try
                                obj3dWindow = ObjSEApp.ActiveWindow
                                obj3dWindow.WindowState = 2
                            Catch ex As Exception

                            End Try

                            If blnResetBodyStyle = True Then
                                Try

                                    Dim objModels1 As SolidEdgePart.Models = Nothing
                                    Dim objModel1 As SolidEdgePart.Model = Nothing
                                    Dim objBody As SolidEdgeGeometry.Body = Nothing

                                    objModels1 = objDoc.models

                                    For Each objModel1 In objModels1
                                        If objModel1.IsModelActive = True Then
                                            objBody = objModel1.Body
                                            objBody.Style = Nothing
                                        End If
                                    Next

                                    Garbage_Collect(objBody)
                                    Garbage_Collect(objModel1)
                                    Garbage_Collect(objModels1)


                                   
                                Catch ex As Exception

                                End Try
                            End If


                            If blnCopyStyles = True Then
                                UpdatePartDocumentStyle(objDoc)
                            End If

                            Try
                                If blnRecompute = True Then
                                    objDoc.recompute()
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If blnFitAndShade = True Then
                                    ObjSEApp.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewISOView)
                                    ObjSEApp.StartCommand(SolidEdgeConstants.PartCommandConstants.PartReferencePlaneHideAllReferencePlanes)
                                    obj3dWindow.View.RenderModeType = SolidEdgeFramework.SeRenderModeType.seRenderModeSmoothVHL
                                    ObjSEApp.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewFit)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If blnTurnOffCss = True Then
                                    ObjSEApp.ActiveDocument.CoordinateSystems.Visible = False
                                End If
                            Catch ex As Exception

                            End Try


                            Try
                                If blnTurnOffGradientBackground = True Then
                                    Dim objSEViewStyle As SolidEdgeFramework.ViewStyle = Nothing
                                    Dim objView As SolidEdgeFramework.View = Nothing

                                    objView = ObjSEApp.ActiveWindow.view
                                    objSEViewStyle = objView.ViewStyle
                                    objSEViewStyle.BackgroundType = SolidEdgeFramework.SeBackgroundType.seBackgroundTypeSolid

                                    Garbage_Collect(objSEViewStyle)
                                    Garbage_Collect(objView)
                                End If
                            Catch ex As Exception

                            End Try

                            If blnCreatePreview = True Then
                                Try
                                    objDoc.createpreview()
                                Catch ex As Exception

                                End Try
                            End If

                            If blnCheckAssemblyForCorruptLinks = True Then
                                blnCorruptLinkFileProcessed = False
                                Try
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")
                                    End If
                                    ObjSEApp.StartCommand(64000)
                                    ObjSEApp.DoIdle()   'during my testing seemed to be a timing thing... sometime did not report an issue  this API or the sleep below seemed to make it work... need to check into this

                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        blnCorruptLinkFileProcessed = True
                                        Dim strInvalidLinks As String = ""
                                        strInvalidLinks = ReadInvalidLinksFile(strTempFolder + "InvalidLinks.txt")
                                        txtStreamReportStatus.WriteLine("CORRUPT LINKS FOUND IN FILE : " + StrFileNames.Item(ii) + " invalid links are->" + strInvalidLinks)
                                    End If
                                    System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")  'remove it
                                Catch ex As Exception
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    txtStreamReportStatus.WriteLine("Error could not delete the file : " + strTempFolder + "InvalidLinks.txt")
                                    GoTo skip
                                End Try
                            End If
                        End If



                        If objDoc.type = SolidEdgeConstants.DocumentTypeConstants.igSheetMetalDocument Then
                            Try
                                obj3dWindow = ObjSEApp.ActiveWindow
                                obj3dWindow.WindowState = 2
                            Catch ex As Exception

                            End Try

                            If blnResetBodyStyle = True Then
                                Try

                                    Dim objModels1 As SolidEdgePart.Models = Nothing
                                    Dim objModel1 As SolidEdgePart.Model = Nothing
                                    Dim objBody As SolidEdgeGeometry.Body = Nothing

                                    objModels1 = objDoc.models

                                    For Each objModel1 In objModels1
                                        If objModel1.IsModelActive = True Then
                                            objBody = objModel1.Body
                                            objBody.Style = Nothing
                                        End If
                                    Next

                                    Garbage_Collect(objBody)
                                    Garbage_Collect(objModel1)
                                    Garbage_Collect(objModels1)



                                Catch ex As Exception

                                End Try
                            End If
                           

                            If blnCopyStyles = True Then
                                UpdateSheetMetalDocumentStyle(objDoc)
                            End If

                            Try
                                If blnRecompute = True Then
                                    objDoc.recompute()
                                    objFlatPatternModels = objDoc.flatpatternmodels
                                    If objFlatPatternModels.Count > 0 Then
                                        For pp = 1 To objFlatPatternModels.Count
                                            objFlatPatternModel = objFlatPatternModels.Item(pp)
                                            objFlatPatternModel.Recompute()
                                            objFlatPatternModel.Update()
                                        Next pp
                                    End If
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If blnFitAndShade = True Then
                                    ObjSEApp.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewISOView)
                                    ObjSEApp.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalReferencePlaneHideAllReferencePlanes)
                                    obj3dWindow.View.RenderModeType = SolidEdgeFramework.SeRenderModeType.seRenderModeSmoothVHL
                                    ObjSEApp.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewFit)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If blnTurnOffGradientBackground = True Then
                                    Dim objSEViewStyle As SolidEdgeFramework.ViewStyle = Nothing
                                    Dim objView As SolidEdgeFramework.View = Nothing
                                    objView = ObjSEApp.ActiveWindow.view
                                    objSEViewStyle = objView.ViewStyle
                                    objSEViewStyle.BackgroundType = SolidEdgeFramework.SeBackgroundType.seBackgroundTypeSolid
                                    objSEViewStyle.
                                    Garbage_Collect(objSEViewStyle)
                                    Garbage_Collect(objView)
                                End If
                            Catch ex As Exception

                            End Try


                            Try
                                If blnTurnOffCss = True Then
                                    ObjSEApp.ActiveDocument.CoordinateSystems.Visible = False
                                End If
                            Catch ex As Exception

                            End Try


                            If blnCreatePreview = True Then
                                Try
                                    objDoc.createpreview()
                                Catch ex As Exception

                                End Try
                            End If


                            If blnCheckAssemblyForCorruptLinks = True Then
                                blnCorruptLinkFileProcessed = False
                                Try
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")
                                    End If
                                    ObjSEApp.StartCommand(64000)
                                    ObjSEApp.DoIdle()   'during my testing seemed to be a timing thing... sometime did not report an issue  this API or the sleep below seemed to make it work... need to check into this

                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        blnCorruptLinkFileProcessed = True
                                        Dim strInvalidLinks As String = ""
                                        strInvalidLinks = ReadInvalidLinksFile(strTempFolder + "InvalidLinks.txt")
                                        txtStreamReportStatus.WriteLine("CORRUPT LINKS FOUND IN FILE : " + StrFileNames.Item(ii) + " invalid links are->" + strInvalidLinks)
                                    End If
                                    System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")  'remove it
                                Catch ex As Exception
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    txtStreamReportStatus.WriteLine("Error could not delete the file : " + strTempFolder + "InvalidLinks.txt")
                                    GoTo skip
                                End Try
                            End If
                        End If

                        If objDoc.type = SolidEdgeConstants.DocumentTypeConstants.igAssemblyDocument Then
                            Try
                                obj3dWindow = ObjSEApp.ActiveWindow
                                obj3dWindow.WindowState = 2
                            Catch ex As Exception

                            End Try

                            Try
                                If blnCopyStyles = True Then
                                    UpdateAssyDocumentStyle(objDoc)
                                End If
                            Catch ex As Exception

                            End Try


                            Dim junk As Integer = objDoc.variables.count
                            Try
                                If blnTurnOffGradientBackground = True Then

                                    Dim objSEViewStyle As SolidEdgeFramework.ViewStyle = Nothing
                                    Dim objView As SolidEdgeFramework.View = Nothing

                                    objView = ObjSEApp.ActiveWindow.view
                                    objSEViewStyle = objView.ViewStyle
                                    objSEViewStyle.BackgroundType = SolidEdgeFramework.SeBackgroundType.seBackgroundTypeSolid

                                    Garbage_Collect(objSEViewStyle)
                                    Garbage_Collect(objView)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If blnFitAndShade = True Then
                                    ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewISOView)
                                    ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyAssemblyToolsHideAllReferencePlanes)
                                    obj3dWindow.View.RenderModeType = SolidEdgeFramework.SeRenderModeType.seRenderModeSmoothVHL
                                    ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)
                                End If
                            Catch ex As Exception

                            End Try


                            Try
                                If blnTurnOffCss = True Then
                                    ' there is no API corresponding to the Hide/Show components....  UI  so have to walk the hiearchy the hard way!
                                    'call the recursive sub routine to do so!
                                    ObjSEApp.ActiveDocument.CoordinateSystems.Visible = False
                                    WalkAssemblyTree(ObjSEApp.ActiveDocument)
                                End If
                            Catch ex As Exception

                            End Try

                            If blnCreatePreview = True Then
                                Try
                                    objDoc.createpreview()
                                Catch ex As Exception

                                End Try
                            End If

                            If blnCheckAssemblyForCorruptLinks = True Then
                                blnCorruptLinkFileProcessed = False
                                Try
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")
                                    End If
                                    ObjSEApp.StartCommand(64000)
                                    ObjSEApp.DoIdle()   'during my testing seemed to be a timing thing... sometime did not report an issue  this API or the sleep below seemed to make it work... need to check into this

                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        blnCorruptLinkFileProcessed = True
                                        Dim strInvalidLinks As String = ""
                                        strInvalidLinks = ReadInvalidLinksFile(strTempFolder + "InvalidLinks.txt")
                                        txtStreamReportStatus.WriteLine("CORRUPT LINKS FOUND IN FILE : " + StrFileNames.Item(ii) + " invalid links are->" + strInvalidLinks)
                                    End If
                                    System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")  'remove it
                                Catch ex As Exception
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    txtStreamReportStatus.WriteLine("Error could not delete the file : " + strTempFolder + "InvalidLinks.txt")
                                    GoTo skip
                                End Try
                            End If


                            If blnTurnOffDisplayNextOccProp = True Then
                                TurnOffNextHighestOccProp(objDoc)
                            End If


                        End If

                        If objDoc.type = SolidEdgeConstants.DocumentTypeConstants.igWeldmentDocument Then

                            Try
                                obj3dWindow = ObjSEApp.ActiveWindow
                                obj3dWindow.WindowState = 2
                            Catch ex As Exception

                            End Try

                            Try
                                If blnFitAndShade = True Then
                                    ObjSEApp.StartCommand(SolidEdgeConstants.WeldmentCommandConstants.WeldmentViewISOView)
                                    ObjSEApp.StartCommand(SolidEdgeConstants.WeldmentCommandConstants.WeldmentReferencePlaneHideAllReferencePlanes)
                                    obj3dWindow.View.RenderModeType = SolidEdgeFramework.SeRenderModeType.seRenderModeSmoothVHL
                                    ObjSEApp.StartCommand(SolidEdgeConstants.WeldmentCommandConstants.WeldmentViewFit)
                                End If
                            Catch ex As Exception

                            End Try


                            Try
                                If blnTurnOffGradientBackground = True Then

                                    Dim objSEViewStyle As SolidEdgeFramework.ViewStyle = Nothing
                                    Dim objView As SolidEdgeFramework.View = Nothing

                                    objView = ObjSEApp.ActiveWindow.view
                                    objSEViewStyle = objView.ViewStyle
                                    objSEViewStyle.BackgroundType = SolidEdgeFramework.SeBackgroundType.seBackgroundTypeSolid

                                    Garbage_Collect(objSEViewStyle)
                                    Garbage_Collect(objView)
                                End If
                            Catch ex As Exception

                            End Try


                            If blnCreatePreview = True Then
                                Try
                                    objDoc.createpreview()
                                Catch ex As Exception

                                End Try
                            End If


                            If blnCheckAssemblyForCorruptLinks = True Then
                                blnCorruptLinkFileProcessed = False
                                Try
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")
                                    End If
                                    ObjSEApp.StartCommand(64000)
                                    ObjSEApp.DoIdle()   'during my testing seemed to be a timing thing... sometime did not report an issue  this API or the sleep below seemed to make it work... need to check into this

                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        blnCorruptLinkFileProcessed = True
                                        Dim strInvalidLinks As String = ""
                                        strInvalidLinks = ReadInvalidLinksFile(strTempFolder + "InvalidLinks.txt")
                                        txtStreamReportStatus.WriteLine("CORRUPT LINKS FOUND IN FILE : " + StrFileNames.Item(ii) + " invalid links are->" + strInvalidLinks)
                                    End If
                                    System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")  'remove it
                                Catch ex As Exception
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    txtStreamReportStatus.WriteLine("Error could not delete the file : " + strTempFolder + "InvalidLinks.txt")
                                    GoTo skip
                                End Try
                            End If

                        End If

                        If objDoc.type = SolidEdgeConstants.DocumentTypeConstants.igWeldmentAssemblyDocument Then

                            Try
                                obj3dWindow = ObjSEApp.ActiveWindow
                                obj3dWindow.WindowState = 2
                            Catch ex As Exception

                            End Try


                            If blnCopyStyles = True Then
                                UpdateAssyDocumentStyle(objDoc)
                            End If


                            Try
                                If blnTurnOffCss = True Then
                                    ' there is no API corresponding to the Hide/Show components....  UI  so have to walk the hiearchy the hard way!
                                    'call the recursive sub routine to do so!
                                    ObjSEApp.ActiveDocument.CoordinateSystems.Visible = False
                                    WalkAssemblyTree(ObjSEApp.ActiveDocument)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If blnTurnOffGradientBackground = True Then
                                    Dim objSEViewStyle As SolidEdgeFramework.ViewStyle = Nothing
                                    Dim objView As SolidEdgeFramework.View = Nothing

                                    objView = ObjSEApp.ActiveWindow.view
                                    objSEViewStyle = objView.ViewStyle
                                    objSEViewStyle.BackgroundType = SolidEdgeFramework.SeBackgroundType.seBackgroundTypeSolid

                                    Garbage_Collect(objSEViewStyle)
                                    Garbage_Collect(objView)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If blnFitAndShade = True Then
                                    ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewISOView)
                                    ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyAssemblyToolsHideAllReferencePlanes)
                                    obj3dWindow.View.RenderModeType = SolidEdgeFramework.SeRenderModeType.seRenderModeSmoothVHL
                                    ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)

                                End If
                            Catch ex As Exception

                            End Try


                            If blnCreatePreview = True Then
                                Try
                                    objDoc.createpreview()
                                Catch ex As Exception

                                End Try
                            End If

                            If blnCheckAssemblyForCorruptLinks = True Then
                                blnCorruptLinkFileProcessed = False
                                Try
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")
                                    End If
                                    ObjSEApp.StartCommand(64000)
                                    ObjSEApp.DoIdle()   'during my testing seemed to be a timing thing... sometime did not report an issue  this API or the sleep below seemed to make it work... need to check into this

                                    If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                        blnCorruptLinkFileProcessed = True
                                        Dim strInvalidLinks As String = ""
                                        strInvalidLinks = ReadInvalidLinksFile(strTempFolder + "InvalidLinks.txt")
                                        txtStreamReportStatus.WriteLine("CORRUPT LINKS FOUND IN FILE : " + StrFileNames.Item(ii) + " invalid links are->" + strInvalidLinks)
                                    End If
                                    System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")  'remove it
                                Catch ex As Exception
                                    Dim strTempFolder As String = System.IO.Path.GetTempPath
                                    txtStreamReportStatus.WriteLine("Error could not delete the file : " + strTempFolder + "InvalidLinks.txt")
                                    GoTo skip
                                End Try
                            End If

                        End If

                        'Try
                        objDoc.Save()
SkipToReadOnly:
                        objDoc.Close(False)
                        'Catch ex As Exception
                        'txtStreamReportStatus.WriteLine("Error saving : File: " + StrFileNames.Item(ii) + "ERR->" + ex.Message)
                        'End Try



                        If blnCheckAssemblyForCorruptLinks = True And blnCorruptLinkFileProcessed = False Then  'due to poissible timing issues check one more time
                            Dim strTempFolder As String = System.IO.Path.GetTempPath
                            If System.IO.File.Exists(strTempFolder + "InvalidLinks.txt") = True Then
                                blnCorruptLinkFileProcessed = True
                                Dim strInvalidLinks As String = ""
                                strInvalidLinks = ReadInvalidLinksFile(strTempFolder + "InvalidLinks.txt")
                                txtStreamReportStatus.WriteLine("CORRUPT LINKS FOUND IN FILE : " + StrFileNames.Item(ii) + " invalid links are->" + strInvalidLinks)
                            End If
                            System.IO.File.Delete(strTempFolder + "InvalidLinks.txt")  'remove it
                        End If


                    End If

                Catch ex As Exception
                    txtStreamReportStatus.WriteLine("Error saving File:" + StrFileNames.Item(ii) + " " + ex.Message)
                    objDoc.Close(False)
                    intCtr = intMaxNumberThenRestart - 1
                    GoTo skip
                End Try

skip:
                intCtr = intCtr + 1
                strtempFilename = ""

                Garbage_Collect(obj2dWindow)
                Garbage_Collect(obj3dWindow)
                Garbage_Collect(objModel)
                Garbage_Collect(objModels)
                Garbage_Collect(objFlatPatternModel)
                Garbage_Collect(objFlatPatternModels)
                Garbage_Collect(objDoc)

            Next


        Catch ex As Exception
            txtStreamReportStatus.WriteLine("Catastrophic failure exiting for this set of files:" + StrFileNames.Item(ii) + " " + ex.Message)
        End Try



        If blnCheckIfAlreadyImported = True Then
            txtStreamReportStatus.WriteLine("List of files NOT already imported:")
            For ii = 0 To arrayOfNotImportedFiles.Count - 1
                txtStreamReportStatus.WriteLine(arrayOfNotImportedFiles(ii))
            Next ii

            txtStreamReportStatus.WriteLine("List of files already imported:")
            For ii = 0 To arrayOfImportedFiles.Count - 1
                txtStreamReportStatus.WriteLine(arrayOfImportedFiles(ii))
            Next ii
        End If


        txtStreamReportStatus.WriteLine("Finished processing files " + Date.Now.ToString)


        txtStreamReportStatus.Close()
        Garbage_Collect(txtStreamReportStatus)
        Garbage_Collect(FSOLog)


        Try
            ObjSEApp.Quit()
            Garbage_Collect(ObjSEApp)
        Catch ex As Exception

        End Try

        If blnRunFromCommandLine = False Then
            OpenSaveForm.Controls.Item("Button4").Enabled = True
        End If

        '-- Reset back to 0 to get normal behavior
        SetKeyValue(HKEY_CURRENT_USER, "Software\Unigraphics Solutions\Solid Edge\Version " & strSEVersion & "\Debug", "DocMgmt_OverrideStatusCheckForFileAccess", "0", REG_DWORD)

        If blnCheckAssemblyForCorruptLinks = True Then
            SetKeyValue(HKEY_CURRENT_USER, "Software\Unigraphics Solutions\Solid Edge\Version " & strSEVersion & "\Debug", "WRITE_BADLINKS_TO_FILE", "0", REG_DWORD)
        End If










        OpenSaveForm.Controls.Item("Label2").Text = "Finished processing."
        OpenSaveForm.Controls.Item("Label2").Refresh()

        '******* Added because of .NET
        Try
            OleMessageFilter.Revoke()
        Catch ex As Exception
            PrintLine("Error revoking the message filter.")
        End Try
        '******* Added because of .NET

        If blnRunFromCommandLine = True Then
            End
        End If


    End Sub

    Public Function ParseExtension(ByVal fname As String) As String

        Dim nstart As Long


        nstart = InStrRev(fname, ".", -1, Microsoft.VisualBasic.CompareMethod.Text)

        If nstart <> 0 Then
            ParseExtension = Mid(fname, nstart + 1, Len(fname))
            Exit Function
        End If

        If nstart = 0 Then
            ParseExtension = ""
            Exit Function
        End If

        ParseExtension = ""


    End Function

    Public Function IsValdidFileToProcess(ByVal oFileExtension As String) As Boolean

        IsValdidFileToProcess = False

        If InStr(UCase(oFileExtension), "ASM", CompareMethod.Text) <> 0 Then
            IsValdidFileToProcess = True
            Exit Function
        End If

        If InStr(UCase(oFileExtension), "PAR", CompareMethod.Text) <> 0 Then
            IsValdidFileToProcess = True
            Exit Function
        End If


        If InStr(UCase(oFileExtension), "PSM", CompareMethod.Text) <> 0 Then
            IsValdidFileToProcess = True
            Exit Function
        End If

        If InStr(UCase(oFileExtension), "PWD", CompareMethod.Text) <> 0 Then
            IsValdidFileToProcess = True
            Exit Function
        End If

        If InStr(UCase(oFileExtension), "DFT", CompareMethod.Text) <> 0 Then
            IsValdidFileToProcess = True
            Exit Function
        End If

    End Function

    Public Function ConnectToSolidEdge(ByVal oCount As Integer) As Boolean
        Dim blnRestart As Boolean

        blnRestart = False

        If oCount >= intMaxNumberThenRestart Then
            If blnRunFromCommandLine = False Then
                OpenSaveForm.Controls.Item("Label2").Text = "Closing and restarting Solid Edge"
                OpenSaveForm.Controls.Item("Label2").Refresh()
            End If
            blnRestart = True
            Try
                intCtr = 0
                IntProcessHardwareCTR = 0
                Do While ObjSEApp.Documents.Count <> 0
                    ObjSEApp.DisplayAlerts = False
                    ObjSEApp.ActiveDocument.close()
                Loop
                ObjSEApp.Quit()
                Garbage_Collect(ObjSEApp)

            Catch ex As Exception
                KillProcess("Edge.exe")
                Garbage_Collect(ObjSEApp)
            End Try

        End If



        'to connect to a running instance of Solid Edge
        Try
StartOver:
            ObjSEApp = Marshal.GetActiveObject("SolidEdge.Application")
            ObjSEApp.DisplayAlerts = False
            ObjSEApp.Visible = True
            ConnectToSolidEdge = True
            ObjSEApp.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalTeamCenterMode, False)


            If blnRestart = True Then
                If blnRunFromCommandLine = False Then
                    OpenSaveForm.Controls.Item("Label2").Text = ""
                    OpenSaveForm.Controls.Item("Label2").Refresh()
                End If

            End If
            ConnectToSolidEdge = True
            Exit Function

        Catch ex As System.Exception

            'SE not running then start it
            Try
                ObjSEApp = Activator.CreateInstance(ObjSEAppType)
                ObjSEApp.DisplayAlerts = False
                ObjSEApp.Visible = True
                ConnectToSolidEdge = True
                ObjSEApp.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalTeamCenterMode, False)
            Catch ex2 As Exception
                MsgBox("Could not Start Solid Edge", MsgBoxStyle.OkOnly)
            End Try

            If blnRestart = True Then
                If blnRunFromCommandLine = False Then
                    OpenSaveForm.Controls.Item("Label2").Text = ""
                    OpenSaveForm.Controls.Item("Label2").Refresh()
                End If

            End If
            ConnectToSolidEdge = True
            Exit Function
        End Try

        If ObjSEApp Is Nothing Then
            MsgBox("Could not Start Solid Edge", MsgBoxStyle.OkOnly)
        End If

        ConnectToSolidEdge = False



    End Function









    Public Function KillProcess(ByVal Name As String) As Long

        Dim LocalProcs As Process()
        Dim Proc As Process
        Dim i As Integer
        Dim blnProcessTerminated As Boolean
        blnProcessTerminated = False

        LocalProcs = System.Diagnostics.Process.GetProcesses
        For Each Proc In LocalProcs
            If UCase(Proc.ProcessName) = UCase(Name) Then
                Try
                    Proc.Kill()
                    KillProcess = 0
                    blnProcessTerminated = True
                Catch ex As System.Exception
                    KillProcess = 1
                    LocalProcs = Nothing
                    Exit Function
                End Try
            End If
            i += 1
        Next

        If blnProcessTerminated = True Then
            KillProcess = 0
        End If

        If blnProcessTerminated = False Then
            KillProcess = -2
        End If

    End Function



    Public Sub ReadTextFile()
        Dim FSOTextFile As Scripting.FileSystemObject = Nothing
        Dim oStream As Scripting.TextStream = Nothing
        'Dim strtmpTCNumber As String
        'Dim strtmpTCREvision As String
        Dim strTmp As String
        Dim pstart As Integer

        Try
            FSOTextFile = New Scripting.FileSystemObject
            oStream = FSOTextFile.OpenTextFile(strInputFilename, Scripting.IOMode.ForReading, False, Scripting.Tristate.TristateUseDefault)
            Do While Not oStream.AtEndOfStream
                strTmp = oStream.ReadLine
                StrFileNames.Add(strTmp)
                pstart = 0

skip:
                strTmp = ""
            Loop

            oStream.Close()

        Catch ex As Exception

        End Try


        Garbage_Collect(oStream)
        Garbage_Collect(FSOTextFile)


    End Sub

    Public Function createLogFile() As Boolean

        Try
            If FSOLog Is Nothing Then
                FSOLog = New Scripting.FileSystemObject
            End If

            If FSOLog.FileExists(strLogFilename) = True Then
                FSOLog.DeleteFile(strLogFilename)
            End If

            txtStreamReportStatus = FSOLog.OpenTextFile(strLogFilename, Scripting.IOMode.ForAppending, True, Scripting.Tristate.TristateUseDefault)
            txtStreamReportStatus.WriteLine("Begin processing files  " + Date.Now.ToString)
            createLogFile = True

            Exit Function
        Catch ex As Exception
            MsgBox("Error creating text file to containing import status " + strLogFilename, MsgBoxStyle.Critical)
            End
        End Try

        createLogFile = False

    End Function

    Public Function GetFilePath(ByVal strFname As String) As String
        Dim intStringLength As Integer
        Dim rstart As Integer
        Dim temp As String
        Dim strPathWithNoFileName As String

        intStringLength = Len(strFname)
        rstart = InStrRev(strFname, "\")
        If rstart = 0 Then
            GetFilePath = ""
            Exit Function
        End If
        temp = Mid(strFname, intStringLength - 3, 1)
        strPathWithNoFileName = Mid(strFname, 1, rstart - 1)
        GetFilePath = Mid(strFname, 1, rstart - 1)


    End Function



    Public Function ParseCommandlineInput(ByVal strCmdLine As String) As String


        Dim nstart As Integer

        Dim kk As Integer = 0

        If InStr(strCmdLine, "-i=") <> 0 Then  'means we have the input filename string
            nstart = InStr(strCmdLine, "-i=")
            ParseCommandlineInput = Mid(strCmdLine, nstart + 3, Len(strCmdLine))
            strCommandLineInputFilename = Mid(strCmdLine, nstart + 3, Len(strCmdLine))
            nstart = 0
            Exit Function

        ElseIf InStr(strCmdLine, "-pdf=") <> 0 Then  ' should be -pdf=TRUE or -pdf=FALSE whether or not to create PDF from draft files
            nstart = InStr(strCmdLine, "-pdf=")
            ParseCommandlineInput = Mid(strCmdLine, nstart + 5, Len(strCmdLine))
            strCommandLineCreatePDFs = Mid(strCmdLine, nstart + 5, Len(strCmdLine))
            nstart = 0
            Exit Function

        ElseIf InStr(strCmdLine, "-checkAssembly=") <> 0 Then  'should be -run=TRUE or -run=FALSE whether or not to run the check for invalid links on only assemblies
            nstart = InStr(strCmdLine, "-checkAssembly=")
            ParseCommandlineInput = Mid(strCmdLine, nstart + 15, Len(strCmdLine))
            strCommandLineInputCheckAssembliesForCorruptLinks = Mid(strCmdLine, nstart + 15, Len(strCmdLine))
            nstart = 0
            Exit Function

        ElseIf InStr(strCmdLine, "-fit=") <> 0 Then  'should be -fit=TRUE or -fit=FALSE whether or not to fit and shade models before saving
            nstart = InStr(strCmdLine, "-fit=")
            ParseCommandlineInput = Mid(strCmdLine, nstart + 5, Len(strCmdLine))
            strCommandLineInputFitAndShade = Mid(strCmdLine, nstart + 5, Len(strCmdLine))
            nstart = 0
            Exit Function


        ElseIf InStr(strCmdLine, "-recompute=") <> 0 Then  'should be -fit=TRUE or -fit=FALSE whether or not to fit and shade models before saving
            nstart = InStr(strCmdLine, "-recompute=")
            ParseCommandlineInput = Mid(strCmdLine, nstart + 11, Len(strCmdLine))
            strCommandLineInputRecompute = Mid(strCmdLine, nstart + 11, Len(strCmdLine))
            nstart = 0
            Exit Function

        ElseIf InStr(strCmdLine, "-updateDraft=") <> 0 Then  'should be -fit=TRUE or -fit=FALSE whether or not to update drawing views before saving
            nstart = InStr(strCmdLine, "-updateDraft=")
            ParseCommandlineInput = Mid(strCmdLine, nstart + 13, Len(strCmdLine))
            strCommandLineInputUpdateDrawingViews = Mid(strCmdLine, nstart + 11, Len(strCmdLine))
            nstart = 0
            Exit Function

        ElseIf InStr(strCmdLine, "-resetStatus=") <> 0 Then  'should be -fit=TRUE or -fit=FALSE whether or not to update drawing views before saving
            nstart = InStr(strCmdLine, "-resetStatus=")
            ParseCommandlineInput = Mid(strCmdLine, nstart + 13, Len(strCmdLine))
            strCommandLineInputResetStatus = Mid(strCmdLine, nstart + 11, Len(strCmdLine))
            nstart = 0
            Exit Function

        Else


            ParseCommandlineInput = ""
        End If




    End Function





    Public Sub ReadFilenamesFromCSVFile(ByVal strCSVFilename As String)

        Dim oCSV As Scripting.TextStream = Nothing
        Dim strtmpLine As String = ""
        Dim strTmpFilename As String = ""
        Dim items() As String
        Dim strOrderNumber As String
        Dim strFileNameWithPath As String
        Dim strFileNameOnly As String
        Dim strPathOnly As String
        Dim strTemp As String
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim strTemp1 As String = ""
        Dim strTemp2 As String = ""
        Dim strTemp3 As String = ""
        Dim strTemp4 As String = ""
        Dim strTemp5 As String = ""
        Dim strtmpTextFilename As String = ""

        Dim blnParseMethod1 As Boolean = False



        Dim nstart As Integer
        Dim pstart As Integer

        If (FSO Is Nothing) Then
            FSO = New Scripting.FileSystemObject
        End If


        strtmpTextFilename = FSO.GetBaseName(strCSVFilename)

        If FSO.FileExists(GetFilePath(strCSVFilename) + "\" + strtmpTextFilename + ".txt") = True Then
            FSO.DeleteFile(GetFilePath(strCSVFilename) + "\" + strtmpTextFilename + ".txt", True)
        End If


        Dim oTxt As Scripting.TextStream = Nothing

        Try
            oCSV = FSO.OpenTextFile(strCSVFilename, Scripting.IOMode.ForReading, False)
            oTxt = FSO.OpenTextFile(GetFilePath(strCSVFilename) + "\" + strtmpTextFilename + ".txt", Scripting.IOMode.ForAppending, True, Scripting.Tristate.TristateUseDefault)
        Catch ex As Exception
            'Beep()
        End Try

        Try

            Do While Not oCSV.AtEndOfStream
                strtmpLine = oCSV.ReadLine
                ' commented out v 29  strtmpLine = Mid(strtmpLine, 2, Len(strtmpLine) - 2)
                items = Split(strtmpLine, Chr(34) + "," + Chr(34))

                If oCSV.Line = 1 Then
                    GoTo skip
                End If
                If oCSV.Line = 2 Then
                    GoTo skip
                End If

                If oCSV.Line = 3 Then
                    GoTo skip
                End If

                If InStr(strtmpLine, "Unordered") <> 0 Then
                    GoTo skip
                End If

                If InStr(strtmpLine, "Ordered") <> 0 Then
                    GoTo skip
                End If

                If InStr(strtmpLine, "Reverse Link") <> 0 Then
                    GoTo skip
                End If

                If InStr(strtmpLine, "Order Number") <> 0 Then
                    GoTo skip
                End If

                If oCSV.Line >= 4 Then
                    '*******   in V29 to fix problem reading the csv file for all locales
                    nstart = InStr(strtmpLine, Chr(34), CompareMethod.Text)

                    If nstart <> 0 Then
                        strTemp1 = Mid(strtmpLine, nstart + 1, Len(strtmpLine))
                        pstart = InStr(strTemp1, Chr(34), CompareMethod.Text)
                        str1 = Mid(strTemp1, 1, pstart - 1)
                        strTemp2 = Mid(strTemp1, pstart + 1, Len(strTemp1))
                        nstart = InStr(strTemp2, Chr(34), CompareMethod.Text)
                        If nstart <> 0 Then
                            strTemp3 = Mid(strTemp2, nstart + 1, Len(strTemp2))
                            pstart = InStr(strTemp3, Chr(34), CompareMethod.Text)
                            str2 = Mid(strTemp3, 1, pstart - 1)
                            strTemp4 = Mid(strTemp3, pstart + 1, Len(strTemp3))
                            nstart = InStr(strTemp4, Chr(34), CompareMethod.Text)
                            If nstart <> 0 Then
                                strTemp5 = Mid(strTemp4, nstart + 1, Len(strTemp2))
                                pstart = InStr(strTemp5, Chr(34), CompareMethod.Text)
                                str3 = Mid(strTemp5, 1, pstart - 1)
                            End If
                        End If
                    End If

                    If str3 <> "" Then   ' means it is a reverse link file so do not add a second time to files to be processed
                        GoTo skip
                    End If

                    strOrderNumber = str1
                    strFileNameWithPath = str2

                    If StrFileNames.Count = 0 Then
                        StrFileNames.Add(strFileNameWithPath)
                        oTxt.WriteLine(strFileNameWithPath)
                        DummySpreadsheetList.Add(UCase(strFileNameWithPath))

                    Else
                        If DummySpreadsheetList.Contains(UCase(strFileNameWithPath)) = False Then
                            StrFileNames.Add(strFileNameWithPath)
                            oTxt.WriteLine(strFileNameWithPath)
                            DummySpreadsheetList.Add(UCase(strFileNameWithPath))

                        End If
                    End If





                End If
skip:
                strtmpLine = ""
                strOrderNumber = ""
                strFileNameWithPath = ""
                strFileNameOnly = ""
                strPathOnly = ""
                strTemp = ""
                nstart = 0
                pstart = 0
                strTmpFilename = ""
                str1 = ""
                str2 = ""
                str3 = ""
                strTemp1 = ""
                strTemp2 = ""
                strTemp3 = ""
                strTemp4 = ""
                strTemp5 = ""


            Loop
            oCSV.Close()
            oTxt.Close()
        Catch ex As Exception
            'Beep()
        End Try
        Garbage_Collect(oTxt)
        Garbage_Collect(oCSV)

    End Sub


    Public Sub ShellandWait(ByVal ProcessPath As String, ByVal arguments As String)

        Try
            Dim myProcess As Process = System.Diagnostics.Process.Start(ProcessPath, arguments)
            myProcess.WaitForExit()
            myProcess.Close()
            myProcess.Dispose()

        Catch ex As Exception
            'possibly need to raise an error
        End Try


    End Sub

    Public Function ResetSEFileStatusToAvailable(ByVal oFilename As String) As Boolean
        Dim Objproperties As SolidEdgeFileProperties.Properties
        Dim objproperty As SolidEdgeFileProperties.Property
        Dim PropertySets As SolidEdgeFileProperties.PropertySets

        Try
            PropertySets = New SolidEdgeFileProperties.PropertySets
            PropertySets.Open(oFilename)
            Objproperties = PropertySets.Item("ExtendedSummaryInformation")
            objproperty = Objproperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igExtSumInfoStatus)
            If objproperty.Value <> SolidEdgeConstants.DocumentStatus.igStatusAvailable Then
                If objproperty.Value = SolidEdgeConstants.DocumentStatus.igStatusReleased Then
                    PropertySets.Item("Custom").add("TC_SEStatus", "Released")
                ElseIf objproperty.Value = SolidEdgeConstants.DocumentStatus.igStatusObsolete Then
                    PropertySets.Item("Custom").add("TC_SEStatus", "Obselete")
                ElseIf objproperty.Value = SolidEdgeConstants.DocumentStatus.igStatusInWork Then
                    PropertySets.Item("Custom").add("TC_SEStatus", "InWork")
                ElseIf objproperty.Value = SolidEdgeConstants.DocumentStatus.igStatusInReview Then
                    PropertySets.Item("Custom").add("TC_SEStatus", "InReview")
                ElseIf objproperty.Value = SolidEdgeConstants.DocumentStatus.igStatusBaselined Then
                    PropertySets.Item("Custom").add("TC_SEStatus", "Baselined")
                End If
                objproperty.Value = SolidEdgeConstants.DocumentStatus.igStatusAvailable
                Objproperties = PropertySets.Item("SummaryInformation")
                objproperty = Objproperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igSumInfoDocSecurity)
                objproperty.Value = SolidEdgeConstants.DocumentStatus.igStatusAvailable
                PropertySets.Save()
                PropertySets.Close()
                Garbage_Collect(objproperty)
                Garbage_Collect(Objproperties)
                Garbage_Collect(PropertySets)
                Exit Try
            End If
            PropertySets.Close()
            Garbage_Collect(objproperty)
            Garbage_Collect(Objproperties)
            Garbage_Collect(PropertySets)
            ResetSEFileStatusToAvailable = True
            Exit Function
        Catch ex As Exception
            txtStreamReportStatus.WriteLine("ERROR->Status can not be set to available:" + oFilename + " " + ex.Message)
        End Try

        ResetSEFileStatusToAvailable = False
    End Function

    Public Function ReadInvalidLinksFile(ByVal oFile As String) As String

        Dim FSO As Scripting.FileSystemObject = Nothing
        Dim oText As Scripting.TextStream = Nothing
        Dim strReadLine As String = ""
        Dim strBadLinks As String = ""
        Try
            FSO = New Scripting.FileSystemObject
            oText = FSO.OpenTextFile(oFile, Scripting.IOMode.ForReading, False, Scripting.Tristate.TristateUseDefault)
            Do While Not oText.AtEndOfStream
                strReadLine = oText.ReadLine
                If InStr(strReadLine, Chr(9), CompareMethod.Text) <> 0 Then
                    strBadLinks = strBadLinks + " " + Mid(strReadLine, InStr(strReadLine, Chr(9), CompareMethod.Text) + 1, Len(strReadLine))
                End If
            Loop
            oText.Close()
            ReadInvalidLinksFile = strBadLinks
            Garbage_Collect(oText)
            Garbage_Collect(FSO)
            Exit Function
        Catch ex As Exception
            Beep()

        End Try


        ReadInvalidLinksFile = ""
    End Function

    Public Sub UpdatePartDocumentStyle(ByRef pDoc As SolidEdgePart.PartDocument)
        Dim strTmp As String = pDoc.Name
        Dim bV As Boolean = True
        Dim ii As Integer

        Try
            Dim pStyles As SolidEdgeFramework.ViewStyles
            Dim NewDefault As SolidEdgeFramework.ViewStyle


            ' Dim strSrc As String = "C:\Users\aspatric\Documents\PLANNING\SE_TEMPLATES\V104\Style_Source\Part and Assembly Styles.asm"


            Dim strSrc As String = strSEInstalledPath + "\Template\More\iso part.par"

            If IO.File.Exists(strSrc) = True Then
                pDoc.ImportStyles(strSrc, bV)
                pStyles = CType(pDoc.ViewStyles, SolidEdgeFramework.ViewStyles)

                For ii = 1 To pStyles.Count
                    If pStyles.Item(ii).StyleName = "Default" Then
                        pStyles.Item(ii).StyleName = "Default_Pre-ST4"
                    End If
                Next

                NewDefault = pStyles.AddFromFile(strSrc, "Default")
                pStyles.AddFromFile(strSrc, "High Quality")
                pStyles.AddFromFile(strSrc, "Perspective")

                Dim pView As SolidEdgeFramework.View
                Dim pWindow As SolidEdgeFramework.Window

                pWindow = CType(pDoc.Windows.Item(1), SolidEdgeFramework.Window)
                pView = CType(pWindow.View, SolidEdgeFramework.View)
                pView.ViewStyle = NewDefault
                ObjSEApp.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewFit)

                Garbage_Collect(pView)
                Garbage_Collect(pWindow)
                Garbage_Collect(NewDefault)
                Garbage_Collect(pStyles)

            End If





        Catch ex As Exception
            MessageBox.Show(ex.Message, "UpdatePartDocument")
        End Try

    End Sub

    Public Sub UpdateSheetMetalDocumentStyle(ByRef pDoc As SolidEdgePart.SheetMetalDocument)
        Dim strTmp As String = pDoc.Name
        Dim bV As Boolean = True

        Try
            Dim pStyles As SolidEdgeFramework.ViewStyles
            Dim pOldDefault As SolidEdgeFramework.ViewStyle
            Dim NewDefault As SolidEdgeFramework.ViewStyle

            ' Dim strSrc As String = "C:\Users\aspatric\Documents\PLANNING\SE_TEMPLATES\V104\Style_Source\Part and Assembly Styles.asm"

            Dim strSrc As String = strSEInstalledPath + "\Template\More\iso sheet metal.psm"


            If IO.File.Exists(strSrc) = True Then
                pDoc.ImportStyles(strSrc, bV)
                pStyles = CType(pDoc.ViewStyles, SolidEdgeFramework.ViewStyles)

                For ii = 1 To pStyles.Count
                    If pStyles.Item(ii).StyleName = "Default" Then
                        pStyles.Item(ii).StyleName = "Default_Pre-ST4"
                    End If
                Next

                NewDefault = pStyles.AddFromFile(strSrc, "Default")
                pStyles.AddFromFile(strSrc, "High Quality")
                pStyles.AddFromFile(strSrc, "Perspective")

                Dim pView As SolidEdgeFramework.View
                Dim pWindow As SolidEdgeFramework.Window

                pWindow = CType(pDoc.Windows.Item(1), SolidEdgeFramework.Window)
                pView = CType(pWindow.View, SolidEdgeFramework.View)
                pView.ViewStyle = NewDefault

                ObjSEApp.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewFit)
                Garbage_Collect(pView)
                Garbage_Collect(pWindow)
                Garbage_Collect(NewDefault)
                Garbage_Collect(pStyles)

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "UpdateSheetMetalDocument")
        End Try

    End Sub


    Public Sub UpdateAssyDocumentStyle(ByRef pDoc As SolidEdgeAssembly.AssemblyDocument)
        Dim strTmp As String = pDoc.Name
        Dim bV As Boolean = True

        Try
            Dim pStyles As SolidEdgeFramework.ViewStyles
            Dim pOldDefault As SolidEdgeFramework.ViewStyle
            Dim NewDefault As SolidEdgeFramework.ViewStyle

            ' Dim strSrc As String = "C:\Users\aspatric\Documents\PLANNING\SE_TEMPLATES\V104\Style_Source\Part and Assembly Styles.asm"



            Dim strSrc As String = strSEInstalledPath + "\Template\More\iso assembly.asm"

            If IO.File.Exists(strSrc) = True Then
                pDoc.ImportStyles(strSrc, bV)
                pStyles = CType(pDoc.ViewStyles, SolidEdgeFramework.ViewStyles)

                For ii = 1 To pStyles.Count
                    If pStyles.Item(ii).StyleName = "Default" Then
                        pStyles.Item(ii).StyleName = "Default_Pre-ST4"
                    End If
                Next

                NewDefault = pStyles.AddFromFile(strSrc, "Default")
                pStyles.AddFromFile(strSrc, "High Quality")
                pStyles.AddFromFile(strSrc, "Perspective")

                Dim pView As SolidEdgeFramework.View
                Dim pWindow As SolidEdgeFramework.Window

                pWindow = CType(pDoc.Windows.Item(1), SolidEdgeFramework.Window)
                pView = CType(pWindow.View, SolidEdgeFramework.View)
                pView.ViewStyle = NewDefault

                ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)

                Garbage_Collect(pView)
                Garbage_Collect(pWindow)
                Garbage_Collect(NewDefault)
                Garbage_Collect(pStyles)

            End If



        Catch ex As Exception
            MessageBox.Show(ex.Message, "UpdateAssyDocument")
        End Try

    End Sub


    Public Sub TurnOffNextHighestOccProp(ByVal oDoc As SolidEdgeAssembly.AssemblyDocument)




        Try

            Dim objSelectSet As SolidEdgeFramework.SelectSet = Nothing
            Dim ii As Integer
            Dim objOcc As SolidEdgeAssembly.Occurrence = Nothing
            Dim oObject As Object = Nothing

            objSelectSet = oDoc.SelectSet
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyAssemblyToolsShowAll)

            oArrayofTotallyHiddenOccs = New ArrayList
            oArrayofVisibleOccs = New ArrayList


            'Top View
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewTopView)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblySelectVisibleParts)
            objSelectSet.Add(ObjSEApp.ActiveSelectSet)

            oArrayOfOccsVisibleInTopView = New ArrayList
            For ii = 1 To objSelectSet.Count
                oObject = TryCast(objSelectSet.Item(ii), SolidEdgeFramework.SelectSet)
                If Not (oObject Is Nothing) Then
                    GoTo skip11
                End If
                If oArrayOfOccsVisibleInTopView.Contains(objSelectSet.Item(ii)) = False Then
                    oArrayOfOccsVisibleInTopView.Add(objSelectSet.Item(ii))
                End If
skip11:
            Next

            objSelectSet.RemoveAll()
            ObjSEApp.DoIdle()
            ' Garbage_Collect(oObject)
            oObject = Nothing


            'Bottom view
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewBottomView)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblySelectVisibleParts)
            objSelectSet.Add(ObjSEApp.ActiveSelectSet)
            oArrayOfOccsVisibleInBottomView = New ArrayList

            For ii = 1 To objSelectSet.Count
                oObject = TryCast(objSelectSet.Item(ii), SolidEdgeFramework.SelectSet)
                If Not (oObject Is Nothing) Then
                    GoTo skip12
                End If
                If oArrayOfOccsVisibleInBottomView.Contains(objSelectSet.Item(ii)) = False Then
                    oArrayOfOccsVisibleInBottomView.Add(objSelectSet.Item(ii))
                End If
skip12:
            Next

            objSelectSet.RemoveAll()
            ObjSEApp.DoIdle()
            ' Garbage_Collect(oObject)
            oObject = Nothing

            'Right view
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewRightView)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblySelectVisibleParts)
            objSelectSet.Add(ObjSEApp.ActiveSelectSet)
            oArrayOfOccsVisibleInRightView = New ArrayList

            For ii = 1 To objSelectSet.Count
                oObject = TryCast(objSelectSet.Item(ii), SolidEdgeFramework.SelectSet)
                If Not (oObject Is Nothing) Then
                    GoTo skip13
                End If
                If oArrayOfOccsVisibleInRightView.Contains(objSelectSet.Item(ii)) = False Then
                    oArrayOfOccsVisibleInRightView.Add(objSelectSet.Item(ii))
                End If
skip13:
            Next

            objSelectSet.RemoveAll()
            ObjSEApp.DoIdle()
            ' Garbage_Collect(oObject)
            oObject = Nothing


            'Left view
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewLeftView)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblySelectVisibleParts)
            objSelectSet.Add(ObjSEApp.ActiveSelectSet)
            oArrayOfOccsVisibleInLeftView = New ArrayList

            For ii = 1 To objSelectSet.Count
                oObject = TryCast(objSelectSet.Item(ii), SolidEdgeFramework.SelectSet)
                If Not (oObject Is Nothing) Then
                    GoTo skip14
                End If
                If oArrayOfOccsVisibleInLeftView.Contains(objSelectSet.Item(ii)) = False Then
                    oArrayOfOccsVisibleInLeftView.Add(objSelectSet.Item(ii))
                End If
skip14:
            Next

            objSelectSet.RemoveAll()
            ObjSEApp.DoIdle()
            ' Garbage_Collect(oObject)
            oObject = Nothing


            'Front view
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFrontView)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblySelectVisibleParts)
            objSelectSet.Add(ObjSEApp.ActiveSelectSet)
            oArrayOfOccsVisibleInFrontView = New ArrayList
            For ii = 1 To objSelectSet.Count
                oObject = TryCast(objSelectSet.Item(ii), SolidEdgeFramework.SelectSet)
                If Not (oObject Is Nothing) Then
                    GoTo skip15
                End If
                If oArrayOfOccsVisibleInFrontView.Contains(objSelectSet.Item(ii)) = False Then
                    oArrayOfOccsVisibleInFrontView.Add(objSelectSet.Item(ii))
                End If
skip15:
            Next

            objSelectSet.RemoveAll()
            ObjSEApp.DoIdle()
            ' Garbage_Collect(oObject)
            oObject = Nothing


            'Back View??????

            ObjSEApp.DoIdle()


            For ii = 0 To oArrayOfOccsVisibleInBottomView.Count - 1
                If oArrayofVisibleOccs.Contains(oArrayOfOccsVisibleInBottomView.Item(ii)) = False Then
                    oArrayofVisibleOccs.Add(oArrayOfOccsVisibleInBottomView.Item(ii))
                End If
            Next

            For ii = 0 To oArrayOfOccsVisibleInTopView.Count - 1
                If oArrayofVisibleOccs.Contains(oArrayOfOccsVisibleInTopView.Item(ii)) = False Then
                    oArrayofVisibleOccs.Add(oArrayOfOccsVisibleInTopView.Item(ii))
                End If
            Next
            For ii = 0 To oArrayOfOccsVisibleInLeftView.Count - 1
                If oArrayofVisibleOccs.Contains(oArrayOfOccsVisibleInLeftView.Item(ii)) = False Then
                    oArrayofVisibleOccs.Add(oArrayOfOccsVisibleInLeftView.Item(ii))
                End If
            Next
            For ii = 0 To oArrayOfOccsVisibleInRightView.Count - 1
                If oArrayofVisibleOccs.Contains(oArrayOfOccsVisibleInRightView.Item(ii)) = False Then
                    oArrayofVisibleOccs.Add(oArrayOfOccsVisibleInRightView.Item(ii))
                End If
            Next
            For ii = 0 To oArrayOfOccsVisibleInFrontView.Count - 1
                If oArrayofVisibleOccs.Contains(oArrayOfOccsVisibleInFrontView.Item(ii)) = False Then
                    oArrayofVisibleOccs.Add(oArrayOfOccsVisibleInFrontView.Item(ii))
                End If
            Next

            Dim oRef As SolidEdgeFramework.Reference = Nothing
            Dim oArrayOfOccs As System.Array = {0}

            Dim oTopOcc As Object = Nothing
            Dim NumSubOccs, NumBoundSubOccs As Integer

            ' ObjSEApp.ScreenUpdating = False

            For ii = 0 To oArrayofVisibleOccs.Count - 1
                oObject = TryCast(oArrayofVisibleOccs.Item(ii), SolidEdgeFramework.SelectSet)
                If Not (oObject Is Nothing) Then
                    GoTo skip1
                End If

                objOcc = TryCast(oArrayofVisibleOccs.Item(ii), SolidEdgeAssembly.Occurrence)
                If Not (objOcc Is Nothing) Then
                    Dim strname1 As String = oArrayofVisibleOccs.Item(ii).name
                    oArrayofVisibleOccs.Item(ii).visible = False
                Else
                    Dim strOccName As String = oArrayofVisibleOccs.Item(ii).object.name
                    oRef = oArrayofVisibleOccs.Item(ii)
                    oArrayOfOccs = {0}
                    oRef.GetOccurrencesInPath(oTopOcc, NumSubOccs, NumBoundSubOccs, oArrayOfOccs)
                    If NumBoundSubOccs <> 0 Then
                        oArrayOfOccs(NumBoundSubOccs - 1).visible = False

                        ' ''Dim oSubOccs As SolidEdgeAssembly.SubOccurrences = oArrayOfOccs(NumBoundSubOccs - 1).suboccurrences
                        ' ''Dim oSubOcc As SolidEdgeAssembly.SubOccurrence = Nothing
                        ' ''For Each oSubOcc In oSubOccs
                        ' ''    If oSubOcc.Name.ToLower = strOccName.ToLower Then
                        ' ''        oSubOcc.ThisAsOccurrence.Visible = False
                        ' ''        Exit For
                        ' ''    End If
                        ' ''Next
                        ' ''Garbage_Collect(oSubOcc)
                        ' ''Garbage_Collect(oSubOccs)
                    End If

                End If
                objOcc = Nothing
                Garbage_Collect(objOcc)
skip1:
                oArrayOfOccs = Nothing
                oTopOcc = Nothing
                NumSubOccs = 0
                NumBoundSubOccs = 0
            Next


            oObject = Nothing

            Garbage_Collect(oTopOcc)
            Garbage_Collect(oRef)

            NumSubOccs = 0
            NumBoundSubOccs = 0


            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblySelectVisibleParts)
            objSelectSet.Add(ObjSEApp.ActiveSelectSet)


            For ii = 1 To objSelectSet.Count
                oObject = TryCast(objSelectSet.Item(ii), SolidEdgeFramework.SelectSet)
                If Not (oObject Is Nothing) Then
                    GoTo skip2
                End If

                objOcc = TryCast(objSelectSet.Item(ii), SolidEdgeAssembly.Occurrence)
                If Not (objOcc Is Nothing) Then
                    objSelectSet.Item(ii).DisplayInSubAssembly = False
                Else

                    oArrayOfOccs = {0}
                    oRef = objSelectSet.Item(ii)
                    oRef.GetOccurrencesInPath(oTopOcc, NumSubOccs, NumBoundSubOccs, oArrayOfOccs)
                    If NumBoundSubOccs <> 0 Then
                        oArrayOfOccs(NumBoundSubOccs - 1).DisplayInSubAssembly = False
                    End If

                End If
                objOcc = Nothing
                ' Garbage_Collect(oObject)
                oObject = Nothing
skip2:
                oArrayOfOccs = Nothing
                NumSubOccs = 0
                NumBoundSubOccs = 0
            Next


            Garbage_Collect(oObject)

            Garbage_Collect(oTopOcc)
            Garbage_Collect(oRef)

            NumSubOccs = 0
            NumBoundSubOccs = 0
            oArrayOfOccs = Nothing
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyAssemblyToolsShowAll)


            Garbage_Collect(objSelectSet)


            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewISOView)
            ObjSEApp.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)
            oDoc.Save()


        Catch ex As Exception

        End Try


    End Sub


    Public Sub WalkAssemblyTree(ByVal oDoc As SolidEdgeAssembly.AssemblyDocument)
        Dim oOccurrences As SolidEdgeAssembly.Occurrences = Nothing
        Dim oOccurrence As SolidEdgeAssembly.Occurrence = Nothing

        Try
            oOccurrences = oDoc.Occurrences
            For Each oOccurrence In oOccurrences
                Dim junk1 As String = oOccurrence.OccurrenceFileName
                oOccurrence.DisplayCoordinateSystems = False

                If oOccurrence.Subassembly = True Then
                    oOccurrence.CoordinateSystemsVisible = False
                    WalkAssemblySubOccurrences(oOccurrence)
                End If
            Next


        Catch ex As Exception

        End Try


    End Sub

    Public Sub WalkAssemblySubOccurrences(ByVal oOcc As Object)

        Dim objSubOccurrences As SolidEdgeAssembly.SubOccurrences = Nothing
        Dim objSubOccurrence As SolidEdgeAssembly.SubOccurrence = Nothing

        Try

            oOcc.DisplayCoordinateSystems = False
            If (oOcc.Subassembly = True) Then
                oOcc.CoordinateSystemsVisible = False
            End If
            objSubOccurrences = oOcc.SubOccurrences

            If Not (objSubOccurrences Is Nothing) Then
                For Each objSubOccurrence In objSubOccurrences
                    Dim junk2 As String = objSubOccurrence.SubOccurrenceFileName
                    objSubOccurrence.DisplayCoordinateSystems = False
                    If (objSubOccurrence.Subassembly = True) Then
                        objSubOccurrence.CoordinateSystemsVisible = False
                        WalkAssemblySubOccurrences(objSubOccurrence)
                    End If
                Next
            End If

        Catch ex As Exception

        End Try


    End Sub
    Public Function FileInUse(ByVal sFile As String) As Boolean
        Dim thisFileInUse As Boolean = False
        If System.IO.File.Exists(sFile) Then
            Try
                Using f As New IO.FileStream(sFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                    ' thisFileInUse = False
                End Using
            Catch ex As System.IO.IOException
                thisFileInUse = True
            End Try
        End If
        Return thisFileInUse
    End Function
    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgePartFile(ByVal Filename As String) As Boolean
        Dim sExtension As String = Nothing
        Dim bPart As Boolean = False

        Try
            sExtension = System.IO.Path.GetExtension(Filename)
            If (sExtension.ToLower = ".par") Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeSheetMetalFile(ByVal Filename As String) As Boolean
        Dim sExtension As String = Nothing
        Dim bPart As Boolean = False

        Try
            sExtension = System.IO.Path.GetExtension(Filename)
            If (sExtension.ToLower = ".psm") Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeDraftFile(ByVal Filename As String) As Boolean
        Dim sExtension As String = Nothing
        Dim bPart As Boolean = False

        Try
            sExtension = System.IO.Path.GetExtension(Filename)
            If (sExtension.ToLower = ".dft") Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeAssemblylFile(ByVal Filename As String) As Boolean
        Dim sExtension As String = Nothing
        Dim bPart As Boolean = False

        Try

            sExtension = System.IO.Path.GetExtension(Filename)
            If (sExtension.ToLower = ".asm") Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeWeldmentFile(ByVal Filename As String) As Boolean
        Dim sExtension As String = Nothing
        Dim bPart As Boolean = False

        Try
            sExtension = System.IO.Path.GetExtension(Filename)
            If (sExtension.ToLower = ".pwd") Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Sub oForceGarbageCollection()
        Try
            GC.Collect(GC.MaxGeneration())
            GC.WaitForPendingFinalizers()
            GC.Collect(GC.MaxGeneration())
        Catch ex As Exception

        End Try
    End Sub

    Private Function ProcessHardwareCheck(oFile) As Boolean
        Dim objProperties As SolidEdgeFileProperties.Properties = Nothing
        Dim objProperty As SolidEdgeFileProperties.Property = Nothing
        Dim objPropertySets As SolidEdgeFileProperties.PropertySets = Nothing

        Dim jj As Long = 0

        Dim blnImportedPreviously As Boolean = False
        Try
            objPropertySets = New SolidEdgeFileProperties.PropertySets
            objPropertySets.Open(oFile, False)
            Try
                objProperties = objPropertySets.Item("DocumentSummaryInformation")
                Dim pulledSEPropertyValue As String = objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igDocSumInfoCategory).value
                If pulledSEPropertyValue.ToLower = strHardwarePropertyValue.ToLower Then
                    objProperties = objPropertySets.Item("ExtendedSummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igExtSumInfoHardwarePart).value = True
                    objPropertySets.Save()
                End If
wrapup1:
                objPropertySets.Close()


            Catch ex As Exception

                'attempt a workaround for OLD files
                IntProcessHardwareCTR = IntProcessHardwareCTR + 1
                objPropertySets.Close()
                Garbage_Collect(objProperty)
                Garbage_Collect(objProperties)
                Garbage_Collect(objPropertySets)


                If ConnectToSolidEdge(IntProcessHardwareCTR) = True Then
                    Dim oProps As SolidEdgeFramework.Properties = Nothing
                    Dim oProp As SolidEdgeFramework.Property = Nothing
                    Dim oPropSets As SolidEdgeFramework.PropertySets = Nothing
                    Dim objdoc As Object = Nothing
                    Dim oPartDoc As SolidEdgePart.PartDocument = Nothing
                    objdoc = ObjSEApp.Documents.Open(oFile)

                    oPartDoc = TryCast(objdoc, SolidEdgePart.PartDocument)

                    If Not oPartDoc Is Nothing Then
                        oPropSets = oPartDoc.Properties
                        oProps = oPropSets.Item("ProjectInformation")
                        Dim tempOriginal As String = oProps.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igProjInfoProjectName).Value


                        If tempOriginal = String.Empty Then
                            tempOriginal = " "
                        Else
                            oProps.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igProjInfoProjectName).Value = "testing"
                            oPropSets.Save()
                        End If

                        Try
                            oProps.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igProjInfoProjectName).Value = tempOriginal
                            oProps = oPropSets.Item("ExtendedSummaryInformation")
                            oProps.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igExtSumInfoHardwarePart).Value = True
                            oPropSets.Save()

                            oPartDoc.Save()
                            oPartDoc.Close()


                            Garbage_Collect(oProp)
                            Garbage_Collect(oProps)
                            Garbage_Collect(oPropSets)
                            Garbage_Collect(oPartDoc)
                            Garbage_Collect(objdoc)

                        Catch ex1 As Exception
                            oProps = oPropSets.Item("ProjectInformation")
                            oProps.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igProjInfoProjectName).Value = tempOriginal
                            oPropSets.Save()
                            'raise error in log file
                            txtStreamReportStatus.WriteLine("ERROR setting hardware checkbox->:" + ex1.Message + "  processing file:" + oFile.ToString)
                        End Try


                    End If


                End If
                
            End Try



wrapup:
            Garbage_Collect(objProperty)
            Garbage_Collect(objProperties)
            Garbage_Collect(objPropertySets)


            If blnImportedPreviously = True Then
                Return True
                Exit Function
            Else
                Return False
                Exit Function
            End If
        Catch ex As Exception

            Garbage_Collect(objProperty)
            Garbage_Collect(objProperties)
            Garbage_Collect(objPropertySets)
        End Try

        Return False



    End Function



    Private Function ProcessFilePropertiesToCheckIfImported(ByVal oFile As String) As Boolean
        Dim objProperties As SolidEdgeFileProperties.Properties = Nothing
        Dim objProperty As SolidEdgeFileProperties.Property = Nothing
        Dim objPropertySets As SolidEdgeFileProperties.PropertySets = Nothing

        Dim jj As Long = 0
       
        Dim blnImportedPreviously As Boolean = False

        
        Try
            
            objPropertySets = New SolidEdgeFileProperties.PropertySets
            objPropertySets.Open(oFile, True)
            Try
                objProperties = objPropertySets.Item("Custom")
                For Each objProperty In objProperties


                    If objProperty.Name = "TC_ImportTime" Then
                        blnImportedPreviously = True
                        GoTo wrapup1
                    End If

                    If objProperty.Name = "InsightXT_ImportTime" Then
                        blnImportedPreviously = True
                        GoTo wrapup1
                    End If

                Next
wrapup1:
                objPropertySets.Close()
            Catch ex As Exception

            End Try
            

          
wrapup:
            Garbage_Collect(objProperty)
            Garbage_Collect(objProperties)
            Garbage_Collect(objPropertySets)


            If blnImportedPreviously = True Then
                Return True
                Exit Function
            Else
                Return False
                Exit Function
            End If
        Catch ex As Exception

            Garbage_Collect(objProperty)
            Garbage_Collect(objProperties)
            Garbage_Collect(objPropertySets)
        End Try

        Return False

    End Function



    Private Sub ProcessFileProperties(ByVal oFile As String)
        Dim objProperties As SolidEdgeFileProperties.Properties = Nothing
        Dim objProperty As SolidEdgeFileProperties.Property = Nothing
        Dim objPropertySets As SolidEdgeFileProperties.PropertySets = Nothing
        CustomPropertiesToDelete = New ArrayList
        Dim jj As Long = 0
        Dim arrData() As String
        Dim strData As String = ""
        arrData = strPropsToRemove.Split(",")
        For Each strData In arrData
            If strData = "" Then
                Continue For
            End If
            If CustomPropertiesToDelete.Contains(strData) = False Then
                CustomPropertiesToDelete.Add(strData)
            End If
        Next

        objPropertySets = New SolidEdgeFileProperties.PropertySets
        objPropertySets.Open(oFile, False)
        Try
            For jj = 0 To CustomPropertiesToDelete.Count - 1
                If CustomPropertiesToDelete(jj).ToString.ToLower = "title" Then
                    objProperties = objPropertySets.Item("SummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igSumInfoTitle).value = " "
                ElseIf CustomPropertiesToDelete(jj).ToString.ToLower = "subject" Then
                    objProperties = objPropertySets.Item("SummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igSumInfoSubject).value = " "
                ElseIf CustomPropertiesToDelete(jj).ToString.ToLower = "author" Then
                    objProperties = objPropertySets.Item("SummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igSumInfoAuthor).value = " "
                ElseIf CustomPropertiesToDelete(jj).ToString.ToLower = "keywords" Then
                    objProperties = objPropertySets.Item("SummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igSumInfoKeyWords).value = " "
                ElseIf CustomPropertiesToDelete(jj).ToString.ToLower = "comments" Then
                    objProperties = objPropertySets.Item("SummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igSumInfoComments).value = " "
                ElseIf CustomPropertiesToDelete(jj).ToString.ToLower = "category" Then
                    objProperties = objPropertySets.Item("DocumentSummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igDocSumInfoCategory).value = " "
                ElseIf CustomPropertiesToDelete(jj).ToString.ToLower = "manager" Then
                    objProperties = objPropertySets.Item("DocumentSummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igDocSumInfoManager).value = " "
                ElseIf CustomPropertiesToDelete(jj).ToString.ToLower = "company" Then
                    objProperties = objPropertySets.Item("DocumentSummaryInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igDocSumInfoCompany).value = " "
                ElseIf CustomPropertiesToDelete(jj).ToString.ToLower = "project name" Then
                    objProperties = objPropertySets.Item("ProjectInformation")
                    objProperties.PropertyByID(SolidEdgeFileProperties.PropertyIDs.igProjInfoProjectName).value = " "
                Else  'must be custom

                    Try
                        objProperties = objPropertySets.Item("Custom")
                        objProperties.Add(CustomPropertiesToDelete(jj), "DeleteMe")
                        objProperty = objProperties.Item(CustomPropertiesToDelete(jj))
                        objProperty.Delete()
                    Catch ex As Exception

                    End Try
                End If
            Next

            objPropertySets.Save()
            objPropertySets.Close()


            Garbage_Collect(objProperty)
            Garbage_Collect(objProperties)
            Garbage_Collect(objPropertySets)
     



        Catch ex As Exception

            Garbage_Collect(objProperty)
            Garbage_Collect(objProperties)
            Garbage_Collect(objPropertySets)
        End Try
    End Sub
End Module
