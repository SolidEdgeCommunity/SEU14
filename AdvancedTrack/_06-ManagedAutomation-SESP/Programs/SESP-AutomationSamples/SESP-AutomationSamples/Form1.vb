Imports System.Runtime.InteropServices
Public Class Form1
    Public strUserName As String = "SPAdmin"
    Public strPassword As String = "sesp"
    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ' Get the type from the Solid Edge ProgID
        objSEType = Type.GetTypeFromProgID("SolidEdge.Application")
        ' Get the type from the Revision Manager ProgID
        objRevManType = Type.GetTypeFromProgID("RevisionManager.Application")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            'turn Off SESP managed mode
            blnSESPMode = False
            objSESP.SetInsightXTMode(blnSESPMode)

            'turn On SESP managed mode
            blnSESPMode = True
            objSESP.SetInsightXTMode(blnSESPMode)

        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                Dim strManagedCacheFolder As String = String.Empty
                objSESP.GetPDMCachePath(strManagedCacheFolder)
            End If



        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")
                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
            End If



        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing
                Dim intnumberofFiles As Integer = 0
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeAssemblylFile(strFilename) = True Then
                                ' 3rd argument ->false will just ckeck-out AND download    true will only download
                                'note: the last argument...  this flag controls what gets downloaded the options are the document only, the first level, ore the entire structure
                                'for an assembly is where this flag is most useful.
                                'in this example this download the entire assembly structure(all linked documents) since we might want to open it later!
                                'Note: Uses default revision rule based on preference
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadAllLevel)


                                'now open the file from the cache folder
                                Dim strCachePath As String = String.Empty
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgePartFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                Dim strCachePath As String = String.Empty
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeSheetMetalFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                Dim strCachePath As String = String.Empty
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeWeldmentFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                Dim strCachePath As String = String.Empty
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeDraftFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                Dim strCachePath As String = String.Empty
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))
                            End If
                        Next ii

                        Do While objSEApp.Documents.Count <> 0
                            objSEApp.ActiveDocument.save()
                            objSEApp.ActiveDocument.close()  'close checks in the document.... no matter what option is set to
                        Loop


                    End If
                End If
            End If



        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then

            Dim strManagedCacheFolder As String = String.Empty
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")
                objSESP.GetPDMCachePath(strManagedCacheFolder)
            End If

            'create a new part document
            Dim objSEPart As SolidEdgePart.PartDocument = Nothing
            objSEPart = objSEApp.Documents.Add("SolidEdge.PartDocument")

            'pull the next available SESP part ID/number
            Dim strItemIDPulledFromSP As String = String.Empty
            Dim strInitialRevisionPulledfromSP = String.Empty
            Dim strContentType As String = String.Empty
            'to take the default 
            'strContentType = ""
            'To specify one
            strContentType = "Part"
            objSESP.AssignItemID(strContentType, strItemIDPulledFromSP, strInitialRevisionPulledfromSP)


            'now push them to the Newly created SE part file.
            objSEPart.Properties.item("ProjectInformation").Item("Document Number").Value = strItemIDPulledFromSP
            objSEPart.Properties.Item("ProjectInformation").Item("Revision").Value = strInitialRevisionPulledfromSP

            'get the SE version loaded
            Dim strSEVersion As String = oGetSolidEdgeVersion()
            Dim splitVersion() As String = Nothing
            Dim SEMajorVersion As String = String.Empty
            splitVersion = strSEVersion.Split(".")
            SEMajorVersion = splitVersion(0)

            'if ST7 and forward use different property name
            If CInt(SEMajorVersion) >= 107 Then
                objSEPart.Properties.Item("Custom").add("SESP URL", "http://eng/Sandbox/IXTLibSandbox/Test")
            Else
                objSEPart.Properties.Item("Custom").add("Insight XT URL", "http://eng/Sandbox/IXTLibSandbox/Test")
            End If

            objSEPart.Properties.Item("Custom").add("Part Content Type", "Part")
            objSEPart.Properties.Item("Custom").add("Content Type", "SE-Part")

            'save these properties to the file property storage
            objSEPart.Properties.save()

            objSEPart.SaveAs(System.IO.Path.GetTempPath + strItemIDPulledFromSP + "_" + strInitialRevisionPulledfromSP + ".par")
            objSEPart.Close()   ' close the file so it can be imported to SP

            Dim strDBURL As String = "http://eng"

            'now import  the new file to Sharepoint
            Dim FilesToImport As System.Array
            FilesToImport = Array.CreateInstance(GetType(Object), 1)
            FilesToImport(0) = System.IO.Path.GetTempPath + strItemIDPulledFromSP + "_" + strInitialRevisionPulledfromSP + ".par"
            objSESP.ImportDocumentsToServer(1, FilesToImport, strUserName, strPassword, "", "", strDBURL, False)
            System.Array.Clear(FilesToImport, 0, 1)


        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            Dim ListOfItemRevIDsForDraft As Object = Nothing
            Dim ListOfItemRevIDsForPart As Object = Nothing
            Dim ListOfItemRevIDsForSheetmetal As Object = Nothing
            Dim ListOfItemRevIDsForAssembly As Object = Nothing
            Dim ListOfItemRevIDsForWeldment As Object = Nothing
            Dim lngNumberOfFilesFound As Long = 1
            Dim StrUserName As String = String.Empty
            Dim strFileName As String = String.Empty


            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'establish the connection to the database if has not been done already
            objSESP.ValidateLogin(StrUserName, strPassword, "", "", "http://eng")


            Dim strType As String = "SE Draft"
            objSESP.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SEDraft, StrUserName, ListOfItemRevIDsForDraft)

            strType = "SE Weldment"
            objSESP.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SEWeldment, StrUserName, ListOfItemRevIDsForWeldment)

            strType = "SE Assembly"
            objSESP.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SEAssembly, StrUserName, ListOfItemRevIDsForAssembly)

            strType = "SE Part"
            objSESP.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SEPart, StrUserName, ListOfItemRevIDsForPart)

            strType = "SE Sheetmetal"
            objSESP.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SESheetmetal, StrUserName, ListOfItemRevIDsForSheetmetal)


            'write to a text file
            strFileName = "Files found in SESP.txt"
            strFileName = System.IO.Path.GetTempPath() + strFileName

            If IO.File.Exists(strFileName) = True Then
                IO.File.Delete(strFileName)
            End If

            For ii = 0 To ListOfItemRevIDsForDraft.length / 2 - 1
                WriteToLogFile(strFileName, ListOfItemRevIDsForDraft(ii, 0) + "," + ListOfItemRevIDsForDraft(ii, 1))
                lngNumberOfFilesFound = lngNumberOfFilesFound + 1
            Next

            For ii = 0 To ListOfItemRevIDsForPart.length / 2 - 1
                WriteToLogFile(strFileName, ListOfItemRevIDsForPart(ii, 0) + "," + ListOfItemRevIDsForPart(ii, 1))
                lngNumberOfFilesFound = lngNumberOfFilesFound + 1
            Next

            For ii = 0 To ListOfItemRevIDsForSheetmetal.length / 2 - 1
                WriteToLogFile(strFileName, ListOfItemRevIDsForSheetmetal(ii, 0) + "," + ListOfItemRevIDsForSheetmetal(ii, 1))
                lngNumberOfFilesFound = lngNumberOfFilesFound + 1
            Next

            For ii = 0 To ListOfItemRevIDsForWeldment.length / 2 - 1
                WriteToLogFile(strFileName, ListOfItemRevIDsForWeldment(ii, 0) + "," + ListOfItemRevIDsForWeldment(ii, 1))
                lngNumberOfFilesFound = lngNumberOfFilesFound + 1
            Next

            For ii = 0 To ListOfItemRevIDsForAssembly.length / 2 - 1
                WriteToLogFile(strFileName, ListOfItemRevIDsForAssembly(ii, 0) + "," + ListOfItemRevIDsForAssembly(ii, 1))
                lngNumberOfFilesFound = lngNumberOfFilesFound + 1
            Next

            MessageBox.Show("Finished Searching. Found " + lngNumberOfFilesFound.ToString + "  See results in " + strFileName)
        End If



        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00655"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing

                Dim ListOfLinkedFileNames(0, 0) As Object
                Dim intnumberofFiles As Integer = 0
                Dim dwDownloadOption As Integer = 0
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeAssemblylFile(strFilename) = True Then
                                objSESP.DownladDocumentsFromServerWithOptions(strPartNumber, strRevisionNumber, strFilename, "Latest", "", True, True, dwDownloadOption, ListOfLinkedFileNames)
                                'now open the file from the cache folder
                                Dim strCachePath As String = String.Empty
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            End If
                        Next ii

                        Do While objSEApp.Documents.Count <> 0
                            objSEApp.ActiveDocument.save()
                            objSEApp.ActiveDocument.close()  'close checks in the document.... if option is set!
                        Loop

                    End If
                End If
            End If



        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00655"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing

                Dim ListOfLinkedFileNames As Object = Nothing
                Dim intnumberofFiles As Integer = 0
                Dim dwDownloadOption As Integer = 0
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeAssemblylFile(strFilename) = True Then
                                objSESP.GetListOfIndirectFilesForGivenFile(strPartNumber, strRevisionNumber, strFilename, "Latest", "", ListOfLinkedFileNames)

                                'since can only be 1 3D file exit loop
                                Exit For
                            End If
                        Next ii
                    End If
                End If
            End If
        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00655"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing
                Dim ListOfItemRevIDs As Object = Nothing
                Dim ListOfLinkedFileNames As Object = Nothing
                Dim intnumberofFiles As Integer = 0
                Dim dwDownloadOption As Integer = 0
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeAssemblylFile(strFilename) = True Then
                                'Note: the 4th argument -> True means traverse entire assembly structure  -> False means only 1 level
                                objSESP.GetBomStructure(strPartNumber, strRevisionNumber, "Latest", True, intnumberofFiles, ListOfItemRevIDs, ListOfLinkedFileNames)

                                'since can only be 1 3D file exit loop
                                Exit For
                            End If
                        Next ii
                    End If
                End If
            End If
        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing
                'Dim ListOfWhereUsedFiles As Object = Nothing
                Dim ListOfWhereUsedFiles(0) As Object
                Dim intnumberofFiles As Integer = 0
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeFile(strFilename) = True Then
                                objSESP.OnGetWhereUsedForAutomation(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision(ii - 1), ListOfWhereUsedFiles)
                            End If
                        Next ii
                    End If
                End If
            End If
        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing


                Dim intnumberofFiles As Integer = 0

                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeFile(strFilename) = True Then
                                If oIsSolidEdgeAssemblylFile(strFilename) = True Then
                                    objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, strFilename, SolidEdgeFramework.DocumentDownloadLevel.SEECDownloadTopLevel)
                                Else
                                    objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, strFilename, SolidEdgeFramework.DocumentDownloadLevel.SEECDownloadTopLevel)
                                End If
                            End If
                        Next ii
                    End If
                End If
            End If
        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing
                Dim listOfFilesToCheckIn As System.Array


                Dim intnumberofFiles As Integer = 0
                Dim SESPUrl As String = String.Empty
                Dim SESPUrlPropertyName As String = String.Empty
                Dim iCheckedOut As Integer


                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            Dim strManagedCacheFolder As String = String.Empty
                            objSESP.GetPDMCachePath(strManagedCacheFolder)
                            strFilename = strManagedCacheFolder + "\" + strFilename

                            If oIsSolidEdgeFile(strFilename) = True Then
                                listOfFilesToCheckIn = Array.CreateInstance(GetType(Object), 1)
                                listOfFilesToCheckIn(0) = strFilename

                                'need to get url from file to pass in
                                'since the property name changed in ST7
                                'get the SE version loaded
                                Dim strSEVersion As String = oGetSolidEdgeVersion()
                                Dim splitVersion() As String = Nothing
                                Dim SEMajorVersion As String = String.Empty
                                splitVersion = strSEVersion.Split(".")
                                SEMajorVersion = splitVersion(0)

                                If SEMajorVersion > 106 Then
                                    SESPUrlPropertyName = "SESP URL"
                                Else
                                    SESPUrlPropertyName = "Insight XT URL"
                                End If

                                'call a little function to read properties from a SE file
                                SESPUrl = oGetSEFileCustomProperty(strFilename, SESPUrlPropertyName)

                                'first need to make sure its checked out to me
                                '0 – Not Checked out to any body.
                                '1 – Checked out to me (i.e. current User)
                                '2 – Checked out to other user.
                                iCheckedOut = objSESP.IsInsightXTFileCheckedOut(strFilename)
                                If iCheckedOut = 1 Then
                                    objSESP.CheckInDocumentsToInsightXTServer(listOfFilesToCheckIn, False, SESPUrl)
                                    System.Array.Clear(listOfFilesToCheckIn, 0, 1)
                                End If
                            End If
                        Next ii
                    End If
                End If
            End If
        End If

        OleMessageFilter.Revoke()
    End Sub


    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strManagedCacheFolder As String = String.Empty
                objSESP.GetPDMCachePath(strManagedCacheFolder)

                Dim strfileName As String = "DH2-00009_A.par"
                strfileName = strManagedCacheFolder + "\" + strfileName
                If oIsSolidEdgeFile(strfileName) = True Then
                    Dim strSESPPartNumber As String = String.Empty
                    Dim strSESPRevisionNumber As String = String.Empty
                    objSESP.GetDocumentUID(strfileName, strSESPPartNumber, strSESPRevisionNumber)
                End If

            End If
        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing
                Dim intnumberofFiles As Integer = 0
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeFile(strFilename) = True Then
                                Dim ListOfMappedProperties As Object = Nothing
                                objSESP.GetMappedPropertiesForGivenFile(strPartNumber, strRevisionNumber, strFilename, ListOfMappedProperties)
                            End If
                        Next ii
                    End If
                End If
            End If
        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim ListOfOODFiles As Object = Nothing
                objSESP.GetOutOfDateDocuments(ListOfOODFiles)
            End If



        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim ListOfOODFiles As Object = Nothing
                objSESP.GetOutOfDateDocuments(ListOfOODFiles)
                If Not ListOfOODFiles(0) Is Nothing Then
                    objSESP.OnSynchronizeFile(ListOfOODFiles, SolidEdgeFramework.SyncOption.SEECSyncAll)
                Else
                    MessageBox.Show("No files in the cache were found ot be out-of-date!")
                    Exit Sub
                End If

                Dim ListOfOODFilesRecheck As Object = Nothing
                objSESP.GetOutOfDateDocuments(ListOfOODFilesRecheck)

                If ListOfOODFilesRecheck(0) Is Nothing Then
                    MessageBox.Show("Successfully synchronized all files in the cache")
                End If

            End If




        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing



                Dim intnumberofFiles As Integer = 0

                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeFile(strFilename) = True Then

                                Dim strManagedCacheFolder As String = String.Empty
                                objSESP.GetPDMCachePath(strManagedCacheFolder)

                                'first need to make sure its not already checked out
                                '0 – Not Checked out to any body.
                                '1 – Checked out to me (i.e. current User)
                                '2 – Checked out to other user.
                                Dim iCheckedOut As Integer = 0
                                iCheckedOut = objSESP.IsInsightXTFileCheckedOut(strManagedCacheFolder + "\" + strFilename)
                                If iCheckedOut = 0 Then
                                    objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, strFilename, SolidEdgeFramework.DocumentDownloadLevel.SEECDownloadTopLevel)

                                    'the specified file should be checked out now
                                    '...  lets do an undo checkout
                                    Dim ListOfFilestoDoUndoCheckout As System.Array
                                    ListOfFilestoDoUndoCheckout = Array.CreateInstance(GetType(Object), 1)
                                    ListOfFilestoDoUndoCheckout(0) = strManagedCacheFolder + "\" + strFilename
                                    objSESP.OnUndoCheckOutDocuments(ListOfFilestoDoUndoCheckout)
                                    System.Array.Clear(ListOfFilestoDoUndoCheckout, 0, 1)
                                End If
                            End If
                        Next ii
                    End If
                End If
            End If
        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing
                Dim intnumberofFiles As Integer = 0
                Dim strCachePath As String = String.Empty
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeAssemblylFile(strFilename) = True Then
                                ' 3rd argument ->false will just ckeck-out AND download    true will only download
                                'note: the last argument...  this flag controls what gets downloaded the options are the document only, the first level, ore the entire structure
                                'for an assembly is where this flag is most useful.
                                'in this example this download the entire assembly structure(all linked documents) since we might want to open it later!
                                'Note: Uses default revision rule based on preference
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadAllLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgePartFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeSheetMetalFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeWeldmentFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeDraftFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))
                            End If
                        Next ii


                        'now should hae a document open
                        'get a new part number and do a save as

                        'pull the next available SESP part ID/number
                        Dim strItemIDPulledFromSP As String = String.Empty
                        Dim strInitialRevisionPulledfromSP = String.Empty
                        Dim strContentType As String = "Part"

                        objSESP.AssignItemID(strContentType, strItemIDPulledFromSP, strInitialRevisionPulledfromSP)

                        ' Dim arrOldAnNewIRs As Object = Nothing
                        Dim arrOldAnNewIRs(0, 0) As Object
                        Dim strFolder As String = " http://eng/Sandbox/IXTLibSandbox/Test"
                        objSESP.SaveAsToInsightXT(strItemIDPulledFromSP, strRevisionNumber, _
                                                  strItemIDPulledFromSP + "_" + strRevisionNumber + oGetFileExtension(ListOfFileNamesInPartRevision(ii - 1)), _
                                                  "", strFolder, arrOldAnNewIRs)

                        Do While objSEApp.Documents.Count <> 0
                            objSEApp.ActiveDocument.save()
                            objSEApp.ActiveDocument.close()  'close checks in the document.... no matter what option is set to
                        Loop


                    End If
                End If
            End If



        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT

            'are we currenlty in SESP manged mode or not?
            Dim blnSESPMode As Boolean = False
            objSESP.GetInsightXTMode(blnSESPMode)

            If blnSESPMode = True Then
                'establish the connection to the database if has not been done already
                objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")

                Dim strPartNumber As String = "DH2-00009"
                Dim strRevisionNumber As String = "A"
                Dim blnFileExists As Boolean = False
                Dim ListOfFileNamesInPartRevision As Object = Nothing
                Dim intnumberofFiles As Integer = 0
                Dim strCachePath As String = String.Empty
                objSESP.DoesInsightXTFileExists(strPartNumber, strRevisionNumber, blnFileExists)
                If blnFileExists = True Then
                    objSESP.GetListOfFilesFromInsightXTServer(strPartNumber, strRevisionNumber, ListOfFileNamesInPartRevision, intnumberofFiles)
                    If intnumberofFiles > 0 Then
                        'loop through the array to fine the SE document held in the part revision
                        Dim ii As Integer = 0
                        For ii = 1 To intnumberofFiles
                            Dim strFilename As String = ListOfFileNamesInPartRevision(ii - 1)
                            If oIsSolidEdgeAssemblylFile(strFilename) = True Then
                                ' 3rd argument ->false will just ckeck-out AND download    true will only download
                                'note: the last argument...  this flag controls what gets downloaded the options are the document only, the first level, ore the entire structure
                                'for an assembly is where this flag is most useful.
                                'in this example this download the entire assembly structure(all linked documents) since we might want to open it later!
                                'Note: Uses default revision rule based on preference
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadAllLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgePartFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeSheetMetalFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeWeldmentFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))

                                'since can only be 1 3D file exit loop
                                Exit For
                            ElseIf oIsSolidEdgeDraftFile(strFilename) = True Then
                                'false will just ckeck-out AND download    true will only download
                                objSESP.CheckOutDocumentsFromInsightXTServer(strPartNumber, strRevisionNumber, False, , SolidEdgeConstants.DocumentDownloadLevel.SEECDownloadTopLevel)

                                'now open the file from the cache folder
                                objSESP.GetPDMCachePath(strCachePath)
                                objSEApp.Documents.Open(strCachePath + "\" + ListOfFileNamesInPartRevision(ii - 1))
                            End If
                        Next ii


                        'now should hae a document open
                        'get a new part number and do a save as

                        'pull the next available SESP part ID/number
                        Dim strItemIDPulledFromSP As String = String.Empty
                        Dim strInitialRevisionPulledfromSP = String.Empty
                        Dim strContentType As String = "Part"

                        objSESP.AssignItemID(strContentType, strItemIDPulledFromSP, strInitialRevisionPulledfromSP)

                        ' Dim arrOldAnNewIRs As Object = Nothing
                        Dim arrOldAnNewIRs(0, 0) As Object
                        Dim strFolder As String = " http://eng/Sandbox/IXTLibSandbox/Test"
                        objSESP.ReviseToInsightXT(strItemIDPulledFromSP, strRevisionNumber, _
                                                  strItemIDPulledFromSP + "_" + strRevisionNumber + oGetFileExtension(ListOfFileNamesInPartRevision(ii - 1)), _
                                                  "", strFolder, arrOldAnNewIRs)

                        Do While objSEApp.Documents.Count <> 0
                            objSEApp.ActiveDocument.save()
                            objSEApp.ActiveDocument.close()  'close checks in the document.... no matter what option is set to
                        Loop


                    End If
                End If
            End If



        End If

        OleMessageFilter.Revoke()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, False) = True Then
            Dim objSESP As SolidEdgeFramework.SolidEdgeInsightXT = Nothing     'Note the existence of the InsightXT nonenclature
            'hook up the to SESP automation object
            objSESP = objSEApp.SolidEdgeInsightXT
            'establish the connection to the database if has not been done already
            objSESP.ValidateLogin(strUserName, strPassword, "", "", "http://eng")
        End If

        OleMessageFilter.Revoke()
    End Sub
End Class
