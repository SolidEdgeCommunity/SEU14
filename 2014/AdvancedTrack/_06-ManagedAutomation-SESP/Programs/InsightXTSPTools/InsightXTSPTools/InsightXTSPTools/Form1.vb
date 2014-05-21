Imports System.Runtime.InteropServices
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.SharePoint




Public Class Form1

    Dim AppPath As String
    Dim IniFileLocation As String
    Dim iniFileName As String
    Dim iniBuffer(1) As String
    Dim line As String
    Dim ctr As Integer
    'Public AppPath As String
    
    















    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' Get the type from the Solid Edge ProgID
        objSEType = Type.GetTypeFromProgID("SolidEdge.Application")
        RevManType = Type.GetTypeFromProgID("RevisionManager.Application")

        Me.Label8.Text = ""
        Me.Label8.Refresh()

        Me.Label12.Text = ""
        Me.Label12.Refresh()


        Me.Text = "Solid Edge Insight XT Tools  Version:" + Application.ProductVersion
        OleMessageFilter.Register()
        Dim strURL As String = String.Empty



        If Me.TabControl1.SelectedTab.Name = "TabPage3" Then
            GoTo skipToHere
        End If





        Try
            objSEApp = Marshal.GetActiveObject("SolidEdge.Application")
        Catch ex As Exception
            MessageBox.Show("Solid Edge is not started.  Please click OK to start it and run this application!")
        End Try

        

        If oConnectToSolidEdge(True, True) = True Then

            Try
                Dim objINsightXT As SolidEdgeFramework.SolidEdgeInsightXT = Nothing
                objINsightXT = objSEApp.SolidEdgeInsightXT
                objINsightXT.GetInsightXTMode(blnIXTModeOn)

                If blnIXTModeOn = False Then
                    objINsightXT.SetInsightXTMode(True)
                End If

                If blnIXTModeOn = True Then
                    strUserName = Environment.UserName
                    objINsightXT.GetPDMCachePath(strCacheLocation)


                    Dim strXMLPath As String = String.Empty
                    Dim nEnd As Integer = 0

                    nEnd = InStr(strCacheLocation, "InsightXT", CompareMethod.Text)
                    strXMLPath = strCacheLocation.Substring(0, nEnd + 8)
                    strXMLFileName = strXMLPath + "\" + strXMLFileName
                    strURL = oReadXMLFile(strXMLFileName, "LastUsedDatabase")

                End If


            Catch ex As Exception
                MessageBox.Show("error is " + ex.Message)
            End Try


        End If



        Me.TextBoxURL.Text = strURL


        Dim arrString() As String = strURL.Split("/")
        Me.TextBoxSQLServer.Text = arrString(2)
        Me.TextBoxSQLServer.Text = "SQL\MSSQLSERVERNEWXT"

skipToHere:
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim connectionString As String = String.Empty
        Dim conn As SqlConnection
        Dim strServerName As String = Me.TextBoxSQLServer.Text
        Dim strContentDataBaseName As String = Me.TextBoxDBName.Text
        Dim strSQLUserName As String = Me.TextBoxSQLUser.Text
        Dim strSQLpassword As String = Me.TextBoxSQLPassword.Text

        connectionString = "Data Source=" + strServerName + ";Initial Catalog=" + strContentDataBaseName + ";User ID=" + strSQLUserName + ";Password=" + strSQLpassword
        'connectionString = "Data Source=" + strServerName + ";Initial Catalog=" + strContentDataBaseName + ";User ID=InsightUser;Password=InsightUser"

        Me.Cursor = Cursors.WaitCursor

        conn = New SqlConnection(connectionString)


        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim SQL As String = ""
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Error connecting to the SQL server driving INsight XT.  The error is " + ex.Message)
            End
        End Try


        'need to find site ID to the site collection used by InsightXT

        Dim strSiteID As String = "B6A4BDA5-4BCF-4223-8BB1-239718937F50"
        Dim strSiteID1 As String = ""
        SQL = "Select * from Webs"
        cmd = New System.Data.SqlClient.SqlCommand(SQL, conn)
        reader = cmd.ExecuteReader()

        While reader.Read
            If reader("Title") = "Search" Then
                strSiteID1 = reader("SiteID").ToString
                Exit While
            End If
        End While


        reader.Close()


        'SQL = "Select  distinct userinfo.tp_Login, Docs.CheckoutUserId , Docs.DirName AS Directory, Docs.LeafName AS FileName From UserInfo inner join Docs on  UserInfo.tp_ID =  CheckoutUserId  WHERE   (Docs.CheckoutUserId IS NOT NULL) and userinfo.tp_siteid = " + Chr(39) + strSiteID + Chr(39) + " and (UserInfo.tp_Login =" + Chr(39) + strUser + Chr(39) + ")"
        SQL = "Select  distinct userinfo.tp_Login, Docs.CheckoutUserId , Docs.DirName AS Directory, Docs.LeafName AS FileName From UserInfo inner join Docs on  UserInfo.tp_ID =  CheckoutUserId  WHERE   (Docs.CheckoutUserId IS NOT NULL) and userinfo.tp_siteid = " + Chr(39) + strSiteID1 + Chr(39)

        cmd = New System.Data.SqlClient.SqlCommand(SQL, conn)

        reader = cmd.ExecuteReader()
        ' Data is accessible through the DataReader object here.

        While reader.Read
            Dim strFilename As String = reader("FileName")
            Dim strCOUser As String = reader("tp_Login")
            Dim strDirName As String = reader("Directory")
            Dim strNameOfServer As String = Me.TextBoxNameOfServer.Text

            If Me.RBAllUsers.Checked = True Then
                Me.ListBox1.Items.Add("Checked Out to " + strCOUser + " file is>" + "http://" + strNameOfServer + "/" + strDirName + "/" + strFilename)
            End If

            If Me.RBSelectUser.Checked = True Then
                If strCOUser.ToUpper = Me.CBSelectedUser.Text.ToUpper Then
                    Me.ListBox1.Items.Add("Checked Out to " + strCOUser + " file is>" + "http://" + strNameOfServer + "/" + strDirName + "/" + strFilename)
                End If

            End If

        End While


        reader.Close()

        conn.Close()
        conn.Dispose()

        Me.TextBoxCount.Text = Me.ListBox1.Items.Count.ToString



        OleMessageFilter.Revoke()

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End

    End Sub

    Private Sub CBSelectedUser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CBSelectedUser.Click
        




        Try
            Dim connectionString As String = String.Empty
            Dim conn As SqlConnection
            Dim strServerName As String = Me.TextBoxSQLServer.Text
            Dim strContentDataBaseName As String = Me.TextBoxDBName.Text
            Dim strSQLUserName As String = Me.TextBoxSQLUser.Text
            Dim strSQLpassword As String = Me.TextBoxSQLPassword.Text

            Me.Cursor = Cursors.WaitCursor
            'connectionString = "Data Source=" + strServerName + ";Initial Catalog=" + strContentDataBaseName + ";User ID=InsightXTUser;Password=InsightXTUser"
            connectionString = "Data Source=" + strServerName + ";Initial Catalog=" + strContentDataBaseName + ";User ID=" + strSQLUserName + ";Password=" + strSQLpassword
            'connectionString = "Data Source=" + strServerName + ";Initial Catalog=" + strContentDataBaseName + ";User ID=InsightUser;Password=InsightUser"

            conn = New SqlConnection(connectionString)

            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader
            Dim SQL As String = ""
            Try
                conn.Open()
            Catch ex As Exception
                MessageBox.Show("Error connecting to the SQL server driving INsight XT.  The error is " + ex.Message)
                End
            End Try



            SQL = "Select distinct UserInfo.tp_login from UserInfo"
            cmd = New System.Data.SqlClient.SqlCommand(SQL, conn)
            reader = cmd.ExecuteReader()

            While reader.Read
                Me.CBSelectedUser.Items.Add(reader("tp_login").ToString)
            End While


            reader.Close()
            conn.Close()
            conn.Dispose()
        Catch ex As Exception
            MessageBox.Show("Error   The error is " + ex.Message)
        End Try

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub CBSelectedUser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBSelectedUser.SelectedIndexChanged
        'run query to find sharepoint users

        Me.ListBox1.Items.Clear()

    End Sub

    Private Sub RBAllUsers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBAllUsers.CheckedChanged
        Me.ListBox1.Items.Clear()
    End Sub

    Private Sub RBSelectUser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBSelectUser.CheckedChanged
        Me.ListBox1.Items.Clear()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CopyListBoxToClipboard(Me.ListBox1)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim strFileName As String = Nothing

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor


        If oConnectToSolidEdge(True, False) = True Then


            

            If Me.RBFindFilesBasedOnType.Checked = True Then

                Me.Label8.Text = "Searching Sharepoint database for ALL file of the specified Solid Edge file type"
                Me.Label8.Refresh()

                If Me.RBSheetmetal.Checked = True Then
                    strType = "SE SheetMetal"
                End If
                If Me.RBAssemblies.Checked = True Then
                    strType = "SE Assembly"
                End If
                If Me.RBDrafts.Checked = True Then
                    strType = "SE Draft"
                End If
                If Me.RBParts.Checked = True Then
                    strType = "SE Part"
                End If
                If Me.RBWeldments.Checked = True Then
                    strType = "SE Weldment"
                End If

                Try
                    intNumberOfLinkedItems = 0
                    ListOfLinkedItems = ""
                    ListOfLinkedItemRevisions = ""
                   

                    '**********  below uses IXT API
                    If oConnectToSolidEdge(True, False, "NO") = True Then
                        Dim objIXT As SolidEdgeFramework.SolidEdgeInsightXT = objSEApp.SolidEdgeInsightXT
                        ' ''Dim blnExists As Boolean = False
                        ' ''objIXT.DoesInsightXTFileExists("PRT-00103", "A", blnExists)

                        Dim strUserName As String = SystemInformation.UserName

                        If strType = "SE Draft" Then
                            objIXT.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SEDraft, "*", ListOfItemRevIDs)
                        ElseIf strType = "SE Weldment" Then
                            objIXT.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SEWeldment, "*", ListOfItemRevIDs)
                        ElseIf strType = "SE Assembly" Then
                            Try
                                'SolidEdgeFramework.TCESETypes.TCE_SEAssembly
                                objIXT.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SEAssembly, "*", ListOfItemRevIDs)
                            Catch ex As Exception

                            End Try
                        ElseIf strType = "SE Part" Then
                            objIXT.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SEPart, "*", _
                                ListOfItemRevIDs)
                        ElseIf strType = "SE SheetMetal" Then
                            objIXT.GetItemRevBasedOnSEType(SolidEdgeFramework.TCESETypes.TCE_SESheetmetal, "*", ListOfItemRevIDs)
                        End If

                        'read into text fileok
                        Try
                            For ii = 0 To ListOfItemRevIDs.length / 2 - 1
                                WriteToLogFile(Me.TxtFileWithListOfFilesFound.Text, ListOfItemRevIDs(ii, 0) + "," + ListOfItemRevIDs(ii, 1))
                            Next

                            intNumberOfLinkedItems = CLng(ListOfItemRevIDs.length) / 2
                        Catch ex As Exception
                            'Beep()
                        End Try

                    End If

                    Me.Label8.Text = "Found " + intNumberOfLinkedItems.ToString + _
                        " files.  See text file containing the list of files."
                    Me.Label8.Refresh()
                    Me.Button5.Enabled = True
                    'MessageBox.Show("Found " + intNumberOfLinkedItems.ToString + " Files", "test", MessageBoxButtons.OK, MessageBoxIcon.Information)


                Catch ex As Exception
                    'Beep()
                End Try



            End If

            If Me.RBValidateSPBOMMismatches.Checked = True Then
                Dim pp As Long
                Dim strFileNameToCheckOut As String = Nothing
                arrayListOfLinkedFilesAccordingToSE = New ArrayList
                ArrayAlreadyAddedToTextFile = New ArrayList
                Dim blnCheckForOccurrenceMismatch As Boolean = False

                arrayTCItems = New ArrayList
                arrayTCRevisions = New ArrayList

                strFileNameToCheckOut = ""
                If Me.RBExcel.Checked = True Then  'read and loop thru each TC item # and revision found in file
                    oReadFile(Me.TxtExcelName.Text)
                    lngNumberOfFilesToProcess = arrayTCItems.Count
                End If

                If Me.RBSpecifySingle.Checked = True Then
                    lngNumberOfFilesToProcess = 1
                    arrayTCItems.Add(Me.TxtTcItem.Text)
                    arrayTCRevisions.Add(Me.TxtTcRev.Text)
                End If


                For pp = 0 To arrayTCItems.Count - 1
                    Dim intNumberOfSPLinkedFiles As Integer = 0
                    Dim ListOfFilenamesInSharepoint As Object = Nothing
                    Dim intNumberOfFilesInSPItemRev As Integer = 0
                    Dim ListOfFilenamesInSharepoint1 As Object = Nothing
                    Dim ListOfDuplicateObjectIDs As Object = Nothing

                    If pp = 0 Then
                        Me.Label8.Text = "Validating 1 of " _
                                           + (lngNumberOfFilesToProcess).ToString + "  ... Please wait."
                        Me.Label8.Refresh()
                    Else
                        Me.Label8.Text = "Validating  " + (pp + 1).ToString + " of " _
                                           + (lngNumberOfFilesToProcess).ToString + "  ... Please wait."
                        Me.Label8.Refresh()
                    End If

                    If oConnectToSolidEdge(True, False, "NO") = True Then
                        Dim objIXT As SolidEdgeFramework.SolidEdgeInsightXT = objSEApp.SolidEdgeInsightXT
                        objIXT.GetListOfFilesFromInsightXTServer(arrayTCItems.Item(pp), _
                        arrayTCRevisions.Item(pp), ListOfFilenamesInSharepoint, intNumberOfFilesInSPItemRev)
                        strFileNameToCheckOut = ""
                        If intNumberOfFilesInSPItemRev > 0 Then
                            For ii = 0 To intNumberOfFilesInSPItemRev - 1
                                If InStr(UCase(ListOfFilenamesInSharepoint(ii)), ".ASM", CompareMethod.Text) <> 0 Then
                                    strFileNameToCheckOut = ListOfFilenamesInSharepoint(ii)  ' this is the one doc in the IR to be checked
                                    Exit For
                                End If
                            Next
                        Else
                            WriteToLogFile(Me.TxtLogFileName.Text, "No Solid Edge file found in ->" + arrayTCItems.Item(pp) + "," + arrayTCRevisions.Item(pp))
                            GoTo skip
                        End If

                        If oIsSolidEdgeAssemblyFile(strFileNameToCheckOut) = False Then
                            WriteToLogFile(Me.TxtLogFileName.Text, "No assembly file found in ->" + arrayTCItems.Item(pp) + "," + arrayTCRevisions.Item(pp))
                            GoTo skip
                        End If

                        If strFileNameToCheckOut = "" Then
                            GoTo skip
                        End If
                        intNumberOfLinkedItems = -1000

                        If Not (ListOfItemRevIDs Is Nothing) Then
                            ListOfItemRevIDs = Nothing
                        End If

                        Try
                            objIXT.GetBomStructure(arrayTCItems.Item(pp), arrayTCRevisions.Item(pp), _
                            Me.CBRevRule.Text, False, intNumberOfSPLinkedFiles, ListOfItemRevIDs, ListOfFilenamesInSharepoint1)
                        Catch ex As Exception
                            'raise error using ITK functions to determined linked TC items
                            WriteToLogFile(Me.TxtLogFileName.Text, "***Error determining if file has SharePoint BOM: " + _
                                ex.Message + "  " + arrayTCItems.Item(pp) + "/" + arrayTCRevisions.Item(pp) + " " + ex.Message)
                            GoTo skip
                        End Try

                        intNumberOfLinkedItems = intNumberOfSPLinkedFiles

                        If intNumberOfLinkedItems = 0 Then
                            BadBomCtr = BadBomCtr + 1
                            WriteToLogFile(Me.TxtLogFileName.Text, "File " + arrayTCItems.Item(pp) + "/" + _
                            arrayTCRevisions.Item(pp) + " (" + strFileNameToCheckOut + ")" + " contains empty SharePoint BOM")
                            WriteToLogFile(txtBadBomFile, arrayTCItems.Item(pp) + "," + arrayTCRevisions.Item(pp))
                        End If

                        If intNumberOfLinkedItems > 0 Then  'download the file and then open with revision manager to see how many occur SE says is there
                            If oConnectToSolidEdge(True, False, "NO") = True Then
                                objIXT.GetPDMCachePath(strCacheFolder)
                                objIXT.CheckOutDocumentsFromInsightXTServer(arrayTCItems.Item(pp), _
                                                                     arrayTCRevisions.Item(pp), True, strFileNameToCheckOut, SolidEdgeFramework.DocumentDownloadLevel.SEECDownloadTopLevel)

                                strFileName = strCacheFolder + "\" + strFileNameToCheckOut
                                If strFileName = "" Then
                                    WriteToLogFile(Me.TxtLogFileName.Text, "could not find in the cache the file for " + arrayTCItems.Item(pp) + "/" + arrayTCRevisions.Item(pp))
                                    GoTo skip
                                End If

                                If strFileName <> "" Then
                                    intSEDOCSays = 0
                                    intSEDOCSays = DetermineNumberOfFirstLevelLinkedDocuments(strFileName)
                                    arrayListOfLinkedFilesAccordingToSE.Clear()
                                End If

                                If intSEDOCSays <> intNumberOfLinkedItems Then
                                    'check if FOA
                                    Dim objPropertySets As SolidEdgeFileProperties.PropertySets = Nothing
                                    Dim blnIsFOA As Boolean = False

                                    objPropertySets = New SolidEdgeFileProperties.PropertySets
                                    objPropertySets.IsFileFamilyOfAssembly(strFileName, blnIsFOA)
                                    oReleaseObject(objPropertySets)

                                    BadBomCtr = BadBomCtr + 1
                                    If ArrayAlreadyAddedToTextFile.Contains(arrayTCItems.Item(pp) + "," + arrayTCRevisions.Item(pp)) = False Then
                                        WriteToLogFile(txtBadBomFile, arrayTCItems.Item(pp) + "," + arrayTCRevisions.Item(pp))
                                    End If
                                    If blnIsFOA = False Then
                                        WriteToLogFile(Me.TxtLogFileName.Text, "***MisMatch in number of occurrences between SharePoint and SE :" + _
                                            arrayTCItems.Item(pp) + "/" + arrayTCRevisions.Item(pp) + " (" + strFileNameToCheckOut + ")" + ">> SP says " + _
                                            intNumberOfLinkedItems.ToString + " line items and SE Document says :" + _
                                            intSEDOCSays.ToString)
                                    End If

                                    If blnIsFOA = True Then
                                        WriteToLogFile(Me.TxtLogFileName.Text, "***FOA Assembly potential false positive for MisMatch in number of occurrences between SharePoint and SE :" + _
                                            arrayTCItems.Item(pp) + "/" + arrayTCRevisions.Item(pp) + " (" + strFileNameToCheckOut + ")" + ">> SP says " + _
                                            intNumberOfLinkedItems.ToString + " line items and SE Document says :" + _
                                            intSEDOCSays.ToString)
                                    End If

                                ElseIf intSEDOCSays = intNumberOfLinkedItems Then
                                    'do nothing...  there is a match
                                    'WriteToLogFile(Me.TxtLogFileName.Text, arrayTCItems.Item(pp) + "/" + arrayTCRevisions.Item(pp) + " (" + strFileNameToCheckOut + ")" + " ->Matches!")
                                End If
                            End If
                        End If

                    End If
skip:
                    strFileName = ""
                    intSEDOCSays = 0
                    intNumberOfFilesInSPItemRev = 0
                    strFileNameToCheckOut = ""
                    intNumberOfLinkedItems = 0
                    intNumberOfSPLinkedFiles = 0


                Next


                WriteToLogFile(Me.TxtLogFileName.Text, "")
                WriteToLogFile(Me.TxtLogFileName.Text, "")
                WriteToLogFile(Me.TxtLogFileName.Text, "Number of possible issues with either count of Zero or Mismatches -> " + BadBomCtr.ToString)
            End If


        End If


        
        Me.Button7.Enabled = True
        Me.Label8.Text = "Finished Processing!"
        Me.Label8.Refresh()
        Me.Cursor = System.Windows.Forms.Cursors.Default

        OleMessageFilter.Revoke()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim strPathToNotpad As String

        Try



            System.IO.Path.GetDirectoryName(Environment.SystemDirectory)

            strPathToNotpad = System.IO.Path.GetDirectoryName(Environment.SystemDirectory) + "\notepad.exe"

            Call Shell(strPathToNotpad + " " + Me.TxtFileWithListOfFilesFound.Text, AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgBox("Error displaying file", MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub RBAssemblies_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBAssemblies.CheckedChanged

        Dim strTypeOfFile As String
        strTypeOfFile = ""
        Me.TxtFileWithListOfFilesFound.Text = ""
        ' add code to name text file to place the results in
        If Me.RBParts.Checked = True Then
            strTypeOfFile = "ListOfParts-"
        End If
        If Me.RBSheetmetal.Checked = True Then
            strTypeOfFile = "ListOfSheetMetal-"
        End If
        If Me.RBAssemblies.Checked = True Then
            strTypeOfFile = "ListOfAssemblies-"
        End If
        If Me.RBDrafts.Checked = True Then
            strTypeOfFile = "ListOfDrafts-"
        End If
        If Me.RBWeldments.Checked = True Then
            strTypeOfFile = "ListOfWeldments-"
        End If
        Try
            System.IO.Path.GetTempPath()
            Me.TxtFileWithListOfFilesFound.Text = System.IO.Path.GetTempPath + strTypeOfFile + Now.Minute.ToString + Now.Second.ToString + ".txt"


        Catch ex As Exception

        End Try
    End Sub

    Private Sub RBFindFilesBasedOnType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBFindFilesBasedOnType.CheckedChanged
        If Me.RBFindFilesBasedOnType.Checked = True Then
            Me.GBDocTypes.Enabled = True


            Me.Label3.Text = "Generate a list of Item IDs and Revisions by querying the Sharepoint database."
            Me.Label3.Refresh()
            Me.Label5.Text = "You must select the Solid Edge file type."
            Me.Label5.Refresh()


        Else
            Me.GBDocTypes.Enabled = False
            Me.TxtFileWithListOfFilesFound.Text = ""
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        With OpenFileDialog1
            .Filter = "Text Files (*.txt) | *.txt"
            .ShowDialog()
        End With
        Me.TxtExcelName.Text = OpenFileDialog1.FileName

        OpenFileDialog1.Dispose()
    End Sub

    
    Private Sub RBValidateSPBOMMismatches_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBValidateSPBOMMismatches.CheckedChanged
        Me.TxtLogFileName.Text = System.IO.Path.GetTempPath + Now.Month.ToString + "_" + Now.Day.ToString + "_" + Now.Year.ToString + "_" + Now.Hour.ToString + "_" + Now.Minute.ToString + "_" + Now.Second.ToString + "_SolidEdgeForSharepoint_Utilities_Report.txt"

        txtBadBomFile = oGetFileNameWithoutExtension(Me.TxtLogFileName.Text)
        txtBadBomFile = txtBadBomFile + "LIST OF BAD BOMs.txt"
        txtBadBomFile = oGetPathOfFilename(Me.TxtLogFileName.Text) + txtBadBomFile
    End Sub

    Private Sub RBExcel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBExcel.CheckedChanged
        If RBExcel.Checked = True Then
            Me.Button6.Enabled = True
            Me.TxtExcelName.Enabled = True
        Else
            Me.Button6.Enabled = False
            Me.TxtExcelName.Enabled = False
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        End

    End Sub

    Private Sub RBSpecifySingle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBSpecifySingle.CheckedChanged
        If RBSpecifySingle.Checked = True Then
            Me.TxtTcRev.Enabled = True
            Me.TxtTcItem.Enabled = True
        Else
            Me.TxtTcRev.Enabled = False
            Me.TxtTcItem.Enabled = False
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Dim strPathToNotpad As String

        Try



            System.IO.Path.GetDirectoryName(Environment.SystemDirectory)

            strPathToNotpad = System.IO.Path.GetDirectoryName(Environment.SystemDirectory) + "\notepad.exe"

            Call Shell(strPathToNotpad + " " + Me.TxtLogFileName.Text, AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgBox("Error displaying file", MsgBoxStyle.OkOnly)
        End Try


        



    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim strSiteCollectionUrl As String = TextBoxChekinURL.Text

        If IsValidSiteCollectionUrl(strSiteCollectionUrl) = True Then
            'MessageBox.Show("Valid URL given")
            arrayUrlOfFilesToCheckIn = New ArrayList
            Dim strtxtFileContainingUrls As String = Me.TextBoxFilesToBeCheckedIn.Text
            Dim iCounter As Long = 0
            Dim ii As Integer = 0
            oReadFileGetURLs(strtxtFileContainingUrls)

            'MessageBox.Show("Found " + arrayUrlOfFilesToCheckIn.Count.ToString + " files")


            If arrayUrlOfFilesToCheckIn.Count > 0 Then
                Dim oSiteCollection As New SPSite(strSiteCollectionUrl)
                Dim oWebs As SPWebCollection = Nothing
                oWebs = oSiteCollection.AllWebs

                For iCounter = 0 To arrayUrlOfFilesToCheckIn.Count - 1
                    Dim srcFile As SPFile = Nothing
                    Label12.Text = " Checking in file #" + iCounter.ToString + "->" + arrayUrlOfFilesToCheckIn(iCounter)
                    Label12.Refresh()
                    For ii = 1 To oWebs.Count - 1
                        Dim srcSite As SPWeb = oWebs.Item(ii)

                        ' MessageBox.Show("searching in web site " + srcSite.Name)
                        Try
                            srcFile = srcSite.GetFile(arrayUrlOfFilesToCheckIn(iCounter))
                            srcFile.CheckIn("Checked in via utility", SPCheckinType.MajorCheckIn)

                        Catch ex As Exception

                        End Try

                    Next ii

                Next iCounter
                oReleaseObject(oWebs)
                oReleaseObject(oSiteCollection)

            End If

            Label12.Text = " Finished processing"
            Label12.Refresh()

            MessageBox.Show("Finished checking in the specified documents")


        End If

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        With OpenFileDialog2
            .Filter = "Text Files (*.txt) | *.txt"
            .ShowDialog()
        End With
        Me.TextBoxFilesToBeCheckedIn.Text = OpenFileDialog2.FileName

        OpenFileDialog2.Dispose()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        End

    End Sub

    Private Sub TextBoxChekinURL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxChekinURL.TextChanged
        Dim strTemp As String = Me.TextBoxChekinURL.Text


        Dim nstart As Integer = 0

        nstart = InStr(strTemp, "//", CompareMethod.Text)
        If nstart <> 0 Then
            strTemp = Mid(strTemp, nstart + 2, Len(strTemp))
            nstart = 0
            nstart = InStr(strTemp, "/", CompareMethod.Text)
            If nstart <> 0 Then
                strTemp = Mid(strTemp, 1, nstart - 1)
                Me.TextBoxNameOfServer.Text = strTemp
                Me.TextBoxNameOfServer.Refresh()
                Exit Sub
            End If
        End If


    End Sub

    Private Sub TextBoxURL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxURL.TextChanged
        Dim strTemp As String = Me.TextBoxURL.Text


        Dim nstart As Integer = 0

        nstart = InStr(strTemp, "//", CompareMethod.Text)
        If nstart <> 0 Then
            strTemp = Mid(strTemp, nstart + 2, Len(strTemp))
            nstart = 0
            nstart = InStr(strTemp, "/", CompareMethod.Text)
            If nstart <> 0 Then
                strTemp = Mid(strTemp, 1, nstart - 1)
                Me.TextBoxNameOfServer.Text = strTemp
                Me.TextBoxNameOfServer.Refresh()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Me.ListBoxGUIDResults.Items.Clear()


        Dim connectionString As String = String.Empty
        Dim conn As SqlConnection
        Dim strServerName As String = Me.TextBoxSQLserverGUID.Text
        Dim strContentDataBaseName As String = Me.TextBoxDBNameGUID.Text
        Dim strSQLUserName As String = Me.TextBoxSQLUserGUID.Text
        Dim strSQLpassword As String = Me.TextBoxSQLPasswordGUID.Text


        connectionString = "Data Source=" + strServerName + ";Initial Catalog=" + strContentDataBaseName + ";User ID=" + strSQLUserName + ";Password=" + strSQLpassword
        'connectionString = "Data Source=" + strServerName + ";Initial Catalog=" + strContentDataBaseName + ";User ID=InsightUser;Password=InsightUser"

        Me.Cursor = Cursors.WaitCursor

        conn = New SqlConnection(connectionString)


        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim SQL As String = ""
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Error connecting to the SQL server driving INsight XT.  The error is " + ex.Message)
            End
        End Try

        Dim strFileName As String = ""
        Dim strFolder As String = ""

        Try
            SQL = "SELECT AllDocs.DirName,AllDocs.LeafName, AllUserData.nvarchar8 FROM AllDocs INNER JOIN AllUserData ON AllDocs.SiteId = AllUserData.tp_SiteId AND AllDocs.DoclibRowId = AllUserData.tp_ID AND AllDocs.ListId = AllUserData.tp_ListId where nvarchar8 = '" + Me.TextBoxGUID.Text + "'"
            cmd = New System.Data.SqlClient.SqlCommand(SQL, conn)
            reader = cmd.ExecuteReader()


            'might want to check to make sure only 1 hit on the GUID....
            While reader.Read
                If reader.HasRows = True Then

                    strFileName = reader("LeafName").ToString
                    strFolder = reader("DirName").ToString
                    Me.ListBoxGUIDResults.Items.Add("Filename is ->" + strFileName)
                    Me.ListBoxGUIDResults.Items.Add("Folder is ->" + strFolder)


                End If
                strFileName = ""
                strFolder = ""
            End While


            reader.Close()
        Catch ex As Exception

        End Try
        

        

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        End

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim strSiteCollectionUrl As String = TextBoxEmptyPartRevisions.Text
        Dim ii As Integer
        Dim jj As Integer
        Dim kk As Integer
        Dim pp As Integer
        Dim blnSEFileFound As Boolean = False
        Dim intCTR As Long = 0

        If IsValidSiteCollectionUrl(strSiteCollectionUrl) = True Then
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

            Me.ListBox2.Items.Clear()
            Me.ListBox2.Items.Add("Part Revisions with no Solid Edge files stored in them OR invalid GUIDS:")


            Dim oSiteCollection As New SPSite(strSiteCollectionUrl)
            Dim oWebs As SPWebCollection = Nothing
            oWebs = oSiteCollection.AllWebs

            Label21.Text = "Processing ....  Please wait."
            Label21.Refresh()
            For ii = 1 To oWebs.Count - 1
                Dim srcSite As SPWeb = oWebs.Item(ii)
                Dim strName As String = srcSite.Name

               
                Me.ListBox2.Items.Add("checking ->" + strName)
                Me.ListBox2.Refresh()



                Try
                    Dim spListCollection As SPListCollection = srcSite.Lists
                    For jj = 0 To spListCollection.Count - 1
                        Dim oList As SPList = spListCollection.Item(jj)
                        oList.EnableThrottling = False  '????

                        Dim oItemCollection As SPListItemCollection = oList.Folders
                        For kk = 0 To oItemCollection.Count - 1
                            Dim oItem As SPListItem = oItemCollection.Item(kk)
                            Dim strName1 As String = oItem.DisplayName
                            Dim strPulledContentType As String = oItem.ContentType.Name.ToLower
                            Label21.Text = "Processing folder " + oItem.DisplayName
                            Label21.Refresh()

                            If strPulledContentType = "part" Then
                                'this gets the "part" folder
                                'can get the folder properties from here!
                            End If

                            If strPulledContentType = "part-revision" Then
                                'this gets the "part-revision" folder
                                'can get the folder properties from here!
                                Dim ofiles As SPFileCollection = oItem.Folder.Files
                                For pp = 0 To ofiles.Count - 1
                                    Dim oFile As SPFile = ofiles.Item(pp)
                                    Dim spFileItem As SPListItem = oFile.Item
                                    Dim strfilename As String = spFileItem.Name
                                    Dim strExt As String = strfilename.Split(".").Last.ToLower
                                    If strExt = "asm" Or strExt = "par" Or strExt = "psm" Or strExt = "dft" Or strExt = "pwd" Then
                                        blnSEFileFound = True
                                    End If

                                    If blnSEFileFound = True Then
                                        Try
                                            If oFile.Properties.ContainsKey("PDM-GUID") = True Then
                                                Dim strJunk As String = oFile.Properties.Item("PDM-GUID").ToString
                                                If strJunk.StartsWith("000000") Then
                                                    blnSEFileFound = False
                                                End If
                                            End If
                                            If oFile.Properties.ContainsKey("Rev-PDM-GUID") = True Then
                                                Dim strJunk1 As String = oFile.Properties.Item("Rev-PDM-GUID").ToString
                                                If strJunk1.StartsWith("000000") Then
                                                    blnSEFileFound = False
                                                End If
                                            End If

                                            If oFile.Properties.ContainsKey("Item-PDM-GUID") = True Then
                                                Dim strJunk2 As String = oFile.Properties.Item("Item-PDM-GUID").ToString
                                                If strJunk2.StartsWith("000000") Then
                                                    blnSEFileFound = False
                                                End If
                                            End If
                                        Catch ex As Exception
                                            MessageBox.Show("error getting GUIDs" + ex.Message)
                                        End Try
                                    End If
                                Next pp

                                If blnSEFileFound = False Then
                                    Me.ListBox2.Items.Add(strName1)
                                    Me.ListBox2.Refresh()
                                    intCTR = intCTR + 1
                                End If
                                blnSEFileFound = False
                            End If

                           

                        Next kk

                    Next jj


                Catch ex As Exception
                    MessageBox.Show("encountered an error -> " + ex.Message)
                End Try





            Next ii

            Dim cnt As Integer = Me.ListBox2.Items.Count - 1
            Me.Label21.Text = "Found " + intCTR.ToString
            Me.Cursor = System.Windows.Forms.Cursors.Default

        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        End
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        CopyListBoxToClipboard(Me.ListBox2)
    End Sub
End Class
