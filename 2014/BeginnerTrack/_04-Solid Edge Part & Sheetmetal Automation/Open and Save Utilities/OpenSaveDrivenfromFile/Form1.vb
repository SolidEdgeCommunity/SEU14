Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.FileIO

Public Class Form1
    Private FolderPath, Filter, Filters(5), strFolders(1), StrUnMngdFolders(1) As String
    Private FilterSet(5) As Boolean
    Private nFilters As Short
    Private strSEFileType() As String = {"*.asm", "*.dft", "*.par", "*.psm", "*.pwd"}
    Dim maxlen As Short
    Dim intTotFileNameCount As Long
    Dim strUpdateFileList, strProcessFiles, Quote As String
    Dim Slash As String



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .Filter = "Input Files (*.txt,*.csv) | *.txt;*.csv"
            .ShowDialog()
        End With
        Me.TxtTextFilename.Text = OpenFileDialog1.FileName
        strInputFilename = Me.TxtTextFilename.Text
        Me.btnProcess.Enabled = True

        strLogFilename = GetFilePath(strInputFilename) + "\OpenSaveLog.txt"
        Me.TxtLogFileName.Text = strLogFilename

        If strInputFilename <> "" Then
            Me.btnProcess.Enabled = True
            Label1.Text = " "
            Label1.Refresh()

            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
        End If



        OpenFileDialog1.Dispose()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim FSO2 As Scripting.FileSystemObject = Nothing
        Dim strPathToNotpad As String

        Try
            FSO2 = New Scripting.FileSystemObject
            strPathToNotpad = FSO2.GetSpecialFolder(Scripting.SpecialFolderConst.WindowsFolder).Path + "\notepad.exe"

            Call Shell(strPathToNotpad + " " + strLogFilename, AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgBox("Error displaying file", MsgBoxStyle.OkOnly)
        End Try


        If Not (FSO2 Is Nothing) Then
            FSO2 = Nothing
        End If


        Garbage_Collect(FSO2)
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Label2.Text = ""
        Me.Label2.Refresh()

        Me.Label1.Text = ""
        Me.Label1.Refresh()

        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        StrFileNames.Clear()

        Quote = Chr(34)
        Slash = "\"

        strUpdateFileList = " Select " & Quote & "Update File List" & Quote
        strProcessFiles = "Files to be processed"

        FolderPath = System.Windows.Forms.Application.StartupPath()
        FolderBrowserDialog1.SelectedPath = FolderPath
        strFolders(0) = FolderPath
        'strDate = System.DateTime.Now.Hour.ToString + "_" + System.DateTime.Now.Minute.ToString + "_" + System.DateTime.Now.Second.ToString + " " + System.DateTime.Now.Month.ToString + "-" + System.DateTime.Now.Day.ToString + "-" + System.DateTime.Now.Year.ToString

        lstPath.Items.Add(FolderPath)

        chkPart.CheckState = CheckState.Checked
        chkSheetmetal.CheckState = CheckState.Checked
        chkAssembly.CheckState = CheckState.Checked
        chkDraft.CheckState = CheckState.Checked
        chkWeldment.CheckState = CheckState.Checked

        Label1.Text = ""
        Label1.Refresh()

        If StrFileNames.Count = 0 Then
            Label2.Text = "No files found, select options or browse to new location"
            Label2.Refresh()
        Else
            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
        End If


        strAppVersion = System.Windows.Forms.Application.ProductVersion
        Me.Text = "Open and Save Data Preparation Utilities    " + strAppVersion


        Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub


   
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcess.Click

        'set the appropriate flags based on user input

        If Me.CheckBoxCalloutBOMFindandReplace.Checked = True Then
            'write values to a text file to be retrieved on later run
            If System.IO.File.Exists(System.AppDomain.CurrentDomain.BaseDirectory + "SearchAndReplaceLinkedPropertyText.txt") = True Then
                System.IO.File.Delete(System.AppDomain.CurrentDomain.BaseDirectory + "SearchAndReplaceLinkedPropertyText.txt")
            End If
            arrDataLinkedPropertyText = New ArrayList
            arrSearchForLinkedPropertyText = New ArrayList
            arrReplaceWithLinkedPropertyText = New ArrayList

            Dim nn As Integer = 0
            For nn = 0 To Me.ListBoxFindReplaceLinkedPropertyText.Items.Count - 1
                WriteToLogFile(System.AppDomain.CurrentDomain.BaseDirectory + "SearchAndReplaceLinkedPropertyText.txt", Me.ListBoxFindReplaceLinkedPropertyText.Items(nn).ToString)
                arrDataLinkedPropertyText.Add(Me.ListBoxFindReplaceLinkedPropertyText.Items(nn).ToString)
            Next


        End If



        If Me.CheckBoxAddHardWare.Checked = True Then
            blnCheckHardwareOption = True
            strHardwarePropertyNameToCheck = Me.TextBoxCheckProperty.Text


            stringArray = Split(strHardwarePropertyNameToCheck, "=")
            strHardwarePropertyNameToCheck = stringArray(0)
            strHardwarePropertyValue = stringArray(1)




        End If


        If Me.CheckBoxCreatPreview.Checked = True Then
            blnCreatePreview = True
        End If

        If Me.CheckBoxResetBodyStyle.Checked = True Then
            blnResetBodyStyle = True
        End If

        If Me.CheckBoxAlreadyImported.Checked = True Then
            blnCheckIfAlreadyImported = True
            arrayOfImportedFiles = New ArrayList
            arrayOfNotImportedFiles = New ArrayList
        End If


        If Me.chkAssembly.Checked = True Then
            blnProcessASMs = True
        End If

        If Me.chkPart.Checked = True Then
            blnProcessPARs = True
        End If

        If Me.chkSheetmetal.Checked = True Then
            blnProcessPSMs = True
        End If

        If Me.chkDraft.Checked = True Then
            blnProcessDFTs = True
        End If

        If Me.chkWeldment.Checked = True Then
            blnProcessPWDs = True
        End If

        If Me.CheckBoxRemoveProperties.Checked = True Then
            blnRemoveProperties = True
            strPropsToRemove = Me.TextBoxPropsToRemove.Text
        End If

        If Me.CBExtractPreview.Checked = True Then
            blnExtractPreviewBMP = True
        End If

        If Me.CBTurnOffDisplayNextHighestAssyOccProp.Checked = True Then
            blnTurnOffDisplayNextOccProp = True
        End If

        If Me.CBCreadtePDFFromDraft.Checked = True Then
            blnCreatePDFsFromDrafts = True
        End If

        If Me.CBXRecompute.Checked = True Then
            blnRecompute = True
        End If

        If Me.CBFitAndShade.Checked = True Then
            blnFitAndShade = True
        End If

        If Me.CBCheckAssemblyForCorruptLink.Checked = True Then
            blnCheckAssemblyForCorruptLinks = True
        End If

        If Me.CBUpdateDrafts.Checked = True Then
            blnUpdateDrawingViews = True
        End If

        If Me.CBRsetFileStatusToAvailable.Checked = True Then
            blnResetStatus = True
        End If

        If Me.RBListFormFile.Checked = True Then
            blnReadFilesFromFile = True
        End If

        If Me.CBCopyStyles.Checked = True Then
            blnCopyStyles = True
        End If

        If Me.CBUpDateAssemblyLinks.Checked = True Then
            blnUpdateLinks = True
        End If

        If Me.CheckBoxTurnOffCSs.Checked = True Then
            blnTurnOffCss = True
        End If

        If Me.CheckBoxReplaceCharacterInFilename.Checked = True Then
            blnReplaceCharacters = True
            strOldChar = Me.TextBoxOldCharacter.Text

            If Me.TextBoxNewCharacter.Text = String.Empty Then
                strNewChar = ""
            Else
                strNewChar = Me.TextBoxNewCharacter.Text

            End If

            If strOldChar = String.Empty Then
                MessageBox.Show("you must enter the old character(s)!")
                Exit Sub
            End If


            If strNewChar = "" Then
                Dim intReturnValue As Integer
                intReturnValue = MessageBox.Show("Are you sure you want to REMOVE not replace the character " + strOldChar + " !", "Question", _
                MessageBoxButtons.OKCancel, MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)

                If (intReturnValue = DialogResult.Cancel) Then
                    Exit Sub
                End If
            End If

        End If


        If Me.CheckBoxCalloutBOMFindandReplace.Checked = True Then
            blnFindReplaceCalloutBOM = True
        End If


        If Me.CheckBoxTurnOffGradientBackground.Checked = True Then
            blnTurnOffGradientBackground = True
        End If



        Call Process()
    End Sub

    Private Sub RBCreatePDF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub

    Private Sub RBListFormFile_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBListFormFile.CheckedChanged
        If Me.RBListFormFile.Checked = True Then
            Me.Button1.Enabled = True
            Me.TxtTextFilename.Enabled = True
            Me.GrpDocumentTypes.Enabled = True
            Me.GrpFolderOptions.Enabled = True
            Me.lstFiles.Enabled = False
            Me.txtFilesFound.Text = ""
            Me.txtFilesFound.Refresh()
            Me.lstPath.Items.Clear()

        Else
            Me.Button1.Enabled = False
            Me.TxtTextFilename.Enabled = False

            Me.GrpDocumentTypes.Enabled = True
            Me.GrpFolderOptions.Enabled = True
            Me.lstFiles.Enabled = True

        End If
    End Sub

    Private Sub TxtTextFilename_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtTextFilename.TextChanged

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FolderBrowserDialog1.ShowDialog()
        FolderPath = FolderBrowserDialog1.SelectedPath
        FileSystem.CurrentDirectory = FolderPath
        strLogFilename = FolderPath + "\OpenSaveLog.txt"
        strFolders(0) = FolderPath

        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        ' if processing subfolders as well, then let's get the folder list here
        If optAllFiles.Checked Then
            GetSubFolders(FolderPath)
        End If

        FolderList()
        LoadSEFiles()

        Label1.Text = ""
        Label1.Refresh()

        If StrFileNames.Count = 0 Then
            Label2.Text = "No files found, select options or browse to new location"
            Label2.Refresh()
        Else
            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
        End If

    End Sub

    Private Sub LoadSEFiles()
        ' Loops through the selected folders to populate the file list
        If optAllInDirectory.Checked And lstPath.SelectedItems.Count = 0 Then
            SetFileList(FolderPath)
        Else
            For i = 0 To UBound(strFolders) - 1
                SetFileList(strFolders(i))
            Next
        End If
    End Sub


    Private Sub FileList(ByVal Path As String, ByVal Filter As String)
        Dim FileName As String
        Dim lenPath, len
        Dim blnUseSearchScope, blnSubfolders As Boolean

        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        blnSubfolders = False
        blnUseSearchScope = False
        lenPath = FolderPath.Length
        ChDrive(Path) 'FileList.Path
        ChDir(Path)

        Label1.Text = "Generating file list"
        Label1.Refresh()
        For Each foundFile As String In FileSystem.GetFiles(Path, SearchOption.SearchTopLevelOnly, Filter)
            len = foundFile.Length
            FileName = foundFile.Substring(lenPath + 1, len - lenPath - 1)
            lstFiles.Items.Add(foundFile)
            StrFileNames.Add(foundFile)
        Next



        If lstFiles.Items.Count > 0 Then
            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
            ' Get total number of files in the list
            intTotFileNameCount = lstFiles.Items.Count
            btnProcess.Enabled = True
        Else
            'Label2.Text = "No files found, select options or browse to new location"
            btnProcess.Enabled = False

        End If

        txtFilesFound.Text = lstFiles.Items.Count
        txtFilesFound.Update()



    End Sub

    Private Sub SetFileList(ByVal strPath)
        If FolderPath = strPath Then lstFiles.Items.Clear()
        If chkPart.Checked Then
            FileList(strPath, "*.par")
        End If
        If chkSheetmetal.Checked Then
            FileList(strPath, "*.psm")
        End If
        If chkAssembly.Checked Then
            FileList(strPath, "*.asm")
            If Me.CheckBox1TimeFixCFGs.Checked = True Then
                FileList(strPath, "*.cfg")  ' temporary for Madison
            End If

        End If
        If chkDraft.Checked Then
            FileList(strPath, "*.dft")
        End If
        If chkWeldment.Checked Then
            FileList(strPath, "*.pwd")
        End If
        SetFilters()

    End Sub

    Sub SetFilters()
        nFilters = -1
        If chkPart.Checked Then
            nFilters = nFilters + 1
            Filters(nFilters) = "*.par"
        End If
        If chkSheetmetal.Checked Then
            nFilters = nFilters + 1
            Filters(nFilters) = "*.psm"
        End If
        If chkAssembly.Checked Then
            nFilters = nFilters + 1
            Filters(nFilters) = "*.asm"
        End If
        If chkDraft.Checked Then
            nFilters = nFilters + 1
            Filters(nFilters) = "*.dft"
        End If
        If chkWeldment.Checked Then
            nFilters = nFilters + 1
            Filters(nFilters) = "*.pwd"
        End If

    End Sub



    Private Sub GetSubFolders(ByVal strDir)
        Dim courier As String
        Dim strNewDirs() As String

        Label1.Text = "Searching for subfolders..."
        Label1.Refresh()

        courier = Dir(strDir & "\", FileAttribute.Directory) ' Retrieve the first entry.
        Do While courier <> "" ' Start the loop.
            ' Ignore the current directory and the encompassing directory.
            If courier <> "." And courier <> ".." Then
                ' Use bitwise comparison to make sure MyName is a directory.
                If (GetAttr(strDir & "\" & courier) And FileAttribute.Directory) = FileAttribute.Directory Then
                    If Not IsInitializedArrayOfStrings(strNewDirs) Then
                        ReDim strNewDirs(0)
                    Else
                        ReDim Preserve strNewDirs(UBound(strNewDirs) + 1)
                    End If
                    strNewDirs(UBound(strNewDirs)) = strDir & "\" & courier

                    ' Let's hold on to all the folders we will consider
                    ReDim Preserve strFolders(UBound(strFolders) + 1)
                    strFolders(UBound(strFolders) - 1) = strDir & "\" & courier
                    Label2.Text = "Number of subfolders found " + (UBound(strFolders) + 1).ToString
                    Label2.Refresh()
                End If ' it represents a directory.
            End If
            courier = Dir() ' Get next entry.
        Loop

        If IsInitializedArrayOfStrings(strNewDirs) Then
            For i = LBound(strNewDirs) To UBound(strNewDirs)
                GetSubFolders(strNewDirs(i))
            Next
        End If

    End Sub

    Private Sub FolderList()
        Dim i As Short
        lstPath.Items.Clear()
        For i = 0 To UBound(strFolders) - 1
            lstPath.Items.Add(strFolders(i))
        Next

    End Sub

    Private Function IsInitializedArrayOfStrings(ByRef strArray() As String) As Boolean

        Dim lL As Integer 'lower bound of array
        Dim lU As Integer 'upper bound of array

        Try
            lL = LBound(strArray)
            lU = UBound(strArray)
            IsInitializedArrayOfStrings = (lU >= lL)
            Exit Function
        Catch ex As Exception
            IsInitializedArrayOfStrings = False
        End Try

    End Function

    Private Sub optAllFiles_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAllFiles.CheckedChanged
        Dim i As Short

        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        StrFileNames.Clear()
        If optAllFiles.Checked Then
            lstFiles.SelectedItems.Clear()
            GetSubFolders(FolderPath)
            FolderList()
            LoadSEFiles()
            Label1.Text = ""
            Label1.Refresh()

            If StrFileNames.Count = 0 Then
                Label2.Text = "No files found, select options or browse to new location"
                Label2.Refresh()
            Else
                Label2.Text = "Files found, proceed with processing or select options"
                Label2.Refresh()
            End If

            Label1.Text = " "
            Label1.Refresh()

            Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Private Sub optAllInDirectory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAllInDirectory.CheckedChanged
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        StrFileNames.Clear()
        If optAllInDirectory.Checked Then
            lstFiles.SelectedItems.Clear()
            lstPath.Items.Clear()
            If FolderPath <> "" Then lstPath.Items.Add(FolderPath)
            ' should not be necessary but optSelected change to other does not update correctly otherwise
            optAllInDirectory.Checked = True
            ReDim strFolders(1)
            strFolders(0) = FolderPath
            LoadSEFiles()

            Label1.Text = ""
            Label1.Refresh()

            If StrFileNames.Count = 0 Then
                Label2.Text = "No files found, select options or browse to new location"
                Label2.Refresh()
            Else
                Label2.Text = "Files found, proceed with processing or select options"
                Label2.Refresh()
            End If

            Cursor.Current = System.Windows.Forms.Cursors.Default

        End If
    End Sub

    Private Sub chkPart_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPart.CheckedChanged
        StrFileNames.Clear()
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If chkPart.Checked Then
            FilterSet(0) = True
        Else
            FilterSet(0) = False
        End If
        SetFilters()
        LoadFileList()

        If StrFileNames.Count = 0 Then
            Label2.Text = "No files found, select options or browse to new location"
            Label2.Refresh()
        Else
            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
        End If

        Label1.Text = " "
        Label1.Refresh()

        Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub LoadFileList()
        lstFiles.Items.Clear()

        LoadSEFiles()
    End Sub

    Private Sub chkSheetmetal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSheetmetal.CheckedChanged
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        StrFileNames.Clear()
        If chkSheetmetal.Checked Then
            FilterSet(1) = True
        Else
            FilterSet(1) = False
        End If

        LoadFileList()

        If StrFileNames.Count = 0 Then
            Label2.Text = "No files found, select options or browse to new location"
            Label2.Refresh()
        Else
            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
        End If

        Label1.Text = " "
        Label1.Refresh()

        Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub chkAssembly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAssembly.CheckedChanged
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        StrFileNames.Clear()
        If chkAssembly.Checked Then
            FilterSet(2) = True
        Else
            FilterSet(2) = False
        End If
        SetFilters()

        LoadFileList()

        If StrFileNames.Count = 0 Then
            Label2.Text = "No files found, select options or browse to new location"
            Label2.Refresh()
        Else
            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
        End If

        Label1.Text = " "
        Label1.Refresh()

        Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub chkDraft_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDraft.CheckedChanged
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        StrFileNames.Clear()
        If chkDraft.Checked Then
            FilterSet(3) = True
        Else
            FilterSet(3) = False
        End If
        SetFilters()

        LoadFileList()

        If StrFileNames.Count = 0 Then
            Label2.Text = "No files found, select options or browse to new location"
            Label2.Refresh()
        Else
            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
        End If

        Label1.Text = " "
        Label1.Refresh()

        Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub chkWeldment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkWeldment.CheckedChanged
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        StrFileNames.Clear()
        If chkWeldment.Checked Then
            FilterSet(4) = True
        Else
            FilterSet(4) = False
        End If
        SetFilters()

        LoadFileList()

        If StrFileNames.Count = 0 Then
            Label2.Text = "No files found, select options or browse to new location"
            Label2.Refresh()
        Else
            Label2.Text = "Files found, proceed with processing or select options"
            Label2.Refresh()
        End If

        Label1.Text = " "
        Label1.Refresh()

        Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub RBTraversFolders_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBTraversFolders.CheckedChanged
        If Me.RBTraversFolders.Checked = True Then
            Me.Button5.Enabled = True
        Else
            Me.Button5.Enabled = False
        End If
    End Sub

    Private Sub CBFitAndShade_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBFitAndShade.CheckedChanged
        If CBFitAndShade.Checked = True Then
            Me.CheckBoxTurnOffCSs.Enabled = True
            Me.CheckBoxTurnOffGradientBackground.Enabled = True
        End If

        If CBFitAndShade.Checked = False Then
            Me.CheckBoxTurnOffCSs.Enabled = False
            Me.CheckBoxTurnOffCSs.Checked = False
            Me.CheckBoxTurnOffGradientBackground.Enabled = False
            Me.CheckBoxTurnOffGradientBackground.Checked = False
        End If
    End Sub

    Private Sub CheckBoxReplaceCharacterInFilename_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxReplaceCharacterInFilename.CheckedChanged
        If CheckBoxReplaceCharacterInFilename.Checked = True Then
            Me.TextBoxNewCharacter.Enabled = True
            Me.TextBoxOldCharacter.Enabled = True
        Else
            Me.TextBoxNewCharacter.Enabled = False
            Me.TextBoxOldCharacter.Enabled = False
        End If
    End Sub

    Private Sub CheckBoxCalloutBOMFindandReplace_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxCalloutBOMFindandReplace.CheckedChanged
        If CheckBoxCalloutBOMFindandReplace.Checked = True Then
            Me.ListBoxFindReplaceLinkedPropertyText.Enabled = True
            Me.Button2.Enabled = True
            Me.Button6.Enabled = True
            Me.TextBoxFindReplaceLinkPropertyText.Enabled = True
            Me.CheckBoxRemoveProperties.Checked = True
            Try
                If System.IO.File.Exists(System.AppDomain.CurrentDomain.BaseDirectory + "SearchAndReplaceLinkedPropertyText.txt") = True Then
                    Dim arrData As String() = System.IO.File.ReadAllLines(System.AppDomain.CurrentDomain.BaseDirectory + "SearchAndReplaceLinkedPropertyText.txt")
                    Dim strData As String

                    For Each strData In arrData
                        If strData = "" Then
                            Continue For
                        End If
                        If InStr(strData.ToUpper, "#", Microsoft.VisualBasic.CompareMethod.Text) <> 0 Then
                            GoTo skip
                        End If

                        If strData <> "" Then
                            'Dim arrStrDataSplit As String() = strData.Split(",")
                            Me.ListBoxFindReplaceLinkedPropertyText.Items.Add(strData)

                            Dim strPulledContents As String = Me.TextBoxPropsToRemove.Text
                            Dim charArray() As String = strData.Split(",")
                            Dim nloc As Integer = 0

                            nloc = InStr(charArray(0), "/", Microsoft.VisualBasic.vbTextCompare)
                            If nloc <> 0 Then
                                If strPulledContents = "" Then
                                    strPulledContents = Trim(charArray(0).ToString.Substring(2, nloc - 3))
                                    GoTo skip1
                                Else
                                    If strPulledContents.Contains(charArray(0).ToString.Substring(2, nloc - 3)) = False Then
                                        strPulledContents = Trim(strPulledContents + "," + charArray(0).ToString.Substring(2, nloc - 3))
                                        GoTo skip1
                                    End If
                                End If
                            End If

                            nloc = InStr(charArray(0), "|", Microsoft.VisualBasic.vbTextCompare)
                            If nloc <> 0 Then
                                If strPulledContents = "" Then
                                    strPulledContents = Trim(charArray(0).ToString.Substring(2, nloc - 3))
                                Else
                                    If strPulledContents.Contains(charArray(0).ToString.Substring(2, nloc - 3)) = False Then
                                        strPulledContents = Trim(strPulledContents + "," + charArray(0).ToString.Substring(2, nloc - 3))
                                    End If

                                    End If
                            End If
skip1:
                            Me.TextBoxPropsToRemove.Text = strPulledContents

                        End If
skip:
                    Next
                End If

            Catch ex As Exception
                MessageBox.Show("Error pulling previous values to populate list box ERR->" + ex.Message)
            End Try




        End If

        If CheckBoxCalloutBOMFindandReplace.Checked = False Then
            Me.ListBoxFindReplaceLinkedPropertyText.Enabled = False
            Me.ListBoxFindReplaceLinkedPropertyText.Items.Clear()
            Me.Button2.Enabled = False
            Me.Button6.Enabled = False
            Me.TextBoxFindReplaceLinkPropertyText.Enabled = False
            Me.CheckBoxRemoveProperties.Checked = False

        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Me.TextBoxFindReplaceLinkPropertyText.Text <> String.Empty Then
            Me.ListBoxFindReplaceLinkedPropertyText.Items.Add(Me.TextBoxFindReplaceLinkPropertyText.Text)

            Dim strPulledContents As String = Me.TextBoxPropsToRemove.Text
            Dim charArray() As String = Me.TextBoxFindReplaceLinkPropertyText.Text.Split(",")
            Dim nloc As Integer = 0

            nloc = InStr(charArray(0), "/", Microsoft.VisualBasic.vbTextCompare)
            If nloc <> 0 Then
                If strPulledContents = "" Then
                    strPulledContents = Trim(charArray(0).ToString.Substring(2, nloc - 3))
                    GoTo Skip
                Else
                    If strPulledContents.Contains(charArray(0).ToString.Substring(2, nloc - 3)) = False Then
                        strPulledContents = Trim(strPulledContents + "," + charArray(0).ToString.Substring(2, nloc - 3))
                        GoTo Skip
                    End If
                End If
            End If

            nloc = InStr(charArray(0), "|", Microsoft.VisualBasic.vbTextCompare)
            If nloc <> 0 Then
                If strPulledContents = "" Then
                    strPulledContents = Trim(charArray(0).ToString.Substring(2, nloc - 3))
                    GoTo Skip
                Else
                    If strPulledContents.Contains(charArray(0).ToString.Substring(2, nloc - 3)) = False Then
                        strPulledContents = Trim(strPulledContents + "," + charArray(0).ToString.Substring(2, nloc - 3))
                        GoTo Skip
                    End If
                End If
            End If

Skip:
            Me.TextBoxPropsToRemove.Text = strPulledContents
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If Me.ListBoxFindReplaceLinkedPropertyText.SelectedItem <> Nothing Then
            Me.ListBoxFindReplaceLinkedPropertyText.Items.RemoveAt(Me.ListBoxFindReplaceLinkedPropertyText.SelectedIndex)
        End If

        Me.TextBoxPropsToRemove.Text = ""


        Dim ii As Integer = 0

        For ii = 0 To Me.ListBoxFindReplaceLinkedPropertyText.Items.Count - 1


            Dim strPulledContents As String = Me.TextBoxPropsToRemove.Text
            Dim charArray() As String = Me.ListBoxFindReplaceLinkedPropertyText.Items(ii).Split(",")
            Dim nloc As Integer = 0

            nloc = InStr(charArray(0), "/", Microsoft.VisualBasic.vbTextCompare)
            If nloc <> 0 Then
                If strPulledContents = "" Then
                    strPulledContents = Trim(charArray(0).ToString.Substring(2, nloc - 3))
                    GoTo skip1
                Else
                    If strPulledContents.Contains(charArray(0).ToString.Substring(2, nloc - 3)) = False Then
                        strPulledContents = Trim(strPulledContents + "," + charArray(0).ToString.Substring(2, nloc - 3))
                        GoTo skip1
                    End If
                End If
            End If

            nloc = InStr(charArray(0), "|", Microsoft.VisualBasic.vbTextCompare)
            If nloc <> 0 Then
                If strPulledContents = "" Then
                    strPulledContents = Trim(charArray(0).ToString.Substring(2, nloc - 3))
                Else
                    If strPulledContents.Contains(charArray(0).ToString.Substring(2, nloc - 3)) = False Then
                        strPulledContents = Trim(strPulledContents + "," + charArray(0).ToString.Substring(2, nloc - 3))
                    End If

                End If
            End If
skip1:
            Me.TextBoxPropsToRemove.Text = strPulledContents


skip:
        Next


        'Next


    End Sub

    Private Sub TextBoxPropsToRemove_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxPropsToRemove.TextChanged

    End Sub

    Private Sub CheckBoxRemoveProperties_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxRemoveProperties.CheckedChanged
        If CheckBoxRemoveProperties.Checked = True Then
            Me.TextBoxPropsToRemove.Enabled = True
        Else
            Me.TextBoxPropsToRemove.Enabled = False
        End If
    End Sub

    Private Sub TextBoxFindReplaceLinkPropertyText_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxFindReplaceLinkPropertyText.TextChanged

    End Sub

    Private Sub ListBoxFindReplaceLinkedPropertyText_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBoxFindReplaceLinkedPropertyText.SelectedIndexChanged

    End Sub

    Private Sub CheckBoxAddHardWare_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBoxAddHardWare.CheckedChanged
        If CheckBoxAddHardWare.Checked = True Then
            Me.TextBoxCheckProperty.Enabled = True
        Else
            Me.TextBoxCheckProperty.Enabled = False
        End If
    End Sub

    Private Sub CBCopyStyles_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CBCopyStyles.CheckedChanged

    End Sub
End Class
