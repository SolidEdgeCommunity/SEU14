Imports System.Runtime.InteropServices
Imports System
Imports System.IO

Public Class Form1
    Public strFilter As String
    Public intFileNameCount As Long
    Public strFileNames As System.Collections.ArrayList
    Public FoundMatchingBGSheet As System.Collections.ArrayList
    Public Filter As String
    Public ObjSEAppType As Type
    Public ObjRevManType As Type
    Public objSEApp As SolidEdgeFramework.Application = Nothing
    Public objRevMan As RevisionManager.Application = Nothing
    Public FSO As Scripting.FileSystemObject
    Public IntNumberOfSheets As Integer
    Public strFileNameNoPath As String
    Public strFileNameNoExtension As String
    Public strPathOnly As String
    Public strSheetName As String
    Public blnDraftFileAttached As Boolean
    Public strSheetSize As String
    Public blnDraftContainsDifferentSizeSheets As Boolean
    Public txtStreamReportStatus As Scripting.TextStream
    Public arrayListOfSheetsToDelete As System.Collections.ArrayList
    'Public NewDraftTemplate As SolidEdgeDraft.DraftDocument = Nothing
    Public objTemplate As SolidEdgeDraft.DraftDocument = Nothing
    Public blnBG_Sheet_Name_Mismatch As Boolean = False


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With FolderBrowserDialog1
            .ShowDialog()
            Me.TxtFolderName.Text = .SelectedPath
            Me.TxtStatusFile.Text = .SelectedPath + "\" + "ReplaceBorderStatusReport.txt"
        End With
        FolderBrowserDialog1.Dispose()

        setfilter()
    End Sub

    Private Sub setfilter()


        Me.ListBox1.Items.Clear()
        Filter = ""

        If Filter <> "" Then
            Filter = Filter & ","
        End If
        Filter = Filter & "*.dft"
        RecursiveDirectoryList(Me.TxtFolderName.Text, False, "*.dft")
        strFilter = Filter


    End Sub



    ' Recursively list all files and subdirectories from the specified source directory

    Private Sub RecursiveDirectoryList(ByVal sourceDir As String, ByVal fRecursive As Boolean, ByVal filter As String)
        Dim sDir As String
        'Dim dDirInfo As IO.DirectoryInfo
        Dim sDirInfo As IO.DirectoryInfo
        Dim sFile As String
        Dim sFileInfo As IO.FileInfo
        'Dim dFileInfo As IO.FileInfo
        ' Add trailing separators to the supplied paths if they don't exist. 
        If Not sourceDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
            sourceDir &= System.IO.Path.DirectorySeparatorChar
        End If

        ' Recursive switch to continue drilling down into directory structure. 
        If fRecursive Then
            ' Get a list of directories from the current parent. 
            For Each sDir In System.IO.Directory.GetDirectories(sourceDir)
                sDirInfo = New System.IO.DirectoryInfo(sDir)
                ' Since we are in recursive mode
                RecursiveDirectoryList(sDirInfo.FullName, fRecursive, filter)
                sDirInfo = Nothing
            Next
        End If

        ' Get the files from the current parent. 
        For Each sFile In System.IO.Directory.GetFiles(sourceDir, filter)
            sFileInfo = New System.IO.FileInfo(sFile)
            Me.ListBox1.Items.Add(sFileInfo.FullName)
            sFileInfo = Nothing
        Next
    End Sub


    ' Recursively travels through a directory structure saving all files into an array.
    Private Sub ReadFileNamesFromDirectory(ByVal sourceDir As String, ByVal fRecursive As Boolean, ByVal filter As String, ByVal strFileNames As ArrayList)
        Dim sDir As String
        'Dim dDirInfo As IO.DirectoryInfo
        Dim sDirInfo As IO.DirectoryInfo
        Dim sFile As String
        Dim sFileInfo As IO.FileInfo
        'Dim dFileInfo As IO.FileInfo
        Dim count As Integer
        ' Add trailing separators to the supplied paths if they don't exist. 
        If Not sourceDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
            sourceDir &= System.IO.Path.DirectorySeparatorChar
        End If

        ' Recursive switch to continue drilling down into directory structure. 
        If fRecursive Then
            ' Get a list of directories from the current parent. 
            For Each sDir In System.IO.Directory.GetDirectories(sourceDir)
                sDirInfo = New System.IO.DirectoryInfo(sDir)
                ' Since we are in recursive mode
                ReadFileNamesFromDirectory(sDirInfo.FullName, fRecursive, filter, strFileNames)
                sDirInfo = Nothing

            Next
        End If

        count = System.IO.Directory.GetFiles(sourceDir, filter).Length
        ' Get the files from the current parent. 
        For Each sFile In System.IO.Directory.GetFiles(sourceDir, filter)
            sFileInfo = New System.IO.FileInfo(sFile)
            intFileNameCount = intFileNameCount + 1
            strFileNames.Add(sFileInfo.FullName)
            sFileInfo = Nothing
        Next
        count = 0
    End Sub

    Private Sub optSelected_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSelected.CheckedChanged
        setfilter()
    End Sub

    Private Sub optAllInDirectory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAllInDirectory.CheckedChanged
        setfilter()
    End Sub

    Private Sub optAllFiles_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAllFiles.CheckedChanged
        setfilter()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim strPathToNotpad As String
        Try
            If FSO Is Nothing Then
                FSO = New Scripting.FileSystemObject
            End If

            strPathToNotpad = FSO.GetSpecialFolder(Scripting.SpecialFolderConst.WindowsFolder).Path + "\notepad.exe"

            Call Shell(strPathToNotpad + " " + Me.TxtStatusFile.Text, AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgBox("Error displaying file", MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        With OpenFileDialog1
            .Filter = "Solid Edge Draft Files (*.dft) | *.dft"
            .ShowDialog()
        End With
        Me.TxtFileContainingNewBorder.Text = OpenFileDialog1.FileName

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        OLEMessageFilter.Register()

        FoundMatchingBGSheet = New ArrayList

        blnDraftFileAttached = False
        Dim i As Long
        Dim intCounter As Long


        'initialize a few things 
        strFileNames = New ArrayList  'initialize array to hold the files to process
        intCounter = 0

        ' Build up list of files to process depending on which option was selected.
        intFileNameCount = 0

        Me.TxtStatus.Text = "Determining list of files to process based on options selected."
        Me.TxtStatus.Refresh()
        If Me.ListBox1.Items.Count > 0 Or optAllFiles.Checked = True Then
            ' Only process the selected files so build up the array so it
            ' contains the currently selected files.
            If optSelected.Checked = True Then
                ' Load the selected filenames into the array.
                For i = 0 To Me.ListBox1.SelectedItems.Count - 1
                    intFileNameCount = intFileNameCount + 1
                    strFileNames.Add(Me.ListBox1.SelectedItems(i))
                Next

            ElseIf optAllInDirectory.Checked = True Then ' Process all the files in the current directory.
                ' Load the selected filenames into the array.
                For i = 0 To Me.ListBox1.Items.Count - 1
                    intFileNameCount = intFileNameCount + 1
                    strFileNames.Add(Me.ListBox1.Items(i))
                Next

            ElseIf optAllFiles.Checked = True Then ' Process all files in the current directory and subdirectories.
                ' Call function to get all files in the current directory and subdirectories.
                Filter = ""
                ReadFileNamesFromDirectory(Me.TxtFolderName.Text, True, "*.dft", strFileNames)

                Me.TxtStatus.Text = ""
                Me.TxtStatus.Refresh()
            End If
        End If

        intFileNameCount = strFileNames.Count
        'Output message that no files to process
        If intFileNameCount = 0 Then 'Check for file to process
            MsgBox("No files to process.  Did you select an option to define the scope by which list of files will be generated?", , "Convert Solid Edge to DWG")
            Exit Sub
        Else 'Files to process
            'do nothing ready to process the specofied files
        End If


        'set global variables

        Try
            If FSO Is Nothing Then
                FSO = New Scripting.FileSystemObject
            End If

            If FSO.FileExists(Me.TxtStatusFile.Text) = True Then
                FSO.DeleteFile(Me.TxtStatusFile.Text)
            End If

            txtStreamReportStatus = FSO.OpenTextFile(Me.TxtStatusFile.Text, Scripting.IOMode.ForAppending, True, Scripting.Tristate.TristateUseDefault)
        Catch ex As Exception
            MsgBox("Error creating text file to contaning import status " + Me.TxtStatusFile.Text, MsgBoxStyle.Critical)
            End
        End Try


        txtStreamReportStatus.WriteLine("Begin Processing.  " + Now.ToLocalTime)
        txtStreamReportStatus.WriteLine("  ")


        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor


        For i = 0 To strFileNames.Count - 1
            txtStreamReportStatus.WriteLine("Processing File: " + strFileNames.Item(i))
            ' Display current status
            Me.TxtStatus.Text = "Processing file " & i + 1 & " of " & intFileNameCount & ": '" & strFileNames.Item(i) & "'"
            Me.TxtStatus.Refresh()
            ' For every XX number of files, close and reopen Solid Edge.  Foor batch jobs with large number of files.  Good idea
            ' to kill SE and restart
            If intCounter Mod CInt(Me.TxtCloseAfter.Text) = 0 Then
                Me.TxtStatus.Text = "Closing and restarting Solid Edge"
                Me.TxtStatus.Refresh()
                If (Not (objSEApp Is Nothing)) Then
                    Try
                        objSEApp.Quit()
                    Catch ex As Exception
                        KillProcess("edge")
                    End Try
                End If
                Garbage_Collect(objSEApp)
                System.Threading.Thread.Sleep(100)
            End If

            Me.TxtStatus.Text = "Processing file " & i + 1 & " of " & intFileNameCount & ": '" & strFileNames.Item(i) & "'"
            Me.TxtStatus.Refresh()

            Try
                ConnectToSolidEdge()
            Catch ex As Exception
                'raise appropriate error message
            End Try



            'check to see if given draft file currently has write access
            If CheckFileAttribute(strFileNames.Item(i), IO.FileAttributes.ReadOnly) Then
                txtStreamReportStatus.WriteLine("Error: File is Read-Only " + strFileNames.Item(i))
                GoTo skip
            End If

            'check to see if given draft file currently has status of released or baselined ... ie un-editable

            Dim Pulled_SE_Status As SolidEdgeFramework.DocumentStatus

            Pulled_SE_Status = GetSEStatus(strFileNames.Item(i))

            If Pulled_SE_Status = SolidEdgeFramework.DocumentStatus.igStatusUnknown Then
                txtStreamReportStatus.WriteLine("Error: Solid Edge internal status is UNKNOWN " + strFileNames.Item(i))
                GoTo skip
            End If

            If Pulled_SE_Status = SolidEdgeFramework.DocumentStatus.igStatusReleased Then
                txtStreamReportStatus.WriteLine("Error: Solid Edge internal status is RELEASED " + strFileNames.Item(i))
                GoTo skip
            End If

            If Pulled_SE_Status = SolidEdgeFramework.DocumentStatus.igStatusObsolete Then
                txtStreamReportStatus.WriteLine("Error: Solid Edge internal status is OBSELETE " + strFileNames.Item(i))
                GoTo skip
            End If

            If Pulled_SE_Status = SolidEdgeFramework.DocumentStatus.igStatusBaselined Then
                txtStreamReportStatus.WriteLine("Error: Solid Edge internal status is BASELINED " + strFileNames.Item(i))
                GoTo skip
            End If

           

            Try
                If Me.TxtFileContainingNewBorder.Text.ToUpper = strFileNames.Item(i).ToString.ToUpper Then
                    GoTo skip
                End If
                objTemplate = objSEApp.Documents.Open(Me.TxtFileContainingNewBorder.Text)
            Catch ex As Exception
                'raise appropriate error message
            End Try

            FoundMatchingBGSheet.Clear()
            'should have write access to the file now go process it!
            ProcessDraftDoc(strFileNames.Item(i))
skip:

            intCounter = intCounter + 1

        Next



        If objSEApp.Documents.Count = 0 Then
            objSEApp.Quit()
        End If
        
        Garbage_Collect(objSEApp)



        Me.TxtStatus.Text = "Finished Processing"
        Me.Refresh()
        txtStreamReportStatus.WriteLine("  ")
        txtStreamReportStatus.WriteLine("Finished Processing.  " + Now.ToLocalTime)
        txtStreamReportStatus.Close()


        System.Windows.Forms.Cursor.Current = Cursors.Default

        OLEMessageFilter.Revoke()
    End Sub


    Public Function ProcessDraftDoc(ByVal filename As String) As Boolean
        Dim objDraftDoc As SolidEdgeDraft.DraftDocument = Nothing


        Try
            ConnectToSolidEdge()
            objSEApp.DoIdle()
            ' '' ''If objSEApp.Documents.Count <> 0 Then 'make sure another document is not already open
            ' '' ''    txtStreamReportStatus.WriteLine("Error: Document left open ... can not continue processing " + filename)
            ' '' ''    ProcessDraftDoc = False
            ' '' ''End If

            Try
                objSEApp.DisplayAlerts = False
                If filename.ToUpper = Me.TxtFileContainingNewBorder.Text.ToUpper Then
                    ProcessDraftDoc = False
                    Exit Function
                End If
                objDraftDoc = objSEApp.Documents.Open(filename)  ' the file where the border needs to be replaced
            Catch ex As Exception
                'MsgBox("Please save the solid edge file.", MsgBoxStyle.OkOnly, "Extract proerties from SE Draft")
                txtStreamReportStatus.WriteLine("Error: opening file " + filename)
                ProcessDraftDoc = False
                Exit Function
            End Try

            objSEApp.DoIdle()

            MakeEachBackGroundSheetActiveAndProcessIt(objDraftDoc)

            If FoundMatchingBGSheet.Contains("FALSE") = True Then
                txtStreamReportStatus.WriteLine("ERROR Mismatch in Background sheet names between the provided new template and : " + filename)
            End If

        Catch ex As Exception

        End Try



        ProcessDraftDoc = False

    End Function

    Public Function MakeEachBackGroundSheetActiveAndProcessIt(ByVal objdoc As SolidEdgeDraft.DraftDocument) As Boolean
        'Dim objsheets As SolidEdgeDraft.Sheets = Nothing
        'Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        'Dim pp As Integer
        'Dim dblsheetHeight As Double
        'Dim dblsheetWidth As Double
        ' Dim objselectset As SolidEdgeFramework.SelectSet = Nothing
        Dim dblTol As Double

        'Dim objBackSheet As SolidEdgeDraft.Sheet = Nothing
        MakeEachBackGroundSheetActiveAndProcessIt = False
        dblTol = 0.00001
        Try    
            For Each objBGSheet As SolidEdgeDraft.Sheet In objdoc.Sections.BackgroundSection.Sheets
                objdoc.Activate()
                objSEApp.DoIdle()
                Dim strBGSheetNameFromDocBeingProcessed As String = objBGSheet.Name
                objBGSheet.Activate()
                objSEApp.DoIdle()
                objdoc.SelectSet.AddAll()
                objdoc.SelectSet.Delete()
                objdoc.SelectSet.RemoveAll()
                objTemplate.Activate()  ' switch to the new template file
                objSEApp.DoIdle()
                For Each objTemplateBGSheet As SolidEdgeDraft.Sheet In objTemplate.Sections.BackgroundSection.Sheets
                    MakeEachBackGroundSheetActiveAndProcessIt = False
                    If objTemplateBGSheet.Name = strBGSheetNameFromDocBeingProcessed Then
                        objTemplateBGSheet.Activate()
                        objSEApp.DoIdle()
                        objTemplate.SelectSet.AddAll()
                        objTemplate.SelectSet.Copy()
                        objTemplate.SelectSet.RemoveAll()
                        Garbage_Collect(objTemplateBGSheet)
                        MakeEachBackGroundSheetActiveAndProcessIt = True
                        FoundMatchingBGSheet.Add("TRUE")
                        Exit For
                    End If

                Next

                If MakeEachBackGroundSheetActiveAndProcessIt = False Then
                    FoundMatchingBGSheet.Add("FALSE")
                    GoTo skiptoHere
                End If

                objdoc.Activate()
                objSEApp.DoIdle()
                objSEApp.ActiveWindow.Paste()
skiptoHere:
                'at this point clipboard still contains the stuff that was pasted.  really need to remove it from the clipboard.
                My.Computer.Clipboard.Clear()
                Garbage_Collect(objBGSheet)

            Next

            objdoc.Sections.WorkingSection.Sheets.Item(1).Activate()
            objSEApp.ActiveWindow.displaybackgroundsheettabs = False

            objdoc.Close(True)
            Garbage_Collect(objdoc)
            objTemplate.Close(False)
            Garbage_Collect(objTemplate)

        Catch ex As Exception
            MakeEachBackGroundSheetActiveAndProcessIt = False
        End Try





    End Function

    Sub Garbage_Collect(ByVal obj As Object)


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
                    KillProcess = -1
                    LocalProcs = Nothing
                    Exit Function
                End Try
            End If
            i += 1
        Next

        If blnProcessTerminated = True Then
            KillProcess = 0
            Exit Function
        End If

        If blnProcessTerminated = False Then
            KillProcess = -2
            Exit Function
        End If

        KillProcess = -1

    End Function



    Public Function CheckFileAttribute(ByVal filename As String, ByVal attribute As IO.FileAttributes) As Boolean
        If IO.File.Exists(filename) Then
            If (IO.File.GetAttributes(filename) And attribute) > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function



    Public Function GetSEStatus(ByVal strFName As String) As Integer
        Dim objPropertySets As SolidEdgeFileProperties.PropertySets = Nothing

        Try
            objPropertySets = New SolidEdgeFileProperties.PropertySets
            Call objPropertySets.Open(strFName, True)
            GetSEStatus = objPropertySets.Item("ExtendedSummaryInformation").item("Status").value
            objPropertySets.Close()
            Garbage_Collect(objPropertySets)
            Exit Function
        Catch ex As Exception
            GetSEStatus = SolidEdgeFramework.DocumentStatus.igStatusUnknown
            Garbage_Collect(objPropertySets)
            Exit Function
        End Try


    End Function

    Public Sub ConnectToSolidEdge()


        'to connect to a running instance of Solid Edge
        Try
           
            objSEApp = Marshal.GetActiveObject("SolidEdge.Application")
            objSEApp.DisplayAlerts = False

            If Me.RBHideSE.Checked = True Then
                objSEApp.Visible = False
            End If
            If Me.RBShowSE.Checked = True Then
                objSEApp.Visible = True
            End If


        Catch ex As System.Exception
            'SE not running then start it
            Try
                Me.TxtStatus.Text = "Starting Solid Edge ..."
                Me.TxtStatus.Refresh()
                objSEApp = Activator.CreateInstance(ObjSEAppType)
                objSEApp.DisplayAlerts = False

                If Me.RBHideSE.Checked = True Then
                    objSEApp.Visible = False
                End If
                If Me.RBShowSE.Checked = True Then
                    objSEApp.Visible = True
                End If
            Catch ex1 As Exception

                If objSEApp Is Nothing Then
                    MsgBox("Could not Start Solid Edge", MsgBoxStyle.OkOnly)
                End If
                Exit Sub
            End Try
        End Try

        If objSEApp Is Nothing Then
            MsgBox("Could not Start Solid Edge", MsgBoxStyle.OkOnly)
            End
        End If



        objSEApp.DisplayAlerts = False


    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ObjSEAppType = Type.GetTypeFromProgID("SolidEdge.Application")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'do necessary garbage collection

        End
    End Sub

    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click
        Me.optSelected.Checked = True
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class
