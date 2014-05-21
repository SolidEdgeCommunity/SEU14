Imports System.Runtime.InteropServices


Class MainWindow

    Dim oApp As Object = Nothing
    Dim oSolidEdge As SolidEdge.Framework.Interop.Application = Nothing

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbTargetFolder.Text = My.Settings.CopyPath
        AttachToSolidEdge()
    End Sub

    Private Sub EnumerateDocuments(asmDoc As SolidEdge.Assembly.Interop.AssemblyDocument)
        On Error Resume Next
        Dim nOccs As Integer = asmDoc.Occurrences.Count
        For Each occ As SolidEdge.Assembly.Interop.Occurrence In asmDoc.Occurrences
            If Not occ.Visible Then
                If Not cbCopyHiddenFiles.IsChecked Then
                    Continue For
                End If
            End If

            Dim strFullName As String = occ.OccurrenceFileName
            If Not lbFileNames.Items.Contains(strFullName) Then
                lbFileNames.Items.Add(strFullName)
            End If

            Dim oSourceDoc As Object = occ.OccurrenceDocument
            If Not oSourceDoc Is Nothing Then
                Dim oDoc As SolidEdge.Framework.Interop.SolidEdgeDocument = oSourceDoc
                If oDoc.Type = SolidEdge.Framework.Interop.DocumentTypeConstants.igAssemblyDocument Then
                    Dim subAsmDoc As SolidEdge.Assembly.Interop.AssemblyDocument = oSourceDoc
                    EnumerateDocuments(oSourceDoc)
                End If
            End If
        Next
    End Sub

    Private Sub btnAttach_Click(sender As Object, e As RoutedEventArgs) Handles btnAttach.Click
        AttachToSolidEdge()
    End Sub

    Private Sub AttachToSolidEdge()
        Dim oldCursor As System.Windows.Input.Cursor = Me.Cursor
        Cursor = Cursors.Wait
        lbFileNames.Items.Clear()
        oApp = Marshal.GetActiveObject("SolidEdge.Application")
        oSolidEdge = oApp
        If oSolidEdge.ActiveDocumentType = SolidEdge.Framework.Interop.DocumentTypeConstants.igAssemblyDocument Then
            Dim oDoc As Object = oSolidEdge.ActiveDocument
            Dim asmDoc As SolidEdge.Assembly.Interop.AssemblyDocument = oDoc
            Dim strFullName As String = asmDoc.FullName
            If Not lbFileNames.Items.Contains(strFullName) Then
                lbFileNames.Items.Add(strFullName)
            End If
            EnumerateDocuments(asmDoc)
        End If
        Cursor = oldCursor
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowse.Click
        Dim browser As New System.Windows.Forms.FolderBrowserDialog()
        browser.SelectedPath = My.Settings.CopyPath
        Dim result As Forms.DialogResult = browser.ShowDialog()
        If (result = Forms.DialogResult.OK) Then
            tbTargetFolder.Text = browser.SelectedPath
            My.Settings.CopyPath = browser.SelectedPath
            My.Settings.Save()
        End If
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As RoutedEventArgs) Handles btnCopy.Click
        Dim di As New System.IO.DirectoryInfo(tbTargetFolder.Text)
        If Not di.Exists Then
            System.Windows.Forms.MessageBox.Show("Target folder does not exist: " & tbTargetFolder.Text)
            Exit Sub
        End If

        If lbFileNames.Items.IsEmpty Then
            AttachToSolidEdge()
        End If

        Dim n As Integer = 1
        Dim copies As String() = New String(lbFileNames.Items.Count) {}

        For Each source As String In lbFileNames.Items
            Dim target As String = System.IO.Path.Combine(tbTargetFolder.Text, System.IO.Path.GetFileName(source))
            Try
                System.IO.File.Copy(source, target, True)
                copies(n) = target
                n = n + 1
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show("Error copying " & source & " to " & target & ": " & ex.ToString())
            End Try
        Next

        lbFileNames.Items.Clear()
        For Each target As String In copies
            lbFileNames.Items.Add(target)
        Next

        My.Settings.CopyPath = tbTargetFolder.Text
        My.Settings.Save()

    End Sub

    Private Sub cbCopyHiddenFiles_Click(sender As Object, e As RoutedEventArgs) Handles cbCopyHiddenFiles.Click
        AttachToSolidEdge()
    End Sub
End Class
