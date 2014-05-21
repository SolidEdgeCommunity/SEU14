Imports System.Runtime.InteropServices
Module CommonFunctions
    Public objSEApp As SolidEdgeFramework.Application = Nothing
    Public objSEType As Type = Nothing
    Public objRevManApp As RevisionManager.Application = Nothing
    Public objRevManType As Type = Nothing


    Public Sub releaseObject(ByRef obj As Object)

        Try
            If (obj IsNot Nothing) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally

        End Try
    End Sub
    Public Sub ForceGarbageCollection()
        Try
            GC.Collect(GC.MaxGeneration())
            GC.WaitForPendingFinalizers()
            GC.Collect(GC.MaxGeneration())
        Catch ex As Exception

        End Try
    End Sub


    Public Function oConnectToSolidEdge(ByVal blnAppVisibility As Boolean, ByVal blnDisplayAlerts As Boolean) As Boolean
        'to connect to a running instance of Solid Edge
        Try
            objSEApp = Marshal.GetActiveObject("SolidEdge.Application")
            objSEApp.DisplayAlerts = blnDisplayAlerts
            objSEApp.Visible = blnAppVisibility
            ' check this.... return might automatically exit function!
            Return True
            Exit Function

        Catch ex As System.Exception

            'SE not running then start it
            Try
                objSEApp = Activator.CreateInstance(objSEType)
                objSEApp.DisplayAlerts = blnDisplayAlerts
                objSEApp.Visible = blnAppVisibility
                Return True
                Exit Function

            Catch ex1 As Exception
                If objSEApp Is Nothing Then
                    MessageBox.Show("Error starting the Solid Edge Application. The error is " + ex1.Message, _
                                    "Error starting Solid Edge", MessageBoxButtons.OK)
                End If
                Return False
                Exit Function
            End Try
        End Try

        If objSEApp Is Nothing Then
            MessageBox.Show("Could not start or connect to the Solid Edge Application", "Error starting or connecting to Solid Edge", _
                            MessageBoxButtons.OK)
        End If

        Return False
    End Function

    Public Function oConnectToRevisionManager(ByVal blnAppVisibility As Boolean, ByVal blnDisplayAlerts As Boolean) As Boolean

        'to connect to a running instance of Revision Manager
        Try
            objRevManApp = Marshal.GetActiveObject("RevisionManager.Application")
            objRevManApp.DisplayAlerts = blnDisplayAlerts
            objRevManApp.Visible = blnAppVisibility
            Return True
            Exit Function

        Catch ex As System.Exception

            'SE not running then start it
            Try
                objRevManApp = Activator.CreateInstance(objRevManType)
                objRevManApp.DisplayAlerts = blnDisplayAlerts
                objRevManApp.Visible = blnAppVisibility
                Return True
                Exit Function

            Catch ex1 As Exception
                If objRevManApp Is Nothing Then
                    MessageBox.Show("Error starting the Revision Manager Application. The error is " + ex1.Message, "Error starting Revision Manager", MessageBoxButtons.OK)
                End If
                Return False
                Exit Function
            End Try
        End Try

        If objRevManApp Is Nothing Then
            MessageBox.Show("Could not start or connect to the Solid Revision Manager", "Error starting or connecting to Revision Manager", MessageBoxButtons.OK)
        End If

        Return False
    End Function

    Public Function oIsValidFileName(ByVal fn As String) As Boolean
        Try
            If System.IO.File.Exists(fn) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeFile(ByVal Filename As String) As Boolean
        Try
            Dim validFileTypes As String() = {"par", "psm", "asm", "dft", "pwd"}
            Dim strFileNameOnly As String = String.Empty

            strFileNameOnly = System.IO.Path.GetFileName(Filename)
            If strFileNameOnly = String.Empty Then
                Return False
                Exit Function
            End If
            If System.IO.Path.HasExtension(strFileNameOnly) Then
                'it has an extension now check to see if it is a valid SE one
                Dim strExtension As String = System.IO.Path.GetExtension(strFileNameOnly)
                For i As Integer = 0 To validFileTypes.Length - 1
                    If strExtension.ToLower = "." & validFileTypes(i).ToLower Then
                        Return True
                        Exit Function
                    End If
                Next
                Return False
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgePartFile(ByVal Filename As String) As Boolean
        Try
            Dim validFileTypes As String() = {"par"}
            Dim strFileNameOnly As String = String.Empty

            strFileNameOnly = System.IO.Path.GetFileName(Filename)
            If strFileNameOnly = String.Empty Then
                Return False
                Exit Function
            End If
            If System.IO.Path.HasExtension(strFileNameOnly) Then
                'it has an extension now check to see if it is a valid SE one
                Dim strExtension As String = System.IO.Path.GetExtension(strFileNameOnly)
                For i As Integer = 0 To validFileTypes.Length - 1
                    If strExtension.ToLower = "." & validFileTypes(i).ToLower Then
                        Return True
                        Exit Function
                    End If
                Next
                Return False
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeSheetMetalFile(ByVal Filename As String) As Boolean
        Try
            Dim validFileTypes As String() = {"psm"}
            Dim strFileNameOnly As String = String.Empty

            strFileNameOnly = System.IO.Path.GetFileName(Filename)
            If strFileNameOnly = String.Empty Then
                Return False
                Exit Function
            End If
            If System.IO.Path.HasExtension(strFileNameOnly) Then
                'it has an extension now check to see if it is a valid SE one
                Dim strExtension As String = System.IO.Path.GetExtension(strFileNameOnly)
                For i As Integer = 0 To validFileTypes.Length - 1
                    If strExtension.ToLower = "." & validFileTypes(i).ToLower Then
                        Return True
                        Exit Function
                    End If
                Next
                Return False
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeDraftFile(ByVal Filename As String) As Boolean
        Try
            Dim validFileTypes As String() = {"dft"}
            Dim strFileNameOnly As String = String.Empty

            strFileNameOnly = System.IO.Path.GetFileName(Filename)
            If strFileNameOnly = String.Empty Then
                Return False
                Exit Function
            End If
            If System.IO.Path.HasExtension(strFileNameOnly) Then
                'it has an extension now check to see if it is a valid SE one
                Dim strExtension As String = System.IO.Path.GetExtension(strFileNameOnly)
                For i As Integer = 0 To validFileTypes.Length - 1
                    If strExtension.ToLower = "." & validFileTypes(i).ToLower Then
                        Return True
                        Exit Function
                    End If
                Next
                Return False
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeAssemblylFile(ByVal Filename As String) As Boolean
        Try
            Dim validFileTypes As String() = {"asm"}
            Dim strFileNameOnly As String = String.Empty

            strFileNameOnly = System.IO.Path.GetFileName(Filename)
            If strFileNameOnly = String.Empty Then
                Return False
                Exit Function
            End If
            If System.IO.Path.HasExtension(strFileNameOnly) Then
                'it has an extension now check to see if it is a valid SE one
                Dim strExtension As String = System.IO.Path.GetExtension(strFileNameOnly)
                For i As Integer = 0 To validFileTypes.Length - 1
                    If strExtension.ToLower = "." & validFileTypes(i).ToLower Then
                        Return True
                        Exit Function
                    End If
                Next
                Return False
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'pass in the full path including filename or just the filename
    Public Function oIsSolidEdgeWeldmentFile(ByVal Filename As String) As Boolean
        Try
            Dim validFileTypes As String() = {"pwd"}
            Dim strFileNameOnly As String = String.Empty

            strFileNameOnly = System.IO.Path.GetFileName(Filename)
            If strFileNameOnly = String.Empty Then
                Return False
                Exit Function
            End If
            If System.IO.Path.HasExtension(strFileNameOnly) Then
                'it has an extension now check to see if it is a valid SE one
                Dim strExtension As String = System.IO.Path.GetExtension(strFileNameOnly)
                For i As Integer = 0 To validFileTypes.Length - 1
                    If strExtension.ToLower = "." & validFileTypes(i).ToLower Then
                        Return True
                        Exit Function
                    End If
                Next
                Return False
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

   ' Append a backslash (or any character) at the end of a path
    ' if it isn't there already
    Function oAddBackslash(ByVal Path As String) As String
        If Path.EndsWith("\") Then
            Return Path
        Else
            Return Path + "\"
        End If
    End Function

    Public Function oKillProcess(ByVal Name As String) As Long

        Dim LocalProcs As Process()
        Dim Proc As Process = Nothing
        Dim i As Integer

        LocalProcs = System.Diagnostics.Process.GetProcesses
        For Each Proc In LocalProcs
            If UCase(Proc.ProcessName) = UCase(Name) Then
                Try
                    Proc.Kill()
                    oRelease_Object(Proc)
                    oRelease_Object(LocalProcs)
                    ForceGarbageCollection()
                    Return 0
                Catch ex As System.Exception
                    oRelease_Object(Proc)
                    oRelease_Object(LocalProcs)
                    ForceGarbageCollection()
                    Return 1
                    Exit Function
                End Try
            End If
            i += 1
        Next
        oRelease_Object(Proc)
        oRelease_Object(LocalProcs)
        ForceGarbageCollection()
        Return 1


    End Function

    Sub oRelease_Object(ByVal obj As Object)
        Try
            If Not (obj Is Nothing) Then
                ' note the API below only be used when programming ADD-ins
                'this should only be used when programming applications that run in their own process space
                Marshal.FinalReleaseComObject(obj)
            End If

        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    Public Function oGetPathOfFilename(ByVal oFileName As String) As String

        Try
            Return System.IO.Path.GetFullPath(oFileName)
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function oGetFileName(ByVal oFileName As String) As String

        Try
            Return System.IO.Path.GetFileName(oFileName)
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function oGetFileExtension(ByVal oFileName As String) As String

        Try
            Return System.IO.Path.GetExtension(oFileName)
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function oGetFileNameWithoutExtension(ByVal oFileName As String) As String

        Try
            Return System.IO.Path.GetFileNameWithoutExtension(oFileName)
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function oGetFolderName(ByVal oFileName As String) As String

        Try
            Return System.IO.Path.GetDirectoryName(oFileName)
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function oReadIniFile(ByVal oFilename As String, ByVal strLineContaining As String) As String
        Try
            Dim arrData As String() = System.IO.File.ReadAllLines(oFilename)
            Dim strData As String

            For Each strData In arrData
                If strData = "" Then
                    Continue For
                End If
                If InStr(strData.ToUpper, "#", Microsoft.VisualBasic.CompareMethod.Text) <> 0 Then
                    GoTo skip
                End If

                If InStr(strData.ToUpper, strLineContaining.ToUpper, Microsoft.VisualBasic.CompareMethod.Text) <> 0 Then
                    Dim arrStrDataSplit As String() = strData.Split("=")
                    Return arrStrDataSplit(1)
                    Exit Function
                End If
skip:
            Next

        Catch ex As Exception
            Return Nothing
        End Try


        Return Nothing
    End Function

    Public Function oGetThisApplicationPath() As String
        Return My.Application.Info.DirectoryPath
    End Function

    Public Function oCheckFileAttribute(ByVal filename As String, ByVal attribute As IO.FileAttributes) As Boolean
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

    Public Function oGetSEStatus(ByVal strFName As String) As SolidEdgeFramework.DocumentStatus
        Dim objPropertySets As SolidEdgePropAuto.PropertySets = Nothing

        Try
            objPropertySets = New SolidEdgePropAuto.PropertySets
            Call objPropertySets.Open(strFName, True)
            Return objPropertySets.Item("ExtendedSummaryInformation").item("Status").value
            objPropertySets.Close()
            oRelease_Object(objPropertySets)
            ForceGarbageCollection()
            Exit Function
        Catch ex As Exception
            Return SolidEdgeFramework.DocumentStatus.igStatusUnknown
            oRelease_Object(objPropertySets)
            ForceGarbageCollection()
            Exit Function
        End Try


    End Function

    Public Function oGetSolidEdgePath() As String
        Try
            Dim install As New SolidEdgeInstallData.SEInstallData
            Dim strSEPath As String = String.Empty

            strSEPath = install.GetInstalledPath
            oRelease_Object(install)
            ForceGarbageCollection()

            Return strSEPath
        Catch ex As Exception
            Return Nothing

        End Try

    End Function

    Public Function oGetSolidEdgeVersion() As String
        Try
            Dim install As New SolidEdgeInstallData.SEInstallData
            Dim strSEVersion As String = String.Empty

            strSEVersion = install.GetVersion.ToString
            oRelease_Object(install)
            ForceGarbageCollection()

            Return strSEVersion
        Catch ex As Exception
            Return Nothing

        End Try

    End Function

End Module
