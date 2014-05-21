Imports System.Runtime.InteropServices
Module CommonFunctions
    Public objSEApp As SolidEdgeFramework.Application = Nothing
    Public objSEType As Type = Nothing
    Public objRevManApp As RevisionManager.Application = Nothing
    Public objRevManType As Type = Nothing
    Public Const InToMeter As Double = 0.0254
    Public Const PI As Double = 3.14159265359






    Public Function oConnectToSolidEdge(ByVal blnAppVisibility As Boolean, ByVal blnDisplayAlerts As Boolean) As Boolean

        'to connect to a running instance of Solid Edge
        Try
            '("Word.Application")
            '("Excel.Application")
            '("RevisionManager.Application")
            objSEApp = Marshal.GetActiveObject("SolidEdge.Application")
            objSEApp.DisplayAlerts = blnDisplayAlerts
            objSEApp.Visible = blnAppVisibility
            Return True


        Catch ex As System.Exception

            'SE not running then start it
            Try
                objSEApp = Activator.CreateInstance(objSEType)
                objSEApp.DisplayAlerts = blnDisplayAlerts
                objSEApp.Visible = blnAppVisibility
                Return True


            Catch ex1 As Exception
                If objSEApp Is Nothing Then
                    MessageBox.Show("Error starting the Solid Edge Application. The error is " + ex1.Message, "Error starting Solid Edge", MessageBoxButtons.OK)
                End If
                Return False

            End Try
        End Try

        If objSEApp Is Nothing Then
            MessageBox.Show("Could not start or connect to the Solid Edge Application", "Error starting or connecting to Solid Edge", MessageBoxButtons.OK)
        End If

        Return False
    End Function

    Public Function oConnectToRevisionManager(ByVal blnAppVisibility As Boolean, ByVal blnDisplayAlerts As Boolean) As Boolean

        'to connect to a running instance of Revision Manager
        Try
            '("Word.Application")
            '("Excel.Application")
            '("RevisionManager.Application")
            objRevManApp = Marshal.GetActiveObject("RevisionManager.Application")
            objRevManApp.DisplayAlerts = blnDisplayAlerts
            objRevManApp.Visible = blnAppVisibility
            Return True
          
        Catch ex As System.Exception

            'SE not running then start it
            Try
                objRevManApp = Activator.CreateInstance(objRevManType)
                objRevManApp.DisplayAlerts = blnDisplayAlerts
                objRevManApp.Visible = blnAppVisibility
                Return True


            Catch ex1 As Exception
                If objRevManApp Is Nothing Then
                    MessageBox.Show("Error starting the Revision Manager Application. The error is " + ex1.Message, "Error starting Revision Manager", MessageBoxButtons.OK)
                End If
                Return False

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
            End If
            If System.IO.Path.HasExtension(strFileNameOnly) Then
                'it has an extension now check to see if it is a valid SE one
                Dim strExtension As String = System.IO.Path.GetExtension(strFileNameOnly)
                For i As Integer = 0 To validFileTypes.Length - 1
                    If strExtension.ToLower = "." + validFileTypes(i).ToLower Then
                        Return True
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
        Dim Proc As Process
        Dim i As Integer

        LocalProcs = System.Diagnostics.Process.GetProcesses
        For Each Proc In LocalProcs
            If UCase(Proc.ProcessName) = UCase(Name) Then
                Try
                    Proc.Kill()
                    oReleaseObject(Proc)
                    oReleaseObject(LocalProcs)
                    oForceGarbageCollection()
                    Return 0
                Catch ex As System.Exception
                    oReleaseObject(Proc)
                    oReleaseObject(LocalProcs)
                    oForceGarbageCollection()
                    Return 1
                    Exit Function
                End Try
            End If
            i += 1
        Next
        oReleaseObject(Proc)
        oReleaseObject(LocalProcs)
        oForceGarbageCollection()
        Return 1


    End Function

    Sub oReleaseObject(ByVal obj As Object)
        Try
            If Not (obj Is Nothing) Then
                'this should only be used when programming applications that run in their own process space
                Marshal.FinalReleaseComObject(obj)
            End If
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub
    Public Sub oForceGarbageCollection()
        Try
            GC.Collect(GC.MaxGeneration())
            GC.WaitForPendingFinalizers()
            GC.Collect(GC.MaxGeneration())
        Catch ex As Exception

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
        Dim objPropertySets As SolidEdgeFileProperties.PropertySets = Nothing

        Try
            objPropertySets = New SolidEdgeFileProperties.PropertySets
            Call objPropertySets.Open(strFName, True)
            Return objPropertySets.Item("ExtendedSummaryInformation").item("Status").value
            objPropertySets.Close()
            oReleaseObject(objPropertySets)
            oForceGarbageCollection()
            Exit Function
        Catch ex As Exception
            Return SolidEdgeFramework.DocumentStatus.igStatusUnknown
            oReleaseObject(objPropertySets)
            oForceGarbageCollection()
            Exit Function
        End Try


    End Function

    Public Function oGetSolidEdgePath() As String

        Dim strSEPath As String = String.Empty
        Dim install As SEInstallDataLib.SEInstallData = Nothing
        Try
            install = New SEInstallDataLib.SEInstallData
            strSEPath = install.GetInstalledPath
        Catch ex As Exception
            Return Nothing
        Finally
            oReleaseObject(install)
            oForceGarbageCollection()
        End Try

        Return strSEPath

    End Function

    Public Function oGetSolidEdgeVersion() As String
        Try
            Dim install As New SEInstallDataLib.SEInstallData
            Dim strSEVersion As String = String.Empty

            strSEVersion = install.GetVersion.ToString
            oReleaseObject(install)
            oForceGarbageCollection()

            Return strSEVersion
        Catch ex As Exception
            Return Nothing

        End Try

    End Function

    Public Function oCreateNewSEDocument(ByVal oApp As SolidEdgeFramework.Application, ByVal strDocType As String) As SolidEdgeFramework.SolidEdgeDocument
        Dim objSEDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing

        Try
            Select Case strDocType.ToUpper
                Case "PART"
                    objSEDoc = oApp.Documents.Add("SolidEdge.PartDocument")
                    Return objSEDoc
                Case "SHEETMETAL"
                    objSEDoc = oApp.Documents.Add("SolidEdge.SheetMetalDocument")
                    Return objSEDoc
                Case "ASSEMBLY"
                    objSEDoc = oApp.Documents.Add("SolidEdge.AssemblyDocument")
                    Return objSEDoc
                Case "DRAFT"
                    objSEDoc = oApp.Documents.Add("SolidEdge.DraftDocument")
                    Return objSEDoc
            End Select

        Catch ex As Exception
            MessageBox.Show("Error ccreating the document.  error is " + ex.Message)
            Return Nothing
        End Try


    End Function

    Public Function IsDocumentMetric(ByVal objDoc As Object) As Boolean

        Dim objUOM As SolidEdgeFramework.UnitsOfMeasure = Nothing
        Dim objUnit As SolidEdgeFramework.UnitOfMeasure = Nothing
        Dim cUnits As SolidEdgeFramework.UnitTypeConstants
        Dim Units As SolidEdgeConstants.UnitOfMeasureLengthReadoutConstants
        Dim Found As Boolean
        Dim Result As Boolean
        Dim Index As Integer

        Result = True

        ' get the UnitsOfMeasure object from the document
        objUOM = objDoc.UnitsOfMeasure
        Index = 1
        While (Not Found And Index < objUOM.Count)
            objUnit = objUOM.Item(Index)
            cUnits = objUnit.Type
            If cUnits = SolidEdgeConstants.UnitTypeConstants.igUnitDistance Then
                Found = True
                Units = objUnit.Units
            End If
            Index = Index + 1
        End While

        If Found Then
            If Units = SolidEdgeConstants.UnitOfMeasureLengthReadoutConstants.seLengthCentimeter Or _
                                  Units = SolidEdgeConstants.UnitOfMeasureLengthReadoutConstants.seLengthKilometer Or _
                                  Units = SolidEdgeConstants.UnitOfMeasureLengthReadoutConstants.seLengthMeter Or _
                                  Units = SolidEdgeConstants.UnitOfMeasureLengthReadoutConstants.seLengthMillimeter Or _
                                  Units = SolidEdgeConstants.UnitOfMeasureLengthReadoutConstants.seLengthNanometer Then
                Result = True
            Else
                Result = False
            End If
        End If


        oReleaseObject(objUnit)
        oReleaseObject(objUOM)
        oForceGarbageCollection()

        IsDocumentMetric = Result

    End Function

    'Returns a number converted from inches to meters
    Public Function ItoM(ByVal InchValue As Double) As Double
        ItoM = InchValue * InToMeter
    End Function

    'Returns a number converted from meters to inches
    Public Function MtoI(ByVal MeterValue As Double) As Double
        MtoI = MeterValue / InToMeter
    End Function

    'Returns a number converted from millimeters to meters
    Public Function MMtoM(ByVal MilliMeterValue As Double) As Double
        MMtoM = MilliMeterValue / 1000.0#
    End Function

    'Returns a number converted from millimeters to meters
    Public Function MtoMM(ByVal MeterValue As Double) As Double
        MtoMM = MeterValue * 1000.0#
    End Function

    'Returns a number converted from millimeters to meters
    Public Function DtoR(ByVal DegreeValue As Double) As Double
        DtoR = DegreeValue * PI / 180.0#
    End Function



End Module
