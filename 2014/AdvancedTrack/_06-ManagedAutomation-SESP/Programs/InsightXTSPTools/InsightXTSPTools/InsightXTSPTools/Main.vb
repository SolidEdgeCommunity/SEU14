Imports System.Runtime.InteropServices
Imports Microsoft.SharePoint

Module Main
    Public lnNumberOfFiles As Long
    Public lngNumberCheckedOut As Long
    Public strStoredLinkedDocSEDocID As String
    Public strStoredLinkedDocURL As String
    Public strCurrentFileBeingProcessed As String
    Public blnIXTModeOn As Boolean = False
    Public strCacheLocation As String = String.Empty
    Public strSQLServerName As String = String.Empty
    Public strUserName As String = String.Empty
    Public strXMLFileName As String = "InsightXTCacheSettings.xml"
    Public strType As String
    Public intNumberOfLinkedItems As Integer
    Public ListOfLinkedItems As String = Nothing
    Public ListOfLinkedItemRevisions As String = Nothing
    Public ListOfItemRevIDs As Object
    Public arrayTCItems As System.Collections.ArrayList
    Public arrayTCRevisions As System.Collections.ArrayList
    Public arrayListOfLinkedFilesAccordingToSE As System.Collections.ArrayList
    Public ArrayAlreadyAddedToTextFile As System.Collections.ArrayList
    Public lngNumberOfFilesToProcess As Long
    Public intSEECCount As Integer
    Public BadBomCtr As Long
    Public txtBadBomFile As String
    Public strCacheFolder As String = String.Empty
    Public intSEDOCSays As Integer
    Public ObjRevMan As RevisionManager.Application
    Public objLinkDocs As RevisionManager.LinkedDocuments
    Public objRevManDoc As RevisionManager.Document
    Public RevManType As Type
    Public arrayUrlOfFilesToCheckIn As ArrayList = Nothing








    Public Sub Main()
        ' Get the type from the Solid Edge ProgID
        objSEType = Type.GetTypeFromProgID("SolidEdge.Application")
        ' Get the type from the Revision Manager ProgID
        objRevManType = Type.GetTypeFromProgID("RevisionManager.Application")

        OleMessageFilter.Register()

        'Add your code here!




        OleMessageFilter.Revoke()

    End Sub



    Public Function DetermineNumberOfFirstLevelLinkedDocuments(ByVal oFilename) As Integer

        Dim nn As Integer
        Dim ctr As Integer
        ctr = 0
        nn = 1
        Dim junk As String



        Try

            If ObjRevMan Is Nothing Then
                ObjRevMan = Activator.CreateInstance(RevManType)
                ObjRevMan.Visible = False
                ObjRevMan.DisplayAlerts = False
                ObjRevMan.ResolveLink = False
            End If
        Catch ex As Exception

            Try
                ObjRevMan = CreateObject("RevisionManager.Application")
                ObjRevMan.Visible = False
                ObjRevMan.DisplayAlerts = False
                ObjRevMan.ResolveLink = False
            Catch ex1 As Exception
                DetermineNumberOfFirstLevelLinkedDocuments = -1
                oReleaseObject(objLinkDocs)
                oReleaseObject(objRevManDoc)
                oReleaseObject(ObjRevMan)
                Exit Function
            End Try
        End Try



        Try

            If System.IO.File.Exists(oFilename) = False Then  ' can not find the file in the SEEC cache
                DetermineNumberOfFirstLevelLinkedDocuments = -100
                oReleaseObject(objLinkDocs)
                oReleaseObject(objRevManDoc)
                'Garbage_Collect(ObjRevMan)
                Exit Function
            End If

            objRevManDoc = ObjRevMan.Open(oFilename)


        Catch ex As Exception
            'could not open the file for some reason...  kil and try it again

            Try
                ObjRevMan.Quit()
                oReleaseObject(objLinkDocs)
                oReleaseObject(objRevManDoc)
                oReleaseObject(ObjRevMan)
            Catch ex2 As Exception

            End Try



            Try

                ObjRevMan = CreateObject("RevisionManager.Application")
                ObjRevMan.Visible = False
                ObjRevMan.DisplayAlerts = False
                ObjRevMan.ResolveLink = False
                objRevManDoc = ObjRevMan.Open(oFilename)
                GoTo skipToOverHere
            Catch ex1 As Exception
                DetermineNumberOfFirstLevelLinkedDocuments = -200
                oReleaseObject(objLinkDocs)
                oReleaseObject(objRevManDoc)
                oReleaseObject(ObjRevMan)
                Exit Function
            End Try


            DetermineNumberOfFirstLevelLinkedDocuments = -2
            oReleaseObject(objLinkDocs)
            oReleaseObject(objRevManDoc)
            oReleaseObject(ObjRevMan)
            Exit Function
        End Try

skipToOverHere:

        Try
            objLinkDocs = objRevManDoc.LinkedDocuments(RevisionManager.LinkTypeConstants.seLinkTypeNormal)
        Catch ex As Exception
            DetermineNumberOfFirstLevelLinkedDocuments = -3
            oReleaseObject(objLinkDocs)
            oReleaseObject(objRevManDoc)
            oReleaseObject(ObjRevMan)
            Exit Function
        End Try


        Try


            'DetermineNumberOfFirstLevelLinkedDocuments = objLinkDocs.Count

            For nn = 1 To objLinkDocs.Count
                'objLinkDocs.Item(nn).occurrences()
                Dim strLinkedFileName As String = objLinkDocs.Item(nn).fullname
                strLinkedFileName = System.IO.Path.GetFileName(strLinkedFileName)

                If arrayListOfLinkedFilesAccordingToSE.Contains(strLinkedFileName.ToUpper) = False Then
                    arrayListOfLinkedFilesAccordingToSE.Add(strLinkedFileName.ToUpper)
                End If
                ctr = ctr + objLinkDocs.Item(nn).occurrences()
            Next
            DetermineNumberOfFirstLevelLinkedDocuments = ctr
            'ObjRevMan.Quit()

            oReleaseObject(objLinkDocs)
            oReleaseObject(objRevManDoc)

            'If Not (ObjRevMan Is Nothing) Then     ***** leave revision manager running
            '    Marshal.ReleaseComObject(ObjRevMan)
            '    ObjRevMan = Nothing
            'End If

            oForceGarbageCollection()

            Exit Function

        Catch ex As System.Exception
            'MsgBox("Error opening assembly in Revision Manager " + fname, MsgBoxStyle.OKOnly)
            DetermineNumberOfFirstLevelLinkedDocuments = -1000
            oReleaseObject(objLinkDocs)
            oReleaseObject(objRevManDoc)
            oReleaseObject(ObjRevMan)
            Exit Function
        End Try




    End Function
End Module
