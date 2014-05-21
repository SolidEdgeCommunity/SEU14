Imports System.Runtime.InteropServices
Imports Microsoft.SharePoint
Imports System.Xml
Imports System.IO

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


    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure

    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" ()
    Public Const RESOURCETYPE_DISK As Long = &H1
    Public Const NO_ERROR As Long = 0









    Public Sub Main()
        ' Get the type from the Solid Edge ProgID
        objSEType = Type.GetTypeFromProgID("SolidEdge.Application")
        ' Get the type from the Revision Manager ProgID
        objRevManType = Type.GetTypeFromProgID("RevisionManager.Application")

        OleMessageFilter.Register()

        'Add your code here!




        OleMessageFilter.Revoke()

    End Sub

    Public Function readContentDBFromXML(oFile As String, strFindValueFor As String) As String

        Dim strReturnString As String = String.Empty





        Try
            Dim txtReader As TextReader = New System.IO.StreamReader(oFile, System.Text.Encoding.UTF8)
            Dim reader As New XmlTextReader(txtReader)


            Do While reader.Read()
                If reader.NodeType = XmlNodeType.Element Then
                    If reader.Name = "ContentDatabase" Then
                        strReturnString = reader.GetAttribute(strFindValueFor)       
                    End If
                End If

            Loop
            reader.Close()

            txtReader.Dispose()
            txtReader.Close()

            oReleaseObject(reader)
            oReleaseObject(txtReader)
        Catch ex As Exception

        End Try



        Return strReturnString

    End Function



    Public Function readNodeFromXML(oFile As String, strnode As String) As String
        Dim m_xmld As XmlDocument
        Dim strReturnString As String = String.Empty
        Dim m_Node As XmlNode = Nothing
        Dim m_ListOfNOdes As XmlNodeList = Nothing
        Dim nn As Integer = 0
        Dim m_PropNode As XmlNode = Nothing
        Dim Pulledstring As String = String.Empty


        Try
            m_xmld = New XmlDocument()
            m_xmld.Load(oFile)
            m_Node = m_xmld.SelectSingleNode("SharePointServer")
            m_ListOfNOdes = m_Node.ChildNodes
            For nn = 0 To m_ListOfNOdes.Count - 1
                Dim strName As String = m_ListOfNOdes.Item(nn).Name
                If strName.ToLower = strnode.ToLower Then
                    Pulledstring = m_ListOfNOdes.Item(nn).InnerText
                    oReleaseObject(m_ListOfNOdes)
                    oReleaseObject(m_Node)
                    oReleaseObject(m_xmld)
                    Return Pulledstring
                End If
            Next

        Catch ex As Exception
            oReleaseObject(m_ListOfNOdes)
            oReleaseObject(m_Node)
            oReleaseObject(m_xmld)
            Return strReturnString
        End Try
        oReleaseObject(m_ListOfNOdes)
        oReleaseObject(m_Node)
        oReleaseObject(m_xmld)
        Return strReturnString

    End Function

    Public Sub MapNetworkDrive(ByVal UncPath As String, ByVal DriveLetter As Char, ByVal Persistent As Boolean, Optional ByVal ConnectionUsername As String = Nothing, Optional ByVal ConnectionPassword As String = Nothing)
        If String.IsNullOrEmpty(UncPath) Then
            Throw New ArgumentException("No UNC path specified", "UncPath")
        End If
        Dim DriveInfo As New NETRESOURCE
        With DriveInfo
            .dwType = RESOURCETYPE_DISK
            .lpLocalName = DriveLetter & ":"
            .lpRemoteName = UncPath
        End With
        Dim flags As UInteger = 0
        If Persistent Then
            flags = &H1
        End If
        Dim Result As UInteger = WNetAddConnection2(DriveInfo, ConnectionPassword, ConnectionUsername, flags)
        If Not Result = NO_ERROR Then
            Throw New System.ComponentModel.Win32Exception(CInt(Result))
        End If
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
