Imports Microsoft.VisualBasic
Imports Microsoft.Win32
Imports System
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms
Imports SolidEdgeFramework
Imports System.Reflection

'Developers Note: .NET has added a FinalReleaseComObject API. The newer API will release the
' actual Solid Edge object immediately so one doesn't have to wait for GC to collect the RCW, which
' is when the call to release the Edge object actually occurs if one calls ReleaseComObject.
'
' BUT BEWARE - FinalReleaseComObject will not only affect one .NET client, but any other .NET client
' that may happen to be using the same Edge object! So only call FinalReleaseComObject on objects for 
' which the addin knows only it should have a reference. When is that? Probably never since any
' automation client can access any object in Edge. That includes the "addin" object passed to an
' addin during connection.

'Developers Note: If you decide to run with the .NET 4 runtime, I advise that you go to the references
' this add-in imports of any solid edge type library and in the properties, set "embed interop types"
' to "False". The property for embedded interop types only shows up if the runtime is 4 (or later).
' Embedding interop types can allow you to not deploy the interop assemblies created from
' a type library. However this has been shown to cause problems. For example using a non dispatch
' based event set will not be possible. If the problem was limited to event sets that is easily overcome
' as one can use, for example, DISEApplicationEvents instead of ISEApplicationEvents. But alas we have
' found that other API calls to edge can also fail. Usually an exception message will contain some
' reference to "embed", "embedded" or "interop" or "type" as a clue that the embedding of interop types
' caused the problem. So if you do embed interop types, make sure you execute every line of code in
' testing. If you encounter a problem, first thing to do is change the property on the reference that
' contains the API that failed so that embedding is not used. Then try the call again.

'Developers Note: The template from which this add-in was created referenced the Solid Edge COM
' type library, framewrk.tlb. This allows the project to build without any changes (so I hope) on
' a machine that has solid edge installed. However the solid edge type library cannot be versioned
' because of this. The reason is that the reference added for a COM type library will include the
' version of the type library. If a user does not have that version on the machine, the add-in will
' fail to resolve the reference and compilation will fail. So Edge does not version the type libraries.
' I strongly recommend that at some point in your development you use the .NET type lib import tool,
' tlbimp.exe, to generate your own interop assemblies from any Solid Edge type libraries you need so
' that you can create unique interop assemblies for your add-in. Develope, test with and deploy those
' interop assemblies with your add-in assembly. The Solid Edge SDK tools directory actually has a bat
' file that generates the interop assemblies give a path to the directory where the type libraries
' reside, an output directory for the generated interop assemblies, and a tag that should be unique.
' For a unique tag, one could use the "short guid", which is the first 8 characters found in the
' GuidAttribute given the add-in below.
'
' Why generate unique interop assemblies? Unfortunately the .NET assembly loader will only load
' one interop assembly with a given signature. So if two add-ins exist with the same interop
' assembly, the first one loaded into edge will have its interop assembly loaded. Any add-in loaded
' after that will find the .NET loader ignored the interop assembly sitting with the add-in assembly.
' This can pose a problem if any API in the assemblies do not match. One will not find this problem
' until the line of code that refererenced such an API is executed. That will happen on some customer
' machine and will not reproduce on another machine unless the other add-in (or macro) is present
' there too. Solid Edge development strives to not change APIs since .NET is not forgiving. However
' APIs have been changed in the past as with VB 6 (or any non .NET clients) can easily handle an
' API that has been modified by adding optional arguments. Unfortunately some .NET add-ins were
' created during the time frame when Edge would change an API by adding a new optional parameter.
' So interop assemblies that are not compatible with ones created with later versions of edge do
' exist. This is rare but it has happened before and is almost impossible to track down. So avoid
' the problem and eventually use the tool to generate your own interops and remove the one the
' wizard added and add a reference to your own.

<System.Runtime.InteropServices.GuidAttribute("8b30aeae-f64c-417c-8cdb-22cef8ec24cf")> _
<System.Runtime.InteropServices.ProgId("CheckDimensionRounding.CheckDimensionRounding")> _
<System.Runtime.InteropServices.ComVisible(True)> Public Class Addin

    Implements SolidEdgeFramework.ISolidEdgeAddIn
    Implements SolidEdgeFramework.ISEApplicationEvents
    Implements SolidEdgeFramework.ISEFileUIEvents

    Implements System.IDisposable

    ' I will keep a copy of the Edge application interface
    Private pApplication As SolidEdgeFramework.Application

    Private Application_CP As IConnectionPoint
    Private Application_CP_Cookie As Integer

    Private FileUI_CP As IConnectionPoint
    Private FileUI_CP_Cookie As Integer

    Dim pAddin As SolidEdgeFramework.AddIn

#Region "SolidEdgeAddInInterface"

    Private Sub ISolidEdgeAddIn_OnConnection(ByVal Application As Object, ByVal ConnectMode As SolidEdgeFramework.SeConnectMode, ByVal AddInInstance As SolidEdgeFramework.AddIn) Implements SolidEdgeFramework.ISolidEdgeAddIn.OnConnection

        pAddin = GetUniqueRCW(AddInInstance)
        pApplication = GetUniqueRCW(Application)

        Application_CP_Cookie = -1
        FileUI_CP_Cookie = -1

        ' TODO: If you want to handle application events and/or FileUI events, uncomment out the call to AddOrRemoveEventHandlers.
        '       Set the variables passed to AddOrRemoveEventHandlers to true for the events you actually want to handle.

        ' add event handlers. Handlers are removed in the Dispose method of the commands object.
        Dim bAddAppEvents As Boolean
        Dim bAddFileUIEvents As Boolean
        '
        bAddAppEvents = True
        bAddFileUIEvents = False
        AddOrRemoveEventHandlers(True, bAddAppEvents, bAddFileUIEvents)

        ' The GUI version should be incremented anytime changes to the UI occur such as adding
        ' new commands to a command bar, removing commands from the add-in etc. Edge will detect
        ' the change and automatically purge the system of any saved command bar data, user assigned
        ' accelerators and other data saved relating to the add-in.
        AddInInstance.GuiVersion = 1
        ' Let the add-in be seen by the user when the add-in manager runs.
        AddInInstance.Visible = True

        Dim Description As String
        Description = GetResourceString(104)
        If Description.Length = 0 Then
            Description = "VB add-in to check dimensions for pre V17 rounding behavior."
        End If

        AddInInstance.Description = Description

    End Sub

    Private Sub ISolidEdgeAddIn_OnConnectToEnvironment(ByVal EnvCatID As String, ByVal pEnvironmentDispatch As Object, ByVal bFirstTime As Boolean) Implements SolidEdgeFramework.ISolidEdgeAddIn.OnConnectToEnvironment

        If Not pEnvironmentDispatch Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pEnvironmentDispatch)
        End If

    End Sub

    Private Sub ISolidEdgeAddIn_OnDisconnection(ByVal DisconnectMode As SolidEdgeFramework.SeDisconnectMode) Implements SolidEdgeFramework.ISolidEdgeAddIn.OnDisconnection
        On Error Resume Next
    End Sub
#End Region

#Region "REGISTRATION"

    ' Code to register/unregister the additional add-in data was derived from Jason Newell (jasonnewell.net)

    ' When Regasm is run on this project, either through IDE or command window, these functions will be called.

    ' Setup the required Solid Edge registry values for an addin that were not automatically
    ' added due to the guid, progid and com visible attributes.
    <ComRegisterFunctionAttribute()> _
    Public Shared Sub RegisterFunction(ByVal t As Type)

        Dim Key As RegistryKey = Registry.ClassesRoot.CreateSubKey("CLSID\{" + t.GUID.ToString() + "}")

        If Not (Key Is Nothing) Then
            ' Tell Edge to automatically connect to the add-in.
            Key.SetValue("AutoConnect", 1)
            ' Set the description
            Key.SetValue("409", "Addin To Check deimension rounding")
            ' Set the summary
            Key.SetValue("Summary", "VB .NET add-in to Check Pre-V17 dimensions in draft files")
            ' Add the Microsoft standard "Implemented Categories" subkey and add ISolidEdgeAddIn as
            ' an implemented category. This is what allows Solid Edge to use the Windows registry
            ' APIs to find an addin registered on the machine.
            Key.CreateSubKey("Implemented Categories\" & CATID_SolidEdgeAddIn)
            ' Set the environment categories to indicate what environments the add-in should
            ' be connected to.
            'Key.CreateSubKey("Environment Categories\" & CATID_SEApplication)
            'Key.CreateSubKey("Environment Categories\" & CATID_SEPart)
            'Key.CreateSubKey("Environment Categories\" & CATID_SEAssembly)
            Key.CreateSubKey("Environment Categories\" & CATID_SEDraft)
            'Key.CreateSubKey("Environment Categories\" & CATID_SESketch)

            Key.Close()
        End If
    End Sub

    ' Remove any registry values specifically added above.
    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub UnregisterFunction(ByVal t As Type)
        Registry.ClassesRoot.DeleteSubKeyTree("CLSID\{" + t.GUID.ToString() + "}")
    End Sub
#End Region

#Region "IDisposable Support "

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free managed resources when explicitly called
                AddOrRemoveEventHandlers(False, True, True)

                If Not pAddin Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pAddin)
                    pAddin = Nothing
                End If

                If Not pApplication Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pApplication)
                    pApplication = Nothing
                End If
            End If

            ' TODO: free shared unmanaged resources
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

#Region "Application Events"
    ' Note that the event methods below will call ReleaseComObject for the passed in RCWs if they are not stored (and if stored by the add-in,
    ' eventually they should be released via that call). With later versions of .NET there is a "FinalReleaseComObject". Don't use that call.
    ' Using that call can actually cause the RCW to be given its final release event if another .NET add-in happens to be holding onto the RCW!
    ' Unfortunately before I realized this, I put out this VB .NET sample full of those calls. No problem as long as no other .NET add-in is
    ' running. But alas, that is not always the case. I jumped on using that API when .NET added it because I found that the actual COM object
    ' inside Solid Edge got a release call when FinalReleaseComObject was called. That avoided problems with calls to the COM object after Edge
    ' unmapped the object from memory (such as a document object when the document was closed) for which all the garbage collection calls in Edge
    ' were trying to address. Note that none of the events here (or code in this add-in) attempt to run garbage collection. Leave that up to Edge
    ' since there can be multiple .NET add-in running and we don't want to bog the system down by having each one run GC on its own. Note that
    ' Edge also handles multiple .NET runtimes loaded due to multiple add-in running in different .NET runtimes (e.g, the 2.0 and 4.0 runtime)
    ' by running GC in each.

    Private Sub AfterActiveDocumentChange(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterActiveDocumentChange
        On Error Resume Next

        If Not theDocument Is Nothing Then
           

            System.Runtime.InteropServices.Marshal.ReleaseComObject(theDocument)
        End If
    End Sub

    Private Sub BeforeDocumentClose(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentClose
        ' This release is for the addref the .NET runtime did when the doc was passed to this event.
        If Not theDocument Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theDocument)
        End If
    End Sub
    Private Sub AfterEnvironmentActivate(ByVal theEnvironment As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterEnvironmentActivate
        On Error Resume Next

        If Not theEnvironment Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theEnvironment)
        End If
    End Sub
    Private Sub AfterWindowActivate(ByVal theWindow As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterWindowActivate
        On Error Resume Next

        If Not theWindow Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theWindow)
        End If
    End Sub

    Private Sub BeforeWindowDeactivate(ByVal theWindow As Object) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeWindowDeactivate
        On Error Resume Next

        If Not theWindow Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theWindow)
        End If
    End Sub
    Private Sub BeforeEnvironmentDeactivate(ByVal theEnvironment As Object) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeEnvironmentDeactivate
        On Error Resume Next

        If Not theEnvironment Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theEnvironment)
        End If
    End Sub
    Private Sub AfterNewDocumentOpen(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterNewDocumentOpen
        On Error Resume Next

        If Not theDocument Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theDocument)
        End If
    End Sub
    Private Sub AfterDocumentOpen(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterDocumentOpen
        'On Error Resume Next

        If Not theDocument Is Nothing Then

            'check if the new active document is a Pre V17 draft file
            Dim strLastSavedVersion As String = String.Empty
            Dim objDraftDoc As SolidEdgeDraft.DraftDocument = Nothing
            Dim arrayOfTerms As String()
            Dim blnNeedToProcessThisDraftFile As Boolean = False


            Try
                objDraftDoc = TryCast(theDocument, SolidEdgeDraft.DraftDocument)

                If Not objDraftDoc Is Nothing Then
                    arrayofBadDimensions = New ArrayList

                    strLastSavedVersion = objDraftDoc.LastSavedVersion
                    Try
                        arrayOfTerms = strLastSavedVersion.Split(".")
                    Catch ex As Exception
                        MessageBox.Show("Error while determining the last saved version of the draft file. error is " + ex.Message)
                        GoTo wrapup
                    End Try

                    If arrayOfTerms(0) = String.Empty Then
                        MessageBox.Show("Could not determine the last saved Major version of the draft file.")
                        GoTo wrapup
                    End If

                    Try
                        Dim intMajorSEVersion As Integer = 0
                        intMajorSEVersion = CInt(arrayOfTerms(0))
                        If intMajorSEVersion < 17 Then
                            blnNeedToProcessThisDraftFile = True
                        Else
                            'GoTo wrapup
                            blnNeedToProcessThisDraftFile = False
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Error while determining the last saved Major version of the draft file. error is " + ex.Message)
                        GoTo wrapup
                    End Try


                    'got to this point so must be a draft file with lasted saved version previous to V17!!!
                    'need to process this drafts dimensions to check for the condition!

                    If oConnectToSolidEdge(True, True) = True Then
                        objSEApp.StatusBar = "Checking draft file for dimensions with pre-V17 rounding behavior"
                    End If

                    '@@@@@@@@  this is set to true ONLY for debug purposes!!!!  comment out the following line when this is delivered.
                    'blnNeedToProcessThisDraftFile = True

                    If blnNeedToProcessThisDraftFile = True Then
                        ProcessDraftDoc(objDraftDoc)
                    End If


                    If arrayofBadDimensions.Count > 0 Then
                        If oConnectToSolidEdge(True, True) = True Then
                            objSEApp.StatusBar = "Found " + (arrayofBadDimensions.Count - 1).ToString + " dimensions where the displayed value might be different in V17 and later SE versions."
                        End If

                    Else
                        If oConnectToSolidEdge(True, True) = True Then
                            objSEApp.StatusBar = "Did not find any effected dimensions."
                        End If
                    End If



                Else
                    MessageBox.Show("Could not determine if the document being opened is a draft file")
                    GoTo wrapup
                End If
            Catch ex As Exception
                MessageBox.Show("Error while determining if the document being opened is a draft file. error is " + ex.Message)
                GoTo wrapup
            End Try

            

wrapup:

            oReleaseObject(objDraftDoc)
            oReleaseObject(theDocument)
            oReleaseObject(objSEApp)

            If blnNeedToProcessThisDraftFile = True Then
                If arrayofBadDimensions.Count = 0 Then
                    MessageBox.Show("This draft file has not been saved since V17 and there were no dimension rounding issues discovered")
                Else
                    MessageBox.Show("This draft file has not been saved since V17 and there were " + (arrayofBadDimensions.Count).ToString + " dimensions rounding issues discovered with possible issues!  Each dimension will be identified by a diamond character!")
                End If
            End If


        End If
    End Sub

    Private Sub AfterDocumentPrint(ByVal theDocument As Object, ByVal hDC As Integer, ByRef ModelToDC As Double, ByRef Rect As Integer) Implements SolidEdgeFramework.ISEApplicationEvents.AfterDocumentPrint
        On Error Resume Next

        If Not theDocument Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theDocument)
        End If
    End Sub
    Private Sub AfterDocumentSave(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterDocumentSave
        On Error Resume Next

        If Not theDocument Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theDocument)
        End If
    End Sub
    Private Sub BeforeDocumentSave(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentSave
        On Error Resume Next

        If Not theDocument Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theDocument)
        End If
    End Sub
    Private Sub AfterNewWindow(ByVal theWindow As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterNewWindow
        On Error Resume Next

        If Not theWindow Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theWindow)
        End If
    End Sub
    Private Sub BeforeDocumentPrint(ByVal theDocument As Object, ByVal hDC As Integer, ByRef ModelToDC As Double, ByRef Rect As Integer) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentPrint
        On Error Resume Next

        If Not theDocument Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(theDocument)
        End If
    End Sub

    Private Sub AfterCommandRun(ByVal theCommandID As Integer) Implements SolidEdgeFramework.ISEApplicationEvents.AfterCommandRun

    End Sub

    Private Sub BeforeCommandRun(ByVal theCommandID As Integer) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeCommandRun

    End Sub

    Private Sub BeforeQuit() Implements SolidEdgeFramework.ISEApplicationEvents.BeforeQuit

    End Sub
#End Region

#Region "Application File UI Events"

    '   FileUIEvents
    '   If the add-in throws the NotImplementedException, Edge gets the COM return code E_NOTIMPL. Edge sees that code and moves on to the
    '   next event listener. If the add-in does not throw that exception, .NET will send Edge the COM return code S_OK. Edge then assumes
    '   the event was handled and will not call out to any other listener. Generally there should only be one listener for File UI events.
    '   Obviously Edge cannot have multiple add-ins showing the user file dialogs and having each return a file to open. If the add-in
    '   shows a file UI and the user cancels the operation, do not throw the exception (so Edge sees S_OK). In that case, simply return
    '   no filename. If all listeners return the exception, Edge will then show its own file UI.
    Private Sub OnCreateInPlacePartUI(ByRef Filename As String, ByRef AppendToTitle As String, ByRef Template As String) Implements SolidEdgeFramework.ISEFileUIEvents.OnCreateInPlacePartUI

        Throw New System.NotImplementedException

    End Sub

    Private Sub OnFileNewUI(ByRef Filename As String, ByRef AppendToTitle As String) Implements SolidEdgeFramework.ISEFileUIEvents.OnFileNewUI

        Throw New System.NotImplementedException

    End Sub

    Private Sub OnFileOpenUI(ByRef Filename As String, ByRef AppendToTitle As String) Implements SolidEdgeFramework.ISEFileUIEvents.OnFileOpenUI

        Throw New System.NotImplementedException

    End Sub

    Private Sub OnFileSaveAsImageUI(ByRef Filename As String, ByRef AppendToTitle As String, ByRef Width As Integer, ByRef Height As Integer, ByRef ImageQuality As SolidEdgeFramework.SeImageQualityType) Implements SolidEdgeFramework.ISEFileUIEvents.OnFileSaveAsImageUI

        Throw New System.NotImplementedException

    End Sub

    Private Sub OnFileSaveAsUI(ByRef Filename As String, ByRef AppendToTitle As String) Implements SolidEdgeFramework.ISEFileUIEvents.OnFileSaveAsUI

        Throw New System.NotImplementedException

    End Sub

    Private Sub OnPlacePartUI(ByRef Filename As String, ByRef AppendToTitle As String) Implements SolidEdgeFramework.ISEFileUIEvents.OnPlacePartUI

        Throw New System.NotImplementedException

    End Sub

#End Region

#Region "Utilities"
    Private Sub AddOrRemoveEventHandlers(ByVal Add As Boolean, ByVal AppEvents As Boolean, ByVal FileUIEvents As Boolean)

        ' For events I have chosen to have the class implement the interface as opposed to the AddHandler/RemoveHandler
        ' .NET paradigm. There are two reasons I have done this. First, when a COM object such as the solid edge
        ' application object has its event set connected to, the object connecting has to supply an interface
        ' that contains event handlers for every event in the interface. So for example if .NET AddHandler is called
        ' for the BeforeCommandRun event, .NET connects an interface to the application object that handles each event.
        ' Only when BeforeCommandRun is called by the application object does the caller to AddHandler get its event
        ' handler called. .NET "stubs" out all the remaining handlers in the interface and simply returns to edge.
        ' If another event handler is added via AddHandler, .NET will connect an entirely different interface to
        ' edge. Hence if ten handlers are added, Edge fires every event ten times to .NET for every event fired.
        ' By using the "Implements" paradigm, .NET only connects the specific interface to Edge. This reduces the
        ' overhead in Edge for firing events.
        '
        ' Actually the above is a bit misleading when I say .NET "stubs" out all the remaining handlers in the interface
        ' and simply returns to edge. If the event that is stubbed out has any COM object passed to .NET from Edge,
        ' .NET still creates the RCW, which is sometimes called a "ghost RCW" since the .NET client programmer never
        ' "sees" (encounters) the RCW since no code was written. By implementing the interface, one is forced to write
        ' minimum code for each event and that means the .net programmer has direct knowledge that the RCW(s) are indeed
        ' being created and hence can call ReleaseComObject.
        '
        ' The second reason to use "Implements" is more subtle and is related to how Solid Edge handles file UI events.
        ' Since the .NET runtime actually connects a new event interface to Edge each time AddHandler is called, 
        ' if the user adds more than one handler to the FileUI event source, when Edge fires events there appears to be 
        ' multiple listeners to Edge and only one returns E_NOTIMPL via throwing the not implemented exception using this 
        ' line of code:

        ' Throw New System.NotImplementedException

        ' The other stubs (unseen by the .net programmer) return S_OK. When S_OK is sent to Edge, Edge thinks 
        ' the user "canceled" the listener's own file UI since no filename is returned. Thus Edge will not show its own
        ' file processing dialog.

        ' This means that if AddHandler is used for file UI events, the .NET programmer can find that everything works fine
        ' if there is only one call to AddHandler for the UI events but can find a problem when AddHandler is called for a
        ' second file UI event.

        ' Whew! That was a lot to say. But now let the show begin.

        Dim i As Type
        Dim EventGuid As Guid

        If AppEvents Then

            i = GetType(SolidEdgeFramework.ISEApplicationEvents)
            EventGuid = i.GUID

            ConnectToEvents(Application_CP, Add, pApplication, Application_CP_Cookie, EventGuid)
        End If

        If FileUIEvents Then

            i = GetType(SolidEdgeFramework.ISEFileUIEvents)
            EventGuid = i.GUID

            ConnectToEvents(FileUI_CP, Add, pApplication, FileUI_CP_Cookie, EventGuid)

        End If
    End Sub

    Private Sub ConnectToEvents(ByRef CP As IConnectionPoint, ByVal Add As Boolean, ByVal obj As Object, ByRef Cookie As Integer, ByVal EventGuid As Guid)
        Dim CPC As IConnectionPointContainer

        Try
            CPC = obj

            If Not CPC Is Nothing Then

                If Add Then
                    CPC.FindConnectionPoint(EventGuid, CP)
                    If Not CP Is Nothing Then
                        CP.Advise(Me, Cookie)
                    End If
                Else
                    If Not CP Is Nothing Then
                        If Not Cookie = -1 Then
                            CP.Unadvise(Cookie)
                            Cookie = -1
                        End If
                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

End Class
