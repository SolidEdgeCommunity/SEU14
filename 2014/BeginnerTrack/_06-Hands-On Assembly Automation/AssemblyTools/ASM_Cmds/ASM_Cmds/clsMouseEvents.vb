Imports SolidEdgeConstants
Imports SolidEdgeFramework
Imports SolidEdgeAssembly

Imports System.Object
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Namespace SolidEdge.ASM_Edgebar_Cmds

Public Enum cmdActions
  NotRunning
  ToggleCsys
  ToggleRefPlanes
  ToggleSketches
End Enum


Public Class clsMouseEvents
  Implements SolidEdgeFramework.ISECommandEvents
  Implements SolidEdgeFramework.ISEMouseEvents
  Implements IDisposable

  Private Command_CP As IConnectionPoint
  Private Command_CP_Cookie As Integer
  Private Mouse_CP As IConnectionPoint
  Private Mouse_CP_Cookie As Integer

  Dim objApp As SolidEdgeFramework.Application
  Dim objDoc As SolidEdgeAssembly.AssemblyDocument

  Dim _theCtrl As ASMEdgebarCtrl
  Dim objCmd As SolidEdgeFramework.Command
  Dim objMouse As SolidEdgeFramework.Mouse

  Dim _cmdAction As cmdActions = cmdActions.NotRunning

  Private _disposed As Boolean = False

  Private _connectionPointDictionary As New Dictionary(Of IConnectionPoint, Integer)()

  Public Sub RunAction(cmdAction As cmdActions, ByRef SEApp As SolidEdgeFramework.Application, ByRef theCtrl As ASMEdgebarCtrl)

      _cmdAction = cmdAction
      _theCtrl = theCtrl

      objApp = SEApp

      objCmd = objApp.CreateCommand(seCmdFlag.seNoDeactivate)
      objDoc = DirectCast(objApp.ActiveDocument, SolidEdgeAssembly.AssemblyDocument)

      objMouse = objCmd.Mouse
      objCmd.Start()

      AddOrRemoveEventHandlers(True, True, True)

      objMouse.InterDocumentLocate = False 'Leave at false since peer component locate is not ready
      objMouse.LocateMode = seLocateModes.seLocateQuickPick
      objMouse.AddToLocateFilter(seLocateFilterConstants.seLocatePart)

      SetPrompt()

  End Sub

  Public Sub EndAction()
      CleanUP()
  End Sub

  Public ReadOnly Property IsRunning As Boolean
      Get
          If _cmdAction = cmdActions.NotRunning Then Return False
          Return True
      End Get
  End Property

  Private Sub SetPrompt()
      Dim tmpStr1 As String = "Select component to toggle "
      Dim tmpStr2 As String = " display on/off, Right click to cancel."

      Select Case _cmdAction
          Case cmdActions.NotRunning
              Return
          Case cmdActions.ToggleCsys
              objApp.StatusBar = tmpStr1 + "Coordinate System" + tmpStr2
          Case cmdActions.ToggleRefPlanes
              objApp.StatusBar = tmpStr1 + "Reference Plane" + tmpStr2
          Case cmdActions.ToggleSketches
              objApp.StatusBar = tmpStr1 + "Sketch display" + tmpStr2
      End Select

  End Sub

#Region "SolidEdgeFramework.ISEMouseEvents"
  Public Sub MouseClick1(sButton As Short, sShift As Short, dX As Double, dY As Double, dZ As Double, pWindowDispatch As Object, lKeyPointType As Integer, pGraphicDispatch As Object) Implements SolidEdgeFramework.ISEMouseEvents.MouseClick
      Select Case sButton

          Case 1 'LMB Click

              If IsNothing(pGraphicDispatch) Then Exit Sub

              Dim TopLevelASM As SolidEdgeAssembly.AssemblyDocument = Nothing

              Dim obj1 As Object = Nothing
              Dim obj2 As Object = Nothing
              Dim objActionObject As Object = Nothing
              'Dim aSubOccurrencesInPath1 As Array
              Dim iBoundSubOccurrencesInPath1 As Integer = 0

              Try
                  If objDoc.InPlaceActivated Then

                      Select Case pGraphicDispatch.Type
                          Case SolidEdgeConstants.ObjectType.igPart, SolidEdgeConstants.ObjectType.igReference
                              'This is an occurrence or sub-occurrence from the active sub-ASM we are IPA'd into.
                              objDoc.GetTopDocumentAndSubOccurrenceOfIPADoc(obj1, obj2)
                              TopLevelASM = CType(obj1, SolidEdgeAssembly.AssemblyDocument)

                              'obj2 is the occurrence we are IPA'd into,
                              If Not obj2.Subassembly Then Exit Sub

                              'Build a reference from the top-level using the select object and the 
                              'Sub-ASM we are IPA'd into.
                              objActionObject = BuildReference(obj2, pGraphicDispatch)

                          Case Else
                              Exit Sub
                      End Select
                  Else
                      'Since we are not IPA we can use the selected item it directly.
                      objActionObject = pGraphicDispatch
                  End If

                  'This routine performs the operation on the selected item.
                  ActionOnSelected(objActionObject)

                  Exit Sub

              Catch ex As Exception
                  MsgBox(ex.Message, , "MouseClick1")
              End Try

          Case 2 'RMB Click
              _theCtrl.StopCommand()
      End Select

  End Sub

  Public Sub MouseDblClick(sButton As Short, sShift As Short, dX As Double, dY As Double, dZ As Double, pWindowDispatch As Object, lKeyPointType As Integer, pGraphicDispatch As Object) Implements SolidEdgeFramework.ISEMouseEvents.MouseDblClick
  End Sub

  Public Sub MouseDown1(sButton As Short, sShift As Short, dX As Double, dY As Double, dZ As Double, pWindowDispatch As Object, lKeyPointType As Integer, pGraphicDispatch As Object) Implements SolidEdgeFramework.ISEMouseEvents.MouseDown
  End Sub

  Public Sub MouseDrag(sButton As Short, sShift As Short, dX As Double, dY As Double, dZ As Double, pWindowDispatch As Object, DragState As Short, lKeyPointType As Integer, pGraphicDispatch As Object) Implements SolidEdgeFramework.ISEMouseEvents.MouseDrag
  End Sub

  Public Sub MouseMove1(sButton As Short, sShift As Short, dX As Double, dY As Double, dZ As Double, pWindowDispatch As Object, lKeyPointType As Integer, pGraphicDispatch As Object) Implements SolidEdgeFramework.ISEMouseEvents.MouseMove
  End Sub

  Public Sub MouseUp1(sButton As Short, sShift As Short, dX As Double, dY As Double, dZ As Double, pWindowDispatch As Object, lKeyPointType As Integer, pGraphicDispatch As Object) Implements SolidEdgeFramework.ISEMouseEvents.MouseUp
  End Sub
#End Region

#Region "SolidEdgeFramework.ISECommandEvents"
  Public Sub Activate1() Implements SolidEdgeFramework.ISECommandEvents.Activate

  End Sub

  Public Sub Deactivate1() Implements SolidEdgeFramework.ISECommandEvents.Deactivate

  End Sub

  Public Sub Idle(lCount As Integer, ByRef pbMore As Boolean) Implements SolidEdgeFramework.ISECommandEvents.Idle

  End Sub

  Public Sub KeyDown1(ByRef KeyCode As Short, Shift As Short) Implements SolidEdgeFramework.ISECommandEvents.KeyDown
  End Sub

  Public Sub KeyPress1(ByRef KeyAscii As Short) Implements SolidEdgeFramework.ISECommandEvents.KeyPress
  End Sub

  Public Sub KeyUp1(ByRef KeyCode As Short, Shift As Short) Implements SolidEdgeFramework.ISECommandEvents.KeyUp
  End Sub

  Public Sub Terminate() Implements SolidEdgeFramework.ISECommandEvents.Terminate
      _theCtrl.StopCommand()
  End Sub
#End Region

#Region "UTILITIES"

  Private Sub AddOrRemoveEventHandlers(ByVal Add As Boolean, ByVal CommandEvents As Boolean, ByVal MouseEvents As Boolean)

      Dim i As Type
      Dim EventGuid As Guid

      Try
          If CommandEvents Then
              i = GetType(SolidEdgeFramework.ISECommandEvents)
              EventGuid = i.GUID

              ConnectToEvents(Command_CP, Add, objCmd, Command_CP_Cookie, EventGuid)

          End If

          If MouseEvents Then
              i = GetType(SolidEdgeFramework.ISEMouseEvents)
              EventGuid = i.GUID

              ConnectToEvents(Mouse_CP, Add, objMouse, Mouse_CP_Cookie, EventGuid)
          End If

      Catch ex As Exception

      End Try
  End Sub

  Private Sub ConnectToEvents(ByRef CP As IConnectionPoint, ByVal Add As Boolean, ByVal obj As Object, ByRef Cookie As Integer, ByVal EventGuid As Guid)
      Dim CPC As IConnectionPointContainer

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
  End Sub

  Public Sub ReleaseComObject(ByRef obj As Object)
      If Not obj Is Nothing Then
          ' Call FinalReleaseComObject. This call means that this tool MUST NOT try to reference the object again, even from another variable.
          System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
          obj = Nothing
      End If
  End Sub

  Private Sub CleanUP()
      ' If user closed the form set Done on the command object to terminate the command in Solid Edge. Also disconnect from the events. 
      ' If the user terminated the command via the Solid Edge UI, say by starting another command or hitting the escape key etc, I
      ' don't really need to set Done to True but I still want to disconnect from events. As long as I disconnect from the events
      ' before I set Done to true things work out. But if I don't, setting Done to true can cause Edge to fire an event (Terminate) to
      ' the command. Also the Idle event is called (a lot) and Solid Edge and this app can both lockup in a classic deadlock
      ' situation.
      AddOrRemoveEventHandlers(False, True, True)

      objApp = Nothing
      ReleaseComObject(objMouse)

      objCmd.Done = True
      ReleaseComObject(objCmd)

      objDoc = Nothing
  End Sub

  Private Sub UpdateDisplay()
      '	Dim pWindows As SolidEdgeFramework.Windows
      Dim pWin As SolidEdgeFramework.Window
      Dim pView As SolidEdgeFramework.View

      pWin = objApp.ActiveWindow
      pView = pWin.View
      pView.Update()

  End Sub

#End Region

#Region "IConnectionPoint helpers"

  Private Sub HookEvents(ByVal eventSource As Object, ByVal eventGuid As Guid)
      Dim container As IConnectionPointContainer = Nothing
      Dim connectionPoint As IConnectionPoint = Nothing
      Dim cookie As Integer = 0

      container = DirectCast(eventSource, IConnectionPointContainer)
      container.FindConnectionPoint(eventGuid, connectionPoint)

      If connectionPoint IsNot Nothing Then
          connectionPoint.Advise(Me, cookie)
          _connectionPointDictionary.Add(connectionPoint, cookie)
      End If
  End Sub

  Private Sub UnhookAllEvents()
      Dim enumerator As Dictionary(Of IConnectionPoint, Integer).Enumerator = _connectionPointDictionary.GetEnumerator()
      Do While enumerator.MoveNext()
          enumerator.Current.Key.Unadvise(enumerator.Current.Value)
      Loop

      _connectionPointDictionary.Clear()
  End Sub

#End Region

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
      If Not Me.disposedValue Then
          If disposing Then
              ' TODO: dispose managed state (managed objects).
              EndAction()
          End If

          ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
          ' TODO: set large fields to null.
      End If
      Me.disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
  End Sub
#End Region

Private Sub ActionOnSelected(objActionObject As Object)
'This sub toggles the display based on the running action.

If _cmdAction = cmdActions.NotRunning Then Return

Dim pReference As SolidEdgeFramework.Reference = Nothing
Dim pSubOccurrence As SolidEdgeAssembly.SubOccurrence = Nothing
Dim pOccurrence As SolidEdgeAssembly.Occurrence = Nothing
Dim SelectedObject As Object = Nothing

SelectedObject = objActionObject

Select Case SelectedObject.Type

  Case SolidEdgeConstants.ObjectType.igPart
      'We are not IPA and this is a top-level occurrence.
      pOccurrence = CType(SelectedObject, SolidEdgeAssembly.Occurrence)

      Select Case _cmdAction
          Case cmdActions.ToggleCsys
              If pOccurrence.DisplayCoordinateSystems Then
                  pOccurrence.DisplayCoordinateSystems = False
              Else
                  pOccurrence.DisplayCoordinateSystems = True
              End If
          Case cmdActions.ToggleRefPlanes
              If pOccurrence.DisplayReferencePlanes Then
                  pOccurrence.DisplayReferencePlanes = False
              Else
                  pOccurrence.DisplayReferencePlanes = True
              End If
          Case cmdActions.ToggleSketches
              If pOccurrence.DisplaySketches Then
                  pOccurrence.DisplaySketches = False
              Else
                  pOccurrence.DisplaySketches = True
              End If
      End Select

      UpdateDisplay()

  Case SolidEdgeConstants.ObjectType.igReference, SolidEdgeConstants.ObjectType.igSubOccurrence

      If SelectedObject.Type = SolidEdgeConstants.ObjectType.igReference Then
          'Have to get it's subOccurrence from the top level ASM.

          pReference = CType(SelectedObject, SolidEdgeFramework.Reference)

          Dim objTopOccurrence1 As SolidEdgeAssembly.Occurrence = Nothing
          Dim iSubOccurrencesInPath1 As Integer
          Dim aSubOccurrencesInPath1 As Array = Array.CreateInstance(GetType(System.Object), 0)
          Dim iBoundSubOccurrencesInPath1 As Integer

          objTopOccurrence1 = pReference.Parent
          pReference.GetOccurrencesInPath(objTopOccurrence1, iSubOccurrencesInPath1, iBoundSubOccurrencesInPath1, aSubOccurrencesInPath1)

          If iSubOccurrencesInPath1 > 0 Then
              Dim obj As Object = aSubOccurrencesInPath1(iSubOccurrencesInPath1 - 1)
              pSubOccurrence = CType(obj, SolidEdgeAssembly.SubOccurrence)
          Else
              Exit Sub
          End If

      ElseIf SelectedObject.Type = SolidEdgeConstants.ObjectType.igSubOccurrence Then
          'Can use it directly.
          pSubOccurrence = CType(SelectedObject, SolidEdgeAssembly.SubOccurrence)
      Else
        Exit Sub
      End If

      Select Case _cmdAction
          Case cmdActions.ToggleCsys
              If pSubOccurrence.DisplayCoordinateSystems Then
                  pSubOccurrence.DisplayCoordinateSystems = False
              Else
                  pSubOccurrence.DisplayCoordinateSystems = True
              End If
          Case cmdActions.ToggleRefPlanes
              If pSubOccurrence.DisplayReferencePlanes Then
                  pSubOccurrence.DisplayReferencePlanes = False
              Else
                  pSubOccurrence.DisplayReferencePlanes = True
              End If
          Case cmdActions.ToggleSketches

             If pSubOccurrence.DisplaySketches Then
                  pSubOccurrence.DisplaySketches = False
              Else
                  pSubOccurrence.DisplaySketches = True
              End If
              pOccurrence = pSubOccurrence.ThisAsOccurrence
              If pOccurrence.DisplaySketches Then
                  pOccurrence.DisplaySketches = False
              Else
                  pOccurrence.DisplaySketches = True
              End If

      End Select

      UpdateDisplay()

  End Select

  End Sub

#Region "ASM Selection and Reference Processing"
  Public Function BuildReference(objIPAOccOrSubOcc As Object, pGraphicDispatch As Object) As Object
      Dim objActionObject As Object = Nothing
      Dim pOcc As SolidEdgeAssembly.Occurrence = Nothing
      Dim pIPASubOcc As SolidEdgeAssembly.SubOccurrence = Nothing
      Dim objAsmDoc As SolidEdgeAssembly.AssemblyDocument

      If IsSubOccurrence(objIPAOccOrSubOcc) Then
          pOcc = objIPAOccOrSubOcc.ThisAsOccurrence()
          objAsmDoc = pOcc.Parent().Parent()
          pGraphicDispatch = CType(objAsmDoc.CreateReference2(pOcc, pGraphicDispatch), SolidEdgeFramework.Reference)

          objIPAOccOrSubOcc = objIPAOccOrSubOcc.Parent()

          'Recurse to this function to get a reference based on the top-level ASM
          'Basically trying to get the occurrence as if we were in the top-level ASM
          objActionObject = BuildReference(objIPAOccOrSubOcc, pGraphicDispatch)

      ElseIf IsOccurrence(objIPAOccOrSubOcc) Then
          objAsmDoc = objIPAOccOrSubOcc.Parent().Parent()
          objActionObject = CType(objAsmDoc.CreateReference2(objIPAOccOrSubOcc, pGraphicDispatch), SolidEdgeFramework.Reference)
      End If

      BuildReference = objActionObject
  End Function

  Public Function IsSubOccurrence(objIPAOccOrSubOcc As Object) As Boolean
      Dim pIPASubOcc As SubOccurrence
      Try
          pIPASubOcc = CType(objIPAOccOrSubOcc, SolidEdgeAssembly.SubOccurrence)
          IsSubOccurrence = True
      Catch ex As Exception
          IsSubOccurrence = False
      End Try
  End Function

  Public Function IsOccurrence(objIPAOccOrSubOcc As Object) As Boolean
      Dim pIPASubOcc As Occurrence
      Try
          pIPASubOcc = CType(objIPAOccOrSubOcc, SolidEdgeAssembly.Occurrence)
          IsOccurrence = True
      Catch ex As Exception
          IsOccurrence = False
      End Try
  End Function

  Public Function IsReference(objIsRef As Object) As Boolean
      Dim pRef As Reference
      Try
          pRef = CType(objIsRef, SolidEdgeFramework.Reference)
          IsReference = True
      Catch ex As Exception
          IsReference = False
      End Try
  End Function
#End Region

End Class

End Namespace