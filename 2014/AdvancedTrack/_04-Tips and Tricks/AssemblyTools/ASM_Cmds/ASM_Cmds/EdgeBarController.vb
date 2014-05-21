Imports SolidEdge.CommonUI
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text

Namespace SolidEdge.ASM_Edgebar_Cmds
	''' <summary>
	''' Controller class for ISolidEdgeBarEx.
	''' </summary>
	Public Class EdgeBarController
		Implements SolidEdgeFramework.ISEAddInEdgeBarEvents, IDisposable

		Private _disposed As Boolean = False
		Private _addInEx As SolidEdgeFramework.ISEAddInEx
		Private _edgeBar As SolidEdgeFramework.ISolidEdgeBarEx
		Private Shared _resourceAssembly As System.Reflection.Assembly
		Private _connectionPointDictionary As New Dictionary(Of IConnectionPoint, Integer)()
		Private _edgeBarPageDictionary As New Dictionary(Of IntPtr, EdgeBarPage)()

'INSTANT VB NOTE: The parameter myAddIn was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
		Public Sub New(ByVal myAddIn_Renamed As SolidEdgeFramework.ISEAddInEx, ByVal resourceAssembly As System.Reflection.Assembly)
			_addInEx = myAddIn_Renamed
			_edgeBar = DirectCast(_addInEx, SolidEdgeFramework.ISolidEdgeBarEx)
			_resourceAssembly = resourceAssembly

			HookEvents(_addInEx, GetType(SolidEdgeFramework.ISEAddInEdgeBarEvents).GUID)
		End Sub

		Protected Overrides Sub Finalize()
			Dispose(False)
		End Sub

#Region "SolidEdgeFramework.ISEAddInEdgeBarEvents"

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISEAddInEdgeBarEvents.AddPage event.
		''' </summary>
    Public Sub AddPage(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEAddInEdgeBarEvents.AddPage

      If theDocument.Type = SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument Then
        'See if we have already create an Edgebar for the doc.
        'We create an edgebar for each ASM doc.
        If NeedtoCreatePage(theDocument) Then
          AddPage(theDocument, New ASMEdgebarCtrl())
        End If
      End If

    End Sub

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISEAddInEdgeBarEvents.IsPageDisplayable event.
		''' </summary>
		Public Sub IsPageDisplayable(ByVal theDocument As Object, ByVal EnvironmentCatID As String, <System.Runtime.InteropServices.Out()> ByRef vbIsPageDisplayable As Boolean) Implements SolidEdgeFramework.ISEAddInEdgeBarEvents.IsPageDisplayable

			' We use the IUnknown pointer of the document as the dictionary key.
			Dim pDocument As IntPtr = Marshal.GetIUnknownForObject(theDocument)
			Marshal.Release(pDocument)

			' If we have an EdgeBarPage, return the EdgeBarControl.IsPageDisplayable property.
			If _edgeBarPageDictionary.ContainsKey(pDocument) Then
				vbIsPageDisplayable = _edgeBarPageDictionary(pDocument).SEEdgeBarControl.IsPageDisplayable
			Else
				' Default to true;
				vbIsPageDisplayable = True
			End If
		End Sub

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISEAddInEdgeBarEvents.RemovePage event.
		''' </summary>
		Public Sub RemovePage(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEAddInEdgeBarEvents.RemovePage
			' We use the IUnknown pointer of the document as the dictionary key.
			Dim pDocument As IntPtr = Marshal.GetIUnknownForObject(theDocument)
			Marshal.Release(pDocument)

			' If we have an EdgeBarPage for the document, remove it.
			If _edgeBarPageDictionary.ContainsKey(pDocument) Then
				RemovePage(pDocument)
      End If

		End Sub

Public Function NeedtoCreatePage(ByVal theDocument As Object) As Boolean
      ' We use the IUnknown pointer of the document as the dictionary key.
      Dim pDocument As IntPtr = Marshal.GetIUnknownForObject(theDocument)
      Marshal.Release(pDocument)

      ' If we have an EdgeBarPage for the document, remove it.
      If _edgeBarPageDictionary.ContainsKey(pDocument) Then
        'Do not need to create it.
        Return False
      End If
    'Need to create the page for the doc
    Return True

End Function

#End Region

#Region "EdgeBarController methods"

		Private Function AddPage(ByVal theDocument As Object, ByVal edgeBarControl As EdgeBarControl) As EdgeBarPage
			Dim edgeBarPage As EdgeBarPage = Nothing
			Dim hWndPage As IntPtr = IntPtr.Zero

			' We use the IUnknown pointer of the document as the dictionary key.
			Dim pDocument As IntPtr = Marshal.GetIUnknownForObject(theDocument)
			Marshal.Release(pDocument)

			' Only add a new EdgeBarPage if one hasn't already been added.
			If Not _edgeBarPageDictionary.ContainsKey(pDocument) Then
				' If ResourceAssembly is null, default to the currently executing assembly.
				If _resourceAssembly Is Nothing Then
					_resourceAssembly = System.Reflection.Assembly.GetExecutingAssembly()
				End If

				hWndPage = New IntPtr(_edgeBar.AddPageEx(theDocument, _resourceAssembly.Location, edgeBarControl.BitmapID, edgeBarControl.ToolTip, 2))

				' ISolidEdgeBarEx.AddPage() may return null.
				If Not hWndPage.Equals(IntPtr.Zero) Then
					edgeBarPage = New EdgeBarPage(hWndPage, theDocument, edgeBarControl)
					_edgeBarPageDictionary.Add(pDocument, edgeBarPage)
				Else
					' AddPageEx() failed to dispose the passed in control.
					edgeBarControl.Dispose()
				End If
			Else
				edgeBarPage = _edgeBarPageDictionary(pDocument)
			End If

			Return edgeBarPage
		End Function

		Private Sub RemovePage(ByVal pDocument As IntPtr)
			If _edgeBarPageDictionary.ContainsKey(pDocument) Then
				Dim edgeBarPage As EdgeBarPage = _edgeBarPageDictionary(pDocument)
				Dim hWndEdgeBarPage As IntPtr = edgeBarPage.Handle

				_edgeBarPageDictionary.Remove(pDocument)

				edgeBarPage.OnRemovePage()
				_edgeBar.RemovePage(edgeBarPage.SEDocument, hWndEdgeBarPage.ToInt32(), 0)

				edgeBarPage.Dispose()
			End If
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

#Region "IDisposable"

		Public Sub Dispose() Implements IDisposable.Dispose
			Dispose(True)
		End Sub

		Public Sub Dispose(ByVal disposing As Boolean)
			If Not _disposed Then
				If disposing Then
					UnhookAllEvents()

					Dim enumerator As Dictionary(Of IntPtr, EdgeBarPage).Enumerator = _edgeBarPageDictionary.GetEnumerator()

					Do While enumerator.MoveNext()
						RemovePage(enumerator.Current.Key)
					Loop

					_edgeBarPageDictionary.Clear()
				End If

				_edgeBarPageDictionary = Nothing
				_edgeBar = Nothing
				_addInEx = Nothing
				_disposed = True
			End If
		End Sub

#End Region
	End Class
End Namespace
