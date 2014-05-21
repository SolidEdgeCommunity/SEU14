Imports SolidEdgeContrib
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace Example1
	Friend Class ApplicationEventWatcher
		Inherits ConnectionPointControllerBase
		Implements SolidEdgeFramework.ISEApplicationEvents, IDisposable

		Private _disposed As Boolean = False
		Private _form As Form1

		Public Sub New(ByVal form As Form1, ByVal application As SolidEdgeFramework.Application)
			_form = form

			Me.AdviseSink(Of SolidEdgeFramework.ISEApplicationEvents)(application)
		End Sub

		#Region "IDisposable implementation"

		Protected Overrides Sub Finalize()
			Dispose(False)
		End Sub

		Public Sub Dispose() Implements IDisposable.Dispose
			Dispose(True)
			GC.SuppressFinalize(Me)
		End Sub

		Private Sub Dispose(ByVal disposing As Boolean)
			If Not _disposed Then
				If disposing Then
					Me.UnadviseAllSinks()
				End If

				_disposed = True
			End If
		End Sub

#End Region

		' Note: Events are fired in a background thread. You cannot update the UI
		' "directly" from a background thread. See ControlExtensions.BeginInvokeIfRequired().
		' Thread.CurrentThread.GetApartmentState() will always be ApartmentState.MTA.
		' OleMessageFilter is not in effect in this thread for two reasons. 1) It's a
		' different thread. 2) It can't be because the ApartmentState = MTA.

		#Region "SolidEdgeFramework.ISEApplicationEvents"

		Public Sub AfterActiveDocumentChange(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterActiveDocumentChange
			_form.BeginInvokeIfRequired(Sub(x) x.OnAfterActiveDocumentChange(theDocument))
		End Sub

		Public Sub AfterCommandRun(ByVal theCommandID As Integer) Implements SolidEdgeFramework.ISEApplicationEvents.AfterCommandRun
			_form.BeginInvokeIfRequired(Sub(x) x.OnAfterCommandRun(theCommandID))
		End Sub

		Public Sub AfterDocumentOpen(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterDocumentOpen
			_form.BeginInvokeIfRequired(Sub(x) x.OnAfterDocumentOpen(theDocument))
		End Sub

		Public Sub AfterDocumentPrint(ByVal theDocument As Object, ByVal hDC As Integer, ByRef ModelToDC As Double, ByRef Rect As Integer) Implements SolidEdgeFramework.ISEApplicationEvents.AfterDocumentPrint
			' Cannot use ref or out parameter 'ModelToDC' inside an anonymous method, lambda expression, or query expression.
			' Cannot use ref or out parameter 'Rect' inside an anonymous method, lambda expression, or query expression.
			Dim a = ModelToDC
			Dim b = Rect

			_form.InvokeIfRequired(Sub(x) x.OnAfterDocumentPrint(theDocument, hDC, a, b))
		End Sub

		Public Sub AfterDocumentSave(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterDocumentSave
			_form.BeginInvokeIfRequired(Sub(x) x.OnAfterDocumentSave(theDocument))
		End Sub

		Public Sub AfterEnvironmentActivate(ByVal theEnvironment As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterEnvironmentActivate
			_form.BeginInvokeIfRequired(Sub(x) x.OnAfterEnvironmentActivate(theEnvironment))
		End Sub

		Public Sub AfterNewDocumentOpen(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterNewDocumentOpen
			_form.BeginInvokeIfRequired(Sub(x) x.OnAfterNewDocumentOpen(theDocument))
		End Sub

		Public Sub AfterNewWindow(ByVal theWindow As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterNewWindow
			_form.BeginInvokeIfRequired(Sub(x) x.OnAfterNewWindow(theWindow))
		End Sub

		Public Sub AfterWindowActivate(ByVal theWindow As Object) Implements SolidEdgeFramework.ISEApplicationEvents.AfterWindowActivate
			_form.BeginInvokeIfRequired(Sub(x) x.OnAfterWindowActivate(theWindow))
		End Sub

		Public Sub BeforeCommandRun(ByVal theCommandID As Integer) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeCommandRun
			_form.BeginInvokeIfRequired(Sub(x) x.OnBeforeCommandRun(theCommandID))
		End Sub

		Public Sub BeforeDocumentClose(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentClose
			_form.BeginInvokeIfRequired(Sub(x) x.OnBeforeDocumentClose(theDocument))
		End Sub

		Public Sub BeforeDocumentPrint(ByVal theDocument As Object, ByVal hDC As Integer, ByRef ModelToDC As Double, ByRef Rect As Integer) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentPrint
			' Cannot use ref or out parameter 'ModelToDC' inside an anonymous method, lambda expression, or query expression.
			' Cannot use ref or out parameter 'Rect' inside an anonymous method, lambda expression, or query expression.
			Dim a = ModelToDC
			Dim b = Rect

			_form.InvokeIfRequired(Sub(x) x.OnBeforeDocumentPrint(theDocument, hDC, a, b))
		End Sub

		Public Sub BeforeDocumentSave(ByVal theDocument As Object) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentSave
			_form.BeginInvokeIfRequired(Sub(x) x.OnBeforeDocumentSave(theDocument))
		End Sub

		Public Sub BeforeEnvironmentDeactivate(ByVal theEnvironment As Object) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeEnvironmentDeactivate
			_form.BeginInvokeIfRequired(Sub(x) x.OnBeforeEnvironmentDeactivate(theEnvironment))
		End Sub

		Public Sub BeforeQuit() Implements SolidEdgeFramework.ISEApplicationEvents.BeforeQuit
			_form.BeginInvokeIfRequired(Sub(x) x.OnBeforeQuit())
		End Sub

		Public Sub BeforeWindowDeactivate(ByVal theWindow As Object) Implements SolidEdgeFramework.ISEApplicationEvents.BeforeWindowDeactivate
			_form.BeginInvokeIfRequired(Sub(x) x.OnBeforeWindowDeactivate(theWindow))
		End Sub

		#End Region
	End Class
End Namespace
