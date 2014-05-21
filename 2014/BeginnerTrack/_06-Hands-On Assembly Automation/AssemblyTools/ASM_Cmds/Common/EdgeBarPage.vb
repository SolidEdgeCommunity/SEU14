Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

Namespace SolidEdge.CommonUI
	''' <summary>
	''' Wrapper class for hWnd returned from ISolidEdgeBarEx.AddPageEx().
	''' </summary>
	Public Class EdgeBarPage
		Inherits NativeWindow
		Implements IDisposable

		Private _disposed As Boolean = False
		Private _document As SolidEdgeFramework.SolidEdgeDocument
		Private _edgeBarControl As EdgeBarControl

		Public Sub New(ByVal hWnd As IntPtr, ByVal theDocument As Object, ByVal edgeBarControl As EdgeBarControl)
			Me.New(hWnd, DirectCast(theDocument, SolidEdgeFramework.SolidEdgeDocument), edgeBarControl)
		End Sub

		Public Sub New(ByVal hWnd As IntPtr, ByVal document As SolidEdgeFramework.SolidEdgeDocument, ByVal edgeBarControl As EdgeBarControl)
			_document = document
			_edgeBarControl = edgeBarControl

			_edgeBarControl.EdgeBarPage = Me

			' Assign the hWnd to this class.
			AssignHandle(hWnd)

			' Reparent child control to this hWnd.
			NativeMethods.SetParent(_edgeBarControl.Handle, Handle)

			' Show the child control.
			NativeMethods.ShowWindow(_edgeBarControl.Handle, ShowWindowCommands.Maximize)
		End Sub

		Protected Overrides Sub Finalize()
			Dispose(False)
		End Sub

		Public Sub Dispose() Implements IDisposable.Dispose
			Dispose(True)
			GC.SuppressFinalize(Me)
		End Sub

		Public Sub Dispose(ByVal disposing As Boolean)
			If Not _disposed Then
				If disposing Then
					If (_edgeBarControl IsNot Nothing) AndAlso ((Not _edgeBarControl.IsDisposed)) Then
						Try
							_edgeBarControl.Dispose()
						Catch
						End Try
					End If
				End If

				Try
					ReleaseHandle()
				Catch
				End Try

				_edgeBarControl = Nothing
				_disposed = True
			End If
		End Sub

		Public Sub OnRemovePage()
			_edgeBarControl.OnRemovePage()
		End Sub

		Public ReadOnly Property SEApplication() As SolidEdgeFramework.Application
			Get
				Return _document.Application
			End Get
		End Property

		Public ReadOnly Property SEDocument() As SolidEdgeFramework.SolidEdgeDocument
			Get
				Return _document
			End Get
		End Property
		Public ReadOnly Property SEEdgeBarControl() As EdgeBarControl
			Get
				Return _edgeBarControl
			End Get
		End Property
	End Class
End Namespace
