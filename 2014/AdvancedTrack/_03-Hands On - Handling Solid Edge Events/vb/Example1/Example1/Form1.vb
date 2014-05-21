Imports SolidEdgeContrib
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

Namespace Example1
	Partial Public Class Form1
		Inherits Form

		Private _application As SolidEdgeFramework.Application = Nothing
		Private _applicationEventWatcher As ApplicationEventWatcher = Nothing
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
			OleMessageFilter.Register()
		End Sub

		Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
			If _applicationEventWatcher IsNot Nothing Then
				_applicationEventWatcher.Dispose()
				_applicationEventWatcher = Nothing
			End If
			_application = Nothing
		End Sub

		Private Sub exitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles exitToolStripMenuItem.Click
			Close()
		End Sub

		Private Sub eventButton_Click(ByVal sender As Object, ByVal e As EventArgs) Handles eventButton.Click
			Try
				If eventButton.Checked Then
					If _application Is Nothing Then
						' On a system where Solid Edge is installed, the COM ProgID will be
						' defined in registry: HKEY_CLASSES_ROOT\SolidEdge.Application
						Dim t As Type = Type.GetTypeFromProgID(SolidEdge.PROGID.Application, throwOnError:= True)

						' Using the discovered Type, create and return a new instance of Solid Edge.
						_application = DirectCast(Activator.CreateInstance(type:= t), SolidEdgeFramework.Application)

						' Show Solid Edge.
						_application.Visible = True
					End If

					_applicationEventWatcher = New ApplicationEventWatcher(Me, _application)
				Else
					_applicationEventWatcher.Dispose()
					_applicationEventWatcher = Nothing
					_application = Nothing
				End If
			Catch ex As System.Exception
				MessageBox.Show(Me, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
			End Try
		End Sub

		Private Sub clearButton_Click(ByVal sender As Object, ByVal e As EventArgs) Handles clearButton.Click
			eventLogTextBox.Clear()
		End Sub

		#Region "SolidEdgeFramework.ISEApplicationEvents"

		Public Sub OnAfterActiveDocumentChange(ByVal theDocument As Object)
			eventLogTextBox.AppendText("OnAfterActiveDocumentChange")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnAfterCommandRun(ByVal theCommandID As Integer)
			eventLogTextBox.AppendText("OnAfterCommandRun")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnAfterDocumentOpen(ByVal theDocument As Object)
			eventLogTextBox.AppendText("OnAfterDocumentOpen")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnAfterDocumentPrint(ByVal theDocument As Object, ByVal hDC As Integer, ByVal ModelToDC As Double, ByVal Rect As Integer)
			eventLogTextBox.AppendText("OnAfterDocumentPrint")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnAfterDocumentSave(ByVal theDocument As Object)
			eventLogTextBox.AppendText("OnAfterDocumentSave")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnAfterEnvironmentActivate(ByVal theEnvironment As Object)
			eventLogTextBox.AppendText("OnAfterEnvironmentActivate")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnAfterNewDocumentOpen(ByVal theDocument As Object)
			eventLogTextBox.AppendText("OnAfterNewDocumentOpen")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnAfterNewWindow(ByVal theWindow As Object)
			eventLogTextBox.AppendText("OnAfterNewWindow")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnAfterWindowActivate(ByVal theWindow As Object)
			eventLogTextBox.AppendText("OnAfterWindowActivate")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnBeforeCommandRun(ByVal theCommandID As Integer)
			eventLogTextBox.AppendText("OnBeforeCommandRun")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnBeforeDocumentClose(ByVal theDocument As Object)
			eventLogTextBox.AppendText("OnBeforeDocumentClose")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnBeforeDocumentPrint(ByVal theDocument As Object, ByVal hDC As Integer, ByVal ModelToDC As Double, ByVal Rect As Integer)
			eventLogTextBox.AppendText("OnBeforeDocumentPrint")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnBeforeDocumentSave(ByVal theDocument As Object)
			eventLogTextBox.AppendText("OnBeforeDocumentSave")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnBeforeEnvironmentDeactivate(ByVal theEnvironment As Object)
			eventLogTextBox.AppendText("OnBeforeEnvironmentDeactivate")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		Public Sub OnBeforeQuit()
			eventLogTextBox.AppendText("OnBeforeQuit")
			eventLogTextBox.AppendText(Environment.NewLine)

			_applicationEventWatcher.Dispose()
			_applicationEventWatcher = Nothing
			_application = Nothing
		End Sub

		Public Sub OnBeforeWindowDeactivate(ByVal theWindow As Object)
			eventLogTextBox.AppendText("OnBeforeWindowDeactivate")
			eventLogTextBox.AppendText(Environment.NewLine)
		End Sub

		#End Region

	End Class
End Namespace
