Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms

Namespace Example2
	Partial Public Class Form1
		Inherits Form

		Private _application As SolidEdgeFramework.Application

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Toggle the label visible state.
			label1.Visible = Not label1.Visible

			' Get a reference to Solid Edge if we don't already have one.
			If _application Is Nothing Then
				Try
					' Attempt to connect to a running instace.
					_application = DirectCast(Marshal.GetActiveObject(SolidEdge.PROGID.Application), SolidEdgeFramework.Application)
				Catch
					' On a system where Solid Edge is installed, the COM ProgID will be
					' defined in registry: HKEY_CLASSES_ROOT\SolidEdge.Application
					Dim t As Type = Type.GetTypeFromProgID(SolidEdge.PROGID.Application, throwOnError:= True)

					' Using the discovered Type, create and return a new instance of Solid Edge.
					_application = DirectCast(Activator.CreateInstance(type:= t), SolidEdgeFramework.Application)
				End Try
			End If

			' Make sure Solid Edge is visible.
			_application.Visible = True

			' See what AppDomain we're currently executing in.
			Dim currentAppDomain = AppDomain.CurrentDomain

			' This will always be the default AppDomain at this point.
			Dim isDefaultAppDomain = currentAppDomain.IsDefaultAppDomain()

			' Create a separate AppDomain and execute our code.
			CreateSeparateAppDomainAndExecuteIsolatedTask(_application)

			' Toggle the label visible state.
			label1.Visible = Not label1.Visible
		End Sub

		Private Sub CreateSeparateAppDomainAndExecuteIsolatedTask(ByVal application As SolidEdgeFramework.Application)
			Dim interopDomain As AppDomain = Nothing

			Try
				Dim thread = New System.Threading.Thread(Sub()
					' Create a custom AppDomain to do COM Interop.
					' Create a new instance of InteropProxy in the isolated application domain.
					interopDomain = AppDomain.CreateDomain("Interop Domain")
					Dim proxyType As Type = GetType(InteropProxy)
					Dim interopProxy As InteropProxy = TryCast(interopDomain.CreateInstanceAndUnwrap(proxyType.Assembly.FullName, proxyType.FullName), InteropProxy)
					Try
						interopProxy.DoIsolatedTask(application)
					Catch ex As System.Exception
						MessageBox.Show(ex.StackTrace, ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error)
					End Try
				End Sub)

				' Important! Set thread apartment state to STA.
				thread.SetApartmentState(System.Threading.ApartmentState.STA)

				' Start the thread.
				thread.Start()

				' Wait for the thead to finish.
				thread.Join()
			Catch ex As System.Exception
				MessageBox.Show(ex.StackTrace, ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error)
			Finally
				If interopDomain IsNot Nothing Then
					' Unload the Interop AppDomain. This will automatically free up any COM references.
					AppDomain.Unload(interopDomain)
				End If
			End Try
		End Sub
	End Class
End Namespace
