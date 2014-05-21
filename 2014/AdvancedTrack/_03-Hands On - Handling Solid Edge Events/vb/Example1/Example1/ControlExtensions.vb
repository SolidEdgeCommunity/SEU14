Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms


Namespace Example1
	Public Module ControlExtensions
		''' <summary>
		''' Synchornous call to control but may load to deadlock if events are firing rapidly.
		''' </summary>
		<System.Runtime.CompilerServices.Extension> _
		Public Sub InvokeIfRequired(Of TControl As Control)(ByVal control As TControl, ByVal action As Action(Of TControl))
			If control.InvokeRequired Then
				control.Invoke(action, control)
			Else
				action(control)
			End If
		End Sub

		''' <summary>
		''' Asynchornous to control and should be safest.
		''' </summary>
		<System.Runtime.CompilerServices.Extension> _
		Public Sub BeginInvokeIfRequired(Of TControl As Control)(ByVal control As TControl, ByVal action As Action(Of TControl))
			If control.InvokeRequired Then
				control.BeginInvoke(action, control)
			Else
				action(control)
			End If
		End Sub
	End Module
End Namespace
