Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace SolidEdge.CommonUI

	Public Class RibbonCommandInfo
		Public ResourceFilename As String = Nothing
		Public EnvironmentCatID As String = Nothing
		Public RibbonTabName As String = Nothing
		Public IDColorBitmapMedium As Integer = -1
		Public IDColorBitmapLarge As Integer = -1
		Public IDMonochromeBitmapMedium As Integer = -1
		Public IDMonochromeBitmapLarge As Integer = -1
		Public RibbonCommands As New List(Of RibbonCommand)()

		''' <summary>
		''' Gets a properly formatted CommandNames array for SetAddInInfoEx().
		''' </summary>
		Public Function GetCommandNames() As Array
			Dim list As New List(Of String)()

			For Each ribbonCommand As RibbonCommand In RibbonCommands
				list.Add(ribbonCommand.ToCommandName())
			Next ribbonCommand

			Return list.ToArray()
		End Function

		''' <summary>
		''' Gets a properly formatted CommandIDs array for SetAddInInfoEx().
		''' </summary>
		Public Function GetCommandIDs() As Array
			Dim list As New List(Of Integer)()

			For Each ribbonCommand As RibbonCommand In RibbonCommands
				list.Add(ribbonCommand.CommandId)
			Next ribbonCommand

			Return list.ToArray()
		End Function

		''' <summary>
		''' Updates each RibbonCommand.SystemCommandId property to the updated values of CommandIDs after SetAddInInfoEx() is called.
		''' </summary>
		Public Sub UpdateEnvironmentCATIDs()
			For Each ribbonCommand As RibbonCommand In RibbonCommands
				ribbonCommand.EnvironmentCATID = EnvironmentCatID
			Next ribbonCommand
		End Sub

		''' <summary>
		''' Updates each RibbonCommand.SystemCommandId property to the updated values of CommandIDs after SetAddInInfoEx() is called.
		''' </summary>
		''' <param name="CommandIDs"></param>
		Public Sub UpdateSystemCommandIDs(ByVal CommandIDs As Array)
			For i As Integer = 0 To CommandIDs.Length - 1
				RibbonCommands(i).SystemCommandId = DirectCast(CommandIDs.GetValue(i), Integer)
			Next i
		End Sub
	End Class
End Namespace
