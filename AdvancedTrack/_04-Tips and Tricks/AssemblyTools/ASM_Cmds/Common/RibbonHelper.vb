Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Windows.Forms

Namespace SolidEdge.CommonUI
	''' <summary>
	''' Helper class to manager ribbon commands for each Solid Edge environment.
	''' </summary>
	Public NotInheritable Class RibbonHelper

		Private Sub New()
		End Sub

		' Static dictionary to hold ribbon commands per environment.
		Private Shared _environmentRibbonCommmands As New Dictionary(Of Guid, List(Of RibbonCommand))()

		Shared Sub New()
		End Sub

		''' <summary>
		''' Lookup ribbon command by addin defined command id.
		''' </summary>
		Public Shared Function LookupRibbonCommandByCommandId(ByVal environmentGuid As Guid, ByVal CommandID As Integer) As RibbonCommand
			If _environmentRibbonCommmands.ContainsKey(environmentGuid) Then
				Return _environmentRibbonCommmands(environmentGuid).Where(Function(x) x.CommandId = CommandID).FirstOrDefault()
			End If

			Return Nothing
		End Function

		''' <summary>
		''' Lookup ribbon command by Solid Edge defined command id.
		''' </summary>
		''' <remarks>Solid Edge will take the CommandID assigned by the addin and define a system CommandID.</remarks>
		Public Shared Function LookupRibbonCommandBySystemCommandId(ByVal environmentGuid As Guid, ByVal CommandID As Integer) As RibbonCommand
			If _environmentRibbonCommmands.ContainsKey(environmentGuid) Then
				Return _environmentRibbonCommmands(environmentGuid).Where(Function(x) x.SystemCommandId = CommandID).FirstOrDefault()
			End If

			Return Nothing
		End Function

		''' <summary>
		''' Handles adding commands to Solid Edge.
		''' </summary>
		''' <remarks>Should only be called from SolidEdgeFramework.ISolidEdgeAddIn.OnConnectToEnvironment().</remarks>
		Public Shared Sub AddRibbonCommands(ByVal addInEx As SolidEdgeFramework.ISEAddInEx, ByVal ribbonCommandInfo As RibbonCommandInfo, ByVal bFirstTime As Boolean)
			' Allocate command arrays. Please see the addin.doc in the SDK folder for details.
			Dim commandNames As Array = ribbonCommandInfo.GetCommandNames()
			Dim commandIDs As Array = ribbonCommandInfo.GetCommandIDs()

			' Inform Solid Edge about the commands to add.
			addInEx.SetAddInInfoEx(ResourceFilename:=ribbonCommandInfo.ResourceFilename, EnvironmentCatID:=ribbonCommandInfo.EnvironmentCatID, CategoryName:=ribbonCommandInfo.RibbonTabName, IDColorBitmapMedium:=ribbonCommandInfo.IDColorBitmapMedium, IDColorBitmapLarge:=ribbonCommandInfo.IDColorBitmapLarge, IDMonochromeBitmapMedium:=ribbonCommandInfo.IDMonochromeBitmapMedium, IDMonochromeBitmapLarge:=ribbonCommandInfo.IDMonochromeBitmapLarge, NumberOfCommands:=commandNames.Length, CommandNames:=commandNames, CommandIDs:=commandIDs)

			' Solid Edge converted our commandIDs array to system command ids. This method will update each RibbonCommand object accordingly.
			ribbonCommandInfo.UpdateSystemCommandIDs(commandIDs)

			' This method stores the environment CATID in each RibbonCommand object for later use.
			ribbonCommandInfo.UpdateEnvironmentCATIDs()

			' If this is the first time to connect to the environment, configure the command bar buttons.
			' TIP: If make changes to your commands, increment your addin's GuiVersion to force bFirstTime to be true;
			If bFirstTime Then
				For Each ribbonCommand As RibbonCommand In ribbonCommandInfo.RibbonCommands
					' Properly format the command bar name string.
					Dim commandBarName As String = String.Format("{0}" & ControlChars.Lf & "{1}", ribbonCommandInfo.RibbonTabName, ribbonCommand.GroupName)

					' Add the command bar button.
					Dim pButton As SolidEdgeFramework.CommandBarButton = addInEx.AddCommandBarButton(ribbonCommandInfo.EnvironmentCatID, commandBarName, ribbonCommand.CommandId)

					' Set the button style.
					If pButton IsNot Nothing Then
						pButton.Style = ribbonCommand.ButtonStyle
					End If
				Next ribbonCommand
			End If

			' Get the GUID of the environment.
			Dim environmentGuid As New Guid(ribbonCommandInfo.EnvironmentCatID)

			' If the dictionary does not already have a key containing the environment guid, add it.
			If Not _environmentRibbonCommmands.ContainsKey(environmentGuid) Then
				_environmentRibbonCommmands.Add(environmentGuid, New List(Of RibbonCommand)())
			End If

			' Store the ribbon commands in the dictionary using the specified environment guid as the key.
			_environmentRibbonCommmands(environmentGuid).AddRange(ribbonCommandInfo.RibbonCommands.ToArray())
		End Sub

		''' <summary>
		''' Resets the class back to it's original state.
		''' </summary>
		''' <remarks>Typically should only be called from SolidEdgeFramework.ISolidEdgeAddIn.OnDisconnection().</remarks>
		Public Shared Sub Reset()
			_environmentRibbonCommmands.Clear()
		End Sub
	End Class
End Namespace
