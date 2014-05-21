Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace SolidEdge.CommonUI
	Public Class RibbonCommand
		Private _environmentCATID As String
		Private _environmentGuid As Guid = Guid.Empty
		Private _groupName As String
		Private _commandId As Integer
		Private _name As String
		Private _caption As String
		Private _statusBarText As String
		Private _tooltip As String
		Private _buttonStyle As SolidEdgeFramework.SeButtonStyle
		Private _macro As String
		Private _macroParameters As String

		Public SystemCommandId As Integer = -1
		Public Enabled As Boolean = True
		Public Checked As Boolean = False
		Public Callback As RibbonCommandDelegate = Nothing

		Public Sub New(ByVal groupName As String, ByVal commandId As Integer, ByVal name As String, ByVal caption As String, ByVal statusBarText As String, ByVal tooltip As String, ByVal buttonStyle As SolidEdgeFramework.SeButtonStyle)
			_groupName = groupName
			_commandId = commandId
			_name = name
			_caption = caption
			_statusBarText = statusBarText
			_tooltip = tooltip
			_buttonStyle = buttonStyle
		End Sub

		Public Sub New(ByVal groupName As String, ByVal commandId As Integer, ByVal name As String, ByVal caption As String, ByVal statusBarText As String, ByVal tooltip As String, ByVal buttonStyle As SolidEdgeFramework.SeButtonStyle, ByVal callback As RibbonCommandDelegate)
			Me.New(groupName, commandId, name, caption, statusBarText, tooltip, buttonStyle)
			Me.Callback = callback
		End Sub

		Public Sub New(ByVal groupName As String, ByVal commandId As Integer, ByVal name As String, ByVal caption As String, ByVal statusBarText As String, ByVal tooltip As String, ByVal macro As String, ByVal macroParameters As String, ByVal buttonStyle As SolidEdgeFramework.SeButtonStyle)
			Me.New(groupName, commandId, name, caption, statusBarText, tooltip, buttonStyle)
			_macro = macro
			_macroParameters = macroParameters
		End Sub

		''' <summary>
		''' Properly formats the CommandName to be used in the CommandNames array of SetAddInInfoEx().
		''' </summary>
		''' <returns></returns>
		Public Function ToCommandName() As String
			Dim sb As New StringBuilder()

			sb.AppendFormat("{0}" & ControlChars.Lf & "{1}" & ControlChars.Lf & "{2}" & ControlChars.Lf & "{3}", Name, Caption, StatusBarText, Tooltip)

			' Append macro info if provided.
			If Not String.IsNullOrEmpty(Macro) Then
				sb.AppendFormat(ControlChars.Lf & "{0}", Macro)

				If Not String.IsNullOrEmpty(MacroParameters) Then
					sb.AppendFormat(ControlChars.Lf & "{0}", MacroParameters)
				End If
			End If

			Return sb.ToString()
		End Function

		Public Property EnvironmentCATID() As String
			Get
				Return _environmentCATID
			End Get
			Set(ByVal value As String)
				_environmentCATID = value
				_environmentGuid = New Guid(value)
			End Set
		End Property
		Public ReadOnly Property EnvironmentGuid() As Guid
			Get
				Return _environmentGuid
			End Get
		End Property
		Public ReadOnly Property GroupName() As String
			Get
				Return _groupName
			End Get
		End Property
		Public ReadOnly Property CommandId() As Integer
			Get
				Return _commandId
			End Get
		End Property
		Public ReadOnly Property Name() As String
			Get
				Return _name
			End Get
		End Property
		Public ReadOnly Property Caption() As String
			Get
				Return _caption
			End Get
		End Property
		Public ReadOnly Property StatusBarText() As String
			Get
				Return _statusBarText
			End Get
		End Property
		Public ReadOnly Property Tooltip() As String
			Get
				Return _tooltip
			End Get
		End Property
		Public ReadOnly Property ButtonStyle() As SolidEdgeFramework.SeButtonStyle
			Get
				Return _buttonStyle
			End Get
		End Property
		Public ReadOnly Property Macro() As String
			Get
				Return _macro
			End Get
		End Property
		Public ReadOnly Property MacroParameters() As String
			Get
				Return _macroParameters
			End Get
		End Property
	End Class
End Namespace
