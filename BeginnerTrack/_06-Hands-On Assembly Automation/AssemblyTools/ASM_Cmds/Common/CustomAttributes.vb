Imports System

Namespace SolidEdge.CommonUI
	''' <summary>
	''' Environment Categories the addin is registered to.
	''' </summary>
	<AttributeUsage(AttributeTargets.Class, AllowMultiple:=True)> _
	Public Class AddInEnvironmentCategoryAttribute
		Inherits System.Attribute

		Private _guid As Guid = Guid.Empty

		''' <summary>
		''' Constructor.
		''' </summary>
		''' <param name="guid">Solid Edge category IDs as defined in \sdk\include\secatids.h.</param>
		Public Sub New(ByVal guid As String)
			Me.New(New Guid(guid))
		End Sub

		''' <summary>
		''' Constructor.
		''' </summary>
		''' <param name="guid">Solid Edge category IDs as defined in \sdk\include\secatids.h.</param>
		Public Sub New(ByVal guid As Guid)
			_guid = guid
		End Sub

		''' <summary>
		''' Solid Edge environment CATID.
		''' </summary>
		Public ReadOnly Property Guid() As Guid
			Get
				Return _guid
			End Get
		End Property
	End Class

	''' <summary>
	''' Information about the addin.
	''' </summary>
	''' 
	<AttributeUsage(AttributeTargets.Class, AllowMultiple:=False)> _
	Public Class AddInInfoAttribute
		Inherits System.Attribute

		Private _title As String = String.Empty
		Private _summary As String = String.Empty
		Private _autoConnect As Boolean = True

		''' <summary>
		''' Constructor.
		''' </summary>
		''' <param name="title">AddIn title. Prepend \n to have top level menu (Ribbon Tab).</param>
		''' <param name="summary">AddIn summary.</param>
		''' <param name="autoConnect">Set AutoConnect DWORD flag. If true, addin be enabled by default.</param>
		Public Sub New(ByVal title As String, ByVal summary As String, ByVal autoConnect As Boolean)
			_title = title
			_summary = summary
			_autoConnect = autoConnect
		End Sub

		''' <summary>
		''' AddIn title.
		''' </summary>
		Public ReadOnly Property Title() As String
			Get
				Return _title
			End Get
		End Property

		''' <summary>
		''' AddIn summary.
		''' </summary>
		Public ReadOnly Property Summary() As String
			Get
				Return _summary
			End Get
		End Property

		''' <summary>
		''' If true, addin be enabled by default.
		''' </summary>
		Public ReadOnly Property AutoConnect() As Boolean
			Get
				Return _autoConnect
			End Get
		End Property
	End Class
End Namespace
