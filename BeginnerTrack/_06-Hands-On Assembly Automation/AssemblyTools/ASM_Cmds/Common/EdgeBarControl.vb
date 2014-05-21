Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Linq
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text

Namespace SolidEdge.CommonUI
	''' <summary>
	''' Base EdgeBar control. Provides core functionality.
	''' </summary>
	''' <remarks>
	''' Not intended to be directly created but rather inherited from.
	''' </remarks>
	Public Class EdgeBarControl
		Inherits System.Windows.Forms.UserControl

		Private _edgeBarPage As EdgeBarPage
		Private _tooltip As String = String.Empty
		Private _bitmapID As Integer
		Private _isPageDisplayable As Boolean = True

		Public Sub New()
			MyBase.New()
		End Sub

		Public Overridable Sub OnRemovePage()
		End Sub

#Region "Properties"

		<Browsable(False)> _
		Public Overridable Property EdgeBarPage() As EdgeBarPage
			Get
				Return _edgeBarPage
			End Get
			Set(ByVal value As EdgeBarPage)
				_edgeBarPage = value
			End Set
		End Property

		<Browsable(False)> _
		Public ReadOnly Property SEDocument() As SolidEdgeFramework.SolidEdgeDocument
			Get
				Return _edgeBarPage.SEDocument
			End Get
		End Property

		<Browsable(False)> _
		Public ReadOnly Property SEApplication() As SolidEdgeFramework.Application
			Get
				Return _edgeBarPage.SEApplication
			End Get
		End Property

		''' <summary>
		''' The ID of the Bitmap to be used in the EdgeBarPage.
		''' </summary>
		''' <remarks>
		''' Win32 resources are located in the Resources.res file.
		''' </remarks>
		<Browsable(True)> _
		Public Property BitmapID() As Integer
			Get
				Return _bitmapID
			End Get
			Set(ByVal value As Integer)
				_bitmapID = value
			End Set
		End Property

		''' <summary>
		''' Called during SolidEdgeFramework.ISEAddInEdgeBarEvents.IsPageDisplayable().
		''' </summary>
		<Browsable(False)> _
		Public Property IsPageDisplayable() As Boolean
			Get
				Return _isPageDisplayable
			End Get
			Set(ByVal value As Boolean)
				_isPageDisplayable = value
			End Set
		End Property

		''' <summary>
		''' The string to be used in the EdgeBarPage caption and tooltip.
		''' </summary>
		<Browsable(True)> _
		Public Property ToolTip() As String
			Get
				Return _tooltip
			End Get
			Set(ByVal value As String)
				_tooltip = value
			End Set
		End Property

#End Region

 Private Sub InitializeComponent()
		Me.SuspendLayout()
		'
		'EdgeBarControl
		'
		Me.Name = "EdgeBarControl"
		Me.ResumeLayout(False)

End Sub
	End Class
End Namespace
