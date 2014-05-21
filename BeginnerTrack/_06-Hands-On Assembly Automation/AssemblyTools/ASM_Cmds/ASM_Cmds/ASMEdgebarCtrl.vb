Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports SolidEdge.CommonUI
Imports SolidEdgeAssembly

Namespace SolidEdge.ASM_Edgebar_Cmds

Public Class ASMEdgebarCtrl
    Inherits EdgeBarControl
		Implements SolidEdgeFramework.ISEDocumentEvents

		Private _Mouse As clsMouseEvents = Nothing
		Private _connectionPointDictionary As New Dictionary(Of IConnectionPoint, Integer)()
		Private _TooltipCtrl As New ToolTip

    Public Sub New()
      InitializeComponent()

   ' Set up the delays for the ToolTip.
     _TooltipCtrl.AutoPopDelay = 5000
     _TooltipCtrl.InitialDelay = 1000
     _TooltipCtrl.ReshowDelay = 500
     ' Force the ToolTip text to be displayed whether or not the form is active.
     _TooltipCtrl.ShowAlways = True

     ' Set up the ToolTip text for the Button and Checkbox.
     _TooltipCtrl.SetToolTip(Me.btnReplaceWithCopy, "Copy and Replace selected components")
     _TooltipCtrl.SetToolTip(Me.btnToggleCsys, "Toggle Coordinate System display")
     _TooltipCtrl.SetToolTip(Me.btnToggleRefPlanes, "Toggle Reference Plane display")
     _TooltipCtrl.SetToolTip(Me.btnToggleSketches, "Toggle Sketch display")
     _TooltipCtrl.SetToolTip(Me.cmbQueryFileName, "File name query string.")
     _TooltipCtrl.SetToolTip(Me.btnRunQuery, "Executes the component file name query.")

    End Sub

	Public Overrides Sub OnRemovePage()
			If Not IsNothing(_Mouse) Then
				_Mouse.Dispose()
				_Mouse = Nothing
			End If
      UnhookAllEvents()
		End Sub

#Region "SolidEdgeFramework.ISEDocumentEvents"

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISEDocumentEvents.AfterSave().
		''' </summary>
		Public Sub AfterSave() Implements SolidEdgeFramework.ISEDocumentEvents.AfterSave
		End Sub

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISEDocumentEvents.BeforeClose().
		''' </summary>
		Public Sub BeforeClose() Implements SolidEdgeFramework.ISEDocumentEvents.BeforeClose
			' BeforeClose will likely never get called because OnRemovePage()
			' will get called 1st which will unhook the ISEDocumentEvents.
		End Sub

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISEDocumentEvents.BeforeSave().
		''' </summary>
		Public Sub BeforeSave() Implements SolidEdgeFramework.ISEDocumentEvents.BeforeSave
		End Sub

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISEDocumentEvents.SelectSetChanged().
		''' </summary>
		Public Sub SelectSetChanged(ByVal SelectSet As Object) Implements SolidEdgeFramework.ISEDocumentEvents.SelectSetChanged

		End Sub

#End Region

#Region "SelectSetEdgeBarControl methods"

		Private Function GetComObjectFullyQualifiedName(ByVal o As Object) As String
			If o Is Nothing Then
				Throw New ArgumentNullException()
			End If

			If Marshal.IsComObject(o) Then
				Dim dispatch As IDispatch = TryCast(o, IDispatch)
				If dispatch IsNot Nothing Then
					Dim typeLib As ITypeLib = Nothing
					Dim typeInfo As ITypeInfo = dispatch.GetTypeInfo(0, 0)
					Dim index As Integer = 0
					typeInfo.GetContainingTypeLib(typeLib, index)

					Dim typeLibName As String = Marshal.GetTypeLibName(typeLib)
					Dim typeInfoName As String = Marshal.GetTypeInfoName(typeInfo)

					Return String.Format("{0}.{1}", typeLibName, typeInfoName)
				End If
			End If

			Return o.GetType().FullName
		End Function

#End Region

#Region "SelectSetEdgeBarControl properties"

		<Browsable(False)> _
		Public Overrides Property EdgeBarPage() As EdgeBarPage
			Get
				Return MyBase.EdgeBarPage
			End Get
			Set(ByVal value As EdgeBarPage)
				MyBase.EdgeBarPage = value

				If (EdgeBarPage IsNot Nothing) AndAlso (EdgeBarPage.SEDocument IsNot Nothing) Then
					HookEvents(EdgeBarPage.SEDocument.DocumentEvents, GetType(SolidEdgeFramework.ISEDocumentEvents).GUID)
				End If
			End Set
		End Property

#End Region

#Region "IConnectionPoint helpers"

		Private Sub HookEvents(ByVal eventSource As Object, ByVal eventGuid As Guid)
'INSTANT VB NOTE: The variable container was renamed since Visual Basic does not handle local variables named the same as class members well:
			Dim container_Renamed As IConnectionPointContainer = Nothing
			Dim connectionPoint As IConnectionPoint = Nothing
			Dim cookie As Integer = 0

			container_Renamed = DirectCast(eventSource, IConnectionPointContainer)
			container_Renamed.FindConnectionPoint(eventGuid, connectionPoint)

			If connectionPoint IsNot Nothing Then
				connectionPoint.Advise(Me, cookie)
				_connectionPointDictionary.Add(connectionPoint, cookie)
			End If
		End Sub

		Private Sub UnhookAllEvents()
			Dim enumerator As Dictionary(Of IConnectionPoint, Integer).Enumerator = _connectionPointDictionary.GetEnumerator()
			Do While enumerator.MoveNext()
				enumerator.Current.Key.Unadvise(enumerator.Current.Value)
			Loop

			_connectionPointDictionary.Clear()
		End Sub
	Friend WithEvents btnToggleCsys As System.Windows.Forms.Button

#End Region

Private Sub InitializeComponent()
    Me.btnReplaceWithCopy = New System.Windows.Forms.Button()
    Me.cmbQueryFileName = New System.Windows.Forms.ComboBox()
    Me.btnRunQuery = New System.Windows.Forms.Button()
    Me.btnToggleSketches = New System.Windows.Forms.Button()
    Me.btnToggleRefPlanes = New System.Windows.Forms.Button()
    Me.btnToggleCsys = New System.Windows.Forms.Button()
    Me.SuspendLayout()
    '
    'btnReplaceWithCopy
    '
    Me.btnReplaceWithCopy.Image = Global.My.Resources.MyResources.btnRPC
    Me.btnReplaceWithCopy.Location = New System.Drawing.Point(232, 1)
    Me.btnReplaceWithCopy.Name = "btnReplaceWithCopy"
    Me.btnReplaceWithCopy.Size = New System.Drawing.Size(20, 20)
    Me.btnReplaceWithCopy.TabIndex = 3
    Me.btnReplaceWithCopy.UseVisualStyleBackColor = True
    '
    'cmbQueryFileName
    '
    Me.cmbQueryFileName.DropDownHeight = 200
    Me.cmbQueryFileName.DropDownWidth = 200
    Me.cmbQueryFileName.FormattingEnabled = True
    Me.cmbQueryFileName.IntegralHeight = False
    Me.cmbQueryFileName.ItemHeight = 13
    Me.cmbQueryFileName.Items.AddRange(New Object() {"<Design>", "<Reference>", "<Not Selectable>"})
    Me.cmbQueryFileName.Location = New System.Drawing.Point(68, 1)
    Me.cmbQueryFileName.Name = "cmbQueryFileName"
    Me.cmbQueryFileName.Size = New System.Drawing.Size(139, 21)
    Me.cmbQueryFileName.TabIndex = 4
    '
    'btnRunQuery
    '
    Me.btnRunQuery.Image = Global.My.Resources.MyResources.btnQ_GO
    Me.btnRunQuery.Location = New System.Drawing.Point(207, 1)
    Me.btnRunQuery.Name = "btnRunQuery"
    Me.btnRunQuery.Size = New System.Drawing.Size(20, 20)
    Me.btnRunQuery.TabIndex = 5
    Me.btnRunQuery.UseVisualStyleBackColor = True
    '
    'btnToggleSketches
    '
    Me.btnToggleSketches.Image = Global.My.Resources.MyResources.btnSketch
    Me.btnToggleSketches.Location = New System.Drawing.Point(42, 1)
    Me.btnToggleSketches.Name = "btnToggleSketches"
    Me.btnToggleSketches.Size = New System.Drawing.Size(20, 20)
    Me.btnToggleSketches.TabIndex = 0
    Me.btnToggleSketches.UseVisualStyleBackColor = True
    '
    'btnToggleRefPlanes
    '
    Me.btnToggleRefPlanes.Image = Global.My.Resources.MyResources.btnRefplanes
    Me.btnToggleRefPlanes.Location = New System.Drawing.Point(21, 1)
    Me.btnToggleRefPlanes.Name = "btnToggleRefPlanes"
    Me.btnToggleRefPlanes.Size = New System.Drawing.Size(20, 20)
    Me.btnToggleRefPlanes.TabIndex = 0
    Me.btnToggleRefPlanes.UseVisualStyleBackColor = True
    '
    'btnToggleCsys
    '
    Me.btnToggleCsys.Image = Global.My.Resources.MyResources.btnCsys
    Me.btnToggleCsys.Location = New System.Drawing.Point(1, 1)
    Me.btnToggleCsys.Name = "btnToggleCsys"
    Me.btnToggleCsys.Size = New System.Drawing.Size(20, 20)
    Me.btnToggleCsys.TabIndex = 0
    Me.btnToggleCsys.UseVisualStyleBackColor = True
    '
    'ASMEdgebarCtrl
    '
    Me.BitmapID = 101
    Me.Controls.Add(Me.btnRunQuery)
    Me.Controls.Add(Me.cmbQueryFileName)
    Me.Controls.Add(Me.btnReplaceWithCopy)
    Me.Controls.Add(Me.btnToggleSketches)
    Me.Controls.Add(Me.btnToggleRefPlanes)
    Me.Controls.Add(Me.btnToggleCsys)
    Me.Name = "ASMEdgebarCtrl"
    Me.Size = New System.Drawing.Size(257, 38)
    Me.ToolTip = "ASM Edgebar"
    Me.ResumeLayout(False)

End Sub

Public Sub StopCommand()
		Try
			If Not IsNothing(_Mouse) Then
				_Mouse.Dispose()
        _Mouse = Nothing
#If True Then
       SEApplication.StartCommand(SolidEdgeFramework.SolidEdgeCommandConstants.seAssemblySelectCommand)
#End If

			End If
		Catch ex As Exception
			MsgBox(ex.Message, , "StopCommand")
		End Try
End Sub

Private Sub btnToggleCsys_Click(sender As Object, e As EventArgs) Handles btnToggleCsys.Click
Try
		If IsNothing(_Mouse) Then
			_Mouse = New clsMouseEvents
		Else
			_Mouse.EndAction()
		End If

		_Mouse.RunAction(cmdActions.ToggleCsys, SEApplication, Me)

	Catch ex As Exception
		MsgBox(ex.Message, , "btnToggleCsysDisplay_Click")
End Try

End Sub

Private Sub btnToggleRefPlanes_Click(sender As Object, e As EventArgs) Handles btnToggleRefPlanes.Click
Try
		If IsNothing(_Mouse) Then
			_Mouse = New clsMouseEvents
		Else
			_Mouse.EndAction()
		End If

		_Mouse.RunAction(cmdActions.ToggleRefPlanes, SEApplication, Me)
	Catch ex As Exception
		MsgBox(ex.Message, , "btnToggleRefPlanes_Click")
End Try

End Sub

Private Sub btnToggleSketches_Click(sender As Object, e As EventArgs) Handles btnToggleSketches.Click
Try
		If IsNothing(_Mouse) Then
			_Mouse = New clsMouseEvents
		Else
			_Mouse.EndAction()
		End If

		_Mouse.RunAction(cmdActions.ToggleSketches, SEApplication, Me)
	Catch ex As Exception
		MsgBox(ex.Message, , "btnToggleSketches_Click")
End Try

End Sub

Friend WithEvents btnToggleRefPlanes As System.Windows.Forms.Button
Friend WithEvents btnToggleSketches As System.Windows.Forms.Button
Friend WithEvents btnReplaceWithCopy As System.Windows.Forms.Button
Private components As System.ComponentModel.IContainer
Friend WithEvents cmbQueryFileName As System.Windows.Forms.ComboBox
Friend WithEvents btnRunQuery As System.Windows.Forms.Button


Private Sub btnReplaceWithCopy_Click(sender As Object, e As EventArgs) Handles btnReplaceWithCopy.Click
  'Connect to the xls for the first time if needed
  'All docs use the same xls
	If IsNothing(m_ExcelNameList) Then
		m_ExcelNameList = New clsExcelData
	End If

	If Not m_ExcelNameList.InitOK Then
		m_ExcelNameList.InitClass()
		If Not m_ExcelNameList.InitOK Then Exit Sub
	End If

	CloneSelectedComponents(SEApplication)

	SEApplication.StartCommand(SolidEdgeFramework.SolidEdgeCommandConstants.seAssemblySelectCommand)

End Sub



Private Sub btnRunQuery_Click(sender As Object, e As EventArgs) Handles btnRunQuery.Click
  Dim pQueries As SolidEdgeAssembly.Queries
  Dim pQQ As SolidEdgeAssembly.Query
  Dim pASMDoc As SolidEdgeAssembly.AssemblyDocument
  Dim strQ As String = ""
  Dim ii As Integer = 0

  strQ = cmbQueryFileName.Text

  

  If strQ = "" Then Exit Sub

  pASMDoc = CType(SEDocument, SolidEdgeAssembly.AssemblyDocument)

  If strQ = "<Reference>" Then
    'Select all references in the file
    Dim pOccs As SolidEdgeAssembly.Occurrences
    Dim pOcc As SolidEdgeAssembly.Occurrence

    pOccs = pASMDoc.Occurrences
    pASMDoc.SelectSet.SuspendDisplay()
    For Each pOcc In pOccs
        If pOcc.DisplayInSubAssembly = False Then
          pASMDoc.SelectSet.Add(pOcc)
        End If
    Next
    pASMDoc.SelectSet.ResumeDisplay()
    pASMDoc.SelectSet.RefreshDisplay()
    SEApplication.StartCommand(SolidEdgeFramework.SolidEdgeCommandConstants.seAssemblySelectCommand)

    Exit Sub
  ElseIf strQ = "<Design>" Then
    'Select all Design in the file
      Dim pOccs As SolidEdgeAssembly.Occurrences
      Dim pOcc As SolidEdgeAssembly.Occurrence

      pOccs = pASMDoc.Occurrences
      pASMDoc.SelectSet.SuspendDisplay()

      For Each pOcc In pOccs
          If pOcc.DisplayInSubAssembly = True Then
            pASMDoc.SelectSet.Add(pOcc)
          End If
      Next
      pASMDoc.SelectSet.ResumeDisplay()

      pASMDoc.SelectSet.RefreshDisplay()
      SEApplication.StartCommand(SolidEdgeFramework.SolidEdgeCommandConstants.seAssemblySelectCommand)
      Exit Sub

    ElseIf strQ = "<Not Selectable>" Then
    'Select all Not Selectable in the file
      Dim pOccs As SolidEdgeAssembly.Occurrences
      Dim pOcc As SolidEdgeAssembly.Occurrence

      pOccs = pASMDoc.Occurrences
      pASMDoc.SelectSet.SuspendDisplay()

      For Each pOcc In pOccs
          If pOcc.Locatable = False Then
            pASMDoc.SelectSet.Add(pOcc)
          End If
      Next
      pASMDoc.SelectSet.ResumeDisplay()

      pASMDoc.SelectSet.RefreshDisplay()
      SEApplication.StartCommand(SolidEdgeFramework.SolidEdgeCommandConstants.seAssemblySelectCommand)
      Exit Sub

  End If

  pQueries = pASMDoc.Queries

  pQQ = pQueries.QuickQuery
  pQQ.RemoveAllCriteria()

  pQQ.Scope = QueryScopeConstants.seQueryScopeAllParts

  pQQ.AddCriteria(QueryPropertyConstants.seQueryPropertyName, "", QueryConditionConstants.seQueryConditionContains, strQ)

  ii = pQQ.MatchesCount

  If ii > 0 Then
    Dim SS As SolidEdgeFramework.SelectSet = pASMDoc.SelectSet
    SS.RefreshDisplay()
    SEApplication.StartCommand(SolidEdgeFramework.SolidEdgeCommandConstants.seAssemblySelectCommand)

    UpdateComboList(strQ)
  End If


End Sub

Private Sub cmbQueryFileName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cmbQueryFileName.KeyPress
  If cmbQueryFileName.Text = "" Then Exit Sub

  If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
    btnRunQuery.PerformClick()
    e.Handled = True
  End If
End Sub

Private Sub UpdateComboList(strQ As String)

Dim tmpStr As String
Dim cmbCol As System.Windows.Forms.ComboBox.ObjectCollection = Me.cmbQueryFileName.Items

   For Each tmpStr In cmbCol
    If strQ.ToLower = tmpStr.ToLower Then Exit Sub
   Next
   'Add to the list
   Me.cmbQueryFileName.Items.Insert(0, strQ)
End Sub

    Private Sub ASMEdgebarCtrl_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
      If Me.Width < 257 Then
        Me.btnReplaceWithCopy.Left = 232
      Else
       Me.btnReplaceWithCopy.Left = Me.Width - 22
      End If
    End Sub
End Class


End Namespace
