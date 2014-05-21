Imports Microsoft.Win32
Imports SolidEdge.CommonUI
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text

Namespace SolidEdge.ASM_Edgebar_Cmds

<Guid("59903533-59FB-4DDB-88CD-60963F016E61"), ProgId("SolidEdge.ASM_Edgebar_Cmds"), _
			ComVisible(True), AddInInfo("SolidEdge.ASM_Edgebar_Cmds", "Solid Edge ASM_Edgebar_Cmds Addin in .NET 4.5.", True), _
			AddInEnvironmentCategory(CategoryIDs.CATID_SEAssembly)> _
Public Class clsASM_Edgebar_Cmds
		Implements SolidEdgeFramework.ISolidEdgeAddIn
	Private _application As SolidEdgeFramework.Application
		Private _addInEx As SolidEdgeFramework.ISEAddInEx
		Private _connectionPointDictionary As New Dictionary(Of IConnectionPoint, Integer)()
		Private _resourceAssembly As System.Reflection.Assembly
		Private _edgeBarController As EdgeBarController

#Region "SolidEdgeFramework.ISolidEdgeAddIn"

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISolidEdgeAddIn.OnConnection().
		''' </summary>
		Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As SolidEdgeFramework.SeConnectMode, ByVal AddInInstance As SolidEdgeFramework.AddIn) Implements SolidEdgeFramework.ISolidEdgeAddIn.OnConnection
			_application = DirectCast(Application, SolidEdgeFramework.Application)
			_addInEx = DirectCast(AddInInstance, SolidEdgeFramework.ISEAddInEx)
			_resourceAssembly = System.Reflection.Assembly.GetExecutingAssembly()
			_edgeBarController = New EdgeBarController(_addInEx, _resourceAssembly)

			' Handle ConnectMode if necessary.
			Select Case ConnectMode
				Case SolidEdgeFramework.SeConnectMode.seConnectAtStartup
				Case SolidEdgeFramework.SeConnectMode.seConnectByUser
				Case SolidEdgeFramework.SeConnectMode.seConnectExternally
			End Select
		End Sub

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISolidEdgeAddIn.OnConnectToEnvironment().
		''' </summary>
		Public Sub OnConnectToEnvironment(ByVal EnvCatID As String, ByVal pEnvironmentDispatch As Object, ByVal bFirstTime As Boolean) Implements SolidEdgeFramework.ISolidEdgeAddIn.OnConnectToEnvironment
			' You can convert the environment CATID string to a .NET Guid object. Easier to compare and work with.
			Dim envGuid As New Guid(EnvCatID)
			' Cast pEnvironmentDispatch as a strongly typed variable.
			Dim environment As SolidEdgeFramework.Environment = DirectCast(pEnvironmentDispatch, SolidEdgeFramework.Environment)

			' Demonstrate working with CategoryIDs.
			If envGuid.Equals(CategoryIDs.SEAssembly) Then
					'Need to init the edgebar control
					_edgeBarController.AddPage(_application.ActiveDocument)
			End If

			' Some things only need to be done during bFirstTime.
			If bFirstTime Then

			End If

		End Sub

		''' <summary>
		''' Implementation of SolidEdgeFramework.ISolidEdgeAddIn.OnDisconnection().
		''' </summary>
		Public Sub OnDisconnection(ByVal DisconnectMode As SolidEdgeFramework.SeDisconnectMode) Implements SolidEdgeFramework.ISolidEdgeAddIn.OnDisconnection
			' Handle DisconnectMode if necessary.
			Select Case DisconnectMode
				Case SolidEdgeFramework.SeDisconnectMode.seDisconnectAtShutdown
				Case SolidEdgeFramework.SeDisconnectMode.seDisconnectByUser
				Case SolidEdgeFramework.SeDisconnectMode.seDisconnectExternally
			End Select

      ' Cleanup.
      _edgeBarController.Dispose()
      _edgeBarController = Nothing

      'Close the Excel Spreadsheet if open here.
      'The xls is global to the addin (open once used everywhere)

       If Not IsNothing(m_ExcelNameList) Then
        m_ExcelNameList.ShutDown()
        m_ExcelNameList = Nothing
      End If

		End Sub

#End Region

#Region "regasm.exe"

		''' <summary>
		''' Implementation of ComRegisterFunction.
		''' </summary>
		''' <remarks>
		''' This method gets called when regasm.exe is executed against the assembly.
		''' </remarks>
		<ComRegisterFunction> _
		Public Shared Sub Register(ByVal t As Type)
			RegistrationHelper.Register(t)
		End Sub

		''' <summary>
		''' Implementation of ComUnregisterFunction.
		''' </summary>
		''' <remarks>
		''' This method gets called when regasm.exe is executed against the assembly.
		''' </remarks>
		<ComUnregisterFunction> _
		Public Shared Sub Unregister(ByVal t As Type)
			RegistrationHelper.Unregister(t)
		End Sub

#End Region

#Region "Properties"

		Public ReadOnly Property Application() As SolidEdgeFramework.Application
			Get
				Return _application
			End Get
		End Property
		Public ReadOnly Property AddIn() As SolidEdgeFramework.ISEAddInEx
			Get
				Return _addInEx
			End Get
		End Property
		Public ReadOnly Property ResourceAssembly() As System.Reflection.Assembly
			Get
				Return _resourceAssembly
			End Get
		End Property

#End Region
End Class


End Namespace
