Imports Microsoft.Win32
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace SolidEdge.ASM_Edgebar_Cmds
	''' <summary>
	''' Helper class for ComRegisterFunction and ComUnregisterFunction attributes.
	''' </summary>
	Public NotInheritable Class RegistrationHelper

		Private Sub New()
		End Sub

		Public Shared Sub Register(ByVal t As Type)
			Dim addinInfo = DirectCast(AddInInfoAttribute.GetCustomAttribute(t, GetType(AddInInfoAttribute)), AddInInfoAttribute)
			Dim environments = CType(AddInEnvironmentCategoryAttribute.GetCustomAttributes(t, GetType(AddInEnvironmentCategoryAttribute)), AddInEnvironmentCategoryAttribute())

			If addinInfo Is Nothing Then
				Throw New System.Exception("Missing AddInInfoAttribute.")
			End If
			If (environments Is Nothing) OrElse (environments.Length = 0) Then
				Throw New System.Exception("Missing AddInEnvironmentCategoryAttribute.")
			End If

			Dim subkey As String = String.Format("CLSID\{0}", t.GUID.ToString("B"))
			Using baseKey As RegistryKey = Registry.ClassesRoot.CreateSubKey(subkey)
				subkey = String.Format("Implemented Categories\{0}", CategoryIDs.CATID_SolidEdgeAddIn)
				Using implementedCategoriesKey As RegistryKey = baseKey.CreateSubKey(subkey)
				End Using

				For Each environment In environments
					subkey = String.Format("Environment Categories\{0}", environment.Guid.ToString("B"))
					Using environmentCategoryKey As RegistryKey = baseKey.CreateSubKey(subkey)
					End Using
				Next environment

				Using summaryKey As RegistryKey = baseKey.CreateSubKey("Summary")
					summaryKey.SetValue("409", addinInfo.Summary)
				End Using

				baseKey.SetValue("AutoConnect", If(addinInfo.AutoConnect, 1, 0))
				baseKey.SetValue("409", addinInfo.Title)
			End Using
		End Sub

		Public Shared Sub Unregister(ByVal t As Type)
			Dim subkey As String = String.Format("CLSID\{0}", t.GUID.ToString("B"))
			Registry.ClassesRoot.DeleteSubKeyTree(subkey, False)
		End Sub
	End Class
End Namespace
