Imports SolidEdgeContrib
Imports SolidEdgeContrib.Extensions
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Runtime.Remoting
Imports System.Text

Namespace Example2
	Public Class InteropProxy
		Inherits MarshalByRefObject

		''' <summary>
		''' Executes interop code in an isolated AppDomain.
		''' </summary>
		''' <remarks>
		''' Notice that we do not have to worry about RCW's. i.e. Marshal.ReleaseComObject.
		''' </remarks>
		Public Sub DoIsolatedTask(ByVal applicationTransparentProxy As SolidEdgeFramework.Application)
			' See what AppDomain we're currently executing in.
			Dim currentAppDomain = AppDomain.CurrentDomain

			' This will never be the default AppDomain at this point.
			Dim isDefaultAppDomain = currentAppDomain.IsDefaultAppDomain()

			' Register with OLE to handle concurrency issues on the current thread.
			OleMessageFilter.Register()

			' RCW's cross AppDomains as TransparentProxies.
			' Unwrap the TransparentProxy object that was passed into this AppDomain.
			Dim application = UnwrapTransparentProxy(Of SolidEdgeFramework.Application)(applicationTransparentProxy)
			Dim documents = application.Documents

			' Add a new part document.
			Dim partDocument = documents.AddPartDocument()

			' Always a good idea to give SE a chance to breathe.
			application.DoIdle()

			' Optional performance improvement tweaks.
			application.DelayCompute = True
			application.ScreenUpdating = False

			' Create a polygon in the part document.
			CreatePolygon(partDocument)

			' Undo performance improvement tweaks.
			application.DelayCompute = False
			application.ScreenUpdating = True

			' Register with OLE to handle concurrency issues on the current thread.
			OleMessageFilter.Unregister()
		End Sub

		Private Sub CreatePolygon(ByVal partDocument As SolidEdgePart.PartDocument)
			' Get a reference to the Application object.
			Dim application = partDocument.Application

			' Get a reference to the RefPlanes collection.
			Dim refPlanes = partDocument.RefPlanes

			' Get a reference to the top RefPlane using extension method.
			Dim refPlane = refPlanes.GetTopPlane()

			' Get a reference to the ProfileSets collection.
			Dim profileSets = partDocument.ProfileSets

			' Add a new ProfileSet.
			Dim profileSet = profileSets.Add()

			' Get a reference to the Profiles collection.
			Dim profiles = profileSet.Profiles

			' Add a new Profile.
			Dim profile = profiles.Add(refPlane)

			' Get a reference to the Relations2d collection.
			Dim relations2d = DirectCast(profile.Relations2d, SolidEdgeFrameworkSupport.Relations2d)

			' Get a reference to the Lines2d collection.
			Dim lines2d = profile.Lines2d

			Dim sides As Integer = 8
			Dim angle As Double = 360 \ sides
			angle = (angle * Math.PI) / 180

			Dim radius As Double =.05
			Dim lineLength As Double = 2 * radius * (Math.Tan(angle) / 2)

			' x1, y1, x2, y2
			Dim points() As Double = { 0.0, 0.0, 0.0, 0.0 }

			Dim x As Double = 0.0
			Dim y As Double = 0.0

			points(2) = -((Math.Cos(angle / 2) * radius) - x)
			points(3) = -((lineLength / 2) - y)

			' Draw each line.
			For i As Integer = 0 To sides - 1
				points(0) = points(2)
				points(1) = points(3)
				points(2) = points(0) + (Math.Sin(angle * i) * lineLength)
				points(3) = points(1) + (Math.Cos(angle * i) * lineLength)

				lines2d.AddBy2Points(points(0), points(1), points(2), points(3))
			Next i

			' Create endpoint relationships.
			For i As Integer = 1 To lines2d.Count
				If i = lines2d.Count Then
					relations2d.AddKeypoint(lines2d.Item(i), CInt(SolidEdgeConstants.KeypointIndexConstants.igLineEnd), lines2d.Item(1), CInt(SolidEdgeConstants.KeypointIndexConstants.igLineStart))
				Else
					relations2d.AddKeypoint(lines2d.Item(i), CInt(SolidEdgeConstants.KeypointIndexConstants.igLineEnd), lines2d.Item(i + 1), CInt(SolidEdgeConstants.KeypointIndexConstants.igLineStart))
					relations2d.AddEqual(lines2d.Item(i), lines2d.Item(i + 1))
				End If
			Next i

			' Get a reference to the ActiveSelectSet.
			Dim selectSet = application.ActiveSelectSet

			' Empty ActiveSelectSet.
			selectSet.RemoveAll()

			' Add all lines to ActiveSelectSet.
			For i As Integer = 1 To lines2d.Count
				selectSet.Add(lines2d.Item(i))
			Next i

			' Switch to ISO view.
			application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewISOView)
		End Sub

		Private Function UnwrapTransparentProxy(Of T As Class)(ByVal rcw As Object) As T
			If RemotingServices.IsTransparentProxy(rcw) Then
				Dim punk As IntPtr = Marshal.GetIUnknownForObject(rcw)

				Try
					Return DirectCast(Marshal.GetObjectForIUnknown(punk), T)
				Finally
					Marshal.Release(punk)
				End Try
			End If

			Return DirectCast(rcw, T)
		End Function
	End Class
End Namespace
