Imports System.Runtime.InteropServices
Public Class Form1
    Private Sub btnStartSE_Click(sender As Object, e As EventArgs) Handles btnStartSE.Click
        Dim objApplication As SolidEdgeFramework.Application = Nothing
        Dim objPartDocument As SolidEdgePart.PartDocument = Nothing
        Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing

        Try
            If chkboxNewSession.Checked = False Then
                objApplication = Marshal.GetActiveObject("SolidEdge.Application")
            Else
                objApplication = Marshal.GetActiveObject("SolidEdge.Application")
                objApplication.Visible = True
            End If

            If radbuttonPart.Checked = True Then
                objPartDocument = objApplication.Documents.Add("SolidEdge.PartDocument")
            End If

            If radbuttonAssembly.Checked = True Then
                objAssemblyDocument = objApplication.Documents.Add("SolidEdge.AssemblyDocument")
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            objApplication = Nothing
            objPartDocument = Nothing
            objAssemblyDocument = Nothing

        End Try
    End Sub
End Class
