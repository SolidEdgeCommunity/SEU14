Imports System
Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Public Class Form1
   
    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Try
            ' Get the type from the Solid Edge ProgID
            objSEType = Type.GetTypeFromProgID("SolidEdge.Application")
            ' Get the type from the Revision Manager ProgID
            objRevManType = Type.GetTypeFromProgID("RevisionManager.Application")

        Catch ex As Exception
            MessageBox.Show("Error getting applicaton type IDs " + ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        '.NET must do
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, True) Then
            If objSEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igDraftDocument Then
                Dim objSelectSet As SolidEdgeFramework.SelectSet = Nothing
                objSelectSet = objSEApp.ActiveDocument.selectset
                If objSelectSet.Count > 0 Then
                    Dim ii As Integer = 0
                    For ii = 1 To objSelectSet.Count
                        Dim objDimSelected As SolidEdgeFrameworkSupport.Dimension = Nothing
                        objDimSelected = objSelectSet.Item(ii)
                        If Me.RadioButtonCritical.Checked = True Then
                            AttachSymbolToDimension(objDimSelected, "Critical")
                        ElseIf Me.RadioButtonMajor.Checked = True Then
                            AttachSymbolToDimension(objDimSelected, "Major")
                        ElseIf Me.RadioButtonMinor.Checked = True Then
                            AttachSymbolToDimension(objDimSelected, "Minor")
                        End If
                    Next
                Else
                    MessageBox.Show("Please select the dimensions to mark as critical, major or minor!", "Label Dimensions", MessageBoxButtons.OK, _
                                    MessageBoxIcon.Exclamation)
                End If
            Else
                MessageBox.Show("You must run this macro from Solid Edge draft!", "Label Dimensions", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            MessageBox.Show("Could not connect to or start Solid Edge!", "Label Dimensions", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If

        '.NET must do
        OleMessageFilter.Revoke()
    End Sub

    Public Sub AttachSymbolToDimension(oDim As SolidEdgeFrameworkSupport.Dimension, strDimType As String)
        Dim objBalloon As SolidEdgeFrameworkSupport.Balloon = Nothing
        Dim objBalloons As SolidEdgeFrameworkSupport.Balloons = Nothing
        Dim objActiveDraftSheet As SolidEdgeDraft.Sheet = Nothing
        Dim oDimDisplayData As SolidEdgeFrameworkSupport.DisplayData = Nothing
        Dim OriginX, OriginY, OriginZ, X_DirX, X_DirY, X_DirZ, Z_DirX, Z_DirY, Z_DirZ As Double
        Dim oText As String = ""
        oDimDisplayData = oDim.GetDisplayData
        Try
            objActiveDraftSheet = objSEApp.ActiveDocument.activesheet
            objBalloons = objActiveDraftSheet.Balloons
            oDimDisplayData = oDim.GetDisplayData
            'index of zero is the dim text
            oDimDisplayData.GetTextAtIndex(0, oText, OriginX, OriginY, OriginZ, X_DirX, X_DirY, X_DirZ, Z_DirX, Z_DirY, Z_DirZ)
            Select Case strDimType
                Case "Critical"  'place a filled in diamond
                    objBalloon = objBalloons.AddByTerminator(oDim, OriginX - 0.002, OriginY + 0.002, 0, True)
                    objBalloon.Callout = True
                    objBalloon.Style.Font = "Arial"
                    objBalloon.BalloonText = ChrW(&H2666)
                Case "Major"  'place a filled in square
                    objBalloon = objBalloons.AddByTerminator(oDim, OriginX - 0.002, OriginY + 0.002, 0, True)
                    objBalloon.Callout = True
                    objBalloon.Style.Font = "Arial"
                    objBalloon.BalloonText = ChrW(&H25A0)  '"U+25A0"
                Case "Minor"  'place a filled in circle
                    objBalloon = objBalloons.AddByTerminator(oDim, OriginX - 0.002, OriginY + 0.002, 0, True)
                    objBalloon.Callout = True
                    objBalloon.Style.Font = "Arial"
                    objBalloon.BalloonText = ChrW(&H25CF)  ' "U+25CF"
            End Select
        Catch ex As Exception
            MessageBox.Show("Error placing callout to identify dimension.  the error is " + ex.Message, "Label Dimensions", _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'release stuff....  since a local variable when goes out of scope should be released but good practice to do it anyway!
        oRelease_Object(oDimDisplayData)
        oRelease_Object(objBalloon)
        oRelease_Object(objBalloons)
        oRelease_Object(objActiveDraftSheet)
        ForceGarbageCollection()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Form1_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        '.NET must do!
        OleMessageFilter.Revoke()


        releaseObject(objSEApp)
        ForceGarbageCollection()

    End Sub
End Class
