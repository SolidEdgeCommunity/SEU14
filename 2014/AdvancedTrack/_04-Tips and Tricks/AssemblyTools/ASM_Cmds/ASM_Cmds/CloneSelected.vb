
Namespace SolidEdge.ASM_Edgebar_Cmds

Module CloneSelected

#Region "Data"
Public m_ExcelNameList As clsExcelData = Nothing
Private m_SEApp As SolidEdgeFramework.Application
Private m_SelectSet As SolidEdgeFramework.SelectSet
Private m_ActiveASMDocument As SolidEdgeAssembly.AssemblyDocument

#End Region

Public Function CloneSelectedComponents(SEApp As SolidEdgeFramework.Application) As Boolean
			' Get the Solid Edge Application and the select set
			Try
					m_SEApp = SEApp
			Catch ex As Exception
					m_SEApp = Nothing
					GoTo NO_DATA
			End Try

			Try
					m_ActiveASMDocument = DirectCast(m_SEApp.ActiveDocument, SolidEdgeAssembly.AssemblyDocument)
			Catch ex As Exception
					m_ActiveASMDocument = Nothing
					GoTo NO_DATA
			End Try

			' Get the active select set
			Try
					m_SelectSet = DirectCast(m_SEApp.ActiveSelectSet, SolidEdgeFramework.SelectSet)
			Catch ex As Exception
					m_SelectSet = Nothing
					MsgBox("Cannot connect to Solid Edge")
					GoTo NO_DATA
			End Try

			PerformCloneAction()
NO_DATA:
			m_SEApp = Nothing
			m_ActiveASMDocument = Nothing
			m_SelectSet = Nothing
			Return False

	End Function


Private Sub PerformCloneAction()

	 'See what we have selected
	 Dim ii As Integer
	 Dim pOccurrence As SolidEdgeAssembly.Occurrence = Nothing
	 Dim pSubOccurrence As SolidEdgeAssembly.SubOccurrence = Nothing
	 Dim strSourceFile As String = ""
	 Dim strNewFileName As String = ""
	 Dim strTargetFolder As String = ""
	 Dim InvalidSelection As Boolean = False
	 Dim ASM_Doc As SolidEdgeAssembly.AssemblyDocument = Nothing
	 Dim Part_Doc As SolidEdgePart.PartDocument = Nothing
	 Dim SM_Doc As SolidEdgePart.SheetMetalDocument = Nothing
	 Dim obj As Object = Nothing
	 Dim ex As Exception

	 For ii = 1 To m_SelectSet.Count

		strSourceFile = ""
		strNewFileName = ""
		strTargetFolder = ""

			Try
					pOccurrence = DirectCast(m_SelectSet.Item(ii), SolidEdgeAssembly.Occurrence)
			Catch ex
					'See if its a component in a sub-ASM (SE Reference)
					Try
							Dim pReference As SolidEdgeFramework.Reference
							pReference = DirectCast(m_SelectSet.Item(ii), SolidEdgeFramework.Reference)

							pOccurrence = DirectCast(pReference.Object, SolidEdgeAssembly.Occurrence)
					Catch ex
							Exit Sub
					End Try
			End Try

			If pOccurrence Is Nothing Then Exit Sub

			'Check if it's an occurrence we can act on.
			If pOccurrence.IsCopy Or pOccurrence.IsPatternItem Or pOccurrence.IsPipeSegment Or _
					pOccurrence.IsPipeFitting Or pOccurrence.IsWire Or pOccurrence.IsTube Or pOccurrence.IsStructuralFrameItem Or _
					pOccurrence.HasBodyOverride _
			Then
				InvalidSelection = True
			Else
				'Process it
					'See if it's a part or sub-ASM
					If pOccurrence.Subassembly Then
							Try
								ASM_Doc = DirectCast(pOccurrence.OccurrenceDocument, SolidEdgeAssembly.AssemblyDocument)

                strTargetFolder = ASM_Doc.Path
                strNewFileName = m_ExcelNameList.GetNewFileName(ASM_Doc.Name, ASM_Doc.Type)
                'Check if the file we are trying to make already exists in the folder we are working in
                'If so get another name.
                While FileExists(strTargetFolder + "\" + strNewFileName)
                  strNewFileName = m_ExcelNameList.GetNewFileName(ASM_Doc.Name, ASM_Doc.Type)
                End While

                ReplaceComponent(pOccurrence, strTargetFolder, strNewFileName)

              Catch ex
                ASM_Doc = Nothing
              End Try

					Else 'See if it's a part or SM file
						Try
							Part_Doc = DirectCast(pOccurrence.OccurrenceDocument, SolidEdgePart.PartDocument)
							If Part_Doc.HardwareFile Then
								InvalidSelection = True
							End If

              strTargetFolder = Part_Doc.Path
							strNewFileName = m_ExcelNameList.GetNewFileName(Part_Doc.Name, Part_Doc.Type)
              'Check if the file we are trying to make already exists in the folder we are working in
              'If so get another name.
               While FileExists(strTargetFolder + "\" + strNewFileName)
                 strNewFileName = m_ExcelNameList.GetNewFileName(Part_Doc.Name, Part_Doc.Type)
               End While

               ReplaceComponent(pOccurrence, strTargetFolder, strNewFileName)

              Try
                SM_Doc = DirectCast(pOccurrence.OccurrenceDocument, SolidEdgePart.SheetMetalDocument)
                If SM_Doc.HardwareFile Then Exit Sub

                strTargetFolder = SM_Doc.Path
                strNewFileName = m_ExcelNameList.GetNewFileName(SM_Doc.Name, SM_Doc.Type)
                'Check if the file we are trying to make already exists in the folder we are working in
                'If so get another name.
                While FileExists(strTargetFolder + "\" + strNewFileName)
                  strNewFileName = m_ExcelNameList.GetNewFileName(SM_Doc.Name, SM_Doc.Type)
                End While

                ReplaceComponent(pOccurrence, strTargetFolder, strNewFileName)

              Catch ex
                SM_Doc = Nothing
              End Try
						Catch ex
							Part_Doc = Nothing
						End Try
					End If
      End If

	 Next ii

If InvalidSelection Then
	If m_SelectSet.Count > 1 Then
	 MsgBox("Some components cannot be replaced because they are:" & vbNewLine & "Copy, Pattern, Pipe, Pipe Fitting, Wire, Frame, Hardware Part or has Assembly Features")
	Else
	 MsgBox("The selected component cannot be replaced because it is a:" & vbNewLine & "Copy, Pattern, Pipe, Pipe Fitting, Wire, Frame, Hardware Part or has Assembly Features")
	End If

End If


End Sub

Private Sub ReplaceComponent(pOccurrence As SolidEdgeAssembly.Occurrence, strTargetFolder As String, strNewFileName As String)
Dim ex As Exception
Dim strSourceFile As String
    Try
            'Copy the file
            strSourceFile = pOccurrence.OccurrenceFileName
            If strSourceFile <> "" And strTargetFolder <> "" And strNewFileName <> "" Then
                CopyAndRename(strSourceFile, strTargetFolder + "\" + strNewFileName)
            End If

            'Get the ASM the occurrence resides in
            Dim TargetASMDocument As SolidEdgeAssembly.AssemblyDocument

            TargetASMDocument = pOccurrence.TopLevelDocument


            Dim arrTarget_Components(1) As SolidEdgeAssembly.Occurrence
            arrTarget_Components(1) = pOccurrence

            'Replace the file
            TargetASMDocument.ReplaceComponents(arrTarget_Components, strTargetFolder + "\" + strNewFileName, _
                                                SolidEdgeAssembly.ConstraintReplacementConstants.seConstraintReplacementDelete)
    Catch ex
      MsgBox(ex.Message, "ReplaceComponent")
    End Try

End Sub



Private Function FileExists(strFname) As Boolean
  If My.Computer.FileSystem.FileExists(strFname) Then
    Return True
  End If
Return False
End Function


Private Sub CopyAndRename(ByVal strSource As String, ByVal strTarget As String)
    My.Computer.FileSystem.CopyFile(strSource, strTarget, False)
End Sub

End Module

End Namespace
