Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Drawing

Public Class Form1
    Public arrayOfStringsToSearchForAndReplace As ArrayList
    Public arrayOfStringsReplace As ArrayList
    Public intNumberOfReplaces As Integer = 0

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        'because of .NET
        OleMessageFilter.Revoke()
    End Sub

    Private Function Search_Themes(ByVal sThemeName As String)
        Dim i As Integer = 0
        For i = 0 To objSEApp.Customization.RibbonBarThemes.Count - 1
            If (objSEApp.Customization.RibbonBarThemes.Item(i).Name = sThemeName) Then
                Search_Themes = True
                Exit Function
            End If
        Next

        Search_Themes = False

    End Function

    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ' Get the type from the Solid Edge ProgID
        objSEType = Type.GetTypeFromProgID("SolidEdge.Application")
        ' Get the type from the Revision Manager ProgID
        objRevManType = Type.GetTypeFromProgID("RevisionManager.Application")


        Me.Label3.Text = ""
        Me.Label3.Refresh()
        'because of .NET
        OleMessageFilter.Register()

        If oConnectToSolidEdge(True, True) Then
            If objSEApp.ActiveDocumentType <> SolidEdgeFramework.DocumentTypeConstants.igDraftDocument Then
                MessageBox.Show("You must run this macro from the Solid Edge draft environment", "Solid Edge Draft Search and Replace", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        End If

        '*****the code below will work for ST5 onward to put the UI on the draft UI
        'Dim strSEVersion As String = oGetSolidEdgeVersion()
        'Dim ar() As String = strSEVersion.Split(".")
        'strSEVersion = ar(0)


        'If CInt(strSEVersion) > 104 Then
        '    Dim objRibbonBarTheme As SolidEdgeFramework.RibbonBarTheme = Nothing
        '    Dim objNewRibbonBarTheme As SolidEdgeFramework.RibbonBarTheme = Nothing
        '    Dim objRibbonBar As SolidEdgeFramework.RibbonBar = Nothing
        '    Dim objNewTab As SolidEdgeFramework.RibbonBarTab = Nothing
        '    Dim objNewGroup As SolidEdgeFramework.RibbonBarGroup = Nothing
        '    Dim objNewControl As SolidEdgeFramework.RibbonBarControl = Nothing



        '    If oConnectToSolidEdge(True, True) Then
        '        Dim bFound As Boolean = False

        '        ' Check to see if they have a custom theme.  If they don't then create one.
        '        If (objSEApp.Customization.RibbonBarThemes.Count > 0) Then
        '            ' Check to see if this tab is on any ribbonbar on any of the themes.
        '            bFound = Search_Themes("My Custom SE Theme")

        '            ' If his theme is there then just continue 
        '            If (bFound) Then
        '                Exit Sub
        '            End If
        '        End If

        '        ' If there isn't a theme then create it.
        '        If (bFound = False) Then
        '            Try


        '                ' Find out where this executable is located.
        '                Dim aAsm As System.Reflection.Assembly = Nothing

        '                Try
        '                    aAsm = System.Reflection.Assembly.GetExecutingAssembly()
        '                Catch ex As Exception

        '                End Try


        '                ' Put this at the beginning  
        '                objSEApp.Customization.BeginCustomization()

        '                objNewRibbonBarTheme = objSEApp.Customization.RibbonBarThemes.Create(objRibbonBarTheme)
        '                objNewRibbonBarTheme.Name = "My Custom SE Theme"



        '                For i = 0 To objNewRibbonBarTheme.RibbonBars.Count - 1
        '                    If ("Detail" = objNewRibbonBarTheme.RibbonBars.Item(i).Environment) Then
        '                        objRibbonBar = objNewRibbonBarTheme.RibbonBars.Item(i)

        '                        ' Add a tab to the ribbon bar
        '                        objNewTab = objNewRibbonBarTheme.RibbonBars.Item(i).RibbonBarTabs.Insert("Extra Draft Tools", objNewRibbonBarTheme.RibbonBars.Item(i).RibbonBarTabs.Count, SolidEdgeFramework.RibbonBarInsertMode.seRibbonBarInsertCreate)
        '                        objNewTab.Visible = True

        '                        ' Add a group to the tab
        '                        objNewGroup = objNewTab.RibbonBarGroups.Insert("Tools", objNewRibbonBarTheme.RibbonBars.Item(i).RibbonBarTabs.Count, SolidEdgeFramework.RibbonBarInsertMode.seRibbonBarInsertCreate)
        '                        objNewGroup.Visible = True

        '                        Dim macroArray(1) As Object
        '                        macroArray(0) = aAsm.Location

        '                        objNewControl = objNewGroup.RibbonBarControls.Insert(macroArray, Nothing, SolidEdgeFramework.RibbonBarInsertMode.seRibbonBarInsertCreateButton) 'ST5 seRibbonBarInsertCreateButton
        '                        objNewControl.Visible = True

        '                        Exit For
        '                    End If
        '                Next

        '                ' Set Greg's theme 
        '                objSEApp.Customization.RibbonBarThemes.ActivateTheme(objNewRibbonBarTheme)
        '                objSEApp.Customization.RibbonBarThemes.Commit()
        '                objSEApp.Customization.EndCustomization()

        '            Catch ex As Exception

        '            End Try

        '        End If

        '    Else
        '        MessageBox.Show("Could not connect to or start Solid Edge!", "Label Dimensions", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '        End
        '    End If
        'End If



    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Dim objdraftDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheets As SolidEdgeDraft.Sheets = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim ii As Integer
        Dim strPulledString As String = String.Empty
        Dim jj As Integer = 0


        intNumberOfReplaces = 0


        If Me.TextBoxStrings.Text = "" Then
            MessageBox.Show("You must enter string to search for")
            Exit Sub
        End If

        If Me.TextBoxStringsReplaceWith.Text = "" Then
            MessageBox.Show("You must enter replacement string")
            Exit Sub
        End If

        'Add your code here!
        If oConnectToSolidEdge(True, True) Then

            Me.Label3.Text = "Searching the document properties...."
            Me.Label3.Refresh()

            arrayOfStringsToSearchForAndReplace = New ArrayList
            Dim strFullList As String = Me.TextBoxStrings.Text
            If strFullList.Contains(",") Then
                'need to splitup string
                Dim arr() As String = strFullList.Split(",")

                For jj = 0 To arr.Length - 1
                    arrayOfStringsToSearchForAndReplace.Add(arr(jj))
                Next
            Else
                arrayOfStringsToSearchForAndReplace.Add(strFullList)
            End If

            arrayOfStringsReplace = New ArrayList
            Dim strFullList1 As String = Me.TextBoxStringsReplaceWith.Text
            If strFullList.Contains(",") Then
                'need to splitup string
                Dim arr() As String = strFullList1.Split(",")

                For jj = 0 To arr.Length - 1
                    arrayOfStringsReplace.Add(arr(jj))
                Next
            Else
                arrayOfStringsReplace.Add(strFullList1)
            End If

            Try
                If objSEApp.ActiveDocumentType <> SolidEdgeFramework.DocumentTypeConstants.igDraftDocument Then
                    MessageBox.Show("The utility must be executed from the Solid Edge Draft environment.  Please open a draft file!")
                    OleMessageFilter.Revoke()
                    End
                End If

                objdraftDoc = objSEApp.ActiveDocument

                Dim blnPropertyChanged As Boolean = False

                'process properties of the document
                Dim objProperties As SolidEdgeFramework.PropertySets = Nothing
                Dim objProperty As SolidEdgeFramework.Properties = Nothing
                objProperties = objdraftDoc.Properties
                For Each objProperty In objProperties
                    For ii = 1 To objProperty.Count
                        Try
                            'strPulledString = objProperty.Item(ii).Name + " : " + objProperty.Item(ii).Value.ToString
                            If IsNothing(objProperty.Item(ii).Value) = False Then
                                strPulledString = objProperty.Item(ii).Value.ToString
                                For jj = 0 To arrayOfStringsToSearchForAndReplace.Count - 1
                                    If strPulledString.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                                        strPulledString = Regex.Replace(strPulledString, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                                        objProperty.Item(ii).Value = strPulledString
                                        intNumberOfReplaces = intNumberOfReplaces + 1
                                        blnPropertyChanged = True
                                    End If


                                Next jj
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Error processing file properties " + ex.Message)
                        End Try
                    Next
                Next

                If blnPropertyChanged = True Then
                    objProperties.Save()
                End If

                oReleaseObject(objProperty)
                oReleaseObject(objProperties)

                objSheets = objdraftDoc.Sheets
                For Each objSheet In objSheets

                    Me.Label3.Text = "Searching the sheet " + objSheet.Name
                    Me.Label3.Refresh()
                    'for each sheet check

                    'callouts
                    Me.Label3.Text = "Searching the callouts and ballons on " + objSheet.Name
                    Me.Label3.Refresh()
                    ProcessCalloutsOnSheet(objSheet)


                    'textboxes
                    Me.Label3.Text = "Searching the textboxes on " + objSheet.Name
                    Me.Label3.Refresh()
                    ProcessTextBoxesOnSheet(objSheet)


                    'dimensions ( prefix, suffix, etc)
                    Me.Label3.Text = "Searching the dimensions on " + objSheet.Name
                    Me.Label3.Refresh()
                    ProcessDimensionsOnSheet(objSheet)


                    'SurfaceFinishSymbols
                    Me.Label3.Text = "Searching the surface finish symbols on " + objSheet.Name
                    Me.Label3.Refresh()
                    ProcessSurfaceFinishSymbolsOnSheet(objSheet)

                    'DatumFrames
                    Me.Label3.Text = "Searching the datum frames on " + objSheet.Name
                    Me.Label3.Refresh()
                    ProcessDatumsOnSheet(objSheet)

                    'corner annotations
                    Me.Label3.Text = "Searching the corner annotations on " + objSheet.Name
                    Me.Label3.Refresh()
                    ProcessCornerAnnotationsOnSheet(objSheet)


                    'Embedded word docs
                    Me.Label3.Text = "Searching the embedded word docs on " + objSheet.Name
                    Me.Label3.Refresh()
                    ProcessEmbeddedWordDocsOnSheet(objSheet)

                Next


                objdraftDoc.UpdatePropertyTextCacheAndDisplay()
                objdraftDoc.UpdatePropertyTextDisplay()

                ' tables 
                Me.Label3.Text = "Searching the user defined tables "
                Me.Label3.Refresh()
                ProcessTablesInDocument(objdraftDoc)


                'partslist
                Me.Label3.Text = "Searching the PartLists "
                Me.Label3.Refresh()

            Catch ex As Exception
                MessageBox.Show("Error processing " + ex.Message)
            End Try


            oReleaseObject(objSheet)
            oReleaseObject(objSheets)
            oReleaseObject(objdraftDoc)
            oReleaseObject(objSEApp)
            oForceGarbageCollection()


            Me.Label3.Text = "Finished processing! Replaced " + intNumberOfReplaces.ToString + " occurrences of " + arrayOfStringsToSearchForAndReplace(0)
            Me.Label3.Refresh()




        End If
    End Sub


    Public Sub ProcessPartsListsInDocument(oDoc As SolidEdgeDraft.DraftDocument)
        Dim objPartsLists As SolidEdgeDraft.PartsLists = Nothing
        Dim objPartsList As SolidEdgeDraft.PartsList = Nothing
        objPartsLists = oDoc.PartsLists

        For Each objPartsList In objPartsLists
            'add code to process the columns in the partslist


        Next



    End Sub

    Public Sub ProcessTablesInDocument(oDoc As SolidEdgeDraft.DraftDocument)
        Dim objTables As SolidEdgeDraft.Tables = Nothing
        Dim objTable As SolidEdgeDraft.Table = Nothing
        objTables = oDoc.Tables
        
        Dim ii As Integer = 0
        Dim jj As Integer = 0

        Dim objTitles As SolidEdgeDraft.TableTitles = Nothing
        Dim objTitle As SolidEdgeDraft.TableTitle = Nothing


        Try
            For Each objTable In objTables

                Dim intNumberOfRows As Integer = objTable.Rows.Count
                Dim intNumberOfCols As Integer = objTable.Columns.Count

                objTitles = objTable.Titles
                For Each objTitle In objTitles
                    If IsNothing(objTitle.value) = False Then
                        Dim strTitle As String = objTitle.value
                        If strTitle.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(0).ToString.ToUpper) Then
                            strTitle = Regex.Replace(strTitle, arrayOfStringsToSearchForAndReplace(0), arrayOfStringsReplace(0), RegexOptions.IgnoreCase)
                            objTitle.value = strTitle
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If
                Next

                For ii = 1 To intNumberOfRows
                    For jj = 1 To intNumberOfCols
                        Dim strCellValue As String = String.Empty

                        If IsNothing(objTable.Cell(ii, jj).value) = False Then
                            strCellValue = objTable.Cell(ii, jj).value
                            If strCellValue.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(0).ToString.ToUpper) Then
                                strCellValue = Regex.Replace(strCellValue, arrayOfStringsToSearchForAndReplace(0), arrayOfStringsReplace(0), RegexOptions.IgnoreCase)
                                objTable.Cell(ii, jj).value = strCellValue
                                intNumberOfReplaces = intNumberOfReplaces + 1
                            End If
                        End If
                    Next jj
                Next ii
                objTable.Update()
            Next


            oReleaseObject(objTable)
            oReleaseObject(objTables)
            oForceGarbageCollection()

        Catch ex As Exception
            MessageBox.Show("Error processing user defined tables " + ex.Message)
        End Try



    End Sub


    Public Sub ProcessEmbeddedWordDocsOnSheet(oSheet As SolidEdgeDraft.Sheet)
        Dim objSmartFrames As SolidEdgeFrameworkSupport.SmartFrames2d = Nothing
        Dim objSmartFrame As SolidEdgeFrameworkSupport.SmartFrame2d = Nothing
        Dim jj As Integer = 0
        Dim strPulledCornerAnnotation As String = String.Empty


        Try
            objSmartFrames = oSheet.SmartFrames2d
            For Each objSmartFrame In objSmartFrames
                Dim oType As Microsoft.Office.Interop.Word.WdDocumentType = objSmartFrame.Object.type
                If oType = 0 Then
                    'check it....  it is a word document
                    Dim oWord As Microsoft.Office.Interop.Word.Application = Nothing
                    Dim WordDoc As Microsoft.Office.Interop.Word.Document = Nothing

                    objSmartFrame.DoVerb(-2) 'opens the word doc in Word
                    'now use WORD API do do search and replace
                    Try
                        oWord = GetObject(, "Word.Application")
                        For jj = 1 To oWord.Documents.Count
                            If oWord.Documents.Item(jj).FullName = objSmartFrame.Object.fullname Then
                                'got the correct word doc.
                                WordDoc = oWord.Documents.Item(jj)


                                'Find and replace some text   
                                'Replace 'VB' with 'Visual Basic'   
                                WordDoc.Content.Find.Execute(FindText:=arrayOfStringsToSearchForAndReplace(0), ReplaceWith:=arrayOfStringsReplace(0), Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                                While WordDoc.Content.Find.Execute(FindText:="  ", Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue)
                                    WordDoc.Content.Find.Execute(FindText:="  ", ReplaceWith:=" ", Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue)
                                End While

                            End If
                        Next jj

                    Catch ex As Exception

                    End Try


                    'now close the word doc.
                    WordDoc.Save()
                    WordDoc.Close()

                    If oWord.Documents.Count = 0 Then
                        oWord.Quit()
                    End If

                    oReleaseObject(oWord)
                    oReleaseObject(WordDoc)

                    objSmartFrame.Update()


                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error processing embedded word docs on " + oSheet.Name + " error is: " + ex.Message)
        End Try



    End Sub












    Public Sub ProcessCornerAnnotationsOnSheet(oSheet As SolidEdgeDraft.Sheet)
        Dim objCornerAnnotations As SolidEdgeFrameworkSupport.CornerAnnotations = Nothing
        Dim objCornerAnnotation As SolidEdgeFrameworkSupport.CornerAnnotation = Nothing
        Dim jj As Integer = 0
        Dim strPulledCornerAnnotation As String = String.Empty



        Try
            objCornerAnnotations = oSheet.CornerAnnotations
            For Each objCornerAnnotation In objCornerAnnotations
                For jj = 0 To arrayOfStringsToSearchForAndReplace.Count - 1
                    If IsNothing(objCornerAnnotation.Direction1Tolerance) = False Then
                        strPulledCornerAnnotation = objCornerAnnotation.Direction1Tolerance
                        If strPulledCornerAnnotation.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledCornerAnnotation = Regex.Replace(strPulledCornerAnnotation, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objCornerAnnotation.Direction1Tolerance = strPulledCornerAnnotation
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objCornerAnnotation.Direction2Tolerance) = False Then
                        strPulledCornerAnnotation = objCornerAnnotation.Direction2Tolerance
                        If strPulledCornerAnnotation.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledCornerAnnotation = Regex.Replace(strPulledCornerAnnotation, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objCornerAnnotation.Direction2Tolerance = strPulledCornerAnnotation
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objCornerAnnotation.UpperToleranceForEnhancedSymbol) = False Then
                        strPulledCornerAnnotation = objCornerAnnotation.UpperToleranceForEnhancedSymbol
                        If strPulledCornerAnnotation.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledCornerAnnotation = Regex.Replace(strPulledCornerAnnotation, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objCornerAnnotation.UpperToleranceForEnhancedSymbol = strPulledCornerAnnotation
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objCornerAnnotation.LowerToleranceForEnhancedSymbol) = False Then
                        strPulledCornerAnnotation = objCornerAnnotation.LowerToleranceForEnhancedSymbol
                        If strPulledCornerAnnotation.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledCornerAnnotation = Regex.Replace(strPulledCornerAnnotation, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objCornerAnnotation.LowerToleranceForEnhancedSymbol = strPulledCornerAnnotation
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                Next jj

            Next

            oReleaseObject(objCornerAnnotation)
            oReleaseObject(objCornerAnnotations)
            oForceGarbageCollection()

        Catch ex As Exception
            MessageBox.Show("Error processing corner annotations on " + oSheet.Name + " error is: " + ex.Message)
        End Try

    End Sub




    Public Sub ProcessDatumsOnSheet(oSheet As SolidEdgeDraft.Sheet)
        Dim objDatums As SolidEdgeFrameworkSupport.DatumFrames = Nothing
        Dim objDatum As SolidEdgeFrameworkSupport.DatumFrame = Nothing
        Dim jj As Integer = 0
        Dim strPulledDatum As String = String.Empty



        Try
            objDatums = oSheet.DatumFrames
            For Each objDatum In objDatums
                For jj = 0 To arrayOfStringsToSearchForAndReplace.Count - 1
                    If IsNothing(objDatum.Datum) = False Then
                        strPulledDatum = objDatum.Datum
                        If strPulledDatum.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledDatum = Regex.Replace(strPulledDatum, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objDatum.Datum = strPulledDatum
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                Next jj
            Next

            oReleaseObject(objDatum)
            oReleaseObject(objDatums)
            oForceGarbageCollection()

        Catch ex As Exception
            MessageBox.Show("Error processing Datum frames on " + oSheet.Name + " error is: " + ex.Message)
        End Try

    End Sub

    Public Sub ProcessSurfaceFinishSymbolsOnSheet(oSheet As SolidEdgeDraft.Sheet)
        Dim objSurfaceFinishes As SolidEdgeFrameworkSupport.SurfaceFinishSymbols = Nothing
        Dim objSurfaceFinish As SolidEdgeFrameworkSupport.SurfaceFinishSymbol = Nothing
        Dim jj As Integer = 0
        Dim strPulledSurfaceFinishText As String = String.Empty



        Try
            objSurfaceFinishes = oSheet.SurfaceFinishSymbols
            For Each objSurfaceFinish In objSurfaceFinishes
                For jj = 0 To arrayOfStringsToSearchForAndReplace.Count - 1
                    If IsNothing(objSurfaceFinish.MachiningAllowance) = False Then
                        strPulledSurfaceFinishText = objSurfaceFinish.MachiningAllowance
                        If strPulledSurfaceFinishText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledSurfaceFinishText = Regex.Replace(strPulledSurfaceFinishText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objSurfaceFinish.MachiningAllowance = strPulledSurfaceFinishText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objSurfaceFinish.ProductionMethod) = False Then
                        strPulledSurfaceFinishText = objSurfaceFinish.ProductionMethod
                        If strPulledSurfaceFinishText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledSurfaceFinishText = Regex.Replace(strPulledSurfaceFinishText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objSurfaceFinish.ProductionMethod = strPulledSurfaceFinishText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objSurfaceFinish.RoughnessValue) = False Then
                        strPulledSurfaceFinishText = objSurfaceFinish.RoughnessValue
                        If strPulledSurfaceFinishText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledSurfaceFinishText = Regex.Replace(strPulledSurfaceFinishText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objSurfaceFinish.RoughnessValue = strPulledSurfaceFinishText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objSurfaceFinish.Requirements) = False Then
                        strPulledSurfaceFinishText = objSurfaceFinish.Requirements
                        If strPulledSurfaceFinishText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledSurfaceFinishText = Regex.Replace(strPulledSurfaceFinishText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objSurfaceFinish.Requirements = strPulledSurfaceFinishText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If
                Next jj
            Next

            oReleaseObject(objSurfaceFinish)
            oReleaseObject(objSurfaceFinishes)
            oForceGarbageCollection()

        Catch ex As Exception
            MessageBox.Show("Error processing surface textures on " + oSheet.Name + " error is: " + ex.Message)
        End Try

    End Sub

    Public Sub ProcessDimensionsOnSheet(oSheet As SolidEdgeDraft.Sheet)

        Dim objDimensions As SolidEdgeFrameworkSupport.Dimensions = Nothing
        Dim objDimension As SolidEdgeFrameworkSupport.Dimension = Nothing
        Dim jj As Integer = 0
        Dim strPulledDimensionText As String = String.Empty


        Try
            objDimensions = oSheet.Dimensions
            For Each objDimension In objDimensions
                For jj = 0 To arrayOfStringsToSearchForAndReplace.Count - 1
                    If IsNothing(objDimension.OverrideString) = False Then
                        strPulledDimensionText = objDimension.OverrideString
                        If strPulledDimensionText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledDimensionText = Regex.Replace(strPulledDimensionText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objDimension.OverrideString = strPulledDimensionText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objDimension.PrefixString) = False Then
                        strPulledDimensionText = objDimension.PrefixString

                        If strPulledDimensionText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledDimensionText = Regex.Replace(strPulledDimensionText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objDimension.PrefixString = strPulledDimensionText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objDimension.SubfixString) = False Then
                        strPulledDimensionText = objDimension.SubfixString.ToString

                        If strPulledDimensionText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledDimensionText = Regex.Replace(strPulledDimensionText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objDimension.SubfixString = strPulledDimensionText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objDimension.SubfixString2) = False Then
                        strPulledDimensionText = objDimension.SubfixString2.ToString

                        If strPulledDimensionText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledDimensionText = Regex.Replace(strPulledDimensionText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objDimension.SubfixString2 = strPulledDimensionText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objDimension.SuffixString) = False Then
                        strPulledDimensionText = objDimension.SuffixString.ToString

                        If strPulledDimensionText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledDimensionText = Regex.Replace(strPulledDimensionText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objDimension.SuffixString = strPulledDimensionText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If
                    If IsNothing(objDimension.SuperfixString) = False Then
                        strPulledDimensionText = objDimension.SuperfixString.ToString

                        If strPulledDimensionText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledDimensionText = Regex.Replace(strPulledDimensionText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objDimension.SuperfixString = strPulledDimensionText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                Next
            Next
            oReleaseObject(objDimension)
            oReleaseObject(objDimensions)
            oForceGarbageCollection()
        Catch ex As Exception
            MessageBox.Show("Error processing dimensions on " + oSheet.Name + " error is: " + ex.Message)
        End Try

    End Sub


    Public Function CalculateWidthNeeded(oFont As String, oFontSize As Double, StrtoMeasure As String) As Double
        ' Determine the correct size for the text box based on its text length
        ' Create a new SizeF object to return the size into
        Dim mySize As New System.Drawing.SizeF

        ' Create a new font based on the font of the textbox we want to resize
        Dim myFont As New System.Drawing.Font(oFont, oFontSize)

        ' Get the size given the string and the font
        mySize = System.Windows.Forms.TextRenderer.MeasureText(StrtoMeasure, myFont)

        oReleaseObject(mySize)
        oReleaseObject(myFont)
        oForceGarbageCollection()
        Return CType(Math.Round(mySize.Width, 0), Integer)






    End Function


    Public Sub ProcessTextBoxesOnSheet(oSheet As SolidEdgeDraft.Sheet)

        Dim objTextBoxes As SolidEdgeFrameworkSupport.TextBoxes = Nothing
        Dim objTextBox As SolidEdgeFrameworkSupport.TextBox = Nothing
        Dim jj As Integer = 0
        Dim strPulledTextBoxText As String = String.Empty

        Try
            objTextBoxes = oSheet.TextBoxes

            For Each objTextBox In objTextBoxes

                For jj = 0 To arrayOfStringsToSearchForAndReplace.Count - 1
                    If IsNothing(objTextBox.Text) = False Then
                        strPulledTextBoxText = objTextBox.Text.ToString
                        If strPulledTextBoxText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            Dim Len1 As Double = Len(strPulledTextBoxText)
                            strPulledTextBoxText = Regex.Replace(strPulledTextBoxText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            Dim Len2 As Double = Len(strPulledTextBoxText)
                            If Len2 > Len1 Then
                                objTextBox.Width = objTextBox.Width + MMtoM(Math.Round((Len2 - Len1) / 2, 3))
                            End If
                            'Dim widthNeeded As Integer = CalculateWidthNeeded(objTextBox.Edit.Font, objTextBox.Edit.TextSize, strPulledTextBoxText)
                            'objTextBox.Width = widthNeeded / 1000
                            objTextBox.Text = strPulledTextBoxText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                Next
            Next

            oReleaseObject(objTextBox)
            oReleaseObject(objTextBoxes)
            oForceGarbageCollection()

        Catch ex As Exception
            MessageBox.Show("Error processing textboxes on " + oSheet.Name + " error is: " + ex.Message)
        End Try


    End Sub


    Public Sub ProcessCalloutsOnSheet(oSheet As SolidEdgeDraft.Sheet)


        Dim objCallouts As SolidEdgeFrameworkSupport.Balloons = Nothing
        Dim objCallout As SolidEdgeFrameworkSupport.Balloon = Nothing
        Dim strPulledCalloutText As String = String.Empty
        Dim jj As Integer = 0

        Try
            objCallouts = oSheet.Balloons
            For Each objCallout In objCallouts

                For jj = 0 To arrayOfStringsToSearchForAndReplace.Count - 1
                    If IsNothing(objCallout.BalloonText) = False Then
                        strPulledCalloutText = objCallout.BalloonText.ToString
                        If strPulledCalloutText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledCalloutText = Regex.Replace(strPulledCalloutText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objCallout.BalloonText = strPulledCalloutText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objCallout.BalloonTextLower) = False Then
                        strPulledCalloutText = objCallout.BalloonTextLower.ToString
                        If strPulledCalloutText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledCalloutText = Regex.Replace(strPulledCalloutText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objCallout.BalloonTextLower = strPulledCalloutText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If
                    If IsNothing(objCallout.BalloonTextPrefix) = False Then
                        strPulledCalloutText = objCallout.BalloonTextPrefix.ToString
                        If strPulledCalloutText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledCalloutText = Regex.Replace(strPulledCalloutText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objCallout.BalloonTextPrefix = strPulledCalloutText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                    If IsNothing(objCallout.BalloonTextSuffix) = False Then
                        strPulledCalloutText = objCallout.BalloonTextSuffix.ToString
                        If strPulledCalloutText.ToUpper.Contains(arrayOfStringsToSearchForAndReplace(jj).ToString.ToUpper) Then
                            strPulledCalloutText = Regex.Replace(strPulledCalloutText, arrayOfStringsToSearchForAndReplace(jj), arrayOfStringsReplace(jj), RegexOptions.IgnoreCase)
                            objCallout.BalloonTextSuffix = strPulledCalloutText
                            intNumberOfReplaces = intNumberOfReplaces + 1
                        End If
                    End If

                Next
            Next
            oReleaseObject(objCallout)
            oReleaseObject(objCallouts)
            oForceGarbageCollection()

        Catch ex As Exception
            MessageBox.Show("Error processing callouts on " + oSheet.Name + " error is: " + ex.Message)
        End Try




    End Sub




    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        OleMessageFilter.Revoke()
        End

    End Sub
End Class
