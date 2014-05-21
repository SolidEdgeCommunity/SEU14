Imports System.Runtime.InteropServices
Module SolidEdge


    Public objSEApp As SolidEdgeFramework.Application = Nothing
    Public objSEType As Type = Nothing
    Public objRevManType As Type = Nothing
    Public Const InToMeter As Double = 0.0254
    Public Const PI As Double = 3.14159265359

    Public arrayofBadDimensions As ArrayList


    ' Define constants for various edge environments commonly used by add-ins. More
    ' GUIDs exist in the edge\sdk\include\secatids.h file.
    Public Const CATID_SolidEdgeAddIn As String = "{26B1D2D1-2B03-11d2-B589-080036E8B802}"

    Public Const CATID_SEApplication As String = "{26618394-09D6-11d1-BA07-080036230602}"

    ' Primary document based environments.
    Public Const CATID_SEPart As String = "{26618396-09D6-11d1-BA07-080036230602}"
    Public Const CATID_SESyncPart As String = "{D9B0BB85-3A6C-4086-A0BB-88A1AAD57A58}"
    Public Const CATID_SEAssembly As String = "{26618395-09D6-11d1-BA07-080036230602}"
    Public Const CATID_SESyncAssembly As String = "{2C3C2A72-3A4A-471d-98B5-E3A8CFA4A2BF}"
    Public Const CATID_SESheetMetal As String = "{26618398-09D6-11D1-BA07-080036230602}"
    Public Const CATID_SESyncSheetMetal As String = "{9CBF2809-FF80-4dbc-98F2-B82DABF3530F}"
    Public Const CATID_SEDraft As String = "{08244193-B78D-11D2-9216-00C04F79BE98}"
    Public Const CATID_SEWeldment As String = "{7313526A-276F-11D4-B64E-00C04F79B2BF}"

    ' Environments accessible via a primary document environment

    ' Sketch is a catch-all for the legacy "profile, profile hole, profile pattern, layout" etc. "layout" was sketch in assembly.
    Public Const CATID_SESketch As String = "{0DDABC90-125E-4cfe-9CB7-DC97FB74CCF4}"
    Public Const CATID_FEAResultsPart As String = "{B5965D1C-8819-4902-8252-64841537A16C}"

    Public Const CATID_FEAResultsAssembly As String = "{986B2512-3AE9-4a57-8513-1D2A1E3520DD}"
    Public Const CATID_SEXpresRoute As String = "{1661432A-489C-4714-B1B2-61E85CFD0B71}"
    Public Const CATID_SEHarness As String = "{5337A0AB-23ED-4261-A238-00E2070406FC}"
    Public Const CATID_SEFrame As String = "{D84119E8-F844-4823-B3A0-D4F31793028A}"

    Public Const CATID_SE2DModel As String = "{F6031120-7D99-48a7-95FC-EEE8038D7996}"
    Public Const CATID_SEDrawingViewEdit As String = "{8DBC3B5F-02D6-4241-BE96-B12EAF83FAE6}"


    ' Use this if you want the add-in to load in every environment (including application/no document env)
    Public Const CATID_SEAll As String = "{C484ED57-DBB6-4a83-BEDB-C08600AF07BF}"

    ' Use this if you want the add-in to load in every environment that has a document (excludes application/no document env)
    Public Const CATID_SEAllDocumentEnvrionments = "{BAD41B8D-18FF-42c9-9611-8A00E6921AE8}"

    ' A few of the less often used are in the sdk\include\secatids.h file. Feel free to add your own.

    Public Function GetUniqueRCW(Obj As Object) As Object

        Try
            Dim ip As IntPtr

            ip = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(Obj)

            GetUniqueRCW = System.Runtime.InteropServices.Marshal.GetUniqueObjectForIUnknown(ip)

            System.Runtime.InteropServices.Marshal.Release(ip)

        Catch ex As Exception
            GetUniqueRCW = Nothing
        End Try
    End Function

    Sub oReleaseObject(ByVal obj As Object)
        Try
            If Not (obj Is Nothing) Then
                'this should only be used when programming applications that run in their own process space
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    Public Function oConnectToSolidEdge(ByVal blnAppVisibility As Boolean, ByVal blnDisplayAlerts As Boolean, Optional ByVal oStopSE As String = "NO") As Boolean
        If oStopSE.ToLower = "no" Then
            'to connect to a running instance of Solid Edge
            Try
                objSEApp = Marshal.GetActiveObject("SolidEdge.Application")
                objSEApp.DisplayAlerts = blnDisplayAlerts
                objSEApp.Visible = blnAppVisibility
                Return True

            Catch ex As System.Exception
                Return False
            End Try
        End If
        Return False
    End Function


    Public Sub ProcessDraftDoc(objDraft As SolidEdgeDraft.DraftDocument)

        Dim oSheets As SolidEdgeDraft.Sheets = Nothing
        Dim oSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objActiveSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objDimensions As SolidEdgeFrameworkSupport.Dimensions = Nothing
        Dim objDimension As SolidEdgeFrameworkSupport.Dimension = Nothing
        Dim dblDimValue As Double = 0
        Dim objDimensionStyle As SolidEdgeFrameworkSupport.DimStyle = Nothing
        Dim NumberofDecimalDigits As SolidEdgeFrameworkSupport.DimRoundOffTypeConstants
        Dim oUnitOfDimension As SolidEdgeFrameworkSupport.DimLinearUnitConstants
        Dim intNumberOfDecimalPlacesShownInDimension As Integer = 0
        Dim dblConversionFactor As Double = 0
        Dim blnWillChange As Boolean = False





        Try
            objActiveSheet = objDraft.ActiveSheet
            oSheets = objDraft.Sheets
            'check for dimensions on each sheet
            For Each oSheet In oSheets
                oSheet.Activate()
                objDimensions = oSheet.Dimensions
                Dim strSheetName As String = oSheet.Name
                Dim intNumberofDimensions = objDimensions.Count
                For Each objDimension In objDimensions
                    Dim strDimName As String = objDimension.DisplayName
                    If objDimension.Type <> SolidEdgeFrameworkSupport.DimTypeConstants.igDimTypeAngular Then
                        dblDimValue = objDimension.Value
                        objDimensionStyle = objDimension.Style
                        'get the number of decimal places shown in the dimension....
                        NumberofDecimalDigits = objDimensionStyle.PrimaryDecimalRoundOff
                        Select Case NumberofDecimalDigits
                            Case 9
                                intNumberOfDecimalPlacesShownInDimension = 7
                            Case 8
                                intNumberOfDecimalPlacesShownInDimension = 6
                            Case 7
                                intNumberOfDecimalPlacesShownInDimension = 5
                            Case 6
                                intNumberOfDecimalPlacesShownInDimension = 4
                            Case 5
                                intNumberOfDecimalPlacesShownInDimension = 3
                            Case 4
                                intNumberOfDecimalPlacesShownInDimension = 2
                            Case 3
                                intNumberOfDecimalPlacesShownInDimension = 1
                            Case 2
                                intNumberOfDecimalPlacesShownInDimension = 0  '1 digit whole number
                            Case 1
                                intNumberOfDecimalPlacesShownInDimension = 0   '2 digit whole number....  not really supported i dont think!
                        End Select                       
                    End If

                    'get the conversion factor for this dimension
                    oUnitOfDimension = objDimension.Style.PrimaryUnits
                    If oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearCM Then
                        dblConversionFactor = 100
                    ElseIf oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearFeet Then
                        dblConversionFactor = 3.280839895
                    ElseIf oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearInches Then
                        dblConversionFactor = 39.37007874
                    ElseIf oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearMeters Then
                        dblConversionFactor = 1
                    ElseIf oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearMM Then
                        dblConversionFactor = 1000
                    End If

                    dblDimValue = dblDimValue * dblConversionFactor

                    'now collected all arguments to pass the checker function...
                    ' if found to be "bad" add it to the array containing bad dimension objects (arrayofBadDimensions)
                    'argument to pass in  display value, and number of decimal places shown in dimension
                    'call function to test the dimension

                    'go ahead and call the function
                    If JDimCheckV17RoundOff(dblDimValue, intNumberOfDecimalPlacesShownInDimension) Then
                        blnWillChange = True
                        Try
                            arrayofBadDimensions.Add(objDimension)
                        Catch ex As Exception

                        End Try
                    End If


                    'if the function returns a value of true then mark the dimension
                    If blnWillChange = True Then
                        AttachSymbolToDimension(objDimension, "Critical")
                    End If

                    'reset for next dimension
                    dblDimValue = 0
                    dblConversionFactor = 0
                    intNumberOfDecimalPlacesShownInDimension = 0
                    blnWillChange = False

                Next objDimension

                'release stuff
                oReleaseObject(objDimensionStyle)
                oReleaseObject(objDimension)
                oReleaseObject(objDimensions)

                'check inside drawing views
                'check for dimensions on each drawing view that might be placed in a drawing view on each sheet
                Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
                Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing

                dblDimValue = 0

                objDrawingViews = oSheet.DrawingViews
                For Each objDrawingView In objDrawingViews
                    Dim strDVName As String = objDrawingView.Name
                    Dim strDVType As SolidEdgeDraft.DrawingViewTypeConstants = objDrawingView.DrawingViewType
                    objDimensions = objDrawingView.Sheet.dimensions
                    For Each objDimension In objDimensions
                        Dim strDimName As String = objDimension.DisplayName
                        If objDimension.Type <> SolidEdgeFrameworkSupport.DimTypeConstants.igDimTypeAngular Then
                            dblDimValue = objDimension.Value
                            objDimensionStyle = objDimension.Style
                            'get the number of decimal places shown in the dimension....
                            NumberofDecimalDigits = objDimensionStyle.PrimaryDecimalRoundOff
                            Select Case NumberofDecimalDigits
                                Case 9
                                    intNumberOfDecimalPlacesShownInDimension = 7
                                Case 8
                                    intNumberOfDecimalPlacesShownInDimension = 6
                                Case 7
                                    intNumberOfDecimalPlacesShownInDimension = 5
                                Case 6
                                    intNumberOfDecimalPlacesShownInDimension = 4
                                Case 5
                                    intNumberOfDecimalPlacesShownInDimension = 3
                                Case 4
                                    intNumberOfDecimalPlacesShownInDimension = 2
                                Case 3
                                    intNumberOfDecimalPlacesShownInDimension = 1
                                Case 2
                                    intNumberOfDecimalPlacesShownInDimension = 0  '1 digit whole number
                                Case 1
                                    intNumberOfDecimalPlacesShownInDimension = 0   '2 digit whole number....  not really supported i dont think!
                            End Select
                        End If

                        'get the conversion factor for this dimension
                        oUnitOfDimension = objDimension.Style.PrimaryUnits
                        If oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearCM Then
                            dblConversionFactor = 100
                        ElseIf oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearFeet Then
                            dblConversionFactor = 3.280839895
                        ElseIf oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearInches Then
                            dblConversionFactor = 39.37007874
                        ElseIf oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearMeters Then
                            dblConversionFactor = 1
                        ElseIf oUnitOfDimension = SolidEdgeFrameworkSupport.DimLinearUnitConstants.igDimStyleLinearMM Then
                            dblConversionFactor = 1000
                        End If

                        dblDimValue = dblDimValue * dblConversionFactor

                        'now collected all arguments to pass the checker function...
                        ' if found to be "bad" add it to the array containing bad dimension objects (arrayofBadDimensions)
                        'argument to pass in  display value, and number of decimal places shown in dimension
                        'call function to test the dimension

                        'go ahead and call the function
                        If JDimCheckV17RoundOff(dblDimValue, intNumberOfDecimalPlacesShownInDimension) Then
                            blnWillChange = True
                            Try
                                arrayofBadDimensions.Add(objDimension)
                            Catch ex As Exception

                            End Try
                        End If


                        'if the function returns a value of true then mark the dimension
                        If blnWillChange = True Then
                            AttachSymbolToDimension(objDimension, "Critical")
                        End If

                        'reset for next dimension
                        dblDimValue = 0
                        dblConversionFactor = 0
                        intNumberOfDecimalPlacesShownInDimension = 0
                        blnWillChange = False

                    Next objDimension

                Next objDrawingView

                'release stuff
                oReleaseObject(objDimensionStyle)
                oReleaseObject(objDimension)
                oReleaseObject(objDimensions)
                oReleaseObject(objDrawingView)
                oReleaseObject(objDrawingViews)


            Next oSheet

            objActiveSheet.Activate()


            'release stuff
            oReleaseObject(objActiveSheet)
            oReleaseObject(oSheet)
            oReleaseObject(oSheets)


        Catch ex As Exception

        End Try



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
            Select Case strDimType.ToUpper
                Case "CRITICAL"  'place a filled in diamond
                    objBalloon = objBalloons.AddByTerminator(oDim, OriginX, OriginY, 0, True)
                    objBalloon.Callout = True
                    objBalloon.Style.Font = "Arial"
                    objBalloon.BalloonText = ChrW(&H2666)
                    objBalloon.Style.FontStyle = SolidEdgeFrameworkSupport.DimTextFontStyleConstants.igDimStyleFontBold
                    objBalloon.TextScale = 2
            End Select
        Catch ex As Exception
            
        End Try

        'release stuff....  since a local variable when goes out of scope should be released but good practice to do it anyway!
        oReleaseObject(oDimDisplayData)
        oReleaseObject(objBalloon)
        oReleaseObject(objBalloons)
        oReleaseObject(objActiveDraftSheet)


    End Sub


    

    Public Function JDimCheckV17RoundOff(dDimValue As Double, digits As Integer) As Boolean

        Dim value1 As Double = 0.0
        Dim value2 As Integer = 0
        Dim distTol As Double = 0.00000001
        Dim dScaleTolV15 As Double = 1.0
        Dim dScaleTolV17 As Double = 1.0
        Dim bWillChange As Boolean = False
        Dim bRoundUpV15 As Boolean = False
        Dim bRoundUpV17 As Boolean = False


        value1 = dDimValue * Math.Pow(10.0, (digits + 1))
        value2 = CInt(Math.Truncate(dDimValue * Math.Pow(10.0, digits)))


        dScaleTolV15 = distTol
        dScaleTolV17 = (distTol * Math.Pow(10.0, digits + 1))

        If EQ((value1 - CDbl(value2 * 10)), 5.0, dScaleTolV15) Then
            bRoundUpV15 = True
        End If

        If EQ((value1 - CDbl(value2 * 10)), 5.0, dScaleTolV17) Then
            bRoundUpV17 = True
        End If

        bWillChange = (bRoundUpV15 <> bRoundUpV17)


        Return bWillChange


    End Function



    Public Function EQ(ByVal val1 As Double, ByVal val2 As Double, ByVal tolerance As Double) As Boolean
        Return Math.Abs(val1 - val2) < tolerance
    End Function

End Module
