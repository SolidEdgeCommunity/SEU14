Imports System.Runtime.InteropServices
Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim application As SolidEdgeFramework.Application = Nothing
        Dim documents As SolidEdgeFramework.Documents = Nothing
        Dim partDocument As SolidEdgePart.PartDocument = Nothing
        Dim models As SolidEdgePart.Models = Nothing
        Dim model As SolidEdgePart.Model = Nothing
        Dim sketches As SolidEdgePart.Sketchs = Nothing
        Dim sketch As SolidEdgePart.Sketch = Nothing
        Dim refPlanes As SolidEdgePart.RefPlanes = Nothing
        Dim refPlane As SolidEdgePart.RefPlane = Nothing
        Dim profileSets As SolidEdgePart.ProfileSets = Nothing
        Dim profileSet As SolidEdgePart.ProfileSet = Nothing
        Dim profiles As SolidEdgePart.Profiles = Nothing
        Dim sketchProfile As SolidEdgePart.Profile = Nothing
        Dim profile As SolidEdgePart.Profile = Nothing
        Dim circles2d As SolidEdgeFrameworkSupport.Circles2d = Nothing

        Dim listPaths As New List(Of SolidEdgePart.Profile)()
        Dim listPathTypes As New List(Of SolidEdgePart.FeaturePropertyConstants)()
        Dim listSections As New List(Of SolidEdgePart.Profile)()
        Dim listSectionTypes As New List(Of SolidEdgePart.FeaturePropertyConstants)()
        Dim listOrigins As New List(Of Integer)()

        Dim objConstructions As SolidEdgePart.Constructions = Nothing
        Dim objSweptSurfs As SolidEdgePart.SweptSurfaces = Nothing
        Dim objSweptSurf As SolidEdgePart.SweptSurface = Nothing
        Try

            ' Connect to or start Solid Edge.
            application = Marshal.getactiveobject("SolidEdge.Application")

            ' Get a reference to the Documents collection.
            documents = application.Documents

            ' Create a new PartDocument.
            partDocument = application.Documents.Add("SolidEdge.PartDocument")

            ' Always a good idea to give SE a chance to breathe.
            application.DoIdle()

            ' Get a reference to the Construction collection.
            objConstructions = partDocument.Constructions

            ' Get a reference to the Sketches collections.
            sketches = DirectCast(partDocument.Sketches, SolidEdgePart.Sketchs)

            ' Get a reference to the profile sets collection.
            profileSets = DirectCast(partDocument.ProfileSets, SolidEdgePart.ProfileSets)

            ' Get a reference to the ref planes collection.
            refPlanes = DirectCast(partDocument.RefPlanes, SolidEdgePart.RefPlanes)

            ' Get a reference to front RefPlane.
            refPlane = refPlanes.Item(1)

            ' Add a new sketch.
            sketch = DirectCast(sketches.Add(), SolidEdgePart.Sketch)

            ' Add profile for sketch on specified refplane.
            sketchProfile = sketch.Profiles.Add(refPlane)

            ' Get a reference to the Circles2d collection.
            circles2d = sketchProfile.Circles2d

            ' Draw the Base Profile.
            circles2d.AddByCenterRadius(0, 0, 0.02)

            ' Close the profile.
            sketchProfile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

            ' Arrays for AddSweptProtrusion().
            listPaths.Add(sketchProfile)
            listPathTypes.Add(SolidEdgePart.FeaturePropertyConstants.igProfileBasedCrossSection)

            ' NOTE: profile is the Curve.
            refPlane = refPlanes.AddNormalToCurve(sketchProfile, SolidEdgePart.ReferenceElementConstants.igCurveEnd, refPlanes.Item(1), SolidEdgePart.ReferenceElementConstants.igPivotEnd, True, System.Reflection.Missing.Value)

            ' Add a new profile set.
            profileSet = DirectCast(profileSets.Add(), SolidEdgePart.ProfileSet)

            ' Get a reference to the profiles collection.
            profiles = DirectCast(profileSet.Profiles, SolidEdgePart.Profiles)

            ' add a new profile.
            profile = DirectCast(profiles.Add(refPlane), SolidEdgePart.Profile)

            ' Get a reference to the Circles2d collection.
            circles2d = profile.Circles2d

            ' Draw the Base Profile.
            circles2d.AddByCenterRadius(0, 0, 0.01)

            ' Close the profile.
            profile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

            ' Arrays for AddSweptProtrusion().
            listSections.Add(profile)
            listSectionTypes.Add(SolidEdgePart.FeaturePropertyConstants.igProfileBasedCrossSection)
            listOrigins.Add(0) 'Use 0 for closed profiles.

            objSweptSurfs = objConstructions.SweptSurfaces
            objSweptSurf = objSweptSurfs.Add(listPaths.Count, listPaths.ToArray(), SolidEdgePart.FeaturePropertyConstants.igEdgeBasedCrossSection, listSections.Count, listSections.ToArray(), SolidEdgePart.FeaturePropertyConstants.igProfileBasedCrossSection, listOrigins.ToArray(), 0, SolidEdgePart.FeaturePropertyConstants.igNone, SolidEdgePart.FeaturePropertyConstants.igNone)

            ' Hide profiles.
            sketchProfile.Visible = False
            profile.Visible = False

            ' Switch to ISO view.
            application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewISOView)
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim objSEApp As SolidEdgeFramework.Application = Nothing
        Dim objDoc As SolidEdgePart.PartDocument = Nothing
        Dim documents As SolidEdgeFramework.Documents = Nothing
        Dim UOM As SolidEdgeFramework.UnitsOfMeasure = Nothing

        ' Connect to or start Solid Edge.
        objSEApp = Marshal.GetActiveObject("SolidEdge.Application")

        ' Get a reference to the Documents collection.
        documents = objSEApp.Documents

        objDoc = oCreateNewSEDocument(objSEApp, "part")
        objDoc.ModelingMode = SolidEdgeConstants.ModelingModeConstants.seModelingModeOrdered

        Dim objReferencePlanes As SolidEdgePart.RefPlanes = Nothing
        Dim objReferencePlane As SolidEdgePart.RefPlane = Nothing

        objReferencePlanes = objDoc.RefPlanes
        For Each objReferencePlane In objReferencePlanes
            'default reference planes 1=Top(XY) 2=Right (YZ), 3=Front (XZ)
            Dim strReferencePlaneName As String = objReferencePlane.Name
            Dim strReferencePlaneSystemName As String = objReferencePlane.SystemName
            Dim strReferencePlaneEdgebarName As String = objReferencePlane.EdgebarName
        Next
        'Lets draw the profile on the front reference plane
        objReferencePlane = objReferencePlanes.Item(3)
        UOM = objDoc.UnitsOfMeasure
        objReferencePlane = objReferencePlanes.AddParallelByDistance(objReferencePlanes.Item(3), _
                                                                       UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 10), _
                                                                       SolidEdgePart.FeaturePropertyConstants.igLeft, , , False)
        Dim IsConsumedInProfile As Boolean = objReferencePlane.Global

        'now ready to actully draw the profile geometry
        Dim objProfile(0) As SolidEdgePart.Profile
        Dim objProfileCircles As SolidEdgeFrameworkSupport.Circles2d = Nothing
        objProfile(0) = objDoc.ProfileSets.Add.Profiles.Add(objReferencePlane)
        objProfileCircles = objProfile(0).Circles2d
        objProfileCircles.AddByCenterRadius(0, 0, UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 10))

        Dim objDimensions As SolidEdgeFrameworkSupport.Dimensions = Nothing
        Dim objDim1 As SolidEdgeFrameworkSupport.Dimension = Nothing
        objDimensions = objProfile(0).Dimensions
        'place a dimension on the profile circle
        objDim1 = objDimensions.AddCircularDiameter(objProfileCircles.Item(1))
        objDim1.Constraint = True 'creates a driving dimension....setting to false will create a driven dimension
        objDim1.TrackDistance = UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, -40)  'pushes dim away from geometry

        Dim objRefLine As SolidEdgeFrameworkSupport.Line2d = Nothing
        objRefLine = objProfile(0).ProjectRefPlane(objReferencePlanes.Item(2)) ' Create dimensionsable object from reference plane

        Dim objRelns As SolidEdgeFrameworkSupport.Relations2d = Nothing
        ' Define Relations among the line objects to make the profile closed
        objRelns = objProfile(0).Relations2d
        ' Adds connect constraint from center of the profile to the midpoin of the reference plane
        Call objRelns.AddKeypoint(objProfileCircles.Item(1), SolidEdgeConstants.KeypointIndexConstants.igCircleCenter, _
                                    objRefLine, SolidEdgeConstants.KeypointIndexConstants.igLineMiddle, )

        ' Check for the Profile Validity
        Dim lngStatus As Long = objProfile(0).End(SolidEdgePart.ProfileValidationType.igProfileClosed)
        If lngStatus <> 0 Then
            MessageBox.Show("Profile not closed")
        End If

        ' Create the base Extruded Protrusion Feature
        Dim objModel = objDoc.Models.AddFiniteExtrudedProtrusion(1, objProfile, SolidEdgePart.FeaturePropertyConstants.igLeft, UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 75))

        objProfile(0).Visible = False

        'Check the status of Base Feature

        If objModel.ExtrudedProtrusions.Item(1).Status <> SolidEdgePart.FeatureStatusConstants.igFeatureOK Then
            MessageBox.Show("Error in the Creation of Base Protrusion Feature object")
        End If


        'create protrusion using Models collection method
        'Dim objMProfile As SolidEdgePart.Profile
        'Dim objL1, objL2, objL3, objL4 As SolidEdgeFrameworkSupport.Line2d
        'Dim objRelations As SolidEdgeFrameworkSupport.Relations2d = Nothing
        'objMProfile = objDoc.ProfileSets.Add.Profiles.Add(objDoc.RefPlanes.Item(3))

        'objL1 = objMProfile.Lines2d.AddBy2Points(-0.015, 0.025, -0.015, 0.035)
        'objL2 = objMProfile.Lines2d.AddBy2Points(-0.015, 0.035, -0.04, 0.035)
        'objL3 = objMProfile.Lines2d.AddBy2Points(-0.04, 0.035, -0.04, 0.025)
        'objL4 = objMProfile.Lines2d.AddBy2Points(-0.04, 0.025, -0.015, 0.025)

        'objRelations = objMProfile.Relations2d
        'Call objRelations.AddKeypoint(objL1, SolidEdgeConstants.KeypointIndexConstants.igLineEnd, objL2, SolidEdgeConstants.KeypointIndexConstants.igLineStart)
        'Call objRelations.AddKeypoint(objL2, SolidEdgeConstants.KeypointIndexConstants.igLineEnd, objL3, SolidEdgeConstants.KeypointIndexConstants.igLineStart)
        'Call objRelations.AddKeypoint(objL3, SolidEdgeConstants.KeypointIndexConstants.igLineEnd, objL4, SolidEdgeConstants.KeypointIndexConstants.igLineStart)
        'Call objRelations.AddKeypoint(objL4, SolidEdgeConstants.KeypointIndexConstants.igLineEnd, objL1, SolidEdgeConstants.KeypointIndexConstants.igLineStart)

        'lngStatus = objMProfile.End(SolidEdgeConstants.ProfileValidationType.igProfileClosed)
        'If lngStatus <> 0 Then
        '    MessageBox.Show("invalid profile")
        'End If

        'now let's create a lofted cutout
        'need to create RP on the end face of the protrusion just created!
        objReferencePlane = objReferencePlanes.AddParallelByDistance(objModel.ExtrudedProtrusions.Item(1).TopCap, _
                                                                       UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 0), _
                                                                       SolidEdgePart.FeaturePropertyConstants.igLeft, , , True)
        ' Dim objProfileCS(0) As SolidEdgePart.Profile
        Dim objProfileCS(6) As SolidEdgePart.Profile
        objProfileCS(0) = objDoc.ProfileSets.Add.Profiles.Add(objReferencePlane)
        objProfileCircles = objProfileCS(0).Circles2d
        objProfileCircles.AddByCenterRadius(UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 6.53), _
                                              UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 5.08), _
                                              UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 4.76))
        'Create Dimension
        objDim1 = objDimensions.AddCircularDiameter(objProfileCircles.Item(1))
        objDim1.Constraint = True ' Creates a driving dimension....setting to false will create a driven dimension
        objDim1.TrackDistance = UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, -40) 'pushes dim away from geometry
        objDim1 = Nothing

        objRefLine = objProfileCS(0).ProjectRefPlane(objReferencePlanes.Item(2)) ' Create dimensionable object from Reference plane

        ' Check for the Profile Validity
        lngStatus = objProfileCS(0).End(SolidEdgePart.ProfileValidationType.igProfileClosed)
        If lngStatus <> 0 Then
            MessageBox.Show("Profile not closed")
        End If

        Dim objRefPln As SolidEdgePart.RefPlane = Nothing

        For i = 1 To 6
          
            objRefPln = objDoc.RefPlanes.AddParallelByDistance(objReferencePlane, _
                                UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 20 / 8 * (i)), SolidEdgeConstants.FeaturePropertyConstants.igLeft)


            objProfileCS(i) = objDoc.ProfileSets.Add.Profiles.Add(objRefPln)
            Call objProfileCS(i).Circles2d.AddByCenterRadius(GetRotatedXCoord(UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 6.53), _
                                                                                 UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 5.08), _
                                                                                 UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitAngle, 180) / 7 * (i + 1)), _
                                                                                 GetRotatedYCoord(UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 6.53), _
                                                                                 UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 5.08), _
                                                                                 UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitAngle, 180) / 7 * (i + 1)), _
                                                                                 UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 4.76))

            ' Check if the Profile is closed
            lngStatus = objProfileCS(i).End(ValidationCriteria:=SolidEdgePart.ProfileValidationType.igProfileClosed)
            If lngStatus <> 0 Then
                MsgBox("Profile not closed")
            End If
        Next i


        Dim SectionTypes(6) As Long
        Dim OriginArray(6) As Object
        For i = 0 To 6
            SectionTypes(i) = SolidEdgePart.FeaturePropertyConstants.igProfileBasedCrossSection
        Next i
        For i = 0 To 6
            OriginArray(i) = 0
        Next i

        Dim objLoft As SolidEdgePart.LoftedCutout = Nothing
        objLoft = objDoc.Models.Item(1).LoftedCutouts.Add(NumSections:=7, CrossSections:=objProfileCS, _
                                CrossSectionTypes:=SectionTypes, _
                                Origins:=OriginArray, _
                                SegmentMaps:=0, MaterialSide:=SolidEdgePart.FeaturePropertyConstants.igLeft, _
                                StartExtentType:=SolidEdgePart.FeaturePropertyConstants.igNone, StartExtentDistance:=0, _
                                StartSurfaceOrRefPlane:=Nothing, EndExtentType:=SolidEdgePart.FeaturePropertyConstants.igNone, _
                                EndExtentDistance:=0, EndSurfaceOrRefPlane:=Nothing, _
                                StartTangentType:=SolidEdgePart.FeaturePropertyConstants.igNormal, StartTangentMagnitude:=0.5, _
                                EndTangentType:=SolidEdgePart.FeaturePropertyConstants.igNormal, EndTangentMagnitude:=0.5)


        'Trun off the profile
        For i = 0 To 6
            objProfileCS(i).Visible = False
        Next i


        If objModel.LoftedCutouts.Item(1).Status <> SolidEdgePart.FeatureStatusConstants.igFeatureOK Then
            MessageBox.Show("Error in the Creation of the lofted cutout")
        End If

        'now pattern the lofted cutout
        Dim FeaturesToPattern(0) As Object
        Dim objPatternProfile As SolidEdgeFrameworkSupport.CircularPatterns2d = Nothing
        Dim objProf As SolidEdgePart.Profile = Nothing
        Const PI = 3.14159265358979

        FeaturesToPattern(0) = objLoft


        objProf = objDoc.ProfileSets.Add.Profiles.Add(objReferencePlane)
        objPatternProfile = objProf.CircularPatterns2d

        Call objPatternProfile.AddByCircle(0, 0, UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 4.76), 0, _
                                           SolidEdgeFrameworkSupport.Geom2dOrientationConstants.igGeom2dOrientClockwise, _
                                           SolidEdgeFrameworkSupport.PatternOffsetTypeConstants.sePatternFixedOffset, 4, PI / 4)


        Call objDoc.Models.Item(1).Patterns.Add(NumberOfFeatures:=1, _
                            FeatureArray:=FeaturesToPattern, Profile:=objProf, PatternType:=SolidEdgePart.PatternTypeConstants.seFastPattern)

        objProf.Visible = False

        If objModel.Patterns.Item(1).Status <> SolidEdgePart.FeatureStatusConstants.igFeatureOK Then
            MessageBox.Show("Error in the Creation of the lofted cutout")
        End If


        'close and save the file
        If objDoc.Path = String.Empty Then  ' The file has never been saved before
            objDoc.SaveAs(NewName:="C:\temp\Drill_Ordered.par")
            objDoc.Close(SaveChanges:=False)
        Else  ' It is an existing document that has been saved at least once so do not need to provide a path/name
            objDoc.Save()
            objDoc.Close(False)
        End If
    End Sub
    Public Function GetRotatedXCoord(xval As Double, yval As Double, ang As Double) As Double
        Dim r As Double


        r = Math.Sqrt(xval ^ 2 + yval ^ 2)
        GetRotatedXCoord = r * Math.Cos(ang)

    End Function

    Public Function GetRotatedYCoord(xval As Double, yval As Double, ang As Double) As Double
        Dim r As Double


        r = Math.Sqrt(xval ^ 2 + yval ^ 2)
        GetRotatedYCoord = r * Math.Sin(ang)

    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim objDoc As SolidEdgePart.PartDocument = Nothing
        Dim UOM As SolidEdgeFramework.UnitsOfMeasure = Nothing

        If oConnectToSolidEdge(True, True) Then
            Do While objSEApp.Documents.Count <> 0
                objSEApp.ActiveDocument.close(False)
            Loop

            objDoc = oCreateNewSEDocument(objSEApp, "part")
            objDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous
            UOM = objDoc.UnitsOfMeasure

            Dim objReferencePlanes As SolidEdgePart.RefPlanes = Nothing
            Dim objReferencePlane As SolidEdgePart.RefPlane = Nothing

            objReferencePlanes = objDoc.RefPlanes
            For Each objReferencePlane In objReferencePlanes
                'default reference planes 1=Top(XY) 2=Right (YZ), 3=Front (XZ)
                Dim strReferencePlaneName As String = objReferencePlane.Name
                Dim strReferencePlaneSystemName As String = objReferencePlane.SystemName
                Dim strReferenceplaneEdgebarName As String = objReferencePlane.EdgebarName
            Next

            'Let's draw the profile on the front reference plane
            objReferencePlane = objReferencePlanes.Item(3)


            'now ready to actually draw the profile geometry
            Dim objProfile(0) As SolidEdgePart.Profile
            Dim objProfileCircles As SolidEdgeFrameworkSupport.Circles2d = Nothing

            objProfile(0) = objDoc.ProfileSets.Add.Profiles.Add(objReferencePlane)
            objProfileCircles = objProfile(0).Circles2d
            objProfileCircles.AddByCenterRadius(0.1, 0.1, UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 6.35))

            Dim objDimensions As SolidEdgeFrameworkSupport.Dimensions = Nothing
            Dim objDim1 As SolidEdgeFrameworkSupport.Dimension = Nothing

            objDimensions = objProfile(0).Dimensions
            'place a dimension on the profile circle
            objDim1 = objDimensions.AddCircularDiameter(objProfileCircles.Item(1))
            objDim1.Constraint = True   ' Creates a driving dimension .... setting to false will create a driven dimension
            objDim1.TrackDistance = UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, -40)   'pushes dim away from geometry

            Dim objRefLine As SolidEdgeFrameworkSupport.Line2d = Nothing
            objRefLine = objProfile(0).ProjectRefPlane(objReferencePlanes.Item(2)) 'Create dimensionable object from Reference plane

            Dim objRelns As SolidEdgeFrameworkSupport.Relations2d = Nothing
            ' Define Relations among the Line objects to make the Profile closed
            objRelns = objProfile(0).Relations2d
            ' Adds connect constraint from center of the profile to the midpoint of the reference plane
            Call objRelns.AddKeypoint(objProfileCircles.Item(1), SolidEdgeConstants.KeypointIndexConstants.igCircleCenter, _
                                      objRefLine, SolidEdgeConstants.KeypointIndexConstants.igLineMiddle, )


            ' Check for the Profile Validity
            Dim lngStatus As Long = objProfile(0).End(ValidationCriteria:=SolidEdgePart.ProfileValidationType.igProfileClosed)
            If lngStatus <> 0 Then
                MessageBox.Show("Profile not closed")
            End If

            oReleaseObject(objRelns)
            oReleaseObject(objRefLine)
            oReleaseObject(objDim1)
            oReleaseObject(objReferencePlane)

            'Create the Base Extruded Protrusion Feature
            Dim objModel As SolidEdgePart.Model = Nothing

            objModel = objDoc.Models.AddFiniteExtrudedProtrusion(NumberOfProfiles:=1, _
            ProfileArray:=objProfile, ProfilePlaneSide:= _
            SolidEdgePart.FeaturePropertyConstants.igLeft, ExtrusionDistance:=UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 75))

            objProfile(0).Visible = False

            ' Check the status of Base Feature

            If objModel.ExtrudedProtrusions.Item(1).Status <> SolidEdgePart.FeatureStatusConstants.igFeatureOK Then
                MessageBox.Show("Error in the Creation of Base Protrusion Feature object")
            End If




            'now let's create a lofted cutout
            'need to create RP on the end face of the protrusion just created!
            'sync model does not have a topcap method
            Dim oFaces As SolidEdgeGeometry.Faces = Nothing
            Dim oFace As SolidEdgeGeometry.Face = Nothing
            oFaces = objModel.ExtrudedProtrusions.Item(1).Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll)
            'now have to loop through each face with logic to fine the one I need
            Dim dblParam(3) As Double
            Dim dblNormal() As Double = {0}
            ' getting the normal at a point (u=0.5, v=0.5) on a Face
            Const PI = 3.14159265358979
            dblParam(0) = 3 * PI / 2
            dblParam(1) = 0.1
            dblParam(2) = PI / 2
            dblParam(3) = 0.05
            For Each oFace In oFaces
                oFace.GetNormal(NumParams:=2, Params:=dblParam, Normals:=dblNormal)
                'if the normal if the face being process in - Y then that's the one I want
                If dblNormal(1) = -1 Then
                    Exit For
                End If
            Next

            'objReferencePlane = objReferencePlanes.AddParallelByDistance(objModel.ExtrudedProtrusions.Item(1).TopCap, _
            '                                                             UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 0), _
            '                                                             SolidEdgePart.FeaturePropertyConstants.igLeft, , , True)

            objReferencePlane = objReferencePlanes.AddParallelByDistance(oFace, _
                                                                        UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 0), _
                                                                        SolidEdgePart.FeaturePropertyConstants.igLeft, , , True)
            oReleaseObject(oFace)
            oReleaseObject(oFaces)


            Dim objProfileCS(6) As SolidEdgePart.Profile

            objProfileCS(0) = objDoc.ProfileSets.Add.Profiles.Add(objReferencePlane)
            objProfileCircles = objProfileCS(0).Circles2d
            objProfileCircles.AddByCenterRadius(UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 6.53), _
                                                UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 5.08), _
                                                UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 4.76))

            'Create Dimension
            objDim1 = objDimensions.AddCircularDiameter(objProfileCircles.Item(1))
            objDim1.Constraint = True   ' Creates a driving dimension .... setting to false will create a driven dimension
            objDim1.TrackDistance = UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, -40)   'pushes dim away from geometry


            ' Check for the Profile Validity
            lngStatus = objProfileCS(0).End(ValidationCriteria:=SolidEdgePart.ProfileValidationType.igProfileClosed)
            If lngStatus <> 0 Then
                MessageBox.Show("Profile not closed")
            End If

            oReleaseObject(objDim1)
            oReleaseObject(objDimensions)

            Dim objRefPln As SolidEdgePart.RefPlane = Nothing

            For i = 1 To 6
                objRefPln = objDoc.RefPlanes.AddParallelByDistance(objReferencePlane, _
                Distance:=UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 20 / 8 * (i)), NormalSide:=SolidEdgeConstants.FeaturePropertyConstants.igLeft)

                objProfileCS(i) = objDoc.ProfileSets.Add.Profiles.Add(pRefPlaneDisp:=objRefPln)
                Call objProfileCS(i).Circles2d.AddByCenterRadius(GetRotatedXCoord(UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 6.53), _
                                                                                  UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 5.08), _
                                                                                  UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitAngle, 180) / 7 * (i + 1)), _
                                                                                  GetRotatedYCoord(UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 6.53), _
                                                                                  UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 5.08), _
                                                                                  UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitAngle, 180) / 7 * (i + 1)), _
                                                                                  UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 4.76))

                ' Check if the Profile is closed
                lngStatus = objProfileCS(i).End(ValidationCriteria:=SolidEdgePart.ProfileValidationType.igProfileClosed)
                If lngStatus <> 0 Then
                    MsgBox("Profile not closed")
                End If
                oReleaseObject(objRefPln)

            Next i


            Dim SectionTypes(6) As Long
            Dim OriginArray(6) As Object
            For i = 0 To 6
                SectionTypes(i) = SolidEdgePart.FeaturePropertyConstants.igProfileBasedCrossSection
            Next i
            For i = 0 To 6
                OriginArray(i) = 0
            Next i

            Dim objLoft As SolidEdgePart.LoftedCutout = Nothing
            objLoft = objDoc.Models.Item(1).LoftedCutouts.Add(NumSections:=7, CrossSections:=objProfileCS, _
                                    CrossSectionTypes:=SectionTypes, _
                                    Origins:=OriginArray, _
                                    SegmentMaps:=0, MaterialSide:=SolidEdgePart.FeaturePropertyConstants.igLeft, _
                                    StartExtentType:=SolidEdgePart.FeaturePropertyConstants.igNone, StartExtentDistance:=0, _
                                    StartSurfaceOrRefPlane:=Nothing, EndExtentType:=SolidEdgePart.FeaturePropertyConstants.igNone, _
                                    EndExtentDistance:=0, EndSurfaceOrRefPlane:=Nothing, _
                                    StartTangentType:=SolidEdgePart.FeaturePropertyConstants.igNormal, StartTangentMagnitude:=0.5, _
                                    EndTangentType:=SolidEdgePart.FeaturePropertyConstants.igNormal, EndTangentMagnitude:=0.5)


            'Trun off the profile
            For i = 0 To 6
                objProfileCS(i).Visible = False
            Next i


            If objModel.LoftedCutouts.Item(1).Status <> SolidEdgePart.FeatureStatusConstants.igFeatureOK Then
                MessageBox.Show("Error in the Creation of the lofted cutout")
            End If

            'now pattern the lofted cutout
            Dim FeaturesToPattern(0) As Object
            Dim objPatternProfile As SolidEdgeFrameworkSupport.CircularPatterns2d = Nothing
            Dim objProf As SolidEdgePart.Profile = Nothing


            FeaturesToPattern(0) = objLoft


            objProf = objDoc.ProfileSets.Add.Profiles.Add(objReferencePlane)
            objPatternProfile = objProf.CircularPatterns2d

            Call objPatternProfile.AddByCircle(0, 0, UOM.ParseUnit(SolidEdgeFramework.UnitTypeConstants.igUnitDistance, 4.76), 0, _
                                               SolidEdgeFrameworkSupport.Geom2dOrientationConstants.igGeom2dOrientClockwise, _
                                               SolidEdgeFrameworkSupport.PatternOffsetTypeConstants.sePatternFixedOffset, 4, PI / 4)


            'ordered equivalent
            'Call objDoc.Models.Item(1).Patterns.Add(NumberOfFeatures:=1, _
            '                FeatureArray:=FeaturesToPattern, Profile:=objProf, _
            '                PatternType:=SolidEdgePart.PatternTypeConstants.seFastPattern)

            'Synchronous equivalent
            Call objDoc.Models.Item(1).Patterns.AddSync(NumberOfFeatures:=1, _
                            FeatureArray:=FeaturesToPattern, Profile:=objProf)

            objProf.Visible = False

            If objModel.Patterns.Item(1).Status <> SolidEdgePart.FeatureStatusConstants.igFeatureOK Then
                MessageBox.Show("Error in the Creation of the lofted cutout")
            End If


            'close and save the file
            If objDoc.Path = String.Empty Then  ' The file has never been saved before
                objDoc.SaveAs(NewName:="C:\temp\Drill_Synchronous.par")
                objDoc.Close(SaveChanges:=False)
            Else  ' It is an existing document that has been saved at least once so do not need to provide a path/name
                objDoc.Save()
                objDoc.Close(False)
            End If


            oReleaseObject(objRefPln)
            oReleaseObject(objLoft)
            oReleaseObject(objProf)
            oReleaseObject(objPatternProfile)
            oReleaseObject(objDoc)
            oForceGarbageCollection()


        End If
    End Sub
End Class
