Imports Inventor

Public Class InitFold
    Public doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Dim refLine, firstLine, secondLine, thirdLine, lastLine, kanteLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim curve3D As Curves3D
    Dim monitor As DesignMonitoring
    Dim medico As DesignDoctor
    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3 As Point
    Dim tg As TransientGeometry
    Dim gap1 As Double
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines, bigFaces As ObjectCollection
    Dim comando As Commands
    Dim mainSketch As InitSketcher
    Dim faceProfile As Profile
    Dim feature As FaceFeature
    Dim foldFeature As FoldFeature
    Dim bendLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim workface, cutFace, nextworkFace As Face
    Dim edgeBand, majorEdge, minorEdge, bendEdge As Edge
    Dim bendAngle As DimensionConstraint
    Dim folded As FoldFeature
    Dim cutFeature As CutFeature
    Dim faceFeature As FaceFeature
    Dim corte As Cortador
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim foldFeatures As FoldFeatures
    Dim lamp As Highlithing
    Dim bender As Doblador
    Dim backwards As Boolean
    Dim initialPlane As Plane
    Dim foldDefinition As FoldDefinition
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        compDef = doku.ComponentDefinition
        sheetMetalFeatures = compDef.Features
        foldFeatures = sheetMetalFeatures.FoldFeatures
        curve3D = New Curves3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        medico = New DesignDoctor(doku)
        mainSketch = New InitSketcher(doku)
        lamp = New Highlithing(doku)
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        bigFaces = app.TransientObjects.CreateObjectCollection

        done = False
    End Sub


    Public Function MakeFirstFold(refDoc As FindReferenceLine) As Boolean
        Dim b As Boolean
        Try
            Try
                cutFeature = MakeInitCut()
            Catch ex As Exception
                sk3D = mainSketch.StartDrawingTranslated(refDoc, 1)
                If mainSketch.done Then
                    doku.ComponentDefinition.Sketches3D.Item("s1").Visible = False
                    If DrawBandStripe().Count > 0 Then
                        bender = New Doblador(doku)
                        If MakeStartingFace(faceProfile).GetHashCode > 0 Then
                            If GetBendLine().Length > 0 Then
                                mainWorkPlane.Visible = False
                                If GetFoldingAngle().Parameter._Value > 0 Then

                                    comando.MakeInvisibleSketches(doku)
                                    comando.MakeInvisibleWorkPlanes(doku)
                                    foldFeature = FoldBand()
                                    If monitor.IsFeatureHealthy(foldFeature) Then
                                        cutFeature = MakeInitCut()
                                        If monitor.IsFeatureHealthy(cutFeature) Then
                                            doku.Save2(True)
                                            done = 1
                                            Return True
                                        End If


                                    End If


                                End If

                            End If

                        End If
                    End If

                End If
            End Try




            Return False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
    Function MakeInitCut() As CutFeature
        foldFeature = foldFeatures.Item(1)
        If monitor.IsFeatureHealthy(foldFeature) Then
            corte = New Cortador(doku)
            cutFace = corte.GetInitCutFace(foldFeature)
            firstLine = compDef.Sketches3D.Item("s1").SketchLines3D.Item(2)
            cutFeature = corte.CutInitFace(firstLine)
            cutFeature.Name = "initCut"
        End If
        Return cutFeature
    End Function
    Public Function MakeFirstFold(refDoc As FindReferenceLine, q As Integer) As Boolean
        Dim b As Boolean
        Try
            If refDoc.foldFeatures.Count > 8 Then
                backwards = True
            Else
                backwards = False
            End If
            Try
                foldFeature = foldFeatures.Item(1)
                cutFeature = MakeInitCut()
            Catch ex As Exception

                sk3D = mainSketch.DrawMainSketch(refDoc, q)
                If mainSketch.done Then
                    doku.ComponentDefinition.Sketches3D.Item("s1").Visible = False
                    If DrawBandStripe().Count > 0 Then
                        kanteLine = refDoc.GetKanteLine()
                        If MakeStartingFace(faceProfile).GetHashCode > 0 Then
                            If GetWorkFace().Evaluator.Area > 0 Then
                                If GetBendLine(workface, mainSketch.thirdLine).Length > 0 Then
                                    mainWorkPlane.Visible = False
                                    If GetFoldingAngle().Parameter._Value > 0 Then
                                        comando.MakeInvisibleSketches(doku)
                                        comando.MakeInvisibleWorkPlanes(doku)
                                        bandLines = mainSketch.bandLines
                                        folded = FoldBand(bandLines.Count)
                                        folded = CheckFoldSide(folded)
                                        folded.Name = "f1"
                                        doku.Update2(True)
                                        If monitor.IsFeatureHealthy(folded) Then
                                            cutFeature = MakeInitCut()
                                            If monitor.IsFeatureHealthy(cutFeature) Then
                                                doku.Save2(True)
                                                done = 1
                                                Return True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End Try




            Return False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function

    Function DrawBandStripe() As Profile
        Dim ps As PlanarSketch

        Dim p1, p2, p3 As SketchPoint
        Dim v As Vector2d
        Dim r As SketchEntitiesEnumerator
        Dim ln, l2d As SketchLine
        Try
            mainWorkPlane = doku.ComponentDefinition.WorkPlanes.AddByThreePoints(mainSketch.firstLine.StartPoint, mainSketch.firstLine.EndPoint, mainSketch.secondLine.EndPoint)
            ps = doku.ComponentDefinition.Sketches.Add(mainWorkPlane)
            p1 = ps.AddByProjectingEntity(mainSketch.firstLine.StartPoint)
            p2 = ps.AddByProjectingEntity(mainSketch.firstLine.EndPoint)
            p3 = ps.AddByProjectingEntity(mainSketch.secondLine.EndPoint)
            ln = ps.AddByProjectingEntity(mainSketch.bandLines.Item(2))
            v = ln.Geometry.Direction.AsVector
            v.ScaleBy(-3 / 10)

            p2.MoveBy(v)
            l2d = ps.SketchLines.AddByTwoPoints(p1.Geometry, p2)
            ps.GeometricConstraints.AddPerpendicular(l2d, ln)
            Dim dc As DimensionConstraint
            dc = ps.DimensionConstraints.AddTwoPointDistance(l2d.StartSketchPoint, l2d.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, p2.Geometry)
            dc.Parameter._Value = GetParameter("b")._Value / 1
            v.Normalize()
            v.ScaleBy(-50)
            p3.MoveBy(v)
            r = ps.SketchLines.AddAsThreePointRectangle(l2d.StartSketchPoint.Geometry, l2d.EndSketchPoint.Geometry, p3.Geometry)

            faceProfile = ps.Profiles.AddForSolid

            Return faceProfile

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeStartingFace(pro As Profile) As FaceFeature
        Dim oSheetMetalCompDef As SheetMetalComponentDefinition
        oSheetMetalCompDef = doku.ComponentDefinition
        oSheetMetalCompDef.UseSheetMetalStyleThickness = False
        Dim oThicknessParam As Parameter
        oThicknessParam = oSheetMetalCompDef.Thickness
        oThicknessParam.Units = UnitsTypeEnum.kMillimeterLengthUnits
        oThicknessParam._Value = 0.03

        Dim oSheetMetalFeatures As SheetMetalFeatures
        oSheetMetalFeatures = doku.ComponentDefinition.Features

        compDef = oSheetMetalCompDef
        Dim oFaceFeatureDefinition As FaceFeatureDefinition
        oFaceFeatureDefinition = oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(pro)
        If CalculateFaceDirection() < 0 Then
            oFaceFeatureDefinition.Direction = PartFeatureExtentDirectionEnum.kPositiveExtentDirection
        Else
            oFaceFeatureDefinition.Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection
        End If

        Dim oFaceFeature As FaceFeature
        oFaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition)
        feature = oFaceFeature
        Return oFaceFeature
    End Function
    Function CalculateFaceDirection() As Integer

        Dim d As Double
        Dim s As Integer

        Try
            d = kanteLine.Geometry.Direction.AsVector.X
            If d < 0 Then
                s = -1
            Else
                s = 1
            End If
            Return s
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function FoldBand() As FoldFeature
        Try
            Dim oSheetMetalFeatures As SheetMetalFeatures
            oSheetMetalFeatures = doku.ComponentDefinition.Features
            Dim oFoldDefinition As FoldDefinition
            If bendAngle.Parameter._Value > Math.PI / 2 Then
                oFoldDefinition = oSheetMetalFeatures.FoldFeatures.CreateFoldDefinition(bendLine, bendAngle.Parameter._Value)
            Else

                oFoldDefinition = oSheetMetalFeatures.FoldFeatures.CreateFoldDefinition(bendLine, Math.PI - bendAngle.Parameter._Value)
            End If

            Dim oFoldFeature As FoldFeature
            oFoldFeature = oSheetMetalFeatures.FoldFeatures.Add(oFoldDefinition)
            folded = oFoldFeature
            Return folded

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetBendLine() As SketchLine
        Try
            Dim maxArea As Double = 0
            Dim oFoldLineSketch As PlanarSketch
            Dim maxface1, maxface2 As Face
            Dim pt1, pt2 As Point
            Dim d1, d2, a As Double
            oFoldLineSketch = compDef.Sketches.Add(mainWorkPlane)
            maxface1 = feature.Faces.Item(1)
            d1 = maxface1.GetClosestPointTo(mainSketch.firstLine.StartSketchPoint.Geometry).DistanceTo(mainSketch.firstLine.StartSketchPoint.Geometry)
            d2 = d1
            maxface2 = maxface1
            For Each f As Face In feature.Faces
                a = f.Evaluator.Area
                If f.Evaluator.Area > maxArea * 0.99 Then
                    maxArea = f.Evaluator.Area
                    maxface2 = maxface1
                    maxface1 = f
                End If
            Next
            d1 = maxface1.GetClosestPointTo(mainSketch.firstLine.StartSketchPoint.Geometry).DistanceTo(mainSketch.firstLine.StartSketchPoint.Geometry)
            d2 = maxface2.GetClosestPointTo(mainSketch.firstLine.StartSketchPoint.Geometry).DistanceTo(mainSketch.firstLine.StartSketchPoint.Geometry)

            If d2 < d1 Then
                maxface1 = maxface2
            End If
            Dim se1, se2 As SketchLine
            Dim e1, e2 As Edge
            Dim maxe1, maxe2 As Double
            maxe1 = 0
            maxe2 = 0
            e1 = maxface1.EdgeLoops.Item(1).Edges.Item(1)
            e2 = e1
            For Each el As EdgeLoop In maxface1.EdgeLoops
                For Each ed As Edge In el.Edges
                    If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) >= maxe1 Then
                        maxe2 = maxe1
                        e2 = e1
                        maxe1 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                        e1 = ed
                    End If
                Next

            Next
            se1 = oFoldLineSketch.AddByProjectingEntity(e1)
            se1.Construction = True
            se2 = oFoldLineSketch.AddByProjectingEntity(e2)
            se2.Construction = True
            edgeBand = e2
            Dim sl, bl As SketchLine
            sl = oFoldLineSketch.AddByProjectingEntity(mainSketch.thirdLine)
            sl.Construction = True
            bl = oFoldLineSketch.SketchLines.AddByTwoPoints(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry)
            oFoldLineSketch.GeometricConstraints.AddCollinear(bl, sl)
            oFoldLineSketch.GeometricConstraints.AddCoincident(bl.StartSketchPoint, se1)
            oFoldLineSketch.GeometricConstraints.AddCoincident(bl.EndSketchPoint, se2)
            bendLine = bl
            Return bl
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetWorkFace() As Face
        Try
            Dim maxArea1, maxArea2, maxArea3 As Double

            Dim maxface1, maxface2, maxface3 As Face


            maxface1 = compDef.Features.Item(compDef.Features.Count).Faces.Item(compDef.Features.Item(compDef.Features.Count).Faces.Count)
            maxface2 = maxface1
            maxface3 = maxface2
            maxArea2 = 0
            maxArea1 = maxArea2
            maxArea3 = maxArea2

            Dim minDis, d As Double

            minDis = 99999999

            For Each f As Face In compDef.Features.Item(compDef.Features.Count).Faces
                If f.Evaluator.Area > maxArea3 Then
                    If f.Evaluator.Area > maxArea2 Then
                        If f.Evaluator.Area > maxArea1 Then
                            maxArea3 = maxArea2
                            maxArea2 = maxArea1
                            maxArea1 = f.Evaluator.Area
                            maxface3 = maxface2
                            maxface2 = maxface1
                            maxface1 = f

                        Else
                            maxArea3 = maxArea2
                            maxArea2 = f.Evaluator.Area
                            maxface3 = maxface2
                            maxface2 = f


                        End If
                    Else

                        maxArea3 = f.Evaluator.Area

                        maxface3 = f

                    End If


                    'lamp.HighLighFace(f)

                End If

            Next
            bigFaces.Add(maxface1)
            bigFaces.Add(maxface2)
            For Each f As Face In bigFaces
                For Each v As Vertex In f.Vertices
                    d = v.Point.DistanceTo(mainSketch.thirdLine.EndSketchPoint.Geometry)
                    If d < minDis Then
                        minDis = d
                        maxface2 = maxface1
                        maxface1 = f

                    Else
                        maxface2 = f
                    End If
                Next
            Next

            workface = maxface1
            initialPlane = tg.CreatePlaneByThreePoints(workface.Vertices.Item(1).Point, workface.Vertices.Item(2).Point, workface.Vertices.Item(3).Point)
            lamp.HighLighFace(workface)

            Return workface
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Public Function GetBendLine(workFace As Face, line As SketchLine3D) As SketchLine
        Try

            Dim ps As PlanarSketch
            Dim minDis As Double = 99999999999
            Dim d As Double
            ps = compDef.Sketches.Add(workFace)

            Dim sl As SketchLine
            sl = ps.AddByProjectingEntity(line)
            sl.Construction = True
            bendLine = ps.SketchLines.AddByTwoPoints(sl.EndSketchPoint.Geometry, sl.StartSketchPoint.Geometry)

            For Each ed As Edge In workFace.Edges
                d = ed.GetClosestPointTo(bendLine.StartSketchPoint.Geometry3d).DistanceTo(bendLine.StartSketchPoint.Geometry3d)
                If d < minDis Then
                    minDis = d
                    edgeBand = ed
                End If
            Next
            'lamp.HighLighObject(edgeBand)
            Return sl
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Public Function GetParameter(name As String) As Parameter
        Dim p As Parameter = Nothing
        Try
            p = doku.ComponentDefinition.Parameters.ModelParameters.Item(name)
        Catch ex As Exception
            Try
                p = doku.ComponentDefinition.Parameters.ReferenceParameters.Item(name)
            Catch ex1 As Exception
                Try
                    p = doku.ComponentDefinition.Parameters.UserParameters.Item(name)
                Catch ex2 As Exception
                    MsgBox(ex2.ToString())
                    MsgBox("Parameter not found: " & name)
                End Try

            End Try

        End Try

        Return p
    End Function
    Function GetFoldingAngle() As DimensionConstraint
        Dim ps As PlanarSketch

        Dim p1, p2, p3 As SketchPoint
        Dim v As Vector2d
        Dim r As SketchEntitiesEnumerator
        Dim sl, el As SketchLine
        Dim wp As WorkPlane
        Try

            wp = compDef.WorkPlanes.AddByNormalToCurve(bendLine, bendLine.StartSketchPoint)

            ps = doku.ComponentDefinition.Sketches.Add(wp)
            Dim fl As SketchLine3D
            fl = mainSketch.bandLines.Item(4)
            sl = ps.AddByProjectingEntity(fl)
            el = ps.AddByProjectingEntity(GetMinorEdge())
            Dim dc As DimensionConstraint
            dc = ps.DimensionConstraints.AddTwoLineAngle(el, sl, sl.EndSketchPoint.Geometry)
            dc.Driven = True
            bendAngle = dc


            Return dc

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetMajorEdge(f As Face) As Edge
        Dim e1, e2, e3 As Edge
        Dim maxe1, maxe2, maxe3 As Double
        maxe1 = 0
        maxe2 = 0
        maxe3 = 0
        e1 = f.Edges.Item(1)
        e2 = e1
        e3 = e2
        For Each ed As Edge In f.Edges
            If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) > maxe2 Then

                If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) > maxe1 Then
                    maxe3 = maxe2
                    e3 = e2
                    maxe2 = maxe1
                    e2 = e1
                    maxe1 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                    e1 = ed
                Else
                    maxe3 = maxe2
                    e3 = e2
                    maxe2 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                    e2 = ed
                End If



            End If

        Next
        'lamp.HighLighObject(e3)
        ' lamp.HighLighObject(e2)
        'lamp.HighLighObject(e1)
        bendEdge = e3
        minorEdge = e2
        majorEdge = e1
        Return e1
    End Function
    Function GetMinorEdge() As Edge
        GetMajorEdge(workface)

        Return minorEdge
    End Function
    Public Function FoldBand(i As Integer) As FoldFeature
        Try
            compDef = doku.ComponentDefinition
            sheetMetalFeatures = compDef.Features
            foldFeatures = sheetMetalFeatures.FoldFeatures


            Dim oFoldFeature As FoldFeature

            oFoldFeature = foldFeatures.Add(AdjustFoldDefinition(i))
            folded = oFoldFeature

            Return folded

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function AdjustFoldDefinition(i As Integer) As FoldDefinition
        Try

            Dim oFoldDefinition As FoldDefinition
            If i > 4 Then

                If bendAngle.Parameter._Value > Math.PI / 2 Then
                    oFoldDefinition = foldFeatures.CreateFoldDefinition(bendLine, bendAngle.Parameter._Value)
                Else
                    oFoldDefinition = foldFeatures.CreateFoldDefinition(bendLine, Math.PI - bendAngle.Parameter._Value)
                End If


            Else
                If bendAngle.Parameter._Value > Math.PI / 2 Then
                    oFoldDefinition = foldFeatures.CreateFoldDefinition(bendLine, Math.PI - bendAngle.Parameter._Value)
                Else

                    oFoldDefinition = foldFeatures.CreateFoldDefinition(bendLine, bendAngle.Parameter._Value)
                End If

            End If
            foldDefinition = oFoldDefinition
            Return oFoldDefinition
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function CheckFoldSide(ff As FoldFeature) As FoldFeature
        Try
            Dim v1, v2 As Vector
            Dim pl2 As Plane
            Dim pt2 As Point
            Dim d As Double
            v1 = initialPlane.Normal.AsVector
            pt2 = GetNextWorkFace(ff).Vertices.Item(1).Point
            pl2 = tg.CreatePlaneByThreePoints(nextworkFace.Vertices.Item(1).Point, nextworkFace.Vertices.Item(2).Point, nextworkFace.Vertices.Item(3).Point)
            v2 = pl2.Normal.AsVector
            d = (v1.CrossProduct(v2)).Length
            If Math.Abs(d) < 0.01 Then
                ff = CorrectFold(ff)
            End If
            Return ff

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function CorrectFold(ff As FoldFeature) As FoldFeature
        Try
            ff.Delete()
            foldDefinition.IsPositiveBendSide = Not foldDefinition.IsPositiveBendSide

            folded = foldFeatures.Add(foldDefinition)
            Return folded

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetNextWorkFace(ff As FoldFeature) As Face
        Try
            Dim maxArea1, maxArea2, maxArea3 As Double

            Dim maxface1, maxface2, maxface3, bf As Face

            maxface1 = ff.Faces.Item(1)
            maxface2 = maxface1
            maxface3 = maxface1
            maxArea2 = 0
            maxArea1 = maxArea2


            For Each f As Face In ff.Faces
                'lamp.HighLighFace(f)
                If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    'lamp.HighLighFace(f)
                    If f.Evaluator.Area > maxArea1 Then
                        maxArea2 = maxArea1
                        maxface2 = maxface1
                        maxArea1 = f.Evaluator.Area
                        maxface1 = f
                    Else
                        maxface2 = f
                    End If
                Else
                    maxface3 = f
                End If
            Next
            maxArea1 = 0
            maxArea2 = 0
            maxArea3 = 0
            bf = maxface2
            'lamp.HighLighFace(maxface2)
            For Each f As Face In bf.TangentiallyConnectedFaces

                If f.Evaluator.Area > maxArea3 Then
                    If f.Evaluator.Area > maxArea2 Then
                        If f.Evaluator.Area > maxArea1 Then
                            maxface3 = maxface2
                            maxface2 = maxface1
                            maxface1 = f
                            maxArea3 = maxArea2
                            maxArea2 = maxArea1
                            maxArea1 = f.Evaluator.Area
                        Else
                            maxface3 = maxface2
                            maxface2 = f
                            maxArea3 = maxArea2
                            maxArea2 = f.Evaluator.Area
                        End If
                    Else
                        maxface3 = f
                        maxArea3 = f.Evaluator.Area
                    End If
                End If
            Next
            nextWorkFace = maxface1
            lamp.HighLighFace(maxface1)
            Return nextWorkFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
End Class
