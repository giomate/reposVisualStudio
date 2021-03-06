﻿Imports Inventor
Imports ThirdFold
Imports FourthFold

Public Class TwistFold7
    Dim doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, connectLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim initialPlane As Plane
    Dim gap1CM, thicknessCM As Double
    Dim partNumber As Integer
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim curvas3D As Curves3D
    Dim nombrador As Nombres
    Dim nextSketch As OriginSketch
    Dim cutProfile As Profile

    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, cutEdge1, cutEdge2, CutEsge3 As Edge
    Dim minorLine, majorLine, cutLine3D, kante3D, tante3D As SketchLine3D
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, twistFace, nextworkFace As Face
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex As DimensionConstraint3D
    Dim folded As FoldFeature
    Dim features As SheetMetalFeatures
    Dim lamp As Highlithing
    Dim bender As Doblador
    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane

    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        curvas3D = New Curves3D(doku)

        compDef = doku.ComponentDefinition
        features = compDef.Features
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        sections = app.TransientObjects.CreateObjectCollection
        esquinas = app.TransientObjects.CreateObjectCollection
        rails = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        thicknessCM = compDef.Thickness._Value
        gap1CM = 3 / 10
        bender = New Doblador(doku)
        nombrador = New Nombres(doku)


        done = False
    End Sub
    Public Function MakeFinalTwist() As Boolean
        Try
            If monitor.IsFeatureHealthy(MakeLastFold()) Then
                If monitor.IsFeatureHealthy(CutLastFace(LastCutProfil())) Then


                    If GetTwistFace().Evaluator.Area > 0 Then
                        If GetTwistProfile().Count > 0 Then
                            If GetRails().Count > 0 Then
                                If monitor.IsFeatureHealthy(MakeFinalLoft()) Then
                                    doku.Update2(True)
                                    If MakeReferenceSketch().Length > 0 Then
                                        comando.MakeInvisibleSketches(doku)
                                        comando.MakeInvisibleWorkPlanes(doku)
                                        doku.Save()
                                        done = 1
                                    End If

                                End If
                            End If

                        End If


                    End If




                End If

            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return 0
    End Function
    Function MakeReferenceSketch() As SketchLine3D
        Try
            Dim l As SketchLine3D
            sk3D = compDef.Sketches3D.Add()
            l = sk3D.Include(connectLine)
            sk3D.Name = "last"
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetRails() As Profile3D
        Try
            Dim l1, l2, l3 As SketchLine3D
            Dim sp As SketchPoint
            Dim d As Double = 999999999
            Dim i, j, k As Integer
            Dim pr3d As Profile3D

            sk3D = compDef.Sketches3D.Add()
            i = 0

            If cutLine.StartSketchPoint.Geometry.DistanceTo(bendLine.EndSketchPoint.Geometry) < cutLine.EndSketchPoint.Geometry.DistanceTo(bendLine.EndSketchPoint.Geometry) Then
                For Each v As Vertex In twistFace.Vertices
                    i = i + 1
                    If cutLine3D.StartSketchPoint.Geometry.DistanceTo(v.Point) < d Then
                        d = cutLine3D.StartSketchPoint.Geometry.DistanceTo(v.Point)
                        j = i
                    End If
                Next
                For index = j To j + 3
                    Math.DivRem(index, 4, k)
                    sp = esquinas.Item(k + 1)
                    sk3D = compDef.Sketches3D.Add()
                    l1 = sk3D.SketchLines3D.AddByTwoPoints(twistFace.Vertices.Item(k + 1).Point, sp.Geometry3d)
                    pr3d = sk3D.Profiles3D.AddOpen
                    rails.Add(pr3d)
                Next

            Else
                For Each v As Vertex In twistFace.Vertices
                    i = i + 1
                    If cutLine3D.EndSketchPoint.Geometry.DistanceTo(v.Point) < d Then
                        d = cutLine3D.StartSketchPoint.Geometry.DistanceTo(v.Point)
                        j = i
                    End If
                Next
                For index = j To j + 3
                    Math.DivRem(index, 4, k)
                    sp = esquinas.Item(k + 1)
                    sk3D = compDef.Sketches3D.Add()
                    l1 = sk3D.SketchLines3D.AddByTwoPoints(twistFace.Vertices.Item(k + 1).Point, sp.Geometry3d)
                    pr3d = sk3D.Profiles3D.AddOpen
                    rails.Add(pr3d)
                Next


            End If


            Return pr3d

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return Nothing
    End Function
    Function CreateRail() As Integer

        Return 0
    End Function
    Public Function MakeLastFold() As FoldFeature
        Try
            If GetWorkFace().SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                If curvas3D.DrawTrobinaCurve(nombrador.GetQNumber(doku), nombrador.GetNextSketchName(doku), 1.05).Construction Then
                    sk3D = curvas3D.sk3D
                    curve = curvas3D.curve
                    If GetMajorEdge(workFace).GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                        If DrawSingleLines().Length > 0 Then

                            nextSketch = New OriginSketch(doku)
                            Dim cl As SketchLine3D
                            cl = constructionLines.Item(1)
                            connectLine = nextSketch.DrawNextMainSketch(refLine, cl.EndSketchPoint.Geometry)
                            If connectLine.Length > 0 Then
                                If ConnectTwistBrigde() Then
                                    Dim sl As SketchLine3D
                                    sl = bandLines(1)
                                    If bender.GetBendLine(workFace, sl).Length > 0 Then
                                        bendLine = bender.bendLine
                                        sl = bandLines(2)
                                        Dim ed As Edge
                                        ed = GetMajorEdge(workFace)

                                        ed = minorEdge

                                        If bender.GetFoldingAngle(ed, sl).Parameter._Value > 0 Then
                                            comando.MakeInvisibleSketches(doku)
                                            comando.MakeInvisibleWorkPlanes(doku)
                                            folded = bender.FoldBand(bandLines.Count)
                                            folded = CheckFoldSide(folded)
                                            folded.Name = "f5"
                                            doku.Update2(True)
                                            If monitor.IsFeatureHealthy(folded) Then
                                                foldFeature = bender.folded

                                                Return foldFeature
                                            End If


                                        End If

                                    End If
                                End If
                            End If

                        End If
                    End If
                End If





            End If
            Return Nothing
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
            nextworkFace = maxface1
            lamp.HighLighFace(maxface1)
            Return nextworkFace
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
                'ff = bender.CorrectFold(ff)
            End If
            Return ff

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetTwistFace() As Face
        Try
            Dim pr As Profile
            Dim ps As PlanarSketch
            Dim el, sl As SketchLine
            Dim minArea As Double = 99999999
            Dim minDis As Double = 99999999

            For Each f As Face In cutfeature.Faces
                If f.Evaluator.Area < minArea Then
                    twistFace = f
                    minArea = f.Evaluator.Area

                End If
                'lamp.HighLighFace(f)

            Next
            ps = doku.ComponentDefinition.Sketches.Add(twistFace)
            For Each ed As Edge In twistFace.Edges

                el = ps.AddByProjectingEntity(ed)
                el.Construction = True

                If el.EndSketchPoint.Geometry3d.Equals(cutLine3D.StartSketchPoint.Geometry) Then

                    kante3D = sk3D.Include(el)


                End If
                sl = ps.SketchLines.AddByTwoPoints(el.StartSketchPoint, el.EndSketchPoint)
            Next
            ' ps.SketchLines.AddAsThreePointRectangle(twistFace.Vertices.Item(1).Poin, twistFace.Vertices.Item(2).Point, twistFace.Vertices.Item(3).Point)



            pr = ps.Profiles.AddForSolid
            sections.Add(pr)
            Return twistFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function
    Function GetTwistProfile() As Profile
        Try
            Dim l, t, il, sl As SketchLine3D
            Dim pr As Profile
            Dim ps As PlanarSketch
            Dim l2d, il2d As SketchLine
            Dim v As Vector
            Dim pt As Point
            Dim r As Boolean

            sk3D = compDef.Sketches3D.Add()

            il = nextSketch.inputLine
            sl = nextSketch.secondLine
            t = nextSketch.constructionLines.Item(2)
            v = t.Geometry.Direction.AsVector
            v.ScaleBy(-1)

            il = sk3D.SketchLines3D.AddByTwoPoints(il.EndPoint, il.StartPoint, False)
            sl = sk3D.SketchLines3D.AddByTwoPoints(sl.StartPoint, sl.EndPoint, False)
            pt = sl.Geometry.MidPoint
            pt.TranslateBy(v)
            l = sk3D.SketchLines3D.AddByTwoPoints(il.StartPoint, pt, False)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, sl)
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, il)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, gap1CM / 10) Then
                tante3D = l
                twistPlane = doku.ComponentDefinition.WorkPlanes.AddByThreePoints(il.StartPoint, il.EndPoint, l.EndPoint)

                ps = doku.ComponentDefinition.Sketches.Add(twistPlane)
                il2d = ps.AddByProjectingEntity(il)
                esquinas.Add(il2d.EndSketchPoint)
                esquinas.Add(il2d.StartSketchPoint)
                l2d = ps.AddByProjectingEntity(l)
                esquinas.Add(l2d.EndSketchPoint)
                ps.SketchLines.AddAsThreePointRectangle(il2d.EndSketchPoint, il2d.StartSketchPoint, l2d.EndSketchPoint.Geometry)

                esquinas.Add(ps.SketchPoints.Item(ps.SketchPoints.Count))



                pr = ps.Profiles.AddForSolid

                sections.Add(pr)
                Return pr
            Else
                Return Nothing
            End If
            Return Nothing
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return Nothing
    End Function
    Function LastCutProfil() As Profile


        Try
            Dim maxArea1 As Double
            Dim pro As Profile
            Dim maxface1 As Face
            maxArea1 = 0

            'lamp.HighLighFace(maxface2)
            For Each f As Face In frontBendFace.TangentiallyConnectedFaces

                If f.Evaluator.Area > maxArea1 Then
                    maxArea1 = f.Evaluator.Area
                    maxface1 = f
                End If
            Next
            cutFace = maxface1
            'lamp.HighLighFace(maxface1)
            Dim ps As PlanarSketch

            ps = compDef.Sketches.Add(cutFace)


            Dim sl, fl, l, r, u, p, cl As SketchLine
            sl = ps.AddByProjectingEntity(secondLine)
            sl.Construction = True
            fl = ps.AddByProjectingEntity(firstLine)
            fl.Construction = True
            u = ps.AddByProjectingEntity(GetCutEdges(cutFace))
            r = ps.AddByProjectingEntity(cutEdge1)
            l = ps.AddByProjectingEntity(cutEdge2)
            p = ps.SketchLines.AddByTwoPoints(sl.Geometry.StartPoint, fl.Geometry.EndPoint)
            p.Construction = True
            cl = ps.SketchLines.AddByTwoPoints(p.StartSketchPoint.Geometry, p.EndSketchPoint.Geometry)
            ps.GeometricConstraints.AddCoincident(cl.EndSketchPoint, l)
            If l.EndSketchPoint.Geometry.DistanceTo(p.Geometry.EndPoint) < l.StartSketchPoint.Geometry.DistanceTo(p.Geometry.EndPoint) Then
                ps.GeometricConstraints.AddCoincident(l.EndSketchPoint, cl)
            Else
                ps.GeometricConstraints.AddCoincident(l.StartSketchPoint, cl)
            End If

            ps.GeometricConstraints.AddCoincident(cl.StartSketchPoint, r)
            ps.GeometricConstraints.AddParallel(cl, p)

            cutLine = cl
            cutLine3D = sk3D.Include(cl)
            pro = ps.Profiles.AddForSolid

            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function

    Function MakeFinalLoft() As LoftFeature

        Dim oLoftDefinition As LoftDefinition
        oLoftDefinition = compDef.Features.LoftFeatures.CreateLoftDefinition(sections, PartFeatureOperationEnum.kNewBodyOperation)
        For Each p As Profile3D In rails
            oLoftDefinition.LoftRails.Add(p)
        Next

        Dim lf As LoftFeature
        lf = compDef.Features.LoftFeatures.Add(oLoftDefinition)
        Return lf
    End Function
    Function CutLastFace(pr As Profile) As CutFeature
        Try
            Dim oSheetMetalCompDef As SheetMetalComponentDefinition
            oSheetMetalCompDef = doku.ComponentDefinition
            Dim oSheetMetalFeatures As SheetMetalFeatures
            oSheetMetalFeatures = doku.ComponentDefinition.Features

            compDef = oSheetMetalCompDef
            Dim oFaceFeatureDefinition As CutDefinition
            oFaceFeatureDefinition = oSheetMetalFeatures.CutFeatures.CreateCutDefinition(pr)
            ' oFaceFeatureDefinition.Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection
            Dim oFaceFeature As CutFeature
            oFaceFeature = oSheetMetalFeatures.CutFeatures.Add(oFaceFeatureDefinition)
            cutfeature = oFaceFeature
            Return oFaceFeature
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing

        End Try


    End Function
    Function ConnectTwistBrigde() As Boolean
        Try
            For Each sl As SketchLine3D In sk3D.SketchLines3D
                If sl.Equals(refLine) Then
                    sl.Delete()
                    Exit For
                End If
            Next
            Dim tl1, tl2 As SketchLine3D
            Dim dc1, dc2 As DimensionConstraint3D
            Dim gc As GeometricConstraint3D

            tl1 = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, connectLine.StartPoint, False)
            tl2 = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartPoint, nextSketch.secondLine.StartSketchPoint.Geometry, False)
            gc = sk3D.GeometricConstraints3D.AddCoincident(tl2.EndPoint, nextSketch.secondLine)
            dc1 = sk3D.DimensionConstraints3D.AddLineLength(tl1)
            dc2 = sk3D.DimensionConstraints3D.AddLineLength(tl2)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc2, tl1.Length) Then
                bandLines.Add(tl1)
                bandLines.Add(tl2)
                Return True
            Else
                dc2.Delete()
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawSingleLines() As SketchLine3D
        Try
            If DrawFirstLine().Length > 0 Then
                If DrawFirstConstructionLine().Construction Then
                    If DrawSecondConstructionLine().Construction Then
                        If DrawSecondLine().Length > 0 Then
                            If DrawThirdConstructionLine().Construction Then

                                If DrawThirdLine.Length > 0 Then
                                    Return bandLines.Item(3)
                                    If DrawFouthLine.Length > 0 Then
                                        If DrawFifthLine().Length > 0 Then
                                            Return bandLines.Item(3)

                                        End If

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function

    Function GetWorkFace() As Face
        Try
            Dim maxArea1, maxArea2, maxArea3 As Double

            Dim maxface1, maxface2, maxface3 As Face


            maxface2 = compDef.Bends.Item(compDef.Bends.Count).BackFaces.Item(1)
            maxface1 = maxface2
            maxArea2 = 0
            maxArea1 = maxArea2





            For Each f As Face In compDef.Features.Item(compDef.Features.Count).Faces
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
            frontBendFace = maxface1
            workFace = maxface3
            adjacentFace = maxface3
            bendFace = maxface2
            'lamp.HighLighFace(bendFace)
            For Each f As Face In maxface2.TangentiallyConnectedFaces

                If f.Evaluator.Area > maxArea3 Then
                    If f.Evaluator.Area > maxArea2 Then
                        If f.Evaluator.Area > maxArea1 Then
                            maxface3 = adjacentFace
                            adjacentFace = workFace
                            workFace = f
                            maxArea3 = maxArea2
                            maxArea2 = maxArea1
                            maxArea1 = f.Evaluator.Area
                        Else
                            maxface3 = adjacentFace
                            adjacentFace = f
                            maxArea3 = maxArea2
                            maxArea2 = f.Evaluator.Area
                        End If
                    Else
                        maxface3 = f
                        maxArea3 = f.Evaluator.Area
                    End If
                End If
            Next
            For Each fb As Face In compDef.Bends.Item(compDef.Bends.Count - 1).BackFaces
                For Each f As Face In fb.TangentiallyConnectedFaces
                    If ((Not f.Equals(workFace)) And (f.SurfaceType = SurfaceTypeEnum.kPlaneSurface)) Then
                        For Each vbf1 As Vertex In f.Vertices
                            For Each vbf0 As Vertex In maxface2.Vertices
                                If vbf0.Equals(vbf1) Then
                                    adjacentFace = f
                                    'lamp.HighLighFace(f)

                                End If
                            Next
                        Next

                    End If

                Next
            Next

            'lamp.HighLighFace(workFace)
            initialPlane = tg.CreatePlaneByThreePoints(workFace.Vertices.Item(1).Point, workFace.Vertices.Item(2).Point, workFace.Vertices.Item(3).Point)
            Return workFace
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
    Function GetCutEdges(f As Face) As Edge
        Dim e1, e2, e3, e4 As Edge
        Dim maxe1, maxe2, maxe3, maxe4 As Double
        maxe1 = 0
        maxe2 = 0
        maxe3 = 0
        maxe4 = 0
        e1 = f.Edges.Item(1)
        e2 = e1
        e3 = e2
        e4 = e3
        For Each ed As Edge In f.Edges
            If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) > maxe3 Then
                If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) > maxe2 Then

                    If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) > maxe1 Then
                        maxe4 = maxe3
                        e4 = e3
                        maxe3 = maxe2
                        e3 = e2
                        maxe2 = maxe1
                        e2 = e1
                        maxe1 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                        e1 = ed
                    Else
                        maxe4 = maxe3
                        e4 = e3
                        maxe3 = maxe2
                        e3 = e2
                        maxe2 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                        e2 = ed
                    End If

                Else
                    maxe4 = maxe3
                    e4 = e3
                    maxe3 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                    e3 = ed

                End If
            Else
                maxe4 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                e4 = ed
            End If


        Next
        'lamp.HighLighObject(e4)
        'lamp.HighLighObject(e2)
        'lamp.HighLighObject(e1)
        cutEdge1 = e1
        cutEdge2 = e2
        CutEsge3 = e4
        Return e4
    End Function
    Function GetStartPoint() As Point
        Dim pt As Point
        Dim v As Vector

        If minorEdge.StartVertex.Point.DistanceTo(bendEdge.StartVertex.Point) < minorEdge.StopVertex.Point.DistanceTo(bendEdge.StartVertex.Point) Then
            pt = minorEdge.StartVertex.Point
            v = minorEdge.StartVertex.Point.VectorTo(minorEdge.StopVertex.Point)
            farPoint = minorEdge.StopVertex.Point
        Else
            pt = minorEdge.StopVertex.Point
            v = minorEdge.StopVertex.Point.VectorTo(minorEdge.StartVertex.Point)
            farPoint = minorEdge.StartVertex.Point

        End If
        'lamp.HighLighObject(pt)

        v.AsUnitVector.AsVector()
        v.ScaleBy(thicknessCM * 5 / 10)
        pt.TranslateBy(v)
        'lamp.HighLighObject(pt)
        point1 = pt
        Return pt
    End Function
    Function DrawFirstLine() As SketchLine3D
        Try

            Dim l As SketchLine3D
            minorLine = sk3D.Include(minorEdge)
            majorLine = sk3D.Include(majorEdge)
            l = sk3D.SketchLines3D.AddByTwoPoints(GetStartPoint(), majorEdge.GetClosestPointTo(GetStartPoint()))
            't = sk3D.SketchLines3D.AddByTwoPoints(point1, minorLine.EndSketchPoint)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, minorLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, majorLine)

            Dim dc As DimensionConstraint3D

            If farPoint.DistanceTo(minorLine.EndSketchPoint.Geometry) < gap1CM Then
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, minorLine.EndPoint)
            Else
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, minorLine.StartPoint)
            End If


            dc.Parameter._Value = minorLine.Length - thicknessCM * 5
            doku.Update2(True)
            bandLines.Add(l)
            firstLine = l
            lastLine = l

            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawFirstConstructionLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing


            Dim v As Vector = firstLine.Geometry.Direction.AsVector
            Dim p As Plane
            Dim optpoint As Point = Nothing
            p = tg.CreatePlane(firstLine.EndSketchPoint.Geometry, v)
            Dim d As Double
            Dim minDis As Double = 999999999999999
            Dim vc, vmjl As Vector
            vmjl = firstLine.EndSketchPoint.Geometry.VectorTo(farPoint)
            ' l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, farPoint, False)dmm
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
                If o.DistanceTo(firstLine.EndSketchPoint.Geometry) < curvas3D.Cr / 10 Then

                    vc = firstLine.EndSketchPoint.Geometry.VectorTo(o)
                    d = vc.CrossProduct(vmjl).Length * vc.DotProduct(vmjl)
                    If d > 0 Then
                        'l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, o, False)
                        If d < minDis Then
                            minDis = d
                            optpoint = o
                        End If
                    End If

                End If


            Next
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, optpoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)

            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, l.EndPoint)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, thicknessCM / 3) Then

                dc.Driven = True
                v = firstLine.StartSketchPoint.Geometry.VectorTo(farPoint).AsUnitVector.AsVector
                v.ScaleBy(thicknessCM * 5)
                l.StartSketchPoint.MoveBy(v)
                dc.Driven = False
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, thicknessCM / 3) Then
                    l.Construction = True
                    constructionLines.Add(l)
                End If
            Else
                dc.Delete()
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                l.Construction = True
                constructionLines.Add(l)

            End If

            gapVertex = dc


            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawSecondConstructionLine() As SketchLine3D
        Try
            Dim l, ol As SketchLine3D
            Dim endPoint As Point
            Dim v As Vector
            Dim pl As Plane
            pl = tg.CreatePlaneByThreePoints(minorLine.EndSketchPoint.Geometry, minorLine.StartSketchPoint.Geometry, majorLine.StartSketchPoint.Geometry)
            v = pl.Normal.AsVector()
            v.ScaleBy(GetParameter("b")._Value / 1)
            endPoint = firstLine.EndSketchPoint.Geometry
            endPoint.TranslateBy(v)
            ol = constructionLines(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, endPoint, False)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, firstLine)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
                gc.Delete()
                l.Construction = True
                constructionLines.Add(l)

            End If





            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawSecondLine() As SketchLine3D
        Try

            Dim l, ol As SketchLine3D
            ol = constructionLines(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(ol.EndSketchPoint.Geometry, firstLine.StartPoint, False)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, ol)
            gc = sk3D.GeometricConstraints3D.AddCoincident(ol.EndPoint, l)
            lastLine = l
            bandLines.Add(l)
            secondLine = lastLine


            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawThirdConstructionLine() As SketchLine3D
        Try

            Dim l, ol, pvl, l3 As SketchLine3D

            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, majorEdge.GetClosestPointTo(firstLine.StartSketchPoint.Geometry), False)


            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, majorLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, minorLine)
            Dim dc1, dc2 As DimensionConstraint3D
            pvl = sk3D.Include(GetAdjacentEdge())
            dc1 = sk3D.DimensionConstraints3D.AddLineLength(l)
            ol = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, pvl.Geometry.MidPoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(ol.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(ol.EndPoint, pvl)
            sk3D.GeometricConstraints3D.AddPerpendicular(ol, secondLine)
            'sk3D.GeometricConstraints3D.AddPerpendicular(ol, pvl)

            dc2 = sk3D.DimensionConstraints3D.AddLineLength(ol)
            If dc2.Parameter._Value < dc1.Parameter._Value Then
                l.Delete()
                If adjuster.AdjustDimensionConstraint3DSmothly(dc2, gap1CM * 3) Then
                    gapFold = dc2
                    ol.Construction = True
                    constructionLines.Add(ol)
                    lastLine = ol
                    Return ol
                End If

            Else
                ol.Delete()

                If adjuster.AdjustDimensionConstraint3DSmothly(dc1, gap1CM * 1) Then
                    gapFold = dc1
                    l.Construction = True
                    constructionLines.Add(l)
                    lastLine = l
                End If
                gapVertex.Driven = True
            End If



            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetAdjacentEdge() As Edge
        Try
            Dim m1, m2 As Double
            Dim va, vb As Vector

            m1 = 0
            vb = minorEdge.StartVertex.Point.VectorTo(minorEdge.StopVertex.Point)
            For Each eda As Edge In adjacentFace.Edges
                va = eda.StartVertex.Point.VectorTo(eda.StopVertex.Point)
                If va.DotProduct(vb) > m1 Then
                    m2 = m1
                    adjacentEdge = eda
                End If
            Next

            Return adjacentEdge
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawThirdLine() As SketchLine3D
        Try
            Dim l, cl, cl2 As SketchLine3D
            Dim ac As DimensionConstraint3D
            Dim dc As DimensionConstraint3D

            cl = constructionLines.Item(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartPoint, cl.EndSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            cl2 = constructionLines.Item(3)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, cl2)
            gapVertex.Driven = False
            If adjuster.GetMinimalDimension(gapVertex) Then


                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, l.Length * 2) Then
                    dc.Driven = True
                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl2)
                    If adjuster.GetMinimalDimension(ac) Then
                        ac.Driven = True
                    Else
                        ac.Delete()
                        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl2)
                        ac.Driven = True
                    End If
                    dc.Driven = True
                Else
                    dc.Delete()
                    dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(secondLine, l)
                    dc.Driven = True
                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl2)
                    If adjuster.GetMinimalDimension(ac) Then
                        ac.Driven = True
                    Else
                        ac.Delete()
                        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl2)
                        ac.Driven = True
                    End If
                    dc.Driven = True
                End If

            End If
            dc.Driven = False
            doku.Update2(True)

            lastLine = l
            bandLines.Add(l)
            refLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawFouthLine() As SketchLine3D
        Try
            Dim l, pl As SketchLine3D
            Dim v As Vector
            Dim p As Plane
            Dim minDis As Double
            Dim optpoint As Point
            pl = bandLines.Item(3)
            v = pl.Geometry.Direction.AsVector
            p = tg.CreatePlane(pl.EndSketchPoint.Geometry, v)

            Dim puntos As ObjectsEnumerator
            puntos = p.IntersectWithCurve(curve.Geometry)
            minDis = 9999999999
            Dim o2 As Point = puntos.Item(puntos.Count)
            optpoint = o2
            For Each o As Point In puntos
                If o.DistanceTo(pl.EndSketchPoint.Geometry) < minDis Then
                    minDis = o.DistanceTo(pl.EndSketchPoint.Geometry)
                    o2 = optpoint
                    optpoint = o
                End If

            Next


            l = sk3D.SketchLines3D.AddByTwoPoints(pl.EndSketchPoint.Geometry, optpoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(pl.EndPoint, l)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawFifthLine() As SketchLine3D
        Try
            Dim l, cl As SketchLine3D
            cl = constructionLines.Item(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.StartPoint, firstLine.EndPoint, False)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(lastLine)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 4 / 3) Then
                dc.Delete()
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
                    dc.Delete()
                    sk3D.GeometricConstraints3D.AddEqual(l, cl)

                End If
            End If




            lastLine = l
            bandLines.Add(l)
            Return l


            Return Nothing
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function AdjustLastAngle() As Boolean
        Try
            Dim fourLine, sixthLine, cl3 As SketchLine3D
            Dim dc As TwoLineAngleDimConstraint3D
            Dim limit As Double = 0.1

            Dim b As Boolean = False

            fourLine = bandLines.Item(2)
            sixthLine = bandLines.Item(4)
            cl3 = constructionLines.Item(3)
            dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)

            If adjuster.AdjustGapSmothly(gapFold, gap1CM, dc) Then
                b = True
            Else
                dc.Driven = True
                gapFold.Delete()
                gapFold = sk3D.DimensionConstraints3D.AddLineLength(cl3)
                b = True
            End If
            Return b
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
End Class
