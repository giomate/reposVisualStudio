﻿Imports Inventor
Imports SecondFold
Public Class MacroFold
    Public doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Dim partNumber As Integer
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim mainSketch As Sketcher3D
    Dim nombrador As Nombres

    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim bendLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim minorEdge, majorEdge, bendEdge As Edge
    Dim minorLine, majorLine As SketchLine3D
    Dim workFace, adjacentFace As Face
    Dim bendAngle As DimensionConstraint
    Dim folded As FoldFeature
    Dim features As SheetMetalFeatures
    Dim lamp As Highlithing
    Dim bender As Doblador
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)

        mainSketch = New Sketcher3D(doku)
        compDef = doku.ComponentDefinition
        features = compDef.Features
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        thicknessCM = compDef.Thickness._Value
        gap1CM = 3 / 10
        bender = New Doblador(doku)
        nombrador = New Nombres(doku)


        done = False
    End Sub
    Public Function MakeThirdFold() As Boolean
        Try
            If GetWorkFace().SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                If mainSketch.DrawTrobinaCurve(nombrador.GetQNumber(doku), nombrador.GetNextSketchName(doku)).Construction Then
                    sk3D = mainSketch.sk3D
                    curve = mainSketch.curve
                    If GetMajorEdge(workFace).GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                        If DrawSingleLines() Then
                            Dim sl As SketchLine3D
                            sl = bandLines(1)
                            If bender.GetBendLine(workFace, sl).Length > 0 Then

                                sl = bandLines(2)
                                Dim ed As Edge
                                ed = GetMajorEdge(workFace)
                                If bandLines.Count > 4 Then
                                    ed = majorEdge
                                Else
                                    ed = minorEdge
                                End If

                                If bender.GetFoldingAngle(ed, sl).Parameter._Value > 0 Then
                                    comando.MakeInvisibleSketches(doku)
                                    comando.MakeInvisibleWorkPlanes(doku)
                                    If monitor.IsFeatureHealthy(bender.FoldBand(bandLines.Count)) Then
                                        'doku.Save2(True)
                                        done = 1
                                        Return True
                                    End If


                                End If

                            End If

                        End If
                    End If
                End If





            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try
    End Function
    Function DrawSingleLines() As Boolean
        Try
            If DrawFirstLine().Length > 0 Then
                If DrawFirstConstructionLine().Construction Then
                    If DrawSecondConstructionLine().Construction Then
                        If DrawSecondLine().Length > 0 Then
                            If DrawThirdConstructionLine().Construction Then
                                If DrawThirdLine.Length > 0 Then
                                    If DrawFouthLine.Length > 0 Then
                                        If DrawFifthLine().Length > 0 Then
                                            Return True

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
            Return False
        End Try
    End Function

    Function GetWorkFace() As Face
        Try
            Dim maxArea1, maxArea2, maxArea3 As Double

            Dim maxface1, maxface2, maxface3 As Face

            maxface1 = compDef.Features.Item(compDef.Features.Count).Faces.Item(1)
            maxface2 = maxface1
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
            workFace = maxface3
            adjacentFace = maxface3
            'lamp.HighLighFace(maxface2)
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
            For Each f As Face In compDef.Features.Item(compDef.Features.Count - 1).Faces
                For Each ft As Face In maxface2.TangentiallyConnectedFaces
                    If ft.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        If ft.Equals(f) Then

                            adjacentFace = f
                            'lamp.HighLighFace(f)

                        End If

                    End If
                Next

            Next


            'lamp.HighLighFace(workFace)

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
    Function GetStartPoint() As Point
        Dim pt As Point
        Dim v As Vector

        If majorEdge.StartVertex.Point.DistanceTo(bendEdge.StartVertex.Point) < majorEdge.StopVertex.Point.DistanceTo(bendEdge.StartVertex.Point) Then
            pt = majorEdge.StartVertex.Point
            v = majorEdge.StartVertex.Point.VectorTo(majorEdge.StopVertex.Point)
            farPoint = majorEdge.StopVertex.Point
        Else
            pt = majorEdge.StopVertex.Point
            v = majorEdge.StopVertex.Point.VectorTo(majorEdge.StartVertex.Point)
            farPoint = majorEdge.StartVertex.Point

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
            l = sk3D.SketchLines3D.AddByTwoPoints(GetStartPoint(), minorEdge.GetClosestPointTo(GetStartPoint()))

            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, majorLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, minorLine)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, majorLine.EndPoint)
            dc.Parameter._Value = majorLine.Length - thicknessCM * 5
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
            Dim minDis As Double = 9999999999
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
                If o.DistanceTo(lastLine.EndSketchPoint.Geometry) < minDis Then
                    minDis = o.DistanceTo(lastLine.EndSketchPoint.Geometry)
                    optpoint = o
                End If

            Next
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, optpoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)

            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, l.EndPoint)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, thicknessCM / 3) Then

                dc.Driven = True
                v = minorLine.Geometry.Direction.AsVector
                v.ScaleBy(-thicknessCM)
                l.StartSketchPoint.MoveBy(v)
                dc.Driven = False
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, thicknessCM / 3) Then
                    l.Construction = True
                    constructionLines.Add(l)
                End If


            End If




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

            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, minorEdge.GetClosestPointTo(firstLine.StartSketchPoint.Geometry), False)


            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, minorLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, minorLine)
            Dim dc1, dc2, dc3 As DimensionConstraint3D
            pvl = sk3D.Include(GetMajorEdge(adjacentFace))
            dc1 = sk3D.DimensionConstraints3D.AddLineLength(l)
            ol = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, pvl.Geometry.MidPoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(ol.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(ol.EndPoint, pvl)
            sk3D.GeometricConstraints3D.AddPerpendicular(ol, secondLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(ol, pvl)

            dc2 = sk3D.DimensionConstraints3D.AddLineLength(ol)
            If dc2.Parameter._Value < dc1.Parameter._Value Then
                l.Delete()
                If adjuster.AdjustDimensionConstraint3DSmothly(dc2, gap1CM) Then
                    ol.Construction = True
                    constructionLines.Add(ol)
                    lastLine = ol
                    Return ol
                End If

            Else
                ol.Delete()
                dc2.Delete()
                If adjuster.AdjustDimensionConstraint3DSmothly(dc1, gap1CM) Then
                    l.Construction = True
                    constructionLines.Add(l)
                    lastLine = l
                End If

            End If



            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawThirdLine() As SketchLine3D
        Try
            Dim l, pl As SketchLine3D

            pl = bandLines(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, pl.StartPoint, False)

            Dim gcpl3 As GeometricConstraint3D
            gcpl3 = sk3D.GeometricConstraints3D.AddPerpendicular(l, firstLine)
            lastLine = l
            bandLines.Add(l)
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
            sk3D.GeometricConstraints3D.AddEqual(l, cl)

            If AdjustlastAngle() Then
                lastLine = l
                bandLines.Add(l)
                Return l
            End If

            Return Nothing
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function AdjustlastAngle() As Boolean
        Dim fourLine, sixthLine As SketchLine3D
        Dim dc As TwoLineAngleDimConstraint3D

        Dim b As Boolean = False

        fourLine = bandLines.Item(2)
        sixthLine = bandLines.Item(4)
        dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)
        If dc.Driven Then
            Return True
        Else
            Return adjuster.AdjustDimensionConstraint3DSmothly(dc, 0.1)

        End If

        Return b
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
