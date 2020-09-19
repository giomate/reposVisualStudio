Imports Inventor

Public Class MicroFold6
    Public doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, minorLine, majorLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3 As Point
    Dim tg As TransientGeometry
    Dim initialPlane As Plane
    Dim gap1CM, thickness, bendGap As Double
    Dim partNumber, paralelLimit As Integer
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
    Dim workFace, adjacentFace, nextWorkFace As Face
    Dim bendAngle As DimensionConstraint
    Dim parallel As GeometricConstraint3D
    Dim tiltAngle As DimensionConstraint3D
    Dim folded As FoldFeature
    Dim sheetMetalFeatures As SheetMetalFeatures
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
        sheetMetalFeatures = compDef.Features
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        thickness = compDef.Thickness._Value
        bendGap = 3 * thickness / 2
        bender = New Doblador(doku)
        nombrador = New Nombres(doku)
        gap1CM = 3 / 10
        paralelLimit = 16
        done = False
    End Sub
    Public Function MakeSecondFold() As Boolean
        Try
            Return MakeEvenFold("f2")
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return False
    End Function
    Public Function MakeFourthFold() As Boolean
        Try
            Return MakeEvenFold("f4")
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return False
    End Function
    Public Function MakeSixthFold() As Boolean
        Try
            Return MakeEvenFold("f6")
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return False
    End Function
    Public Function MakeEighthFold() As Boolean
        Try
            Return MakeEvenFold("f8")
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return False
    End Function
    Public Function MakeEvenFold(s As String) As Boolean
        Try
            If GetWorkFace().SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                lamp.LookAtFace(workFace)
                If mainSketch.DrawTrobinaCurve(nombrador.GetQNumber(doku), nombrador.GetNextSketchName(doku)).Construction Then
                    sk3D = mainSketch.sk3D
                    curve = mainSketch.curve
                    lamp.ZoomSelected(curve)
                    If GetMinorEdge(workFace).GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                        If DrawFirstLine().Length > 0 Then
                            If DrawSecondLine().Length > 0 Then
                                If DrawThirdLine().Length > 0 Then
                                    If DrawFirstConstructionLine().Construction Then
                                        If IsTiltAngleOk() Then
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
                                                    folded = bender.FoldBand(bandLines.Count)
                                                    folded = CheckFoldSide(folded)
                                                    folded.Name = s
                                                    ' lamp.LookAtFace(workFace)
                                                    doku.Update2(True)
                                                    If monitor.IsFeatureHealthy(folded) Then
                                                        comando.UnfoldBand(doku)
                                                        If compDef.HasFlatPattern Then
                                                            comando.RefoldBand(doku)
                                                            doku.Save2(True)
                                                            done = 1
                                                            Return True
                                                        Else
                                                            Return False
                                                        End If

                                                    End If
                                                End If
                                            End If
                                        End If
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
            Return Nothing
        End Try


        Return False
    End Function
    Function CheckFoldSide(ff As FoldFeature) As FoldFeature
        Try
            Dim v1, v2 As Vector
            Dim pl2 As Plane
            Dim pt2 As Point
            Dim d As Double
            v1 = initialPlane.Normal.AsVector
            pt2 = GetNextWorkFace(ff).Vertices.Item(1).Point
            pl2 = tg.CreatePlaneByThreePoints(nextWorkFace.Vertices.Item(1).Point, nextWorkFace.Vertices.Item(2).Point, nextWorkFace.Vertices.Item(3).Point)
            v2 = pl2.Normal.AsVector
            d = (v1.CrossProduct(v2)).Length
            If Math.Abs(d) < 0.01 Then
                ff = bender.CorrectFold(ff)
            End If
            Return ff

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
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


            For Each f As Face In sheetMetalFeatures.FoldFeatures.Item(sheetMetalFeatures.FoldFeatures.Count).Faces
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
            lamp.HighLighFace(workFace)
            initialPlane = tg.CreatePlaneByThreePoints(workFace.Vertices.Item(1).Point, workFace.Vertices.Item(2).Point, workFace.Vertices.Item(3).Point)
            Return workFace
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
    Function GetMinorEdge(f As Face) As Edge
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
        ' lamp.HighLighObject(e3)
        ' lamp.HighLighObject(e2)
        ' lamp.HighLighObject(e1)
        bendEdge = e3
        minorEdge = e2
        majorEdge = e1
        Return e2
    End Function
    Function DrawFirstLine() As SketchLine3D
        Try

            Dim l As SketchLine3D
            Dim pt As Point = GetStartPoint()
            minorLine = sk3D.Include(minorEdge)
            majorLine = sk3D.Include(majorEdge)
            l = sk3D.SketchLines3D.AddByTwoPoints(pt, majorEdge.GetClosestPointTo(pt))

            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, minorLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, majorLine)

            Dim dc As DimensionConstraint3D

            If farPoint.DistanceTo(minorLine.EndSketchPoint.Geometry) < gap1CM Then
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, minorLine.EndPoint)
            Else
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, minorLine.StartPoint)
            End If

            dc.Parameter._Value = minorLine.Length - bendGap
            sk3D.Solve()
            bandLines.Add(l)
            firstLine = l
            lastLine = l
            ' lamp.FitView(doku)



            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function DrawSecondLine() As SketchLine3D
        Try
            Dim v, vmnl, vc As Vector
            v = lastLine.Geometry.Direction.AsVector()
            Dim p As Plane
            Dim optpoint As Point = Nothing
            Dim f As Double
            p = tg.CreatePlane(lastLine.EndSketchPoint.Geometry, v)
            Dim minDis As Double = 9999999999
            Dim d As Double
            vmnl = firstLine.StartSketchPoint.Geometry.VectorTo(farPoint)
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
                vc = lastLine.EndSketchPoint.Geometry.VectorTo(o)
                d = vmnl.CrossProduct(vc).Length * vc.DotProduct(vmnl)
                If d > 0 Then
                    If o.DistanceTo(lastLine.EndSketchPoint.Geometry) < minDis Then
                        minDis = o.DistanceTo(lastLine.EndSketchPoint.Geometry)
                        optpoint = o
                    End If
                End If
            Next
            Dim l As SketchLine3D = Nothing
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, optpoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(firstLine, minorLine)

            If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 9 / 10) Then
            Else
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 11 / 12)
            End If
            dc.Delete()
            dc = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 7 / 6) Then
            Else
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 9 / 8)
            End If
            dc.Delete()

            If sk3D.Name = "s8" Then
                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(firstLine, minorLine)
                Try

                    adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 4 / 5)
                Catch ex As Exception

                End Try
                dc.Delete()

            End If
            tiltAngle = sk3D.DimensionConstraints3D.AddTwoLineAngle(firstLine, minorLine)
            tiltAngle.Driven = True
            If l.Length > 35 / 10 Then
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                adjuster.AdjustDimConstrain3DSmothly(dc, 25 / 10)
                dc.Delete()
            End If
            point3 = l.EndSketchPoint.Geometry
            lastLine = l
            bandLines.Add(l)
            secondLine = lastLine


            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawThirdLine() As SketchLine3D
        Try
            Dim dc, ac As DimensionConstraint3D
            Dim gc As GeometricConstraint3D
            direction = lastLine.Geometry.Direction.AsVector()
            direction.ScaleBy(thickness * 10)
            Dim pt As Point
            Dim climit As Integer = 0
            pt = firstLine.StartSketchPoint.Geometry
            pt.TranslateBy(direction)
            Dim l As SketchLine3D = Nothing
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.StartPoint, pt, False)
            Try
                gc = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            Catch ex As Exception
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 2)
                dc.Driven = True
                gc = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            End Try
            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, minorLine)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, 1 / 32)
                ac.Delete()

            Catch ex As Exception
                Try
                    ac.Delete()
                Catch ex2 As Exception

                End Try
            End Try

            TryParallel(l)




            lastLine = l
            bandLines.Add(l)
            thirdLine = lastLine


            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            parallel.driven = True
            Return Nothing
        End Try
    End Function
    Function DrawFirstConstructionLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing
            l = sk3D.SketchLines3D.AddByTwoPoints(thirdLine.EndPoint, firstLine.EndSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, secondLine)
            Dim dc, ac As DimensionConstraint3D
            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(secondLine, l)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()
                sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            Catch ex As Exception
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(secondLine, l)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, 3 * Math.PI / 8)
                ac.Delete()
                sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            End Try


            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
                dc.Parameter._Value = GetParameter("b")._Value
            Else
                dc.Driven = True
                Try
                    parallel.Delete()
                Catch ex As Exception

                End Try

                dc.Driven = False
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value) Then
                    dc.Parameter._Value = GetParameter("b")._Value
                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(thirdLine, l)
                    If adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2) Then
                    Else
                        ac.Driven = True
                    End If
                    ac.Delete()
                Else
                    dc.Driven = True

                End If
                TryParallel(thirdLine)
            End If
            Try
                sk3D.Solve()
                If monitor.IsSketch3DHealthy(sk3D) Then
                Else
                    Try
                        parallel.Delete()
                    Catch ex As Exception
                    End Try
                    TryParallel(thirdLine)
                End If
            Catch ex As Exception

            End Try
            l.Construction = True
            constructionLines.Add(l)
            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function IsTiltAngleOk() As Boolean
        Try
            Try
                sk3D.Solve()
                If monitor.IsSketch3DHealthy(sk3D) Then
                Else
                    Try
                        parallel.Delete()
                    Catch ex As Exception
                    End Try
                    TryParallel(thirdLine)
                End If
            Catch ex As Exception

            End Try

            If tiltAngle.Parameter._Value > (4 * Math.PI / 6) Then
                Try
                    parallel.Delete()
                    adjuster.AdjustDimensionConstraint3DSmothly(tiltAngle, Math.PI / 2)
                    Return True
                Catch ex As Exception
                    Return adjuster.AdjustDimensionConstraint3DSmothly(tiltAngle, Math.PI / 2)

                End Try
            Else
                Return True

            End If

        Catch ex As Exception

        End Try

        Return 0
    End Function
    Function TryParallel(l As SketchLine3D) As GeometricConstraint3D
        Try
            parallel = sk3D.GeometricConstraints3D.AddParallel(l, secondLine)
            Return parallel
        Catch ex As Exception
            Dim errorCounter As Integer = 0
            Dim ac1, dcl2 As DimensionConstraint3D
            Dim cl As SketchLine3D
            cl = sk3D.SketchLines3D.AddByTwoPoints(firstLine.StartPoint, farPoint)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddParallel(cl, secondLine)
            ac1 = sk3D.DimensionConstraints3D.AddTwoLineAngle(cl, l)
            ac1.Driven = True
            dcl2 = sk3D.DimensionConstraints3D.AddLineLength(secondLine)
            dcl2.Driven = True
            If ac1.Parameter._Value < Math.PI / 2 Then
                While (ac1.Parameter._Value > 0 And errorCounter < paralelLimit)
                    Try
                        ac1.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(ac1, ac1.Parameter._Value / 4)
                        ac1.Driven = True
                    Catch ex2 As Exception
                        ac1.Driven = True
                        errorCounter = errorCounter + 1
                    End Try
                    Try
                        dcl2.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl2, dcl2.Parameter._Value * (paralelLimit - 4) / paralelLimit)
                        dcl2.Driven = True
                    Catch ex3 As Exception
                        dcl2.Driven = True
                        errorCounter = errorCounter + 1
                    End Try
                    Try
                        parallel = sk3D.GeometricConstraints3D.AddParallel(l, secondLine)
                        errorCounter = paralelLimit + 1
                    Catch ex3 As Exception
                        errorCounter = errorCounter + 1
                    End Try
                    errorCounter = errorCounter + 1
                End While
                errorCounter = 0
            Else
                While (ac1.Parameter._Value < Math.PI And errorCounter < paralelLimit)
                    Try
                        ac1.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(ac1, Math.PI)
                        ac1.Driven = True
                    Catch ex2 As Exception
                        ac1.Driven = True
                        errorCounter = errorCounter + 1
                    End Try
                    Try
                        dcl2.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl2, dcl2.Parameter._Value * (paralelLimit - 4) / paralelLimit)
                        dcl2.Driven = True
                    Catch ex3 As Exception
                        dcl2.Driven = True
                        errorCounter = errorCounter + 1
                    End Try
                    Try
                        parallel = sk3D.GeometricConstraints3D.AddParallel(l, secondLine)
                        errorCounter = paralelLimit + 1
                    Catch ex3 As Exception
                        errorCounter = errorCounter + 1
                    End Try
                    errorCounter = errorCounter + 1
                End While
                errorCounter = 0

            End If
            cl.Construction = True
            ac1.Driven = True
            Try
                parallel = sk3D.GeometricConstraints3D.AddParallel(l, secondLine)
            Catch ex3 As Exception
            End Try
            dcl2.Delete()
            Return gc
        End Try
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


        v.AsUnitVector.AsVector()
        v.ScaleBy(bendGap / 10)
        pt.TranslateBy(v)
        'lamp.HighLighObject(pt)
        point1 = pt
        Return pt
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
