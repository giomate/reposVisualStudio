Imports Inventor
Imports Subina_Design_Helpers


Public Class MacroFold5
    Public doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, bendLine3D As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Public direction As Integer
    Dim monitor As DesignMonitoring

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim initialPlane, adjacentPlane As Plane
    Dim gap1CM, thicknessCM As Double
    Dim partNumber As Integer
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim mainSketch As Sketcher3D
    Dim nombrador As Nombres
    Dim curvas3D As Curves3D
    Dim pro As Profile
    Dim feature As FaceFeature
    Dim bendLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, leadingEdge, followEdge As Edge
    Dim minorLine, majorLine As SketchLine3D
    Dim workFace, adjacentFace, bendFace, nextworkFace As Face
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex As DimensionConstraint3D
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
        thicknessCM = compDef.Thickness._Value
        gap1CM = 3 / 10
        bender = New Doblador(doku)
        nombrador = New Nombres(doku)
        curvas3D = New Curves3D(doku)


        done = False
    End Sub
    Public Function MakeThirdFold() As Boolean
        Try
            Return MakeOddFold("f3")
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return False
    End Function
    Public Function MakeFifthFold() As Boolean
        Try
            Return MakeOddFold("f5")
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return False
    End Function
    Public Function MakeSeventhFold() As Boolean
        Try
            Return MakeOddFold("f7")
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return False
    End Function

    Public Function MakeOddFold(s As String) As Boolean
        Try
            If GetWorkFace().SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                lamp.LookAtFace(workFace)
                If mainSketch.DrawTrobinaCurve(nombrador.GetQNumber(doku), nombrador.GetNextSketchName(doku), direction).Construction Then
                    sk3D = mainSketch.sk3D
                    curve = mainSketch.curve
                    lamp.ZoomSelected(curve)
                    If GetMajorEdge(workFace).GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                        If DrawSingleLines() Then
                            If AdjustLastAngle() Then
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
                                        'lamp.LookAtFace(workFace)
                                        doku.Update2(True)
                                        If monitor.IsFeatureHealthy(folded) Then
                                            comando.UnfoldBand(doku)
                                            If compDef.HasFlatPattern Then
                                                comando.RefoldBand(doku)
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

            maxface1 = sheetMetalFeatures.FoldFeatures.Item(sheetMetalFeatures.FoldFeatures.Count).Faces.Item(1)
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
            For Each f As Face In sheetMetalFeatures.FoldFeatures.Item(sheetMetalFeatures.FoldFeatures.Count - 1).Faces
                For Each ft As Face In maxface2.TangentiallyConnectedFaces
                    If ft.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        If ft.Equals(f) Then

                            adjacentFace = f
                            lamp.HighLighFace(f)

                        End If

                    End If
                Next

            Next
            adjacentPlane = adjacentFace.Geometry
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
    Function GetMinorAdjacentEdge() As Edge
        Dim e1, e2, e3 As Edge
        Dim min1, min2, min3 As Double
        min1 = 99999
        min2 = 99999
        min3 = 99999
        e1 = adjacentFace.Edges.Item(1)
        e2 = e1
        e3 = e2
        For Each ed As Edge In adjacentFace.Edges
            If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) < min2 Then

                If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) < min1 Then
                    min3 = min2
                    e3 = e2
                    min2 = min1
                    e2 = e1
                    min1 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                    e1 = ed
                Else
                    min3 = min2
                    e3 = e2
                    min2 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                    e2 = ed
                End If
            Else

                min3 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                e3 = ed

            End If

        Next

        Return e1
    End Function
    Function GetStartPoint() As Point
        Try

            Dim pt As Point
            Dim v As Vector

            leadingEdge = GetLeadingEdge()
            pt = point1


            If leadingEdge.StartVertex.Point.DistanceTo(pt) < leadingEdge.StopVertex.Point.DistanceTo(pt) Then
                pt = leadingEdge.StartVertex.Point
                v = leadingEdge.StartVertex.Point.VectorTo(leadingEdge.StopVertex.Point)
                farPoint = leadingEdge.StopVertex.Point
            Else
                pt = leadingEdge.StopVertex.Point
                v = leadingEdge.StopVertex.Point.VectorTo(leadingEdge.StartVertex.Point)
                farPoint = leadingEdge.StartVertex.Point

            End If
            'lamp.HighLighObject(pt)
            lamp.HighLighObject(leadingEdge)
            lamp.HighLighObject(followEdge)
            v.AsUnitVector.AsVector()
            v.ScaleBy(thicknessCM * 3 / 10)
            pt.TranslateBy(v)
            'lamp.HighLighObject(pt)

            point1 = pt
            Return pt
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetLeadingEdge() As Edge
        Try
            Dim pt, opt As Point
            Dim minDis As Double = 9999999999
            Dim minDisAdj As Double = 999999999
            Dim minOpt As Double = 999999999
            Dim minEdge As Double = minDis
            Dim l As SketchLine3D
            bendLine3D = sk3D.Include(bendEdge)
            For Each o As Point In initialPlane.IntersectWithCurve(curve.Geometry)
                'l = sk3D.SketchLines3D.AddByTwoPoints(o, bendLine3D.EndSketchPoint.Geometry)
                If o.DistanceTo(bendEdge.GetClosestPointTo(o)) < minDis Then
                    minDis = o.DistanceTo(bendEdge.GetClosestPointTo(o))
                    pt = bendEdge.GetClosestPointTo(o)
                    opt = GetClosestPointTrobina(pt, o)
                    minOpt = 99999999
                    minEdge = minOpt
                    If opt.DistanceTo(pt) < minDisAdj Then
                        minDisAdj = opt.DistanceTo(pt)
                        For Each ed As Edge In workFace.Edges
                            If ed.Equals(bendEdge) Then
                            Else
                                If ed.GetClosestPointTo(o).DistanceTo(opt) < minOpt Then
                                    minOpt = ed.GetClosestPointTo(o).DistanceTo(opt)
                                    If ed.GetClosestPointTo(o).DistanceTo(bendEdge.GetClosestPointTo(o)) < bendLine3D.Length + gap1CM * 2 Then
                                        If ed.GetClosestPointTo(opt).DistanceTo(opt) < minEdge Then
                                            minEdge = ed.GetClosestPointTo(opt).DistanceTo(opt)
                                            followEdge = leadingEdge
                                            leadingEdge = ed
                                        Else
                                            followEdge = ed
                                        End If
                                    Else
                                        followEdge = ed
                                    End If
                                Else
                                    followEdge = ed
                                End If
                            End If
                        Next

                    End If


                End If
            Next
            point1 = pt
            bendLine3D.Construction = True
            Return leadingEdge
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetClosestPointTrobina(b As Point, o As Point) As Point
        Try
            Dim l As SketchLine3D
            Dim dc, dce As DimensionConstraint3D
            Dim cp As Point
            l = sk3D.SketchLines3D.AddByTwoPoints(b, o, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, bendLine3D)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            dce = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, GetMinorAdjacentEdge().StartVertex)
            adjuster.GetMinimalDimension(dc)
            cp = l.EndSketchPoint.Geometry
            Try
                dc.Delete()
                dce.Delete()
                l.Delete()
            Catch ex As Exception
            End Try
            Return cp
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function

    Function DrawFirstLine() As SketchLine3D
        Try

            Dim l As SketchLine3D
            Dim pt As Point = GetStartPoint()
            minorLine = sk3D.Include(followEdge)
            majorLine = sk3D.Include(leadingEdge)
            l = sk3D.SketchLines3D.AddByTwoPoints(pt, followEdge.GetClosestPointTo(pt))

            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, majorLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, minorLine)
            Dim dc As DimensionConstraint3D
            If farPoint.DistanceTo(majorLine.EndSketchPoint.Geometry) < gap1CM Then
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, majorLine.EndPoint)
            Else
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, majorLine.StartPoint)
            End If


            dc.Parameter._Value = majorLine.Length - thicknessCM * 5
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
    End Function
    Function DrawFirstConstructionLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing
            Dim v As Vector = firstLine.Geometry.Direction.AsVector
            Dim p As Plane
            Dim optpoint As Point = Nothing
            p = tg.CreatePlane(firstLine.EndSketchPoint.Geometry, v)
            Dim d As Double
            Dim minDis As Double = 999999999999
            Dim vc, vmjl As Vector
            vmjl = firstLine.EndSketchPoint.Geometry.VectorTo(farPoint)
            ' l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, farPoint, False)dmm
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
                If o.DistanceTo(firstLine.EndSketchPoint.Geometry) < curvas3D.Cr / 4 Then

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


            Else
                dc.Delete()
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, l.EndPoint)
                adjuster.GetMinimalDimension(dc)

            End If
            dc.Driven = True
            v = firstLine.StartSketchPoint.Geometry.VectorTo(farPoint).AsUnitVector.AsVector
            v.ScaleBy(thicknessCM * 5)
            l.StartSketchPoint.MoveBy(v)
            dc.Driven = False
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, thicknessCM / 3) Then
            Else
                dc.Delete()
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, l.EndPoint)
                adjuster.GetMinimalDimension(dc)

            End If
            dc.Delete()
            gapVertex = sk3D.DimensionConstraints3D.AddLineLength(l)
            Try
                adjuster.GetMinimalDimension(gapVertex)
                gapVertex.Delete()
                gapVertex = sk3D.DimensionConstraints3D.AddLineLength(l)
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
    Function DrawSecondConstructionLine() As SketchLine3D
        Dim gc As GeometricConstraint3D
        Dim cl1 As SketchLine3D = constructionLines(1)
        Try
            Dim l As SketchLine3D
            Dim endPoint As Point
            Dim v As Vector
            Dim pl As Plane
            pl = tg.CreatePlaneByThreePoints(minorLine.EndSketchPoint.Geometry, minorLine.StartSketchPoint.Geometry, majorLine.StartSketchPoint.Geometry)
            v = initialPlane.Normal.AsVector
            v.ScaleBy(-1)
            endPoint = firstLine.EndSketchPoint.Geometry
            endPoint.TranslateBy(v)

            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, endPoint, False)
            Try
                gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, firstLine)
            Catch ex As Exception
                l.Delete()
                gapVertex.Delete()
                gapVertex = sk3D.DimensionConstraints3D.AddLineLength(cl1)
                l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, endPoint, False)
                gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, firstLine)
            End Try


            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            Try
                dc.Parameter._Value = GetParameter("b")._Value
            Catch ex As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1)
                dc.Parameter._Value = GetParameter("b")._Value
            End Try

            GC.Delete()
            l.Construction = True
            constructionLines.Add(l)




            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawSecondLine() As SketchLine3D
        Try

            Dim l, cl2 As SketchLine3D
            Dim dc As DimensionConstraint3D
            cl2 = constructionLines(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(cl2.EndSketchPoint.Geometry, firstLine.StartPoint, False)

            gapVertex.Driven = True
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddCoincident(cl2.EndPoint, l)
            Try
                gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, cl2)
            Catch ex As Exception
                gapVertex.Driven = True
                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl2)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI / 2)
                dc.Driven = True
                gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, cl2)
                gapVertex.Driven = False
            End Try

            gapVertex.Driven = False
            Try
                adjuster.GetMinimalDimension(gapVertex)
            Catch ex As Exception
                gapVertex.Driven = True
            End Try
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

            Dim l, ol, adjl, l3, cl2 As SketchLine3D
            Dim ac, dcl, dcol As DimensionConstraint3D
            Dim d, e, f, b As Double
            Dim limit As Integer = 0
            Dim vbl2, vnap, vcc, vfl, ve, vf As Vector
            b = GetParameter("b")._Value
            cl2 = constructionLines.Item(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, minorEdge.GetClosestPointTo(firstLine.StartSketchPoint.Geometry), False)

            gapVertex.Driven = True
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, minorLine)
            dcl = sk3D.DimensionConstraints3D.AddLineLength(l)
            d = Math.Abs(initialPlane.Normal.AsVector.DotProduct(adjacentPlane.Normal.AsVector))
            If d > 0.01 Then
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 31 / 32)
                If d > 0.1 Then
                    adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 15 / 16)
                    If d > 0.6 Then
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 7 / 8)
                        If d > 0.8 Then
                            adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 3 / 4)
                            If d > 0.9 Then
                                adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 2 / 3)
                            End If
                        End If
                    End If
                End If
            End If
            dcl.Delete()
            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, secondLine)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()
                sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            Catch ex As Exception
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, secondLine)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()
            End Try

            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, minorLine)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()
                sk3D.GeometricConstraints3D.AddPerpendicular(l, minorLine)
            Catch ex As Exception
                dcl = sk3D.DimensionConstraints3D.AddLineLength(l)
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, gap1CM)
                dcl.Delete()
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, minorLine)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Driven = True
                Try
                    sk3D.GeometricConstraints3D.AddPerpendicular(l, minorLine)
                Catch ex3 As Exception
                    ac.Driven = False
                End Try
                adjuster.GetMinimalDimension(gapVertex)
                gapVertex.Driven = True
            End Try

            Dim dc1, dc2 As DimensionConstraint3D
            adjl = sk3D.Include(GetAdjacentEdge())
            f = adjacentEdge.GetClosestPointTo(firstLine.EndSketchPoint.Geometry).DistanceTo(firstLine.EndSketchPoint.Geometry)
            If f < b Then
                If d < 0.76 Then

                    vbl2 = secondLine.Geometry.Direction.AsVector
                    vbl2.ScaleBy(-1)
                    vnap = adjacentPlane.Normal.AsVector
                    vnap.ScaleBy(-1)
                    vfl = firstLine.Geometry.Direction.AsVector
                    vf = vfl.CrossProduct(vbl2)
                    ve = vnap.CrossProduct(vf)
                    e = ve.DotProduct(vfl)
                    If (e < -1 / 16 And firstLine.Length < GetParameter("b")._Value * 2) Then
                        dcl = sk3D.DimensionConstraints3D.AddLineLength(l)
                        While e < (1 / (limit * limit + 4)) And limit < 4
                            Try
                                adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 17 / 16)
                                vbl2 = secondLine.Geometry.Direction.AsVector
                                vbl2.ScaleBy(-1)
                                vnap = adjacentPlane.Normal.AsVector
                                vnap.ScaleBy(-1)
                                vfl = firstLine.Geometry.Direction.AsVector
                                vf = vfl.CrossProduct(vbl2)
                                ve = vnap.CrossProduct(vf)
                                e = ve.DotProduct(vfl)
                                limit = limit + 1
                            Catch ex As Exception
                                limit = limit + 2
                            End Try
                        End While
                        limit = 0
                        dcl.Delete()

                    End If

                End If

                ol = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, adjl.Geometry.MidPoint, False)
                sk3D.GeometricConstraints3D.AddCoincident(ol.StartPoint, secondLine)
                sk3D.GeometricConstraints3D.AddCoincident(ol.EndPoint, adjl)

                Try
                    gapVertex.Driven = False
                    adjuster.GetMinimalDimension(gapVertex)
                    gapVertex.Driven = True
                Catch ex As Exception
                    gapVertex.Driven = True
                End Try

                dcl = sk3D.DimensionConstraints3D.AddLineLength(l)
                dcol = sk3D.DimensionConstraints3D.AddTwoLineAngle(ol, secondLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dcol, Math.PI / 2)
                dcol.Delete()
                sk3D.GeometricConstraints3D.AddPerpendicular(ol, secondLine)

                Try
                    dcol = sk3D.DimensionConstraints3D.AddTwoLineAngle(ol, adjl)
                    adjuster.AdjustDimensionConstraint3DSmothly(dcol, Math.PI / 2)
                    dcol.Delete()
                    sk3D.GeometricConstraints3D.AddPerpendicular(ol, adjl)
                    dcl.Delete()
                Catch ex As Exception
                    Try
                        dcl.Delete()
                    Catch ex4 As Exception
                    End Try
                    dcl = sk3D.DimensionConstraints3D.AddLineLength(ol)
                    adjuster.AdjustDimensionConstraint3DSmothly(dcl, gap1CM)
                    dcl.Delete()
                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(ol, adjl)
                    adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                    ac.Driven = True
                    Try
                        sk3D.GeometricConstraints3D.AddPerpendicular(ol, adjl)
                    Catch ex3 As Exception
                        ac.Driven = False
                    End Try
                    adjuster.GetMinimalDimension(gapVertex)
                    gapVertex.Driven = True
                End Try
                Try
                    gapVertex.Driven = False
                    adjuster.GetMinimalDimension(gapVertex)
                    gapVertex.Driven = True
                Catch ex As Exception
                    gapVertex.Driven = True
                End Try
                dc1 = sk3D.DimensionConstraints3D.AddLineLength(l)
                dc2 = sk3D.DimensionConstraints3D.AddLineLength(ol)
                If dc2.Parameter._Value < dc1.Parameter._Value Then
                    l.Delete()
                    Try

                        adjuster.AdjustDimensionConstraint3DSmothly(dc2, gap1CM * 4)
                        gapFold = dc2
                        ol.Construction = True
                        constructionLines.Add(ol)
                        lastLine = ol
                        l = ol

                    Catch ex As Exception
                        adjuster.AdjustDimensionConstraint3DSmothly(dc2, gap1CM * 4)
                        gapFold = dc2
                        ol.Construction = True
                        constructionLines.Add(ol)
                        lastLine = ol
                        l = ol

                    End Try
                Else
                    ol.Delete()
                    Try

                        adjuster.AdjustDimensionConstraint3DSmothly(dc1, gap1CM * 4)
                        gapFold = dc1
                        l.Construction = True
                        constructionLines.Add(l)
                        lastLine = l

                    Catch ex As Exception
                        adjuster.AdjustDimensionConstraint3DSmothly(dc1, gap1CM * 4)
                        gapFold = dc1
                        l.Construction = True
                        constructionLines.Add(l)
                        lastLine = l

                    End Try

                End If
            Else
                Try
                    dc1 = sk3D.DimensionConstraints3D.AddLineLength(l)
                    adjuster.AdjustDimensionConstraint3DSmothly(dc1, gap1CM * 4)
                    gapFold = dc1
                    l.Construction = True
                    constructionLines.Add(l)
                    lastLine = l

                Catch ex As Exception


                End Try
            End If
            Try
                gapVertex.Driven = False
                adjuster.GetMinimalDimension(gapVertex)
            Catch ex As Exception
                gapVertex.Driven = True
            End Try


            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetAdjacentEdge() As Edge
        Try
            Dim m1, m2, d, e, f As Double
            Dim va, vb, vc, v1, v2 As Vector
            v1 = bendEdge.StartVertex.Point.VectorTo(bendEdge.StopVertex.Point)
            m1 = 0
            vb = point1.VectorTo(farPoint)
            For Each eda As Edge In adjacentFace.Edges
                va = eda.StartVertex.Point.VectorTo(eda.StopVertex.Point)
                v2 = v1.CrossProduct(va)
                vc = eda.GetClosestPointTo(firstLine.EndSketchPoint.Geometry).VectorTo(firstLine.EndSketchPoint.Geometry)
                e = eda.GetClosestPointTo(point1).DistanceTo(point1)
                f = eda.GetClosestPointTo(firstLine.EndSketchPoint.Geometry).DistanceTo(firstLine.EndSketchPoint.Geometry)
                d = 1 / (1 + f) * v2.Length * e * e
                If d > m1 Then
                    m1 = d
                    adjacentEdge = eda
                End If
            Next
            lamp.HighLighObject(adjacentEdge)
            Return adjacentEdge
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawThirdLine() As SketchLine3D
        Try
            Dim l, pl As SketchLine3D
            Dim ac As DimensionConstraint3D
            pl = bandLines(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, pl.StartPoint, False)

            Dim gcppl3 As GeometricConstraint3D
            Try
                gcppl3 = sk3D.GeometricConstraints3D.AddPerpendicular(l, firstLine)
            Catch ex As Exception
                gapVertex.Driven = True
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, firstLine)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Driven = True
                Try
                    gcppl3 = sk3D.GeometricConstraints3D.AddPerpendicular(l, firstLine)
                Catch ex3 As Exception
                    ac.Driven = False
                End Try
                gapVertex.Driven = False
                adjuster.GetMinimalDimension(gapVertex)
            End Try

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
            Dim l, bl3 As SketchLine3D
            Dim v As Vector
            Dim p As Plane
            Dim dcl As DimensionConstraint3D
            Dim minDis As Double
            Dim optpoint As Point
            bl3 = bandLines.Item(3)
            v = bl3.Geometry.Direction.AsVector
            p = tg.CreatePlane(bl3.EndSketchPoint.Geometry, v)

            Dim puntos As ObjectsEnumerator
            puntos = p.IntersectWithCurve(curve.Geometry)
            minDis = 9999999999
            Dim o2 As Point = puntos.Item(puntos.Count)
            optpoint = o2
            For Each o As Point In puntos
                If o.DistanceTo(bl3.EndSketchPoint.Geometry) < minDis Then
                    minDis = o.DistanceTo(bl3.EndSketchPoint.Geometry)
                    o2 = optpoint
                    optpoint = o
                End If

            Next


            l = sk3D.SketchLines3D.AddByTwoPoints(bl3.EndSketchPoint.Geometry, optpoint, False)
            Try
                sk3D.GeometricConstraints3D.AddCoincident(bl3.EndPoint, l)
            Catch ex As Exception
                gapVertex.Driven = True
                sk3D.Solve()
                sk3D.GeometricConstraints3D.AddCoincident(bl3.EndPoint, l)
            End Try

            Try
                sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            Catch ex As Exception
                gapVertex.Driven = True
                dcl = sk3D.DimensionConstraints3D.AddLineLength(l)
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 7 / 8)
                dcl.Delete()
                sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
                gapVertex.Driven = False
                adjuster.GetMinimalDimension(gapVertex)
            End Try

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
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 5 / 4) Then
                dc.Delete()
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
                    dc.Delete()
                    Try
                        sk3D.GeometricConstraints3D.AddEqual(l, cl)
                    Catch ex As Exception
                        gapVertex.Driven = True
                        dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                        adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value)
                        dc.Delete()
                        sk3D.GeometricConstraints3D.AddEqual(l, cl)
                        gapVertex.Driven = False
                    End Try


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
        Dim currentQ As Integer = GetParameter("currentQ")._Value
        Dim oddQ As Integer
        Dim tensionFactor As Double = 2
        Try

            If sheetMetalFeatures.FoldFeatures.Count > 3 Then
                Math.DivRem(currentQ, 2, oddQ)
                If oddQ = 0 Then
                    tensionFactor = 1
                Else
                    tensionFactor = 3
                End If
            End If

            Dim fourLine, sixthLine, cl3, cl2 As SketchLine3D
            Dim ac As TwoLineAngleDimConstraint3D
            Dim limit As Double = 0.16
            Dim counterLimit As Integer = 0
            Dim angleLimit As Double
            If GetParameter("q")._Value = 31 Then
                angleLimit = 2.67
            Else
                angleLimit = 2.2
            End If

            Dim lastAngle, lastValue As Double

            Dim d As Double

            Dim b As Boolean = False
            fourLine = bandLines.Item(2)
            sixthLine = bandLines.Item(4)
            cl3 = constructionLines.Item(3)
            cl2 = constructionLines.Item(2)
            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)
            Catch ex As Exception
                adjuster.AdjustDimConstrain3DSmothly(gapFold, gapFold.Parameter._Value * 17 / 16)
                gapFold.Driven = True
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)
                gapFold.Driven = False
            End Try

            ac.Driven = True
            gapVertex.Driven = False
            d = CalculateRoof()
            If d > 0 Then
                If adjuster.AdjustGapSmothly(gapFold, gap1CM * tensionFactor, ac) Then
                    Try
                        While (ac.Parameter._Value < angleLimit And ac.Parameter._Value > angleLimit / 2) And counterLimit < 4
                            lastAngle = ac.Parameter._Value
                            adjuster.AdjustGapSmothly(gapFold, gap1CM * tensionFactor, ac)
                            If ac.Parameter._Value < lastAngle Then
                                counterLimit = 4
                            End If
                            counterLimit = counterLimit + 1
                        End While
                        counterLimit = 0
                    Catch ex As Exception
                        ac.Driven = True

                    End Try
                    b = True
                Else
                    gapVertex.Driven = True
                    adjuster.AdjustGapSmothly(gapFold, gap1CM * tensionFactor, ac)
                    ac.Driven = True
                    gapFold.Delete()
                    gapFold = sk3D.DimensionConstraints3D.AddLineLength(cl3)
                    b = True
                End If

            Else
                Try
                    ac.Driven = True
                    While ((d < 0 Or ac.Parameter._Value > Math.PI - limit / 1) And counterLimit < 32)
                        Try
                            gapFold.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value * 17 / 16)
                            gapFold.Driven = True
                            sk3D.Solve()

                        Catch ex2 As Exception
                            counterLimit = counterLimit + 1
                            gapFold.Driven = True
                        End Try
                        d = CalculateRoof()
                        counterLimit = counterLimit + 1
                    End While
                    gapFold.Driven = False
                    b = True
                Catch ex As Exception
                    Try
                        gapFold.Delete()
                    Catch ex3 As Exception
                    End Try
                    gapFold = sk3D.DimensionConstraints3D.AddLineLength(cl3)
                    b = True
                End Try
            End If

            Try
                gapFold.Driven = True
                ac.Driven = False
                Try
                    gapVertex.Driven = False
                    counterLimit = 0
                    lastValue = 0
                    While (Math.Abs(gapVertex.Parameter._Value - lastValue) > gap1CM / 64) And counterLimit < 64
                        lastValue = gapVertex.Parameter._Value
                        adjuster.GetMinimalDimension(gapVertex)
                        counterLimit = counterLimit + 1
                    End While

                Catch ex As Exception
                    gapVertex.Driven = True
                End Try
                ac.Driven = True
                gapFold.Driven = False

            Catch ex As Exception
                gapFold.Driven = True
                Try
                    'gapVertex.Driven = False
                    counterLimit = 0
                    lastValue = 0
                    While (Math.Abs(gapVertex.Parameter._Value - lastValue) > gap1CM / 64) And counterLimit < 64
                        lastValue = gapVertex.Parameter._Value
                        adjuster.GetMinimalDimension(gapVertex)
                        counterLimit = counterLimit + 1
                    End While
                Catch ex4 As Exception
                    gapVertex.Driven = True
                End Try
            End Try
            RegenerateGapVertex(True)
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function RegenerateGapVertex(driven As Boolean) As DimensionConstraint3D
        Dim cl1 As SketchLine3D = constructionLines(1)
        gapVertex.Delete()
        sk3D.Solve()
        gapVertex = sk3D.DimensionConstraints3D.AddLineLength(cl1)
        gapVertex.Driven = driven
        Return gapVertex
    End Function
    Function CalculateRoof() As Double
        Dim fourLine, sixthLine, cl3, cl2 As SketchLine3D


        Dim v1, v2, v3, v4, vfl, vnwf, vflnwf, vcp As Vector
        Dim d, e, f, g As Double
        Dim i As Integer


        fourLine = bandLines.Item(2)
        sixthLine = bandLines.Item(4)
        cl3 = constructionLines.Item(3)
        cl2 = constructionLines.Item(2)
        v3 = firstLine.StartSketchPoint.Geometry.VectorTo(sixthLine.EndSketchPoint.Geometry)
        v2 = fourLine.EndSketchPoint.Geometry.VectorTo(fourLine.StartSketchPoint.Geometry)
        v4 = cl2.Geometry.Direction.AsVector
        v1 = v2.CrossProduct(v3)
        vfl = firstLine.Geometry.Direction.AsVector
        vnwf = initialPlane.Normal.AsVector
        vnwf.ScaleBy(-1)
        vflnwf = vfl.CrossProduct(vnwf)
        vcp = compDef.WorkPoints.Item(1).Point.VectorTo(firstLine.EndSketchPoint.Geometry)
        e = vcp.DotProduct(vflnwf)
        d = v1.DotProduct(v4)
        g = Math.Abs(initialPlane.Normal.AsVector.DotProduct(adjacentPlane.Normal.AsVector))
        If e < 0 And g < 0.11 Then
            f = firstLine.Geometry.Direction.AsVector.CrossProduct(v3).DotProduct(vnwf) * -1
        Else
            f = 1
        End If

        Return d * e * f
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
                ff = bender.CorrectFold(ff)
            End If
            Return ff

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
End Class
