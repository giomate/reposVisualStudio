Imports Inventor
Imports SecondFold
Public Class MacroFold
    Public doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim curvas3D As Curves3D
    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim initialPlane As Plane
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
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, leadingEdge, followEdge As Edge
    Dim leadingLine, followLine As SketchLine3D
    Dim workFace, adjacentFace, nextworkface As Face
    Dim bendAngle As DimensionConstraint
    Dim folded As FoldFeature
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim lamp As Highlithing
    Dim bender As Doblador
    Dim gapFold, gapVertex As DimensionConstraint3D
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
            If GetWorkFace().SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                If mainSketch.DrawTrobinaCurve(nombrador.GetQNumber(doku), nombrador.GetNextSketchName(doku)).Construction Then
                    sk3D = mainSketch.sk3D
                    curve = mainSketch.curve
                    If GetMajorEdge(workFace).GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                        If DrawSingleLines() Then
                            If AdjustlastAngle() Then
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
                                        folded.Name = "f3"
                                        doku.Update2(True)
                                        If monitor.IsFeatureHealthy(folded) Then
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
                            'lamp.HighLighFace(f)

                        End If

                    End If
                Next

            Next


            lamp.HighLighFace(workFace)
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
    Function GetStartPoint() As Point
        Dim vbl, v As Vector
        Dim bl As SketchLine3D

        bl = sk3D.Include(bendEdge)

        vbl = bl.Geometry.Direction.AsVector
        Dim ps, pe As Plane
        Dim pt As Point = Nothing
        ps = tg.CreatePlane(bl.StartSketchPoint.Geometry, vbl)
        Dim minDis As Double = 9999999999
        Dim minEdge As Double = minDis
        pe = tg.CreatePlane(bl.EndSketchPoint.Geometry, vbl)

        For Each o As Point In pe.IntersectWithCurve(curve.Geometry)
            If o.DistanceTo(bl.EndSketchPoint.Geometry) < minDis Then
                minDis = o.DistanceTo(bl.EndSketchPoint.Geometry)
                pt = o
                For Each ed As Edge In workFace.Edges
                    If ed.Equals(bendEdge) Then
                    Else
                        If ed.GetClosestPointTo(o).DistanceTo(bl.EndSketchPoint.Geometry) < bl.Length + thicknessCM Then
                            If ed.GetClosestPointTo(o).DistanceTo(o) < minEdge Then
                                minEdge = ed.GetClosestPointTo(o).DistanceTo(o)
                                followEdge = leadingEdge
                                leadingEdge = ed
                            Else
                                followEdge = ed
                            End If

                        End If
                    End If


                Next
            End If
        Next
        'lamp.HighLighObject(leadingEdge)
        'lamp.HighLighObject(followEdge)
        minEdge = 9999999999
        For Each o As Point In ps.IntersectWithCurve(curve.Geometry)
            If o.DistanceTo(bl.StartSketchPoint.Geometry) < minDis Then
                minDis = o.DistanceTo(bl.StartSketchPoint.Geometry)
                pt = o
                For Each ed As Edge In workFace.Edges
                    If ed.Equals(bendEdge) Then
                    Else
                        If ed.GetClosestPointTo(o).DistanceTo(bl.StartSketchPoint.Geometry) < bl.Length + thicknessCM Then
                            If ed.GetClosestPointTo(o).DistanceTo(o) < minEdge Then
                                minEdge = ed.GetClosestPointTo(o).DistanceTo(o)
                                followEdge = leadingEdge
                                leadingEdge = ed
                            Else
                                followEdge = ed
                            End If

                        End If
                    End If
                Next
            End If
        Next
        'lamp.HighLighObject(leadingEdge)
        'lamp.HighLighObject(followEdge)


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
        v.AsUnitVector.AsVector()
        v.ScaleBy(thicknessCM * 5 / 10)
        pt.TranslateBy(v)
        'lamp.HighLighObject(pt)
        bl.Construction = True
        point1 = pt
        Return pt
    End Function
    Function DrawFirstLine() As SketchLine3D
        Try

            Dim l As SketchLine3D
            Dim pt As Point = GetStartPoint()
            leadingLine = sk3D.Include(leadingEdge)
            followLine = sk3D.Include(followEdge)
            l = sk3D.SketchLines3D.AddByTwoPoints(pt, followEdge.GetClosestPointTo(pt))

            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, leadingLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, followLine)
            Dim dc As DimensionConstraint3D
            If farPoint.DistanceTo(leadingLine.EndSketchPoint.Geometry) < gap1CM Then
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, leadingLine.EndPoint)
            Else
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, leadingLine.StartPoint)
            End If

            dc.Parameter._Value = leadingLine.Length - thicknessCM * 5 / 1
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
            gapVertex = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, l.EndPoint)
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
        Try
            Dim l, ol As SketchLine3D
            Dim endPoint As Point
            Dim v1, v2, v3 As Vector
            Dim pl As Plane
            v1 = firstLine.StartSketchPoint.Geometry.VectorTo(farPoint)
            v2 = firstLine.Geometry.Direction.AsVector
            v3 = v1.CrossProduct(v2)
            v3.ScaleBy(1 / GetParameter("b")._Value)
            endPoint = firstLine.EndSketchPoint.Geometry
            endPoint.TranslateBy(v3)
            ol = constructionLines(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, endPoint, False)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, firstLine)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
            Else
                dc.Driven = True

            End If
            dc.Delete()
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            dc.Parameter._Value = GetParameter("b")._Value
            gc.Delete()
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

            Dim l, ol As SketchLine3D
            ol = constructionLines.Item(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(ol.EndSketchPoint.Geometry, firstLine.StartPoint, False)
            gapVertex.Driven = True
            Dim gc As GeometricConstraint3D

            gc = sk3D.GeometricConstraints3D.AddCoincident(ol.EndPoint, l)
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, ol)
            gapVertex.Driven = False
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

            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, followEdge.GetClosestPointTo(firstLine.StartSketchPoint.Geometry), False)


            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, followLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, followLine)
            Dim dc1, dc2 As DimensionConstraint3D
            pvl = sk3D.Include(GetAdjacentEdge())
            dc1 = sk3D.DimensionConstraints3D.AddLineLength(l)
            ol = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, pvl.Geometry.MidPoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(ol.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(ol.EndPoint, pvl)
            sk3D.GeometricConstraints3D.AddPerpendicular(ol, secondLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(ol, pvl)

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

                If adjuster.AdjustDimensionConstraint3DSmothly(dc1, gap1CM * 3) Then
                    gapFold = dc1
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
    Function AdjustlastAngle() As Boolean
        Try
            Dim fourLine, sixthLine, cl3, cl2 As SketchLine3D
            Dim dc As TwoLineAngleDimConstraint3D
            Dim limit As Double = 0.1
            Dim counterLimit As Integer = 0

            Dim d As Double

            Dim b As Boolean = False
            fourLine = bandLines.Item(2)
            sixthLine = bandLines.Item(4)
            cl3 = constructionLines.Item(3)
            cl2 = constructionLines.Item(2)

            d = CalculateRoof()
            If d > 0 Then
                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)

                If adjuster.AdjustGapSmothly(gapFold, gap1CM, dc) Then
                    b = True
                Else
                    dc.Driven = True
                    gapFold.Delete()
                    gapFold = sk3D.DimensionConstraints3D.AddLineLength(cl3)
                    b = True
                End If
            Else
                Try
                    dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)
                    dc.Driven = True
                    While (d < 0 And counterLimit < 32)
                        Try
                            gapFold.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value * 17 / 16)
                            gapFold.Driven = True
                            doku.Update2()

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



            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function CalculateRoof() As Double
        Dim fourLine, sixthLine, cl3, cl2 As SketchLine3D


        Dim v1, v2, v3, v4 As Vector
        Dim d As Double


        fourLine = bandLines.Item(2)
        sixthLine = bandLines.Item(4)
        cl3 = constructionLines.Item(3)
        cl2 = constructionLines.Item(2)
        v3 = firstLine.StartSketchPoint.Geometry.VectorTo(sixthLine.EndSketchPoint.Geometry)
        v2 = fourLine.EndSketchPoint.Geometry.VectorTo(fourLine.StartSketchPoint.Geometry)
        v4 = cl2.Geometry.Direction.AsVector
        v1 = v2.CrossProduct(v3)
        d = v1.DotProduct(v4)

        Return d
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
