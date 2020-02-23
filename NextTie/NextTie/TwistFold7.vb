Imports Inventor
Imports ThirdFold
Imports FourthFold

Public Class TwistFold7
    Dim doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, connectLine, nextLine, centroLine, bendLine3D, tangentLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, startPoint As Point
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
    Dim lastCut As Cortador
    Dim cutProfile As Profile

    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, cutEdge1, cutEdge2, CutEsge3, leadingEdge, followEdge As Edge
    Dim minorLine, majorLine, cutLine3D, kante3D, tante3D As SketchLine3D
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, twistFace, nextworkFace As Face
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex, gapTwist, gapTry, perpendicular1, cruzeta As DimensionConstraint3D
    Dim equalCl4 As GeometricConstraint3D
    Dim folded As FoldFeature
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim lamp As Highlithing
    Dim bender As Doblador
    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails As ObjectCollection
    Dim manager As FoldingEvaluator

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
        sheetMetalFeatures = compDef.Features
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
        manager = New FoldingEvaluator(doku)


        done = False
    End Sub
    Public Function MakeFinalTwist() As Boolean
        Try
            If (manager.GetLastFold(doku).Name = "f7" Or manager.GetLastFold(doku).Name = "f9") Then
                foldFeature = manager.GetLastFold(doku)
            Else
                foldFeature = MakeLastFold()
            End If
            If monitor.IsFeatureHealthy(foldFeature) Then
                firstLine = compDef.Sketches3D.Item("firstLine").SketchLines3D.Item(1)
                secondLine = compDef.Sketches3D.Item("secondLine").SketchLines3D.Item(1)
                nextLine = compDef.Sketches3D.Item("nextLine").SketchLines3D.Item(1)
                connectLine = compDef.Sketches3D.Item("connectLine").SketchLines3D.Item(1)
                tangentLine = compDef.Sketches3D.Item("tangentLine").SketchLines3D.Item(1)
                If sheetMetalFeatures.CutFeatures.Count > 0 Then
                    If sheetMetalFeatures.CutFeatures.Item(sheetMetalFeatures.CutFeatures.Count).Name = "lastTwistCut" Then
                        cutfeature = sheetMetalFeatures.CutFeatures.Item(sheetMetalFeatures.CutFeatures.Count)
                    Else
                        cutfeature = MakeLastCut()
                    End If
                Else

                    cutfeature = MakeLastCut()
                End If

                If monitor.IsFeatureHealthy(cutfeature) Then

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
        Return done
    End Function
    Function MakeLastCut() As CutFeature
        If monitor.IsFeatureHealthy(foldFeature) Then
            lastCut = New Cortador(doku)
            cutFace = lastCut.GetLastCutFace(foldFeature)

            cutfeature = lastCut.MakeLastCut(firstLine, secondLine)
            cutfeature.Name = "lastTwistCut"
            cutLine = lastCut.lastCutLine
            cutLine3D = lastCut.cutLine3D
        End If
        Return cutfeature
    End Function
    Function MakeReferenceSketch() As SketchLine3D
        Try
            Dim l As SketchLine3D
            sk3D = compDef.Sketches3D.Add()
            l = sk3D.Include(nextLine)
            sk3D.Name = "last"
            sk3D = compDef.Sketches3D.Add()
            l = sk3D.Include(tante3D)
            sk3D.Name = "kante"
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
            Dim i, j, k, l, f, g, s As Integer
            Dim pr3d As Profile3D

            sk3D = compDef.Sketches3D.Add()
            sk3D.Name = "rails"
            i = 0
            bendLine = foldFeature.Definition.BendLine
            If cutLine.StartSketchPoint.Geometry.DistanceTo(bendLine.EndSketchPoint.Geometry) < cutLine.EndSketchPoint.Geometry.DistanceTo(bendLine.EndSketchPoint.Geometry) Then
                For Each v As Vertex In twistFace.Vertices
                    i = i + 1
                    If cutLine3D.StartSketchPoint.Geometry.DistanceTo(v.Point) < d Then
                        d = cutLine3D.StartSketchPoint.Geometry.DistanceTo(v.Point)
                        j = i
                    End If
                Next


            Else
                For Each v As Vertex In twistFace.Vertices
                    i = i + 1
                    If cutLine3D.EndSketchPoint.Geometry.DistanceTo(v.Point) < d Then
                        d = cutLine3D.StartSketchPoint.Geometry.DistanceTo(v.Point)
                        j = i
                    End If
                Next



            End If
            s = RotationDirection()


            If s > 0 Then
                g = 1
                For index = j To j + 3
                    Math.DivRem(index, 4, k)
                    Math.DivRem(k + g, 4, l)
                    sp = esquinas.Item(l + 1)
                    sk3D = compDef.Sketches3D.Add()
                    l1 = sk3D.SketchLines3D.AddByTwoPoints(twistFace.Vertices.Item(k + 1).Point, sp.Geometry3d)
                    pr3d = sk3D.Profiles3D.AddOpen
                    rails.Add(pr3d)
                Next
            Else
                g = 1
                For index = j To j + 3
                    Math.DivRem(index, 4, k)
                    Math.DivRem(k + g, 4, l)
                    sp = esquinas.Item(4 - l)
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
    Function RotationDirection() As Integer
        Dim sign As Integer
        Dim pl As Plane
        Dim vcl5, vcl4, v3 As Vector


        'vnpl.ScaleBy(-1)
        vcl4 = nextLine.Geometry.Direction.AsVector
        vcl5 = tangentLine.Geometry.Direction.AsVector
        v3 = vcl4.CrossProduct(vcl5)
        If v3.Z > 0 Then
            sign = -1
        Else
            sign = 1

        End If

        Return sign
    End Function
    Public Function MakeLastFold() As FoldFeature
        Dim cl5 As SketchLine3D
        Try
            If GetWorkFace().SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                If curvas3D.DrawTrobinaCurve(nombrador.GetQNumber(doku), nombrador.GetNextSketchName(doku)).Construction Then
                    sk3D = curvas3D.sk3D
                    curve = curvas3D.curve
                    If GetMajorEdge(workFace).GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                        If DrawSingleLines().Length > 0 Then

                            nextSketch = New OriginSketch(doku)
                            Dim tl As SketchLine3D
                            tl = bandLines.Item(4)
                            cl5 = constructionLines.Item(5)
                            connectLine = nextSketch.DrawNextMainSketch(refLine, tl, firstLine, cl5)
                            If connectLine.Length > 0 Then
                                centroLine = nextSketch.centroLine
                                If ConnectTwistBrigde() Then
                                    Dim sl As SketchLine3D
                                    sl = bandLines(1)
                                    If bender.GetBendLine(workFace, sl).Length > 0 Then
                                        bendLine = bender.bendLine
                                        sl = bandLines(2)
                                        If bender.GetFoldingAngle(leadingEdge, sl).Parameter._Value > 0 Then
                                            comando.MakeInvisibleSketches(doku)
                                            comando.MakeInvisibleWorkPlanes(doku)
                                            folded = bender.FoldBand(bandLines.Count)
                                            folded = CheckFoldSide(folded)
                                            Try
                                                folded.Name = "f7"
                                            Catch ex As Exception
                                                folded.Name = "f9"
                                            End Try
                                            doku.Update2(True)
                                            If monitor.IsFeatureHealthy(folded) Then
                                                cl5 = constructionLines.Item(5)
                                                foldFeature = bender.folded
                                                sk3D = compDef.Sketches3D.Add()
                                                sk3D.Include(firstLine)
                                                sk3D.Name = "firstLine"
                                                sk3D = compDef.Sketches3D.Add()
                                                sk3D.Include(secondLine)
                                                sk3D.Name = "secondLine"
                                                sk3D = compDef.Sketches3D.Add()
                                                nextLine = constructionLines.Item(4)
                                                sk3D.Include(nextLine)
                                                sk3D.Name = "nextline"
                                                sk3D = compDef.Sketches3D.Add()
                                                connectLine = nextSketch.secondLine
                                                sk3D.Include(connectLine)
                                                sk3D.Name = "connectLine"
                                                sk3D = compDef.Sketches3D.Add()
                                                sk3D.Include(cl5)
                                                sk3D.Name = "tangentLine"
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
                ff = bender.CorrectFold(ff)
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
            Dim l, il, sl As SketchLine3D
            Dim pr As Profile
            Dim ps As PlanarSketch
            Dim l2d, il2d As SketchLine
            Dim gc As GeometricConstraint3D
            Dim dc As DimensionConstraint3D
            Dim v1, v2, v3 As Vector
            Dim pt As Point = nextLine.EndSketchPoint.Geometry


            sk3D = compDef.Sketches3D.Add()
            sk3D.Name = "twistProfile"
            il = sk3D.Include(nextLine)
            sl = sk3D.Include(connectLine)
            v1 = il.Geometry.Direction.AsVector
            v2 = sl.Geometry.Direction.AsVector
            v3 = v2.CrossProduct(v1)
            pt.TranslateBy(v3)
            l = sk3D.SketchLines3D.AddByTwoPoints(il.EndPoint, pt, False)
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, il)
            'gc.Delete()
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, sl)

            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, gap1CM / 10) Then
                tante3D = l
                twistPlane = doku.ComponentDefinition.WorkPlanes.AddByThreePoints(il.StartPoint, il.EndPoint, l.EndPoint)

                ps = doku.ComponentDefinition.Sketches.Add(twistPlane)
                l2d = ps.AddByProjectingEntity(l)
                esquinas.Add(l2d.EndSketchPoint)
                il2d = ps.AddByProjectingEntity(il)
                esquinas.Add(il2d.EndSketchPoint)
                esquinas.Add(il2d.StartSketchPoint)
                ps.SketchLines.AddAsThreePointRectangle(il2d.StartSketchPoint, il2d.EndSketchPoint, l2d.EndSketchPoint.Geometry)

                esquinas.Add(ps.SketchPoints.Item(ps.SketchPoints.Count))



                pr = ps.Profiles.AddForSolid

                sections.Add(pr)
                kante3D = l
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
            Dim pro As Profile
            Dim maxArea1, maxArea2, maxArea3 As Double
            Dim maxface1, maxface2, maxface3 As Face
            maxface2 = compDef.Bends.Item(compDef.Bends.Count).FrontFaces.Item(1)
            maxface1 = maxface2
            maxArea2 = 0
            maxArea1 = maxArea2
            For Each f As Face In foldFeature.Faces
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
            pro = ps.Profiles.AddForSolid
            sk3D = compDef.Sketches3D.Add()
            cutLine3D = sk3D.Include(cl)
            sk3D.Name = "lastCutLine"
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
            If monitor.IsFeatureHealthy(cutfeature) Then
                cutfeature.Name = "lastCut"
            End If
            Return oFaceFeature
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing

        End Try


    End Function
    Function ConnectTwistBrigde() As Boolean
        Try

            Dim cnl, nssl As SketchLine3D
            Dim dc1, dccl, accl, dcgvc As DimensionConstraint3D
            Dim gc, gcpp As GeometricConstraint3D
            Dim limit As Integer = 0
            Try
                gapVertex.Driven = False
                doku.Update2()
                adjuster.GetMinimalDimension(gapVertex)


            Catch ex As Exception
                gapVertex.Driven = True
            End Try
            gapVertex.Driven = True

            cnl = sk3D.SketchLines3D.AddByTwoPoints(refLine.EndPoint, connectLine.EndPoint, False)
            ' tl2 = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartPoint, nextSketch.secondLine.StartSketchPoint.Geometry, False)
            'gc = sk3D.GeometricConstraints3D.AddCoincident(cnl1.EndPoint, nextSketch.secondLine)
            dc1 = sk3D.DimensionConstraints3D.AddLineLength(cnl)
            ' dc2 = sk3D.DimensionConstraints3D.AddLineLength(tl2)
            Try
                adjuster.AdjustDimensionConstraint3DSmothly(dc1, dc1.Parameter._Value / 2)
            Catch ex As Exception
                dc1.Driven = True
            End Try

            dc1.Driven = True
            ' cnl1.Delete()
            Try
                gapVertex.Driven = False
                doku.Update2()
                adjuster.GetMinimalDimension(gapVertex)
            Catch ex As Exception
                gapVertex.Driven = True
            End Try
            dccl = sk3D.DimensionConstraints3D.AddLineLength(centroLine)
            dccl.Driven = True
            nssl = nextSketch.secondLine
            accl = sk3D.DimensionConstraints3D.AddTwoLineAngle(cnl, nssl)
            Try
                While ((gapVertex.Parameter._Value > gap1CM Or accl.Parameter._Value < Math.PI) And limit < 2)
                    Try
                        accl.Driven = False
                        gapVertex.Driven = True

                        accl.Driven = False
                        gapVertex.Driven = True
                        adjuster.AdjustDimensionConstraint3DSmothly(accl, Math.PI)
                        accl.Driven = True
                        gapVertex.Driven = False
                        adjuster.GetMinimalDimension(gapVertex)
                        gapVertex.Driven = True


                        limit = limit + 1
                    Catch ex As Exception
                        limit = limit + 1
                        accl.Driven = True
                        gapVertex.Driven = True
                    End Try
                End While
                limit = 0
            Catch ex As Exception
                accl.Driven = True
                gapVertex.Driven = True
            End Try
            Try
                Try
                    dcgvc = sk3D.DimensionConstraints3D.AddTwoPointDistance(firstLine.EndPoint, bendEdge.GetClosestPointTo(firstLine.EndSketchPoint.Geometry))
                    If dcgvc.Parameter._Value < gap1CM / 2 Then
                        gapVertex.Driven = True
                        adjuster.AdjustDimensionConstraint3DSmothly(dcgvc, gap1CM / 2)
                    Else
                        Try
                            dcgvc.Driven = True
                        Catch ex As Exception
                        End Try
                    End If
                Catch ex As Exception
                    Try
                        If dcgvc.Parameter._Value < gap1CM / 2 Then
                            gapVertex.Driven = True
                            adjuster.AdjustDimensionConstraint3DSmothly(dcgvc, gap1CM / 2)
                        Else
                            Try
                                dcgvc.Driven = True
                            Catch ex2 As Exception
                            End Try
                        End If
                    Catch ex4 As Exception
                    End Try
                End Try

            Catch ex As Exception

            End Try
            accl.Driven = True
            cnl.Construction = True

            Return True
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function TurnCrossAngle() As DimensionConstraint3D
        Try
            Dim d As Double
            Dim limit As Integer = 0
            d = CalculateTwistFactor()
            cruzeta.Driven = False
            While (d < (1 / (limit + 1)) And limit < 16)
                Try
                    adjuster.AdjustDimensionConstraint3DSmothly(cruzeta, cruzeta.Parameter._Value * 5 / 4)
                    d = CalculateTwistFactor()
                    limit = limit + 1
                Catch ex As Exception
                    cruzeta.Driven = True
                    limit = limit + 2
                End Try

            End While
            cruzeta.Driven = True
        Catch ex As Exception
            cruzeta.Driven = True
        End Try
        Return cruzeta
    End Function
    Function DrawTwistAligner() As SketchLine3D

        Try
            Dim l, cl2 As SketchLine3D
            cl2 = constructionLines.Item(2)
            Dim pl As Plane
            Dim vnpl As Vector
            pl = tg.CreatePlaneByThreePoints(firstLine.EndSketchPoint.Geometry, firstLine.StartSketchPoint.Geometry, secondLine.StartSketchPoint.Geometry)
            vnpl = pl.Normal
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndSketchPoint.Geometry, secondLine.StartSketchPoint.Geometry, False)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            Dim dc As DimensionConstraint3D
            ' dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            ' If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
            '  dc.Delete()
            sk3D.GeometricConstraints3D.AddEqual(l, cl2)
            l.Construction = True
            constructionLines.Add(l)

            ' End If




            refLine = l
            lastLine = l
            Return l
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
                                If DrawFourthConstructionLine().Construction Then
                                    If DrawThirdLine.Length > 0 Then
                                        'Return bandLines.Item(3)
                                        If DrawFouthLine.Length > 0 Then
                                            'Return bandLines.Item(4)
                                            If DrawFifthConstructionLine().Length > 0 Then
                                                Return bandLines.Item(4)

                                            End If

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
        Try
            Dim pt As Point
            Dim v As Vector

            leadingEdge = GetLeadingEdge()
            pt = point1

            lamp.HighLighObject(leadingEdge)
            lamp.HighLighObject(followEdge)


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
            'lamp.HighLighObject(leadingEdge)
            v.AsUnitVector.AsVector()
            v.ScaleBy(thicknessCM * 5 / 10)
            pt.TranslateBy(v)
            'lamp.HighLighObject(pt)


            point1 = pt
            Return pt
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Function
    Function GetLeadingEdge() As Edge
        Try
            Dim pt, opt, blep, blsp As Point
            Dim minDis As Double = 999999999
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
                                End If
                            Else
                                followEdge = ed
                            End If
                        End If
                    Next
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
    Function GetClosestPointTrobina(o As Point) As Point
        Try
            Dim l As SketchLine3D
            Dim dc As DimensionConstraint3D
            Dim cp As Point
            l = sk3D.SketchLines3D.AddByTwoPoints(bendLine3D.EndPoint, o, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            adjuster.GetMinimalDimension(dc)
            cp = l.EndSketchPoint.Geometry
            Try
                dc.Delete()
                l.Delete()
            Catch ex As Exception
            End Try
            Return cp
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function GetClosestPointTrobina(b As Point, o As Point) As Point
        Try
            Dim l As SketchLine3D
            Dim dc As DimensionConstraint3D
            Dim cp As Point
            l = sk3D.SketchLines3D.AddByTwoPoints(b, o, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, bendLine3D)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            adjuster.GetMinimalDimension(dc)
            cp = l.EndSketchPoint.Geometry
            Try
                dc.Delete()
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
            v = initialPlane.Normal.AsVector
            v.ScaleBy(-1)
            endPoint = firstLine.EndSketchPoint.Geometry
            endPoint.TranslateBy(v)
            ol = constructionLines(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, endPoint, False)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, firstLine)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            Try
                dc.Parameter._Value = GetParameter("b")._Value
            Catch ex As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1)
                dc.Parameter._Value = GetParameter("b")._Value
            End Try



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
    Function DrawFourthConstructionLine() As SketchLine3D
        Try
            Dim l, cl2 As SketchLine3D
            cl2 = constructionLines.Item(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndSketchPoint.Geometry, secondLine.StartSketchPoint.Geometry, False)
            Dim gc As GeometricConstraint3D
            gc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            Dim dc As DimensionConstraint3D
            ' dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            ' If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
            '  dc.Delete()
            equalCl4 = sk3D.GeometricConstraints3D.AddEqual(l, cl2)
            l.Construction = True
            constructionLines.Add(l)

            ' End If




            refLine = l
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

            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, followEdge.GetClosestPointTo(firstLine.StartSketchPoint.Geometry), False)


            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, minorLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, minorLine)
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
                If adjuster.AdjustDimensionConstraint3DSmothly(dc2, gap1CM * 1) Then
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
            Dim m1, m2, d, e As Double
            Dim va, vb, vc, v1, v2 As Vector
            v1 = bendEdge.StartVertex.Point.VectorTo(bendEdge.StopVertex.Point)
            m1 = 0
            vb = point1.VectorTo(farPoint)
            For Each eda As Edge In adjacentFace.Edges
                va = eda.StartVertex.Point.VectorTo(eda.StopVertex.Point)
                v2 = v1.CrossProduct(va)
                vc = eda.GetClosestPointTo(firstLine.EndSketchPoint.Geometry).VectorTo(firstLine.EndSketchPoint.Geometry)
                e = eda.GetClosestPointTo(point1).DistanceTo(point1)
                d = vb.DotProduct(vc) * v2.Length * e
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
            Dim l, cl4, cl3, bl2, cl1, cl2 As SketchLine3D
            Dim ac As DimensionConstraint3D
            Dim dc, dc2, dccl4 As DimensionConstraint3D
            Dim gc, cc, gccll2 As GeometricConstraint3D
            bl2 = bandLines.Item(2)
            cl4 = constructionLines.Item(4)
            cl3 = constructionLines.Item(3)
            cl2 = constructionLines.Item(2)
            cl1 = constructionLines.Item(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartPoint, cl4.EndPoint, False)


            gapVertex.Driven = False
            If adjuster.GetMinimalDimension(gapVertex) Then
                gapVertex.Driven = True
                equalCl4.Delete()
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(bl2, l)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI * 3 / 4)
                ac.Driven = True
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)

                Try
                    gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                Catch ex As Exception
                    Try
                        ac.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI * 15 / 16)
                        ac.Driven = True
                    Catch ex2 As Exception
                        If adjuster.AdjustDimensionConstraint3DSmothly(ac, ac.Parameter._Value * 5 / 4) Then
                        Else
                            adjuster.AdjustDimensionConstraint3DSmothly(ac, ac.Parameter._Value * 9 / 8)
                        End If
                        ac.Driven = True
                    End Try
                    gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                End Try

                Try
                    dc.Delete()
                Catch ex As Exception

                End Try


                ac.Driven = True
                'gc = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                ' gc.Delete()
                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl4)

                Try
                    'adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI / 2)
                Catch ex As Exception
                    dc.Driven = True
                End Try
                dc.Driven = True
                'cc.Delete()

                'sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, cl3)
                dc2 = sk3D.DimensionConstraints3D.AddLineLength(l)
                Try
                    adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value * 1)
                    Try
                        gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, cl4)
                    Catch ex As Exception
                        dc2.Driven = True
                        dc.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI / 2)
                        dc.Driven = True
                        gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, cl4)
                        dc2.Driven = False
                    End Try
                    doku.Update()
                    cc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, cl3)
                    gapVertex.Driven = False
                    Try
                        gapVertex.Driven = False
                        adjuster.GetMinimalDimension(gapVertex)
                    Catch ex As Exception
                        gapVertex.Driven = True
                    End Try
                    gapVertex.Driven = True
                    If sk3D.Name = "s7" Then
                        Try
                            dc2.Parameter._Value = GetParameter("b")._Value * 2
                            doku.Update()

                        Catch ex3 As Exception
                            adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value * 2)
                        End Try
                    End If

                Catch ex As Exception
                    adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value * 1)
                    If sk3D.Name = "s7" Then
                        Try
                            dc2.Parameter._Value = GetParameter("b")._Value * 2
                            doku.Update()
                        Catch ex3 As Exception
                            adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value * 2)
                        End Try
                    End If
                End Try
                Try
                    dc2.Delete()
                Catch ex As Exception

                End Try
                Try
                    gc.Delete()
                Catch ex As Exception

                End Try
                Try
                    gccll2.Delete()
                Catch ex As Exception

                End Try
                Try
                    equalCl4 = sk3D.GeometricConstraints3D.AddEqual(cl4, cl2)

                Catch ex As Exception
                    dccl4 = sk3D.DimensionConstraints3D.AddLineLength(cl4)
                    adjuster.AdjustDimensionConstraint3DSmothly(dccl4, GetParameter("b")._Value)
                    dccl4.Driven = True
                    equalCl4 = sk3D.GeometricConstraints3D.AddEqual(cl4, cl2)

                End Try

                Try
                    gapVertex.Driven = False
                    Try
                        adjuster.GetMinimalDimension(gapVertex)
                    Catch ex As Exception
                    End Try
                Catch ex As Exception
                    gapVertex.Delete()
                    gapVertex = sk3D.DimensionConstraints3D.AddLineLength(cl1)
                    Try
                        adjuster.GetMinimalDimension(gapVertex)
                    Catch ex5 As Exception

                    End Try
                End Try

                '  End If
                ' dc.Delete()
            Else



                gapVertex.Driven = True

            End If
            Try
                gc.Delete()
            Catch ex As Exception

            End Try



            doku.Update2(True)
            thirdLine = l
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
            Dim l, cl4 As SketchLine3D
            Dim v As Vector
            Dim p As Plane
            Dim dc As DimensionConstraint3D
            Dim gc As GeometricConstraint3D
            Dim c As Integer = 0

            cl4 = constructionLines.Item(4)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, cl4.StartPoint, False)
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            Try
                While (Math.Abs(l.Length - thirdLine.Length) > thicknessCM * 1 And c < 128)
                    Try
                        gapVertex.Driven = True
                        adjuster.AdjustDimensionConstraint3DSmothly(dc, thirdLine.Length)
                        gapVertex.Driven = False
                        doku.Update2()
                    Catch ex As Exception
                        gapVertex.Driven = True
                    End Try
                    c = c + 1
                End While
            Catch ex1 As Exception
                gapVertex.Driven = True
            End Try

            dc.Delete()


            gc = sk3D.GeometricConstraints3D.AddEqual(l, thirdLine)
            'dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(thirdLine, secondLine)
            'If adjuster.GetMaximalDimension(dc) Then
            'dc.Driven = True
            ' Else
            'dc.Delete()
            ' End If
            Try
                gapVertex.Driven = False
                doku.Update2()
                adjuster.GetMinimalDimension(gapVertex)


            Catch ex As Exception
                gapVertex.Driven = True
            End Try
            gapVertex.Driven = True
            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawFifthConstructionLine() As SketchLine3D
        Try
            Dim l, l2, l3, l4, l7, cl4, cl2, cl3, bl3, bl4 As SketchLine3D
            Dim v As Vector
            Dim gapTry, dc2, ac, dccl5, dc5 As DimensionConstraint3D
            Dim gc, gcclcl7 As GeometricConstraint3D
            Dim c As Integer = 0
            Dim tlimt As Integer = 4
            Dim pl As Plane
            Dim vnpl, vcl3, vbl3, vcl4, vl, vcl7, vbl4, vcl7l4 As Vector
            Dim endPoint As Point
            Dim d, tangent As Double
            Dim limit As Integer = 0
            tangent = gap1CM
            pl = workFace.Geometry
            cl4 = constructionLines.Item(4)
            cl3 = constructionLines.Item(3)
            vcl4 = cl4.Geometry.Direction.AsVector
            cl2 = constructionLines.Item(2)
            bl3 = bandLines.Item(3)
            bl4 = bandLines.Item(4)
            vbl3 = bl3.Geometry.Direction.AsVector()
            vcl3 = vcl4.CrossProduct(vbl3)
            vcl3 = cl3.Geometry.Direction.AsVector
            vcl3.ScaleBy(tangent)
            endPoint = bl3.EndSketchPoint.Geometry
            endPoint.TranslateBy(vcl3)
            l = sk3D.SketchLines3D.AddByTwoPoints(bl3.EndPoint, endPoint, False)
            l.Construction = True

            sk3D.GeometricConstraints3D.AddPerpendicular(bl3, l)
            Try
                sk3D.GeometricConstraints3D.AddPerpendicular(cl4, l)

            Catch ex As Exception
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl4)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(cl4, l)
            End Try


            gapTwist = sk3D.DimensionConstraints3D.AddLineLength(l)
            gapTwist.Driven = True

            l2 = sk3D.SketchLines3D.AddByTwoPoints(bl3.StartPoint, l.EndPoint, False)
            l2.Construction = True



            vcl4.ScaleBy(-1)
            endPoint = bl4.EndSketchPoint.Geometry
            endPoint.TranslateBy(vcl4)
            l3 = sk3D.SketchLines3D.AddByTwoPoints(bl4.EndPoint, endPoint, False)
            l3.Construction = True




            sk3D.GeometricConstraints3D.AddParallel(minorLine, l3)
            'sk3D.GeometricConstraints3D.AddGround(l3)
            dc5 = sk3D.DimensionConstraints3D.AddLineLength(l3)
            'sk3D.GeometricConstraints3D.AddCoincident(l3.EndPoint, bl4)
            vl = l.Geometry.Direction.AsVector
            vcl7 = l3.Geometry.Direction.AsVector

            vcl3 = cl3.Geometry.Direction.AsVector
            endPoint = bl4.EndSketchPoint.Geometry
            endPoint.TranslateBy(vcl3)
            l7 = sk3D.SketchLines3D.AddByTwoPoints(bl4.EndPoint, endPoint, False)
            l7.Construction = True

            sk3D.GeometricConstraints3D.AddPerpendicular(l7, l3)
            sk3D.GeometricConstraints3D.AddPerpendicular(l7, bl4)
            dccl5 = sk3D.DimensionConstraints3D.AddTwoLineAngle(cl4, l7)
            d = l.Geometry.Direction.AsVector.DotProduct(l3.Geometry.Direction.AsVector)
            Try
                gapTwist.Driven = False
                While d < 1 / (limit * limit + 1) And limit < tlimt * 8
                    adjuster.AdjustDimensionConstraint3DSmothly(dccl5, dccl5.Parameter._Value * 7 / 8)
                    d = l.Geometry.Direction.AsVector.DotProduct(vcl7)
                    limit = limit + 1
                End While

                dccl5.Driven = True
            Catch ex As Exception
                dccl5.Driven = True

            End Try
            limit = 0
            gapTry = sk3D.DimensionConstraints3D.AddTwoLineAngle(l3, cl4)
            gapVertex.Driven = True
            dc2 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
            dc2.Driven = True
            ' gcclcl7 = sk3D.GeometricConstraints3D.AddCoincident(l3.EndPoint, cl4)
            Try
                While ((gapTwist.Parameter._Value + gapTry.Parameter._Value > 3 * gap1CM / 2) And limit < tlimt)
                    Try
                        Try
                            dc2.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(dc2, dc2.Parameter._Value * 9 / 8)
                            dc2.Driven = True
                        Catch ex As Exception
                            dc2.Driven = True
                        End Try

                        gapTry.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(gapTry, gapTry.Parameter._Value * 7 / 8)
                        gapTry.Driven = True

                        limit = limit + 1
                    Catch ex As Exception
                        limit = limit + 2
                        gapTry.Driven = True
                        gapTwist.Driven = True
                    End Try

                End While
                limit = 0
            Catch ex As Exception
                gapTry.Driven = True
                gapTwist.Driven = True
            End Try
            vl = l.Geometry.Direction.AsVector
            vcl7 = l3.Geometry.Direction.AsVector
            If vcl7.DotProduct(vl) < 0 Then
                endPoint = l.StartSketchPoint.Geometry
                vl.ScaleBy(-1)
                endPoint.TranslateBy(vl)
                l.Delete()
                l = sk3D.SketchLines3D.AddByTwoPoints(bl3.EndPoint, endPoint, False)
                l.Construction = True
                gapTwist = sk3D.DimensionConstraints3D.AddLineLength(l)
                adjuster.AdjustDimensionConstraint3DSmothly(gapTwist, gap1CM)
                gapTwist.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(bl3, l)
                Try
                    sk3D.GeometricConstraints3D.AddPerpendicular(cl4, l)

                Catch ex As Exception
                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl4)
                    adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                    ac.Driven = True
                    sk3D.GeometricConstraints3D.AddPerpendicular(cl4, l)
                End Try
                l2.Delete()
                l2 = sk3D.SketchLines3D.AddByTwoPoints(bl3.StartPoint, l.EndPoint, False)
                l2.Construction = True
            End If
            constructionLines.Add(l)
            constructionLines.Add(l2)
            constructionLines.Add(l3)
            constructionLines.Add(l7)
            Try

                gapTwist.Driven = False
                l4 = sk3D.SketchLines3D.AddByTwoPoints(l.EndPoint, l2.Geometry.MidPoint, False)
                l4.Construction = True
                sk3D.GeometricConstraints3D.AddCollinear(l4, l2)
                TryFixTwistGap(l4)
                constructionLines.Add(l4)
                gapTwist.Driven = True
                Try
                    gapTwist.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(gapTwist, gap1CM)

                Catch ex As Exception
                    gapTwist.Driven = True
                End Try
            Catch ex As Exception
                gapTwist.Driven = True
            End Try
            Try
                gapTry.Driven = True
                gapTwist.Driven = False
                gapVertex.Driven = False
                adjuster.GetMinimalDimension(gapVertex)
            Catch ex As Exception
                gapVertex.Driven = True
            End Try
            lastLine = l
            constructionLines.Add(l)
            l.Construction = True
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function TryFixTwistGap(l4 As SketchLine3D) As GeometricConstraint3D
        Dim bl4 As SketchLine3D
        Dim t As GeometricConstraint3D
        Dim dc2 As DimensionConstraint3D

        Try
            bl4 = bandLines.Item(4)
            t = sk3D.GeometricConstraints3D.AddCoincident(l4.EndPoint, bl4)
        Catch ex As Exception
            Try
                dc2.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(dc2, dc2.Parameter._Value * 17 / 16)
                dc2.Driven = True
            Catch ex2 As Exception
                dc2.Driven = True
            End Try
            Try
                gapTry.Driven = False
                adjuster.GetMinimalDimension(gapTry)
                gapTry.Driven = True
            Catch ex3 As Exception
                gapTry.Driven = True
            End Try

            TryFixTwistGap(l4)

        End Try


        Return t
    End Function
    Function CalculateTwistFactor() As Double
        Dim d As Double
        Dim pl As Plane
        Dim vnpl, vl3cl4, vl3, vcl4 As Vector

        Dim cl4, bl3 As SketchLine3D
        pl = workFace.Geometry

        vnpl = pl.Normal.AsVector
        vnpl.ScaleBy(-1)
        cl4 = constructionLines.Item(4)
        vcl4 = cl4.Geometry.Direction.AsVector

        bl3 = bandLines.Item(3)
        vl3 = bl3.Geometry.Direction.AsVector()
        vl3cl4 = vcl4.CrossProduct(vl3)
        d = vnpl.DotProduct(vl3cl4)
        Return d
    End Function
    Sub TryDeletePerpendicular1()
        Try
            perpendicular1.Delete()
        Catch ex As Exception
        End Try
    End Sub
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
