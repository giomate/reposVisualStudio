Imports Inventor


Public Class TwistFold7
    Dim doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, connectLine, nextLine, centroLine, bendLine3D, tangentLine, zAxisLine, radius As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, startPoint, centroPoint As Point
    Dim tg As TransientGeometry
    Dim initialPlane, adjacentPlane As Plane
    Dim gap1CM, thicknessCM, diff1, diff2, radiusValue As Double
    Dim partNumber As Integer
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines, intersectionPoints As ObjectCollection
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
    Dim sptRBack, sptRFront As SketchPoint3D
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, twistFace, nextworkFace As Face
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex, gapTwist, gapTry, perpendicular1, cruzeta, dcCentroLine, angleTangent As DimensionConstraint3D
    Public outletGap As DimensionConstraint3D
    Dim equalCl4, punchBandLine3 As GeometricConstraint3D
    Dim folded As FoldFeature
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim lamp As Highlithing
    Dim bender As Doblador
    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails As ObjectCollection
    Dim manager As FoldingEvaluator
    Dim cilindro As RodMaker

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
        intersectionPoints = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        thicknessCM = compDef.Thickness._Value
        gap1CM = 3 / 10
        bender = New Doblador(doku)
        nombrador = New Nombres(doku)
        manager = New FoldingEvaluator(doku)
        cilindro = New RodMaker(doku)

        diff1 = Math.PI / 2 - 1.64
        diff2 = Math.PI / 2 - 1.43
        radiusValue = 20 / 10
        done = False
    End Sub
    Public Function MakeFinalTwist() As Boolean
        Dim ws As WorkSurface
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
                connectLine = compDef.Sketches3D.Item("kanteLine").SketchLines3D.Item(1)
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
                    If MakeReferenceSketch().Length > 0 Then


                        comando.MakeInvisibleSketches(doku)
                        comando.MakeInvisibleWorkPlanes(doku)
                        doku.Save()
                        done = 1
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

            cutfeature = lastCut.MakeLastCut(secondLine)
            cutfeature.Name = "lastTwistCut"
            cutLine = lastCut.lastCutLine
            cutLine3D = lastCut.cutLine3D
            lamp.FitView(doku)
        End If
        Return cutfeature
    End Function
    Function MakeReferenceSketch() As SketchLine3D
        Try
            Dim l As SketchLine3D
            sk3D = compDef.Sketches3D.Add()
            l = sk3D.Include(nextLine)
            sk3D.Name = "introLine"
            tante3D = connectLine
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
        Dim cl5, bl3 As SketchLine3D
        Try
            If GetWorkFace().SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                lamp.LookAtFace(workFace)
                If curvas3D.DrawTrobinaCurve(nombrador.GetQNumber(doku), nombrador.GetNextSketchName(doku)).Construction Then
                    sk3D = curvas3D.sk3D
                    curve = curvas3D.curve
                    lamp.ZoomSelected(curve)
                    If GetMajorEdge(workFace).GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                        If DrawSingleLines().Length > 0 Then

                            nextSketch = New OriginSketch(doku)
                            Dim tl As SketchLine3D
                            tl = bandLines.Item(3)
                            cl5 = constructionLines.Item(constructionLines.Count)
                            bl3 = bandLines.Item(3)
                            gapFold.Driven = True
                            connectLine = nextSketch.DrawNextStartSketch(refLine, tl, firstLine, zAxisLine, sptRFront, outletGap)
                            If connectLine.Length > 0 Then
                                centroLine = nextSketch.centroLine
                                If ReduceGap(nextSketch.gapFold) Then
                                    Dim sl As SketchLine3D
                                    sl = bandLines(1)
                                    If bender.GetBendLine(workFace, sl).Length > 0 Then
                                        bendLine = bender.bendLine
                                        sl = bandLines(2)
                                        If bender.GetFoldingAngle(leadingEdge, sl).Parameter._Value > 0 Then
                                            comando.MakeInvisibleSketches(doku)
                                            comando.MakeInvisibleWorkPlanes(doku)
                                            folded = bender.FoldBand(bandLines.Count + 1)
                                            folded = CheckFoldSide(folded)
                                            Try
                                                folded.Name = "f7"
                                            Catch ex As Exception
                                                folded.Name = "f9"
                                            End Try
                                            doku.Update2(True)
                                            If monitor.IsFeatureHealthy(folded) Then
                                                comando.UnfoldBand(doku)
                                                If compDef.HasFlatPattern Then
                                                    comando.RefoldBand(doku)
                                                    cl5 = constructionLines.Item(5)
                                                    foldFeature = bender.folded
                                                    sk3D = compDef.Sketches3D.Add()
                                                    sk3D.Include(firstLine)
                                                    sk3D.Name = "firstLine"
                                                    sk3D = compDef.Sketches3D.Add()
                                                    sk3D.Include(secondLine)
                                                    sk3D.Name = "secondLine"
                                                    sk3D = compDef.Sketches3D.Add()
                                                    nextLine = connectLine
                                                    sk3D.Include(nextLine)
                                                    sk3D.Name = "nextline"
                                                    sk3D = compDef.Sketches3D.Add()
                                                    connectLine = nextSketch.constructionLines.Item(5)
                                                    sk3D.Include(connectLine)
                                                    sk3D.Name = "kanteLine"
                                                    sk3D = compDef.Sketches3D.Add()
                                                    sk3D.Include(bl3)
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

    Function ReduceGap(ng As DimensionConstraint3D) As Boolean
        Dim reduced As Boolean = False

        Dim inside As Boolean = IsFirstLineInside()
        Dim limit As Integer = 0
        Dim dc As DimensionConstraint3D
        Dim cl5 As SketchLine3D = nextSketch.constructionLines.Item(5)
        Try


            gapFold.Driven = True
            nextSketch.CorrectFirstLine()
            Do

                nextSketch.CorrectEntryGap()
                'adjuster.AdjustDimensionConstraint3DSmothly(nextSketch.gapFold, nextSketch.gapFold.Parameter._Value * 15 / 16)
                'nextSketch.gapFold.Driven = True
                nextSketch.CorrectGap()

                CorrectGaptFold()
                nextSketch.CorrectSecondLine()
                TryPerpendicularGap(nextSketch.angleTangent, nextSketch.gapFold)
                nextSketch.angleTangent.Driven = True
                CorrectFoldAngle()
                reduced = (TryPerpendicularGap(angleTangent, gapFold) Or reduced)
                limit = limit + 1
            Loop Until (limit > 16 Or (gapFold.Parameter._Value < 2 * gap1CM) Or (nextSketch.gapFold.Parameter._Value < gap1CM) Or (Math.Abs(angleTangent.Parameter._Value - Math.PI / 2) < 1 / 128))
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(nextSketch.firstLine.StartPoint, cl5.StartPoint)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, 1 / 10)
            dc.Delete()
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return reduced
    End Function
    Function CorrectFoldAngle() As Boolean
        Dim cl4 As SketchLine3D = constructionLines(4)
        Dim vcl4 As Vector = cl4.Geometry.Direction.AsVector

        Dim vpl As Vector = initialPlane.Normal.AsVector
        Dim d As Double = vcl4.DotProduct(vpl)
        If d > 0 Then
            angleTangent.Driven = True
            For i = 1 To 16
                CorrectFoldAngle = adjuster.AdjustDimConstrain3DSmothly(gapFold, gap1CM * 3)
                vcl4 = cl4.Geometry.Direction.AsVector
                d = vcl4.DotProduct(vpl)
                If d < 0 Then
                    gapFold.Driven = True
                    Return CorrectFoldAngle
                End If

            Next
        Else
            Return True
        End If

        Return False
    End Function
    Function CorrectFirstLine() As Boolean

        If IsFirstLineInside() Then
            Return ForceFirstLineOutside()

        Else
            Return True
        End If

    End Function

    Function IsFirstLineInside() As Boolean
        Dim l As Line = compDef.WorkAxes.Item(3).Line
        Dim b As Double = GetParameter("b")._Value
        If l.DistanceTo(connectLine.StartSketchPoint.Geometry) < b Then
            Return True
        Else
            Return False

        End If
        Return False
    End Function
    Function ForceFirstLineOutside() As Integer
        Dim hecho As Boolean
        Dim l As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(zAxisLine.Geometry.MidPoint, connectLine.StartPoint, False)
        sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, zAxisLine)
        TryPerpendicular(zAxisLine, l)
        Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(l)
        Dim b As Double = GetParameter("b")._Value
        hecho = adjuster.AdjustDimensionConstraint3DSmothly(dc, b * 5 / 4)
        dc.Delete()
        l.Delete()
        Return hecho
        Return hecho
    End Function
    Function CorrectGaptFold() As Boolean
        Dim state As Boolean = gapFold.Driven
        If gapFold.Parameter._Value < gap1CM / 16 Then
            If state Then
                adjuster.AdjustDimensionConstraint3DSmothly(gapFold, 2 * gap1CM)
                Try
                    gapFold.Driven = True
                    Return True
                Catch ex As Exception

                End Try
            Else
                Return adjuster.AdjustDimensionConstraint3DSmothly(gapFold, 2 * gap1CM)
            End If

        Else
            Return True
        End If

    End Function
    Function CorrectGaptFold(dc As DimensionConstraint3D) As Boolean
        Dim state As Boolean = dc.Driven
        If gapFold.Parameter._Value < gap1CM / 8 Then
            If state Then
                adjuster.AdjustDimensionConstraint3DSmothly(dc, gap1CM)
                Try
                    dc.Driven = True
                    Return True
                Catch ex As Exception

                End Try
            Else
                Return adjuster.AdjustDimensionConstraint3DSmothly(dc, gap1CM)
            End If

        Else
            Return True
        End If

    End Function
    Function IsGapFoldInside() As Boolean
        Dim l As Line = compDef.WorkAxes.Item(3).Line
        Dim b As Double = GetParameter("b")._Value
        If l.DistanceTo(connectLine.StartSketchPoint.Geometry) < b Then
            Return True
        Else
            Return False

        End If
        Return False
    End Function
    Function ConnectTwistBrigde() As Boolean
        Try

            Dim cnl, nssl, bl4, zAxisLine As SketchLine3D
            Dim dc1, dccl, accl, dcgvc, dcl4, acza As DimensionConstraint3D
            Dim gc, gcpp As GeometricConstraint3D
            Dim limit As Integer = 0
            Dim d, e As Double
            Dim b As Double = GetParameter("b")._Value
            bl4 = bandLines.Item(4)
            Try
                gapVertex.Driven = False
                sk3D.Solve()
                adjuster.GetMinimalDimension(gapVertex)


            Catch ex As Exception
                gapVertex.Driven = True
            End Try
            gapVertex.Driven = True

            cnl = sk3D.SketchLines3D.AddByTwoPoints(refLine.EndPoint, connectLine.EndPoint, False)

            dc1 = sk3D.DimensionConstraints3D.AddLineLength(cnl)
            ' dc2 = sk3D.DimensionConstraints3D.AddLineLength(tl2)
            Try
                adjuster.AdjustDimensionConstraint3DSmothly(dc1, dc1.Parameter._Value * 2 / 3)
            Catch ex As Exception
                dc1.Driven = True
            End Try

            dc1.Driven = True
            ' cnl1.Delete()
            Try
                gapVertex.Driven = False
                sk3D.Solve()
                adjuster.GetMinimalDimension(gapVertex)
            Catch ex As Exception
                gapVertex.Driven = True
            End Try
            dcCentroLine = sk3D.DimensionConstraints3D.AddLineLength(centroLine)
            dcCentroLine.Driven = True
            nssl = nextSketch.secondLine
            accl = sk3D.DimensionConstraints3D.AddTwoLineAngle(cnl, nssl)
            Try
                While ((gapVertex.Parameter._Value > gap1CM Or accl.Parameter._Value < Math.PI) And limit < 2)
                    Try
                        accl.Driven = False
                        gapVertex.Driven = True

                        accl.Driven = False
                        gapVertex.Driven = True
                        If accl.Parameter._Value < Math.PI * 3 / 4 Then
                            adjuster.AdjustDimensionConstraint3DSmothly(accl, Math.PI * 3 / 4)
                        Else
                            adjuster.AdjustDimensionConstraint3DSmothly(accl, Math.PI)
                        End If

                        accl.Driven = True
                        CorrectCentroLine()
                        gapVertex.Driven = False
                        adjuster.GetMinimalDimension(gapVertex)
                        gapVertex.Driven = True
                        CorrectCentroLine()

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
                    gapVertex.Driven = True
                    e = firstLine.EndSketchPoint.Geometry.DistanceTo(bendEdge.GetClosestPointTo(firstLine.EndSketchPoint.Geometry))
                    If e < gap1CM / 2 Then
                        d = Math.Abs(initialPlane.Normal.AsVector.DotProduct(adjacentPlane.Normal.AsVector))
                        If d < 0.5 Then
                            dcl4 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                            While (e < gap1CM / 2) And limit < 4
                                adjuster.AdjustDimensionConstraint3DSmothly(dcl4, dcl4.Parameter._Value * 80 / 103)
                                e = firstLine.EndSketchPoint.Geometry.DistanceTo(bendEdge.GetClosestPointTo(firstLine.EndSketchPoint.Geometry))
                                limit = limit + 1
                            End While
                            limit = 0
                        Else
                            dcl4 = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                            gapVertex.Driven = True
                            While (e < gap1CM / 2) And limit < 4
                                adjuster.AdjustDimensionConstraint3DSmothly(dcl4, dcl4.Parameter._Value * 9 / 8)
                                e = firstLine.EndSketchPoint.Geometry.DistanceTo(bendEdge.GetClosestPointTo(firstLine.EndSketchPoint.Geometry))
                                limit = limit + 1
                            End While
                            limit = 0

                        End If
                        dcl4.Driven = True
                    Else
                        Try
                            gapVertex.Driven = True
                        Catch ex As Exception
                        End Try
                    End If
                Catch ex As Exception
                    Try
                        dcl4 = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                        gapVertex.Driven = True
                        limit = 0
                        While (e < gap1CM / 2) And limit < 4
                            adjuster.AdjustDimensionConstraint3DSmothly(dcl4, dcl4.Parameter._Value * 9 / 8)
                            e = firstLine.EndSketchPoint.Geometry.DistanceTo(bendEdge.GetClosestPointTo(firstLine.EndSketchPoint.Geometry))
                            limit = limit + 1
                        End While
                        limit = 0
                    Catch ex4 As Exception
                    End Try
                End Try

            Catch ex As Exception

            End Try
            If dcCentroLine.Parameter._Value > 3 * b Then
                dcCentroLine.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(dcCentroLine, b * 2)
                dcCentroLine.Driven = True
            End If
            Try
                zAxisLine = sk3D.SketchLines3D.AddByTwoPoints(nextSketch.firstLine.StartPoint, nextSketch.centroLine.EndSketchPoint.Geometry, False)
                sk3D.GeometricConstraints3D.AddParallelToZAxis(zAxisLine)
                zAxisLine.Construction = True
                acza = sk3D.DimensionConstraints3D.AddTwoLineAngle(nextSketch.thirdLine, zAxisLine)
                Try
                    adjuster.AdjustDimensionConstraint3DSmothly(acza, Math.PI / 2 * Math.Exp((-acza.Parameter._Value + Math.PI / 2) / 4))

                Catch ex As Exception

                End Try

                acza.Driven = True
                If Not monitor.IsSketch3DHealthy(sk3D) Then
                    adjuster.RecoveryUnhealthySketch(sk3D)
                End If
            Catch ex As Exception
                acza.Driven = True
            End Try

            Try
                gapVertex.Driven = False
                sk3D.Solve()
                adjuster.GetMinimalDimension(gapVertex)

            Catch ex As Exception
                gapVertex.Driven = True
            End Try



            cnl.Construction = True

            Return True
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CorrectCentroLine() As DimensionConstraint3D
        Try
            If dcCentroLine.Parameter._Value > 3 * GetParameter("b")._Value Then
                dcCentroLine.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(dcCentroLine, 5 * GetParameter("b")._Value / 4)
                dcCentroLine.Driven = True
            End If
        Catch ex As Exception

        End Try
        Return dcCentroLine
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
                                    If DrawTangentLine.Length > 0 Then
                                        Return bandLines.Item(3)
                                        If DrawTangentParallel.Length > 0 Then
                                            Return bandLines.Item(3)
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


            maxface2 = sheetMetalFeatures.FoldFeatures.Item(sheetMetalFeatures.FoldFeatures.Count).Faces.Item(1)
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
            For Each fb As Face In sheetMetalFeatures.FoldFeatures.Item(sheetMetalFeatures.FoldFeatures.Count - 1).Faces
                For Each f As Face In bendFace.TangentiallyConnectedFaces
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
            adjacentPlane = adjacentFace.Geometry
            'lamp.HighLighFace(workFace)
            lamp.HighLighFace(adjacentFace)
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
                    minOpt = 99999999
                    minEdge = minOpt
                    For Each ed As Edge In workFace.Edges
                        If ed.Equals(bendEdge) Then
                        Else
                            If ed.GetClosestPointTo(o).DistanceTo(opt) < minOpt Then
                                'lamp.HighLighObject(ed)
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
            adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 15 / 16)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 7 / 8)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 3 / 4)
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
            Try
                adjuster.GetMinimalDimension(gapVertex)
            Catch ex As Exception

            End Try

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
            Dim l, cl3, cl2, bl2, l2 As SketchLine3D
            Dim dc, dccl4, ac As DimensionConstraint3D
            Dim gc, gcpp As GeometricConstraint3D
            Dim pt As Point = secondLine.StartSketchPoint.Geometry
            Dim v As Vector

            cl3 = constructionLines.Item(3)
            cl2 = constructionLines.Item(2)
            v = cl3.Geometry.Direction.AsVector
            v.ScaleBy(-1)
            pt.TranslateBy(v)
            bl2 = bandLines.Item(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, pt, False)

            Try
                gc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, cl2)

                gcpp = TryPerpendicular(l, cl2)
                Try
                    gcpp = sk3D.GeometricConstraints3D.AddPerpendicular(l, bl2)
                Catch ex As Exception
                    l2 = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartPoint, l.StartPoint, False)
                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l2, l)
                    adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                    ac.Delete()
                    l2.Construction = True
                    gcpp = sk3D.GeometricConstraints3D.AddPerpendicular(l, l2)
                End Try

            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                Return Nothing
            End Try
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, 1)
            l.Construction = True
            constructionLines.Add(l)





            refLine = l
            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function TryPerpendicular(l As SketchLine3D, bl As SketchLine3D) As GeometricConstraint3D
        Dim ac As DimensionConstraint3D
        Dim gc As GeometricConstraint3D
        Dim limit As Integer = 0
        Dim delta As Double = 99999
        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(bl, l)
        adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
        ac.Delete()
        gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, bl)
        Return gc
    End Function
    Function TryPerpendicularGap(ac As DimensionConstraint3D, dc As DimensionConstraint3D) As Boolean

        Dim delta As Double = 99999
        Dim hecho As Boolean
        delta = (Math.PI / 2 - ac.Parameter._Value) / (4 * Math.PI)
        hecho = adjuster.AdjustDimensionConstraint3DSmothly(ac, ac.Parameter._Value * Math.Exp(delta))
        CorrectGaptFold(dc)

        ' gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, bl)
        Return hecho
    End Function
    Function CorrectTangent() As Boolean
        Dim e As Double = Math.Cos(Math.PI * (diff1 + diff2) / 2)
        Dim d As Double = CalculateEntryRodFactor()

        If d < e Then
            Dim l As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(nextSketch.firstLine.StartPoint, sptRFront, False)
            Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(l)
            Dim limit As Integer = 0
            Dim hecho As Boolean
            Do
                hecho = adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 7 / 8)
                nextSketch.CorrectFirstLine()
                nextSketch.CorrectGap()
                d = CalculateEntryRodFactor()
                limit = limit + 1
            Loop Until (d > e Or limit > 16)
            dc.Delete()
            l.Delete()
            Return hecho
        Else
            Return True

        End If
        Return 0
    End Function
    Function CalculateEntryRodFactor() As Double
        Dim vz As Vector = tg.CreateVector(0, 0, -1)
        vz.Normalize()
        Dim vnsbl3 As Vector = nextSketch.thirdLine.Geometry.Direction.AsVector
        vnsbl3.Normalize()
        Dim d As Double = vnsbl3.DotProduct(vz)
        Return d
    End Function
    Function DrawSecondLine() As SketchLine3D
        Try

            Dim l, cl2 As SketchLine3D
            Dim dc As DimensionConstraint3D
            cl2 = constructionLines(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(cl2.EndSketchPoint.Geometry, firstLine.StartPoint, False)
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
            Dim vbl2, vnap, vcl2, vfl, ve, vf As Vector
            b = GetParameter("b")._Value
            cl2 = constructionLines.Item(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, followEdge.GetClosestPointTo(firstLine.StartSketchPoint.Geometry), False)

            gapVertex.Driven = True
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, minorLine)
            dcl = sk3D.DimensionConstraints3D.AddLineLength(l)
            d = Math.Abs(initialPlane.Normal.AsVector.DotProduct(adjacentPlane.Normal.AsVector))
            If d > 0.01 Then
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 15 / 16)
                If d > 0.1 Then
                    adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 7 / 8)
                    If d > 0.6 Then
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 3 / 4)
                        If d > 0.8 Then
                            adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 2 / 3)
                            If d > 0.9 Then
                                adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value * 1 / 2)
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
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value / 2)
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
                    If e < -1 / 16 Then
                        dcl = sk3D.DimensionConstraints3D.AddLineLength(l)
                        While e < (1 / (limit * limit + 4)) And limit < 2
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
                If d < 0.1 Then
                    Try
                        gapVertex.Driven = False
                        adjuster.GetMinimalDimension(gapVertex)
                        gapVertex.Driven = True
                    Catch ex As Exception
                        gapVertex.Driven = True
                    End Try
                End If
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
                    adjuster.AdjustDimensionConstraint3DSmothly(dcl, dcl.Parameter._Value / 2)
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
                    gapVertex.Driven = True

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
                    gapVertex.Driven = True
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
            refLine = l
            gapFold.Driven = True
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
    Function DrawTangentLine() As SketchLine3D
        Dim l, cl4, cl3, bl2, cl2, cl1 As SketchLine3D

        Dim dc As DimensionConstraint3D


        Dim spt As SketchPoint3D
        Dim c As Cylinder
        Dim b As Double = GetParameter("b")._Value
        Dim r As Double = radiusValue

        Try
            bl2 = bandLines.Item(2)

            cl3 = constructionLines.Item(3)
            cl2 = constructionLines.Item(2)
            centroPoint = compDef.WorkPoints.Item(1).Point

            zAxisLine = sk3D.SketchLines3D.AddByTwoPoints(compDef.WorkPoints.Item(1), bl2.EndSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddParallelToZAxis(zAxisLine)
            zAxisLine.Construction = True
            Try
                c = tg.CreateCylinder(centroPoint, zAxisLine.Geometry.Direction, r)
                intersectionPoints.Clear()

                For Each pt As Point In tg.CurveSurfaceIntersection(curve.Geometry, c)
                    spt = sk3D.SketchPoints3D.Add(pt)
                    If spt.Geometry.Z > 0 Then
                        sptRFront = spt
                    Else
                        sptRBack = spt
                    End If
                    intersectionPoints.Add(spt)
                Next
                radius = sk3D.SketchLines3D.AddByTwoPoints(zAxisLine.Geometry.MidPoint, sptRBack.Geometry, False)
                sk3D.GeometricConstraints3D.AddCoincident(radius.StartPoint, zAxisLine)
                dc = sk3D.DimensionConstraints3D.AddLineLength(radius)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, r)
                dc.Parameter._Value = r
                outletGap = sk3D.DimensionConstraints3D.AddTwoPointDistance(firstLine.EndPoint, sptRBack)
                outletGap.Driven = True
                TryPerpendicular(zAxisLine, radius)
                radius.Construction = True
                l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartPoint, radius.EndSketchPoint.Geometry, False)
                sk3D.GeometricConstraints3D.AddCoincident(radius.EndPoint, l)
                TryPerpendicular(radius, l)

            Catch ex2 As Exception

            End Try

            gapVertex.Driven = False
            Try
                adjuster.GetMinimalDimension(gapVertex)
            Catch ex As Exception
            End Try
            Try
                sk3D.Solve()
            Catch ex As Exception
                cl1 = constructionLines(1)
                gapVertex.Delete()
                gapVertex = sk3D.DimensionConstraints3D.AddLineLength(cl1)
            End Try

            thirdLine = l
            lastLine = l
            bandLines.Add(l)

            Return l

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function

    Function DrawThirdLine() As SketchLine3D
        Try
            Dim l, cl4, cl3, bl2, cl1, cl2 As SketchLine3D
            Dim acbl2bl3 As DimensionConstraint3D
            Dim acbl3cl4, acl2l3, dcl, dccl4 As DimensionConstraint3D
            Dim gc, cc, gccll2, gcplcl4 As GeometricConstraint3D
            Dim d, e As Double
            Dim v1, v2 As Vector
            Dim limit As Integer = 0
            bl2 = bandLines.Item(2)
            cl4 = constructionLines.Item(4)
            cl3 = constructionLines.Item(3)
            cl2 = constructionLines.Item(2)
            cl1 = constructionLines.Item(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartPoint, cl4.EndPoint, False)
            d = Math.Abs(initialPlane.Normal.AsVector.DotProduct(adjacentPlane.Normal.AsVector))

            gapVertex.Driven = False
            adjuster.GetMinimalDimension(gapVertex)


            gapVertex.Driven = True
            Try
                equalCl4.Delete()
            Catch ex As Exception

            End Try

            acbl2bl3 = sk3D.DimensionConstraints3D.AddTwoLineAngle(bl2, l)
            adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI * 3 / 4)
            acbl2bl3.Driven = True
            acbl3cl4 = sk3D.DimensionConstraints3D.AddLineLength(l)

            Try
                If d < 0.5 Then
                    acbl2bl3.Driven = False
                    While (gapVertex.Parameter._Value < 3 * gap1CM) And limit < 4
                        adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI / 1 * (1 - Math.Exp(-8 * acbl2bl3.Parameter._Value / (Math.PI * 1))))
                        If gapVertex.Parameter._Value > 2 * gap1CM Then
                            adjuster.AdjustDimensionConstraint3DSmothly(gapVertex, gap1CM)
                        End If
                        limit = limit + 1
                    End While
                    acbl2bl3.Driven = True
                    gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                Else
                    acbl2bl3.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI * 15 / 16)
                    acbl2bl3.Driven = True
                    gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                End If

            Catch ex As Exception
                Try
                    acbl2bl3.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI * 15 / 16)
                    acbl2bl3.Driven = True
                Catch ex2 As Exception
                    If adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, acbl2bl3.Parameter._Value * 5 / 4) Then
                    Else
                        adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, acbl2bl3.Parameter._Value * 9 / 8)
                    End If
                    acbl2bl3.Driven = True
                End Try
                Try
                    gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                Catch ex6 As Exception
                    acbl2bl3.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI)
                    acbl2bl3.Driven = True
                    gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                End Try

            End Try

            Try
                acbl3cl4.Delete()
            Catch ex As Exception

            End Try


            acbl2bl3.Driven = True
            'gc = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
            ' gc.Delete()
            acbl3cl4 = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl4)

            Try
                'adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI / 2)
            Catch ex As Exception
                acbl3cl4.Driven = True
            End Try
            acbl3cl4.Driven = True
            'cc.Delete()

            'sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, cl3)
            dcl = sk3D.DimensionConstraints3D.AddLineLength(l)
            Try
                Try
                    adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 1)
                Catch ex As Exception
                    Try
                        dcl.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 1)
                    Catch ex3 As Exception

                    End Try

                End Try
                Try
                    gapVertex.Driven = False
                    adjuster.GetMinimalDimension(gapVertex)
                    gapVertex.Driven = True
                Catch ex5 As Exception
                    gapVertex.Driven = True
                End Try

                Try
                    dcl.Driven = True
                    acbl3cl4.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(acbl3cl4, Math.PI / 2)
                    acbl3cl4.Driven = True
                    v1 = bl2.Geometry.Direction.AsVector
                    v2 = l.Geometry.Direction.AsVector
                    e = v2.DotProduct(v1)
                    If e > 0.99 Then
                        dccl4 = sk3D.DimensionConstraints3D.AddLineLength(cl4)
                        adjuster.AdjustDimensionConstraint3DSmothly(dccl4, dccl4.Parameter._Value * 3 / 2)
                        dccl4.Delete()

                        If (acbl2bl3.Parameter._Value < Math.PI / 2) Then
                            gccll2.Delete()
                            acbl2bl3.Driven = False
                            acbl2bl3.Parameter._Value = 0.01
                            sk3D.Solve()
                            adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI / 3)
                            adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI * 2 / 3)
                            adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI)
                            acbl2bl3.Driven = True
                            gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                        Else
                            gccll2.Delete()
                            acbl2bl3.Driven = False
                            acbl2bl3.Parameter._Value = 0.01
                            sk3D.Solve()

                            acbl2bl3.Driven = True
                            gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                        End If
                    Else
                        gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, cl4)
                    End If


                    dcl.Driven = False
                Catch ex As Exception
                    Try
                        If e > 0.99 Then
                            dccl4 = sk3D.DimensionConstraints3D.AddLineLength(cl4)
                            adjuster.AdjustDimensionConstraint3DSmothly(dccl4, dccl4.Parameter._Value * 3 / 2)
                            dccl4.Delete()
                            If (acbl2bl3.Parameter._Value < Math.PI / 2) Then
                                gccll2.Delete()
                                acbl2bl3.Driven = False
                                acbl2bl3.Parameter._Value = 0.01
                                sk3D.Solve()
                                adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI / 3)
                                adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI * 2 / 3)
                                adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI)
                                acbl2bl3.Driven = True
                                gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                            Else
                                gccll2.Delete()
                                acbl2bl3.Driven = False
                                acbl2bl3.Parameter._Value = 0.01
                                sk3D.Solve()

                                acbl2bl3.Driven = True
                                gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                            End If
                        Else
                            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, cl4)
                        End If

                    Catch ex6 As Exception

                    End Try
                    dcl.Driven = True
                    Try
                        gccll2.Delete()
                    Catch ex3 As Exception
                    End Try
                    acbl3cl4.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(acbl3cl4, Math.PI / 2)
                    acbl3cl4.Driven = True
                    Try
                        gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, cl4)

                    Catch ex3 As Exception

                    End Try
                    Try
                        dcl.Driven = False
                    Catch ex10 As Exception

                    End Try
                    Try
                        gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                    Catch ex6 As Exception

                    End Try


                End Try
                Try
                    sk3D.Solve()
                Catch ex As Exception
                End Try

                If dcl.Parameter._Value < GetParameter("b")._Value Then
                    Try
                        dcl.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 1)
                    Catch ex3 As Exception
                    End Try
                End If

                Try
                    v1 = bl2.Geometry.Direction.AsVector
                    v2 = l.Geometry.Direction.AsVector
                    e = v2.DotProduct(v1)
                    Try
                        If e > 0.99 Then
                            dccl4 = sk3D.DimensionConstraints3D.AddLineLength(cl4)
                            adjuster.AdjustDimensionConstraint3DSmothly(dccl4, dccl4.Parameter._Value * 3 / 2)
                            dccl4.Delete()
                            If (acbl2bl3.Parameter._Value < Math.PI / 2) Then
                                gccll2.Delete()
                                acbl2bl3.Driven = False
                                acbl2bl3.Parameter._Value = 0.01
                                sk3D.Solve()
                                adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI / 3)
                                adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI * 2 / 3)
                                adjuster.AdjustDimensionConstraint3DSmothly(acbl2bl3, Math.PI)
                                acbl2bl3.Driven = True
                                gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                            Else
                                gccll2.Delete()
                                acbl2bl3.Driven = False
                                acbl2bl3.Parameter._Value = 0.01
                                sk3D.Solve()

                                acbl2bl3.Driven = True
                                gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                            End If
                        Else
                            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, cl4)
                        End If

                    Catch ex6 As Exception

                    End Try
                    cc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, cl3)
                Catch ex As Exception
                    gapVertex.Driven = True
                    Try
                        gccll2.Delete()
                    Catch ex7 As Exception

                    End Try

                    Try
                        cc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, cl3)
                    Catch ex6 As Exception
                    End Try
                    gccll2 = sk3D.GeometricConstraints3D.AddCollinear(l, bl2)
                End Try

                Try
                    gapVertex.Driven = False
                    adjuster.GetMinimalDimension(gapVertex)
                    gapVertex.Driven = True
                Catch ex As Exception
                    gapVertex.Driven = True
                End Try
                gapVertex.Driven = True
                If sk3D.Name = "s7" Then
                    Try
                        Try
                            gcplcl4 = sk3D.GeometricConstraints3D.AddParallel(cl4, minorLine)
                            gcplcl4.Delete()
                        Catch ex As Exception
                            Try
                                gcplcl4.Delete()
                            Catch ex6 As Exception
                            End Try
                        End Try
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 3 / 2)
                        Try
                            sk3D.Solve()
                        Catch ex As Exception

                        End Try


                    Catch ex3 As Exception
                        Try
                            gcplcl4 = sk3D.GeometricConstraints3D.AddParallel(cl4, minorLine)
                            gcplcl4.Delete()
                        Catch ex As Exception
                            Try
                                gcplcl4.Delete()
                            Catch ex6 As Exception
                            End Try
                        End Try
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 3 / 2)
                    End Try
                End If

            Catch ex As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 1)
                If sk3D.Name = "s7" Then
                    Try
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 3 / 2)

                        Try
                            gcplcl4 = sk3D.GeometricConstraints3D.AddParallel(cl4, minorLine)
                            gcplcl4.Delete()
                        Catch ex6 As Exception
                            Try
                                gcplcl4.Delete()
                            Catch ex8 As Exception
                            End Try
                        End Try
                        sk3D.Solve()
                    Catch ex3 As Exception
                        Try
                            gcplcl4 = sk3D.GeometricConstraints3D.AddParallel(cl4, minorLine)
                            gcplcl4.Delete()
                        Catch ex7 As Exception
                            Try
                                gcplcl4.Delete()
                            Catch ex6 As Exception
                            End Try
                        End Try
                        adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 3 / 2)
                    End Try
                End If
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

                If dcl.Parameter._Value > GetParameter("b")._Value * 2 Then
                    adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 3 / 2)
                End If

            Catch ex As Exception
                dccl4 = sk3D.DimensionConstraints3D.AddLineLength(cl4)
                adjuster.AdjustDimensionConstraint3DSmothly(dccl4, GetParameter("b")._Value)
                dccl4.Driven = True
                equalCl4 = sk3D.GeometricConstraints3D.AddEqual(cl4, cl2)

            End Try
            Try
                dcl.Delete()
            Catch ex As Exception
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

            Try
                gc.Delete()
            Catch ex As Exception

            End Try



            sk3D.Solve()
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
            Dim l, cl4, cl3, bl3 As SketchLine3D
            Dim vbl3, vbl4 As Vector
            Dim p As Plane
            Dim dc, dcbl3, ac As DimensionConstraint3D
            Dim gc, gcplbl4, gcel As GeometricConstraint3D
            Dim c, limit As Integer
            Dim lastLength, d As Double
            cl3 = constructionLines.Item(3)
            cl4 = constructionLines.Item(4)
            bl3 = bandLines.Item(3)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, cl4.StartPoint, False)

            Try
                dcbl3 = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl3.StartPoint, cl3.StartPoint)
                adjuster.AdjustDimensionConstraint3DSmothly(dcbl3, gap1CM / 16)
                Try
                    dcbl3.Delete()
                Catch ex As Exception
                End Try
                Try
                    punchBandLine3 = sk3D.GeometricConstraints3D.AddCoincident(bl3.StartPoint, cl3)
                Catch ex As Exception
                End Try

            Catch ex As Exception
                Try
                    dcbl3.Delete()
                Catch ex5 As Exception
                End Try
            End Try

            Try
                dcbl3 = sk3D.DimensionConstraints3D.AddLineLength(bl3)
                If dcbl3.Parameter._Value > GetParameter("b")._Value * 3 Then
                    adjuster.AdjustDimensionConstraint3DSmothly(dcbl3, GetParameter("b")._Value * 2)
                End If
                dcbl3.Delete()
            Catch ex As Exception
                Try
                    dcbl3.Delete()
                Catch ex5 As Exception
                End Try
            End Try

            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            Try
                c = 0
                While (Math.Abs(l.Length - thirdLine.Length) > thicknessCM * 1 And c < 64)
                    lastLength = l.Length
                    Try
                        gapVertex.Driven = True
                        If adjuster.AdjustDimensionConstraint3DSmothly(dc, thirdLine.Length) Then
                        Else
                            If Math.Abs(l.Length - lastLength) < thicknessCM / 10 And c > 8 Then
                                Try
                                    gcel = sk3D.GeometricConstraints3D.AddEqual(l, thirdLine)
                                    gcel.Delete()
                                    'Exit While
                                Catch ex2 As Exception
                                    Try
                                        gcel.Delete()
                                    Catch ex3 As Exception

                                    End Try
                                End Try
                            End If

                        End If

                        gapVertex.Driven = False
                        sk3D.Solve()
                        If gapVertex.Parameter._Value > gap1CM * 3 Then
                            adjuster.GetMinimalDimension(gapVertex)
                        End If
                        gapVertex.Driven = True
                        vbl3 = thirdLine.Geometry.Direction.AsVector
                        vbl4 = l.Geometry.Direction.AsVector
                        d = vbl3.DotProduct(vbl4)
                        If d < 0 Then
                            limit = 0
                            While (d < 0.25 And limit < 16)
                                Try
                                    gapVertex.Driven = True
                                    dc.Driven = True
                                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, firstLine)
                                    adjuster.AdjustDimensionConstraint3DSmothly(ac, ac.Parameter._Value * 4 / 3)
                                    ac.Delete()
                                    gapVertex.Driven = False
                                    dc.Driven = False
                                    sk3D.Solve()
                                    vbl3 = thirdLine.Geometry.Direction.AsVector
                                    vbl4 = l.Geometry.Direction.AsVector
                                    d = vbl3.DotProduct(vbl4)
                                    limit = limit + 1
                                Catch ex As Exception
                                    limit = limit + 1
                                    Try
                                        ac.Delete()
                                    Catch ex4 As Exception

                                    End Try
                                End Try
                            End While
                            limit = 0
                            adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value)

                        End If

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
            TryCL4Parallel()

            Try
                gapVertex.Driven = False
                sk3D.Solve()
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
    Function DrawTangentParallel() As SketchLine3D
        Dim l, bl3, cl4, cl1 As SketchLine3D
        Dim gc As GeometricConstraint3D
        Try

            bl3 = bandLines.Item(3)
            cl4 = constructionLines.Item(4)
            l = sk3D.SketchLines3D.AddByTwoPoints(cl4.StartPoint, bl3.EndSketchPoint.Geometry, False)
            gc = sk3D.GeometricConstraints3D.AddParallel(l, bl3)
            angleTangent = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl4)
            angleTangent.Driven = True
            '  gc = TryPerpendicularGap(l, cl4)
            l.Construction = True
            Try
                adjuster.GetMinimalDimension(gapVertex)
            Catch ex As Exception
            End Try
            Try
                sk3D.Solve()
            Catch ex As Exception
                cl1 = constructionLines(1)
                gapVertex.Delete()
                gapVertex = sk3D.DimensionConstraints3D.AddLineLength(cl1)
            End Try
            gapFold.Driven = True
            lastLine = l
            constructionLines.Add(l)

            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawFifthConstructionLine() As SketchLine3D
        Try
            Dim l, l2, l3, l4, l5, l6, cl4, cl2, cl3, bl3, bl4, clgv As SketchLine3D
            Dim v As Vector
            Dim gapTry, dc2, ac, accl7cl4, dc5, dcbl3, dccl, dcbl4, dcgt As DimensionConstraint3D
            Dim dcl7, dccl7, acl7, accl7, acflbl4 As DimensionConstraint3D
            Dim gc, gcclcl7, gcplcl4cl7 As GeometricConstraint3D
            Dim c As Integer = 0
            Dim tlimt As Integer = 4
            Dim cUpwards As Integer
            Dim pl As Plane
            Dim vnpl, vcl3, vbl3, vcl4, vl, vcl7, vbl4, vcl7l4, vwf, vfl As Vector
            Dim endPoint As Point
            Dim d, tangent, f, k, b, lastValue As Double
            Dim limit As Integer = 0
            b = GetParameter("b")._Value
            If sk3D.Name = "s7" Then
                tlimt = 6
            End If
            tangent = gap1CM
            pl = workFace.Geometry
            vwf = pl.Normal.AsVector
            vfl = firstLine.Geometry.Direction.AsVector
            cl4 = constructionLines.Item(4)
            cl3 = constructionLines.Item(3)
            vcl4 = cl4.Geometry.Direction.AsVector
            cl2 = constructionLines.Item(2)
            clgv = constructionLines.Item(1)
            bl3 = bandLines.Item(3)
            bl4 = bandLines.Item(4)
            vbl4 = bl4.Geometry.Direction.AsVector
            vbl3 = bl3.Geometry.Direction.AsVector()
            vcl3 = cl3.Geometry.Direction.AsVector
            vfl.ScaleBy(-1)
            f = (vfl).DotProduct(vcl4)
            vcl3.ScaleBy(tangent * f)

            endPoint = bl3.EndSketchPoint.Geometry
            endPoint.TranslateBy(vcl3)
            l = sk3D.SketchLines3D.AddByTwoPoints(bl3.EndPoint, endPoint, False)
            l.Construction = True
            gapTwist = sk3D.DimensionConstraints3D.AddLineLength(l)
            adjuster.AdjustDimensionConstraint3DSmothly(gapTwist, gap1CM)

            Try
                sk3D.GeometricConstraints3D.AddPerpendicular(bl3, l)

            Catch ex As Exception
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl3)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Driven = True
                Try
                    sk3D.GeometricConstraints3D.AddPerpendicular(bl3, l)
                Catch ex7 As Exception
                    ac.Driven = False

                End Try

            End Try
            Try
                sk3D.GeometricConstraints3D.AddPerpendicular(cl4, l)

            Catch ex As Exception
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl4)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Driven = True
                Try
                    sk3D.GeometricConstraints3D.AddPerpendicular(cl4, l)
                Catch ex8 As Exception
                    ac.Driven = False
                End Try

            End Try



            'gapTwist.Driven = True

            l2 = sk3D.SketchLines3D.AddByTwoPoints(bl3.StartPoint, l.EndPoint, False)
            l2.Construction = True



            vcl4.ScaleBy(-1)

            endPoint = bl4.EndSketchPoint.Geometry
            endPoint.TranslateBy(vcl4)
            l3 = sk3D.SketchLines3D.AddByTwoPoints(bl4.EndPoint, farPoint, False)
            l3.Construction = True




            gcplcl4cl7 = sk3D.GeometricConstraints3D.AddParallel(minorLine, l3)
            'sk3D.GeometricConstraints3D.AddGround(l3)
            dc5 = sk3D.DimensionConstraints3D.AddLineLength(l3)
            Try
                dc5.Parameter._Value = GetParameter("b")._Value
                sk3D.Solve()
            Catch ex As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(dc5, GetParameter("b")._Value)
                dc5.Parameter._Value = GetParameter("b")._Value
                sk3D.Solve()
            End Try

            'sk3D.GeometricConstraints3D.AddCoincident(l3.EndPoint, bl4)
            vl = l.Geometry.Direction.AsVector
            vcl7 = l3.Geometry.Direction.AsVector
            vcl4 = cl4.Geometry.Direction.AsVector
            vcl3 = cl3.Geometry.Direction.AsVector
            vbl3 = bl3.Geometry.Direction.AsVector()
            f = (vfl).DotProduct(vcl4)
            vcl3.ScaleBy(tangent)
            endPoint = bl4.EndSketchPoint.Geometry
            endPoint.TranslateBy(vcl3)
            l5 = sk3D.SketchLines3D.AddByTwoPoints(bl4.EndPoint, endPoint, False)
            l5.Construction = True
            d = Math.Abs(initialPlane.Normal.AsVector.DotProduct(adjacentPlane.Normal.AsVector))
            If d > 0.5 Then
                sk3D.GeometricConstraints3D.AddPerpendicular(l5, l3)
                Try
                    dcl7 = sk3D.DimensionConstraints3D.AddLineLength(l5)
                    sk3D.GeometricConstraints3D.AddPerpendicular(l5, bl4)

                Catch ex As Exception
                    Try
                        dcl7 = sk3D.DimensionConstraints3D.AddLineLength(l5)
                    Catch ex8 As Exception

                    End Try

                    dcl7.Parameter._Value = dcl7.Parameter._Value * 2
                    dcl7.Driven = True
                    Try
                        acl7 = sk3D.DimensionConstraints3D.AddTwoLineAngle(l5, bl4)
                        adjuster.AdjustDimensionConstraint3DSmothly(acl7, Math.PI / 2)
                        acl7.Driven = True
                    Catch ex1 As Exception
                        Try
                            acl7.Driven = True
                        Catch ex3 As Exception

                        End Try
                    End Try
                    sk3D.GeometricConstraints3D.AddPerpendicular(l5, bl4)
                End Try
            Else
                sk3D.GeometricConstraints3D.AddParallel(l5, cl3)
                Try

                    dccl7 = sk3D.DimensionConstraints3D.AddLineLength(l5)
                    Try
                        accl7 = sk3D.DimensionConstraints3D.AddTwoLineAngle(l5, l3)
                    Catch ex As Exception

                    End Try

                Catch ex As Exception

                End Try
            End If
            accl7cl4 = sk3D.DimensionConstraints3D.AddTwoLineAngle(cl4, l5)
            d = l.Geometry.Direction.AsVector.DotProduct(l3.Geometry.Direction.AsVector)
            dccl = sk3D.DimensionConstraints3D.AddTwoPointDistance(compDef.WorkPoints.Item(1), bl4.EndPoint)
            dccl.Driven = True
            Try
                gapTwist.Driven = False
                While d < 1 / (limit * limit + 1) And limit < tlimt * 8
                    lastValue = bl4.Length
                    If adjuster.AdjustDimensionConstraint3DSmothly(accl7cl4, accl7cl4.Parameter._Value * 8 / 9) Then
                        Try
                            dcbl4 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                            adjuster.AdjustDimensionConstraint3DSmothly(dcbl4, dcbl4.Parameter._Value * 1.01)
                            dcbl4.Delete()
                        Catch ex As Exception
                            Try
                                dcbl4.Delete()
                            Catch ex5 As Exception

                            End Try
                        End Try
                    Else
                        Try
                            dcbl4 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                            adjuster.AdjustDimensionConstraint3DSmothly(dcbl4, dcbl4.Parameter._Value * 1.02)
                            dcbl4.Delete()
                        Catch ex As Exception
                            Try
                                dcbl4.Delete()
                            Catch ex5 As Exception

                            End Try
                        End Try
                    End If

                    If dccl.Parameter._Value > GetParameter("b")._Value * 3 Then
                        Try
                            dccl.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(dccl, b * 2)
                            dccl.Driven = True
                        Catch ex As Exception
                            dccl.Driven = True
                        End Try

                    End If
                    Try
                        gapVertex.Driven = False
                        sk3D.Solve()
                        If gapVertex.Parameter._Value > gap1CM * 3 Then
                            If adjuster.AdjustDimensionConstraint3DSmothly(gapVertex, gap1CM) Then
                                gapVertex.Driven = True
                            Else
                                Try
                                    dcbl4 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                                    adjuster.AdjustDimensionConstraint3DSmothly(dcbl4, dcbl4.Parameter._Value * (1.04))
                                    dcbl4.Delete()
                                Catch ex As Exception
                                    Try
                                        dcbl4.Delete()
                                    Catch ex5 As Exception

                                    End Try
                                End Try
                            End If

                        End If

                    Catch ex As Exception
                        gapVertex.Driven = True
                    End Try
                    If IsBl4Upwards() Then
                        Try
                            ForceBl4Downwards()
                            Try
                                dcbl4 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                                adjuster.AdjustDimensionConstraint3DSmothly(dcbl4, GetParameter("b")._Value)
                                dcbl4.Delete()
                            Catch ex As Exception
                                Try
                                    dcbl4.Delete()
                                Catch ex5 As Exception

                                End Try
                            End Try
                            limit = 1
                        Catch ex As Exception
                            Try
                                dcbl4.Delete()
                            Catch ex5 As Exception

                            End Try
                        End Try
                    End If
                    If lastValue > bl4.Length Then
                        dcbl4 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                        adjuster.AdjustDimensionConstraint3DSmothly(dcbl4, lastValue * 8 / 7)
                        dcbl4.Delete()
                    End If
                    limit = limit + 1
                    d = l.Geometry.Direction.AsVector.DotProduct(l3.Geometry.Direction.AsVector)
                End While

                accl7cl4.Driven = True
            Catch ex As Exception
                accl7cl4.Driven = True

            End Try
            limit = 0
            Try
                gapVertex.Driven = False
                adjuster.GetMinimalDimension(gapVertex)
                gapVertex.Driven = True
            Catch ex As Exception
                gapVertex.Driven = True
            End Try
            gapTry = sk3D.DimensionConstraints3D.AddTwoLineAngle(l3, cl4)
            gapVertex.Driven = True
            dc2 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
            dc2.Driven = True
            Try

                gapTwist.Driven = False
                l4 = sk3D.SketchLines3D.AddByTwoPoints(l.EndPoint, l2.Geometry.MidPoint, False)
                l4.Construction = True
                sk3D.GeometricConstraints3D.AddCollinear(l4, l2)
                l6 = sk3D.SketchLines3D.AddByTwoPoints(l4.EndPoint, bl4.Geometry.MidPoint, False)
                l6.Construction = True
                Try
                    sk3D.GeometricConstraints3D.AddCoincident(l6.EndPoint, bl4)
                Catch ex As Exception

                End Try


            Catch ex As Exception
                gapTwist.Driven = True
            End Try
            If sk3D.Name = "s7" Then
                k = 6 / 5
            Else
                k = 8 / 7
            End If
            ' gcclcl7 = sk3D.GeometricConstraints3D.AddCoincident(l3.EndPoint, cl4)
            Try
                While ((gapTwist.Parameter._Value + gapTry.Parameter._Value > 4 * gap1CM / 3) And limit < tlimt) And l6.Length > gap1CM
                    Try
                        Try
                            dc2.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(dc2, dc2.Parameter._Value * k)
                            dc2.Driven = True
                        Catch ex As Exception
                            dc2.Driven = True
                        End Try

                        gapTry.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(gapTry, gapTry.Parameter._Value * 8 / 9)
                        gapTry.Driven = True
                        If dccl.Parameter._Value > b * 3 Then
                            Try
                                dccl.Driven = False
                                adjuster.AdjustDimensionConstraint3DSmothly(dccl, b * 2)
                                dccl.Driven = True
                            Catch ex As Exception
                                dccl.Driven = True
                            End Try

                        End If
                        Try
                            gapVertex.Driven = False
                            sk3D.Solve()
                            If gapVertex.Parameter._Value > gap1CM * 3 Then
                                adjuster.GetMinimalDimension(gapVertex)
                            End If
                            gapVertex.Driven = True
                        Catch ex As Exception
                            gapVertex.Driven = True
                        End Try


                        limit = limit + 1
                    Catch ex As Exception
                        limit = limit + 2
                        gapTry.Driven = True
                        'gapTwist.Driven = True
                    End Try

                End While
                limit = 0
            Catch ex As Exception
                gapTry.Driven = True
                ' gapTwist.Driven = True
            End Try
            Try
                dccl.Delete()
            Catch ex As Exception
            End Try
            vl = l.Geometry.Direction.AsVector
            vcl7 = l3.Geometry.Direction.AsVector

            constructionLines.Add(l)
            constructionLines.Add(l2)
            constructionLines.Add(l3)
            constructionLines.Add(l4)
            constructionLines.Add(l5)
            constructionLines.Add(l6)
            Try
                gapTwist.Driven = False
                TryFixTwistGap(l4, dc2)
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
                gapVertex.Delete()
                gapVertex = sk3D.DimensionConstraints3D.AddLineLength(clgv)
                Try
                    adjuster.GetMinimalDimension(gapVertex)
                Catch ex6 As Exception
                End Try

            End Try
            Try
                dc2.Delete()
            Catch ex As Exception

            End Try
            lastLine = l
            constructionLines.Add(l)

            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function TryCL4Parallel() As Boolean

        Dim ac1 As DimensionConstraint3D
        Dim clm, cl4 As SketchLine3D




        Try
            cl4 = constructionLines.Item(4)
            clm = sk3D.SketchLines3D.AddByTwoPoints(cl4.StartPoint, farPoint, False)
            sk3D.GeometricConstraints3D.AddParallel(clm, minorLine)

            ac1 = sk3D.DimensionConstraints3D.AddTwoLineAngle(clm, cl4)
            adjuster.AdjustDimensionConstraint3DSmothly(ac1, ac1.Parameter._Value * 5 / 4)
            adjuster.AdjustDimensionConstraint3DSmothly(ac1, ac1.Parameter._Value * 9 / 8)
            adjuster.AdjustDimensionConstraint3DSmothly(ac1, Math.PI)
            clm.Delete()

        Catch ex As Exception
            Try
                clm.Delete()
            Catch ex2 As Exception

            End Try
        End Try

    End Function
    Function IsBl4Upwards() As Boolean
        Dim vfl, vbl4 As Vector
        Dim fl, bl4 As SketchLine3D
        Dim d As Double
        bl4 = bandLines.Item(4)
        vfl = firstLine.Geometry.Direction.AsVector
        vbl4 = bl4.Geometry.Direction.AsVector
        d = vfl.DotProduct(vbl4)
        If d < 0 Then
            Return True
        Else
            Return False

        End If

    End Function
    Function ForceBl4Downwards() As Boolean

        Dim fl, bl4 As SketchLine3D
        Dim ac As DimensionConstraint3D


        Dim limit As Integer
        fl = firstLine
        bl4 = bandLines.Item(4)
        Try
            limit = 0
            While (IsBl4Upwards() And limit < 8)
                Try
                    gapVertex.Driven = True

                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(fl, bl4)
                    adjuster.AdjustDimensionConstraint3DSmothly(ac, ac.Parameter._Value * 4 / 3)
                    ac.Delete()
                    gapVertex.Driven = False

                    sk3D.Solve()

                    limit = limit + 1
                Catch ex As Exception
                    limit = limit + 1
                    Try
                        ac.Delete()
                    Catch ex4 As Exception
                    End Try
                End Try
            End While
        Catch ex As Exception
            Try
                ac.Delete()
            Catch ex4 As Exception
            End Try
        End Try
        Return (Not IsBl4Upwards())

    End Function
    Function TryFixTwistGap(l4 As SketchLine3D, dc2 As DimensionConstraint3D) As GeometricConstraint3D
        Dim bl4 As SketchLine3D
        Dim t As GeometricConstraint3D

        Dim d As Double

        Try
            bl4 = bandLines.Item(4)
            d = bl4.Length
            t = sk3D.GeometricConstraints3D.AddCoincident(l4.EndPoint, bl4)

            If d > bl4.Length Then

                adjuster.AdjustDimensionConstraint3DSmothly(dc2, d * 17 / 16)

            End If


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

            TryFixTwistGap(l4, dc2)

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
