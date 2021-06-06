Imports System.IO
Imports Inventor
Imports Subina_Design_Helpers
Public Class Wedges
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk, curvesSketch As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, coneLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim corte As IntersectionCurve
    Dim lastSeq As SketchEquationCurve3D
    Dim skptMax, skptMin As SketchPoint3D

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile
    Private converter As Conversions

    Public trobinaCurve As Curves3D
    Dim palito As RodMaker

    Public wp1, wp2, wp3, wpConverge, wptHigh, wptLow, wptConeBase As WorkPoint
    Public farPoint, curvePoint, convergePoint As Point
    Dim tg As TransientGeometry
    Dim freeCollisionAngle As Double
    Public partNumber, qNext, qLastTie As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres

    Dim cutProfile As Profile


    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Public compDef As PartComponentDefinition
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim mainWorkPlane As WorkPlane
    Dim workPointFace As WorkPoint
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, cutEdge1, cutEdge2, CutEsge3, closestEdge As Edge
    Dim ringLine, ovalLine As SketchLine3D
    Dim workFace, adjacentFace, bendFace, closestFace, corteFace, cylinderFace, sideMajorFace, sideMinorFace As Face
    Dim workFaces As FaceCollection
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex, angleConeLineTangent, angleConeNormalLine As DimensionConstraint3D
    Dim folded As FoldFeature
    Public foldFeatures As FoldFeatures
    Dim features As PartFeatures
    Dim lamp As Highlithing
    Dim di As System.IO.DirectoryInfo
    Dim fi As System.IO.File
    Dim nf As System.IO.Path
    Dim bandFaces, tangentSurfaces As WorkSurface

    Dim foldFeature As FoldFeature
    Public sections, endPoints, rails, caras, surfaceBodies, coneAxes, highPoints, lowPoints As ObjectCollection
    Dim interPoints As ObjectCollection
    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim adjuster As SketchAdjust
    Dim coneCounter As Integer



    Dim fullFileNames As String()
    Structure DesignParam
        Public p As Integer
        Public q As Integer
        Public b As Double
        Public Dmax As Double
        Public Dmin As Double

    End Structure
    Public DP As DesignParam
    Dim Tr As Double
    Dim Cr As Double

    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        invFile = New InventorFile(app)
        converter = New Conversions(app)
        projectManager = app.DesignProjectManager

        compDef = doku.ComponentDefinition

        features = doku.ComponentDefinition.Features


        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        sections = app.TransientObjects.CreateObjectCollection
        endPoints = app.TransientObjects.CreateObjectCollection
        rails = app.TransientObjects.CreateObjectCollection
        caras = app.TransientObjects.CreateObjectCollection
        surfaceBodies = app.TransientObjects.CreateObjectCollection
        coneAxes = app.TransientObjects.CreateObjectCollection
        lowPoints = app.TransientObjects.CreateObjectCollection
        highPoints = app.TransientObjects.CreateObjectCollection
        interPoints = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        adjuster = New SketchAdjust(doku)

        nombrador = New Nombres(doku)

        trobinaCurve = New Curves3D(doku)
        palito = New RodMaker(doku)
        DP.Dmax = 200 / 10
        DP.Dmin = 1 / 10
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 11
        DP.q = 23
        DP.b = 25
        coneCounter = 1

        done = False
        wptHigh = compDef.WorkPoints.Item("wptHigh")
        wptLow = compDef.WorkPoints.Item("wptLow")
        tangentSurfaces = compDef.WorkSurfaces.Item("tangents")
    End Sub
    Function DocUpdate(docu As PartDocument) As PartDocument
        doku = docu

        compDef = doku.ComponentDefinition
        lamp = New Highlithing(doku)

        Return doku
    End Function
    Public Function MakeWedgesIteration(i As Integer) As PartDocument
        Try
            Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
            Dim q As Integer
            Dim nd As String

            nd = String.Concat(p, "\Iteration", i.ToString)

            If (Directory.Exists(nd)) Then

                fullFileNames = Directory.GetFiles(nd, "*.ipt")


                For Each s As String In fullFileNames
                    If nombrador.ContainBand(s) Then
                        If Not IsWedgeDone(nombrador.MakeWedgeFileName(s)) Then
                            doku = MakeSingleWedge(s)
                            doku.ComponentDefinition.WorkSurfaces.Item(2).Visible = False
                            'FillVoids()

                            CombineBodies()

                            '  CombineBodiesDuo()
                            palito.MakeAllRods(doku.ComponentDefinition.WorkSurfaces.Item(1))

                            monitor = New DesignMonitoring(doku)
                            doku.Update2(True)

                            MakeHole()

                            bandFaces = doku.ComponentDefinition.WorkSurfaces.Item(1)

                            If monitor.IsFeatureHealthy(EmbossNumber(s)) Then
                                app.SilentOperation = True
                                doku.SaveAs(nombrador.MakeWedgeFileName(s), False)
                                doku.Save()
                                doku.Close(True)
                                app.Documents.CloseAll()
                                app.SilentOperation = False

                            End If




                        End If
                    End If

                Next

                done = True
            End If

            Return doku
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function IsWedgeDone(s As String) As Boolean
        For Each ffn As String In fullFileNames
            If ffn.Equals(s) Then
                Return True
            End If
        Next

        Return False
    End Function
    Function GetStartingface(ws As WorkSurface) As Face
        Dim min2, min1 As Double
        Dim fmin1, fmin2 As Face
        Try
            min1 = ws.SurfaceBodies.Item(1).Faces.Item(1).Evaluator.Area
            min2 = min1
            fmin1 = ws.SurfaceBodies.Item(1).Faces.Item(1)
            fmin2 = fmin1
            For Each sb As SurfaceBody In ws.SurfaceBodies
                For Each fc As Face In sb.Faces
                    If fc.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        min2 = fc.Evaluator.Area
                        If min2 < min1 Then
                            min1 = min2
                            fmin2 = fmin1
                            fmin1 = fc
                        Else
                            fmin2 = fc
                        End If
                    End If
                Next
            Next
            lamp.HighLighFace(fmin1)
            Return fmin1
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function RemoveExcessFaces(ws As WorkSurface) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim wpf, wpt, wptpf As WorkPoint
        Dim pt As Point
        Dim pl, plw As Plane
        Dim v As Vector
        Try
            workFaces = GetStartingface(ws).TangentiallyConnectedFaces
            For Each f As Face In workFaces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    Try
                        wpf = doku.ComponentDefinition.WorkPoints.AddAtCentroid(f.EdgeLoops.Item(1))
                        pl = f.Geometry

                        ' If IsPointContained(pt, doku.ComponentDefinition.SurfaceBodies.Item(1)) Then

                        For Each fw As Face In ws.SurfaceBodies.Item(1).Faces
                            If fw.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                                If Not fw.Equals(f) Then
                                    If Math.Abs(f.Evaluator.Area - fw.Evaluator.Area) < 0.01 Then
                                        plw = fw.Geometry
                                        If Math.Abs(plw.Normal.AsVector.DotProduct(pl.Normal.AsVector)) > 0.999 Then
                                            wptpf = doku.ComponentDefinition.WorkPoints.AddAtCentroid(fw.EdgeLoops.Item(1))
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        v = wpf.Point.VectorTo(wptpf.Point)
                        v.ScaleBy(f.Evaluator.Area * 24)
                        pt = wpf.Point
                        pt.TranslateBy(v)
                        wpt = doku.ComponentDefinition.WorkPoints.AddFixed(pt)
                        ef = RemoveFakeMaterial(f, wpt)
                        If monitor.IsFeatureHealthy(ef) Then
                            RemoveBorders(f, wpt, ef)
                            v.ScaleBy(1)
                            wpt.Visible = False
                        Else
                            ef.Delete()
                            Try
                                wpt.Delete()
                            Catch ex As Exception

                            End Try
                        End If


                        ' End If
                        wpf.Delete()
                        wptpf.Delete()
                    Catch ex As Exception
                        Try
                            wpf.Delete()
                            wptpf.Delete()
                        Catch ex2 As Exception

                        End Try
                    End Try
                End If

            Next
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return ef
    End Function
    Function GetRealNormal(f As Face, ws As WorkSurface) As Vector
        Dim v As Vector
        Dim wpf, wpt, wptpf As WorkPoint
        Dim pt As Point
        Dim pl, plw As Plane
        Try
            Try
                wpf = doku.ComponentDefinition.WorkPoints.AddAtCentroid(f.EdgeLoops.Item(1))
                pl = f.Geometry

                ' If IsPointContained(pt, doku.ComponentDefinition.SurfaceBodies.Item(1)) Then

                For Each fw As Face In ws.SurfaceBodies.Item(1).Faces
                    If fw.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        If Not fw.Equals(f) Then
                            If Math.Abs(f.Evaluator.Area - fw.Evaluator.Area) < 0.01 Then
                                plw = fw.Geometry
                                If Math.Abs(plw.Normal.AsVector.DotProduct(pl.Normal.AsVector)) > 0.999 Then
                                    wptpf = doku.ComponentDefinition.WorkPoints.AddAtCentroid(fw.EdgeLoops.Item(1))
                                    Exit For
                                End If
                            End If
                        End If

                    End If

                Next
                v = wpf.Point.VectorTo(wptpf.Point)
                Try
                    wpf.Delete()
                    wptpf.Delete()
                Catch ex2 As Exception

                End Try
                Return v
            Catch ex As Exception
                Try
                    wpf.Delete()
                    wptpf.Delete()
                Catch ex2 As Exception
                    MsgBox(ex.ToString())
                    Return Nothing
                End Try
            End Try
        Catch ex As Exception
            Try
                wpf.Delete()
                wptpf.Delete()
            Catch ex2 As Exception
                MsgBox(ex.ToString())
                Return Nothing
            End Try
        End Try
        Return Nothing
    End Function
    Function RemoveBorders(f As Face, wpt As WorkPoint, ef As ExtrudeFeature) As ExtrudeFeature


        Dim d, e, dmin, dp, npls As Double
        Dim ve, vc, vd As Vector
        Dim mpt As Point
        Dim ls As LineSegment
        Dim plfr, plfd As Plane

        Dim spl As PlanarSketch
        Dim skl As SketchLine
        Dim fr, fd As Face
        dmin = 9999999
        Try



            For Each ed As Edge In f.Edges
                e = CalculateEntryDistance(ed)
                If e < dmin Then
                    dmin = e
                End If



            Next

            If ef.SideFaces.Count > 1 Then

                fr = GetMajorFace(ef)
                fd = sideMinorFace
                plfd = sideMinorFace.Geometry
                plfr = sideMajorFace.Geometry
                npls = (plfd.Normal.AsVector).CrossProduct(plfr.Normal.AsVector).Length
                If npls > 0.5 Then
                    Dim sb As SurfaceBody = doku.ComponentDefinition.WorkSurfaces.Item(2).SurfaceBodies.Item(1)
                    For Each ed As Edge In f.Edges
                        d = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                        For Each fs As Face In sb.Faces
                            If fs.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                                For Each eds As Edge In fs.Edges
                                    e = eds.StartVertex.Point.DistanceTo(eds.StopVertex.Point)
                                    If Math.Abs(d - e) < 1 / 128 Then
                                        Dim vg, vs As Vector
                                        vg = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point)
                                        vs = eds.StartVertex.Point.VectorTo(eds.StopVertex.Point)
                                        dp = Math.Abs(vg.DotProduct(vs))
                                        If dp > 0.9 Then
                                            If eds.GetClosestPointTo(ed.StartVertex.Point).DistanceTo(ed.StartVertex.Point) < 25 / 10 Then
                                                Return MakeBandEntrance(ef, ed)
                                            End If

                                        End If
                                    End If
                                Next
                            End If
                        Next
                        e = CalculateEntryDistance(ed)
                        If e <= dmin Then
                            Return MakeBandEntrance(ef, ed)
                        End If
                    Next
                End If
            End If



            Try
                spl.Visible = False
            Catch ex As Exception

            End Try

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function
    Function MakeBandEntrance(ef As ExtrudeFeature, ed As Edge) As ExtrudeFeature
        Dim spl As PlanarSketch
        Dim skl As SketchLine
        Dim edt As Edge = ed
        Dim fr, fd As Face
        Dim mpt As Point
        Dim ls As LineSegment
        Dim plfr, plfd As Plane
        Dim ve, vc, vd As Vector
        Dim pt1, pt2, pt3, pt4 As Point
        Dim spt1, spt2, spt3 As SketchPoint3D
        Dim dmin As Double = 9999999
        Dim d, e As Double
        Try
            sk3D = doku.ComponentDefinition.Sketches3D.Add()
            d = CalculateClosestFace(GetMajorFace(ef))
            e = CalculateClosestFace(sideMinorFace)
            If d > e Then
                fd = sideMinorFace
                plfd = sideMinorFace.Geometry
                fr = sideMajorFace
                plfr = sideMajorFace.Geometry
            Else
                fd = sideMajorFace
                plfd = sideMajorFace.Geometry
                fr = sideMinorFace
                plfr = sideMinorFace.Geometry

            End If

            lamp.HighLighObject(edt)
            For Each vt As Vertex In fd.Vertices
                d = vt.Point.DistanceTo(edt.StartVertex.Point)
                e = vt.Point.DistanceTo(edt.StopVertex.Point)
                If (d < dmin) Or (e < dmin) Then
                    If e < d Then
                        dmin = e
                        pt4 = edt.StartVertex.Point
                    Else
                        dmin = d
                        pt4 = edt.StopVertex.Point
                    End If
                    pt1 = vt.Point
                End If
            Next
            spt1 = sk3D.SketchPoints3D.Add(pt1)
            vc = plfd.Normal.AsVector.CrossProduct(plfr.Normal.AsVector)
            pt2 = pt1
            pt2.TranslateBy(vc)




            spt2 = sk3D.SketchPoints3D.Add(pt2)
            pt3 = pt1
            pt3.TranslateBy(plfd.Normal.AsVector)
            spt3 = sk3D.SketchPoints3D.Add(pt3)
            sk3D.Visible = False
            Dim wpl As WorkPlane = doku.ComponentDefinition.WorkPlanes.AddByThreePoints(spt1, spt2, spt3)
            wpl.Visible = False

            d = wpl.Plane.DistanceTo(pt4)
            If d < 0 Then
                d = d * -1
            Else
                d = 1 / 128
            End If
            spl = doku.ComponentDefinition.Sketches.Add(wpl)
            'RemoveFaceExtend(fr)
            Return RemoveEntryMaterial(spl, edt, d)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CalculateEntryDistance(ed As Edge) As Double

        Dim ve, vc As Vector
        Dim mpt As Point
        Dim ls As LineSegment
        Dim e As Double
        ls = ed.Geometry
        mpt = ls.MidPoint
        ve = ls.Direction.AsVector
        vc = doku.ComponentDefinition.WorkPoints.Item(1).Point.VectorTo(mpt)
        e = (Math.Abs(vc.DotProduct(ve))) * vc.Length

        Return e
    End Function
    Function CalculateClosestFace(f As Face) As Double

        Dim vp, vc As Vector
        Dim mpt As Point
        Dim pl As Plane
        Dim wptfc As WorkPoint

        Dim e As Double
        pl = f.Geometry
        wptfc = doku.ComponentDefinition.WorkPoints.AddAtCentroid(f.EdgeLoops.Item(1))
        wptfc.Visible = False
        mpt = wptfc.Point
        vp = pl.Normal.AsVector
        vc = doku.ComponentDefinition.WorkPoints.Item(1).Point.VectorTo(mpt)
        e = Math.Abs(vc.DotProduct(vp)) / (vc.Length)

        Return e
    End Function
    Function RemovePatch(spl As PlanarSketch, skl As SketchLine, wpt As WorkPoint) As ExtrudeFeature
        Try

            Dim v As Vector2d = skl.Geometry.Direction.AsVector
            Dim spt As SketchPoint = spl.SketchPoints.Add(skl.Geometry.MidPoint)
            Dim m As Matrix2d = tg.CreateMatrix2d
            m.SetToRotation(Math.PI / 2, spt.Geometry)
            v.TransformBy(m)
            v.ScaleBy(skl.Length)
            Dim pt2d As Point2d = skl.StartSketchPoint.Geometry
            pt2d.TranslateBy(v)
            spl.SketchPoints.Add(pt2d)
            spl.SketchLines.AddAsThreePointCenteredRectangle(spt, skl.StartSketchPoint, pt2d)

            Dim pro As Profile

            pro = spl.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
            oExtrudeDef.SetToExtent(wpt)
            Dim oExtrude As ExtrudeFeature
            oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)
            Return oExtrude
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetWorkFace() As Face
        Try
            Dim maxArea1, maxArea2, maxArea3 As Double
            Dim sb As SurfaceBody
            Dim maxface1, maxface2, maxface3 As Face

            Dim b, b2 As Boolean
            Dim aMax, eMax As Double

            features = doku.ComponentDefinition.Features
            sb = doku.ComponentDefinition.SurfaceBodies.Item(1)
            maxface1 = sb.Faces.Item(1)
            maxface2 = maxface1
            maxArea2 = 0
            maxArea1 = maxArea2
            maxArea3 = maxArea2
            aMax = 0
            eMax = 0

            For Each f As Face In sb.Faces
                If f.Evaluator.Area > aMax Then
                    aMax = f.Evaluator.Area
                End If
            Next


            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    If f.Evaluator.Area > aMax / 2 Then
                        b = False
                        For Each fb As Face In bandFaces.SurfaceBodies.Item(1).Faces
                            If Math.Abs(fb.Evaluator.Area - f.Evaluator.Area) < 0.001 Then
                                b = True
                            End If
                        Next
                    Else
                        b = False
                    End If

                    If b Then
                        If b Then
                            If f.Evaluator.Area > maxArea3 Then
                                If f.Evaluator.Area > maxArea2 Then
                                    If f.Evaluator.Area > maxArea1 Then
                                        maxface3 = maxface2
                                        maxface2 = workFace
                                        workFace = f
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
                        End If
                    End If
                End If
            Next


            lamp.HighLighFace(workFace)
            Return workFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function GetMajorFace(ef As ExtrudeFeature) As Face

        Try
            Dim maxArea1, maxArea2, maxArea3 As Double

            Dim maxface1, maxface2, maxface3 As Face


            maxface1 = ef.SideFaces.Item(1)
            maxface2 = maxface1
            maxArea2 = 0
            maxArea1 = maxArea2
            maxArea3 = maxArea2





            For Each f As Face In ef.SideFaces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
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

                End If
            Next
            sideMinorFace = maxface2
            sideMajorFace = maxface1

            lamp.HighLighFace(sideMajorFace)
            lamp.HighLighFace(sideMinorFace)
            Return sideMajorFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try




    End Function
    Function GetWorkFaceComplex() As Face
        Try
            Dim maxArea1, maxArea2, maxArea3 As Double
            Dim sb As SurfaceBody
            Dim maxface1, maxface2, maxface3 As Face
            Dim fc As FaceCollection
            Dim b, b2 As Boolean
            Dim d1, d2, aMax, eMax As Double
            Dim v1, v2 As Vector

            maxface1 = sheetMetalFeatures.FoldFeatures.Item(sheetMetalFeatures.FoldFeatures.Count).Faces.Item(1)
            maxface2 = maxface1
            maxArea2 = 0
            maxArea1 = maxArea2
            maxArea3 = maxArea2
            aMax = 0
            eMax = 0
            features = doku.ComponentDefinition.Features
            sb = doku.ComponentDefinition.SurfaceBodies.Item(1)
            For Each f As Face In sb.Faces
                If f.Evaluator.Area > aMax Then
                    aMax = f.Evaluator.Area
                End If
            Next
            For Each f As Face In bandFaces.SurfaceBodies.Item(1).Faces
                For Each edb As Edge In f.Edges
                    If edb.StartVertex.Point.DistanceTo(edb.StopVertex.Point) > eMax Then
                        eMax = edb.StartVertex.Point.DistanceTo(edb.StopVertex.Point)
                    End If
                Next
            Next

            For Each f As Face In sb.Faces
                If f.Evaluator.Area > aMax / 2 Then
                    b = False
                    For Each fb As Face In bandFaces.SurfaceBodies.Item(1).Faces
                        If Math.Abs(fb.Evaluator.Area - f.Evaluator.Area) < 0.001 Then
                            b = False
                        End If
                    Next
                Else
                    b = False
                End If

                If b Then
                    b2 = False
                    While Not b2
                        For Each ed As Edge In f.Edges
                            If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) > eMax / 2 Then
                                d1 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                                v1 = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point).AsUnitVector.AsVector
                                For Each fb As Face In bandFaces.SurfaceBodies.Item(1).Faces
                                    While Not b2
                                        For Each ed2 As Edge In fb.Edges
                                            d2 = ed2.StartVertex.Point.DistanceTo(ed2.StopVertex.Point)
                                            v2 = ed2.StartVertex.Point.VectorTo(ed2.StopVertex.Point).AsUnitVector.AsVector
                                            If Math.Abs(d1 - d2) < 0.01 Then
                                                If Math.Abs(v1.DotProduct(v2)) > 0.9 Then
                                                    b2 = True
                                                    Exit For
                                                End If

                                            End If
                                        Next
                                        If Not b2 Then
                                            Exit While
                                        End If
                                    End While
                                    If b2 Then
                                        Exit For
                                    End If
                                Next

                            End If
                            If b2 Then
                                Exit For
                            End If

                        Next
                        If Not b2 Then
                            Exit While
                        End If
                    End While
                    If b Then
                        If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                            If f.Evaluator.Area > maxArea3 Then
                                If f.Evaluator.Area > maxArea2 Then
                                    If f.Evaluator.Area > maxArea1 Then
                                        maxface3 = maxface2
                                        maxface2 = workFace
                                        workFace = f
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
                        End If
                    End If


                End If


            Next


            lamp.HighLighFace(workFace)
            Return workFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function


    Function EmbossNumber(s As String) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim q As Integer
        Try
            q = nombrador.GetQNumberString(s)
            ef = ExtrudeNumber(SketchNumber(q))

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return ef
    End Function
    Function SketchNumber(q As Integer) As Profile
        Dim oProfile As Profile

        Dim a, b As Double
        Try
            Dim wpt As WorkPoint
            wpt = doku.ComponentDefinition.WorkPoints.AddAtCentroid(GetWorkFace().EdgeLoops.Item(1))
            wpt.Visible = False

            Dim spt As SketchPoint
            Dim oSketch As PlanarSketch
            Dim l As Line
            Dim edMax As Edge
            edMax = GetClosestEdge(workFace)
            oSketch = doku.ComponentDefinition.Sketches.AddWithOrientation(workFace, edMax, True, True, edMax.StartVertex,)
            spt = oSketch.AddByProjectingEntity(wpt)
            l = oSketch.AxisEntityGeometry

            Dim oTextBox As TextBox
            Dim oStyle As TextStyle
            Dim sText As String
            sText = CStr(q)

            oTextBox = oSketch.TextBoxes.AddFitted(spt.Geometry, sText)
            oStyle = oSketch.TextBoxes.Item(1).Style
            oTextBox.Delete()
            oStyle.FontSize = 0.81197733879089
            oStyle.Bold = True
            doku.Update2(True)
            Dim ptBox As Point2d
            ptBox = tg.CreatePoint2d(spt.Geometry.X - oStyle.FontSize / 2, -1 / 10)
            oTextBox = oSketch.TextBoxes.AddFitted(ptBox, sText, oStyle)
            Dim reflex As TextBox
            ptBox = tg.CreatePoint2d(spt.Geometry.X - oStyle.FontSize / 2, 1 / 10)
            reflex = oSketch.TextBoxes.AddFitted(ptBox, sText, oStyle)
            reflex.Rotation = Math.PI


            ' oTextBox.Rotation = Math.PI



            Dim oPaths As ObjectCollection
            oPaths = app.TransientObjects.CreateObjectCollection
            oPaths.Add(oTextBox)
            oPaths.Add(reflex)


            oProfile = oSketch.Profiles.AddForSolid(False, oPaths)


        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return oProfile
    End Function
    Function GetMajorEdge(f As Face) As Edge
        Dim e1, e2, e3 As Edge
        Dim maxe1, maxe2, maxe3 As Double
        maxe1 = 0
        maxe2 = 0
        maxe3 = 0
        e1 = f.EdgeLoops.Item(1).Edges.Item(1)
        e2 = e1
        e3 = e2
        For Each ed As Edge In f.EdgeLoops.Item(1).Edges
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
        lamp.HighLighObject(e3)
        lamp.HighLighObject(e2)
        lamp.HighLighObject(e1)
        bendEdge = e3
        minorEdge = e2
        majorEdge = e1
        Return e1
    End Function
    Function GetMinorEdge(f As Face) As Edge
        Dim e1, e2, e3 As Edge
        Dim min1, min2, min3 As Double
        min1 = 999999
        min2 = 99999
        min3 = 999999
        e1 = f.EdgeLoops.Item(1).Edges.Item(1)
        e2 = e1
        e3 = e2
        For Each ed As Edge In f.EdgeLoops.Item(1).Edges
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



            End If

        Next

        minorEdge = e1

        Return e1
    End Function
    Function GetClosestEdge(f As Face) As Edge
        Dim e1, e2, e3 As Edge
        Dim mine1, mine2, mine3, d, e As Double
        Dim ve, vc As Vector
        Dim pt1 As Point
        Dim ls As LineSegment
        mine1 = 99999
        mine2 = 9999999
        mine3 = 999999
        e1 = f.EdgeLoops.Item(1).Edges.Item(1)
        e2 = e1
        e3 = e2
        pt1 = f.GetClosestPointTo(wpConverge.Point)
        e = pt1.DistanceTo(wpConverge.Point)
        ' sk3D.SketchPoints3D.Add(pt1)
        For Each ed As Edge In f.EdgeLoops.Item(1).Edges
            d = ed.GetClosestPointTo(wpConverge.Point).DistanceTo(wpConverge.Point)
            If Math.Abs(d - e) < 1 / 10 Then
                ve = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point)
                ls = tg.CreateLineSegment(ed.StartVertex.Point, ed.StopVertex.Point)
                vc = ls.MidPoint.VectorTo(wpConverge.Point)
                d = ls.MidPoint.DistanceTo(wpConverge.Point) / (ve.CrossProduct(vc).Length)
                If d < mine2 Then
                    If d < mine1 Then
                        mine3 = mine2
                        e3 = e2
                        mine2 = mine1
                        e2 = e1
                        mine1 = d
                        e1 = ed
                    Else
                        mine3 = mine2
                        e3 = e2
                        mine2 = d
                        e2 = ed
                    End If
                Else
                    mine3 = d
                    e3 = ed
                End If

            End If


        Next
        lamp.HighLighObject(e1)

        closestEdge = e1
        Return e1
    End Function
    Function GetClosestFace(wpti As WorkPoint) As Face
        Dim f1, f2, f3 As Face
        Dim pl As Plane
        Dim mine1, mine2, mine3, d As Double

        mine1 = 99999
        mine2 = 9999999
        mine3 = 999999
        f1 = compDef.SurfaceBodies(1).Faces(1)
        f2 = f1
        f3 = f2
        For Each f As Face In compDef.SurfaceBodies(1).Faces
            If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                'pl = f.Geometry
                d = f.GetClosestPointTo(wpti.Point).DistanceTo(wpti.Point)
                If d < mine2 Then
                    If d < mine1 Then
                        mine3 = mine2
                        f3 = f2
                        mine2 = mine1
                        f2 = f1
                        mine1 = d
                        f1 = f
                    Else
                        mine3 = mine2
                        f3 = f2
                        mine2 = d
                        f2 = f
                    End If
                Else
                    mine3 = d
                    f3 = f
                End If
            End If




        Next
        sk3D.SketchPoints3D.Add(f1.PointOnFace)
        lamp.HighLighFace(f1)

        closestFace = f1
        Return f1
    End Function
    Function ExtrudeNumber(pro As Profile) As ExtrudeFeature
        Dim oExtrudeDef As ExtrudeDefinition
        oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
        oExtrudeDef.SetThroughAllExtent(PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        'oExtrudeDef.SetDistanceExtent(0.12, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude As ExtrudeFeature
        oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)



        Return oExtrude
    End Function
    Function RemoveFakeMaterial(f As Face, wpt As WorkPoint) As ExtrudeFeature
        Dim spl As PlanarSketch
        spl = doku.ComponentDefinition.Sketches.Add(f)
        Dim skl As SketchLine
        For Each ed As Edge In f.Edges
            skl = spl.AddByProjectingEntity(ed)
        Next


        Dim pro As Profile

        pro = spl.Profiles.AddForSolid
        Dim oExtrudeDef As ExtrudeDefinition
        oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
        oExtrudeDef.SetToExtent(wpt)
        Dim oExtrude As ExtrudeFeature
        oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)



        Return oExtrude
    End Function
    Function RemoveFaceExtend(f As Face) As ExtrudeFeature
        Dim spl As PlanarSketch
        spl = doku.ComponentDefinition.Sketches.Add(f)
        Dim skl As SketchLine
        For Each ed As Edge In f.Edges
            skl = spl.AddByProjectingEntity(ed)
        Next
        Dim pro As Profile
        pro = spl.Profiles.AddForSolid
        Dim oExtrudeDef As ExtrudeDefinition
        oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
        oExtrudeDef.SetDistanceExtent(25 / 10, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude As ExtrudeFeature
        oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)

        Return oExtrude
    End Function
    Function RemoveEntryMaterial(spl As PlanarSketch, ed As Edge, d As Double) As ExtrudeFeature
        Try
            Dim skl As SketchLine
            lamp.HighLighObject(ed)
            skl = spl.AddByProjectingEntity(ed)
            skl.Construction = True
            Dim v As Vector2d = skl.Geometry.Direction.AsVector
            Dim pt2 As Point2d = skl.EndSketchPoint.Geometry
            Dim v2 As Vector2d = v
            v2.ScaleBy(25 / 10)
            pt2.TranslateBy(v2)
            Dim spt2 As SketchPoint = spl.SketchPoints.Add(pt2)
            Dim spt As SketchPoint = spl.SketchPoints.Add(skl.Geometry.MidPoint)
            Dim m As Matrix2d = tg.CreateMatrix2d
            m.SetToRotation(Math.PI / 2, spt2.Geometry)
            v.TransformBy(m)
            v.ScaleBy(skl.Length * 1 / 2)
            Dim pt3 As Point2d = pt2
            pt3.TranslateBy(v)
            spl.SketchPoints.Add(pt3)
            spl.SketchLines.AddAsThreePointCenteredRectangle(spt, spt2, pt3)

            Dim pro As Profile

            pro = spl.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
            oExtrudeDef.SetDistanceExtent(25 / 10, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            oExtrudeDef.SetDistanceExtentTwo(d)
            Dim oExtrude As ExtrudeFeature
            oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)



            Return oExtrude
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function MakeHole() As RevolveFeature
        Dim rf As RevolveFeature

        Try
            rf = RevolveHole(DrawHole())
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return rf
    End Function
    Function MakeRing() As RevolveFeature
        Dim rf As RevolveFeature

        Try
            rf = RevolveRing(DrawRing())
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return rf
    End Function
    Function DrawHole() As Profile
        Dim pro As Profile
        Dim ps As PlanarSketch
        Dim spt As SketchPoint
        Try
            ps = doku.ComponentDefinition.Sketches.Add(doku.ComponentDefinition.WorkPlanes.Item(2))
            spt = ps.SketchPoints.Add(tg.CreatePoint2d(Tr, 0))
            ps.SketchCircles.AddByCenterRadius(spt, Cr / 2 - 3 / 10)
            pro = ps.Profiles.AddForSolid
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return pro
    End Function
    Function DrawRing() As Profile
        Dim pro As Profile
        Dim ps As PlanarSketch
        Dim spt As SketchPoint
        Try
            ps = doku.ComponentDefinition.Sketches.Add(doku.ComponentDefinition.WorkPlanes.Item(2))
            spt = ps.SketchPoints.Add(tg.CreatePoint2d(Tr, 0))
            ps.SketchCircles.AddByCenterRadius(spt, Cr / 4 - 3 / 10)
            pro = ps.Profiles.AddForSolid
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return pro
    End Function
    Function RevolveHole(pro As Profile) As RevolveFeature
        Dim rf As RevolveFeature
        Try
            rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, doku.ComponentDefinition.WorkAxes.Item(3), PartFeatureOperationEnum.kCutOperation)

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return rf
    End Function
    Function RevolveRing(pro As Profile) As RevolveFeature
        Dim rf As RevolveFeature
        Try
            rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, doku.ComponentDefinition.WorkAxes.Item(3), PartFeatureOperationEnum.kJoinOperation)

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return rf
    End Function
    Function GetConvergePoint(q As Integer) As Point
        Dim wpt As WorkPoint
        Dim pt As Point = tg.CreatePoint(Math.Cos(2 * Math.PI * DP.p * q / DP.q) * Tr, Math.Sin(2 * Math.PI * DP.p * q / DP.q) * Tr, 0)
        wpt = doku.ComponentDefinition.WorkPoints.AddFixed(pt)
        wpt.Visible = False
        wpConverge = wpt
        Return pt
    End Function
    Function CombineBodies() As CombineFeature
        Dim cf As CombineFeature
        If doku.ComponentDefinition.SurfaceBodies.Count > 1 Then
            surfaceBodies.Clear()

            For index = 2 To doku.ComponentDefinition.SurfaceBodies.Count
                surfaceBodies.Add(doku.ComponentDefinition.SurfaceBodies.Item(index))
            Next
            Try
                cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)

            Catch ex As Exception


            End Try
        End If

        ' cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)

        Return cf
    End Function
    Function CombineBodiesDuo() As CombineFeature
        Dim cf As CombineFeature
        Dim imax, limit, k, j, l As Integer
        l = doku.ComponentDefinition.SurfaceBodies.Count
        limit = 0
        While doku.ComponentDefinition.SurfaceBodies.Count > 1 And limit < Math.Pow(l, 2)
            imax = GetMaxBody()
            Math.DivRem(limit, doku.ComponentDefinition.SurfaceBodies.Count, j)
            If imax = j + 1 Then
                k = 1
            Else
                k = 0
            End If
            ' For j = imax + k To doku.ComponentDefinition.SurfaceBodies.Count
            surfaceBodies.Clear()
            ' If Not j = imax Then
            Try
                surfaceBodies.Add(doku.ComponentDefinition.SurfaceBodies.Item(j + k + 1))
                Try
                    cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(imax), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)
                    limit = limit + 1

                Catch ex As Exception
                    limit = limit + 1
                End Try
            Catch ex As Exception
                limit = limit + 1
            End Try


            '  End If
            '  Next

        End While

        Return cf
    End Function
    Function GetMaxBody() As Integer
        Dim vMax As Double = 0
        Dim imax As Integer
        Try
            For j = 1 To doku.ComponentDefinition.SurfaceBodies.Count
                Try
                    If (doku.ComponentDefinition.SurfaceBodies.Item(j)).Volume(0.01) > vMax Then
                        vMax = (doku.ComponentDefinition.SurfaceBodies.Item(j)).Volume(0.01)
                        imax = j
                    End If
                Catch ex As Exception

                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return imax
    End Function

    Function MakeSingleWedge(s As String) As PartDocument
        Dim p As PartDocument
        Dim q As Integer
        Dim derivedDefinition As DerivedPartDefinition
        Dim newComponent As DerivedPartComponent
        Try
            p = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
            converter.SetUnitsToMetric(p)
            derivedDefinition = p.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsWorkSurface
            newComponent = p.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)
            p.Update2(True)
            doku = DocUpdate(p)
            q = nombrador.GetQNumberString(s)
            convergePoint = GetConvergePoint(q + 1)
            LoftFaces(compDef.WorkSurfaces.Item(1))

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return p
    End Function

    Function TryLoft(f1 As Face, f2 As Face) As LoftFeature

        Try
            sections.Clear()


            sections.Add(GetSpikeProfile(f1))
            sections.Add(GetSpikeProfile(f2))
            Return LoftFilling()

        Catch ex As Exception

            Return Nothing
        End Try

    End Function

    Function LoftFaces(ws As WorkSurface) As Integer
        Dim vc, vfc As Vector
        Dim ptc, ptf As Point
        Dim pl As Plane
        Dim min2, min1 As Double
        Dim wp As WorkPoint
        Dim fmin1, fmin2 As Face
        Try
            ptc = doku.ComponentDefinition.WorkPoints.Item(1).Point
            caras.Clear()
            min1 = ws.SurfaceBodies.Item(1).Faces.Item(1).Evaluator.Area
            min2 = min1
            fmin1 = ws.SurfaceBodies.Item(1).Faces.Item(1)
            fmin2 = fmin1
            For Each sb As SurfaceBody In ws.SurfaceBodies
                For Each fc As Face In sb.Faces
                    If fc.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then

                        min2 = fc.Evaluator.Area
                        If min2 < min1 Then
                            min1 = min2
                            fmin2 = fmin1
                            fmin1 = fc
                        Else
                            fmin2 = fc
                        End If


                    End If
                Next
            Next
            ' lamp.HighLighFace(fmin1)
            cylinderFace = fmin1
            workFaces = fmin1.TangentiallyConnectedFaces
            For Each f As Face In fmin1.TangentiallyConnectedFaces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    'lamp.HighLighFace(f)
                    MakeSpike(f)
                End If
            Next

            Return caras.Count
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function GetMinimosEdges(f As Face) As Double
        Dim e1, e2, e3, e4 As Edge
        Dim mine1, mine2, mine3, mine4 As Double
        mine1 = f.Edges.Item(1).StartVertex.Point.DistanceTo(f.Edges.Item(1).StopVertex.Point)
        mine2 = mine1
        mine3 = mine2
        mine4 = mine3
        e1 = f.Edges.Item(1)
        e2 = e1
        e3 = e2
        e4 = e3
        For Each ed As Edge In f.Edges
            If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) < mine3 Then
                If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) < mine2 Then

                    If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) < mine1 Then
                        mine4 = mine3
                        e4 = e3
                        mine3 = mine2
                        e3 = e2
                        mine2 = mine1
                        e2 = e1
                        mine1 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                        e1 = ed
                    Else
                        mine4 = mine3
                        e4 = e3
                        mine3 = mine2
                        e3 = e2
                        mine2 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                        e2 = ed
                    End If

                Else
                    mine4 = mine3
                    e4 = e3
                    mine3 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                    e3 = ed

                End If
            Else
                mine4 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                e4 = ed
            End If


        Next
        'lamp.HighLighObject(e4)
        'lamp.HighLighObject(e2)
        lamp.HighLighObject(e1)

        Return (mine1 + mine2)

    End Function

    Function MakeSpike(fc As Face) As LoftFeature
        sections.Clear()
        Dim pr As Profile = GetSpikeProfile(fc)
        Dim ps As PlanarSketch
        Dim spt2d As SketchPoint
        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        Dim ed As Edge = GetClosestEdge(fc)
        'GetRails(fc)
        sections.Add(pr)
        Dim pt As Point


        Dim v As Vector
        v = GetRealNormal(fc, doku.ComponentDefinition.WorkSurfaces.Item(1))
        v.ScaleBy(-8 * fc.Evaluator.Area)
        ' ps = doku.ComponentDefinition.Sketches.Add(fc)
        '   spt2d = ps.AddByProjectingEntity(wpConverge)
        '   ps.Visible = False
        pt = GetOppositeVertex(ed, fc).Point
        'pt.TranslateBy(v)
        Dim wptlf As WorkPoint = doku.ComponentDefinition.WorkPoints.AddFixed(pt)
        wptlf.Visible = False
        sections.Add(wptlf)
        Return LoftSingleSpike()
    End Function

    Function MakeCone(skpt As SketchPoint3D, wpt As WorkPoint) As LoftFeature
        Dim skl As SketchLine3D
        Dim pr As Profile
        skl = coneAxes.Item(coneAxes.Count)
        sections.Clear()
        If skl.Length < 10 / 10 Then
            sections.Add(wpt)
        Else
            pr = DrawCircleOnLine(coneLine, 0.3 / 10)
            sections.Add(pr)
        End If
        pr = GetConeProfile(skpt)
        sections.Add(pr)
        Return LoftSingleCone(wpt)
    End Function

    Function GetOppositeVertex(edi As Edge, fi As Face) As Vertex
        Dim ver1, ver2 As Vertex
        Dim ws As WorkSurface = compDef.WorkSurfaces.Item(1)
        Dim sb As SurfaceBody = ws.SurfaceBodies.Item(1)
        Dim vnpl, vnfi, v As Vector
        Dim d, e, eMin, dMin1, dMin2, dis As Double
        Dim pt1, pt2, pt3 As Point
        Dim ls As LineSegment
        Dim ed2 As Edge
        vnfi = GetRealNormal(fi, ws)
        vnfi.Normalize()
        pt1 = fi.PointOnFace

        'sk3D.SketchPoints3D.Add(pt1)
        dMin1 = 9999
        dMin2 = 9999
        eMin = 99999
        ver1 = fi.Vertices.Item(1)
        ls = tg.CreateLineSegment(edi.StartVertex.Point, edi.StopVertex.Point)
        pt3 = ls.MidPoint
        sk3D.SketchPoints3D.Add(pt3)
        Try
            For Each f As Face In fi.TangentiallyConnectedFaces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    If Not (fi.Equals(f)) Then
                        vnpl = GetRealNormal(f, ws)
                        vnpl.Normalize()
                        d = vnpl.DotProduct(vnfi)
                        If d < -0.5 Then
                            ed2 = GetClosestEdge(f)
                            pt2 = ed2.GetClosestPointTo(pt3)
                            sk3D.SketchPoints3D.Add(pt2)
                            e = pt3.DistanceTo(pt2)
                            If e < 18 / 10 Then
                                If e < eMin Then
                                    eMin = e
                                    pt1 = ed2.GetClosestPointTo(wpConverge.Point)
                                    sk3D.SketchPoints3D.Add(pt1)
                                    For Each vr As Vertex In f.Vertices
                                        dis = pt1.DistanceTo(vr.Point)
                                        If dis < dMin2 Then
                                            If dis < dMin1 Then
                                                dMin2 = dMin1
                                                dMin1 = dis
                                                ver2 = ver1
                                                ver1 = vr
                                                'lamp.HighLighObject(vr)
                                            Else
                                                dMin2 = dis
                                                ver2 = vr

                                            End If

                                        End If
                                    Next

                                End If


                            End If
                        End If
                    End If
                End If
            Next
            sk3D.Visible = False
            lamp.HighLighObject(ver1)
            Return ver1
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeCones(wpt1 As WorkPoint, wpt2 As WorkPoint, wptc As WorkPoint) As LoftFeature
        Dim lf As LoftFeature
        lf = MakeCone(wptc, wpt1)
        lf = MakeCone(wptc, wpt2)
        Return lf
    End Function
    Function MakeSupports(p As RodMaker, skt As Sketch3D) As LoftFeature
        Try
            wp1 = p.wp1
            wp2 = p.wp2
            curvesSketch = skt
            palito = p
            ovalLine = palito.ovalLine

            Dim skl As SketchLine3D
            Dim lf As LoftFeature
            If compDef.Features.LoftFeatures.Count > 0 Then
                lf = compDef.Features.LoftFeatures.Item(compDef.Features.LoftFeatures.Count)
            End If

            Dim skpt As SketchPoint3D = OptimalAlignedPoint(wp1)


            If Not IsFeatureSimilar(wp1, skpt) Then
                lf = MakeCone(skpt, wp1)
                If IsConeFree(lf, skpt) Then
                    If monitor.IsFeatureHealthy(lf) Then
                        skpt = OptimalAlignedPoint(wp2)

                        If Not IsFeatureSimilar(wp2, skpt) Then
                            lf = MakeCone(skpt, wp2)
                            If IsConeFree(lf, skpt) Then
                                Return lf
                            Else
                                Return CorrectCone(lf, wp2)
                            End If

                        Else
                            Return lf
                        End If
                    End If
                Else
                    lf = CorrectCone(lf, wp1)

                    If monitor.IsFeatureHealthy(lf) Then
                        skpt = OptimalAlignedPoint(wp2)

                        If Not IsFeatureSimilar(wp2, skpt) Then
                            lf = MakeCone(skpt, wp2)
                            If IsConeFree(lf, skpt) Then
                                Return lf
                            Else
                                Return CorrectCone(lf, wp2)
                            End If

                        Else
                            Return lf
                        End If
                    End If
                End If

            Else
                skpt = OptimalAlignedPoint(wp2)
                endPoints.Add(skpt)
                If Not IsFeatureSimilar(wp2, skpt) Then
                    lf = MakeCone(skpt, wp2)
                    If IsConeFree(lf, skpt) Then
                        Return lf
                    Else
                        Return CorrectCone(lf, wp2)
                    End If
                Else
                    Return lf
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeSingleSupport(p As RodMaker, skt As Sketch3D) As LoftFeature
        Try
            wp1 = p.wp1
            wp2 = p.wp2
            curvesSketch = skt
            palito = p
            ovalLine = palito.ovalLine


            Dim lf As LoftFeature = Nothing
            If compDef.Features.LoftFeatures.Count > 0 Then
                lf = compDef.Features.LoftFeatures.Item(compDef.Features.LoftFeatures.Count)
            End If

            Dim skpt As SketchPoint3D = OptimalAlignedPoint(wp2)


            If Not IsFeatureSimilar(wp2, skpt) Then
                lf = MakeLoftCone(palito.wp2Face, skpt)
#If Entry_Type = "Unilateral" Then
                If monitor.IsFeatureHealthy(lf) Then
                    Return lf
                End If
#Else
                If IsConeFree(lf, skpt) Then
                    If monitor.IsFeatureHealthy(lf) Then
                        Return lf
                    End If
                Else
                    Return CorrectLoftCone(lf, palito.wp2Face)
                End If
#End If

            End If
            Return lf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CorrectCone(lf As LoftFeature, wpt As WorkPoint) As LoftFeature
        Dim skl As SketchLine3D
        Dim skpt As SketchPoint3D

        For i = 1 To 4
            For Each skpti As SketchPoint3D In interPoints
                lf.Delete(False, True, True)
                skl = MakeConeFree(coneAxes.Item(coneAxes.Count), skpti)
                sk3D.Visible = False
                coneLine = skl
                coneAxes.Add(skl)
                skpt = skl.EndSketchPoint
                lf = MakeCone(skl.EndSketchPoint, wpt)
                If IsConeFree(lf, skpt) Then
                    CorrectCone = lf
                    Exit For
                End If
            Next
            If IsConeFree(lf, skpt) Then
                CorrectCone = lf
                Exit For
            End If
        Next


        Return CorrectCone
    End Function
    Function CorrectLoftCone(lf As LoftFeature, fi As Face) As LoftFeature
        Dim skl As SketchLine3D
        Dim skpt As SketchPoint3D = coneLine.EndSketchPoint

        For i = 1 To 4
            For Each skpti As SketchPoint3D In interPoints
                lf.Delete(False, True, True)
                skl = MakeConeFree(coneAxes.Item(coneAxes.Count), skpti)
                sk3D.Visible = False
                coneLine = skl
                coneAxes.Add(skl)
                skpt = skl.EndSketchPoint

                lf = MakeLoftCone(fi, skl.EndSketchPoint)
                If IsConeFree(lf, skpt) Then
                    CorrectLoftCone = lf
                    Exit For
                End If

            Next
            If IsConeFree(lf, skpt) Then
                CorrectLoftCone = lf
                Exit For
            End If
        Next


        Return CorrectLoftCone
    End Function
    Function IsConeFree(lf As LoftFeature, skpti As SketchPoint3D) As Boolean
        Dim ic As IntersectionCurve
        Dim skpt As SketchPoint3D
        Dim ssl As SketchSpline
        Dim ssl3D As SketchSpline3D
        Dim d, dMin, dMax As Double
        dMin = 9999
        sk3D = compDef.Sketches3D.Add
        dMax = 0

        For Each sb As SurfaceBody In tangentSurfaces.SurfaceBodies
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    For Each fc As Face In lf.SideFaces
                        Try
                            ic = sk3D.IntersectionCurves.Add(f, fc)
                            corte = ic
                            corteFace = f
                            interPoints.Clear()

                            For Each se As SketchEntity3D In corte.SketchEntities
                                se.Construction = True
                                If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                                    skpt = se
                                    skpt = sk3D.SketchPoints3D.Add(skpt.Geometry)
                                    interPoints.Add(skpt)
                                ElseIf se.Type = ObjectTypeEnum.kSketchSplineObject Then
                                    ssl = se
                                    If ssl.Length > dMax Then
                                        dMax = ssl.Length
                                    End If
                                ElseIf se.Type = ObjectTypeEnum.kSketchSpline3DObject Then
                                    ssl3D = se
                                    If ssl3D.Length > dMax Then
                                        dMax = ssl3D.Length
                                    End If

                                End If


                            Next

                            If dMax > 3 / 10 Then
                                ic.BreakLink()
                                Return False
                            End If

                        Catch ex As Exception

                        End Try
                    Next
                End If

            Next
        Next
        endPoints.Add(skpti)
        If skpti.Geometry.Z > 0 Then
            highPoints.Add(skpti.Geometry)
        Else
            lowPoints.Add(skpti.Geometry)
        End If
        Return True
    End Function
    Function IsFeatureSimilar(wpt As WorkPoint, skpt As SketchPoint3D) As Boolean
        Dim v1, v2 As Vector
        Dim skpt2 As SketchPoint3D
        Dim sklNew, sklOld As SketchLine3D
        Dim d As Double
        Try
            sklNew = coneAxes.Item(coneAxes.Count)
            v2 = sklNew.StartSketchPoint.Geometry.VectorTo(sklNew.EndSketchPoint.Geometry)
            If coneAxes.Count > 1 Then
                For i = 1 To coneAxes.Count - 1
                    sklOld = coneAxes.Item(i)
                    v1 = sklOld.StartSketchPoint.Geometry.VectorTo(sklOld.EndSketchPoint.Geometry)

                    If Math.Abs(sklNew.Length - sklOld.Length) < 1 / 10 Then
                        v1.Normalize()
                        v2.Normalize()
                        d = v1.DotProduct(v2)
                        If d > 0.96 Then
                            IsFeatureSimilar = True
                            Return True
                        End If

                    End If
                Next
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function OptimalPoint(wpti As WorkPoint) As SketchPoint3D
        Dim e, d, dMin As Double
        Dim skeq As SketchEquationCurve3D
        Dim wptt As WorkPoint
        Dim ptOpt As Point
        Try
            sk3D = compDef.Sketches3D.Add
            dMin = 9999999
            Dim pt As Point = wpti.Point
            d = pt.DistanceTo(wptHigh.Point)
            e = pt.DistanceTo(wptLow.Point)
            If d < e Then
                skeq = curvesSketch.SketchEquationCurves3D.Item(2)
                ptOpt = wptHigh.Point
                For Each pto As Point In highPoints
                    d = pto.DistanceTo(pt)
                    If d < dMin Then
                        dMin = d
                        ptOpt = pto
                    End If

                Next

                lastSeq = skeq
                OptimalPoint = SketchOptimalPoint(wpti, ptOpt, skeq, palito.wp2Face).EndSketchPoint
                '  highPoints.Add(OptimalPoint.Geometry)
            Else
                skeq = curvesSketch.SketchEquationCurves3D.Item(3)
                ptOpt = wptLow.Point
                For Each pto As Point In lowPoints
                    d = pto.DistanceTo(pt)
                    If d < dMin Then
                        dMin = d
                        ptOpt = pto
                    End If

                Next
                lastSeq = skeq
                OptimalPoint = SketchOptimalPoint(wpti, ptOpt, skeq, palito.wp2Face).EndSketchPoint
                '  lowPoints.Add(OptimalPoint.Geometry)
            End If


            sk3D.Visible = False
            Return OptimalPoint
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function OptimalAlignedPoint(wpti As WorkPoint) As SketchPoint3D
        Dim e, d, dMin As Double
        Dim skeq As SketchEquationCurve3D
        Dim wptt As WorkPoint
        Dim ls As LineSegment
        Dim vp, vHigh, vLow As Vector

        Dim ptOpt As Point
        Try
            sk3D = compDef.Sketches3D.Add
            Dim skpt As SketchPoint3D = sk3D.SketchPoints3D.Add(ovalLine.StartSketchPoint.Geometry)
            Dim wpl As WorkPlane = compDef.WorkPlanes.AddByThreePoints(wp1, wp2, skpt)
            wpl.Visible = False
            lamp.LookAtPlane(wpl)
            dMin = 9999999
            Dim pt As Point = wpti.Point
            vp = wp2.Point.VectorTo(wp1.Point)
            vHigh = wp2.Point.VectorTo(wptHigh.Point)
            vLow = wp2.Point.VectorTo(wptLow.Point)
            Dim rad As Integer = 2
            d = pt.DistanceTo(wptHigh.Point) * Math.Pow(vp.CrossProduct(vHigh).Length, 1 / rad)
            e = pt.DistanceTo(wptLow.Point) * Math.Pow(vp.CrossProduct(vLow).Length, 1 / rad)
            If d < e Then
                skeq = curvesSketch.SketchEquationCurves3D.Item(2)

                Try
                    wptt = compDef.WorkPoints.AddFixed(GetClosestIntersectionPoint(skeq, wpl, skpt))
                    wptt.Visible = False
                    ptOpt = wptt.Point
                    ls = tg.CreateLineSegment(ptOpt, wptHigh.Point)
                    ptOpt = ls.MidPoint
                Catch ex As Exception
                    Try
                        wptt = compDef.WorkPoints.AddFixed(GetClosestPointCurve(skeq, wptHigh, wpti))
                        wptt.Visible = False
                        ptOpt = wptt.Point
                        '  ls = tg.CreateLineSegment(ptOpt, wptHigh.Point)
                        '  ptOpt = ls.MidPoint
                    Catch ex2 As Exception
                        ptOpt = wptHigh.Point
                        For Each pto As Point In highPoints
                            d = pto.DistanceTo(pt)
                            If d < dMin Then
                                dMin = d
                                ptOpt = pto
                            End If

                        Next
                    End Try

                End Try


                '  highPoints.Add(OptimalPoint.Geometry)
            Else
                skeq = curvesSketch.SketchEquationCurves3D.Item(3)
                Try
                    wptt = compDef.WorkPoints.AddFixed(GetClosestIntersectionPoint(skeq, wpl, skpt))
                    wptt.Visible = False
                    ptOpt = wptt.Point
                    ls = tg.CreateLineSegment(ptOpt, wptLow.Point)
                    ptOpt = ls.MidPoint
                Catch ex As Exception
                    Try
                        wptt = compDef.WorkPoints.AddFixed(GetClosestPointCurve(skeq, wptLow, wpti))
                        wptt.Visible = False
                        ptOpt = wptt.Point
                        'ls = tg.CreateLineSegment(ptOpt, wptLow.Point)
                        ' ptOpt = ls.MidPoint
                    Catch ex2 As Exception
                        ptOpt = wptLow.Point
                        For Each pto As Point In lowPoints
                            d = pto.DistanceTo(pt)
                            If d < dMin Then
                                dMin = d
                                ptOpt = pto
                            End If

                        Next
                    End Try

                End Try

                '  lowPoints.Add(OptimalPoint.Geometry)
            End If
            lastSeq = skeq
            OptimalAlignedPoint = SketchOptimalPoint(wpti, ptOpt, skeq, palito.wp2Face).EndSketchPoint

            sk3D.Visible = False
            Return OptimalAlignedPoint
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetClosestIntersectionPoint(seq As SketchEquationCurve3D, wpl As WorkPlane, skpt As SketchPoint3D) As Point
        Dim pl As Plane = wpl.Plane
        Dim ptMin As Point = seq.StartSketchPoint.Geometry
        Dim d, dMin As Double
        dMin = 9999999
        For Each pt As Point In pl.IntersectWithCurve(seq.Geometry)
            d = pt.DistanceTo(skpt.Geometry)
            If d < dMin Then
                dMin = d
                ptMin = pt
            End If
        Next
        Return ptMin
    End Function
    Function GetClosestPointCurve(seq As SketchEquationCurve3D, wpti As WorkPoint, wptiOval As WorkPoint) As Point

        Dim skl As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(wptiOval, wpti.Point, False)
        Dim gc As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddCoincident(skl.EndPoint, seq)
        Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(skl)
        Dim d As Double = dc.Parameter._Value
        For index = 1 To 16
            adjuster.AdjustDimConstrain3DSmothly(dc, 31 / 32 * dc.Parameter._Value)
            If dc.Parameter._Value = d Then
                Exit For
            End If

        Next
        Dim ptMin As Point = skl.EndSketchPoint.Geometry
        dc.Delete()
        skl.Delete()

        Return ptMin
    End Function
    Function SketchOptimalPoint(wpti As WorkPoint, ptt As Point, skeq As SketchEquationCurve3D, fi As Face) As SketchLine3D
        Dim v1 As Vector
        Try
            Dim v As Vector = palito.lref.StartSketchPoint.Geometry.VectorTo(palito.lref.EndSketchPoint.Geometry)
            Dim l As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(wpti, ptt, False)

            Dim gc As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, skeq)
            Dim pt As Point = wpti.Point
            pt.TranslateBy(v)
            Dim lr As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(l.StartPoint, pt, False)
            Dim gcgr As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddGround(lr)
            Dim pl As Plane = fi.Geometry
            Dim v2 As Vector = pl.Normal.AsVector
            pt = wpti.Point
            pt.TranslateBy(v2)
            Dim ln As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(l.StartPoint, pt, False)
            sk3D.GeometricConstraints3D.AddGround(ln)
            Dim a, c As Double
            Dim acn As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, ln)
            FirstConeLineAdjust(ln, l, acn)
            acn = AdjustNormal(ln, l, acn, Math.Cos(Math.PI / 4) * 2)


            acn.Driven = True
            Dim vc As Vector = compDef.WorkPoints(1).Point.VectorTo(l.EndSketchPoint.Geometry)
            Dim vz As Vector = tg.CreateVector(0, 0, l.EndSketchPoint.Geometry.Z)
            Dim vr As Vector = vc.CrossProduct(vz)
            vr.Normalize()
            Dim vl As Vector = l.Geometry.Direction.AsVector
            vl.Normalize()
            pt = l.EndSketchPoint.Geometry
            pt.TranslateBy(vz)
            Dim lz As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(l.EndPoint, pt, False)
            sk3D.GeometricConstraints3D.AddParallelToZAxis(lz)
            lz.Construction = True
            Dim acz As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(lz, l)
            acz = AdjustZetaAngle(l, ln, acn, acz, 1)
            Dim lc As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(l.EndPoint, compDef.WorkPoints.Item(1), False)
            lc.Construction = True

            Dim lrd As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(l.EndPoint, l.StartSketchPoint.Geometry, False)
            lrd.Construction = True
            Dim ac As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(lc, lrd)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Delete()
            sk3D.GeometricConstraints3D.AddPerpendicular(lc, lrd)
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(lz, lrd)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Delete()
            sk3D.GeometricConstraints3D.AddPerpendicular(lz, lrd)
            Dim acr As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, lrd)
            acr = AdjustTangentAngle(l, ln, acn, acr, acz, 1)
            Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(l)
            dc = AdjustLineLength(lz, l, ln, acr, acn, acz, dc)

            angleConeLineTangent = acr
            angleConeNormalLine = acn
            coneLine = l
            coneAxes.Add(l)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function FirstConeLineAdjust(ln As SketchLine3D, l As SketchLine3D, acn As DimensionConstraint3D) As Integer
        Dim b As Double = 25 / 10

        For i = 1 To 32
            If l.Length > b / 2 And l.Length < 2 * b Then
                adjuster.AdjustDimConstrain3DSmothly(acn, acn.Parameter._Value * 15 / 16)
            Else
                Exit For
            End If
        Next
        Return 0
    End Function
    Function AdjustLineLength(lz As SketchLine3D, l As SketchLine3D, ln As SketchLine3D, acr As DimensionConstraint3D,
                  acn As DimensionConstraint3D, acz As DimensionConstraint3D, dc As DimensionConstraint3D) As DimensionConstraint3D
        Dim j, k As Integer
        j = 0
        k = 0
        For i = 1 To 16
            If dc.Parameter._Value < 10 / 10 Then
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 17 / 16)
                dc.Driven = True
                AdjustConeLineSmoothly(l, ln, acr, acn, acz, dc, k, j, i)


            ElseIf dc.Parameter._Value > 50 / 10 Then
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 15 / 16)
                dc.Driven = True
                AdjustConeLineSmoothly(l, ln, acr, acn, acz, dc, k, j, i)
            Else
                Exit For
            End If


        Next
        dc.Driven = True

        Return dc
    End Function
    Function AdjustConeLineSmoothly(l As SketchLine3D, ln As SketchLine3D, acr As DimensionConstraint3D,
                  acn As DimensionConstraint3D, acz As DimensionConstraint3D, lc As DimensionConstraint3D, ByRef k As Integer, ByRef j As Integer, ByRef i As Integer) As DimensionConstraint3D

        If (acr.Parameter._Value > 3 * Math.PI / (8)) And (acr.Parameter._Value < (Math.PI - 3 * Math.PI / (8))) Then
            If j > 15 Then
                j = 16
            Else
                j += 1
            End If
            AdjustTangentAngle(l, ln, acn, acr, acz, (256 - j) / 256)

            i = 1
        End If


        Return lc
    End Function

    Function AdjustTangentAngle(l As SketchLine3D, ln As SketchLine3D, acn As DimensionConstraint3D,
                             acr As DimensionConstraint3D, acz As DimensionConstraint3D, k As Double) As DimensionConstraint3D
        Dim j As Integer = 0
        For i = 1 To 16
            If (acr.Parameter._Value > k * 3 * Math.PI / (8)) And (acr.Parameter._Value < (Math.PI - k * 3 * Math.PI / (8))) Then
                If (acr.Parameter._Value < Math.PI / 2) Then
                    adjuster.AdjustDimensionConstraint3DSmothly(acr, acr.Parameter._Value * 15 / 16)
                Else
                    adjuster.AdjustDimensionConstraint3DSmothly(acr, acr.Parameter._Value * 17 / 16)
                End If

                acr.Driven = True
                If (acz.Parameter._Value < Math.PI / (12)) Or (acz.Parameter._Value > (Math.PI - (Math.PI / (12)))) Then
                    If j > 15 Then
                        j = 16
                    Else
                        j += 1
                    End If
                    AdjustZetaAngle(l, ln, acn, acz, (256 - j) / 256)
                    i = 1
                End If



            Else
                Exit For
            End If
        Next
        acr.Driven = True
        Return acr
    End Function
    Function AdjustZetaAngle(l As SketchLine3D, ln As SketchLine3D, acn As DimensionConstraint3D,
                             acz As DimensionConstraint3D, k As Double) As DimensionConstraint3D
        Dim j As Integer = 0
        For i = 1 To 16
            If (acz.Parameter._Value < k * Math.PI / (12)) Or (acz.Parameter._Value > (Math.PI - k * Math.PI / (12))) Then
                If (acz.Parameter._Value < k * Math.PI / 12) Then
                    adjuster.AdjustDimensionConstraint3DSmothly(acz, acz.Parameter._Value * 17 / 16)
                Else
                    adjuster.AdjustDimensionConstraint3DSmothly(acz, acz.Parameter._Value * 15 / 16)
                End If

                acz.Driven = True
                If Math.Abs(Math.Cos(acn.Parameter._Value)) < 1 / 2 Then
                    If j > 15 Then
                        j = 16
                    Else
                        j += 1
                    End If
                    acn = AdjustNormal(ln, l, acn, (256 - j) / 256)
                    i = 1
                End If


            Else
                Exit For
            End If
        Next
        acz.Driven = True
        Return acz
    End Function
    Function AdjustNormal(ln As SketchLine3D, li As SketchLine3D, acn As DimensionConstraint3D, k As Double) As DimensionConstraint3D
        Dim v1, v2 As Vector
        v2 = ln.Geometry.Direction.AsVector

        Dim a, c As Double

        For i = 1 To 16
            v1 = li.Geometry.Direction.AsVector
            a = v1.AngleTo(v2)
            c = Math.Cos(a)
            If c > 0 Then
                If c > k / (2) Then
                    Exit For

                Else
                    adjuster.AdjustDimensionConstraint3DSmothly(acn, acn.Parameter._Value * 15 / 16)
                End If
            Else
                If c < -1 * k / (2) Then
                    Exit For
                Else
                    If (c > -1 / 12) And i < 4 And k = 1 Then
                        adjuster.AdjustDimensionConstraint3DSmothly(acn, acn.Parameter._Value * 15 / 16)
                    Else
                        If acn.Parameter._Value > Math.PI / 2 Then
                            adjuster.AdjustDimensionConstraint3DSmothly(acn, acn.Parameter._Value * 17 / 16)
                        Else
                            adjuster.AdjustDimensionConstraint3DSmothly(acn, acn.Parameter._Value * 15 / 16)
                        End If
                    End If


                End If

            End If

        Next
        acn.Driven = True
        Return acn
    End Function
    Function IsCollision(skl As SketchLine3D, aci As DimensionConstraint3D, seq As SketchEquationCurve3D) As Boolean
        Dim wpt As WorkPoint
        Dim l As SketchLine3D
        Dim ac, dc As DimensionConstraint3D
        Dim skpt As SketchPoint3D
        Dim v1, v2 As Vector
        Dim gc As GeometricConstraint3D
        Return False

        Try
            For Each sb As SurfaceBody In tangentSurfaces.SurfaceBodies
                For Each f As Face In sb.Faces
                    If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        Try
                            wpt = compDef.WorkPoints.AddByCurveAndEntity(skl, f)
                            If IsPointOnFace(f, wpt.Point) Then
                                If IsPointInLine(skl, wpt.Point) Then
                                    Try
                                        v1 = palito.lref.StartSketchPoint.Geometry.VectorTo(palito.lref.EndSketchPoint.Geometry)
                                        l = sk3D.SketchLines3D.AddByTwoPoints(skl.StartPoint, skl.EndSketchPoint.Geometry, False)
                                        l.Construction = True
                                        gc = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, seq)
                                        dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                                        For index = 1 To 32
                                            If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 31 / 32) Then
                                                doku.Update2(True)
                                            Else
                                                Exit For
                                            End If

                                        Next
                                        doku.Update2(True)
                                        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, skl)
                                        aci.Driven = True
                                        For index = 1 To 16
                                            If IsPointOnFace(f, wpt.Point) Then
                                                If IsPointInLine(skl, wpt.Point) Then
                                                    skpt = sk3D.SketchPoints3D.Add(wpt.Point)

                                                    adjuster.AdjustDimensionConstraint3DSmothly(ac, ac.Parameter._Value * 15 / 16)
                                                Else

                                                    If FitCone(skl, skpt, f, ac) Then
                                                        Exit For
                                                    Else
                                                        adjuster.AdjustDimensionConstraint3DSmothly(ac, ac.Parameter._Value * 15 / 16)
                                                    End If
                                                End If
                                            Else

                                                If FitCone(skl, skpt, f, ac) Then
                                                    Exit For
                                                Else
                                                    adjuster.AdjustDimensionConstraint3DSmothly(ac, ac.Parameter._Value * 15 / 16)
                                                End If

                                            End If
                                        Next
                                        ac.Delete()
                                        doku.Update()
                                        v2 = skl.StartSketchPoint.Geometry.VectorTo(skpt.Geometry)
                                        freeCollisionAngle = aci.Parameter._Value
                                        dc.Delete()
                                        l.Delete()
                                        wpt.Delete()
                                        Return True
                                    Catch ex As Exception
                                        MsgBox(ex.ToString())
                                        Return Nothing
                                    End Try

                                End If
                            End If


                            wpt.Delete()

                        Catch ex As Exception
                        End Try
                    End If
                Next
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function FitCone(skl As SketchLine3D, skpti As SketchPoint3D, f As Face, aci As DimensionConstraint3D) As Boolean
        Dim ac2, ac1 As DimensionConstraint3D
        Dim l As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.StartPoint, skpti.Geometry, False)

        Dim lc As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, compDef.WorkPoints.Item(1), False)
        lc.Construction = True
        Dim lz As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, l.EndSketchPoint.Geometry, False)
        lz.Construction = True
        Dim gc As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddParallelToZAxis(lz)
        Dim lr As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, skl.StartSketchPoint.Geometry, False)
        lr.Construction = True
        Dim ac As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(lc, lr)
        adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
        ac.Delete()
        sk3D.GeometricConstraints3D.AddPerpendicular(lc, lr)
        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(lz, lr)
        adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
        ac.Delete()
        sk3D.GeometricConstraints3D.AddPerpendicular(lz, lr)
        Dim lb As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, l.EndPoint, False)
        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(lb, lr)
        adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
        ac.Delete()
        sk3D.GeometricConstraints3D.AddPerpendicular(lb, lr)
        Dim lr2 As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, lc.EndSketchPoint.Geometry, False)
        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(lr2, lz)
        adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
        ac.Delete()
        sk3D.GeometricConstraints3D.AddPerpendicular(lr2, lz)
        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(lr2, lr)
        adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
        ac.Delete()
        sk3D.GeometricConstraints3D.AddPerpendicular(lr2, lr)
        lr2.Construction = True
        Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(lb)
        adjuster.AdjustDimensionConstraint3DSmothly(dc, 8 / 10)
        Dim wpt As WorkPoint
        Dim skpt As SketchPoint3D
        dc = sk3D.DimensionConstraints3D.AddLineLength(skl)
        Try
            wpt = compDef.WorkPoints.AddByCurveAndEntity(l, f)
            aci.Driven = True
            For index = 1 To 16
                If IsPointOnFace(f, wpt.Point) Then
                    If IsPointInLine(l, wpt.Point) Then
                        skpt = sk3D.SketchPoints3D.Add(wpt.Point)
                        FitCone = adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 15 / 16)
                    Else
                        ac1 = sk3D.DimensionConstraints3D.AddTwoLineAngle(lr2, lb)
                        ac1.Driven = True
                        ac2 = sk3D.DimensionConstraints3D.AddTwoLineAngle(lz, lb)
                        If IsRotableCone(dc, ac2, ac1, wpt, f, l) Then
                            FitCone = True
                            Exit For
                        Else
                            FitCone = adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 15 / 16)
                        End If
                        ac2.Delete()
                        ac1.Delete()
                    End If
                Else
                    ac1 = sk3D.DimensionConstraints3D.AddTwoLineAngle(lr2, lb)
                    ac1.Driven = True
                    ac2 = sk3D.DimensionConstraints3D.AddTwoLineAngle(lz, lb)
                    If IsRotableCone(dc, ac2, ac1, wpt, f, l) Then
                        FitCone = True
                        Exit For
                    Else
                        FitCone = adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 15 / 16)
                    End If
                    ac2.Delete()
                    ac1.Delete()

                End If
            Next
            wpt.Delete()
            dc.Delete()
            Return FitCone
        Catch ex As Exception

            Return True
        End Try
    End Function
    Function MakeConeFree(skli As SketchLine3D, skpti As SketchPoint3D) As SketchLine3D

        Dim d, dMin As Double
        Dim fi As Face = corteFace
        Dim pl As Plane = fi.Geometry
        Dim seq As SketchEquationCurve3D = lastSeq
        Dim puntos As ObjectsEnumerator
        Dim ptOpt As Point



        Try
            dMin = 99999
            comando.WireFrameView(doku)
            Try
                puntos = pl.IntersectWithCurve(seq.Geometry)
                ptOpt = puntos(1)
            Catch ex As Exception
                ptOpt = skli.EndSketchPoint.Geometry
            End Try
            Dim skptOpt As SketchPoint3D = sk3D.SketchPoints3D.Add(ptOpt)
            sk3D.GeometricConstraints3D.AddGround(skptOpt)
            Dim lref As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skli.StartPoint, skptOpt)
            lref.Construction = True
            Dim skl As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skli.StartPoint, skli.EndSketchPoint.Geometry, False)
            Dim gci As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddCoincident(skl.EndPoint, seq)
            Dim acRef As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(lref, skl)
            Dim l As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.StartPoint, skpti.Geometry, False)

            Dim lc As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, compDef.WorkPoints.Item(1), False)
            lc.Construction = True
            Dim lz As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, l.EndSketchPoint.Geometry, False)
            lz.Construction = True
            Dim gc As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddParallelToZAxis(lz)
            Dim lr As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, skl.StartSketchPoint.Geometry, False)
            lr.Construction = True
            Dim ac As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(lc, lr)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Delete()
            sk3D.GeometricConstraints3D.AddPerpendicular(lc, lr)
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(lz, lr)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Delete()
            sk3D.GeometricConstraints3D.AddPerpendicular(lz, lr)
            Dim lb As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.EndPoint, l.EndPoint, False)
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(lb, lr)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Delete()
            sk3D.GeometricConstraints3D.AddPerpendicular(lb, lr)
            Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(lb)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, 8 / 10)
            Dim wpt As WorkPoint
            Dim skpt As SketchPoint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(skl)
            Try
                wpt = compDef.WorkPoints.AddByCurveAndEntity(l, fi)

                For index = 1 To 16
                    If IsPointOnFace(fi, wpt.Point) Then
                        If IsPointInLine(l, wpt.Point) Then
                            skpt = sk3D.SketchPoints3D.Add(wpt.Point)
                            dc.Driven = True
                            adjuster.AdjustDimensionConstraint3DSmothly(acRef, acRef.Parameter._Value * 15 / 16)
                            acRef.Driven = True
                            adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 15 / 16)
                        Else

                        End If
                    Else


                    End If
                Next
                wpt.Delete()
                dc.Delete()
                comando.RealisticView(doku)
                Return skl
            Catch ex2 As Exception


            End Try
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function IsRotableCone(dci As DimensionConstraint3D, aci As DimensionConstraint3D, acc As DimensionConstraint3D, wpt As WorkPoint, f As Face, l As SketchLine3D) As Boolean
        Dim skpt As SketchPoint3D
        Dim sign, pieces As Integer
        pieces = 6
        If aci.Parameter._Value < Math.PI / pieces Then
            sign = -1
        Else
            sign = 1
        End If

        For i = 1 To pieces
            If IsPointOnFace(f, wpt.Point) Then
                If IsPointInLine(l, wpt.Point) Then
                    i = 1
                    skpt = sk3D.SketchPoints3D.Add(wpt.Point)
                    IsRotableCone = adjuster.AdjustDimensionConstraint3DSmothly(dci, dci.Parameter._Value * 31 / 32)
                    acc.Driven = True
                    aci.Driven = False
                Else
                    If aci.Parameter._Value < Math.PI / pieces Or aci.Parameter._Value > (pieces - 1) * Math.PI / pieces Or (Not acc.Driven) Then
                        aci.Driven = True
                        adjuster.AdjustDimensionConstraint3DSmothly(acc, acc.Parameter._Value + sign * Math.PI / pieces)
                        If acc.Parameter._Value > 11 * Math.PI / 12 Then
                            acc.Driven = True
                            aci.Driven = False
                        End If
                    Else
                        acc.Driven = True
                        adjuster.AdjustDimensionConstraint3DSmothly(aci, aci.Parameter._Value - sign * Math.PI / 12)
                    End If

                    IsRotableCone = True
                End If
            Else
                If aci.Parameter._Value < Math.PI / pieces Or aci.Parameter._Value > (pieces - 1) * Math.PI / pieces Or (Not acc.Driven) Then
                    aci.Driven = True
                    adjuster.AdjustDimensionConstraint3DSmothly(acc, acc.Parameter._Value + sign * Math.PI / pieces)
                    If acc.Parameter._Value > 11 * Math.PI / 12 Then
                        acc.Driven = True
                        aci.Driven = False
                    End If
                Else
                    acc.Driven = True
                    adjuster.AdjustDimensionConstraint3DSmothly(aci, aci.Parameter._Value - sign * Math.PI / pieces)
                End If
                IsRotableCone = True
            End If
        Next
        Return IsRotableCone
    End Function
    Function IsPointInLine(l As SketchLine3D, pt As Point) As Boolean

        Dim d As Double = l.StartSketchPoint.Geometry.DistanceTo(pt)
        Dim e As Double = l.EndSketchPoint.Geometry.DistanceTo(pt)
        If Math.Abs(d + e - l.Length) > 1 / 128 Then
            Return False
        Else
            Return True
        End If

    End Function
    Function IsPointOnFace(f As Face, pt As Point) As Boolean
        Dim spt As SketchPoint3D = sk3D.SketchPoints3D.Add(f.GetClosestPointTo(pt))
        Dim d As Double = f.GetClosestPointTo(pt).DistanceTo(pt)

        If (d > f.Evaluator.Area / 128) Then
            Return False
        Else
            Return True
        End If

    End Function
    Function GetSpikeProfile(fc As Face) As Profile
        Try
            Dim pr As Profile
            Dim ps As PlanarSketch
            Dim el, sl As SketchLine

            ps = doku.ComponentDefinition.Sketches.Add(fc)
            For Each ed As Edge In fc.Edges

                el = ps.AddByProjectingEntity(ed)
                el.Construction = True


                sl = ps.SketchLines.AddByTwoPoints(el.StartSketchPoint, el.EndSketchPoint)
            Next
            ' ps.SketchLines.AddAsThreePointRectangle(twistFace.Vertices.Item(1).Poin, twistFace.Vertices.Item(2).Point, twistFace.Vertices.Item(3).Point)



            pr = ps.Profiles.AddForSolid

            Return pr
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function GetConeProfile(spt As SketchPoint3D) As Profile
        Try
            Dim pr As Profile
            Dim ps As PlanarSketch
            wptConeBase = compDef.WorkPoints.AddByPoint(spt)
            wptConeBase.Visible = False
            ' sk3D = compDef.Sketches3D.Add
            'ringLine = sk3D.SketchLines3D.AddByTwoPoints(compDef.WorkPoints.Item(1), spt)
            'Dim skpt As SketchPoint3D = sk3D.SketchPoints3D.Add(tg.CreatePoint(0, 0, 1))
            'Dim wpl As WorkPlane = compDef.WorkPlanes.AddByThreePoints(compDef.WorkPoints.Item(1), skpt, spt)
            Dim wpl As WorkPlane = compDef.WorkPlanes.AddByNormalToCurve(coneLine, spt)
            ps = doku.ComponentDefinition.Sketches.Add(wpl)
            lamp.LookAtPlane(wpl)
            pr = DrawSingleCircle(ps, 6 / 10)
            ps.Visible = False
            'sk3D.Visible = False
            wpl.Visible = False
            Return pr
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawSingleCircle(ps As Sketch, r As Double) As Profile
        Dim pro As Profile

        ' Dim sl As SketchLine = ps.AddByProjectingEntity(coneLine)
        Dim spt As SketchPoint = ps.AddByProjectingEntity(wptConeBase)
        Try
            '  sl.Construction = True
            ps.SketchCircles.AddByCenterRadius(spt, r)
            pro = ps.Profiles.AddForSolid

            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawCircleOnLine(skl As SketchLine3D, r As Double) As Profile
        Dim pro As Profile
        Dim ps As PlanarSketch
        Dim wpl As WorkPlane
        Dim spt As SketchPoint
        Try
            wpl = doku.ComponentDefinition.WorkPlanes.AddByNormalToCurve(skl, skl.StartPoint)
            wpl.Visible = False
            lamp.LookAtPlane(wpl)
            ps = doku.ComponentDefinition.Sketches.Add(wpl)

            spt = ps.AddByProjectingEntity(skl.StartPoint)
            ps.SketchCircles.AddByCenterRadius(spt, r)
            pro = ps.Profiles.AddForSolid

            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetRails(fc As Face) As Profile3D
        Try
            Dim l1 As SketchLine3D
            Dim pt As Point = doku.ComponentDefinition.WorkPoints.Item(1).Point
            Dim d As Double = 999999999

            Dim pr3d As Profile3D
            rails.Clear()

            For Each v As Vertex In fc.Vertices
                sk3D = compDef.Sketches3D.Add()
                l1 = sk3D.SketchLines3D.AddByTwoPoints(v.Point, pt)
                pr3d = sk3D.Profiles3D.AddOpen
                rails.Add(pr3d)
            Next



            Return pr3d

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return Nothing
    End Function
    Function LoftSingleSpike() As LoftFeature
        Try
            Dim oLoftDefinition As LoftDefinition
            oLoftDefinition = compDef.Features.LoftFeatures.CreateLoftDefinition(sections, PartFeatureOperationEnum.kJoinOperation)

            ' oLoftDefinition.Closed = True

            Dim lf As LoftFeature
            Try
                lf = compDef.Features.LoftFeatures.Add(oLoftDefinition)
            Catch ex As Exception
                lf = CorrectLoftSingleSpike()
            End Try


            Return lf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function CorrectLoftSingleSpike() As LoftFeature

        sections.Remove(sections.Count)
        wptConeBase = TryStraigthSpike()
        doku.Update()
        sections.Add(wptConeBase)
        Dim oLoftDefinition As LoftDefinition
        oLoftDefinition = compDef.Features.LoftFeatures.CreateLoftDefinition(sections, PartFeatureOperationEnum.kJoinOperation)
        Try
            Return compDef.Features.LoftFeatures.Add(oLoftDefinition)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return Nothing
    End Function
    Function TryTangentSpike() As DimensionConstraint3D
        Dim d As Double = angleConeLineTangent.Parameter._Value
        For i = 1 To 32
            adjuster.AdjustDimConstrain3DSmothly(angleConeLineTangent, angleConeLineTangent.Parameter._Value * 15 / 16)
            If Math.Abs(d - angleConeLineTangent.Parameter._Value) < Math.PI / 256 Then
                Exit For
            Else
                d = angleConeLineTangent.Parameter._Value
            End If
        Next
        angleConeLineTangent.Driven = True
        Return angleConeLineTangent
    End Function
    Function TryStraigthSpike() As WorkPoint
        Dim d As Double = angleConeNormalLine.Parameter._Value
        Dim pt As Point
        Dim wptStraigth As WorkPoint
        Dim skpt As SketchPoint3D
        Dim v As Vector
        For i = 1 To 32
            adjuster.AdjustDimConstrain3DSmothly(angleConeNormalLine, angleConeNormalLine.Parameter._Value * 15 / 16)
            If Math.Abs(d - angleConeNormalLine.Parameter._Value) < Math.PI / 256 Then
                Exit For
            Else
                d = angleConeNormalLine.Parameter._Value
            End If
        Next
        angleConeNormalLine.Driven = True
        v = coneLine.Geometry.Direction.AsVector()
        v.ScaleBy(coneLine.Length)
        pt = coneLine.EndSketchPoint.Geometry
        pt.TranslateBy(v)
        skpt = sk3D.SketchPoints3D.Add(pt)
        wptStraigth = compDef.WorkPoints.AddByPoint(skpt)
        wptStraigth.Visible = False
        Return wptStraigth
    End Function

    Function LoftSingleCone(wpti As WorkPoint) As LoftFeature
        Try
            Dim oLoftDefinition As LoftDefinition
            oLoftDefinition = compDef.Features.LoftFeatures.CreateLoftDefinition(sections, PartFeatureOperationEnum.kJoinOperation)
            Dim lf As LoftFeature
            Try
                'lf = compDef.Features.LoftFeatures.Add(oLoftDefinition)
                lf = LoftSpecialCone(wpti)
            Catch ex As Exception
                lf = LoftSpecialCone(wpti)
            End Try

            Return lf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function LoftSpecialCone(wpti As WorkPoint) As LoftFeature
        Dim f As Face = GetClosestFace(wpti)
        Dim ps As PlanarSketch = compDef.Sketches.Add(f)
        lamp.LookAtFace(f)
        For Each ed As Edge In f.Edges
            ps.AddByProjectingEntity(ed)
        Next
        Dim pro As Profile = ps.Profiles.AddForSolid
        Dim pro2 As Profile = sections(2)
        sections.Clear()
        sections.Add(pro)
        sections.Add(pro2)
        ps.Visible = False
        Return LoftSingleSpike()
    End Function
    Function MakeLoftCone(f As Face, skpt As SketchPoint3D) As LoftFeature
        Dim lf As LoftFeature
        Dim ps As PlanarSketch = compDef.Sketches.Add(f)
        lamp.LookAtFace(f)
        Dim se As SketchEntity
        Dim sl As SketchLine
        Dim ssl As SketchSpline
        Dim sea As SketchEllipticalArc
        Dim spt As SketchPoint
        Dim ska As SketchArc
        Dim sce As SketchConstraintsEnumerator
        For Each ed As Edge In f.Edges
            se = ps.AddByProjectingEntity(ed)

            Select Case se.Type
                Case ObjectTypeEnum.kSketchPointObject
                    spt = se
                    sce = spt.Constraints
                    For Each gc As GeometricConstraint In sce
                        gc.Delete()
                    Next
                Case ObjectTypeEnum.kSketchLineObject
                    sl = se
                    sce = sl.Constraints
                    For Each gc As GeometricConstraint In sce
                        gc.Delete()
                    Next
                Case ObjectTypeEnum.kSketchSplineObject
                    ssl = se
                    sce = ssl.Constraints
                    For Each gc As GeometricConstraint In sce
                        gc.Delete()
                    Next
                Case ObjectTypeEnum.kSketchEllipticalArcObject
                    sea = se
                    sce = sea.Constraints
                    For Each gc As GeometricConstraint In sce
                        gc.Delete()
                    Next
                Case ObjectTypeEnum.kSketchArcObject
                    ska = se
                    sce = ska.Constraints
                    For Each gc As GeometricConstraint In sce
                        gc.Delete()
                    Next
                Case Else

            End Select
        Next
        For Each gc As GeometricConstraint In ps.GeometricConstraints
            gc.Delete()
        Next
        Try
            ps.BreakLink()
        Catch ex As Exception

        End Try

        Dim pro As Profile = ps.Profiles.AddForSolid
        Dim pro2 As Profile = GetConeProfile(skpt)
        sections.Clear()
        sections.Add(pro)
        sections.Add(pro2)
        ps.Visible = False
        lf = LoftSingleSpike()

        Return lf
    End Function
    Function LoftFilling() As LoftFeature
        Try
            Dim oLoftDefinition As LoftDefinition
            oLoftDefinition = compDef.Features.LoftFeatures.CreateLoftDefinition(sections, PartFeatureOperationEnum.kJoinOperation)

            ' oLoftDefinition.Closed = True


            Dim lf As LoftFeature
            lf = compDef.Features.LoftFeatures.Add(oLoftDefinition)
            Return lf
        Catch ex As Exception
            Return Nothing
        End Try


    End Function
    Function IsPointContained(Point As Point, body As SurfaceBody) As Boolean

        Dim tol As Double = 0.0001
        IsPointContained = False
        Dim TxBrep As TransientBRep
        TxBrep = app.TransientBRep
        Dim ptBody As SurfaceBody
        ptBody = TxBrep.CreateSolidSphere(Point, tol)
        Dim vol1 As Double
        vol1 = ptBody.Volume(tol / 10)
        Dim vol2 As Double
        vol2 = body.Volume(tol / 10)
        TxBrep.DoBoolean(ptBody, body, BooleanTypeEnum.kBooleanTypeUnion)
        Dim volRes As Double
        volRes = ptBody.Volume(tol / 10)
        If (volRes < vol1 + vol2) Then
            Return True
        Else
            Return False

        End If
    End Function

End Class
