Imports Inventor

Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory
Public Class RodMaker
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Public centroLine, lref As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy, smallOval As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile
    Dim adjuster As SketchAdjust
    Dim cortoCounter As Integer



    Public wp1, wp2, wp3, wpHigh, wpLow As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM, angleRod As Double
    Public partNumber, qNext, qLastTie, cutside As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Dim cylinderCenter As SketchPoint3D

    Dim cutProfile, faceProfile, rodProfile As Profile

    Dim profile As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Public ovalLine As SketchLine3D
    Public compDef As PartComponentDefinition
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim mainWorkPlane As WorkPlane
    Dim workAxis As WorkAxis
    Dim faceRod As Face
    Dim adjacentFace, bendFace, frontBendFace, cutFace, twistFace, endface, cutTipFace1, cutTipface2 As Face
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, cutEdge1, cutEdge2, CutEsge3, leadingEdge, followEdge As Edge
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex As DimensionConstraint3D
    Dim largo As DimensionConstraint
    Dim folded As FoldFeature
    Public foldFeatures As FoldFeatures
    Dim features As SheetMetalFeatures
    Dim lamp As Highlithing
    Dim di As System.IO.DirectoryInfo
    Dim fi As System.IO.File
    Dim nf As System.IO.Path
    Public wp2Face, wp1Face As Face
    Dim foldFeature As FoldFeature
    Public rodLines, esquinas, rails, caras, surfaceBodies, arcPoints As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane

    Dim arrayFunctions As Collection
    Dim fullFileNames As String()
    Dim extrudeRod As ExtrudeFeature
    Dim tangentSurfaces As WorkSurface
    Dim endFaceKey As Long
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        invFile = New InventorFile(app)
        projectManager = app.DesignProjectManager
        cortoCounter = 1
        compDef = doku.ComponentDefinition

        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        rodLines = app.TransientObjects.CreateObjectCollection
        esquinas = app.TransientObjects.CreateObjectCollection
        rails = app.TransientObjects.CreateObjectCollection
        caras = app.TransientObjects.CreateObjectCollection
        surfaceBodies = app.TransientObjects.CreateObjectCollection
        arcPoints = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)


        nombrador = New Nombres(doku)



        done = False
    End Sub

    Function MakeAllRods(ws As WorkSurface) As Integer
        Dim tf As Face
        Dim fc As FaceCollection
        Dim fi As Integer
        doku = ws.Parent.Document
        compDef = doku.ComponentDefinition


        lamp = New Highlithing(doku)

        Try

            tf = GetStartingface(ws)
            MakeSingleRod(tf)
            sk3D.Visible = False
            tf = GetStartingface(ws)
            fc = tf.TangentiallyConnectedFaces
            fi = fc.Count
            For i = 1 To fi
                If fc.Item(i).SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    ' lamp.HighLighFace(fc.Item(i))
                    MakeSingleRod(fc.Item(i))
                    sk3D.Visible = False
                    tf = GetStartingface(ws)
                    fc = tf.TangentiallyConnectedFaces
                End If
            Next



            Return doku.ComponentDefinition.Features.RevolveFeatures.Count
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function MakeSkeletonRods(ws As WorkSurface) As Integer

        Dim sb As SurfaceBody = ws.SurfaceBodies.Item(1)

        lamp = New Highlithing(doku)

        Try
            wpHigh = compDef.WorkPoints.Item("wptHigh")
            wpLow = compDef.WorkPoints.Item("wptLow")

            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    ' lamp.HighLighFace(fc.Item(i))
                    MakeSingleRod(f)


                End If
            Next



            Return doku.ComponentDefinition.Features.RevolveFeatures.Count
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
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
            Return fmin1
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function MakeSingleRod(f As Face) As RevolveFeature
        Dim wa As WorkAxis
        Dim le As SketchLine3D
        Try
            comando.WireFrameView(doku)
            ' lamp.HighLighFace(f)
            wa = doku.ComponentDefinition.WorkAxes.AddByRevolvedFace(f)
            wa.Visible = False
            workAxis = wa

            le = DrawRodSketch3D(f, wa)
            rodProfile = DrawPlannSketch(centroLine, le)
            comando.RealisticView(doku)
            Return RevolveFace(rodProfile)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Public Function ExtrudeEgg(fi As Face) As ExtrudeFeature
        Dim rf As RevolveFeature = MakeSingleRod(fi)
        Dim d, e, dMin, eMin As Double
        Dim pro As Profile
        Dim f1 As Face
        Dim ef As ExtrudeFeature
        Try
            dMin = 9999
            eMin = dMin
            For Each f As Face In rf.SideFaces
                lamp.HighLighFace(f)

                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    d = f.GetClosestPointTo(wp1.Point).DistanceTo(wp1.Point)
                    If d < dMin Then
                        dMin = d
                        f1 = f

                    End If
                    e = f.GetClosestPointTo(wp2.Point).DistanceTo(wp2.Point)
                    If e < eMin Then
                        eMin = e
                        wp2Face = f
                    End If
                End If
            Next
            wp1Face = f1
            lamp.LookAtFace(f1)
            pro = DrawOvalSketch(wp1Face, fi)
            If smallOval Then

                ef = ExtrudeReverseOval(ChangeFaceOval(pro))
                ef.Name = String.Concat("egg", cortoCounter)
                rf.Name = String.Concat("shortRod", cortoCounter)
                cortoCounter = cortoCounter + 1


            Else
                ef = ExtrudeOval(pro)
            End If
            extrudeRod = ef
            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function MakeTipCut(topSide As Integer) As RevolveFeature
        Dim f As Face
        Dim ed As Edge
        Dim pro As Profile
        cutside = topSide
        Try
            f = GetTipCutFace(topSide)
            ed = GetTipCutEdge(f)
            pro = DrawTipCutProfile(f, ed)
            MakeTipCut = RevolveCutTip(pro)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return MakeTipCut
    End Function

    Public Function GetTipCutFace(side As Integer) As Face
        Dim pl As Plane = GetCutTipEndFace().Geometry
        Dim vnef As Vector = wp2.Point.VectorTo(wp1.Point)
        Dim vnsf, vcp As Vector
        Dim d, dMax, aMax1, aMax2, dis, minDis1, minDis2 As Double
        minDis1 = 99999
        dMax = -99999
        minDis2 = 99999
        cutTipface2 = extrudeRod.SideFaces.Item(1)
        For Each f As Face In extrudeRod.SideFaces
            If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                dis = f.PointOnFace.DistanceTo(endface.PointOnFace) / f.Evaluator.Area

                If dis < minDis2 Then
                    pl = f.Geometry
                    vnsf = pl.Normal.AsVector
                    vcp = vnef.CrossProduct(vnsf)
                    d = vcp.Z * f.Evaluator.Area
                    If d > dMax Then
                        dMax = d
                        GetTipCutFace = f
                        cutTipface2 = cutTipFace1
                        cutTipFace1 = f
                        lamp.HighLighFace(f)

                    Else
                        cutTipface2 = f

                    End If
                    If dis < minDis1 Then
                        minDis2 = minDis1
                        minDis1 = dis

                    Else
                        minDis2 = dis

                    End If


                Else

                End If

            End If
        Next
        lamp.HighLighFace(cutTipface2)
        lamp.HighLighFace(GetTipCutFace)

        Return GetTipCutFace
    End Function
    Function GetTipCutEdge(fi As Face) As Edge
        tangentSurfaces = compDef.WorkSurfaces.Item("tangents")
        Dim vRod As Vector = wp2.Point.VectorTo(wp1.Point)
        Dim pl As Plane
        Dim ve, vpl, vfi, v As Vector
        Dim d, dMax, eMax, dis, minDis As Double
        pl = fi.Geometry
        vfi = pl.Normal.AsVector
        dMax = 0
        eMax = 0
        minDis = 9999
        Dim sf As Face
        For Each sb As SurfaceBody In tangentSurfaces.SurfaceBodies
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    dis = f.PointOnFace.DistanceTo(fi.PointOnFace)
                    If dis < minDis Then
                        minDis = dis
                        pl = f.Geometry
                        vpl = pl.Normal.AsVector
                        d = Math.Abs(vpl.DotProduct(vfi))
                        If d > eMax Then
                            eMax = d

                            sf = f
                        End If
                    End If



                End If
            Next
        Next
        lamp.HighLighFace(sf)
        dMax = 0
        pl = sf.Geometry
        vpl = pl.Normal.AsVector
        For Each ed As Edge In sf.Edges
            ve = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point)
            d = Math.Abs((ve.CrossProduct(vRod)).DotProduct(vpl))
            If d > dMax Then
                dMax = d
                GetTipCutEdge = ed
            End If
        Next


        lamp.HighLighObject(GetTipCutEdge)

        Return GetTipCutEdge
    End Function
    Function DrawTipCutProfile(fi As Face, edi As Edge) As Profile
        Dim ps As PlanarSketch
        Dim sl, slRef As SketchLine
        Dim spt, sptMid, sptVertex As SketchPoint
        Dim gc As GeometricConstraint
        Dim dc As DimensionConstraint
        Dim see As SketchEntitiesEnumerator
        '   Dim ve As Vector2d = edi.StartVertex.Point.VectorTo(edi.StopVertex.Point)
        Dim pt, ptMid As Point2d
        Dim m As Matrix2d = tg.CreateMatrix2d()
        Dim sign As Integer
        'Dim vePer = tg.CreateVector2d(ve.Y, -ve.X)
        Try
            lamp.LookAtFace(fi)

            ps = compDef.Sketches.Add(fi)
            slRef = ps.AddByProjectingEntity(edi)
            slRef.Construction = True
            spt = ps.AddByProjectingEntity(wp1)
            sptVertex = ps.AddByProjectingEntity(GetClosestVertexCut(edi))
            sl = ps.SketchLines.AddByTwoPoints(sptVertex, spt.Geometry)
            gc = ps.GeometricConstraints.AddParallel(sl, slRef)
            dc = ps.DimensionConstraints.AddTwoPointDistance(sl.StartSketchPoint, sl.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, sl.Geometry.MidPoint, False)
            adjuster.AdjustDimension2DSmothly(dc, 25 / 10)
            sl.Construction = True
            pt = sl.Geometry.StartPoint
            If sptVertex.Geometry3d.Z > 0 Then
                sign = 1
            Else
                sign = -1
            End If

            m.SetToRotation(sign * Math.PI / 2, pt)
            ptMid = sl.Geometry.MidPoint
            ptMid.TransformBy(m)
            sptMid = ps.SketchPoints.Add(ptMid)
            see = ps.SketchLines.AddAsThreePointRectangle(sptVertex, sl.EndSketchPoint, sptMid.Geometry)
            DrawTipCutProfile = ps.Profiles.AddForSolid()
            Return DrawTipCutProfile
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function GetClosestVertexCut(edi As Edge) As Vertex
        Dim f As Face = GetCutTipEndFace()
        Dim d, e, dMin As Double
        dMin = 999999999
        Try
            For Each v As Vertex In f.Vertices
                d = v.Point.DistanceTo(edi.StartVertex.Point)
                e = v.Point.DistanceTo(edi.StopVertex.Point)
                If d < e Then
                    If d < dMin Then
                        dMin = d
                        GetClosestVertexCut = v
                    End If
                Else
                    If e < dMin Then
                        dMin = e
                        GetClosestVertexCut = v
                    End If
                End If

            Next



            Return GetClosestVertexCut
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Function MakeBridge(f As Face) As ExtrudeFeature
        Dim wa As WorkAxis
        Dim le As SketchLine3D
        Dim pro As Profile
        Try
            ' lamp.HighLighFace(f)
            wa = doku.ComponentDefinition.WorkAxes.AddByRevolvedFace(f)
            workAxis = wa
            wa.Visible = False
            GetArcPoints(f)
            pro = DrawArcSketch(f, wa)
            Return ExtrudeSurface(pro)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function GetArcPoints(fi As Face) As Point
        Dim min1, min2, min3, min4, d As Double
        Dim pt1, pt2, pt3, pt4, ptc As Point
        Dim ver1, ver2 As Vertex
        Dim l As Line
        Dim c As Cylinder
        Dim fc As FaceCollection = app.TransientObjects.CreateFaceCollection
        Dim b As Boolean = False
        min1 = 99999999
        min2 = min1
        min3 = min2
        min4 = min3


        pt1 = fi.Vertices.Item(1).Point
        pt2 = pt1
        pt3 = pt2
        pt4 = pt3
        If fi.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
            c = fi.Geometry
            ptc = c.BasePoint
        End If
        l = tg.CreateLine(c.BasePoint, c.AxisVector.AsVector)
        arcPoints.Clear()
        ver1 = fi.Vertices.Item(1)
        fc = fi.TangentiallyConnectedFaces
        For Each f As Face In fc
            If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                For Each ver As Vertex In f.Vertices
                    d = l.DistanceTo(ver.Point)
                    If d < min2 Then
                        If d < min1 Then
                            min1 = d

                            pt2 = pt1
                            pt1 = ver.Point
                            ver2 = ver1
                            ver1 = ver
                        Else
                            min2 = d
                            pt2 = ver.Point
                            ver2 = ver
                        End If

                    End If
                Next
            End If


        Next

        sk3D = doku.ComponentDefinition.Sketches3D.Add()

        arcPoints.Add(sk3D.SketchPoints3D.Add(pt1))
        arcPoints.Add(sk3D.SketchPoints3D.Add(pt2))
        For Each f As Face In ver1.Faces

            For Each f3 As Face In fc
                If f.Equals(f3) Then
                    fc.RemoveByObject(f3)
                    b = True
                    Exit For
                End If

            Next
            If b Then
                Exit For
            End If
        Next
        min1 = 99999999
        min2 = min1

        For Each f As Face In fc
            If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                For Each ver As Vertex In f.Vertices
                    d = l.DistanceTo(ver.Point)
                    If d < min2 Then
                        If d < min1 Then
                            min1 = d

                            pt2 = pt1
                            pt1 = ver.Point
                            ver2 = ver1
                            ver1 = ver
                        Else
                            min2 = d
                            pt2 = ver.Point
                            ver2 = ver
                        End If

                    End If
                Next
            End If

        Next

        arcPoints.Add(sk3D.SketchPoints3D.Add(pt1))
        arcPoints.Add(sk3D.SketchPoints3D.Add(pt2))

        Return pt1
    End Function
    Function GetRadiusPoint(pt As Point, r As Double) As Double
        Dim radius As Double = Math.Pow(Math.Pow(pt.X, 2) + Math.Pow(pt.Y, 2), 1 / 2)
        Return radius
    End Function
    Function DrawRodSketch3D(f As Face, a As WorkAxis) As SketchLine3D
        Dim l, le As SketchLine3D
        Dim gc As GeometricConstraint3D
        sk3D = doku.ComponentDefinition.Sketches3D.Add()

        le = sk3D.Include(GetMajorEdge(f))
        le.Construction = True
        l = sk3D.SketchLines3D.AddByTwoPoints(le.StartSketchPoint.Geometry, le.EndSketchPoint.Geometry, False)
        l.Construction = True
        gc = sk3D.GeometricConstraints3D.AddCollinear(l, a)
        centroLine = l
        sk3D.Visible = False
        Return le
    End Function
    Function ExtrudeSurface(pro As Profile) As ExtrudeFeature
        Dim oExtrudeDef As ExtrudeDefinition
        Dim spt As SketchPoint3D = GetFarPoint()
        Try

            oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kSurfaceOperation)
            oExtrudeDef.SetToExtent(spt)

            'oExtrudeDef.SetDistanceExtent(0.12, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)
            Return oExtrude
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function ChangeFaceOval(pro As Profile) As Profile

        Dim pr As Profile
        Try
            Dim psi As Sketch = pro.Parent

            lamp.LookAtFace(wp2Face)
            Dim ps As PlanarSketch = compDef.Sketches.Add(wp2Face)
            For Each se As SketchEntity In psi.SketchEntities
                If se.Construction = False Then
                    ps.AddByProjectingEntity(se)
                End If

            Next

            pr = ps.Profiles.AddForSolid
            ps.Visible = False
            Return pr
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function ExtrudeOval(pro As Profile) As ExtrudeFeature
        Dim oExtrudeDef As ExtrudeDefinition
        Dim spt As SketchPoint3D = sk3D.SketchPoints3D.Add(wp2.Point)
        Try
            oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kJoinOperation)
            oExtrudeDef.SetToExtent(spt)
            'oExtrudeDef.SetDistanceExtent(0.12, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oExtrude As ExtrudeFeature
            Try
                oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)
            Catch ex As Exception
                oExtrude = ExtrudeReverseOval(ChangeFaceOval(pro))
            End Try
            Return oExtrude
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function ExtrudeReverseOval(pro As Profile) As ExtrudeFeature
        Dim oExtrudeDef As ExtrudeDefinition


        Try
            Dim ls As LineSegment = tg.CreateLineSegment(cylinderCenter.Geometry, wp1.Point)
            Dim spt As SketchPoint3D = sk3D.SketchPoints3D.Add(wp1.Point)
            Dim d As Double = spt.Geometry.DistanceTo(wp2.Point) * 15 / 16
            oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kJoinOperation)
            oExtrudeDef.SetDistanceExtent(d, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)

            'oExtrudeDef.SetDistanceExtent(0.12, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oExtrude As ExtrudeFeature
            Try
                oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)
            Catch ex As Exception
                oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kJoinOperation)
                oExtrudeDef.SetToExtent(cylinderCenter)
                oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)
            End Try

            Return oExtrude
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function GetFarPoint() As SketchPoint3D
        Dim ptMax2 As SketchPoint3D
        Dim d, dMax As Double
        dMax = 0
        Dim fpt As SketchPoint3D = arcPoints.Item(1)
        Try
            For Each spt As SketchPoint3D In arcPoints
                d = fpt.Geometry.DistanceTo(spt.Geometry)
                If d > dMax Then
                    dMax = d
                    ptMax2 = spt
                End If
            Next

            Return ptMax2
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawArcSketch(f As Face, wa As WorkAxis) As Profile
        Try
            Dim wpl As WorkPlane
            Dim ps As PlanarSketch
            Dim v As Vector
            Dim skpt3DStart As SketchPoint3D
            Dim sp1, sp2, sp3, spc As SketchPoint
            Dim le2d, l2d, l32d, mne As SketchLine
            Dim skAxis As SketchLine3D
            Dim gc, gccc As GeometricConstraint
            Dim r As SketchEntitiesEnumerator
            Dim sa As SketchArc
            Dim dc As DimensionConstraint
            Dim pro As Profile
            Dim pt3D As Point
            Dim d As Double = 0
            Dim i As Integer
            Dim c As Cylinder
            Dim pl As Plane
            Dim dis As Double
            Dim b As Boolean
            c = f.Geometry
            sk3D.Visible = False

            pt3D = c.BasePoint
            pt3D.TranslateBy(c.AxisVector.AsVector)

            skAxis = sk3D.SketchLines3D.AddByTwoPoints(c.BasePoint, pt3D, False)
            skAxis.Construction = True
            wpl = doku.ComponentDefinition.WorkPlanes.AddByNormalToCurve(skAxis, skAxis.StartPoint)
            wpl.Visible = False
            skpt3DStart = arcPoints.Item(1)
            wpl = compDef.WorkPlanes.AddByPlaneAndPoint(wpl, skpt3DStart)
            wpl.Visible = False
            ps = doku.ComponentDefinition.Sketches.Add(wpl)
            sp1 = ps.AddByProjectingEntity(arcPoints.Item(1))
            sp2 = ps.AddByProjectingEntity(arcPoints.Item(3))
            spc = ps.AddByProjectingEntity(skAxis.StartPoint)
            v = c.BasePoint.VectorTo(sp1.Geometry3d)
            If v.DotProduct(c.AxisVector.AsVector) < 0 Then
                b = True
            Else
                b = False
            End If
            sa = ps.SketchArcs.AddByCenterStartEndPoint(spc, sp1, sp2, b)
            If sa.Length > c.Radius Then
                sa.Delete()
                b = Not b
                sa = ps.SketchArcs.AddByCenterStartEndPoint(spc, sp1, sp2, b)
            End If

            pro = ps.Profiles.AddForSurface(sa)
            ps.Visible = False
            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawOvalSketch(fi As Face, fa As Face) As Profile

        Dim ps As PlanarSketch
        Dim v1, v2, v3, v4 As Vector2d
        Dim skpt3D, skpt2, skpt3, skpt1 As SketchPoint3D
        Dim sp1, sp2, sp3, sp4, spc As SketchPoint
        Dim l12d, l22d, cl1, slme As SketchLine
        Dim skpoints As ObjectCollection = app.TransientObjects.CreateObjectCollection

        Dim gc, gccl As GeometricConstraint
        Dim sen As SketchEntitiesEnumerator
        Dim sa, saf As SketchArc

        Dim dc As DimensionConstraint
        Dim pro As Profile

        Dim pt2D, pt2DCl As Point2d
        Dim d As Double = 0
        Dim i As Integer
        Dim v As Vector
        Dim pl As Plane
        Dim dis As Double
        Dim b As Boolean
        Dim skl As SketchLine3D
        Dim angle As Double
        Dim sa0 As SketchArc
        Dim ssl As SketchSpline
        Try
            comando.WireFrameView(doku)
            Dim ro As Double = 2 / 10
            ps = compDef.Sketches.Add(fi)
            Dim c As Cylinder = fa.Geometry
            sk3D = compDef.Sketches3D.Add
            sk3D.Visible = False
            skpt3D = sk3D.SketchPoints3D.Add(c.BasePoint)
            cylinderCenter = skpt3D
            sp3 = ps.AddByProjectingEntity(skpt3D)
            Dim edMin As Edge = GetMajorArcEdge(fa)
            adjuster = New SketchAdjust(doku)
            Dim se As SketchEntity = ps.AddByProjectingEntity(edMin)
            Try
                ssl = se
                ssl.Construction = True
                sa0 = ps.SketchArcs.AddByCenterStartEndPoint(sp3, ssl.StartSketchPoint, ssl.EndSketchPoint, False)
                If sa0.Length > 1.01 * ssl.Length Then
                    sa0.Delete()
                    sa0 = ps.SketchArcs.AddByCenterStartEndPoint(sp3, ssl.StartSketchPoint, ssl.EndSketchPoint, True)
                End If

            Catch ex As Exception

                slme = se
                slme.Construction = True
                v1 = slme.StartSketchPoint.Geometry.VectorTo(sp3.Geometry)
                v2 = tg.CreateVector2d(-1, 0)
                angle = v1.AngleTo(v2)

                sa0 = ps.SketchArcs.AddByCenterStartSweepAngle(sp3, 0.3 / 10, -angle - Math.PI / 32, Math.PI / 16)

                dc = ps.DimensionConstraints.AddArcLength(sa0, sa0.EndSketchPoint.Geometry)
                'dc = ps.DimensionConstraints.AddRadius(sa0, sa0.EndSketchPoint.Geometry)
                'adjuster.AdjustDimension2DSmothly(dc, 0.3 / 10)
                'dc.Parameter._Value = 0.3 / 10
            End Try

            ps.GeometricConstraints.AddGround(sa0)

            sp1 = sa0.StartSketchPoint

            v1 = sp1.Geometry.VectorTo(sp3.Geometry)
            sp2 = sa0.EndSketchPoint
            v2 = sp2.Geometry.VectorTo(sp3.Geometry)
            v3 = v2
            v3.AddVector(v1)
            pt2DCl = sp3.Geometry
            v4 = v3
            v4.ScaleBy(16)
            pt2DCl.TranslateBy(v4)
            sp4 = ps.SketchPoints.Add(pt2DCl)
            ps.GeometricConstraints.AddGround(sp4)
            cl1 = ps.SketchLines.AddByTwoPoints(sp3, sp4)
            cl1.Construction = True
            angle = v1.AngleTo(v2)
            angleRod = angle
            d = wp1.Point.DistanceTo(wp2.Point)
            If (angle < 1.2 * 0.9) And (d < 45 / 10) Then
                ro = 2 / 10
                smallOval = True
            Else
                smallOval = False
            End If
            pt2D = sp1.Geometry
            pt2D.TranslateBy(v3)
            l12d = ps.SketchLines.AddByTwoPoints(sp1, pt2D)
            gc = ps.GeometricConstraints.AddTangent(sa0, l12d)
            pt2D = sp2.Geometry
            pt2D.TranslateBy(v3)
            l22d = ps.SketchLines.AddByTwoPoints(sp2, pt2D)
            gc = ps.GeometricConstraints.AddTangent(sa0, l22d)
            pt2D = sp3.Geometry
            pt2D.TranslateBy(v3)
            sa = ps.SketchArcs.AddByThreePoints(l12d.EndSketchPoint, pt2D, l22d.EndSketchPoint)
            gccl = ps.GeometricConstraints.AddCoincident(sa.CenterSketchPoint, cl1)
            dc = ps.DimensionConstraints.AddTwoPointDistance(sp4, sa.CenterSketchPoint, DimensionOrientationEnum.kAlignedDim, sa.CenterSketchPoint.Geometry)
            adjuster.AdjustDimensionConstraint2D(dc, dc.Parameter.Value / 16)
            dc.Delete()
            gc = ps.GeometricConstraints.AddTangent(sa, l22d)
            gccl.Delete()
            Try
                gc = ps.GeometricConstraints.AddTangent(sa, l12d)
            Catch ex As Exception
                dc = ps.DimensionConstraints.AddTwoPointDistance(sp3, sa.CenterSketchPoint, DimensionOrientationEnum.kAlignedDim, sa.CenterSketchPoint.Geometry)
                adjuster.AdjustDimensionConstraint2D(dc, dc.Parameter.Value / 4)
                dc.Delete()
                gc = ps.GeometricConstraints.AddTangent(sa, l12d)
            End Try

            dc = ps.DimensionConstraints.AddRadius(sa, pt2D)

            adjuster.AdjustDimensionConstraint2D(dc, ro)
            dc.Parameter._Value = ro
            pro = ps.Profiles.AddForSolid()
            ps.Visible = False
            point1 = sp3.Geometry3d
            skpt1 = sk3D.SketchPoints3D.Add(point1)
            point2 = sa.CenterSketchPoint.Geometry3d
            skpt2 = sk3D.SketchPoints3D.Add(point2)
            lref = sk3D.SketchLines3D.AddByTwoPoints(skpt1, skpt2, False)
            lref.Construction = True
            v = wp1.Point.VectorTo(wp2.Point)
            point3 = skpt2.Geometry
            point3.TranslateBy(v)
            skpt3 = sk3D.SketchPoints3D.Add(point3)
            skl = sk3D.SketchLines3D.AddByTwoPoints(skpt2, skpt3, False)
            ovalLine = skl
            skl.Construction = True
            rodLines.Add(skl)
            comando.RealisticView(doku)
            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetMinorEdge(f As Face) As Edge
        Dim e1, e2, e3 As Edge
        Dim min1, min2, min3 As Double
        min1 = 999999
        min2 = 99999
        min3 = 999999
        e1 = f.Edges.Item(1)
        e2 = e1
        e3 = e2
        For Each ed As Edge In f.Edges
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
    Function DrawPlannSketch(cl As SketchLine3D, le As SketchLine3D) As Profile
        Try
            Dim wpl As WorkPlane
            Dim ps As PlanarSketch
            Dim p1, p2, p3 As SketchPoint
            Dim le2d, cl2d, l32d, mne As SketchLine
            Dim gc, gccc As GeometricConstraint
            Dim r As SketchEntitiesEnumerator
            Dim dc As DimensionConstraint
            Dim spt As SketchPoint
            Dim d As Double = 0
            Dim i As Integer
            Dim sken As SketchEntitiesEnumerator
            wpl = doku.ComponentDefinition.WorkPlanes.AddByThreePoints(cl.EndPoint, cl.StartPoint, le.StartPoint)
            ps = doku.ComponentDefinition.Sketches.Add(wpl)
            lamp.LookAtPlane(wpl)
            le2d = ps.AddByProjectingEntity(le)
            le2d.Construction = True
            cl2d = ps.AddByProjectingEntity(cl)
            cl2d.Construction = True
            mne = ps.AddByProjectingEntity(minorEdge)
            mne.Construction = True


            If minorEdge.GetClosestPointTo(majorEdge.StartVertex.Point).DistanceTo(majorEdge.StartVertex.Point) > minorEdge.GetClosestPointTo(majorEdge.StopVertex.Point).DistanceTo(majorEdge.StopVertex.Point) Then
                l32d = ps.SketchLines.AddByTwoPoints(le2d.StartSketchPoint.Geometry, cl2d.StartSketchPoint.Geometry)
                l32d.Construction = True
                gccc = ps.GeometricConstraints.AddCoincident(l32d.StartSketchPoint, le2d)

            Else
                l32d = ps.SketchLines.AddByTwoPoints(le2d.EndSketchPoint.Geometry, cl2d.EndSketchPoint.Geometry)
                l32d.Construction = True
                gccc = ps.GeometricConstraints.AddCoincident(l32d.StartSketchPoint, le2d)

            End If

            gc = ps.GeometricConstraints.AddPerpendicular(l32d, le2d)
            If mne.EndSketchPoint.Geometry.DistanceTo(l32d.EndSketchPoint.Geometry) > mne.StartSketchPoint.Geometry.DistanceTo(l32d.EndSketchPoint.Geometry) Then
                sken = ps.SketchLines.AddAsThreePointRectangle(l32d.StartSketchPoint.Geometry, l32d.EndSketchPoint.Geometry, mne.EndSketchPoint.Geometry)

                ' dc = ps.DimensionConstraints.AddTwoPointDistance(l32d.EndSketchPoint, mne.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, mne.EndSketchPoint.Geometry)
            Else
                sken = ps.SketchLines.AddAsThreePointRectangle(l32d.StartSketchPoint.Geometry, l32d.EndSketchPoint.Geometry, mne.StartSketchPoint.Geometry)
                ' dc = ps.DimensionConstraints.AddTwoPointDistance(l32d.EndSketchPoint, mne.StartSketchPoint, DimensionOrientationEnum.kAlignedDim, mne.StartSketchPoint.Geometry)
            End If
            i = ps.SketchLines.Count
            If ps.SketchLines.Item(i).Length > ps.SketchLines.Item(i - 1).Length Then

                largo = ps.DimensionConstraints.AddTwoPointDistance(ps.SketchLines.Item(i).StartSketchPoint, ps.SketchLines.Item(i).EndSketchPoint, DimensionOrientationEnum.kAlignedDim, mne.StartSketchPoint.Geometry)
            Else
                largo = ps.DimensionConstraints.AddTwoPointDistance(ps.SketchLines.Item(i - 1).StartSketchPoint, ps.SketchLines.Item(i - 1).EndSketchPoint, DimensionOrientationEnum.kAlignedDim, mne.StartSketchPoint.Geometry)
            End If


            faceProfile = ps.Profiles.AddForSolid
            GetSidePoints(sken, cl2d)
            workAxis.Visible = False
            wpl.Visible = False
            ps.Visible = False
            Return faceProfile
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetSidePoints(sken As SketchEntitiesEnumerator, cl As SketchLine) As WorkPoint

        Dim skl, sklAxis As SketchLine
        sklAxis = cl
        Dim d, e, dmin As Double
        dmin = 9999
        For Each ske As SketchEntity In sken
            If ske.Type = ObjectTypeEnum.kSketchLineObject Then
                skl = ske
                d = skl.Geometry.MidPoint.DistanceTo(cl.Geometry.MidPoint)
                If d < dmin Then
                    dmin = d
                    sklAxis = skl

                End If
            End If
        Next
        d = TorusRadius(sklAxis.StartSketchPoint.Geometry3d)
        e = TorusRadius(sklAxis.EndSketchPoint.Geometry3d)
        If d > e Then
            wp1 = compDef.WorkPoints.AddByPoint(sklAxis.StartSketchPoint)
            wp2 = compDef.WorkPoints.AddByPoint(sklAxis.EndSketchPoint)
        Else
            wp1 = compDef.WorkPoints.AddByPoint(sklAxis.EndSketchPoint)
            wp2 = compDef.WorkPoints.AddByPoint(sklAxis.StartSketchPoint)
        End If
        wp1.Visible = False
        wp2.Visible = False
        Return wp1
    End Function
    Function TorusRadius(pt As Point) As Double

        Return Math.Pow(Math.Pow(Math.Pow(Math.Pow(pt.X, 2) + Math.Pow(pt.Y, 2), 1 / 2) - 43 / 10, 2) + Math.Pow(pt.Z, 2), 1 / 2)
    End Function
    Function RevolveCutTip(pro As Profile) As RevolveFeature

        Try

            RevolveCutTip = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, workAxis, PartFeatureOperationEnum.kCutOperation)
            'RevolveCutTip = doku.ComponentDefinition.Features.RevolveFeatures.AddByFromToExtent(pro, workAxis, cutTipFace1, True, cutTipface2, True, PartFeatureExtentDirectionEnum.kNegativeExtentDirection, True, PartFeatureOperationEnum.kCutOperation)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return RevolveCutTip
    End Function
    Function RevolveFace(pro As Profile) As RevolveFeature
        Dim rf As RevolveFeature
        Try
            rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, workAxis, PartFeatureOperationEnum.kJoinOperation)

        Catch ex As Exception
            Try
                largo.Parameter._Value = largo.Parameter._Value * 1.01
                rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, workAxis, PartFeatureOperationEnum.kJoinOperation)

            Catch ex2 As Exception
                Try
                    largo.Parameter._Value = largo.Parameter._Value / 1.01
                    largo.Driven = True
                    Dim dc1, dc2 As DimensionConstraint
                    Dim ps As PlanarSketch
                    Dim skpt, lpt1, lpt2 As SketchPoint
                    Dim l1 As TwoPointDistanceDimConstraint
                    Dim ls As LineSegment2d
                    Dim gc As GeometricConstraint
                    l1 = largo
                    ps = largo.Parent
                    lpt1 = l1.PointOne
                    lpt2 = l1.PointTwo
                    ls = tg.CreateLineSegment2d(lpt1.Geometry, lpt2.Geometry)
                    skpt = ps.SketchPoints.Add(ls.MidPoint)
                    gc = ps.GeometricConstraints.AddGround(skpt)
                    dc1 = ps.DimensionConstraints.AddTwoPointDistance(lpt1, skpt, DimensionOrientationEnum.kAlignedDim, lpt1.Geometry)
                    Try
                        dc1.Parameter._Value = dc1.Parameter._Value * 1.01
                        rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, workAxis, PartFeatureOperationEnum.kJoinOperation)
                    Catch ex4 As Exception
                        Try
                            dc1.Parameter._Value = dc1.Parameter._Value / 1.02
                            dc2 = ps.DimensionConstraints.AddTwoPointDistance(lpt2, skpt, DimensionOrientationEnum.kAlignedDim, lpt2.Geometry)
                            dc2.Parameter._Value = dc2.Parameter._Value * 1.04
                            doku.Update2(True)
                            Try
                                rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, workAxis, PartFeatureOperationEnum.kJoinOperation)

                            Catch ex5 As Exception
                                dc1.Parameter._Value = dc1.Parameter._Value * 24 / 27
                                dc2.Parameter._Value = dc2.Parameter._Value * 24 / 28
                                doku.Update2(True)
                                rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, workAxis, PartFeatureOperationEnum.kJoinOperation)

                            End Try

                        Catch ex5 As Exception
                            MsgBox(ex.ToString())
                            Return Nothing
                        End Try
                    End Try


                Catch ex3 As Exception
                    MsgBox(ex.ToString())
                    Return Nothing
                End Try

            End Try

        End Try
        If wp2.Point.DistanceTo(rf.SideFaces.Item(1).PointOnFace) < wp2.Point.DistanceTo(rf.SideFaces.Item(2).PointOnFace) Then
            endface = rf.SideFaces.Item(1)

        Else
            endface = rf.SideFaces.Item(2)
        End If
        lamp.HighLighFace(endface)
        endFaceKey = endface.TransientKey
        Return rf
    End Function
    Public Function GetCutTipEndFace() As Face
        Dim pl As Plane = wp2Face.Geometry
        Dim vef As Vector = pl.Normal.AsVector
        Dim vplf As Vector
        Dim d, dMax As Double
        dMax = 0
        For Each sb As SurfaceBody In compDef.SurfaceBodies
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    Try
                        pl = f.Geometry
                        vplf = pl.Normal.AsVector
                        d = vef.DotProduct(vplf)
                        If d > dMax Then
                            dMax = d
                            lamp.HighLighFace(f)
                            GetCutTipEndFace = f
                        End If
                    Catch ex As Exception

                    End Try

                End If

            Next
        Next
        lamp.HighLighFace(GetCutTipEndFace)
        endface = GetCutTipEndFace
        Return GetCutTipEndFace
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
        lamp.HighLighObject(e2)
        lamp.HighLighObject(e1)
        bendEdge = e3
        minorEdge = e2
        majorEdge = e1
        Return e1
    End Function
    Function GetMajorArcEdge(f As Face) As Edge
        Dim e1, e2, e3 As Edge
        Dim maxe1, maxe2, maxe3 As Double
        maxe1 = 0
        maxe2 = 0
        maxe3 = 0
        e1 = f.Edges.Item(1)
        e2 = e1
        e3 = e2
        For Each ed As Edge In f.Edges
            If ed.CurveType = CurveTypeEnum.kBSplineCurve Then
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
            End If


        Next
        'lamp.HighLighObject(e3)
        lamp.HighLighObject(e2)
        lamp.HighLighObject(e1)
        bendEdge = e3
        minorEdge = e2
        majorEdge = e1
        Return e1
    End Function
End Class
