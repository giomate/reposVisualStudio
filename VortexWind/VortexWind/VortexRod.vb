Imports Inventor

Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory

Public Class VortexRod
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile
    Dim windings, driftAngle, passes As Double
    Dim estampa As Stanzer



    Public wp1, wp2, wp3, wpConverge As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim skpt1, skpt2, skpt3 As SketchPoint3D
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Public partNumber, qNext, qLastTie As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres

    Dim cutProfile, faceProfile, rodProfile As Profile
    Dim circle As SketchCircle3D

    Dim pro As Profile
    Dim direction As Vector
    Dim lastWorkPlane, nextWorkPlane, currentWorkPlane As WorkPlane
    Dim lastWorkAxis, nextWorkAxis As WorkAxis

    Public compDef As PartComponentDefinition

    Dim mainWorkPlane As WorkPlane
    Dim workAxis As WorkAxis
    Dim faceRod As Face
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, twistFace As Face
    Dim minorEdge, majorEdge, inputEdge As Edge
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex As DimensionConstraint3D
    Dim largo As DimensionConstraint
    Dim folded As FoldFeature
    Public foldFeatures As FoldFeatures
    Dim features As PartFeatures
    Dim verMax1, verMax2 As Vertex
    Dim lamp As Highlithing
    Dim di As System.IO.DirectoryInfo
    Dim fi As System.IO.File
    Dim nf As System.IO.Path
    Dim pValue, qvalue As Integer

    Dim foldFeature As FoldFeature
    Dim sections, stamPoints, rails, caras, guidePoints As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim spt2dHigh, spt2dLow As SketchPoint
    Dim sptHigh, sptLow, sptHigh2, sptLow2 As SketchPoint3D


    Dim arrayFunctions As Collection
    Dim fullFileNames As String()
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        invFile = New InventorFile(app)
        projectManager = app.DesignProjectManager

        compDef = doku.ComponentDefinition
        features = compDef.Features
        tg = app.TransientGeometry
        sections = app.TransientObjects.CreateObjectCollection
        guidePoints = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        windings = 257
        passes = 109
        driftAngle = 2 * Math.PI * passes / windings
        gap1CM = 3 / 10
        pValue = 0
        qvalue = 0
        done = False
        guidePoints.Clear()
    End Sub
    Function SetConvergePoint(wp As WorkPoint) As WorkPoint
        Try
            wpConverge = wp
            Return wpConverge
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetHighestVertex(sk As Sketch3D) As Vertex
        Dim dMax1 As Double = 0
        Dim dMax2 As Double = 0
        Dim sb As SurfaceBody = doku.ComponentDefinition.SurfaceBodies.Item(1)
        Dim pMax1, pMax2 As Point
        Dim v As Vector
        Dim dMin As Double = 999999
        Dim cpt As Point = compDef.WorkPoints.Item(1).Point
        Dim e As Double = 0
        Try
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    For Each ver As Vertex In f.Vertices
                        If ver.Point.Z > 46 / 10 Then
                            e = ver.Point.Z / ver.Point.DistanceTo(cpt)
                            If (e > dMax2) Then
                                If (e > dMax1) Then
                                    dMax2 = dMax1
                                    dMax1 = e
                                    verMax2 = verMax1
                                    verMax1 = ver
                                    pMax2 = pMax1
                                    pMax1 = ver.Point
                                Else
                                    dMax2 = e
                                    verMax2 = ver
                                    pMax2 = ver.Point
                                End If
                            End If
                        End If

                    Next
                End If

            Next

            skpt1 = sk.SketchPoints3D.Add(pMax1)
            Return verMax1
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DeriveMainPart(s As String) As DerivedPartComponent
        Dim derivedDefinition As DerivedPartDefinition
        Dim dpc As DerivedPartComponent
        Try



            derivedDefinition = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
            derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
            ' derivedDefinition.IncludeAllSurfaces = DerivedComponentOptionEnum.kDerivedExcludeAll
            ' derivedDefinition.IncludeAllParameters = DerivedComponentOptionEnum.kDerivedExcludeAll
            ' derivedDefinition.IncludeBody = True
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
            'derivedDefinition.BodyAsSolidBody = True
            app.SilentOperation = True
            dpc = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)
            app.SilentOperation = False



            Return dpc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetInputEdge(ver As Vertex) As Edge
        Dim ve, vi, vver As Vector
        Dim d As Double = 0
        Dim eMax As Edge = ver.Edges.Item(1)
        Dim e As Double
        Try
            vi = ver.Point.VectorTo(doku.ComponentDefinition.WorkPoints.Item(1).Point)
            vver = verMax1.Point.VectorTo(verMax2.Point)
            For Each ed As Edge In ver.Edges
                ve = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point)
                e = ve.CrossProduct(vi).Length * Math.Abs(ve.DotProduct(vver))
                If ve.CrossProduct(vi).Length > d Then
                    d = ve.CrossProduct(vi).Length
                    eMax = ed
                End If
            Next
            inputEdge = eMax
            lamp.HighLighObject(eMax)
            Return inputEdge
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function GetSixPoints() As SketchPoint3D
        Dim ver As Vertex
        Dim ed As Edge
        Dim ve As Vector
        Dim skpt As SketchPoint3D
        Dim pt As Point

        Try
            ve = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point)
            ve.AsUnitVector.AsVector()
            ve.ScaleBy(0.3 * 2 / 10)
            If ver.Point.DistanceTo(ed.StartVertex.Point) > ver.Point.DistanceTo(ed.StopVertex.Point) Then
                ve.ScaleBy(-1)
            End If
            pt = ver.Point
            guidePoints.Clear()

            For i = 1 To 6
                pt.TranslateBy(ve)
                skpt = sk3D.SketchPoints3D.Add(pt)
                guidePoints.Add(skpt)
            Next
            Return skpt
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawInletLine(skpt As SketchPoint3D) As SketchLine3D
        Dim l, cl As SketchLine3D
        Dim v, vz As Vector
        Dim pt As Point
        Dim gc As GeometricConstraint3D

        Try

            cl = sk3D.SketchLines3D.AddByTwoPoints(wpConverge, doku.ComponentDefinition.WorkPoints.Item(1), False)
            cl.Construction = True
            v = cl.Geometry.Direction.AsVector
            vz = tg.CreateVector(0, 0, 1)
            pt = DrawMainCircle().CenterPoint
            v = vz.CrossProduct(v)
            pt.TranslateBy(v)
            l = sk3D.SketchLines3D.AddByTwoPoints(skpt, pt, False)
            gc = sk3D.GeometricConstraints3D.AddTangent(l, circle)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DocUpdate(docu As PartDocument) As PartDocument
        doku = docu
        compDef = docu.ComponentDefinition
        features = docu.ComponentDefinition.Features
        lamp = New Highlithing(doku)
        monitor = New DesignMonitoring(doku)
        estampa = New Stanzer(doku)
        'adjuster = New SketchAdjust(doku)

        Return doku
    End Function
    Public Function MakeAllWiresGuides(docu As PartDocument) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim w As Parameter


        doku = DocUpdate(docu)
        comando.WireFrameView(doku)

        Try
            Try
                skpt1 = compDef.Sketches3D.Item("HighestPoint").SketchPoints3D.Item(1)
            Catch ex2 As Exception
                sk3D = compDef.Sketches3D.Add()
                sk3D.Name = "HighestPoint"
                If GetHighestVertex(sk3D).Point.Z > 0 Then
                    skpt1 = sk3D.SketchPoints3D.Item(1)
                End If
            End Try


            If compDef.Features.ExtrudeFeatures.Count < 1 Then
                sk3D = compDef.Sketches3D.Add()
                sk3D.Name = "FirstWire"

                ef = RemoveFirstWire(skpt1)
                If monitor.IsFeatureHealthy(ef) Then
                    qvalue = 1
                    w = GetParameter("wq")
                    w._Value = qvalue
                    ef.Name = "rw1"
                    doku.Update2(True)
                    doku.Save2(True)
                    ef = MakeNextWireHole(qvalue)
                End If
            Else
                qvalue = GetParameter("wq")._Value
                ef = MakeNextWireHole(qvalue)
            End If





            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeNextWireHole(q As Integer) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim w As Parameter
        Try
            ef = compDef.Features.ExtrudeFeatures.Item(compDef.Features.ExtrudeFeatures.Count)
            If monitor.IsFeatureHealthy(ef) Then
                lastWorkAxis = compDef.WorkAxes.Item(compDef.WorkAxes.Count)
                For Each wpl As WorkPlane In compDef.WorkPlanes
                    If wpl.Name = String.Concat("wpl", (q - 0).ToString) Then
                        lastWorkPlane = wpl
                        Exit For
                    End If
                Next

                While (compDef.Parameters.Item("wq")._Value < windings + 1 And Not done)
                    ef = RemoveNextWire(lastWorkAxis, lastWorkPlane)
                    If monitor.IsFeatureHealthy(ef) Then
                        qvalue = qvalue + 1
                        w = GetParameter("wq")
                        w._Value = qvalue
                        ef.Name = String.Concat("rw", CInt(w._Value).ToString)
                        doku.Update2(True)
                        doku.Save2(True)
                        lastWorkAxis = compDef.WorkAxes.Item(compDef.WorkAxes.Count)
                        For Each wpl As WorkPlane In compDef.WorkPlanes
                            If wpl.Name = String.Concat("wpl", (qvalue - 0).ToString) Then
                                lastWorkPlane = wpl
                                Exit For
                            End If
                        Next


                    Else
                        done = True
                    End If
                End While

            End If
            Return ef
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function RemoveFirstWire(skpt As SketchPoint3D) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim spt2, spt3 As SketchPoint3D
        Dim pt1, pt2, pt3 As Point
        Dim skl, sklz, sklw As SketchLine3D
        Dim wpl As WorkPlane
        Dim wa As WorkAxis

        Dim v, vz, vpp As Vector

        Try
            pt1 = compDef.WorkPoints.Item(1).Point
            pt2 = skpt.Geometry
            v = pt1.VectorTo(pt2)
            vz = tg.CreateVector(0, 0, 1)
            vpp = v.CrossProduct(vz)
            vpp.Normalize()
            vpp.ScaleBy(25 / 10)
            pt3 = pt1
            pt3.TranslateBy(vpp)
            spt3 = sk3D.SketchPoints3D.Add(pt3)
            skl = sk3D.SketchLines3D.AddByTwoPoints(spt3, skpt, False)
            skl.Construction = True
            pt2 = pt3
            pt2.TranslateBy(vz)
            spt2 = sk3D.SketchPoints3D.Add(pt2)
            wa = compDef.WorkAxes.AddByTwoPoints(spt3, spt2)
            wa.Visible = False
            lastWorkAxis = wa
            wpl = compDef.WorkPlanes.AddByThreePoints(spt3, skpt, spt2)
            wpl.Visible = False
            wpl.Name = "wpl1"
            currentWorkPlane = wpl
            sklw = DrawWireAxis(spt3, wpl)
            ef = RemoveSingleWire(sklw)
            MakeGuideHoles(spt3, sklw)
            guidePoints.Clear()
            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function RemoveNextWire(wai As WorkAxis, wpli As WorkPlane) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim spt2, spt3, spt1 As SketchPoint3D
        Dim pt1, pt2, pt3 As Point
        Dim skl, sklz, sklw As SketchLine3D
        Dim m As Matrix
        Dim wa As WorkAxis
        Dim wpl As WorkPlane
        Dim v, vz, vpp As Vector


        Try
            sk3D = compDef.Sketches3D.Add
            pt1 = compDef.WorkPlanes.Item(3).Plane.IntersectWithLine(wai.Line)
            m = tg.CreateMatrix()
            m.SetToIdentity()
            vz = tg.CreateVector(0, 0, 1)
            m.SetToRotation(driftAngle, vz, compDef.WorkPoints.Item(1).Point)
            pt1.TransformBy(m)
            spt1 = sk3D.SketchPoints3D.Add(pt1)
            pt2 = pt1
            pt2.TranslateBy(vz)
            spt2 = sk3D.SketchPoints3D.Add(pt2)
            wa = compDef.WorkAxes.AddByTwoPoints(spt1, spt2)
            wa.Visible = False
            nextWorkAxis = wa
            wpl = GetNextWorkPlane(wa, wpli)
            wpl.Visible = False
            wpl.Name = String.Concat("wpl", (qvalue + 1).ToString)
            currentWorkPlane = wpl
            sklw = DrawWireAxis(spt1, wpl)
            ef = RemoveSingleWire(sklw)
            MakeGuideHoles(spt1, sklw)
            guidePoints.Clear()
            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetNextWorkPlane(wa As WorkAxis, wpli As WorkPlane) As WorkPlane
        Dim wplo As WorkPlane

        wplo = compDef.WorkPlanes.AddByLinePlaneAndAngle(wa, wpli, driftAngle)
        wplo.Visible = False
        Return wplo
    End Function
    Function DrawWireAxis(skpt As SketchPoint3D, wpl As WorkPlane) As SketchLine3D
        Dim skl As SketchLine3D
        Try
            sk3D.Visible = False
            sk3D = compDef.Sketches3D.Add
            skl = sk3D.SketchLines3D.AddByTwoPoints(GetHighestPointPlane(skpt, wpl), sptLow, False)
            sk3D.Visible = False
            Return skl
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawGuideAxis(skpt As SketchPoint3D, skptc As SketchPoint3D, skli As SketchLine3D) As SketchLine3D
        Dim skl, cl As SketchLine3D
        Dim gc As GeometricConstraint3D
        Try
            sk3D.Visible = False
            sk3D = compDef.Sketches3D.Add()
            skl = sk3D.SketchLines3D.AddByTwoPoints(skpt, skptc.Geometry, False)
            cl = sk3D.Include(skli)
            cl.Construction = True
            gc = sk3D.GeometricConstraints3D.AddParallel(skl, cl)
            sk3D.Visible = False
            Return skl
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Function GetHighestPointPlane(skpt As SketchPoint3D, wpl As WorkPlane) As SketchPoint3D
        Dim dMax1, dMax2, dMin1, dMin2, e, dis As Double
        Dim pt, ptMax1, ptMax2, ptMin1, ptMin2 As Point
        dMax1 = 0
        dMax2 = dMax1
        Dim f1, f2, b1, b2 As Face
        Dim sb As SurfaceBody = compDef.SurfaceBodies.Item(1)
        Dim vnp, vpt, vr As Vector
        Dim cpt As Point = compDef.WorkPoints.Item(1).Point
        Dim spt As SketchPoint3D
        Dim ic As IntersectionCurve

        Try
            vnp = wpl.Plane.Normal.AsVector
            ptMax1 = skpt.Geometry
            ptMin1 = ptMax1
            ptMin2 = ptMin1
            ptMax2 = ptMax1
            f1 = sb.Faces.Item(1)
            f2 = f1
            b1 = f1
            b2 = f2
            For Each f As Face In sb.Faces
                If f.Evaluator.Area > 0.42 Then
                    If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        Try
                            pt = f.GetClosestPointTo(skpt.Geometry)
                            dis = wpl.Plane.DistanceTo(pt)
                            If Math.Abs(dis) < 13 / 10 Then
                                ' lamp.HighLighFace(f)
                                If pt.Z > 28 / 10 Then
                                    vpt = skpt.Geometry.VectorTo(pt)
                                    vr = vpt.CrossProduct(vnp)
                                    If vr.Z < 0 Then
                                        vr = vr.CrossProduct(vpt)
                                        vr.Normalize()
                                        vpt.Normalize()
                                        e = vr.DotProduct(vnp)
                                        If e > 0.9 Then
                                            Try
                                                ic = sk3D.IntersectionCurves.Add(wpl, f)
                                                For Each se As SketchEntity3D In ic.SketchEntities
                                                    se.Construction = True
                                                    If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                                                        spt = se
                                                        pt = spt.Geometry
                                                        e = pt.Z / pt.DistanceTo(skpt.Geometry)
                                                        If (e > dMax2) Then
                                                            If (e > dMax1) Then
                                                                dMax2 = dMax1
                                                                dMax1 = e
                                                                f2 = f1
                                                                f1 = f
                                                                ptMax2 = ptMax1
                                                                ptMax1 = pt
                                                                lamp.HighLighObject(se)
                                                            Else
                                                                dMax2 = e
                                                                f2 = f
                                                                ptMax2 = pt
                                                            End If
                                                        End If
                                                    End If

                                                Next
                                            Catch ex As Exception

                                            End Try



                                        End If
                                    End If
                                Else
                                    If pt.Z < -22 / 10 Then
                                        vpt = skpt.Geometry.VectorTo(pt)
                                        vr = vpt.CrossProduct(vnp)
                                        lamp.HighLighFace(f)
                                        If vr.Z > 0 Then
                                            vr = vr.CrossProduct(vpt)
                                            vr.Normalize()
                                            vpt.Normalize()
                                            e = vr.DotProduct(vnp)
                                            If e > 0.9 Then
                                                Try
                                                    'lamp.HighLighFace(f)
                                                    ic = sk3D.IntersectionCurves.Add(wpl, f)
                                                    For Each se As SketchEntity3D In ic.SketchEntities
                                                        se.Construction = True
                                                        If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                                                            spt = se
                                                            pt = spt.Geometry
                                                            e = pt.Z / pt.DistanceTo(skpt.Geometry)
                                                            If (e < dMin2) Then
                                                                If (e < dMin1) Then
                                                                    dMin2 = dMin1
                                                                    dMin1 = e
                                                                    b2 = b1
                                                                    b1 = f
                                                                    ptMin2 = ptMin1
                                                                    ptMin1 = pt
                                                                    lamp.HighLighObject(se)
                                                                Else
                                                                    dMin2 = e
                                                                    b2 = f
                                                                    ptMin2 = pt
                                                                End If
                                                            End If
                                                        End If

                                                    Next
                                                Catch ex As Exception

                                                End Try
                                            End If
                                        End If


                                    End If
                                End If

                            End If
                        Catch ex As Exception

                        End Try

                    End If
                End If


            Next
            sptLow = sk3D.SketchPoints3D.Add(ptMin1)
            point2 = ptMin2

            sptLow = DrawEntryPoint(b1, wpl, sptLow, skpt)

            sptHigh = sk3D.SketchPoints3D.Add(ptMax1)
            point2 = ptMax2

            Return DrawEntryPoint(f1, wpl, sptHigh, skpt)



        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetLowerstPointPlane(skpt As SketchPoint3D, wpl As WorkPlane) As SketchPoint3D
        Dim dMin1, dMin2, e, dis As Double
        Dim pt, ptMin1, ptMin2 As Point
        Dim f1, f2 As Face
        dMin1 = 99999999
        dMin2 = dMin1
        Dim sb As SurfaceBody = compDef.SurfaceBodies.Item(1)
        Dim vnp, vpt, vr As Vector
        Dim cpt As Point = compDef.WorkPoints.Item(1).Point
        Dim pl As Plane
        Dim oe As ObjectsEnumerator
        Dim ic As IntersectionCurve
        Dim ea As SketchEllipticalArc3D
        Dim spt As SketchPoint3D
        Try
            vnp = wpl.Plane.Normal.AsVector
            ptMin1 = skpt.Geometry
            f1 = sb.Faces.Item(1)
            f2 = f1
            For Each f As Face In sb.Faces
                If f.Evaluator.Area > 0.25 Then
                    If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        Try
                            pt = f.GetClosestPointTo(skpt.Geometry)
                            dis = wpl.Plane.DistanceTo(pt)
                            If Math.Abs(dis) < 1 / 10 Then
                                lamp.HighLighFace(f)
                                If pt.Z < -36 / 10 Then
                                    vpt = skpt.Geometry.VectorTo(pt)
                                    vr = vpt.CrossProduct(vnp)
                                    If vr.Z > 0 Then
                                        vr = vr.CrossProduct(vpt)
                                        vr.Normalize()
                                        vpt.Normalize()
                                        e = vr.DotProduct(vnp)
                                        If e > 0.9 Then
                                            Try
                                                lamp.HighLighFace(f)
                                                ic = sk3D.IntersectionCurves.Add(wpl, f)
                                                For Each se As SketchEntity3D In ic.SketchEntities
                                                    se.Construction = True
                                                    If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                                                        spt = se
                                                        pt = spt.Geometry
                                                        e = pt.Z / pt.DistanceTo(skpt.Geometry)
                                                        If (e < dMin2) Then
                                                            If (e < dMin1) Then
                                                                dMin2 = dMin1
                                                                dMin1 = e
                                                                f2 = f1
                                                                f1 = f
                                                                ptMin2 = ptMin1
                                                                ptMin1 = pt
                                                                lamp.HighLighObject(se)
                                                            Else
                                                                dMin2 = e
                                                                f2 = f
                                                                ptMin2 = pt
                                                            End If
                                                        End If
                                                    End If

                                                Next
                                            Catch ex As Exception

                                            End Try
                                        End If
                                    End If
                                Else

                                End If

                            End If
                        Catch ex As Exception

                        End Try

                    End If
                End If


            Next
            sptLow = sk3D.SketchPoints3D.Add(ptMin1)

            Return DrawEntryPoint(f1, wpl, skpt, sptLow)



        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetRadiusPoint(pt As Point) As Double
        Dim radius As Double = Math.Pow(Math.Pow(pt.X, 2) + Math.Pow(pt.Y, 2), 1 / 2)
        Return radius
    End Function
    Function DrawEntryPoint(f As Face, wpl As WorkPlane, sptFace As SketchPoint3D, sptCenter As SketchPoint3D) As SketchPoint3D
        Dim ic As IntersectionCurve
        Dim spt, spt2 As SketchPoint3D
        Dim pt, ptMax1, ptMax2, pttf As Point
        Dim f1, f2 As Face
        Dim e As Double = 0
        Dim dMax1, dMax2, dis1, dis2 As Double
        dMax1 = 0
        dMax2 = 0
        ptMax1 = sptFace.Geometry
        ptMax2 = point2

        dis2 = Math.Abs(sptFace.Geometry.Z) / ((GetRadiusPoint(sptFace.Geometry)))
        f1 = f
        Try
            For Each tf As Face In f.TangentiallyConnectedFaces

                pttf = tf.GetClosestPointTo(sptCenter.Geometry)

                If (Math.Abs(sptFace.Geometry.Z - pttf.Z)) < 25 / 10 Then
                    dis1 = Math.Abs(pttf.Z) / (GetRadiusPoint(pttf))
                    sk3D.SketchPoints3D.Add(pttf)
                    If dis1 > dis2 Then

                        lamp.HighLighFace(tf)
                        Try
                            ic = sk3D.IntersectionCurves.Add(wpl, tf)
                            For Each se As SketchEntity3D In ic.SketchEntities
                                se.Construction = True
                                If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                                    spt = se
                                    pt = spt.Geometry
                                    e = Math.Abs(pt.Z) / (GetRadiusPoint(pt))
                                    If (e > dMax2) Then
                                        If (e > dMax1) Then
                                            dMax2 = dMax1
                                            dMax1 = e
                                            ptMax2 = ptMax1
                                            ptMax1 = pt
                                            f2 = f1
                                            f1 = tf
                                            lamp.HighLighObject(se)
                                            dis2 = e
                                        Else
                                            dMax2 = e
                                            f2 = tf
                                            ptMax2 = pt
                                        End If
                                    End If
                                End If

                            Next
                        Catch ex As Exception

                        End Try
                    End If
                End If


            Next
            spt = sk3D.SketchPoints3D.Add(ptMax1)
            spt2 = sk3D.SketchPoints3D.Add(ptMax2)
            estampa.EmbossLetter(qvalue, DrawStampPoint(f1, wpl, sptFace), f1)
            'DrawGuidePoints(spt, spt2)
            Return spt
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawStampPoint(f As Face, wpl As WorkPlane, sptFace As SketchPoint3D) As SketchPoint3D
        Dim ic As IntersectionCurve
        Dim sk As Sketch3D = compDef.Sketches3D.Add
        Dim spt, spt2 As SketchPoint3D
        Dim pt, ptMax1, ptMax2, pttf, ptMaxZ As Point
        Dim f1, f2 As Face
        Dim e As Double = 0
        Dim dMax1, dMax2, dis1, dis2, dMin, d As Double
        dMax1 = 0
        dMax2 = 0
        dMin = 99999
        ptMax1 = sptFace.Geometry
        ptMax2 = point2
        Dim tc As Circle

        tc = tg.CreateCircle(compDef.WorkPoints.Item(1), tg.CreateUnitVector(0, 0, 1), 100 / 10)

        For Each ptz As Point In wpl.Plane.IntersectWithCurve(tc)
            d = ptz.DistanceTo(sptFace.Geometry)
            If d < dMin Then
                dMin = d
                ptMaxZ = ptz

            End If
        Next



        dis2 = (GetRadiusPoint(sptFace.Geometry) / (Math.Abs(sptFace.Geometry.Z) + 0.000001))
        f1 = f
        Try
            For Each tf As Face In f.TangentiallyConnectedFaces
                If tf.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    pttf = tf.GetClosestPointTo(ptMaxZ)
                    dis1 = (GetRadiusPoint(pttf) / (Math.Abs(pttf.Z) + 0.000001))
                    sk.SketchPoints3D.Add(pttf)
                    If dis1 > dis2 Then

                        lamp.HighLighFace(tf)
                        Try
                            ic = sk.IntersectionCurves.Add(wpl, tf)
                            For Each se As SketchEntity3D In ic.SketchEntities
                                se.Construction = True
                                If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                                    spt = se
                                    pt = spt.Geometry
                                    e = (GetRadiusPoint(pt) / (Math.Abs(pt.Z) + 0.000001))
                                    If (e > dMax2) Then
                                        If (e > dMax1) Then
                                            dMax2 = dMax1
                                            dMax1 = e
                                            ptMax2 = ptMax1
                                            ptMax1 = pt
                                            f2 = f1
                                            f1 = tf
                                            lamp.HighLighObject(se)
                                            dis2 = e
                                        Else
                                            dMax2 = e
                                            f2 = tf
                                            ptMax2 = pt
                                        End If
                                    End If
                                End If

                            Next
                        Catch ex As Exception

                        End Try

                    End If
                End If
            Next
            spt = sk.SketchPoints3D.Add(ptMax1)
            spt2 = sk.SketchPoints3D.Add(ptMax2)
            sk.Visible = False
            Return spt
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawGuidePoints(skpt1 As SketchPoint3D, skpt2 As SketchPoint3D) As SketchPoint3D
        Dim spt As SketchPoint3D
        Dim sk As Sketch3D = compDef.Sketches3D.Add
        Dim ve, vz, vr As Vector
        Dim pt As Point
        Dim modulo As Integer
        Dim d As Double
        Try

            vz = tg.CreateVector(0, 0, 1)
            ve = skpt1.Geometry.VectorTo(skpt2.Geometry)
            ve.Normalize()
            d = Math.Abs(ve.DotProduct(vz))
            If d > 0.5 Then
                If skpt1.Geometry.Z > 0 Then
                    vr = vz.CrossProduct(currentWorkPlane.Plane.Normal.AsVector)
                Else
                    vr = currentWorkPlane.Plane.Normal.AsVector.CrossProduct(vz)

                End If

                vr.Normalize()
                ve = vr

            End If
            ve.ScaleBy(0.6 * 2 / 10)
            pt = skpt1.Geometry

            If skpt1.Geometry.Z > 0 Then
                Math.DivRem(qvalue + 1, 5, modulo)
            Else
                Math.DivRem(qvalue, 5, modulo)
            End If

            If modulo > 0 Then
                For i = 1 To modulo
                    pt.TranslateBy(ve)
                    spt = sk.SketchPoints3D.Add(pt)
                    guidePoints.Add(spt)

                Next

            End If
            sk.Visible = False
            Return spt



        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Function MakeGuideHoles(skptc As SketchPoint3D, skli As SketchLine3D) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim t As Integer = guidePoints.Count
        Dim skpto As SketchPoint3D
        Try
            If t > 0 Then
                For i = 1 To t
                    skpto = guidePoints.Item(i)

                    ef = MakeSingleHole(DrawGuideAxis(skpto, skptc, skli))

                Next
            End If


            Return ef
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function RemoveSingleWire(skl As SketchLine3D) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim ed As ExtrudeDefinition
        Dim pro As Profile = DrawRemoveCircle(skl)
        Try
            ed = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
            ed.SetThroughAllExtent(PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
            ef = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(ed)
            Return ef
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeSingleHole(skl As SketchLine3D) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim ed As ExtrudeDefinition
        Dim pro As Profile = DrawGuideHole(skl)
        Try
            ed = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
            ed.SetDistanceExtent(6 / 10, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
            ef = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(ed)
            If Not monitor.IsFeatureHealthy(ef) Then
                ef.Definition.SetDistanceExtent(24 / 10, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
                ' ed.SetDistanceExtent(25 / 10, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
                ' ef = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(ed)
            End If
            Return ef
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawRemoveCircle(skl As SketchLine3D) As Profile
        Dim pro As Profile
        Dim ps As PlanarSketch
        Dim wpl As WorkPlane
        Dim spt As SketchPoint
        Try
            wpl = doku.ComponentDefinition.WorkPlanes.AddByNormalToCurve(skl, skl.EndPoint)
            wpl.Visible = False
            ps = doku.ComponentDefinition.Sketches.Add(wpl)
            spt = ps.AddByProjectingEntity(skl.EndPoint)
            ps.SketchCircles.AddByCenterRadius(spt, 0.5 / 10)
            pro = ps.Profiles.AddForSolid

            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawGuideHole(skl As SketchLine3D) As Profile
        Dim pro As Profile
        Dim ps As PlanarSketch
        Dim wpl As WorkPlane
        Dim spt As SketchPoint
        Try
            wpl = doku.ComponentDefinition.WorkPlanes.AddByNormalToCurve(skl, skl.StartPoint)
            wpl.Visible = False
            ps = doku.ComponentDefinition.Sketches.Add(wpl)
            spt = ps.AddByProjectingEntity(skl.EndPoint)
            ps.SketchCircles.AddByCenterRadius(spt, 0.3 / 10)
            pro = ps.Profiles.AddForSolid

            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetLowerstPoint() As Point
        Dim dMin As Double = 99999
        Dim sb As SurfaceBody = doku.ComponentDefinition.SurfaceBodies.Item(1)
        Dim pMax As Point
        Try
            For Each f As Face In sb.Faces
                For Each v As Vertex In f.Vertices
                    If v.Point.Z < dMin Then
                        dMin = v.Point.Z
                        pMax = v.Point

                    End If
                Next
            Next
            skpt2 = sk3D.SketchPoints3D.Add(pMax)
            Return pMax
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Function DrawMainCircle() As SketchCircle3D

        Dim v As Vector
        Dim skpt
        Try
            skpt = sk3D.SketchPoints3D.Add(doku.ComponentDefinition.WorkPoints.Item(1).Point)
            v = tg.CreateVector(0, 0, 1)
            circle = sk3D.SketchCircles3D.AddByCenterRadius(skpt, v.AsUnitVector, 50 / 10)

            Return circle
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function GetParameter(name As String) As Parameter
        Dim p As Parameter
        Try
            p = compDef.Parameters.ModelParameters.Item(name)
        Catch ex As Exception
            Try
                p = compDef.Parameters.ReferenceParameters.Item(name)
            Catch ex1 As Exception
                Try
                    p = compDef.Parameters.UserParameters.Item(name)
                Catch ex2 As Exception
                    MsgBox(ex2.ToString())
                    MsgBox("Parameter not found: " & name)
                End Try

            End Try

        End Try

        Return p
    End Function
End Class
