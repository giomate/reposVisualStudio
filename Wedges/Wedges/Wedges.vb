Imports Inventor

Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory

Public Class Wedges
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, connectLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile

    Public trobinaCurve As Curves3D
    Dim palito As RodMaker

    Public wp1, wp2, wp3, wpConverge As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, convergePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Public partNumber, qNext, qLastTie As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Dim nextSketch As OriginSketch
    Dim cutProfile As Profile


    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Public compDef As ComponentDefinition
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim mainWorkPlane As WorkPlane
    Dim workPointFace As WorkPoint
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, cutEdge1, cutEdge2, CutEsge3, closestEdge As Edge
    Dim minorLine, majorLine, cutLine3D, kante3D, tante3D As SketchLine3D
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, cylinderFace As Face
    Dim workFaces As FaceCollection
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex As DimensionConstraint3D
    Dim folded As FoldFeature
    Public foldFeatures As FoldFeatures
    Dim features As PartFeatures
    Dim lamp As Highlithing
    Dim di As System.IO.DirectoryInfo
    Dim fi As System.IO.File
    Dim nf As System.IO.Path
    Dim bandFaces As WorkSurface

    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails, caras, surfaceBodies As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim refDoc As FindReferenceLine

    Dim manager As FoldingEvaluator
    Dim giro As TwistFold7
    Dim arrayFunctions As Collection
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
        projectManager = app.DesignProjectManager

        compDef = doku.ComponentDefinition

        features = doku.ComponentDefinition.Features
        sheetMetalFeatures = compDef.Features
        foldFeatures = sheetMetalFeatures.FoldFeatures
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        sections = app.TransientObjects.CreateObjectCollection
        esquinas = app.TransientObjects.CreateObjectCollection
        rails = app.TransientObjects.CreateObjectCollection
        caras = app.TransientObjects.CreateObjectCollection
        surfaceBodies = app.TransientObjects.CreateObjectCollection

        lamp = New Highlithing(doku)
        thicknessCM = compDef.Thickness._Value
        gap1CM = 3 / 10
        refDoc = New FindReferenceLine(doku)
        nombrador = New Nombres(doku)
        manager = New FoldingEvaluator(doku)
        trobinaCurve = New Curves3D(doku)
        palito = New RodMaker(doku)
        DP.Dmax = 200 / 10
        DP.Dmin = 1 / 10
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 13
        DP.q = 31
        DP.b = 25

        done = False
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
                        RemoveExcessFaces(doku.ComponentDefinition.WorkSurfaces.Item(1))
                        If monitor.IsFeatureHealthy(EmbossNumber(s)) Then
                            doku.SaveAs(nombrador.MakeWedgeFileName(s), False)
                            doku.Close(True)
                            app.Documents.CloseAll()

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
                        v.ScaleBy(f.Evaluator.Area * 3)
                        pt = wpf.Point
                        pt.TranslateBy(v)
                        wpt = doku.ComponentDefinition.WorkPoints.AddFixed(pt)
                        ef = RemoveFakeMaterial(f, wpt)
                        If monitor.IsFeatureHealthy(ef) Then
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
    Function GetWorkFace() As Face
        Try
            Dim maxArea1, maxArea2, maxArea3 As Double
            Dim sb As SurfaceBody
            Dim maxface1, maxface2, maxface3 As Face
            Dim fc As FaceCollection
            Dim b, b2 As Boolean
            Dim d1, d2, aMax, eMax As Double
            Dim v1, v2 As Vector
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
    Function GetClosestEdge(f As Face) As Edge
        Dim e1, e2, e3 As Edge
        Dim mine1, mine2, mine3 As Double
        mine1 = 99999
        mine2 = 9999999
        mine3 = 999999
        e1 = f.EdgeLoops.Item(1).Edges.Item(1)
        e2 = e1
        e3 = e2
        For Each ed As Edge In f.EdgeLoops.Item(1).Edges
            If ed.GetClosestPointTo(wpConverge.Point).DistanceTo(wpConverge.Point) < mine2 Then

                If ed.GetClosestPointTo(wpConverge.Point).DistanceTo(wpConverge.Point) < mine1 Then
                    mine3 = mine2
                    e3 = e2
                    mine2 = mine1
                    e2 = e1
                    mine1 = ed.GetClosestPointTo(wpConverge.Point).DistanceTo(wpConverge.Point)
                    e1 = ed
                Else
                    mine3 = mine2
                    e3 = e2
                    mine2 = ed.GetClosestPointTo(wpConverge.Point).DistanceTo(wpConverge.Point)
                    e2 = ed
                End If
            Else
                mine3 = ed.GetClosestPointTo(wpConverge.Point).DistanceTo(wpConverge.Point)
                e3 = ed



            End If

        Next
        lamp.HighLighObject(e1)

        closestEdge = e1
        Return e1
    End Function
    Function ExtrudeNumber(pro As Profile) As ExtrudeFeature
        Dim oExtrudeDef As ExtrudeDefinition
        oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
        oExtrudeDef.SetDistanceExtent(0.12, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
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
            Conversions.SetUnitsToMetric(p)
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
    Function FillVoids() As PartDocument
        Dim ilf As Integer
        Dim wp1, wp2 As WorkPoint
        Dim sb As SurfaceBody
        Dim pli, plj As Plane
        Dim ls As LineSegment
        Dim pti, pti2, ptj, ptj2, ptm, ptmi, ptmj, ptcj As Point
        Dim ptinside(3) As Double
        Dim d, aMax As Double
        Dim limit As Integer = 0
        sb = doku.ComponentDefinition.SurfaceBodies.Item(1)
        ilf = sb.Faces.Count

        aMax = 0
        Try

            For Each f As Face In sb.Faces
                If f.Evaluator.Area > aMax Then
                    aMax = f.Evaluator.Area

                End If
            Next
            For i = 1 To ilf - 1
                If sb.Faces.Item(i).Evaluator.Area > aMax * 1 / 3 Then
                    lamp.HighLighFace(sb.Faces.Item(i))
                    For j = i + 1 To ilf

                        If sb.Faces.Item(j).Evaluator.Area > aMax * 1 / 3 Then
                            lamp.HighLighFace(sb.Faces.Item(j))
                            wp1 = doku.ComponentDefinition.WorkPoints.AddAtCentroid(sb.Faces.Item(i).EdgeLoops.Item(1))
                            pti = wp1.Point
                            pti2 = sb.Faces.Item(i).GetClosestPointTo(pti)
                            wp1.Delete()
                            pti = pti2
                            wp1 = doku.ComponentDefinition.WorkPoints.AddFixed(pti)
                            wp2 = doku.ComponentDefinition.WorkPoints.AddAtCentroid(sb.Faces.Item(j).EdgeLoops.Item(1))
                            ptj = wp2.Point
                            ptj2 = sb.Faces.Item(j).GetClosestPointTo(ptj)
                            wp2.Delete()
                            ptj = ptj2
                            wp2 = doku.ComponentDefinition.WorkPoints.AddFixed(ptj)

                            If pti.DistanceTo(ptj) < 4 * gap1CM Then
                                ptcj = sb.Faces.Item(j).GetClosestPointTo(pti)
                                If ptcj.DistanceTo(pti) < gap1CM Then
                                    wp2.Delete()
                                    ptj = ptcj
                                    wp2 = doku.ComponentDefinition.WorkPoints.AddFixed(ptj)
                                    ls = tg.CreateLineSegment(pti, ptj)
                                    ptm = ls.MidPoint
                                    If Not IsPointContained(ptm, sb) Then
                                        ls = tg.CreateLineSegment(pti, ptm)
                                        ptmi = ls.MidPoint
                                        ls = tg.CreateLineSegment(pti, ptmi)
                                        ptmi = ls.MidPoint
                                        If Not IsPointContained(ptmi, sb) Then
                                            ls = tg.CreateLineSegment(ptj, ptm)
                                            ptmj = ls.MidPoint
                                            ls = tg.CreateLineSegment(ptj, ptmj)
                                            ptmj = ls.MidPoint
                                            If Not IsPointContained(ptmj, sb) Then
                                                pli = sb.Faces.Item(i).Geometry
                                                plj = sb.Faces.Item(j).Geometry
                                                d = pli.Normal.AsVector.DotProduct(plj.Normal.AsVector)
                                                If Math.Abs(d) > 0.5 Then
                                                    'lamp.HighLighFace(sb.Faces.Item(i))
                                                    'lamp.HighLighFace(sb.Faces.Item(j))
                                                    Try
                                                        If monitor.IsFeatureHealthy(TryLoft(sb.Faces.Item(i), sb.Faces.Item(j))) Then
                                                            i = 1
                                                            sb = doku.ComponentDefinition.SurfaceBodies.Item(1)
                                                            ilf = sb.Faces.Count
                                                            For Each f As Face In sb.Faces
                                                                If f.Evaluator.Area > aMax Then
                                                                    aMax = f.Evaluator.Area
                                                                End If
                                                            Next
                                                        End If


                                                    Catch ex As Exception
                                                        doku.ComponentDefinition.Sketches.Item(doku.ComponentDefinition.Sketches.Count).Visible = False
                                                        doku.ComponentDefinition.Sketches.Item(doku.ComponentDefinition.Sketches.Count - 1).Visible = False
                                                    End Try

                                                End If
                                            End If
                                            End If

                                    End If
                                End If

                            End If
                            wp1.Delete()
                            wp2.Delete()

                        End If
                    Next
                End If
                sb = doku.ComponentDefinition.SurfaceBodies.Item(1)
                ilf = sb.Faces.Count
            Next




        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return doku
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
            workfaces = fmin1.TangentiallyConnectedFaces
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
        'GetRails(fc)
        sections.Add(pr)
        Dim pt As Point
        Dim pl As Plane
        pl = fc.Geometry
        Dim v As Vector
        v = pl.Normal.AsVector
        v.ScaleBy(-1 * fc.Evaluator.Area / 16)
        pt = convergePoint
        pt.TranslateBy(v)
        Dim wptlf As WorkPoint = doku.ComponentDefinition.WorkPoints.AddFixed(pt)
        wptlf.Visible = False
        sections.Add(wptlf)
        Return LoftSingleSpike()
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
            lf = compDef.Features.LoftFeatures.Add(oLoftDefinition)
            Return lf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


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
