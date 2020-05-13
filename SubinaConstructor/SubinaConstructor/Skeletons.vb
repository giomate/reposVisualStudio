Imports Inventor

Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory
Public Class Skeletons
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim ringLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile

    Dim trobinaCurve As Curves3D
    Dim palitos As RodMaker
    Dim conos As Wedges
    Dim barrido As Sweeper

    Dim wp1, wp2, wp3, wpConverge, wptHigh, wptLow As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, convergePoint As Point
    Dim skptHigh, skptLow As SketchPoint3D
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Public partNumber, qNext, qLastTie, qvalue As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Dim caraTrabajo As Surfacer
    Dim estampa As Stanzer
    Dim adjuster As SketchAdjust

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
    Dim CutEsge3, closestEdge, stampEdge As Edge
    Dim minorLine, majorLine, cutLine3D, kante3D, stampLine As SketchLine3D
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, cylinderFace, stampFace As Face
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
    Dim bandFaces, tangentSurfaces As WorkSurface

    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails, caras, surfaceBodies As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane


    Dim arrayFunctions As Collection
    Dim garras As VortexRod
    Dim fullFileNames As String()
    Structure DesignParam
        Public p As Integer
        Public q As Integer
        Public b As Double
        Public Dmax As Double
        Public Dmin As Double

    End Structure
    Dim DP As DesignParam
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


        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        sections = app.TransientObjects.CreateObjectCollection
        esquinas = app.TransientObjects.CreateObjectCollection
        rails = app.TransientObjects.CreateObjectCollection
        caras = app.TransientObjects.CreateObjectCollection
        surfaceBodies = app.TransientObjects.CreateObjectCollection

        lamp = New Highlithing(doku)

        gap1CM = 3 / 10

        nombrador = New Nombres(doku)

        trobinaCurve = New Curves3D(doku)
        palitos = New RodMaker(doku)
        DP.Dmax = 200 / 10
        DP.Dmin = 1 / 10
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 11
        DP.q = 23
        DP.b = 25

        done = False
    End Sub
    Public Function MakeSkeletonIteration(i As Integer) As PartDocument
        Try
            Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
            Dim q As Integer
            Dim nd As String
            Dim ef As ExtrudeFeature
            Dim rf As RevolveFeature

            nd = String.Concat(p, "\Iteration", i.ToString)

            If (Directory.Exists(nd)) Then

                fullFileNames = Directory.GetFiles(nd, "*.ipt")


                For Each s As String In fullFileNames
                    If nombrador.ContainBand(s) Then
                        If Not IsSkeletonDone(nombrador.MakeSkeletonFileName(s)) Then
                            doku = MakeSingleSkeleton(s)
                            monitor = New DesignMonitoring(doku)
                            doku.Update2(True)
                            If monitor.IsFeatureHealthy(MakeFrameLetters(s)) Then

                                comando.RealisticView(doku)
                                If monitor.IsFeatureHealthy(RemoveExcessMaterial(s)) Then
                                    comando.RealisticView(doku)
                                    If monitor.IsFeatureHealthy(MakeTwoHoles()) Then
                                        comando.RealisticView(doku)
                                        SaveAsSilent(s)
                                        DokuClose()
                                    End If
                                End If
                            End If
                        Else


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
    Public Function RecoverSkeletonIteration(i As Integer) As PartDocument
        Try
            Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
            Dim q As Integer
            Dim nd As String
            Dim ef As ExtrudeFeature
            Dim rf As RevolveFeature
            Dim cf As CombineFeature
            Dim s As String
            nd = String.Concat(p, "\Iteration", i.ToString)

            If (Directory.Exists(nd)) Then

                s = String.Concat(nd, "\Band8.ipt")

                Try
                    doku = app.Documents.Open(String.Concat(nd, "\SkeletonTest.ipt"))
                    DocUpdate(doku)
                    ef = compDef.Features.ExtrudeFeatures.Item("last")
                    DokuClose()
                Catch ex As Exception

                    If compDef.Features.LoftFeatures.Count > 2 Then
                        wptHigh = compDef.WorkPoints.Item("wptHigh")
                        wptLow = compDef.WorkPoints.Item("wptLow")
                        Try
                            If monitor.IsFeatureHealthy(RemoveExcessMaterial(s)) Then
                                comando.RealisticView(doku)
                                If monitor.IsFeatureHealthy(MakeTwoHoles()) Then
                                    comando.RealisticView(doku)
                                    SaveAsSilent(s)
                                    DokuClose()
                                End If
                            End If
                        Catch ex2 As Exception
                            If monitor.IsFeatureHealthy(MakeTwoHoles()) Then
                                If monitor.IsFeatureHealthy(RemoveExcessMaterial(s)) Then
                                    DokuClose()
                                End If
                            End If
                        End Try
                    End If
                End Try

            End If

            Return doku
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DokuClose() As Integer
        SaveSilent()
        doku.Close()
        app.Documents.CloseAll()
        Return 0
    End Function
    Function RemoveExcessMaterial(s As String) As ExtrudeFeature
        Try
            barrido = New Sweeper(doku)
            RemoveExcessMaterial = barrido.RemoveAll(s)
            RemoveExcessMaterial.Name = "last"
            Return RemoveExcessMaterial
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function SaveSilent() As PartDocument
        app.SilentOperation = True
        doku.Save()

        app.SilentOperation = False
        Return doku
    End Function
    Function SaveAsSilent(s As String) As PartDocument
        app.SilentOperation = True
        doku.SaveAs(nombrador.MakeSkeletonFileName(s), False)
        doku.Save()

        app.SilentOperation = False
        Return doku
    End Function
    Function MakeFrameLetters(s As String) As ExtrudeFeature
        estampa = New Stanzer(doku)
        Dim ef As ExtrudeFeature = estampa.ExtrudeFrameLetter(s, wptHigh, GetStampLine(wptHigh))
        ef.Name = "LetterHigh"
        If monitor.IsFeatureHealthy(ef) Then
            ef = estampa.ExtrudeFrameLetter(s, wptLow, GetStampLine(wptLow))
            ef.Name = "LetterLow"
            If monitor.IsFeatureHealthy(ef) Then
                Return ef

            End If
        End If
        Return Nothing
    End Function
    Function GetStampFace(wpt As WorkPoint) As Face

        Dim sb As SurfaceBody = tangentSurfaces._SurfaceBody
        Dim d, e, factor, dMax As Double
        Dim v1, v2, v3 As Vector
        dMax = 0
        For Each f As Face In sb.Faces
            d = f.GetClosestPointTo(wpt.Point).DistanceTo(wpt.Point)
            For Each ed As Edge In f.Edges
                e = ed.GetClosestPointTo(wpt.Point).DistanceTo(wpt.Point)
                v1 = ed.GetClosestPointTo(wpt.Point).VectorTo(wpt.Point)
                v2 = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point)
                v3 = v1.CrossProduct(v2)
                factor = v3.Length * v2.Length / (v1.Length + 0.0000001)
                If factor > dMax Then
                    dMax = factor
                    stampEdge = ed
                    stampFace = f
                End If
            Next

        Next


        Return stampFace
    End Function
    Function GetStampLine(wpt As WorkPoint) As SketchLine3D


        Dim d, e, factor, dMax, ave As Double
        ave = 0
        dMax = 0
        For Each skl As SketchLine3D In palitos.rodLines
            ave = skl.Length + ave
        Next
        ave = ave / palitos.rodLines.Count
        For Each skl As SketchLine3D In palitos.rodLines
            factor = skl.Length / (skl.Geometry.DistanceTo(wpt.Point) + Math.Exp(Math.Abs(skl.Geometry.MidPoint.Z) - 35 / 10))
            If skl.Length > ave Then
                If factor > dMax Then
                    dMax = factor
                    stampLine = skl
                End If
            End If
        Next

        ' lamp.HighLighObject(stampEdge)

        Return stampLine
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
    Function IsSkeletonDone(s As String) As Boolean
        For Each ffn As String In fullFileNames
            If ffn.Equals(s) Then
                Return True
            End If
        Next

        Return False
    End Function
    Function MakeSingleSkeleton(s As String) As PartDocument
        Dim p As PartDocument
        Dim q As Integer
        Dim derivedDefinition As DerivedPartDefinition
        Dim newComponent As DerivedPartComponent
        Dim rf As RevolveFeature
        Dim lf As LoftFeature
        Dim pt As Point
        Dim d, e As Double
        Dim skt As Sketch3D
        Dim cc As Integer = 1


        Try
            p = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
            Conversions.SetUnitsToMetric(p)
            derivedDefinition = p.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsWorkSurface
            newComponent = p.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)
            p.Update2(True)
            doku = DocUpdate(p)
            qvalue = nombrador.GetQNumberString(s)
            convergePoint = GetConvergePoints(qvalue + 1)
            caraTrabajo = New Surfacer(doku)
            compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count).Visible = False
            tangentSurfaces = caraTrabajo.GetTangentials(GetInitialFace(compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)))
            palitos = New RodMaker(doku)
            conos = New Wedges(doku)
            conos.coneAxes.Clear()
            palitos.rodLines.Clear()
            conos.highPoints.Clear()
            conos.lowPoints.Clear()
            conos.highPoints.Add(wptHigh.Point)
            conos.lowPoints.Add(wptLow.Point)
            For Each sb As SurfaceBody In tangentSurfaces.SurfaceBodies
                For Each f As Face In sb.Faces
                    If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        ' lamp.HighLighFace(fc.Item(i))
                        If monitor.IsFeatureHealthy(palitos.ExtrudeEgg(f)) Then
                            If palitos.smallOval Then
                                skt = compDef.Sketches3D.Item("curvas")
                                lf = conos.MakeSingleSupport(palitos, skt)
                                If monitor.IsFeatureHealthy(lf) Then
                                    lf.Name = String.Concat("cone", cc.ToString)
                                    cc = cc + 1
                                    palitos.smallOval = False
                                Else
                                    Exit For
                                End If
                            End If

                        End If



                    End If
                Next
            Next

            Return doku
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return p
    End Function
    Function MakeTwoHoles() As RevolveFeature
        Dim rf As RevolveFeature
        rf = MakeRingEntrance(wptHigh)
        If monitor.IsFeatureHealthy(rf) Then
            rf.Name = "holeHigh"
            rf = MakeRingEntrance(wptLow)
            If monitor.IsFeatureHealthy(rf) Then
                rf.Name = "holeLow"
                Return rf
            End If
        End If


        Return Nothing
    End Function
    Function DocUpdate(docu As PartDocument) As PartDocument
        doku = docu

        compDef = doku.ComponentDefinition
        lamp = New Highlithing(doku)

        Return doku
    End Function
    Function GetConvergePoints(q As Integer) As Point
        Dim wpt As WorkPoint
        Dim pt As Point = tg.CreatePoint(Math.Cos(2 * Math.PI * DP.p * q / DP.q) * DP.Dmax / 2, Math.Sin(2 * Math.PI * DP.p * q / DP.q) * DP.Dmax / 2, 0)
        wpt = doku.ComponentDefinition.WorkPoints.AddFixed(pt)
        wpt.Visible = False
        wpConverge = wpt
        GetConvergePointHiLo(q)
        Dim wpl As WorkPlane = compDef.WorkPlanes.AddByThreePoints(wpt, skptHigh, skptLow)
        wpl.Visible = False
        Dim seq As SketchEquationCurve3D = trobinaCurve.DrawXYPlaneRing(sk3D, 55 / 1, 20 / 1, q)
        sk3D.GeometricConstraints3D.AddGround(seq)
        wptHigh = compDef.WorkPoints.AddByCurveAndEntity(seq, wpl)
        wptHigh.Name = String.Concat("wptHigh")
        wptHigh.Visible = False
        seq = trobinaCurve.DrawXYPlaneRing(sk3D, 55 / 1, -20 / 1, q)
        sk3D.GeometricConstraints3D.AddGround(seq)
        wptLow = compDef.WorkPoints.AddByCurveAndEntity(seq, wpl)
        wptLow.Name = "wptLow"
        wptLow.Visible = False


        Return pt
    End Function
    Function GetConvergePointHiLo(q As Integer) As Point
        sk3D = compDef.Sketches3D.Add
        sk3D.Name = "curvas"
        Dim sc As SketchEquationCurve3D = trobinaCurve.DrawHalfTrobinaCurve(sk3D, q)
        sk3D.GeometricConstraints3D.AddGround(sc)
        skptHigh = sc.StartSketchPoint
        skptLow = sc.EndSketchPoint
        sk3D.Visible = False
        Return sc.StartSketchPoint.Geometry
    End Function
    Function GetInitialFace(ws As WorkSurface) As Face
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
            lamp.HighLighFace(fmin1)
            cylinderFace = fmin1


            Return cylinderFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function MakeRingEntrance(wpt As WorkPoint) As RevolveFeature
        Dim rf As RevolveFeature

        Try
            rf = RevolveHole(DrawSingleEntrace(wpt))
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return rf
    End Function
    Function DrawSingleEntrace(wpt As WorkPoint) As Profile
        Try
            Dim pr As Profile
            Dim ps As PlanarSketch
            Dim sl As SketchLine
            sk3D = compDef.Sketches3D.Add
            ringLine = sk3D.SketchLines3D.AddByTwoPoints(compDef.WorkPoints.Item(1), wpt, False)
            Dim skpt As SketchPoint3D = sk3D.SketchPoints3D.Add(tg.CreatePoint(0, 0, 1))
            Dim wpl As WorkPlane = compDef.WorkPlanes.AddByThreePoints(compDef.WorkPoints.Item(1), skpt, wpt)
            lamp.LookAtPlane(wpl)
            ps = doku.ComponentDefinition.Sketches.Add(wpl)
            wpl.Visible = False
            pr = DrawEntranceCircle(ps, 5 / 10)

            ps.Visible = False
            sk3D.Visible = False
            Return pr
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawSingleCircle(ps As Sketch, r As Double) As Profile
        Dim pro As Profile

        Dim sl As SketchLine = ps.AddByProjectingEntity(ringLine)
        Try
            sl.Construction = True
            ps.SketchCircles.AddByCenterRadius(sl.EndSketchPoint, r)
            pro = ps.Profiles.AddForSolid

            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawEntranceCircle(ps As Sketch, r As Double) As Profile
        Dim pro As Profile
        Dim slup, sldown As SketchLine


        Try
            adjuster = New SketchAdjust(doku)
            comando.WireFrameView(doku)
            Dim sl As SketchLine = ps.AddByProjectingEntity(ringLine)
            sl.Construction = True
            Dim c As SketchCircle = ps.SketchCircles.AddByCenterRadius(sl.EndSketchPoint, r)
            c.Construction = True
            Dim slxy As SketchLine = ps.AddByProjectingEntity(compDef.WorkPlanes.Item(3))
            slxy.Construction = True
            Dim slref As SketchLine = ps.SketchLines.AddByTwoPoints(sl.EndSketchPoint, sl.StartSketchPoint.Geometry)
            slref.Construction = True
            Dim gc As GeometricConstraint = ps.GeometricConstraints.AddParallel(slref, slxy)
            Dim pt As Point2d = slref.EndSketchPoint.Geometry
            Dim dc As DimensionConstraint = ps.DimensionConstraints.AddTwoPointDistance(slref.StartSketchPoint, slref.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, slref.EndSketchPoint.Geometry)
            Try
                dc.Parameter._Value = r
            Catch ex As Exception
                adjuster.AdjustDimension2DSmothly(dc, r)
            End Try


            Dim sl2 As SketchLine = ps.SketchLines.AddByTwoPoints(sl.EndSketchPoint, slref.EndSketchPoint.Geometry)

            Dim ac As DimensionConstraint = ps.DimensionConstraints.AddTwoLineAngle(slref, sl2, sl2.EndSketchPoint.Geometry)
            ac.Parameter._Value = Math.PI / 36
            adjuster.AdjustDimension2DSmothly(ac, Math.PI / 3)
            gc = ps.GeometricConstraints.AddCoincident(sl2.EndSketchPoint, c)
            Dim sl3 As SketchLine = ps.SketchLines.AddByTwoPoints(sl.EndSketchPoint, slref.EndSketchPoint.Geometry)
            Dim ls As LineSegment2d = tg.CreateLineSegment2d(sl3.EndSketchPoint.Geometry, sl2.EndSketchPoint.Geometry)
            ac = ps.DimensionConstraints.AddTwoLineAngle(sl2, sl3, ls.MidPoint)
            adjuster.AdjustDimension2DSmothly(ac, 2 * Math.PI / 3)
            gc = ps.GeometricConstraints.AddCoincident(sl3.EndSketchPoint, c)
            Dim sai As SketchArc = ps.SketchArcs.AddByCenterStartEndPoint(c.CenterSketchPoint, sl2.EndSketchPoint, sl3.EndSketchPoint, True)
            sl2.Construction = True
            sl3.Construction = True
            gc = ps.GeometricConstraints.AddGround(sai)

            slup = DrawEntraceRingEdge(ps, sl2, slref, pt, 5 * r)
            sldown = DrawEntraceRingEdge(ps, sl3, slref, pt, 5 * r)
            Dim sltop As SketchLine = ps.SketchLines.AddByTwoPoints(slup.EndSketchPoint, sldown.EndSketchPoint)

            pro = ps.Profiles.AddForSolid

            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawEntraceRingEdge(ps As PlanarSketch, sli As SketchLine, slref As SketchLine, pt As Point2d, r As Double) As SketchLine
        Dim sl As SketchLine = ps.SketchLines.AddByTwoPoints(sli.EndSketchPoint, pt)
        Dim gc As GeometricConstraint = ps.GeometricConstraints.AddParallel(sl, slref)
        Dim dc As DimensionConstraint = ps.DimensionConstraints.AddTwoPointDistance(sl.StartSketchPoint, sl.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, sl.EndSketchPoint.Geometry)
        Try
            dc.Parameter._Value = r
        Catch ex As Exception
            adjuster.AdjustDimension2DSmothly(dc, r)
        End Try

        'dc.Driven = True
        Return sl
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
    Function DrawSingleCircle(ps As Sketch, r As Double, c As WorkPoint) As Profile
        Dim pro As Profile

        Dim spt As SketchPoint = ps.AddByProjectingEntity(c)
        Try
            ps.SketchCircles.AddByCenterRadius(spt, r)
            pro = ps.Profiles.AddForSolid

            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

End Class
