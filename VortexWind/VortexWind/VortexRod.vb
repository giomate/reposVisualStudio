Imports System.Text.RegularExpressions
Imports Inventor

Public Class VortexRod
    Public doku, reference, banda As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk As Sketch3D
    Dim sabana As Surfacer
    Dim adjuster As SketchAdjust


    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy, outlet, inlet As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile
    Dim windings, driftAngle, passes, startAngle, tangle, cutAngle As Double
    Dim estampa As Stanzer
    Dim puente As RodMaker
    Dim lista As ExcelInterface

    Public wp1, wp2, wp3, wpConverge, startWorkPoint, currentWorkPoint As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, ptRReference, ptzBack, ptZMax, pointCutRadiusMax, ptRMin As Point
    Dim skpt1, skpt2, skpt3 As SketchPoint3D
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Public partNumber, qNext, qLastTie As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim tangents, bands, rods, currentWorkSurface As WorkSurface
    Dim comando As Commands
    Public nombrador As Nombres

    Dim cutProfile, faceProfile, rodProfile As Profile
    Dim circle As SketchCircle3D

    Dim pro As Profile
    Dim direction As Vector
    Dim lastWorkPlane, nextWorkPlane, currentWorkPlane, startWorkPlane As WorkPlane
    Dim lastWorkAxis, nextWorkAxis, startWorkAxis, currentWorkAxis As WorkAxis

    Public compDef As PartComponentDefinition

    Dim mainWorkPlane As WorkPlane
    Dim workAxis As WorkAxis
    Dim faceRod As Face
    Dim workFace, adjacentFace, bendFace, frontBendFace, stampPlanarFace, stampCurveFace As Face
    Dim tangentFaces As FaceCollection
    Dim minorEdge, majorEdge, inputEdge, secondEdge As Edge
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
    Dim pValue, qValue, sbValue, bandNumbers(), nTanFaces As Integer
    Public iteration As Integer
    Dim foldFeature As FoldFeature
    Dim sections, stamPoints, surfaceBodies, bandas, guidePoints, cylinders, planarFaces, wedges, tangentials, arcPoints, surfacesSculpt As ObjectCollection
    Dim affectedBodies As ObjectCollection
    Dim tanKeys(), bandKeys(), rodKeys(), keyPlanarFace As Long
    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim spt2dHigh, spt2dLow As SketchPoint
    Dim sptHigh, sptLow, sptHigh2, sptLow2 As SketchPoint3D
    Dim freeRadius As Double
    Public CW As Boolean
    Dim rotationDirection As Double

    Dim arrayFunctions As Collection
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
        sabana = New Surfacer(doku)
        estampa = New Stanzer(doku)
        adjuster = New SketchAdjust(doku)

        projectManager = app.DesignProjectManager

        compDef = doku.ComponentDefinition
        features = compDef.Features
        tg = app.TransientGeometry
        sections = app.TransientObjects.CreateObjectCollection
        guidePoints = app.TransientObjects.CreateObjectCollection
        cylinders = app.TransientObjects.CreateObjectCollection
        planarFaces = app.TransientObjects.CreateObjectCollection
        wedges = app.TransientObjects.CreateObjectCollection
        tangentials = app.TransientObjects.CreateObjectCollection
        arcPoints = app.TransientObjects.CreateObjectCollection
        bandas = app.TransientObjects.CreateObjectCollection
        surfacesSculpt = app.TransientObjects.CreateObjectCollection
        surfaceBodies = app.TransientObjects.CreateObjectCollection
        affectedBodies = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        windings = 211
        passes = 89
        gap1CM = 3 / 10
        pValue = 0
        qValue = 0
        done = False
        guidePoints.Clear()
        DP.Dmax = 171 * 20 / 194
        ' DP.Dmax = 200 / 10
        DP.Dmin = 1 / 10
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 11
        DP.q = 23
        DP.b = 25
        rotationDirection = Math.Pow(-1, (1 + CInt(CW)))
        freeRadius = 20 / 10
        startAngle = rotationDirection * 4 * Math.PI * DP.p / (DP.q) + 0 * Math.PI / 2
        tangle = Math.PI - 2 * Math.Asin(freeRadius * 2 / DP.Dmax)
        cutAngle = Math.IEEERemainder(startAngle + 1 * rotationDirection * tangle / 2, Math.PI * 2)

        driftAngle = Math.PI * 2 * passes / windings

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


        nombrador = New Nombres(doku)
        puente = New RodMaker(doku)
        'adjuster = New SketchAdjust(doku)

        Return doku
    End Function
    Public Function MakeAllWiresGuides(docu As PartDocument) As ExtrudeFeature
        Dim ef As ExtrudeFeature = Nothing
        Dim w As Parameter
        Dim wpt1 As WorkPoint



        doku = DocUpdate(docu)
        comando.WireFrameView(doku)


        Try
            Try
                wpt1 = compDef.WorkPoints.Item("StartPoint")
                Try
                    tangents = compDef.WorkSurfaces.Item("tangents")
                    ' CreateTangenFaceCollection()
                    Try
                        rods = compDef.WorkSurfaces.Item("cylinders")
                        Try
                            'bands = compDef.WorkSurfaces.Item("bands")
                            Try
                                ' ReadKeys()
                                CreateKeys(rods, rodKeys)
                                'CreateKeys(bands, bandKeys)
                            Catch ex As Exception

                            End Try
                        Catch ex As Exception
                            'bands = MakePlanarFaces()
                            CreateKeys(rods, rodKeys)
                            ' CreateKeys(bands, bandKeys)
                        End Try
                    Catch ex As Exception
                        rods = MakeCylinders()

                    End Try

                Catch ex As Exception
                    tangents = MakeTangentials()
                    rods = MakeCylinders()
                    'bands = MakePlanarFaces()
                    CreateKeys(rods, rodKeys)
                    'CreateKeys(bands, bandKeys)

                End Try
            Catch ex2 As Exception

                If GetRadiusPoint(GetStartWorkPoint.Point) > 20 / 10 Then

                End If
            End Try

            Try
                If monitor.IsFeatureHealthy(compDef.Features.ExtrudeFeatures.Item("rw1")) Then
                    qValue = FindLastSW()
                    ef = MakeNextWireHole(qValue)
                End If
            Catch ex As Exception
                sk3D = compDef.Sketches3D.Add()
                sk3D.Name = "FirstWire"

                ef = RemoveFirstWire(skpt1)
                If monitor.IsFeatureHealthy(ef) Then
                    qValue = 1
                    w = GetParameter("wq")
                    w._Value = qValue
                    ef.Name = "rw1"
                    doku.Update2(True)
                    doku.Save2(True)
                    ef = MakeNextWireHole(qValue)
                End If
            End Try






            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function StampAllWireGuides(docu As PartDocument) As ExtrudeFeature
        Dim ef As ExtrudeFeature = Nothing

        Dim wpt1 As WorkPoint
        doku = DocUpdate(docu)
        comando.WireFrameView(doku)
        comando.HideSketches(doku)
        Try
            Try
                wpt1 = compDef.WorkPoints.Item("wpt1")

            Catch ex2 As Exception

                If GetRadiusPoint(GetStartWorkPoint().Point) >= freeRadius - 1 / 1024 Then
                    wpt1 = startWorkPoint
                End If
            End Try

            Try

                qValue = FindLastSW()
                If qValue > 0 Then
                    ef = StampNextWire(CInt(qValue / 2) + 1)
                Else
                    ef = StampNextWire(1)
                End If

            Catch ex As Exception

            End Try






            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function StampAllWireGuides(docu As PartDocument, cwi As Boolean) As CutFeature
        Dim cf As CutFeature = Nothing
        CW = cwi
        rotationDirection = Math.Pow(-1, (2 + CInt(CW)))
        startAngle = Math.IEEERemainder(4 * Math.PI * DP.p / (DP.q), Math.PI * 2)
        cutAngle = Math.IEEERemainder(startAngle + 1 * rotationDirection * tangle / 2, Math.PI * 2)

        Dim wpt1 As WorkPoint
        doku = DocUpdate(docu)
        comando.WireFrameView(doku)
        ' comando.HideSketches(doku)
        Try
            Try
                wpt1 = compDef.WorkPoints.Item("wpt1")
                startWorkPoint = wpt1
            Catch ex2 As Exception

                If GetRadiusPoint(GetStartWorkPoint().Point) >= freeRadius - 1 / 1024 Then
                    wpt1 = startWorkPoint
                End If
            End Try

            Try

                qValue = FindLastSW()
                If qValue > 0 Then
                    cf = StampNextWire(Math.Floor(qValue / 2) + 1)
                Else
                    cf = StampNextWire(1)
                End If

            Catch ex As Exception

            End Try






            Return cf

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function ResumeStampLetters(docu As PartDocument) As ExtrudeFeature
        Dim ef As ExtrudeFeature = Nothing

        Dim wpt1 As WorkPoint
        doku = DocUpdate(docu)
        comando.WireFrameView(doku)
        comando.HideSketches(doku)
        Try
            Try
                wpt1 = compDef.WorkPoints.Item("wpt1")

            Catch ex2 As Exception

                If GetRadiusPoint(GetStartWorkPoint.Point) >= freeRadius - 1 / 1024 Then
                    wpt1 = startWorkPoint
                End If
            End Try

            Try

                qValue = FindLastSW()
                If qValue > 0 Then
                    ef = StampNextWire(qValue + 1)
                Else
                    ef = StampNextWire(1)
                End If

            Catch ex As Exception

            End Try






            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetStartWorkPoint() As WorkPoint
        Dim wpt As WorkPoint
        Dim pt As Point = tg.CreatePoint(Math.Cos(startAngle) * freeRadius, Math.Sin(startAngle) * freeRadius, 0)
        wpt = doku.ComponentDefinition.WorkPoints.AddFixed(pt)
        wpt.Visible = False
        wpt.Name = "wpt1"
        startWorkPoint = wpt
        Return wpt
    End Function
    Public Function ResumeWiresGuidesReference(docu As PartDocument, ref As PartDocument) As ExtrudeFeature
        reference = ref
        tangents = sabana.ReferencedTangentials(reference)
        rods = sabana.ReferencedCylinders(reference)
        Return ResumeWiresGuides(docu)
    End Function
    Public Function ResumeWiresGuides(docu As PartDocument) As ExtrudeFeature
        Dim ef As ExtrudeFeature = Nothing




        doku = DocUpdate(docu)
        doku.Activate()
        comando.WireFrameView(doku)


        Try
            Try
                'skpt1 = compDef.Sketches3D.Item("HighestPoint").SketchPoints3D.Item(1)
                Try
                    tangents = compDef.WorkSurfaces.Item("tangents")
                    ' CreateTangenFaceCollection()
                    Try
                        rods = compDef.WorkSurfaces.Item("cylinders")
                        Try
                            'bands = compDef.WorkSurfaces.Item("bands")
                            Try
                                'ReadKeys()
                                CreateKeys(rods, rodKeys)
                                'CreateKeys(bands, bandKeys)
                            Catch ex As Exception

                            End Try
                        Catch ex As Exception
                            'bands = MakePlanarFaces()
                            ' CreateKeys(rods, rodKeys)
                            ' CreateKeys(bands, bandKeys)
                        End Try
                    Catch ex As Exception
                        'rods = sabana.ReferencedCylinders(reference)
                        ' CreateKeys(rods, rodKeys)
                    End Try

                Catch ex As Exception
                    ' tangents = sabana.ReferencedTangentials(reference)
                    'rods = sabana.ReferencedCylinders(reference)
                    'bands = MakePlanarFaces()
                    CreateKeys(rods, rodKeys)
                    'CreateKeys(bands, bandKeys)

                End Try
            Catch ex2 As Exception

            End Try

            Try
                Try
                    qValue = FindLastWpl()
                    ef = ResumeWireHole(qValue)

                Catch ex As Exception

                End Try

            Catch ex As Exception

            End Try






            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CreateTangenFaceCollection() As Integer
        Dim n As Integer = compDef.Features.NonParametricBaseFeatures.Item("tangentials").Faces.Count
        Dim i As Integer = 0
        ReDim tanKeys(n - 1)
        tangentFaces = app.TransientObjects.CreateFaceCollection
        tangentFaces.Clear()


        For Each sb As SurfaceBody In tangents.SurfaceBodies
            For Each f As Face In sb.Faces
                i = i + 1
                tangentFaces.Add(f)

                tanKeys(i - 1) = f.TransientKey
            Next
        Next
        nTanFaces = n
        Return n
    End Function
    Function CreateKeys(ws As WorkSurface, ByRef keys() As Long) As Integer

        Dim n As Integer = compDef.Features.NonParametricBaseFeatures.Item("tangentials").Faces.Count
        Dim i As Integer = 0

        Dim j As Integer = 0
        Dim k As Integer

        If True Then

        End If
        nTanFaces = n

        Dim fc As FaceCollection = app.TransientObjects.CreateFaceCollection
        Try
            For Each sb As SurfaceBody In ws.SurfaceBodies
                For Each fr As Face In sb.Faces
                    fc.Add(fr)
                Next
            Next
            k = fc.Count
            ReDim tanKeys(n - 1)
            ReDim keys(n - 1)
            For Each sb As SurfaceBody In tangents.SurfaceBodies
                For Each f As Face In sb.Faces
                    i = i + 1
                    tanKeys(i - 1) = f.TransientKey

                    keys(i - 1) = 0
                    If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        If fc.Count > 0 Then
                            For Each f2 As Face In fc
                                ' lamp.HighLighFace(f)
                                If IsSameFace(f, f2) Then
                                    '  lamp.HighLighFace(f)
                                    '  lamp.HighLighFace(f2)
                                    keys(i - 1) = f2.TransientKey
                                    fc.RemoveByObject(f2)
                                    j = j + 1
                                    Exit For
                                End If

                            Next
                        End If
                    End If

                Next

            Next



            If j = k Then
                SaveKeys(tanKeys, rodKeys)
                Return j

            Else
                MsgBox(j.ToString())
                Return Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return 0
    End Function
    Function SaveKeys(tans() As Long, rods() As Long) As Integer
        Dim path As String = projectManager.ActiveDesignProject.WorkspacePath
        Try
            lista = New ExcelInterface(path)
            lista.SaveArray(tans, rods)
            Return 0
        Catch ex As Exception

        End Try

    End Function

    Function IsSameFace(f1 As Face, f2 As Face) As Boolean
        Dim b As Boolean = False
        Dim d As Double
        Dim c1, c2 As Cylinder
        Dim pl1, pl2 As Plane
        Try
            If f1.Vertices.Count = f2.Vertices.Count Then
                d = Math.Abs(f1.Evaluator.Area - f2.Evaluator.Area)
                If d < 0.001 Then
                    If f1.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        If f2.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                            c1 = f1.Geometry
                            c2 = f2.Geometry
                            d = Math.Abs(c1.AxisVector.AsVector.DotProduct(c2.AxisVector.AsVector))
                            If d > 0.99 Then
                                d = c1.BasePoint.DistanceTo(c2.BasePoint)
                                If d < 1 / 1000 Then
                                    b = True
                                End If

                            End If
                        End If
                    Else
                        If f1.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                            If f2.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                                pl1 = f1.Geometry
                                pl2 = f2.Geometry
                                d = Math.Abs(pl1.Normal.AsVector.DotProduct(pl2.Normal.AsVector))
                                If d > 0.99 Then
                                    d = pl1.RootPoint.DistanceTo(pl2.RootPoint)
                                    If d < 1 / 1000 Then
                                        b = True
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
        Return b
    End Function
    Function AreDifferentFaces(f1 As Face, f2 As Face) As Double
        Dim b As Boolean = False
        Dim d, e, f As Double
        Dim c1, c2 As Cylinder
        Dim pl1, pl2 As Plane
        d = 9999999
        e = d
        f = e
        Try
            If f1.Vertices.Count = f2.Vertices.Count Then
                d = Math.Abs(f1.Evaluator.Area - f2.Evaluator.Area)
                If d < 0.1 Then
                    If f1.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        If f2.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                            c1 = f1.Geometry
                            c2 = f2.Geometry
                            e = (c1.AxisVector.AsVector.CrossProduct(c2.AxisVector.AsVector)).Length
                            f = c1.BasePoint.DistanceTo(c2.BasePoint)
                        End If
                    Else
                        If f1.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                            If f2.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                                pl1 = f1.Geometry
                                pl2 = f2.Geometry
                                e = pl1.Normal.AsVector.CrossProduct(pl2.Normal.AsVector).Length
                                f = pl1.RootPoint.DistanceTo(pl2.RootPoint)
                            End If
                        End If
                    End If
                End If
            End If
            Return e * d * f
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return b
    End Function
    Function MakeCylinders() As WorkSurface
        Dim ws As WorkSurface
        Dim sb1 As SurfaceBody
        Dim sb As SurfaceBody = compDef.SurfaceBodies.Item(1)
        Dim np As NonParametricBaseFeature
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition
        Try
            cylinders.Clear()
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    If f.Evaluator.Area > 0.5 And f.Evaluator.Area < 1 Then
                        cylinders.Add(f)
                    End If
                End If
            Next

            npDef.BRepEntities = cylinders
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
            np.Name = "rodillos"
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sb1 = ws.SurfaceBodies.Item(1)
            sb1.Name = "cylinders"
            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function MakePlanarFaces() As WorkSurface
        Dim ws As WorkSurface
        Dim sb1 As SurfaceBody
        Dim sb As SurfaceBody = compDef.SurfaceBodies.Item(1)
        Dim np As NonParametricBaseFeature
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition
        Try
            planarFaces.Clear()
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    If f.Evaluator.Area > 2 And f.Evaluator.Area < 9 Then
                        If f.TangentiallyConnectedFaces.Count > 7 Then
                            planarFaces.Add(f)
                        End If

                    End If
                End If
            Next

            npDef.BRepEntities = planarFaces
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
            np.Name = "bandas"
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sb1 = ws.SurfaceBodies.Item(1)
            sb1.Name = "bands"
            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function MakeTangentials() As WorkSurface
        Dim ws As WorkSurface

        Dim sb As SurfaceBody = compDef.SurfaceBodies.Item(1)
        Dim np As NonParametricBaseFeature
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition


        Try
            tangentials.Clear()
            For Each f As Face In sb.Faces
                If f.Evaluator.Area > 0.09 And f.Evaluator.Area < 9 Then
                    If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Or f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        Try
                            If f.TangentiallyConnectedFaces.Count > 13 Then
                                tangentials.Add(f)
                            End If
                        Catch ex As Exception

                        End Try
                    End If



                End If

            Next

            npDef.BRepEntities = tangentials
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)

            ' npAUx = CheckSurfaceValidity(np)
            '  If monitor.IsFeatureHealthy(npAUx) Then
            '   np = npAUx
            '   End If

            np.Name = "tangentials"
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sb = ws.SurfaceBodies.Item(1)
            sb.Name = "tangents"

            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function CheckSurfaceValidity(npi As NonParametricBaseFeature) As NonParametricBaseFeature
        Dim fc As FaceCollection = app.TransientObjects.CreateFaceCollection
        Dim badFaces As ObjectCollection = app.TransientObjects.CreateObjectCollection
        Dim goodFaces As ObjectCollection = app.TransientObjects.CreateObjectCollection
        Dim oe As ObjectsEnumerator = Nothing
        Dim ef As ExtrudeFeature
        Dim sb As SurfaceBody

        Dim bf As Face
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition
        Dim wd As Boolean
        Dim d As Double
        Dim c1, c2 As Cylinder


        Dim np As NonParametricBaseFeature
        Dim count As Integer = npi.Faces.Count

        Try
            badFaces.Clear()
            goodFaces.Clear()
            fc.Clear()

            For Each f As Face In npi.Faces

                sb = f.SurfaceBody
                If Not sb.IsEntityValid(f, , oe) Then
                    badFaces.Add(f)
                    lamp.HighLighFace(f)
                    ef = puente.MakeBridge(f)
                    If monitor.IsFeatureHealthy(ef) Then
                        lamp.HighLighFace(ef.Faces.Item(1))
                        goodFaces.Add(ef.Faces.Item(1))
                        lamp.HighLighObject(goodFaces.Item(goodFaces.Count))
                        'compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count).SurfaceBodies.Item(1).Visible = False
                    End If

                End If

            Next

            'npi.Delete()
            tangentials.Clear()
            For Each f As Face In compDef.SurfaceBodies.Item(1).Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Or f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    If f.Evaluator.Area > 0.09 And f.Evaluator.Area < 9 Then
                        Try
                            If f.TangentiallyConnectedFaces.Count > 13 Then
                                wd = False
                                If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                                    For i = 1 To badFaces.Count

                                        bf = badFaces.Item(i)
                                        If f.Vertices.Count = bf.Vertices.Count Then
                                            d = Math.Abs(f.Evaluator.Area - bf.Evaluator.Area)

                                            If d < 0.001 Then
                                                c1 = f.Geometry
                                                c2 = bf.Geometry
                                                d = Math.Abs(c1.AxisVector.AsVector.DotProduct(c2.AxisVector.AsVector))
                                                If d > 0.99 Then
                                                    lamp.HighLighFace(f)
                                                    tangentials.Add(goodFaces.Item(i))
                                                    wd = True
                                                    Exit For
                                                End If

                                            End If
                                        End If


                                    Next
                                End If

                                If Not wd Then
                                    tangentials.Add(f)
                                End If

                            End If
                        Catch ex As Exception

                        End Try


                    End If
                End If


            Next



            npDef.BRepEntities = tangentials
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            Try
                np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
            Catch ex As Exception
                npDef.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType
                Try
                    np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
                Catch ex2 As Exception
                    npDef.DeleteOriginal = True
                    npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
                    Try
                        np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
                    Catch ex3 As Exception
                        npDef.BRepEntities = goodFaces
                        np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
                    End Try
                End Try

            End Try
            If Not monitor.IsFeatureHealthy(np) Then
                CheckSurfaceValidity(np)
            Else
                npi.Delete()
            End If

            Return np
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetArcPoints(fi As Face) As Point
        Dim min1, min2, min3, min4, d As Double
        Dim pt1, pt2, pt3, pt4 As Point
        min1 = 99999999
        min2 = min1
        min3 = min2
        min4 = min3
        pt1 = fi.Vertices.Item(1).Point
        pt2 = pt1
        pt3 = pt2
        pt4 = pt3
        For Each f As Face In fi.TangentiallyConnectedFaces
            For Each ver As Vertex In fi.Vertices
                For Each ver2 As Vertex In f.Vertices
                    d = ver.Point.DistanceTo(ver2.Point)
                    If d < min4 Then
                        If d < min3 Then
                            If d < min2 Then
                                If d < min1 Then
                                    min1 = d
                                    pt4 = pt3
                                    pt2 = pt1
                                    pt2 = pt1
                                    pt1 = ver
                                End If
                            End If

                        End If

                    End If
                Next

            Next
        Next
        Return pt1
    End Function
    Function FindLastSW() As Integer
        Dim sws As Integer
        Dim pattern As String = "sw"
        Dim s As String
        Dim skl As Sketch3D

        Try
            sws = 0

            For Each sk As Sketch3D In compDef.Sketches3D
                s = String.Concat("sw", CInt(sws + 1).ToString)
                Try
                    skl = compDef.Sketches3D.Item(s)

                    sws += 1



                Catch ex As Exception
                    Return (sws)
                End Try




            Next


            Return sws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return 1
    End Function
    Function FindLastWpl() As Integer
        Dim wpls As Integer
        Dim pattern As String = "wpl"
        Dim s As String

        Try
            wpls = 0

            For Each wpl As WorkPlane In compDef.WorkPlanes
                If Regex.IsMatch(wpl.Name, pattern) Then
                    s = String.Concat(pattern, CInt(wpls + 1).ToString)
                    Try
                        If Not compDef.WorkPlanes.Item(s).IsOwnedByFeature Then
                            wpls = wpls + 1
                        Else
                            Return (windings + 1)
                        End If
                    Catch ex As Exception
                        Return (wpls)
                    End Try



                End If
            Next


            Return wpls
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return 1
    End Function
    Function MakeNextWireHole(q As Integer) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim w As Parameter
        Try
            ef = compDef.Features.ExtrudeFeatures.Item(compDef.Features.ExtrudeFeatures.Count)
            startWorkAxis = compDef.WorkAxes.Item(4)
            If monitor.IsFeatureHealthy(ef) Then

                While (q < windings + 1 And Not done)
                    ef = RemoveNextWire()
                    If monitor.IsFeatureHealthy(ef) Then
                        If ef.SideFaces.Count > 1 Then
                            Me.qValue = Me.qValue + 1
                            w = GetParameter("wq")
                            w._Value = Me.qValue
                            ef.Name = String.Concat("rw", CInt(w._Value).ToString)
                            doku.Update2(True)
                            doku.Save2(True)
                            Me.qValue = FindLastSW()
                        Else
                            done = True
                        End If




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
    Function StampNextWire(q As Integer) As CutFeature
        Dim cf As CutFeature = Nothing

        Try

            Try
                startWorkAxis = compDef.WorkAxes.Item("wa1")
                startWorkPlane = compDef.WorkPlanes.Item("wpl1")
            Catch ex As Exception
                startWorkAxis = DrawFirstAxis()
            End Try

            While (q < windings + 1 And Not done)
                comando.WireFrameView(doku)
                currentWorkPlane = DrawReferences(q)
                cf = StampLetters(q)
                If done Then
                    done = False

                    doku.Update2(True)
                    doku.Save2(True)
                    q += 1
                    ' q = FindLastSW()
                Else
                    Exit While

                End If




            End While


            Return cf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function ResumeWireHole(q As Integer) As ExtrudeFeature
        Dim ef As ExtrudeFeature

        Dim i As Integer
        Try
            i = compDef.Features.ExtrudeFeatures.Count
            If i > 0 Then
                ef = compDef.Features.ExtrudeFeatures.Item(compDef.Features.ExtrudeFeatures.Count)
            Else
                i = reference.ComponentDefinition.Features.ExtrudeFeatures.Count
                ef = reference.ComponentDefinition.Features.ExtrudeFeatures.Item(i)

            End If


            If monitor.IsFeatureHealthy(ef) Then

                While (q < windings + 0 And Not done)
                    ef = RemoveNextWire()
                    If monitor.IsFeatureHealthy(ef) Then
                        If ef.SideFaces.Count > 1 Then
                            Me.qValue = Me.qValue + 1

                            ef.Name = String.Concat("rw", Me.qValue.ToString)
                            doku.Update2(True)
                            doku.Save2(True)
                            Me.qValue = FindLastWpl()
                            If q > windings Then
                                done = True
                            End If
                        Else
                            done = True
                        End If




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
        Dim skl, sklw As SketchLine3D
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
            wa.Name = "wa1"
            wpl = compDef.WorkPlanes.AddByThreePoints(spt3, skpt, spt2)
            wpl.Visible = False
            wpl.Name = "wpl1"
            currentWorkPlane = wpl
            sklw = DrawWireAxis(spt3, wpl)
            ef = RemoveSingleWire(sklw)

            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawFirstAxis() As WorkAxis

        Dim spt1, spt2, spt3 As SketchPoint3D
        Dim pt1, pt2, pt3 As Point
        Dim skl As SketchLine3D
        Dim wpl As WorkPlane
        Dim wa As WorkAxis

        Dim v, vz, vpp As Vector

        Try
            sk3D = compDef.Sketches3D.Add
            sk3D.Name = "sk3D1"
            pt1 = compDef.WorkPoints.Item(1).Point
            spt1 = sk3D.SketchPoints3D.Add(pt1)
            pt2 = startWorkPoint.Point
            spt2 = sk3D.SketchPoints3D.Add(pt2)
            v = pt1.VectorTo(pt2)
            vz = tg.CreateVector(0, 0, 1)
            vpp = v.CrossProduct(vz)
            vpp.Normalize()
            vpp.ScaleBy(freeRadius)
            pt3 = pt1
            pt3.TranslateBy(vpp)
            spt3 = sk3D.SketchPoints3D.Add(pt3)
            skl = sk3D.SketchLines3D.AddByTwoPoints(spt1, spt2, False)
            skl.Construction = True
            pt3 = pt2
            pt3.TranslateBy(vz)
            spt3 = sk3D.SketchPoints3D.Add(pt3)
            wa = compDef.WorkAxes.AddByTwoPoints(spt2, spt3)
            wa.Visible = False
            startWorkAxis = wa
            wa.Name = "wa1"
            wpl = compDef.WorkPlanes.AddByNormalToCurve(skl, skl.EndPoint)
            lamp.LookAtPlane(wpl)
            wpl.Visible = False
            wpl.Name = "wpl1"
            startWorkPlane = wpl
            ' currentWorkPlane = wpl
            lamp.LookAtPlane(compDef.WorkPlanes(3))
            sk3D.Visible = False
            Return wa

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function RemoveNextWire() As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim spt2, spt1 As SketchPoint3D
        Dim pt1, pt2 As Point
        Dim sklw As SketchLine3D
        Dim m As Matrix
        Dim wa As WorkAxis
        Dim wpl As WorkPlane
        Dim vz As Vector

        Try
            sk3D = compDef.Sketches3D.Add
            sk3D.Name = String.Concat("sk3D", (qValue + 1).ToString)
            pt1 = compDef.WorkPlanes.Item(3).Plane.IntersectWithLine(compDef.WorkAxes.Item("wa1").Line)
            m = tg.CreateMatrix()
            m.SetToIdentity()
            vz = tg.CreateVector(0, 0, 1)
            m.SetToRotation(driftAngle * qValue, vz, compDef.WorkPoints.Item(1).Point)
            pt1.TransformBy(m)
            spt1 = sk3D.SketchPoints3D.Add(pt1)
            pt2 = pt1
            pt2.TranslateBy(vz)
            spt2 = sk3D.SketchPoints3D.Add(pt2)
            wa = compDef.WorkAxes.AddByTwoPoints(spt1, spt2)
            wa.Visible = False
            nextWorkAxis = wa
            wa.Name = String.Concat("wa", (qValue + 1).ToString)
            wpl = GetNextWorkPlane(wa, 1)
            wpl.Visible = False
            wpl.Name = String.Concat("wpl", (qValue + 1).ToString)
            currentWorkPlane = wpl
            sklw = DrawWireAxis(spt1, wpl)
            ef = RemoveSingleWire(sklw)

            Return ef

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawReferences(q As Integer) As WorkPlane

        Dim spt2, spt1 As SketchPoint3D
        Dim pt1, pt2 As Point

        Dim m As Matrix
        Dim wa As WorkAxis
        Dim wpl As WorkPlane
        Dim vz As Vector
        Dim wpt As WorkPoint


        Try


            If q > 1 Then
                Try
                    sk3D = compDef.Sketches3D.Item(String.Concat("sk3D", (q).ToString))
                    wpt = compDef.WorkPoints.Item(String.Concat("wpt", (q).ToString))
                    wa = compDef.WorkAxes.Item(String.Concat("wa", (q).ToString))
                    wpl = compDef.WorkPlanes.Item(String.Concat("wpl", (q).ToString))
                Catch ex As Exception
                    sk3D = compDef.Sketches3D.Add
                    sk3D.Name = String.Concat("sk3D", (q).ToString)
                    pt1 = compDef.WorkPlanes.Item(3).Plane.IntersectWithLine(compDef.WorkAxes.Item("wa1").Line)
                    m = tg.CreateMatrix()
                    m.SetToIdentity()
                    vz = tg.CreateVector(0, 0, 1)
                    m.SetToRotation(driftAngle * rotationDirection * (q - 1), vz, compDef.WorkPoints.Item(1).Point)
                    pt1.TransformBy(m)
                    spt1 = sk3D.SketchPoints3D.Add(pt1)
                    wpt = compDef.WorkPoints.AddByPoint(spt1)
                    wpt.Name = String.Concat("wpt", (q).ToString)
                    wpt.Visible = False
                    pt2 = pt1
                    pt2.TranslateBy(vz)
                    spt2 = sk3D.SketchPoints3D.Add(pt2)
                    wa = compDef.WorkAxes.AddByTwoPoints(spt1, spt2)
                    wa.Name = String.Concat("wa", (q).ToString)
                    wa.Visible = False
                    wpl = GetNextWorkPlane(wa, q)
                    lamp.LookAtPlane(wpl)
                    wpl.Visible = False
                    wpl.Name = String.Concat("wpl", (q).ToString)
                    sk3D.Visible = False
                End Try

            Else
                wpt = startWorkPoint
                wa = startWorkAxis
                wpl = startWorkPlane
            End If
            currentWorkAxis = wa
            currentWorkPoint = wpt
            currentWorkPlane = wpl
            nextWorkAxis = wa
            lamp.LookAtPlane(compDef.WorkPlanes(3))




            Return wpl

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetNextWorkPlane(wai As WorkAxis, q As Integer) As WorkPlane
        Dim wplo As WorkPlane

        wplo = compDef.WorkPlanes.AddByLinePlaneAndAngle(wai, compDef.WorkPlanes.Item("wpl1"), driftAngle * (q - 1) * rotationDirection)
        wplo.Visible = False
        Return wplo
    End Function
    Function DrawWireAxis(skpt As SketchPoint3D, wpl As WorkPlane) As SketchLine3D
        Dim skl As SketchLine3D
        Try
            sk3D.Visible = False
            sk3D = compDef.Sketches3D.Add
            skl = sk3D.SketchLines3D.AddByTwoPoints(GetSectionPoint(currentWorkPoint, wpl), sptLow, False)
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
    Function EstimateInletBandNumbers(q As Integer, level As Boolean) As Integer
        Dim ansu, r, rMins() As Double
        Dim nb As Integer = 6
        ReDim rMins(nb)
        Dim sign, qi, j As Integer
        ReDim bandNumbers(nb)
        Dim ws As WorkSurface = compDef.WorkSurfaces(1)

        Try
            lamp.LookAtPlane(compDef.WorkPlanes(3))
            lamp.FitView(doku)
            comando.HideAllSurfaces(doku)
            For i = 0 To rMins.Length - 1
                rMins(i) = 9999
                bandNumbers(i) = i
            Next
            Dim antr As Double = Math.IEEERemainder(rotationDirection * (q - 1) * driftAngle + cutAngle - (Math.Abs(CDbl(level)) * tangle * rotationDirection), 2 * Math.PI)
            For i = 1 To DP.q

                ansu = Math.IEEERemainder(2 * Math.PI * (i + 1) * DP.p / DP.q, 2 * Math.PI)
                r = Math.Min(Math.Abs(ansu - antr), Math.Abs(-ansu + antr + 2 * Math.PI))
                If r < rMins(CInt(nb / 2)) Then
                    rMins(CInt(nb / 2)) = r
                    qi = i
                    Try
                        ws.Visible = False
                        ws = compDef.WorkSurfaces.Item(String.Concat("ws", i.ToString))
                        ws.Visible = True
                    Catch ex As Exception

                    End Try

                End If
            Next
            bandNumbers(0) = qi
            For i = 1 To bandNumbers.Length - 1
                sign = CInt(Math.Pow(-1, i + 1))
                If sign > 0 Then
                    Math.DivRem(qi + i + 23, 23, j)
                    j += 1
                Else
                    Math.DivRem(qi - i + 22, 23, j)
                    j += 1
                End If
                bandNumbers(i) = j

            Next
            ws.Visible = False
            Return bandNumbers(0)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function GetBandSurfaces() As WorkSurface
        Try
            Dim nb As Integer = bandNumbers.Length
            Dim s As String
            Dim ws As WorkSurface = compDef.WorkSurfaces(1)
            bandas.Clear()

            For i = 0 To nb - 1
                s = String.Concat("ws", bandNumbers(i).ToString)
                ws = compDef.WorkSurfaces.Item(s)
                bandas.Add(ws)
            Next
            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function

    Function GetSectionPoint(wpti As WorkPoint, wpl As WorkPlane) As SketchPoint3D
        Dim dMax1, dMax2, rMinFront, rMinBack, e As Double
        Dim pt, ptMax1, ptMax2, ptMin1, ptMin2, ptRBack, ptRStamp As Point
        Dim f1, f2, b1, b2, fs, fb As Face
        Dim vnp, vpt, vr, vtc As Vector
        Dim cpt As Point = compDef.WorkPoints.Item(1).Point
        Dim sptRStamp As SketchPoint3D
        Dim ic As IntersectionCurve
        Dim tc As Circle

        Dim s As String
        Dim ws As WorkSurface

        rMinBack = 999999
        rMinFront = rMinBack
        dMax1 = 0
        dMax2 = dMax1
        sk3D = compDef.Sketches3D.Add
        Try
            comando.WireFrameView(doku)
            lamp.FitView(doku)
            vnp = cpt.VectorTo(wpti.Point)
            ptMax1 = wpti.Point
            ptRBack = ptMax1
            ptRStamp = ptRBack
            ptMin1 = ptMax1
            ptMin2 = ptMin1
            ptMax2 = ptMax1
            f1 = compDef.SurfaceBodies.Item(1).Faces.Item(1)
            f2 = f1
            b1 = f1
            b2 = f2
            fb = f1
            fs = f1
            tc = tg.CreateCircle(compDef.WorkPoints.Item(1).Point, tg.CreateUnitVector(0, 0, 1), DP.Dmax / 2)
            For Each ptz As Point In wpl.Plane.IntersectWithCurve(tc)
                vtc = cpt.VectorTo(ptz)
                vr = vnp.CrossProduct(vtc)
                If vr.Z * rotationDirection * (Math.Pow(-1, CDbl(outlet))) > 0 Then
                    ptRReference = ptz

                End If
                sk3D.SketchPoints3D.Add(ptz)
            Next
            For i = 1 To bandNumbers.Length
                s = String.Concat("ws", bandNumbers(i - 1).ToString)
                Try
                    ws = compDef.WorkSurfaces.Item(s)
                    For Each sb As SurfaceBody In ws.SurfaceBodies
                        For Each f As Face In sb.Faces

                            If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                                If Not IsShorterBend(f) Then
                                    If f.Evaluator.Area > 0.18 Then
                                        pt = f.GetClosestPointTo(wpti.Point)
                                        If (pt.Z) * (Math.Pow(-1, CDbl(outlet))) > 0 Then
                                            Try
                                                ic = sk3D.IntersectionCurves.Add(wpl, f)
                                                lamp.HighLighFace(f)
                                                vpt = wpti.Point.VectorTo(GetOuterRadialPoint(ic))
                                                vr = vnp.CrossProduct(vpt)
                                                If vr.Z * (Math.Pow(-1, CDbl(outlet))) * rotationDirection > 0 Then

                                                    e = Math.Pow(pointCutRadiusMax.DistanceTo(ptRReference), 1 / 4) / (Math.Pow(CalculateOutPostionFactor(pointCutRadiusMax), 1 / 4) * Math.Pow(GetStampLength(f, wpl, pointCutRadiusMax, wpti, bandNumbers(i - 1)), 2))
                                                    If e < rMinFront Then
                                                        rMinFront = e
                                                        ptRStamp = pointCutRadiusMax
                                                        fs = f
                                                        currentWorkSurface = ws
                                                        sbValue = bandNumbers(i - 1)
                                                        lamp.HighLighFace(f)
                                                    End If
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If
                                    End If
                                End If



                            End If



                        Next
                    Next
                Catch ex As Exception

                End Try

            Next
            currentWorkSurface.Visible = True
            lamp.HighLighFace(fs)
            stampCurveFace = fs
            sptRStamp = sk3D.SketchPoints3D.Add(ptRStamp)
            sk3D.Visible = False

            Return sptRStamp



        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Function Rod2TangentFace(fi As Face) As Face

        Try


            For i = 0 To tanKeys.Length - 1
                If rodKeys(i) > 0 Then
                    If rodKeys(i) = fi.TransientKey Then
                        For Each sb As SurfaceBody In tangents.SurfaceBodies
                            For Each f As Face In sb.Faces
                                If f.TransientKey = tanKeys(i) Then
                                    Return f
                                End If
                            Next
                        Next

                    End If
                End If

            Next





            Return fi
        Catch ex As Exception
            Return fi
        End Try





    End Function
    Function GetLargerPoints(f As Face, ic As IntersectionCurve) As Point
        Dim ptZ As Point = compDef.WorkPoints.Item(1).Point
        Dim ptR As Point = ptZ
        Dim spt As SketchPoint3D
        Dim pt As Point
        Dim dMax, r As Double
        dMax = 0
        Try

            For Each se As SketchEntity3D In ic.SketchEntities
                se.Construction = True
                If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                    spt = se
                    pt = spt.Geometry

                    If Math.Abs(pt.Z) > Math.Abs(ptZ.Z) Then
                        ptZ = pt
                    End If
                    r = GetRadiusPoint(pt)
                    If (r > dMax) Then
                        dMax = r
                        ptR = pt
                    End If
                End If

            Next
            ptZMax = ptZ
            pointCutRadiusMax = ptR
            Return ptZMax
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetOuterRadialPoint(ic As IntersectionCurve) As Point
        Dim ptZ As Point = compDef.WorkPoints.Item(1).Point
        Dim ptR As Point = ptZ
        Dim spt As SketchPoint3D
        Dim pt As Point
        Dim dMax, r As Double
        dMax = 0
        Try

            For Each se As SketchEntity3D In ic.SketchEntities
                se.Construction = True
                If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                    spt = se
                    pt = spt.Geometry
                    r = GetRadiusPoint(pt)
                    If (r > dMax) Then
                        dMax = r
                        ptR = pt
                    End If
                End If

            Next
            pointCutRadiusMax = ptR
            Return pointCutRadiusMax
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


        Dim ic As IntersectionCurve

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
    Function DrawEntryPoint(fi As Face, wpl As WorkPlane, sptFace As SketchPoint3D, sptCenter As SketchPoint3D) As SketchPoint3D
        Dim ic As IntersectionCurve
        Dim spt, spt2 As SketchPoint3D
        Dim pt, ptMax1, ptMax2, pttf As Point

        Dim f1, f2 As Face
        Dim e As Double = 0


        Dim dMax1, dMax2, dis1, dis2 As Double

        Try

            dMax1 = 0
            dMax2 = 0
            ptMax1 = sptFace.Geometry
            ptMax2 = point2

            dis2 = Math.Abs(sptFace.Geometry.Z) / ((GetRadiusPoint(sptFace.Geometry)))
            f1 = fi

            For Each tf As Face In f1.TangentiallyConnectedFaces
                If tf.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    pttf = tf.GetClosestPointTo(sptCenter.Geometry)
                    If (Math.Abs(sptFace.Geometry.Z - pttf.Z)) < 25 / 10 Then
                        dis1 = Math.Abs(pttf.Z) / (GetRadiusPoint(pttf))
                        sk3D.SketchPoints3D.Add(pttf)
                        If dis1 > dis2 Then

                            'lamp.HighLighFace(tf)
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
                                                lamp.LookAtFace(tf)
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
                End If

            Next


            spt = sk3D.SketchPoints3D.Add(ptMax1)
            spt2 = sk3D.SketchPoints3D.Add(ptMax2)

            'DrawGuidePoints(spt, spt2)
            Return spt
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetStampLength(fi As Face, wpl As WorkPlane, stampPointLocal As Point, wptCenter As WorkPoint, bandNumber As Integer) As Double
        Dim ic As IntersectionCurve
        Dim skl As SketchLine3D

        Dim cpt As Point = compDef.WorkPoints.Item(1).Point
        Dim f1, f2, fout As Face

        Dim d, lMax, dMin, e As Double
        Dim ls As LineSegment
        Dim ptf As Point



        Try
            lMax = 0
            dMin = 9999999

            f1 = GetSimilarFace(fi, bandNumber)
            fout = f1
            Try
                For Each tf As Face In f1.TangentiallyConnectedFaces
                    If tf.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        lamp.HighLighFace(tf)
                        Try
                            ic = sk3D.IntersectionCurves.Add(wpl, tf)

                            For Each se As SketchEntity3D In ic.SketchEntities
                                If se.Type = ObjectTypeEnum.kSketchLine3DObject Then
                                    skl = se
                                    ls = skl.Geometry
                                    ptf = tf.GetClosestPointTo(stampPointLocal)
                                    If skl.StartSketchPoint.Geometry.DistanceTo(stampPointLocal) < skl.EndSketchPoint.Geometry.DistanceTo(stampPointLocal) Then
                                        d = skl.StartSketchPoint.Geometry.DistanceTo(stampPointLocal)
                                    Else
                                        d = skl.EndSketchPoint.Geometry.DistanceTo(stampPointLocal)
                                    End If

                                    e = skl.Length * Math.Exp(-8 * d)
                                    If e > lMax Then
                                        If tf.Equals(fout) Then
                                            lMax = e * 7 / 8
                                        Else
                                            lMax = e
                                        End If

                                        fout = tf
                                    End If
                                    'End If
                                End If
                            Next
                        Catch ex As Exception

                        End Try

                    End If

                Next
                If fout.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    lMax = 3 / 100
                End If


                Return lMax
            Catch ex As Exception
                MsgBox(ex.ToString())
                Return Nothing
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetSimilarFace(fi As Face, bandNumber As Integer) As Face
        Dim area, diff, dMin, dMax, d As Double
        Dim fo As Face = fi
        dMin = 999999
        dMax = 0
        Dim ci As Cylinder = fi.Geometry
        Dim c As Cylinder
        Try
            Dim sb As SurfaceBody = compDef.SurfaceBodies(bandNumber)
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    area = f.Evaluator.Area
                    diff = Math.Abs(area - fi.Evaluator.Area)
                    If diff < dMin Then
                        c = f.Geometry
                        d = Math.Abs((c.AxisVector.AsVector).DotProduct(ci.AxisVector.AsVector))
                        If d > dMax Then
                            dMax = d
                            dMin = diff
                            fo = f
                        End If

                    End If
                End If
            Next
            lamp.HighLighFace(fo)
            Return fo
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function DrawStampPoints(fi As Face, wpl As WorkPlane, sptFace As SketchPoint3D, wptCenter As WorkPoint) As SketchPoint3D
        Dim ic As IntersectionCurve
        Dim sk As Sketch3D = compDef.Sketches3D.Add
        Dim spt As SketchPoint3D
        Dim skl As SketchLine3D
        Dim conti As Boolean
        Dim pt, ptMax1, ptMax2, pttf, ptRMaxLocal As Point
        Dim cpt As Point = compDef.WorkPoints.Item(1).Point
        Dim f1, f2 As Face
        Dim e As Double = 0
        Dim dMax1, dMax2, dis1, dis2, dMin, d, lMax As Double
        Dim vnp, vpt, vr, vtf As Vector
        Dim pl As Plane


        Dim tc As Circle

        Try
            vnp = cpt.VectorTo(wptCenter.Point)
            dMax1 = 0
            dMax2 = 0
            lMax = 3 / 100
            dMin = 99999
            ptMax1 = sptFace.Geometry
            ptMax2 = point2
            tc = tg.CreateCircle(compDef.WorkPoints.Item(1).Point, tg.CreateUnitVector(0, 0, 1), DP.Dmax / 2)
            ptRMaxLocal = sptFace.Geometry
            For Each ptz As Point In wpl.Plane.IntersectWithCurve(tc)
                d = ptz.DistanceTo(sptFace.Geometry)
                If d < dMin Then
                    dMin = d
                    ptRMaxLocal = ptz

                End If
            Next
            dis2 = CalculateOutPostionFactor(sptFace.Geometry) * Math.Pow(lMax, 1 / 4)
            f1 = fi
            Try
                For Each tf As Face In f1.TangentiallyConnectedFaces
                    If tf.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        pttf = tf.GetClosestPointTo(ptRMaxLocal)
                        vpt = wptCenter.Point.VectorTo(pttf)
                        vr = vnp.CrossProduct(vpt)

                        If (pttf.Z * vr.Z) * rotationDirection > 0 Then
                            dis1 = CalculateOutPostionFactor(pttf)
                            sk.SketchPoints3D.Add(pttf)
                            If dis1 > dis2 Then

                                lamp.HighLighFace(tf)
                                Try
                                    ic = sk.IntersectionCurves.Add(wpl, tf)
                                    conti = True
                                    For Each se As SketchEntity3D In ic.SketchEntities
                                        If se.Type = ObjectTypeEnum.kSketchLine3DObject Then
                                            skl = se
                                            lMax = skl.Length
                                            Exit For
                                        End If
                                    Next

                                    For Each se As SketchEntity3D In ic.SketchEntities
                                            se.Construction = True
                                            If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                                                spt = se
                                                pt = spt.Geometry
                                            e = CalculateOutPostionFactor(pt) * Math.Pow(lMax, 1 / 2)
                                            If (e > dMax2) Then
                                                    If (e > dMax1) Then
                                                        dMax2 = dMax1
                                                        dMax1 = e
                                                        ptMax2 = ptMax1
                                                        ptMax1 = pt
                                                        f2 = f1
                                                        f1 = tf
                                                        ptRMin = GetRminPoint(ic, pt)
                                                        lamp.LookAtFace(tf)
                                                        lamp.HighLighFace(tf)
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

                    End If
                Next
                If f1.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    f1 = GetClosestFacePoint(f1, sptFace)
                    ptMax1 = f1.GetClosestPointTo(sptFace.Geometry)
                End If

                spt = sk.SketchPoints3D.Add(ptMax1)

                lamp.HighLighFace(f1)
                lamp.LookAtFace(f1)
                ' spt2 = sk.SketchPoints3D.Add(ptMax2)
                '   lamp.HighLighFace(f1)
                '   lamp.HighLighObject(spt)

                stampPlanarFace = f1
                keyPlanarFace = f1.TransientKey
                sk.Visible = False
                Return spt
            Catch ex As Exception
                MsgBox(ex.ToString())
                Return Nothing
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function SculptRemove(wsi As WorkSurface) As SculptFeature
        Dim sf As SculptFeature = Nothing
        Dim ss As SculptSurface
        Dim asb As SurfaceBody = compDef.SurfaceBodies(sbValue)

        Try
            surfacesSculpt.Clear()
            ss = compDef.Features.SculptFeatures.CreateSculptSurface(wsi, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            surfacesSculpt.Add(ss)
            Try
                sf = compDef.Features.SculptFeatures.Add(surfacesSculpt, PartFeatureOperationEnum.kCutOperation, asb)
            Catch ex As Exception
                CorrectSculpt()
            End Try
            If Not monitor.IsFeatureHealthy(sf) Then
                sf.Delete()
                sf = compDef.Features.SculptFeatures(compDef.Features.SculptFeatures.Count)
            End If

            Return sf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function RemoveExcessStamp(efi As ExtrudeFeature, sbi As Integer) As ExtrudeFeature
        Dim aMax, d As Double
        Dim pt, pti, pto As Point
        Dim fi As Face

        Try
            aMax = 0
            Try
                fi = efi.Faces(1)
                For Each f As Face In efi.Faces
                    If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        If f.Evaluator.Area > aMax Then
                            aMax = f.Evaluator.Area
                            fi = f
                        End If
                    End If

                Next
                lamp.HighLighFace(fi)
            Catch ex As Exception
                fi = efi.Faces(1)

            End Try
            pto = fi.PointOnFace
            Dim ef As ExtrudeFeature = compDef.Features.ExtrudeFeatures(compDef.Features.ExtrudeFeatures.Count)
            RemoveExcessStamp = ef
            Dim s As String = String.Concat("ws", sbi.ToString)
            Dim ws As WorkSurface = compDef.WorkSurfaces.Item(s)
            For Each sb As SurfaceBody In ws.SurfaceBodies
                For Each f As Face In sb.Faces

                    If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        If Not f.TransientKey = keyPlanarFace Then
                            pt = f.GetClosestPointTo(pto)
                            pti = fi.GetClosestPointTo(pt)
                            d = pt.DistanceTo(pti)
                            If d < 12 / 10 Then
                                lamp.FocusFace(f)

                                Try
                                    ef = RemoveFaceMaterial(f, sbi)
                                    If Not monitor.IsFeatureHealthy(ef) Then
                                        ef.Delete()
                                    Else
                                        RemoveExcessStamp = ef
                                        Return RemoveExcessStamp
                                    End If
                                Catch ex As Exception

                                End Try

                            End If
                        End If




                    End If
                Next
            Next

            Return RemoveExcessStamp
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function RemoveFaceMaterial(fc As Face, sb As Integer) As ExtrudeFeature
        Dim ps As PlanarSketch = Nothing
        Try
            ps = doku.ComponentDefinition.Sketches.Add(fc)
            Dim skl As SketchLine
            For Each ed As Edge In fc.Edges
                skl = ps.AddByProjectingEntity(ed)
            Next
            ps.Visible = False
            Dim pro As Profile
            affectedBodies.Clear()
            pro = ps.Profiles.AddForSolid
            Dim edef As ExtrudeDefinition
            edef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
            edef.SetDistanceExtent(2 / 10, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            affectedBodies.Add(compDef.SurfaceBodies(sb))
            edef.AffectedBodies = affectedBodies
            Dim oExtrude As ExtrudeFeature
            oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(edef)

            Return oExtrude
        Catch ex As Exception
            ps.Delete()
            '  MsgBox(ex.ToString())
            Return compDef.Features.ExtrudeFeatures(compDef.Features.ExtrudeFeatures.Count)
        End Try

    End Function
    Function CorrectSculpt() As SculptFeature
        Dim sf As SculptFeature
        Try

            sf = compDef.Features.SculptFeatures(compDef.Features.SculptFeatures.Count)

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return Nothing
    End Function
    Function CombineBodies() As CombineFeature
        Dim cf As CombineFeature = Nothing
        If compDef.SurfaceBodies.Count > 1 Then
            surfaceBodies.Clear()

            For index = 2 To doku.ComponentDefinition.SurfaceBodies.Count
                surfaceBodies.Add(compDef.SurfaceBodies.Item(index))
            Next
            Try
                cf = doku.ComponentDefinition.Features.CombineFeatures.Add(compDef.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)

            Catch ex As Exception


            End Try
        End If

        ' cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)

        Return cf
    End Function

    Function StampLetters(q As Integer) As CutFeature
        Dim cf As CutFeature = Nothing
        Dim path As String = projectManager.ActiveDesignProject.WorkspacePath

        Try
            For i = 0 To 1
                Try
                    sk3D = compDef.Sketches3D.Item(String.Concat("sw", ((q * 2) - 1).ToString))
                    If i = 0 Then
                        Continue For
                    Else
                        cf = StampOneLetter(q, 1 - i)

                    End If

                Catch ex As Exception
                    cf = StampOneLetter(q, 1 - i)
                End Try
            Next
            Return cf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function StampOneLetter(q As Integer, i As Integer) As CutFeature
        Dim cf As CutFeature = Nothing
        Dim skpt As SketchPoint3D
        Dim path As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim bffn As String
        Try
            EstimateInletBandNumbers(q, CBool(i))
            outlet = CBool(i)
            skpt = GetSectionPoint(currentWorkPoint, currentWorkPlane)
            qValue = q
            skpt = DrawStampPoints(stampCurveFace, currentWorkPlane, skpt, currentWorkPoint)

            bffn = String.Concat(path, "\Iteration", iteration.ToString, "\Band", sbValue.ToString, ".ipt")
            banda = app.Documents.Open(bffn, True)
            banda.Activate()
            estampa = New Stanzer(banda)
            cf = estampa.EmbossColumnNumber(q, stampPlanarFace, skpt)
            If monitor.IsFeatureHealthy(cf) Then

                banda.Save()
                banda.Close()
                doku.Activate()
                currentWorkSurface.Visible = False
                doku.Update()

                If i = 1 Then
                    compDef.Sketches3D.Item(compDef.Sketches3D.Count).Name = String.Concat("sw", ((2 * q) - 1).ToString)
                Else
                    compDef.Sketches3D.Item(compDef.Sketches3D.Count).Name = String.Concat("sw", ((2 * q)).ToString)
                End If
            End If
            done = estampa.done
            Return cf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function

    Function GetRminPoint(ic As IntersectionCurve, pti As Point) As Point
        Dim pt As Point
        Dim spt As SketchPoint3D
        Dim e, dMin As Double
        Try
            dMin = 9999
            For Each se As SketchEntity3D In ic.SketchEntities
                se.Construction = True
                If se.Type = ObjectTypeEnum.kSketchPoint3DObject Then
                    spt = se
                    pt = spt.Geometry
                    If pt.DistanceTo(pti) > 1 / 128 Then
                        e = GetRadiusPoint(pt)
                        If (e < dMin) Then
                            dMin = e
                            ptRMin = pt
                        End If
                    End If
                End If
            Next
            Return ptRMin
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function GetClosestFacePoint(fi As Face, spt As SketchPoint3D) As Face
        Dim d, dMin, eMin As Double
        Dim fMin As Face = fi
        dMin = 999999
        eMin = dMin
        For Each f As Face In fi.TangentiallyConnectedFaces
            If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                d = f.GetClosestPointTo(spt.Geometry).DistanceTo(spt.Geometry)
                If d < dMin Then
                    dMin = d
                    fMin = f
                    ptRMin = f.GetClosestPointTo(currentWorkPoint.Point)

                End If

            End If
        Next
        Return fMin
    End Function
    Function CalculateOutPostionFactor(pt As Point) As Double

        Return (Math.Pow(GetRadiusPoint(pt), 1 / 2) * Math.Pow(TorusRadius(pt), 2)) / (1 / (1 + Math.Exp(-4 * Cr / 5 + Math.Abs(pt.Z))))
    End Function
    Function TorusRadius(pt As Point) As Double

        Return Math.Pow(Math.Pow(Math.Pow(Math.Pow(pt.X, 2) + Math.Pow(pt.Y, 2), 1 / 2) - Tr, 2) + Math.Pow(pt.Z, 2), 1 / 2)
    End Function
    Function IsShorterBend(fi As Face) As Boolean
        Dim smallOval As Boolean
        Dim d As Double
        Dim angle, e As Double
        Dim cyl As Cylinder = fi.Geometry
        Dim r As Double = cyl.Radius
        Dim ptc As Point = cyl.BasePoint
        Try

            Dim edMax As Edge = GetMajorEdge(fi)
            Dim edMin As Edge = GetMinorEdge(fi)
            Dim pt As Point = edMax.GetClosestPointTo(ptc)
            Dim pt2 As Point = secondEdge.GetClosestPointTo(ptc)

            e = pt.DistanceTo(pt2)

            angle = 2 * Math.Asin(e / (2 * r))
            d = edMax.StartVertex.Point.DistanceTo(edMax.StopVertex.Point)

            If (angle < 1.2 * 0.9) And (d < 45 / 10) Then
                smallOval = True
            Else
                smallOval = False
            End If

            Return smallOval
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
        e1 = f.Edges(1)
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
        secondEdge = e2
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
    Function DrawGuidePoints(skpt1 As SketchPoint3D, skpt2 As SketchPoint3D) As SketchPoint3D
        Dim spt As SketchPoint3D = skpt1
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
                Math.DivRem(qValue + 1, 5, modulo)
            Else
                Math.DivRem(qValue, 5, modulo)
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
        Dim ef As ExtrudeFeature = Nothing
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
            lamp.LookAtPlane(wpl)
            ps = doku.ComponentDefinition.Sketches.Add(wpl)
            ps.Name = String.Concat("wire", (qValue + 1).ToString)
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
        Dim pMax As Point = Nothing
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
        Dim p As Parameter = Nothing
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
