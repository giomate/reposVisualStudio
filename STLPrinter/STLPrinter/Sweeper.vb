Imports Inventor

Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory
Public Class Sweeper

    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk, curvesSketch As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, coneLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile

    Public trobinaCurve As Curves3D
    Dim palito As RodMaker

    Public wp1, wp2, wp3, wpConverge, wptHigh, wptLow As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, convergePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM, freeCollisionAngle As Double
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
    Dim ringLine As SketchLine3D
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, cylinderFace, sideMajorFace, sideMinorFace As Face
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
    Dim bandsurfaces, tangentSurfaces As WorkSurface
    Dim surfacesScuplt, facesToDelete As ObjectCollection
    Dim foldFeature As FoldFeature
    Public sections, endPoints, rails, caras, surfaceBodies, coneAxes, highPoints, lowPoints As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim adjuster As SketchAdjust
    Dim PreviousFaceKeys(), combinedKeys() As Long

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
        surfacesScuplt = app.TransientObjects.CreateObjectCollection
        facesToDelete = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)

        gap1CM = 3 / 10

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

        done = False

    End Sub
    Function GetMajorEdge(wsi As WorkSurface) As Edge
        Dim e1, e2, e3 As Edge
        Dim maxe1, maxe2, maxe3 As Double
        maxe1 = 0
        maxe2 = 0
        maxe3 = 0
        e1 = wsi.SurfaceBodies(1).Faces(1).Edges(1)
        e2 = e1
        e3 = e2
        For Each sb As SurfaceBody In wsi.SurfaceBodies
            For Each f As Face In sb.Faces
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
            Next
        Next

        lamp.HighLighObject(e3)
        lamp.HighLighObject(e2)
        lamp.HighLighObject(e1)
        bendEdge = e3
        minorEdge = e2
        majorEdge = e1
        Return e1
    End Function
    Function GetMinorEdge(wsi As WorkSurface) As Edge
        Dim e1, e2, e3 As Edge
        Dim min1, min2, min3 As Double
        min1 = 999999
        min2 = 99999
        min3 = 999999
        e1 = wsi.SurfaceBodies(1).Faces(1).Edges(1)
        e2 = e1
        e3 = e2
        For Each sb As SurfaceBody In wsi.SurfaceBodies
            For Each f As Face In sb.Faces
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
                    Else
                        min3 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                        e3 = ed

                    End If

                Next
            Next
        Next



        minorEdge = e1

        Return e1
    End Function
    Public Function RemoveAll(s As String) As ExtrudeFeature
        compDef = doku.ComponentDefinition
        Dim cf As CombineFeature
        Dim ef As ExtrudeFeature
        Dim sf As SculptFeature
        tangentSurfaces = compDef.WorkSurfaces.Item("tangents")
        Dim ws As WorkSurface = tangentSurfaces
        doku.Update2(True)
        Try
            cf = compDef.Features.CombineFeatures(1)
            If monitor.IsFeatureHealthy(cf) Then
                ef = RemoveRemaining(cf)
                If monitor.IsFeatureHealthy(ef) Then
                    Return ef
                End If

            End If
        Catch ex As Exception
            Try
                sf = compDef.Features.SculptFeatures(1)
                cf = RemoveCombine(DeriveBand(s), s)
                If monitor.IsFeatureHealthy(cf) Then
                    ef = RemoveRemaining(cf)
                    If monitor.IsFeatureHealthy(ef) Then
                        Return ef
                    End If

                End If
            Catch ex2 As Exception
                If monitor.IsFeatureHealthy(SculptRemove(ws)) Then

                    cf = RemoveCombine(DeriveBand(s), s)
                    If monitor.IsFeatureHealthy(cf) Then
                        ef = RemoveRemaining(cf)
                        If monitor.IsFeatureHealthy(ef) Then
                            Return ef
                        End If

                    End If
                End If
            End Try
        End Try




        Return Nothing
    End Function
    Function GetLargerFaceShell() As FaceShell
        Dim fsMax As FaceShell = doku.ComponentDefinition.SurfaceBodies(1).FaceShells(1)

        Dim nFaceMax As Integer = 0

        For Each fsi As FaceShell In doku.ComponentDefinition.SurfaceBodies(1).FaceShells
            If fsi.Faces.Count > nFaceMax Then
                nFaceMax = fsi.Faces.Count
                fsMax = fsi
            End If
        Next
        Return fsMax
    End Function
    Public Function DocUpdate(docu As PartDocument) As PartDocument
        doku = docu

        compDef = doku.ComponentDefinition
        lamp = New Highlithing(doku)

        Return doku
    End Function
    Function CutSmallBodies() As Integer
        Try
            Dim vMax As Double = 0
            Dim sbMax As SurfaceBody = compDef.SurfaceBodies(1)
            Dim fsMax As FaceShell = sbMax.FaceShells(1)

            Dim ws As WorkSurface
            Dim sf As SculptFeature
            Dim nFaceMax As Integer = 0
            Dim nFaceShells, k As Integer
            Dim fs As FaceShell
            k = 1

            doku.Update2(True)
            comando.WireFrameView(doku)

            If sbMax.FaceShells.Count > 1 Then

                nFaceShells = doku.ComponentDefinition.SurfaceBodies(1).FaceShells.Count
                For i = 1 To nFaceShells * 2

                    fs = doku.ComponentDefinition.SurfaceBodies(1).FaceShells(k)
                    fsMax = GetLargerFaceShell()
                    If Not fs.Equals(fsMax) Then
                        facesToDelete.Clear()
                        For Each f As Face In fs.Faces
                            lamp.HighLighFace(f)
                            facesToDelete.Add(f)
                        Next
                        If monitor.IsFeatureHealthy(CreateNonParametricBody()) Then
                            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)

                            sf = SculptRemove(ws)
                            If monitor.IsFeatureHealthy(sf) Then
                                doku.Update2(True)
                                nFaceShells = doku.ComponentDefinition.SurfaceBodies(1).FaceShells.Count
                                If nFaceShells < 2 Then
                                    Exit For
                                Else
                                    k = 1
                                    DocUpdate(doku)
                                End If
                            Else
                                fs = doku.ComponentDefinition.SurfaceBodies(1).FaceShells(k)
                                If fs.IsVoid Then
                                    sf.Delete(True, True, True)
                                    sf = SculptAdd(ws)
                                    If monitor.IsFeatureHealthy(sf) Then
                                        doku.Update2(True)
                                        nFaceShells = doku.ComponentDefinition.SurfaceBodies(1).FaceShells.Count
                                        If nFaceShells < 2 Then
                                            Exit For
                                        Else
                                            k = 1
                                            DocUpdate(doku)
                                        End If
                                    Else
                                        MsgBox(sf.ToString)
                                        Return Nothing

                                    End If
                                End If

                            End If

                        End If
                    Else
                        k = k + 1

                    End If
                    If doku.ComponentDefinition.SurfaceBodies(1).FaceShells.Count < 2 Then
                        Exit For
                    End If
                Next



            End If



            Return doku.ComponentDefinition.SurfaceBodies(1).FaceShells.Count
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CreateNonParametricBody() As NonParametricBaseFeature
        Dim np As NonParametricBaseFeature
        Try
            Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition

            npDef.BRepEntities = facesToDelete
            npDef.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
            Return np
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function SculptRemove(wsi As WorkSurface) As SculptFeature
        Dim sf As SculptFeature
        Dim ss As SculptSurface

        Try
            surfacesScuplt.Clear()
            ss = compDef.Features.SculptFeatures.CreateSculptSurface(wsi, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            surfacesScuplt.Add(ss)
            Try
                sf = compDef.Features.SculptFeatures.Add(surfacesScuplt, PartFeatureOperationEnum.kCutOperation)
            Catch ex As Exception
                sf = CorrectSculpt(ss)
            End Try
            If monitor.IsFeatureHealthy(sf) Then
                CombineBodies()
            End If

            Return sf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function SculptAdd(wsi As WorkSurface) As SculptFeature
        Dim sf As SculptFeature
        Dim ss As SculptSurface

        Try
            surfacesScuplt.Clear()
            ss = compDef.Features.SculptFeatures.CreateSculptSurface(wsi, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            surfacesScuplt.Add(ss)
            Try
                sf = compDef.Features.SculptFeatures.Add(surfacesScuplt, PartFeatureOperationEnum.kJoinOperation)
            Catch ex As Exception
                sf = CorrectSculpt(ss)
            End Try
            If monitor.IsFeatureHealthy(sf) Then
                CombineBodies()
            End If

            Return sf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CorrectSculpt(ss As SculptSurface) As SculptFeature
        Dim rf As RevolveFeature = compDef.Features.RevolveFeatures(1)
        Dim pf As PartFeature = compDef.Features(1)
        Dim sf As SculptFeature = Nothing
        Dim n As Integer = compDef.Features.RevolveFeatures.Count
        Dim s As String
        Dim ef As ExtrudeFeature

        Dim lf As LoftFeature
        Dim d As Double

        Try
            For i = 1 To n
                rf = compDef.Features.RevolveFeatures.Item(n - i + 1)
                If nombrador.ContainCorto(rf.Name) Then
                    Try
                        s = String.Concat("egg", nombrador.GetRodNumber(rf).ToString)
                        ef = compDef.Features.ExtrudeFeatures.Item(s)

                        d = GetExtendEggDistance(ef)
                        ef.Definition.SetDistanceExtent(d * 3 / 4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
                        rf.Operation = PartFeatureOperationEnum.kNewBodyOperation
                        ' doku.Update2(True)
                        Try
                            sf = compDef.Features.SculptFeatures.Add(surfacesScuplt, PartFeatureOperationEnum.kCutOperation, compDef.SurfaceBodies(1))
                            Exit For

                        Catch ex As Exception

                            If ef.Name = "egg2" Then
                                rf.Operation = PartFeatureOperationEnum.kJoinOperation
                                lf = compDef.Features.LoftFeatures.Item("cone2")
                                lf.Definition.Operation = PartFeatureOperationEnum.kNewBodyOperation
                                ef.Operation = PartFeatureOperationEnum.kNewBodyOperation
                                rf.Operation = PartFeatureOperationEnum.kNewBodyOperation
                                Try
                                    sf = compDef.Features.SculptFeatures.Add(surfacesScuplt, PartFeatureOperationEnum.kCutOperation, compDef.SurfaceBodies(1))
                                    Exit For

                                Catch ex3 As Exception

                                End Try
                            Else
                                ef.Definition.SetDistanceExtent(d, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
                                rf.Operation = PartFeatureOperationEnum.kJoinOperation

                            End If

                        End Try
                    Catch ex As Exception

                    End Try
                End If


            Next


        Catch ex As Exception
            MsgBox(ex.ToString() & n.ToString)
            Return Nothing
        End Try
        CombineBodies()
        Return sf
    End Function
    Function GetExtendEggDistance(ef As ExtrudeFeature) As Double
        Dim d, dmin, dMax, Amin As Double



        dmin = 9999
        dMax = 0
        Amin = 9999
        Dim n As Integer = ef.SideFaces.Count


        Try
            For Each p As Parameter In ef.Parameters
                d = p._Value
                If d > dMax Then
                    dMax = d
                End If
            Next




            Return dMax
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function ReorderPane(rf As RevolveFeature, sf As SculptFeature) As Integer
        Dim m As Integer
        Dim pf As PartFeature
        Dim pane As BrowserPane

        Dim dependantNode As BrowserNode
        Dim lastNode As BrowserNode

        Dim revolNode As BrowserNode
        Try
            m = compDef.Features.Count
            pane = doku.BrowserPanes("Model")
            ' nodeDef = doku.BrowserPanes.GetNativeBrowserNodeDefinition(rf)

            revolNode = pane.GetBrowserNodeFromObject(rf)
            lastNode = pane.GetBrowserNodeFromObject(sf)

            For j = 1 To m
                pf = compDef.Features.Item(m - j + 1)
                If pf.Suppressed Then

                    dependantNode = pane.GetBrowserNodeFromObject(pf)
                    pane.Reorder(lastNode, False, revolNode, dependantNode)
                    ' pane.Refresh()
                    Exit For
                End If
            Next
            m = compDef.Features.Count
            For j = 1 To m
                pf = compDef.Features(j)
                If pf.Suppressed Then
                    pf.Suppressed = False
                End If
            Next
            Return m
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Function DeriveBand(s As String) As DerivedPartComponent
        Dim derivedDefinition As DerivedPartUniformScaleDef
        Dim dpc As DerivedPartComponent
        Try
            derivedDefinition = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateUniformScaleDef(s)
            derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
            derivedDefinition.ScaleFactor = 1
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
            dpc = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)
            Return dpc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


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
    Function RemoveCombine(dp As DerivedPartComponent, s As String) As CombineFeature

        Dim t As Integer = compDef.SurfaceBodies.Count
        Dim cf As CombineFeature = Nothing

        Dim dpc As DerivedPartComponent = dp
        Dim n As Integer = compDef.SurfaceBodies(1).Faces.Count
        Dim factor As Double
        Try

            surfaceBodies.Clear()

            surfaceBodies.Add(compDef.SurfaceBodies(t))
            If doku.ComponentDefinition.SurfaceBodies.Count > 1 Then
                Try
                    cf = compDef.Features.CombineFeatures.Add(compDef.SurfaceBodies(1), surfaceBodies, PartFeatureOperationEnum.kCutOperation)
                Catch ex As Exception
                    For i = 3 To 8
                        dpc.Delete()
                        factor = (Math.Pow(10, 8 - i + 3) - 1) / (Math.Pow(10, 8 - i + 3))
                        dpc = DeriveBandFactor(s, factor)
                        surfaceBodies.Clear()
                        surfaceBodies.Add(compDef.SurfaceBodies(t))
                        Try
                            cf = compDef.Features.CombineFeatures.Add(compDef.SurfaceBodies(1), surfaceBodies, PartFeatureOperationEnum.kCutOperation)
                            Exit For
                        Catch ex2 As Exception

                        End Try
                    Next


                End Try

            End If
            ' cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)
            Return cf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DeriveBandFactor(s As String, d As Double) As DerivedPartComponent
        Dim derivedDefinition As DerivedPartUniformScaleDef
        Dim dpc As DerivedPartComponent
        Try
            derivedDefinition = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateUniformScaleDef(s)
            derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
            derivedDefinition.ScaleFactor = d
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
            dpc = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)
            Return dpc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function RemoveRemaining(cf As CombineFeature) As ExtrudeFeature
        Try
            Dim n As Integer = cf.Faces.Count
            ReDim combinedKeys(n - 1)
            Dim found As Boolean = False

            Dim ef As ExtrudeFeature = compDef.Features.ExtrudeFeatures(1)
            RemoveRemaining = ef
            For Each f As Face In cf.Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    '  lamp.LookAtFace(f)

                    Try
                        ef = RemoveFaceMaterial(f, GetClosestRefFace(f))
                        If Not monitor.IsFeatureHealthy(ef) Then
                            ef.Delete()
                        Else
                            RemoveRemaining = ef
                        End If
                    Catch ex As Exception

                    End Try

                End If

            Next
            Return RemoveRemaining
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Function GetClosestTangentFace(fi As Face) As Face
        Dim v1, v2, v3 As Vector
        Dim pl1, pl2 As Plane
        Dim dMax, d As Double
        dMax = 0
        pl1 = fi.Geometry
        v1 = pl1.Normal.AsVector
        v1.Normalize()
        GetClosestTangentFace = fi
        For Each sb As SurfaceBody In tangentSurfaces.SurfaceBodies
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    pl2 = f.Geometry
                    v2 = pl2.Normal.AsVector
                    v2.Normalize()
                    v3 = v1.CrossProduct(v2)
                    d = v3.Length / (pl1.RootPoint.DistanceTo(pl2.RootPoint) + 0.000001)
                    If d > dMax Then
                        dMax = d
                        GetClosestTangentFace = f
                    End If
                End If

            Next
        Next

        Return GetClosestTangentFace
    End Function
    Function GetClosestRefFace(fi As Face) As Face
        Dim v1, v2, v3 As Vector
        Dim pl1, pl2 As Plane
        Dim dMax, d As Double
        Dim fb As Face
        dMax = 0
        pl1 = fi.Geometry
        v1 = pl1.Normal.AsVector
        v1.Normalize()
        fb = fi
        For Each sb As SurfaceBody In tangentSurfaces.SurfaceBodies
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    pl2 = f.Geometry
                    v2 = pl2.Normal.AsVector
                    v2.Normalize()
                    v3 = v1.CrossProduct(v2)
                    d = Math.Abs(v2.DotProduct(v1)) / (pl1.RootPoint.DistanceTo(pl2.RootPoint) + 0.000001)
                    If d > dMax Then
                        dMax = d
                        fb = f
                    End If
                End If

            Next
        Next
        lamp.HighLighFace(fb)
        Return fb
    End Function
    Function RemoveFaceMaterial(fc As Face, ft As Face) As ExtrudeFeature
        Dim ps As PlanarSketch
        ps = doku.ComponentDefinition.Sketches.Add(ft)

        Dim skl As SketchLine
        For Each ed As Edge In ft.Edges
            skl = ps.AddByProjectingEntity(ed)
        Next

        ps.Visible = False
        Dim pro As Profile


        pro = ps.Profiles.AddForSolid
        Dim oExtrudeDef As ExtrudeDefinition
        oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
        oExtrudeDef.SetDistanceExtent(2 / 10, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude As ExtrudeFeature
        oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)



        Return oExtrude
    End Function
    Function GetFarestVertexFace(fc As Face, ft As Face) As Vertex
        Dim fv As Vertex
        Dim d, e, dMax, dMin As Double
        dMax = 0
        dMin = 9999
        Dim ptCommon As Point = fc.Vertices(1).Point
        Try
            For Each veri As Vertex In fc.Vertices
                For Each vero As Vertex In ft.Vertices
                    d = veri.Point.DistanceTo(vero.Point)
                    If d < dMin Then
                        dMin = d
                        ptCommon = vero
                    Else
                        e = veri.Point.DistanceTo(ptCommon)
                        If e > dMax Then
                            dMax = e
                            fv = veri
                        End If

                    End If


                Next
            Next

            lamp.HighLighObject(fv)
            Return fv
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
End Class
