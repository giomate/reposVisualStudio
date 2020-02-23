Imports Inventor
Public Class Cortador
    Public doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Dim refLine, firstLine, secondLine, thirdLine, lastLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim curve3D As Curves3D
    Dim monitor As DesignMonitoring
    Dim medico As DesignDoctor
    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3 As Point
    Dim tg As TransientGeometry
    Dim gap1 As Double
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines, bigFaces As ObjectCollection
    Dim comando As Commands
    Dim mainSketch As InitSketcher
    Dim cutProfile As Profile
    Dim feature As FaceFeature
    Dim foldFeature As FoldFeature
    Dim cutFeature As CutFeature
    Dim bendLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim workFace, adjacentFace, bendFace, cutBend, frontBendFace, cutFace, twistFace, nextworkFace As Face
    Dim edgeBand, majorEdge, minorEdge, bendEdge As Edge
    Dim adjacentEdge, cutEdge1, cutEdge2, CutEdge3, leadingEdge, followEdge As Edge
    Dim bendAngle As DimensionConstraint
    Public initCutLine, lastCutLine As SketchLine
    Public cutLine3D As SketchLine3D

    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim lamp As Highlithing
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        compDef = doku.ComponentDefinition
        sheetMetalFeatures = compDef.Features
        curve3D = New Curves3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        medico = New DesignDoctor(doku)
        mainSketch = New InitSketcher(doku)
        lamp = New Highlithing(doku)
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        bigFaces = app.TransientObjects.CreateObjectCollection

        done = False
    End Sub
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
        lamp.HighLighObject(e4)
        lamp.HighLighObject(e2)
        lamp.HighLighObject(e1)
        cutEdge1 = e1
        cutEdge2 = e2
        CutEdge3 = e4
        Return e4
    End Function
    Function GetInitFaceCutEdges(f As Face) As Edge


        Dim e1, e2, e3, e4 As Edge
        Dim maxe1, maxe2, maxe3, maxe4, d As Double
        Dim v1, v2, v3 As Vector
        Dim sl2d As SketchLine
        sl2d = foldFeature.Definition.BendLine
        v1 = sl2d.Geometry3d.Direction.AsVector
        maxe1 = 0
        maxe2 = 0
        maxe3 = 0
        maxe4 = 0
        e1 = f.Edges.Item(1)
        e2 = e1
        e3 = e2
        e4 = e3
        lamp.HighLighFace(cutFace)
        For Each ed As Edge In f.Edges
            v2 = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point).AsUnitVector.AsVector
            v3 = v1.CrossProduct(v2)
            d = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) * v3.Length
            If d > maxe3 Then
                If d > maxe2 Then
                    If d > maxe1 Then
                        maxe4 = maxe3
                        e4 = e3
                        maxe3 = maxe2
                        e3 = e2
                        maxe2 = maxe1
                        e2 = e1
                        maxe1 = d
                        e1 = ed
                    Else
                        maxe4 = maxe3
                        e4 = e3
                        maxe3 = maxe2
                        e3 = e2
                        maxe2 = d
                        e2 = ed
                    End If
                Else
                    maxe4 = maxe3
                    e4 = e3
                    maxe3 = d
                    e3 = ed

                End If
            Else
                maxe4 = d
                e4 = ed
            End If
        Next

        lamp.HighLighObject(e3)
        lamp.HighLighObject(e2)
        lamp.HighLighObject(e1)
        cutEdge1 = e1
        cutEdge2 = e2
        CutEdge3 = e3
        Return e1
    End Function
    Function GetInitCutFace(ff As FoldFeature) As Face
        Dim maxArea1, maxArea2, maxArea3 As Double

        Dim maxface1, maxface2, maxface3 As Face

        foldFeature = ff
        maxface2 = compDef.Bends.Item(compDef.Bends.Count).BackFaces.Item(1)
        maxface1 = maxface2
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
        bendFace = maxface2
        cutBend = maxface1
        For Each f As Face In cutBend.TangentiallyConnectedFaces

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
        cutFace = maxface2
        lamp.HighLighFace(cutFace)
        Return cutFace
    End Function
    Function LastCutProfil(firstLine As SketchLine3D, secondLine As SketchLine3D) As Profile
        Try
            Dim pro As Profile
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
            lastCutLine = cl
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
    Function GetLastCutFace(ff As FoldFeature) As Face
        Try
            foldFeature = ff
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

            Return cutFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function InitCutProfil(rl As SketchLine3D) As Profile
        Try
            'lamp.HighLighFace(maxface1)
            Dim ps As PlanarSketch
            ps = compDef.Sketches.Add(cutFace)
            Dim sl, fl, l, r, u, pfl, cl As SketchLine
            fl = ps.AddByProjectingEntity(rl)
            fl.Construction = True
            l = ps.AddByProjectingEntity(GetInitFaceCutEdges(cutFace))
            u = ps.AddByProjectingEntity(cutEdge2)

            r = ps.AddByProjectingEntity(CutEdge3)
            pfl = ps.SketchLines.AddByTwoPoints(fl.Geometry.StartPoint, fl.Geometry.EndPoint)
            pfl.Construction = True
            cl = ps.SketchLines.AddByTwoPoints(pfl.StartSketchPoint.Geometry, pfl.EndSketchPoint.Geometry)
            ps.GeometricConstraints.AddCoincident(cl.StartSketchPoint, r)
            If r.EndSketchPoint.Geometry.DistanceTo(pfl.Geometry.StartPoint) < r.StartSketchPoint.Geometry.DistanceTo(pfl.Geometry.StartPoint) Then
                ps.GeometricConstraints.AddCoincident(r.EndSketchPoint, cl)
            Else
                ps.GeometricConstraints.AddCoincident(r.StartSketchPoint, cl)
            End If

            ps.GeometricConstraints.AddCoincident(cl.EndSketchPoint, l)
            ps.GeometricConstraints.AddParallel(cl, pfl)

            initCutLine = cl


            cutProfile = ps.Profiles.AddForSolid
            sk3D = compDef.Sketches3D.Add()
            cutLine3D = sk3D.Include(cl)
            sk3D.Name = "initCutLine"
            Return cutProfile
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Public Function CutLastFace(pr As Profile) As CutFeature
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
            cutFeature = oFaceFeature
            If monitor.IsFeatureHealthy(cutFeature) Then
                cutFeature.Name = "lastCut"
            End If
            Return oFaceFeature
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing

        End Try


    End Function
    Public Function MakeLastCut(fl As SketchLine3D, sl As SketchLine3D) As CutFeature
        Try
            Dim pr As Profile
            pr = LastCutProfil(fl, sl)
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
            cutFeature = oFaceFeature
            Return oFaceFeature
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing

        End Try


    End Function
    Function CutInitFace(l As SketchLine3D) As CutFeature
        Try
            Dim pr As Profile
            pr = InitCutProfil(l)
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
            cutFeature = oFaceFeature
            If monitor.IsFeatureHealthy(cutFeature) Then
                cutFeature.Name = "initCut"
            End If
            Return oFaceFeature
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing

        End Try


    End Function

End Class
