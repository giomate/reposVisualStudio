﻿Imports Inventor
Imports DrawingMainSketch
Imports GetInitialConditions
Imports DrawInitialSketch
Public Class InitFold
    Public doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Dim refLine, firstLine, secondLine, thirdLine, lastLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim curve3D As Curve3D
    Dim monitor As DesignMonitoring
    Dim medico As DesignDoctor
    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3 As Point
    Dim tg As TransientGeometry
    Dim gap1 As Double
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim mainSketch As InitSketcher
    Dim pro As Profile
    Dim feature As FaceFeature
    Dim bendLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim workface As Face
    Dim edgeBand, majorEdge, minorEdge, bendEdge As Edge
    Dim bendAngle As DimensionConstraint
    Dim folded As FoldFeature

    Dim features As SheetMetalFeatures
    Dim lamp As Highlithing
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        compDef = doku.ComponentDefinition
        features = compDef.Features
        curve3D = New Curve3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        medico = New DesignDoctor(doku)
        mainSketch = New InitSketcher(doku)
        lamp = New Highlithing(doku)
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection

        done = False
    End Sub


    Public Function MakeFirstFold(refDoc As FindReferenceLine) As Boolean
        Dim b As Boolean
        Try


            sk3D = mainSketch.StartDrawingTranslated(refDoc, 1)
            If mainSketch.done Then
                doku.ComponentDefinition.Sketches3D.Item("s0").Visible = False
                If DrawBandStripe().Count > 0 Then
                    If MakeStartingFace(pro).GetHashCode > 0 Then
                        If GetBendLine().Length > 0 Then
                            mainWorkPlane.Visible = False
                            If GetFoldingAngle().Parameter._Value > 0 Then
                                comando.MakeInvisibleSketches(doku)
                                comando.MakeInvisibleWorkPlanes(doku)
                                If monitor.IsFeatureHealthy(FoldBand()) Then
                                    doku.Save2(True)
                                    done = 1
                                    Return True
                                End If


                            End If

                        End If

                    End If
                End If

            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
    Public Function MakeFirstFold(refDoc As FindReferenceLine, q As Integer) As Boolean
        Dim b As Boolean
        Try


            sk3D = mainSketch.DrawMainSketch(refDoc, q)
            If mainSketch.done Then
                doku.ComponentDefinition.Sketches3D.Item("s0").Visible = False
                If DrawBandStripe().Count > 0 Then
                    If MakeStartingFace(pro).GetHashCode > 0 Then
                        If GetWorkFace().Evaluator.Area > 0 Then
                            If GetBendLine(workface, mainSketch.thirdLine).Length > 0 Then
                                mainWorkPlane.Visible = False
                                If GetFoldingAngle().Parameter._Value > 0 Then
                                    comando.MakeInvisibleSketches(doku)
                                    comando.MakeInvisibleWorkPlanes(doku)
                                    bandLines = mainSketch.bandLines
                                    folded = FoldBand(bandLines.Count)
                                    folded.Name = "f1"
                                    doku.Update2(True)
                                    If monitor.IsFeatureHealthy(folded) Then
                                        doku.Save2(True)
                                        done = 1
                                        Return True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
    Function DrawBandStripe() As Profile
        Dim ps As PlanarSketch

        Dim p1, p2, p3 As SketchPoint
        Dim v As Vector2d
        Dim r As SketchEntitiesEnumerator
        Dim ln, l2d As SketchLine
        Try
            mainWorkPlane = doku.ComponentDefinition.WorkPlanes.AddByThreePoints(mainSketch.firstLine.StartPoint, mainSketch.firstLine.EndPoint, mainSketch.secondLine.EndPoint)
            ps = doku.ComponentDefinition.Sketches.Add(mainWorkPlane)
            p1 = ps.AddByProjectingEntity(mainSketch.firstLine.StartPoint)
            p2 = ps.AddByProjectingEntity(mainSketch.firstLine.EndPoint)
            p3 = ps.AddByProjectingEntity(mainSketch.secondLine.EndPoint)
            ln = ps.AddByProjectingEntity(mainSketch.bandLines.Item(2))
            v = ln.Geometry.Direction.AsVector
            v.ScaleBy(-3 / 10)
            p1.MoveBy(v)
            l2d = ps.SketchLines.AddByTwoPoints(p1, p2.Geometry)
            ps.GeometricConstraints.AddPerpendicular(l2d, ln)
            Dim dc As DimensionConstraint
            dc = ps.DimensionConstraints.AddTwoPointDistance(p1, l2d.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, p2.Geometry)
            dc.Parameter._Value = GetParameter("b")._Value / 1
            v.Normalize()
            v.ScaleBy(-50)
            p3.MoveBy(v)
            r = ps.SketchLines.AddAsThreePointRectangle(p1.Geometry, l2d.EndSketchPoint.Geometry, p3.Geometry)

            pro = ps.Profiles.AddForSolid

            Return pro

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeStartingFace(pro As Profile) As FaceFeature
        Dim oSheetMetalCompDef As SheetMetalComponentDefinition
        oSheetMetalCompDef = doku.ComponentDefinition
        oSheetMetalCompDef.UseSheetMetalStyleThickness = False
        Dim oThicknessParam As Parameter
        oThicknessParam = oSheetMetalCompDef.Thickness
        oThicknessParam.Units = UnitsTypeEnum.kMillimeterLengthUnits
        oThicknessParam._Value = 0.03

        Dim oSheetMetalFeatures As SheetMetalFeatures
        oSheetMetalFeatures = doku.ComponentDefinition.Features

        compDef = oSheetMetalCompDef
        Dim oFaceFeatureDefinition As FaceFeatureDefinition
        oFaceFeatureDefinition = oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(pro)
        oFaceFeatureDefinition.Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection
        Dim oFaceFeature As FaceFeature
        oFaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition)
        feature = oFaceFeature
        Return oFaceFeature
    End Function
    Function FoldBand() As FoldFeature
        Try
            Dim oSheetMetalFeatures As SheetMetalFeatures
            oSheetMetalFeatures = doku.ComponentDefinition.Features
            Dim oFoldDefinition As FoldDefinition
            If bendAngle.Parameter._Value > Math.PI / 2 Then
                oFoldDefinition = oSheetMetalFeatures.FoldFeatures.CreateFoldDefinition(bendLine, bendAngle.Parameter._Value)
            Else

                oFoldDefinition = oSheetMetalFeatures.FoldFeatures.CreateFoldDefinition(bendLine, Math.PI - bendAngle.Parameter._Value)
            End If

            Dim oFoldFeature As FoldFeature
            oFoldFeature = oSheetMetalFeatures.FoldFeatures.Add(oFoldDefinition)
            folded = oFoldFeature
            Return folded

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetBendLine() As SketchLine
        Try
            Dim maxArea As Double = 0
            Dim oFoldLineSketch As PlanarSketch
            Dim maxface1, maxface2 As Face
            Dim pt1, pt2 As Point
            Dim d1, d2, a As Double
            oFoldLineSketch = compDef.Sketches.Add(mainWorkPlane)
            maxface1 = feature.Faces.Item(1)
            d1 = maxface1.GetClosestPointTo(mainSketch.firstLine.StartSketchPoint.Geometry).DistanceTo(mainSketch.firstLine.StartSketchPoint.Geometry)
            d2 = d1
            maxface2 = maxface1
            For Each f As Face In feature.Faces
                a = f.Evaluator.Area
                If f.Evaluator.Area > maxArea * 0.99 Then
                    maxArea = f.Evaluator.Area
                    maxface2 = maxface1
                    maxface1 = f
                End If
            Next
            d1 = maxface1.GetClosestPointTo(mainSketch.firstLine.StartSketchPoint.Geometry).DistanceTo(mainSketch.firstLine.StartSketchPoint.Geometry)
            d2 = maxface2.GetClosestPointTo(mainSketch.firstLine.StartSketchPoint.Geometry).DistanceTo(mainSketch.firstLine.StartSketchPoint.Geometry)

            If d2 < d1 Then
                maxface1 = maxface2
            End If
            Dim se1, se2 As SketchLine
            Dim e1, e2 As Edge
            Dim maxe1, maxe2 As Double
            maxe1 = 0
            maxe2 = 0
            e1 = maxface1.EdgeLoops.Item(1).Edges.Item(1)
            e2 = e1
            For Each el As EdgeLoop In maxface1.EdgeLoops
                For Each ed As Edge In el.Edges
                    If ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point) >= maxe1 Then
                        maxe2 = maxe1
                        e2 = e1
                        maxe1 = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
                        e1 = ed
                    End If
                Next

            Next
            se1 = oFoldLineSketch.AddByProjectingEntity(e1)
            se1.Construction = True
            se2 = oFoldLineSketch.AddByProjectingEntity(e2)
            se2.Construction = True
            edgeBand = e2
            Dim sl, bl As SketchLine
            sl = oFoldLineSketch.AddByProjectingEntity(mainSketch.thirdLine)
            sl.Construction = True
            bl = oFoldLineSketch.SketchLines.AddByTwoPoints(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry)
            oFoldLineSketch.GeometricConstraints.AddCollinear(bl, sl)
            oFoldLineSketch.GeometricConstraints.AddCoincident(bl.StartSketchPoint, se1)
            oFoldLineSketch.GeometricConstraints.AddCoincident(bl.EndSketchPoint, se2)
            bendLine = bl
            Return bl
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetWorkFace() As Face
        Try
            Dim maxArea1, maxArea2, maxArea3 As Double

            Dim maxface1, maxface2, maxface3 As Face


            maxface1 = compDef.Features.Item(compDef.Features.Count).Faces.Item(compDef.Features.Item(compDef.Features.Count).Faces.Count)
            maxArea2 = 0
            maxArea1 = maxArea2

            Dim minDis, d As Double

            minDis = 99999999

            For Each f As Face In compDef.Features.Item(compDef.Features.Count).Faces

                'lamp.HighLighFace(f)
                If f.Evaluator.Area > maxArea1 Then
                    maxArea2 = maxArea1
                    maxArea1 = f.Evaluator.Area
                    For Each v As Vertex In f.Vertices
                        d = v.Point.DistanceTo(mainSketch.thirdLine.EndSketchPoint.Geometry)
                        If d < minDis Then
                            minDis = d
                            maxface2 = maxface1
                            maxface1 = f

                        Else
                            maxface2 = f
                        End If
                    Next

                End If

            Next
            workface = maxface1
            ' lamp.HighLighFace(workFace)

            Return workface
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Public Function GetBendLine(workFace As Face, line As SketchLine3D) As SketchLine
        Try

            Dim ps As PlanarSketch
            Dim minDis As Double = 99999999999
            Dim d As Double
            ps = compDef.Sketches.Add(workFace)

            Dim sl As SketchLine
            sl = ps.AddByProjectingEntity(line)
            sl.Construction = True
            bendLine = ps.SketchLines.AddByTwoPoints(sl.EndSketchPoint.Geometry, sl.StartSketchPoint.Geometry)

            For Each ed As Edge In workFace.Edges
                d = ed.GetClosestPointTo(bendLine.StartSketchPoint.Geometry3d).DistanceTo(bendLine.StartSketchPoint.Geometry3d)
                If d < minDis Then
                    minDis = d
                    edgeBand = ed
                End If
            Next
            'lamp.HighLighObject(edgeBand)
            Return sl
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
    Function GetFoldingAngle() As DimensionConstraint
        Dim ps As PlanarSketch

        Dim p1, p2, p3 As SketchPoint
        Dim v As Vector2d
        Dim r As SketchEntitiesEnumerator
        Dim sl, el As SketchLine
        Dim wp As WorkPlane
        Try

            wp = compDef.WorkPlanes.AddByNormalToCurve(bendLine, bendLine.StartSketchPoint)

            ps = doku.ComponentDefinition.Sketches.Add(wp)
            Dim fl As SketchLine3D
            fl = mainSketch.bandLines.Item(4)
            sl = ps.AddByProjectingEntity(fl)
            el = ps.AddByProjectingEntity(GetMinorEdge())
            Dim dc As DimensionConstraint
            dc = ps.DimensionConstraints.AddTwoLineAngle(el, sl, sl.EndSketchPoint.Geometry)
            dc.Driven = True
            bendAngle = dc


            Return dc

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
    Function GetMinorEdge() As Edge
        GetMajorEdge(workface)

        Return minorEdge
    End Function
    Public Function FoldBand(i As Integer) As FoldFeature
        Try



            Dim oFoldFeature As FoldFeature
            oFoldFeature = features.FoldFeatures.Add(AdjustFoldDefinition(i))
            folded = oFoldFeature

            Return folded

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function AdjustFoldDefinition(i As Integer) As FoldDefinition
        Try
            features = doku.ComponentDefinition.Features
            Dim oFoldDefinition As FoldDefinition
            If i > 4 Then

                If bendAngle.Parameter._Value > Math.PI / 2 Then
                    oFoldDefinition = features.FoldFeatures.CreateFoldDefinition(bendLine, bendAngle.Parameter._Value)
                Else
                    oFoldDefinition = features.FoldFeatures.CreateFoldDefinition(bendLine, Math.PI - bendAngle.Parameter._Value)
                End If

                If Not oFoldDefinition.IsPositiveBendSide Then
                    oFoldDefinition.IsPositiveBendSide = Not oFoldDefinition.IsPositiveBendSide

                End If
            Else
                If bendAngle.Parameter._Value > Math.PI / 2 Then
                    oFoldDefinition = features.FoldFeatures.CreateFoldDefinition(bendLine, Math.PI - bendAngle.Parameter._Value)
                Else

                    oFoldDefinition = features.FoldFeatures.CreateFoldDefinition(bendLine, bendAngle.Parameter._Value)
                End If
                If oFoldDefinition.IsPositiveBendSide Then
                    oFoldDefinition.IsPositiveBendSide = Not oFoldDefinition.IsPositiveBendSide
                End If
            End If

            Return oFoldDefinition
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
End Class
