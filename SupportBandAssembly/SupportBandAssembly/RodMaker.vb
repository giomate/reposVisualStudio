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

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, connectLine, centroLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile

    Public trobinaCurve As Curves3D

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Public partNumber, qNext, qLastTie As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Dim nextSketch As OriginSketch
    Dim cutProfile, faceProfile, rodProfile As Profile

    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Public compDef As ComponentDefinition
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim mainWorkPlane As WorkPlane
    Dim workAxis As WorkAxis
    Dim faceRod As Face
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, twistFace As Face
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

    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails, caras, surfaceBodies As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim refDoc As FindReferenceLine

    Dim manager As FoldingEvaluator
    Dim giro As TwistFold7
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
            ' lamp.HighLighFace(f)
            wa = doku.ComponentDefinition.WorkAxes.AddByRevolvedFace(f)
            workAxis = wa

            le = DrawRodSketch(f, wa)
            rodProfile = DrawPlannSketch(centroLine, le)
            Return RevolveFace(rodProfile)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawRodSketch(f As Face, a As WorkAxis) As SketchLine3D
        Dim l, le As SketchLine3D
        Dim gc As GeometricConstraint3D
        sk3D = doku.ComponentDefinition.Sketches3D.Add()

        le = sk3D.Include(GetMajorEdge(f))
        l = sk3D.SketchLines3D.AddByTwoPoints(le.StartSketchPoint.Geometry, le.EndSketchPoint.Geometry, False)
        gc = sk3D.GeometricConstraints3D.AddCollinear(l, a)
        centroLine = l
        Return le
    End Function
    Function DrawPlannSketch(cl As SketchLine3D, le As SketchLine3D) As Profile
        Try
            Dim wpl As WorkPlane
            Dim ps As PlanarSketch
            Dim p1, p2, p3 As SketchPoint
            Dim le2d, l2d, l32d, mne As SketchLine
            Dim gc, gccc As GeometricConstraint
            Dim r As SketchEntitiesEnumerator
            Dim dc As DimensionConstraint
            Dim d As Double = 0
            Dim i As Integer
            wpl = doku.ComponentDefinition.WorkPlanes.AddByThreePoints(cl.EndPoint, cl.StartPoint, le.StartPoint)
            ps = doku.ComponentDefinition.Sketches.Add(wpl)

            le2d = ps.AddByProjectingEntity(le)
            le2d.Construction = True
            l2d = ps.AddByProjectingEntity(cl)
            l2d.Construction = True
            mne = ps.AddByProjectingEntity(minorEdge)
            mne.Construction = True


            If minorEdge.GetClosestPointTo(majorEdge.StartVertex.Point).DistanceTo(majorEdge.StartVertex.Point) > minorEdge.GetClosestPointTo(majorEdge.StopVertex.Point).DistanceTo(majorEdge.StopVertex.Point) Then
                l32d = ps.SketchLines.AddByTwoPoints(le2d.StartSketchPoint.Geometry, l2d.StartSketchPoint.Geometry)
                l32d.Construction = True
                gccc = ps.GeometricConstraints.AddCoincident(l32d.StartSketchPoint, le2d)

            Else
                l32d = ps.SketchLines.AddByTwoPoints(le2d.EndSketchPoint.Geometry, l2d.EndSketchPoint.Geometry)
                l32d.Construction = True
                gccc = ps.GeometricConstraints.AddCoincident(l32d.StartSketchPoint, le2d)

            End If

            gc = ps.GeometricConstraints.AddPerpendicular(l32d, le2d)
            If mne.EndSketchPoint.Geometry.DistanceTo(l32d.EndSketchPoint.Geometry) > mne.StartSketchPoint.Geometry.DistanceTo(l32d.EndSketchPoint.Geometry) Then
                ps.SketchLines.AddAsThreePointRectangle(l32d.StartSketchPoint.Geometry, l32d.EndSketchPoint.Geometry, mne.EndSketchPoint.Geometry)
                ' dc = ps.DimensionConstraints.AddTwoPointDistance(l32d.EndSketchPoint, mne.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, mne.EndSketchPoint.Geometry)
            Else
                ps.SketchLines.AddAsThreePointRectangle(l32d.StartSketchPoint.Geometry, l32d.EndSketchPoint.Geometry, mne.StartSketchPoint.Geometry)
                ' dc = ps.DimensionConstraints.AddTwoPointDistance(l32d.EndSketchPoint, mne.StartSketchPoint, DimensionOrientationEnum.kAlignedDim, mne.StartSketchPoint.Geometry)
            End If
            i = ps.SketchLines.Count
            If ps.SketchLines.Item(i).Length > ps.SketchLines.Item(i - 1).Length Then
                largo = ps.DimensionConstraints.AddTwoPointDistance(ps.SketchLines.Item(i).StartSketchPoint, ps.SketchLines.Item(i).EndSketchPoint, DimensionOrientationEnum.kAlignedDim, mne.StartSketchPoint.Geometry)
            Else
                largo = ps.DimensionConstraints.AddTwoPointDistance(ps.SketchLines.Item(i - 1).StartSketchPoint, ps.SketchLines.Item(i - 1).EndSketchPoint, DimensionOrientationEnum.kAlignedDim, mne.StartSketchPoint.Geometry)
            End If


            faceProfile = ps.Profiles.AddForSolid
            workAxis.Visible = False
            wpl.Visible = False
            Return faceProfile
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

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
                            rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, workAxis, PartFeatureOperationEnum.kJoinOperation)
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

        Return rf
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
End Class
