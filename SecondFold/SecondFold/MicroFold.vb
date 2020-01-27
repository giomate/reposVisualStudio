Imports Inventor
Imports DrawingMainSketch
Imports GetInitialConditions
Imports DrawInitialSketch
Public Class MicroFold
    Dim doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Dim refLine, firstLine, secondLine, thirdLine, lastLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring

    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3 As Point
    Dim tg As TransientGeometry
    Dim gap1, thickness As Double
    Dim partNumber As Integer
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim mainSketch As Sketcher3D

    Dim pro As Profile
    Dim feature As FaceFeature
    Dim bendLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim minorEdge, majorEdge, bendEdge As Edge
    Dim workFace As Face
    Dim bendAngle As DimensionConstraint
    Dim folded As FoldFeature
    Dim features As SheetMetalFeatures
    Dim lamp As Highlithing

    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(app)

        mainSketch = New Sketcher3D(doku)
        compDef = doku.ComponentDefinition
        features = compDef.Features
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        thickness = compDef.Thickness._Value


        done = False
    End Sub
    Function GetQNumber(docu As Inventor.Document) As Integer
        Dim s() As String
        s = Strings.Split(docu.FullFileName, "Band")
        s = Strings.Split(s(1), ".ipt")

        Return CInt(s(0)) + 1
    End Function
    Function GetSketchName(docu As Inventor.Document) As String
        Dim s(), sn As String
        sn = compDef.Sketches3D.Item(compDef.Sketches3D.Count).Name
        s = Strings.Split(sn, "s")
        sn = String.Concat("s" & (CInt(s(1)) + 1).ToString)

        Return sn
    End Function

    Public Function MakeSecondFold() As Boolean
        Try
            If GetWorkFace().SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                If mainSketch.DrawTrobinaCurve(GetQNumber(doku), GetSketchName(doku)).Construction Then
                    sk3D = mainSketch.sk3D
                    If GetMinorEdge(workFace).GeometryType = CurveTypeEnum.kLineSegmentCurve Then
                        If DrawFirstLine().Length > 0 Then

                        End If

                    End If

                End If

                Return True
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try


        Return False
    End Function
    Function GetWorkFace() As Face
        Try
            Dim maxArea1, maxArea2 As Double

            Dim maxface1, maxface2 As Face

            maxface1 = compDef.Features.Item(compDef.Features.Count).Faces.Item(compDef.Features.Item(compDef.Features.Count).Faces.Count)
            maxface2 = maxface1
            maxArea2 = maxface2.Evaluator.Area
            maxArea1 = maxArea2
            For Each f As Face In compDef.Features.Item(compDef.Features.Count).Faces

                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then

                    If f.Evaluator.Area > maxArea1 Then
                        maxArea2 = maxArea1
                        maxface2 = maxface1
                        maxArea1 = f.Evaluator.Area
                        maxface1 = f

                    End If
                End If

            Next
            workFace = maxface2

            'HighLighFace(maxface2)
            Return maxface2
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function GetMinorEdge(f As Face) As Edge
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
        'lamp.HighLighObject(e2)
        'lamp.HighLighObject(e1)
        bendEdge = e3
        minorEdge = e2
        majorEdge = e1
        Return e2
    End Function
    Function DrawFirstLine() As SketchLine3D
        Try

            Dim l, mn, mj As SketchLine3D
            mn = sk3D.Include(minorEdge)
            mj = sk3D.Include(majorEdge)
            l = sk3D.SketchLines3D.AddByTwoPoints(GetStartPoint(), majorEdge.GetClosestPointTo(GetStartPoint()))

            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, mn)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, mj)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, mn.EndPoint)
            dc.Parameter._Value = mn.Length - thickness
            doku.Update2(True)
            bandLines.Add(l)
                firstLine = l
                lastLine = l




            Return l
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function
    Function GetStartPoint() As Point
        Dim pt As Point
        Dim v As Vector

        If minorEdge.StartVertex.Point.DistanceTo(bendEdge.StartVertex.Point) < minorEdge.StopVertex.Point.DistanceTo(bendEdge.StartVertex.Point) Then
            pt = minorEdge.StartVertex.Point
            v = minorEdge.StartVertex.Point.VectorTo(minorEdge.StopVertex.Point)
        Else
            pt = minorEdge.StopVertex.Point
            v = minorEdge.StopVertex.Point.VectorTo(minorEdge.StartVertex.Point)

        End If
        'lamp.HighLighObject(pt)

        v.AsUnitVector.AsVector()
        v.ScaleBy(thickness / 10)
        pt.TranslateBy(v)
        'lamp.HighLighObject(pt)
        point1 = pt
        Return pt
    End Function



End Class
