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

    Public trobinaCurve As Curves3D

    Public wp1, wp2, wp3, wpConverge As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim skpt1, skpt2, skpt3 As SketchPoint3D
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Public partNumber, qNext, qLastTie As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Dim nextSketch As OriginSketch
    Dim cutProfile, faceProfile, rodProfile As Profile
    Dim circle As SketchCircle3D

    Dim pro As Profile
    Dim direction As Vector


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

    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails, caras, inputPoints As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim refDoc As FindReferenceLine


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

        inputPoints = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)

        gap1CM = 3 / 10

        done = False
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
    Function GetHighestVertex() As Vertex
        Dim dMax As Double = 0
        Dim sb As SurfaceBody = doku.ComponentDefinition.SurfaceBodies.Item(1)
        Dim pMax1, pMax2 As Point

        Dim v As Vector
        Dim dMin As Double = 999999
        Dim cpt As Point = compDef.WorkPoints.Item(1).Point
        Dim e As Double = 0
        Try
            For Each f As Face In sb.Faces
                For Each ver As Vertex In f.Vertices
                    e = ver.Point.Z / ver.Point.DistanceTo(cpt)
                    If (e > dMax) Then
                        dMax = e
                        verMax2 = verMax1
                        verMax1 = ver
                        pMax2 = pMax1
                        pMax1 = ver.Point
                    Else
                        verMax2 = ver
                        pMax2 = ver.Point
                    End If
                Next
            Next

            skpt1 = sk3D.SketchPoints3D.Add(pMax1)
            Return verMax1
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
        Dim ver As Vertex = GetHighestVertex()
        Dim ed As Edge = GetInputEdge(ver)
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
            inputPoints.Clear()

            For i = 1 To 6
                pt.TranslateBy(ve)
                skpt = sk3D.SketchPoints3D.Add(pt)
                inputPoints.Add(skpt)
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
    Function MakeAllWiresGuide() As ExtrudeFeature
        Dim ef As ExtrudeFeature
        sk3D = compDef.Sketches3D.Add()
        GetSixPoints()
        Dim total As Integer = inputPoints.Count
        Dim skpt As SketchPoint3D
        Try
            For i = 1 To total
                skpt = inputPoints.Item(i)
                ef = RemoveSingleWire(DrawInletLine(skpt))
            Next
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
            ed.SetThroughAllExtent(PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            ef = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(ed)

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawRemoveCircle(skl As SketchLine3D) As Profile3D
        Dim pro As Profile3D
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
End Class
