﻿Imports Inventor

Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory
Imports Subina_Design_Helpers

Public Class Stanzer
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, connectLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim adjuster As SketchAdjust
    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile

    Public trobinaCurve As Curves3D


    Public wp1, wp2, wp3, wpConverge As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, convergePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Public partNumber, qNext, qLastTie As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Private converter As Conversions

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
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, cutEdge1, cutEdge2, outerEdge, closestEdge As Edge
    Dim minorLine, majorLine, cutLine3D, kante3D, tante3D As SketchLine3D
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
    Dim bandFaces As WorkSurface

    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails, caras, surfaceBodies As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane

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
        converter = New Conversions(app)

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
        adjuster = New SketchAdjust(doku)
        gap1CM = 3 / 10

        nombrador = New Nombres(doku)

        trobinaCurve = New Curves3D(doku)

        DP.Dmax = 200 / 10
        DP.Dmin = 1 / 10
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 11
        DP.q = 23
        DP.b = 25

        done = False
    End Sub
    Function DocUpdate(docu As PartDocument) As PartDocument
        doku = docu

        compDef = doku.ComponentDefinition
        lamp = New Highlithing(doku)
        adjuster = New SketchAdjust(doku)
        Return doku
    End Function

    Function IsWedgeDone(s As String) As Boolean
        For Each ffn As String In fullFileNames
            If ffn.Equals(s) Then
                Return True
            End If
        Next

        Return False
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

    Function GetRealNormal(f As Face, ws As WorkSurface) As Vector
        Dim v As Vector
        Dim wpf, wpt, wptpf As WorkPoint
        Dim pt As Point
        Dim pl, plw As Plane
        Try
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
                Try
                    wpf.Delete()
                    wptpf.Delete()
                Catch ex2 As Exception

                End Try
                Return v
            Catch ex As Exception
                Try
                    wpf.Delete()
                    wptpf.Delete()
                Catch ex2 As Exception
                    MsgBox(ex.ToString())
                    Return Nothing
                End Try
            End Try
        Catch ex As Exception
            Try
                wpf.Delete()
                wptpf.Delete()
            Catch ex2 As Exception
                MsgBox(ex.ToString())
                Return Nothing
            End Try
        End Try
        Return Nothing
    End Function

    Function CalculateEntryDistance(ed As Edge) As Double

        Dim ve, vc As Vector
        Dim mpt As Point
        Dim ls As LineSegment
        Dim e As Double
        ls = ed.Geometry
        mpt = ls.MidPoint
        ve = ls.Direction.AsVector
        vc = doku.ComponentDefinition.WorkPoints.Item(1).Point.VectorTo(mpt)
        e = (Math.Abs(vc.DotProduct(ve))) * vc.Length

        Return e
    End Function
    Function CalculateClosestFace(f As Face) As Double

        Dim vp, vc As Vector
        Dim mpt As Point
        Dim pl As Plane
        Dim wptfc As WorkPoint

        Dim e As Double
        pl = f.Geometry
        wptfc = doku.ComponentDefinition.WorkPoints.AddAtCentroid(f.EdgeLoops.Item(1))
        wptfc.Visible = False
        mpt = wptfc.Point
        vp = pl.Normal.AsVector
        vc = doku.ComponentDefinition.WorkPoints.Item(1).Point.VectorTo(mpt)
        e = Math.Abs(vc.DotProduct(vp)) / (vc.Length)

        Return e
    End Function
    Function RemovePatch(spl As PlanarSketch, skl As SketchLine, wpt As WorkPoint) As ExtrudeFeature
        Try

            Dim v As Vector2d = skl.Geometry.Direction.AsVector
            Dim spt As SketchPoint = spl.SketchPoints.Add(skl.Geometry.MidPoint)
            Dim m As Matrix2d = tg.CreateMatrix2d
            m.SetToRotation(Math.PI / 2, spt.Geometry)
            v.TransformBy(m)
            v.ScaleBy(skl.Length)
            Dim pt2d As Point2d = skl.StartSketchPoint.Geometry
            pt2d.TranslateBy(v)
            spl.SketchPoints.Add(pt2d)
            spl.SketchLines.AddAsThreePointCenteredRectangle(spt, skl.StartSketchPoint, pt2d)

            Dim pro As Profile

            pro = spl.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
            oExtrudeDef.SetToExtent(wpt)
            Dim oExtrude As ExtrudeFeature
            oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)
            Return oExtrude
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

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
    Function GetMajorFace(ef As ExtrudeFeature) As Face

        Try
            Dim maxArea1, maxArea2, maxArea3 As Double

            Dim maxface1, maxface2, maxface3 As Face


            maxface1 = ef.SideFaces.Item(1)
            maxface2 = maxface1
            maxArea2 = 0
            maxArea1 = maxArea2
            maxArea3 = maxArea2





            For Each f As Face In ef.SideFaces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
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

                End If
            Next
            sideMinorFace = maxface2
            sideMajorFace = maxface1

            lamp.HighLighFace(sideMajorFace)
            lamp.HighLighFace(sideMinorFace)
            Return sideMajorFace
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
    Public Function EmbossLetter(q As Integer, spt As SketchPoint3D, f As Face) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim s As String
        Try
            If spt.Geometry.Z > 0 Then
                s = nombrador.ConvertQNumberLetter(q + 1)
            Else
                s = nombrador.ConvertQNumberLetter(q)

            End If
            lamp.HighLighFace(f)
            ef = ExtrudeLetter(SketchLetter(s, f, spt))

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return ef
    End Function
    Public Function ExtrudeFrameLetter(name As String, wpt As WorkPoint, skl As SketchLine3D, fi As Face) As ExtrudeFeature
        Dim ef As ExtrudeFeature
        Dim s As String
        Dim q As Integer
        Try
            q = nombrador.GetQNumberString(name)
            s = nombrador.ConvertQNumberLetter(q + 64 - 33)

            ' lamp.HighLighObject(ed)
            ef = ExtrudeLetter(SketchFrameLetter(s, skl, wpt, fi))

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
    Function SketchFrameLetter(s As String, skli As SketchLine3D, wpt As WorkPoint, fi As Face) As Profile
        Dim oProfile As Profile
        Dim d, w, h As Double

        Try

            Dim spt, sptw As SketchPoint
            Dim skl As SketchLine
            Dim ps As PlanarSketch

            Dim v, vy, ve As Vector

            ' Dim wpl As WorkPlane = compDef.WorkPlanes.AddByThreePoints(skli.StartPoint, skli.EndPoint, sklp.EndPoint)
            ps = doku.ComponentDefinition.Sketches.Add(fi)
            lamp.LookAtFace(fi)


            skl = ps.AddByProjectingEntity(skli)
            skl.Construction = True
            sk3D = compDef.Sketches3D.Add
            Dim sklRef As SketchLine3D = sk3D.SketchLines3D.AddByTwoPoints(skl.StartSketchPoint.Geometry3d, skl.EndSketchPoint.Geometry3d, False)
            sk3D.GeometricConstraints3D.AddGround(sklRef)
            Dim sklp As SketchLine3D = CorrectFrameLetterPosition(sklRef, wpt, fi)
            ps.Visible = False
            Dim wpl As WorkPlane = compDef.WorkPlanes.AddByThreePoints(sklRef.StartPoint, sklRef.EndPoint, sklp.EndPoint)
            ps = doku.ComponentDefinition.Sketches.Add(wpl)
            spt = ps.AddByProjectingEntity(sklp.StartPoint)
            Dim skpt As SketchPoint3D = sk3D.SketchPoints3D.Add(sklp.EndSketchPoint.Geometry)
            sptw = ps.AddByProjectingEntity(skpt)

            Dim oTextBox As TextBox
            Dim oStyle As TextStyle
            Dim sText As String
            sText = s

            oTextBox = ps.TextBoxes.AddFitted(spt.Geometry, sText)
            oStyle = ps.TextBoxes.Item(1).Style
            oTextBox.Delete()
            oStyle.FontSize = sklp.Length
            oStyle.Bold = True
            doku.Update2(True)
            Dim ptBox As Point2d
            ptBox = tg.CreatePoint2d(spt.Geometry.X - oStyle.FontSize / 4, -0.5 / 10)
            oTextBox = ps.TextBoxes.AddFitted(ptBox, sText, oStyle)
            w = oTextBox.Width
            h = oTextBox.Height

            oTextBox.Origin = tg.CreatePoint2d(spt.Geometry.X - w / 2, 0)
            doku.Update2(True)
            Dim reflex As TextBox
            ptBox = tg.CreatePoint2d(spt.Geometry.X - w / 2, h)
            reflex = ps.TextBoxes.AddFitted(ptBox, sText, oStyle)



            ' oTextBox.Rotation = Math.PI



            Dim oPaths As ObjectCollection
            oPaths = app.TransientObjects.CreateObjectCollection

            If oTextBox.Origin.DistanceTo(sptw.Geometry) < reflex.Origin.DistanceTo(sptw.Geometry) Then
                oPaths.Add(oTextBox)
            Else
                oPaths.Add(reflex)
            End If




            oProfile = ps.Profiles.AddForSolid(False, oPaths)
            wpl.Visible = False
            ps.Visible = False
            sk3D.Visible = False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return oProfile
    End Function
    Function CorrectFrameLetterPosition(skli As SketchLine3D, wpt As WorkPoint, fi As Face) As SketchLine3D
        Try
            Dim seq As SketchEquationCurve3D
            Dim vz As Vector = tg.CreateVector(0, 0, 1)
            Dim curvesSketch As Sketch3D = compDef.Sketches3D.Item("curvas")
            Dim sklRing, sklPlane, sklo, sklc, sklr, sklz, sklt, sklt2 As SketchLine3D
            Dim ac, ac2, dcm As DimensionConstraint3D
            sklPlane = sk3D.SketchLines3D.AddByTwoPoints(skli.Geometry.MidPoint, wpt.Point, False)
            Dim gc As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddOnFace(sklPlane, fi)
            Dim sb As SurfaceBody = compDef.SurfaceBodies(1)
            Dim pt As Point
            Dim dpt(2) As Double
            sk3D.GeometricConstraints3D.AddCoincident(sklPlane.StartPoint, skli)
            sklPlane.Construction = True
            comando.WireFrameView(doku)

            If wpt.Point.Z > 0 Then
                seq = curvesSketch.SketchEquationCurves3D.Item(2)

            Else
                seq = curvesSketch.SketchEquationCurves3D.Item(3)

            End If

            sklRing = sk3D.SketchLines3D.AddByTwoPoints(sklPlane.EndPoint, wpt.Point, False)

            sk3D.GeometricConstraints3D.AddCoincident(sklRing.EndPoint, seq)
            sklRing.Construction = True
            sklz = sk3D.SketchLines3D.AddByTwoPoints(sklRing.EndPoint, sklPlane.EndSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddParallelToZAxis(sklz)
            sklc = sk3D.SketchLines3D.AddByTwoPoints(compDef.WorkPoints(1), sklRing.EndPoint, False)
            sklr = sk3D.SketchLines3D.AddByTwoPoints(sklRing.EndPoint, wpt.Point, False)
            TryPerpendicular(sklr, sklc)
            TryPerpendicular(sklr, sklz)
            TryPerpendicular(sklRing, sklr)

            TryPerpendicular(sklPlane, skli)
            sklt = sk3D.SketchLines3D.AddByTwoPoints(sklPlane.EndPoint, wpt.Point, False)
            sk3D.GeometricConstraints3D.AddOnFace(sklt, fi)
            TryPerpendicular(sklPlane, sklt)
            dcm = sk3D.DimensionConstraints3D.AddLineLength(sklt)
            adjuster.AdjustDimConstrain3DSmothly(dcm, sklPlane.Length / 6)
            sklt2 = sk3D.SketchLines3D.AddByTwoPoints(sklPlane.EndPoint, wpt.Point, False)
            ' sk3D.GeometricConstraints3D.AddOnFace(sklt2, fi)
            TryOpposite(sklt2, sklt)
            sk3D.GeometricConstraints3D.AddEqual(sklt, sklt2)

            ForceLetterMiddle(sklPlane, skli)
            Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(sklPlane)
            adjuster.AdjustDimConstrain3DSmothly(dc, 10 / 10)
            dc.Driven = True
            dc = sk3D.DimensionConstraints3D.AddLineLength(sklRing)
            For i = 1 To 32
                adjuster.AdjustDimConstrain3DSmothly(dc, dc.Parameter._Value * 15 / 16)
                dc.Driven = True
                adjuster.AdjustDimConstrain3DSmothly(dcm, sklPlane.Length / 4)
                dcm.Driven = True
                pt = sklt.EndSketchPoint.Geometry
                dpt(0) = pt.X
                dpt(1) = pt.Y
                dpt(2) = pt.Z
                If (sb.IsPointInside(dpt, False) = ContainmentEnum.kInsideContainment) Or (sklRing.Length < 12 / 10) Then
                    Exit For
                Else
                    pt = sklt2.EndSketchPoint.Geometry
                    dpt(0) = pt.X
                    dpt(1) = pt.Y
                    dpt(2) = pt.Z
                    If (sb.IsPointInside(dpt, False) = ContainmentEnum.kInsideContainment) Then
                        Exit For
                    Else
                        pt = sklPlane.EndSketchPoint.Geometry
                        dpt(0) = pt.X
                        dpt(1) = pt.Y
                        dpt(2) = pt.Z
                        If (sb.IsPointInside(dpt, False) = ContainmentEnum.kInsideContainment) Then
                            Exit For
                        End If
                    End If

                End If

            Next


            Return sklPlane
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function ForceLetterMiddle(lp As SketchLine3D, lr As SketchLine3D) As Boolean
        Dim dc As DimensionConstraint3D
        Dim skpt As SketchPoint3D
        If Math.Abs(lr.StartSketchPoint.Geometry.Z) < Math.Abs(lr.EndSketchPoint.Geometry.Z) Then
            skpt = sk3D.SketchPoints3D.Add(lr.StartSketchPoint.Geometry)
        Else
            skpt = sk3D.SketchPoints3D.Add(lr.EndSketchPoint.Geometry)
        End If
        sk3D.GeometricConstraints3D.AddGround(skpt)
        dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(lp.StartPoint, skpt)
        ForceLetterMiddle = adjuster.AdjustDimConstrain3DSmothly(dc, lr.Length / 4)
        dc.Delete()
        Return ForceLetterMiddle
    End Function
    Function TryPerpendicular(l1 As SketchLine3D, l2 As SketchLine3D) As GeometricConstraint3D
        Try
            Dim ac As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(l1, l2)
            Dim gc As GeometricConstraint3D
            If adjuster.AdjustDimConstrain3DSmothly(ac, Math.PI / 2) Then
                ac.Driven = True
                Try
                    gc = sk3D.GeometricConstraints3D.AddPerpendicular(l1, l2)
                    ac.Delete()
                Catch ex2 As Exception
                    If adjuster.AdjustDimConstrain3DSmothly(ac, Math.PI / 2) Then
                        ac.Driven = True
                        Try
                            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l1, l2)
                        Catch ex3 As Exception
                            ac.Driven = False
                        End Try

                    Else
                        ac.Driven = False

                    End If
                End Try


            Else
                ac.Driven = False
            End If
            Return gc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function TryOpposite(l1 As SketchLine3D, l2 As SketchLine3D) As GeometricConstraint3D
        Dim ac As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoLineAngle(l1, l2)
        Dim gc As GeometricConstraint3D
        If adjuster.AdjustDimConstrain3DSmothly(ac, Math.PI) Then
            ac.Delete()
            gc = sk3D.GeometricConstraints3D.AddCollinear(l1, l2)
        Else
            ac.Driven = False
        End If
        Return gc
    End Function
    Function SketchLetter(s As String, f As Face, spti As SketchPoint3D) As Profile
        Dim oProfile As Profile


        Dim a, b, d, w, h As Double
        Dim natural As Boolean
        Dim pt2d As Point2d
        Dim v2d As Vector2d
        Try

            Dim spt As SketchPoint
            Dim skl As SketchLine
            Dim oSketch As PlanarSketch
            Dim v, vy, ve As Vector
            Dim pl As Plane = f.Geometry

            Dim l As Line
            Dim edMax As Edge
            edMax = GetOuterEdge(f, spti.Geometry)
            ve = edMax.StartVertex.Point.VectorTo(edMax.StopVertex.Point)
            vy = ve.CrossProduct(pl.Normal.AsVector)
            v = spti.Geometry.VectorTo(compDef.WorkPoints.Item(1).Point)
            d = v.DotProduct(vy)
            If d < 0 Then
                natural = False
            Else
                natural = True
            End If
            oSketch = doku.ComponentDefinition.Sketches.AddWithOrientation(f, edMax, True, True, edMax.StartVertex,)

            spt = oSketch.AddByProjectingEntity(spti)

            skl = oSketch.AddByProjectingEntity(edMax)
            skl.Construction = True
            Dim oTextBox As TextBox
            Dim oStyle As TextStyle
            Dim sText As String
            sText = s

            oTextBox = oSketch.TextBoxes.AddFitted(spt.Geometry, sText)
            oStyle = oSketch.TextBoxes.Item(1).Style
            oTextBox.Delete()
            oStyle.FontSize = 0.2
            oStyle.Bold = True
            doku.Update2(True)
            Dim ptBox As Point2d
            ptBox = tg.CreatePoint2d(spt.Geometry.X - oStyle.FontSize / 4, -0.5 / 10)
            oTextBox = oSketch.TextBoxes.AddFitted(ptBox, sText, oStyle)
            w = oTextBox.Width
            h = oTextBox.Height

            oTextBox.Origin = tg.CreatePoint2d(spt.Geometry.X - w / 2, -0.5 / 10)
            doku.Update2(True)
            Dim reflex As TextBox
            ptBox = tg.CreatePoint2d(spt.Geometry.X - w / 2, 0.5 / 10 + h)
            reflex = oSketch.TextBoxes.AddFitted(ptBox, sText, oStyle)



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
    Function GetClosestVertexFace(f As Face, spt As SketchPoint3D) As Vertex
        Dim cv As Vertex
        Dim d, dMin As Double
        dMin = 999999
        Try
            For Each ver As Vertex In f.Vertices
                d = spt.Geometry.DistanceTo(ver.Point)
                If d < dMin Then
                    dMin = d
                    cv = ver
                End If
            Next
            lamp.HighLighObject(cv)
            Return cv
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetClosestVertexEdge(ed As Edge, spt As SketchPoint3D) As Vertex
        Dim cv As Vertex
        Dim d, e, dMin As Double
        dMin = 999999
        Try

            d = spt.Geometry.DistanceTo(ed.StartVertex.Point)
            e = spt.Geometry.DistanceTo(ed.StopVertex.Point)
            If d < e Then
                Return ed.StartVertex
            Else
                Return ed.StopVertex
            End If

            lamp.HighLighObject(cv)
            Return cv
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
        e1 = f.EdgeLoops.Item(1).Edges.Item(1)
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
        lamp.HighLighObject(e3)
        lamp.HighLighObject(e2)
        lamp.HighLighObject(e1)
        bendEdge = e3
        minorEdge = e2
        majorEdge = e1
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
    Function GetClosestEdge(f As Face) As Edge
        Dim e1, e2, e3 As Edge
        Dim mine1, mine2, mine3, d, e As Double
        Dim ve, vc As Vector
        Dim pt1, pt2, pt3 As Point
        Dim ls As LineSegment
        mine1 = 99999
        mine2 = 9999999
        mine3 = 999999
        e1 = f.EdgeLoops.Item(1).Edges.Item(1)
        e2 = e1
        e3 = e2
        pt1 = f.GetClosestPointTo(wpConverge.Point)
        e = pt1.DistanceTo(wpConverge.Point)
        ' sk3D.SketchPoints3D.Add(pt1)
        For Each ed As Edge In f.EdgeLoops.Item(1).Edges
            d = ed.GetClosestPointTo(wpConverge.Point).DistanceTo(wpConverge.Point)
            If Math.Abs(d - e) < 1 / 10 Then
                ve = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point)
                ls = tg.CreateLineSegment(ed.StartVertex.Point, ed.StopVertex.Point)
                vc = ls.MidPoint.VectorTo(wpConverge.Point)
                d = ls.MidPoint.DistanceTo(wpConverge.Point) / (ve.CrossProduct(vc).Length)
                If d < mine2 Then
                    If d < mine1 Then
                        mine3 = mine2
                        e3 = e2
                        mine2 = mine1
                        e2 = e1
                        mine1 = d
                        e1 = ed
                    Else
                        mine3 = mine2
                        e3 = e2
                        mine2 = d
                        e2 = ed
                    End If
                Else
                    mine3 = d
                    e3 = ed
                End If

            End If


        Next
        lamp.HighLighObject(e1)

        closestEdge = e1
        Return e1
    End Function
    Function GetOuterEdge(f As Face, pti As Point) As Edge
        Dim e1, e2, e3 As Edge
        Dim mine1, mine2, mine3, d, e, dis, dMin As Double
        Dim ve, vc As Vector
        Dim pt1, pt2, pt3 As Point
        Dim ls As LineSegment
        mine1 = 99999
        mine2 = 9999999
        mine3 = 999999
        e1 = GetMajorEdge(f)
        e2 = e1
        e3 = e2
        pt1 = f.GetClosestPointTo(pti)
        e = pt1.DistanceTo(pti)
        dMin = e1.StartVertex.Point.DistanceTo(e1.StopVertex.Point)
        ' sk3D.SketchPoints3D.Add(pt1)
        For Each ed As Edge In f.Edges
            dis = ed.StartVertex.Point.DistanceTo(ed.StopVertex.Point)
            If dis > dMin / 2 Then

                d = ed.GetClosestPointTo(pti).DistanceTo(pti)
                If Math.Abs(d - e) < 1 / 10 Then

                    If d < mine2 Then
                        If d < mine1 Then
                            mine3 = mine2
                            e3 = e2
                            mine2 = mine1
                            e2 = e1
                            mine1 = d
                            e1 = ed
                        Else
                            mine3 = mine2
                            e3 = e2
                            mine2 = d
                            e2 = ed
                        End If
                    Else
                        mine3 = d
                        e3 = ed
                    End If

                End If
            End If



        Next
        lamp.HighLighObject(e1)

        outerEdge = e1
        Return e1
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
    Function ExtrudeLetter(pro As Profile) As ExtrudeFeature
        Dim oExtrudeDef As ExtrudeDefinition
        Try
            oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kJoinOperation)
            oExtrudeDef.SetDistanceExtent(3 / 10, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            'oExtrudeDef.SetDistanceExtent(0.12, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)
            Return oExtrude
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try





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
    Function RemoveFaceExtend(f As Face) As ExtrudeFeature
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
        oExtrudeDef.SetDistanceExtent(25 / 10, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude As ExtrudeFeature
        oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)

        Return oExtrude
    End Function
    Function RemoveEntryMaterial(spl As PlanarSketch, ed As Edge, d As Double) As ExtrudeFeature
        Try
            Dim skl As SketchLine
            lamp.HighLighObject(ed)
            skl = spl.AddByProjectingEntity(ed)
            skl.Construction = True
            Dim v As Vector2d = skl.Geometry.Direction.AsVector
            Dim pt2 As Point2d = skl.EndSketchPoint.Geometry
            Dim v2 As Vector2d = v
            v2.ScaleBy(25 / 10)
            pt2.TranslateBy(v2)
            Dim spt2 As SketchPoint = spl.SketchPoints.Add(pt2)
            Dim spt As SketchPoint = spl.SketchPoints.Add(skl.Geometry.MidPoint)
            Dim m As Matrix2d = tg.CreateMatrix2d
            m.SetToRotation(Math.PI / 2, spt2.Geometry)
            v.TransformBy(m)
            v.ScaleBy(skl.Length * 1 / 2)
            Dim pt3 As Point2d = pt2
            pt3.TranslateBy(v)
            spl.SketchPoints.Add(pt3)
            spl.SketchLines.AddAsThreePointCenteredRectangle(spt, spt2, pt3)

            Dim pro As Profile

            pro = spl.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            oExtrudeDef = doku.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(pro, PartFeatureOperationEnum.kCutOperation)
            oExtrudeDef.SetDistanceExtent(25 / 10, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            oExtrudeDef.SetDistanceExtentTwo(d)
            Dim oExtrude As ExtrudeFeature
            oExtrude = doku.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef)



            Return oExtrude
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


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
            converter.SetUnitsToMetric(p)
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
        Dim ps As PlanarSketch
        Dim spt2d As SketchPoint
        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        Dim ed As Edge = GetClosestEdge(fc)
        'GetRails(fc)
        sections.Add(pr)
        Dim pt As Point


        Dim v As Vector
        v = GetRealNormal(fc, doku.ComponentDefinition.WorkSurfaces.Item(1))
        v.ScaleBy(-8 * fc.Evaluator.Area)
        ' ps = doku.ComponentDefinition.Sketches.Add(fc)
        '   spt2d = ps.AddByProjectingEntity(wpConverge)
        '   ps.Visible = False
        pt = GetOppositeVertex(ed, fc).Point
        'pt.TranslateBy(v)
        Dim wptlf As WorkPoint = doku.ComponentDefinition.WorkPoints.AddFixed(pt)
        wptlf.Visible = False
        sections.Add(wptlf)
        Return LoftSingleSpike()
    End Function
    Function GetOppositeVertex(edi As Edge, fi As Face) As Vertex
        Dim ver1, ver2 As Vertex
        Dim ws As WorkSurface = compDef.WorkSurfaces.Item(1)
        Dim sb As SurfaceBody = ws.SurfaceBodies.Item(1)
        Dim vnpl, vnfi, v As Vector
        Dim d, e, eMin, dMin1, dMin2, dis As Double
        Dim pt1, pt2, pt3 As Point
        Dim ls As LineSegment
        Dim ed2 As Edge
        vnfi = GetRealNormal(fi, ws)
        vnfi.Normalize()
        pt1 = fi.PointOnFace

        'sk3D.SketchPoints3D.Add(pt1)
        dMin1 = 9999
        dMin2 = 9999
        eMin = 99999
        ver1 = fi.Vertices.Item(1)
        ls = tg.CreateLineSegment(edi.StartVertex.Point, edi.StopVertex.Point)
        pt3 = ls.MidPoint
        sk3D.SketchPoints3D.Add(pt3)
        Try
            For Each f As Face In fi.TangentiallyConnectedFaces
                If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                    If Not (fi.Equals(f)) Then
                        vnpl = GetRealNormal(f, ws)
                        vnpl.Normalize()
                        d = vnpl.DotProduct(vnfi)
                        If d < -0.5 Then
                            ed2 = GetClosestEdge(f)
                            pt2 = ed2.GetClosestPointTo(pt3)
                            sk3D.SketchPoints3D.Add(pt2)
                            e = pt3.DistanceTo(pt2)
                            If e < 18 / 10 Then
                                If e < eMin Then
                                    eMin = e
                                    pt1 = ed2.GetClosestPointTo(wpConverge.Point)
                                    sk3D.SketchPoints3D.Add(pt1)
                                    For Each vr As Vertex In f.Vertices
                                        dis = pt1.DistanceTo(vr.Point)
                                        If dis < dMin2 Then
                                            If dis < dMin1 Then
                                                dMin2 = dMin1
                                                dMin1 = dis
                                                ver2 = ver1
                                                ver1 = vr
                                                'lamp.HighLighObject(vr)
                                            Else
                                                dMin2 = dis
                                                ver2 = vr

                                            End If

                                        End If
                                    Next

                                End If


                            End If
                        End If
                    End If
                End If
            Next
            sk3D.Visible = False
            lamp.HighLighObject(ver1)
            Return ver1
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

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
