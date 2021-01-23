Imports Inventor
Imports System
Imports Subina_Design_Helpers


Public Class InitSketcher
    Dim doku As PartDocument
    Dim app As Application
    Public sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Public refLine, firstLine, secondLine, thirdLine, lastLine, twistLine, inputLine, kanteLine As SketchLine3D
    Public curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim curve3D As Curves3D
    Dim monitor As DesignMonitoring
    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3 As Point
    Dim tg As TransientGeometry
    Dim gapFoldCM As Double
    Dim adjuster As SketchAdjust
    Public bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim nombrador As Nombres
    Dim dimConstrainBandLine2, gapFold, dcThirdLine, acCl7Bl4 As DimensionConstraint3D
    Dim startPoint, cl3Equal As GeometricConstraint3D
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim compDef As SheetMetalComponentDefinition
    Public ID1, ID2, qvalue, direction As Integer
    Dim k1(), k2() As Byte
    Dim backwards As Boolean
    Dim lamp As Highlithing




    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        lamp = New Highlithing(doku)
        curve3D = New Curves3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        gapFoldCM = 3 / 10
        done = False
    End Sub
    Public Sub New(docu As Inventor.Document, tl As SketchLine3D)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        compDef = doku.ComponentDefinition
        sheetMetalFeatures = compDef.Features
        curve3D = New Curves3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        gapFoldCM = 3 / 10
        twistLine = tl
        done = False
    End Sub



    Public Function DrawNextMainSketch(rl As SketchLine3D, pt As Point) As SketchLine3D

        twistLine = rl
        sk3D = doku.ComponentDefinition.Sketches3D.Item(doku.ComponentDefinition.Sketches3D.Count)
        refLine = sk3D.SketchLines3D.AddByTwoPoints(twistLine.EndPoint, pt, False)
        refLine.Construction = True
        curve = sk3D.SketchEquationCurves3D.Item(sk3D.SketchEquationCurves3D.Count)
        If refLine.Length > 0 Then
            If DrawSingleLines() Then
                point1 = firstLine.StartSketchPoint.Geometry
                point2 = firstLine.EndSketchPoint.Geometry
                point3 = secondLine.EndSketchPoint.Geometry
                For Each sl As SketchLine3D In sk3D.SketchLines3D
                    If sl.Equals(firstLine) Then
                        sk3D = doku.ComponentDefinition.Sketches3D.Add()
                        sk3D.Name = "firstLine"
                        sk3D.Include(sl)

                    Else

                        If sl.Equals(secondLine) Then
                            sk3D = doku.ComponentDefinition.Sketches3D.Add()
                            sk3D.Name = "secondLine"
                            sk3D.Include(sl)

                        End If
                    End If
                Next
                inputLine = firstLine
                done = 1
            End If
        End If

        Return firstLine
    End Function
    Public Function DrawMainSketch(refDoc As FindReferenceLine, q As Integer, direction As Integer) As Sketch3D

        Dim s As String
        Dim ac As DimensionConstraint3D

        refLine = refDoc.GetKeyLine()
        kanteLine = refDoc.GetKanteLine()
        doku.Activate()
        compDef = doku.ComponentDefinition
        sheetMetalFeatures = compDef.Features
        curve = DrawTrobinaCurveFitted(q, direction)
        lamp.ZoomSelected(curve)
        DrawInitialLine(refLine)
        refLine.Construction = True
        If refLine.Length > 0 Then
            If DrawSingleLines() Then
                point1 = firstLine.StartSketchPoint.Geometry
                point2 = firstLine.EndSketchPoint.Geometry
                point3 = secondLine.EndSketchPoint.Geometry
                ac = AdjustlastAngle()
                If gapFold.Parameter._Value > gapFoldCM * 2 Then
                    CorrectLastAngle(ac)
                End If
                If gapFold.Parameter._Value < gapFoldCM * 4 Then
                    Return sk3D
                Else
                    Return Nothing
                End If
            End If
        End If
        Return Nothing

    End Function
    Function DrawReferenceCurve(sk As Sketch3D) As SketchEquationCurve3D

        Return Nothing
    End Function



    Function DrawSingleLines() As Boolean
        Try
            If DrawFirstLine().Length > 0 Then
                If DrawSecondLine().Length > 0 Then
                    If DrawFirstConstructionLine().Construction Then
                        If DrawThirdLine().Length > 0 Then
                            If DrawSecondConstructionLine().Construction Then
                                If DrawFourthLine().Length > 0 Then
                                    If DrawFifthLine().Length > 0 Then
                                        If DrawThirdConstructionLine().Construction Then
                                            If DrawFourthConstructionLine().Construction Then
                                                'Return True
                                                If DrawSixthLine().Length > 0 Then
                                                    If DrawSeventhLine().Length > 0 Then
                                                        Return True
                                                    End If
                                                End If
                                            End If

                                        End If
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
        Return False
    End Function
    Function DrawInitialLine(line As SketchLine3D) As SketchLine3D
        Dim l As SketchLine3D
        l = sk3D.SketchLines3D.AddByTwoPoints(line.StartSketchPoint.Geometry, line.EndSketchPoint.Geometry)
        l.Construction = True
        sk3D.GeometricConstraints3D.AddGround(l)
        Return line
    End Function
    Function DrawTrobinaCurve(q As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(q, "s1")
    End Function
    Function DrawTrobinaCurveFitted(q As Integer, s As Integer, direction As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(q, "s1", s, direction)
    End Function
    Function DrawTrobinaCurveFitted(q As Integer, direction As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(q, "s1", 0, direction)
    End Function
    Function DrawTrobinaCurve(q As Integer, s As String) As SketchEquationCurve3D

        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = curve3D.DrawTrobinaCurve(sk3D, q, direction)
        sk3D.GeometricConstraints3D.AddGround(curve)
        Return curve
    End Function
    Function DrawTrobinaCurve(q As Integer, s As String, f As Integer, direction As String) As SketchEquationCurve3D

        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = curve3D.DrawTrobinaCurve(sk3D, q, f, direction)
        sk3D.GeometricConstraints3D.AddGround(curve)
        Return curve
    End Function
    Function DrawFirstLine() As SketchLine3D
        Dim r1, r2 As Double
        Dim dc As DimensionConstraint3D
        Dim dir As Integer = GetParameter("direction")._Value
        Try

            Dim l As SketchLine3D
            Dim gc, gpc As GeometricConstraint3D
            Dim pt1, pt2 As Point
            r1 = TorusRadius(refLine.StartSketchPoint.Geometry)
            r2 = TorusRadius(refLine.EndSketchPoint.Geometry)
            If r1 > r2 Then
                pt1 = refLine.StartSketchPoint.Geometry
                pt2 = refLine.EndSketchPoint.Geometry
            Else
                pt1 = refLine.EndSketchPoint.Geometry
                pt2 = refLine.StartSketchPoint.Geometry
            End If

            If (doku.FullFileName.Contains("Band1.ipt") And (dir < 0)) Then
                ' pt1 = GetStartingPoint().Geometry
            End If

            'rl = sk3D.Include(refLine)
            l = sk3D.SketchLines3D.AddByTwoPoints(pt1, pt2)
            If (doku.FullFileName.Contains("Band1.ipt") And (dir < 0)) Then
                'dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.StartPoint, curve.StartSketchPoint)
                'adjuster.AdjustDimConstrain3DSmothly(dc, dc.Parameter._Value / 4)
                'dc.Delete()
            End If
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            gc = sk3D.GeometricConstraints3D.AddGround(l.StartPoint)

            Try
                '  gpc = sk3D.GeometricConstraints3D.AddParallel(l, refLine)
            Catch ex As Exception

            End Try




            'lamp.FitView(doku)

            point1 = l.StartSketchPoint.Geometry
            point2 = l.EndSketchPoint.Geometry
            bandLines.Add(l)
            firstLine = l
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If l.Length > 25 / 10 Then
                adjuster.AdjustDimConstrain3DSmothly(dc, 25 / 10)
            End If
            dc.Delete()
            ForceFirstLineInside()
            lastLine = l

            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function
    Function GetStartingPoint() As SketchPoint3D
        Dim c As Cylinder
        Dim r As Double = 40 / 10
        c = tg.CreateCylinder(compDef.WorkPoints.Item(1).Point, compDef.WorkAxes.Item(3).Line.Direction, r)
        Dim zMax As Double = 0
        Dim spt As SketchPoint3D
        Try
            For Each pt As Point In tg.CurveSurfaceIntersection(curve.Geometry, c)
                spt = sk3D.SketchPoints3D.Add(pt)
                If spt.Geometry.Z > zMax Then
                    zMax = spt.Geometry.Z
                    GetStartingPoint = spt
                Else

                End If

            Next
            Return GetStartingPoint
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try




    End Function
    Function IsFirstLineInside() As Boolean

        If firstLine.EndSketchPoint.Geometry.DistanceTo(curve.EndSketchPoint.Geometry) < firstLine.StartSketchPoint.Geometry.DistanceTo(curve.EndSketchPoint.Geometry) Then
            Return True
        Else
            Return False

        End If

    End Function
    Function ForceFirstLineInside() As Boolean
        Dim limit As Integer = 0
        Dim dc As DimensionConstraint3D
        Try
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(firstLine.EndPoint, curve.EndSketchPoint)
            If dc.Parameter._Value > 75 / 10 Then
                adjuster.AdjustDimConstrain3DSmothly(dc, 75 / 10)
            End If
            While (Not IsFirstLineInside() And limit < 8)
                Try

                    adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 3 / 4)


                    limit += 1
                Catch ex As Exception
                    limit += 1
                End Try

            End While
            dc.Delete()
        Catch ex As Exception
            Try
                dc.Delete()
            Catch ex2 As Exception

            End Try
        End Try
        Return IsFirstLineInside()
    End Function
    Function CalculateOutPostionFactor(pt As Point) As Double

        Return (GetRadiusPoint(pt) * TorusRadius(pt)) / (Math.Pow(Math.Abs(pt.Z) + 0.000001, 1 / 2))
    End Function
    Function TorusRadius(pt As Point) As Double

        Return Math.Pow(Math.Pow(Math.Pow(Math.Pow(pt.X, 2) + Math.Pow(pt.Y, 2), 1 / 2) - 50 / 10, 2) + Math.Pow(pt.Z, 2), 1 / 2)
    End Function
    Function GetRadiusPoint(pt As Point) As Double
        Dim radius As Double = Math.Pow(Math.Pow(pt.X, 2) + Math.Pow(pt.Y, 2), 1 / 2)
        Return radius
    End Function
    Function DrawFirstFloatingLine() As SketchLine3D
        Try

            Dim l As SketchLine3D = Nothing
            Dim dc As DimensionConstraint3D
            Dim gc As GeometricConstraint3D


            l = sk3D.SketchLines3D.AddByTwoPoints(refLine.StartPoint, refLine.EndSketchPoint.Geometry, False)
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            dc.Parameter._Value = curve3D.DP.b / 10
            sk3D.Solve()
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, twistLine)
            sk3D.Solve()
            gc.Delete()
            dimConstrainBandLine2 = dc
            bandLines.Add(l)
            firstLine = l
            lastLine = l

            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function
    Function DrawSecondLine() As SketchLine3D
        Try
            Dim v1, v2, v3 As Vector
            Dim p As Plane
            Dim ls As LineSegment
            Dim l As SketchLine3D = Nothing
            Dim optpoint As Point = curve.EndSketchPoint.Geometry
            Dim dc, ac, dcfl As DimensionConstraint3D
            Dim gc As GeometricConstraint3D
            Dim d As Double
            Dim b As Double = GetParameter("b")._Value
            v1 = lastLine.StartSketchPoint.Geometry.VectorTo(lastLine.EndSketchPoint.Geometry)
            v3 = curve.StartSketchPoint.Geometry.VectorTo(curve.EndSketchPoint.Geometry)
            If firstLine.Length < b Then
                p = tg.CreatePlane(lastLine.EndSketchPoint.Geometry, v1)
            Else
                ls = firstLine.Geometry
                p = tg.CreatePlane(ls.MidPoint, v1)
            End If

            Dim minDis As Double = 9999999999
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
                v2 = lastLine.EndSketchPoint.Geometry.VectorTo(o)
                d = v1.CrossProduct(v2).Length * v2.DotProduct(v3)
                If d > 0 Then
                    If o.DistanceTo(lastLine.EndSketchPoint.Geometry) < minDis Then
                        minDis = o.DistanceTo(lastLine.EndSketchPoint.Geometry)
                        optpoint = o
                    End If

                End If
                'l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, o, False)


            Next
            If Math.Abs(firstLine.Length - b) < 1 / b Then
                dc = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value * 24 / 25)
                dc.Delete()
            End If
            Try
                l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, optpoint, False)
            Catch ex As Exception
                dc = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value * 23 / 25)
                dc.Delete()
                l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, optpoint, False)

            End Try

            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            If l.Length > GetParameter("b")._Value * 3 / 2 Then
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, firstLine)
                dcfl = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 3 / 4)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 3 / 4)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value * 4 / 3)
                dc.Delete()
                ac.Delete()
                dcfl.Delete()

            End If
            gc = TryPerpendicular(l, firstLine)
            gc.Delete()
            dimConstrainBandLine2 = sk3D.DimensionConstraints3D.AddLineLength(l)
            dimConstrainBandLine2.Driven = True
            CorrectSecondLine()

            lastLine = l
            bandLines.Add(l)
            secondLine = lastLine


            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawFirstConstructionLine() As SketchLine3D
        Dim gc As GeometricConstraint3D
        Try
            Dim l As SketchLine3D = Nothing
            Dim endPoint As Point
            Dim v As Vector
            v = firstLine.Geometry.Direction.AsVector
            v.ScaleBy(GetParameter("b")._Value / 1)
            endPoint = firstLine.StartSketchPoint.Geometry
            endPoint.TranslateBy(v)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.StartPoint, endPoint, False)
            endPoint.TranslateBy(v)

            sk3D.DimensionConstraints3D.AddLineLength(l, endPoint, False)
            gc = TryPerpendicular(l, secondLine)
            gc.Delete()
            sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, secondLine)
            l.Construction = True
            constructionLines.Add(l)



            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawThirdConstructionLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing
            Dim bl4, cl As SketchLine3D
            Dim dc, dc2 As DimensionConstraint3D
            Dim b As Double = GetParameter("b")._Value
            bl4 = bandLines.Item(4)
            cl = constructionLines.Item(1)
            CorrectSecondLine()
            AdjustThirdLine(2 * b)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndPoint, bl4.Geometry.MidPoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, bl4)
            dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
            adjuster.AdjustDimConstrain3DSmothly(dc, Math.PI / 2)
            dc.Delete()
            sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
            AdjustThirdLine(2 * b)

            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            dc2 = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
            adjuster.AdjustDimConstrain3DSmothly(dc, 25 / 10)
            dc.Delete()
            Try
                cl3Equal = sk3D.GeometricConstraints3D.AddEqual(l, cl)
            Catch ex As Exception
                gapFold.Driven = True
                CorrectSecondLine()
                adjuster.AdjustDimConstrain3DSmothly(gapFold, 3 * gapFoldCM)
                gapFold.Driven = True
                cl3Equal = sk3D.GeometricConstraints3D.AddEqual(l, cl)
                adjuster.AdjustDimConstrain3DSmothly(gapFold, 3 * gapFoldCM)
                gapFold.Driven = False
            End Try

            dc2.Delete()
            l.Construction = True
            sk3D.Solve()
            constructionLines.Add(l)

            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawFourthConstructionLine() As SketchLine3D
        Try
            Dim l, fl As SketchLine3D
            Dim bl4, cl3, cl1 As SketchLine3D
            Dim dc, dc2, dcaux, ac As DimensionConstraint3D
            Dim gcel As GeometricConstraint3D
            bl4 = bandLines.Item(4)
            cl3 = constructionLines.Item(3)
            cl1 = constructionLines.Item(1)
            Dim ls As LineSegment = thirdLine.Geometry
            Dim b As Double = GetParameter("b")._Value
            AdjustThirdLine(2 * b)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndSketchPoint.Geometry, bl4.StartSketchPoint.Geometry, False)
            Try
                sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            Catch ex As Exception
                doku.Update()
                l.Delete()
                l = sk3D.SketchLines3D.AddByTwoPoints(ls.MidPoint, bl4.StartSketchPoint.Geometry, False)

                sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            End Try

            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, bl4)
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.EndPoint, cl3.EndPoint)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value / 2)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, 1 / 10)
            dc2 = sk3D.DimensionConstraints3D.AddLineLength(l)
            If dc2.Parameter._Value > curve3D.DP.b * 3 / 20 Then
                If adjuster.AdjustDimensionConstraint3DSmothly(dc2, curve3D.DP.b * 5 / 40) Then
                    dc2.Driven = True
                Else
                    dc2.Delete()
                End If
            Else
                dc2.Driven = True
            End If
            dcaux = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
            adjuster.AdjustDimConstrain3DSmothly(dcaux, 2 * b)
            dcaux.Delete()
            gapFold.Driven = True
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Driven = True
            Try
                sk3D.Solve()
            Catch ex As Exception
                Try
                    cl3Equal.Delete()
                    dcaux = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
                    adjuster.AdjustDimConstrain3DSmothly(dcaux, 2 * b)
                    dcaux.Delete()
                    dcaux = sk3D.DimensionConstraints3D.AddLineLength(cl3)
                    adjuster.AdjustDimConstrain3DSmothly(dcaux, b)
                    dcaux.Delete()
                    cl3Equal = sk3D.GeometricConstraints3D.AddEqual(cl1, cl3)
                    sk3D.Solve()
                Catch ex2 As Exception
                    cl3Equal.Delete()
                    dcaux = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
                    adjuster.AdjustDimConstrain3DSmothly(dcaux, 2 * b)
                    dcaux.Delete()
                    dcaux = sk3D.DimensionConstraints3D.AddLineLength(cl3)
                    adjuster.AdjustDimConstrain3DSmothly(dcaux, b)
                    sk3D.Solve()
                End Try
            End Try
            sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
            dc2.Driven = False
            adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
            dc2.Driven = True
            gapFold.Driven = False
            Try
                sk3D.GeometricConstraints3D.AddEqual(l, cl3)
            Catch ex2 As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 4)
                gapFold.Driven = True
                dc2.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                dc2.Driven = True
                gapFold.Driven = False
                Try
                    gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)

                Catch ex3 As Exception
                    adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 4)
                    gapFold.Driven = True
                    dc2.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                    dc2.Driven = True
                    gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                End Try

                gapFold.Driven = False
                Try
                    gapFold.Parameter._Value = gapFold.Parameter._Value * 1.1
                    adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM)
                    If gapFold.Parameter._Value < gapFoldCM * 0.9 Then
                        gcel.Delete()
                        adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM)
                        adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM)
                    End If
                Catch ex As Exception
                    gcel.Delete()
                    adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM)
                    adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM)
                End Try

                adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM)
                adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM)
            End Try
            gapFold.Driven = True
            l.Construction = True

            constructionLines.Add(l)

            lastLine = l
            ac.Delete()
            dc2.Delete()
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function AdjustThirdLine(bi As Double) As DimensionConstraint3D
        Dim dc As DimensionConstraint3D
        Dim b As Double = GetParameter("b")._Value

        If thirdLine.Length > 2 * b Or thirdLine.Length < 3 * b / 2 Then
            dc = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, bi)
            dc.Delete()
        End If
        Return dc
    End Function
    Function DrawThirdLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing
            Dim pl As SketchLine3D = Nothing
            pl = sk3D.SketchLines3D.Item(sk3D.SketchLines3D.Count - 1)
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.EndPoint, firstLine.StartPoint, False)
            thirdLine = l
            bandLines.Add(l)
            CorrectThirdLine()
            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CorrectSecondLine() As Boolean
        Dim dc As DimensionConstraint3D
        Dim dis As Double = firstLine.StartSketchPoint.Geometry.DistanceTo(curve.StartSketchPoint.Geometry)
        Dim b As Double = GetParameter("b")._Value
        dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(firstLine.EndPoint, curve.StartSketchPoint)
        For i = 1 To 8
            If (dc.Parameter._Value > dis Or dc.Parameter._Value < dis / 2) Then
                If dc.Parameter._Value > 2 * b Then
                    CorrectSecondLine = adjuster.AdjustDimConstrain3DSmothly(dc, 7 * dc.Parameter._Value / 8)
                Else
                    CorrectSecondLine = adjuster.AdjustDimConstrain3DSmothly(dc, 9 * dc.Parameter._Value / 8)
                End If
            Else

                Exit For
            End If
        Next
        dc.Delete()
        Return CorrectSecondLine
    End Function
    Function CorrectThirdLine() As Boolean
        Dim dc, dc2 As DimensionConstraint3D
        Dim b As Double = GetParameter("b")._Value
        dc = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
        For i = 1 To 4
            If (thirdLine.Length > 2 * b) Or (thirdLine.Length < b) Then
                CorrectThirdLine = adjuster.AdjustDimConstrain3DSmothly(dc, 2 * b)
                dc2 = sk3D.DimensionConstraints3D.AddTwoPointDistance(firstLine.EndPoint, curve.EndSketchPoint)
                If dc2.Parameter._Value > 4 * b Then
                    CorrectThirdLine = adjuster.AdjustDimConstrain3DSmothly(dc2, 4 * b)
                End If
                dc2.Delete()
            Else
                dc2 = sk3D.DimensionConstraints3D.AddTwoPointDistance(firstLine.EndPoint, curve.EndSketchPoint)
                If dc2.Parameter._Value > 4 * b Then
                    'CorrectThirdLine = adjuster.AdjustDimConstrain3DSmothly(dc2, 3 * b)
                End If
                dc2.Delete()
                Exit For
            End If
        Next
        dc.Delete()
        Return CorrectThirdLine
    End Function
    Function DrawSecondConstructionLine() As SketchLine3D
        Try
            Dim v1, v2, v3, vk As Vector
            Dim l, kl As SketchLine3D
            Dim endPoint As Point
            Dim dc As DimensionConstraint3D
            Dim dir As Integer = GetParameter("direction")._Value
            Dim currentQ As Integer = GetParameter("currentQ")._Value
            Dim oddQ As Integer
            Dim gc As GeometricConstraint3D

            CorrectSecondLine()
            CorrectThirdLine()

            kl = sk3D.SketchLines3D.AddByTwoPoints(kanteLine.StartSketchPoint.Geometry, kanteLine.EndSketchPoint.Geometry, False)
            vk = kl.Geometry.Direction.AsVector
            endPoint = secondLine.StartSketchPoint.Geometry
            If (doku.FullFileName.Contains("Band1.ipt")) Then
                v1 = firstLine.EndSketchPoint.Geometry.VectorTo(firstLine.StartSketchPoint.Geometry)
                v2 = secondLine.EndSketchPoint.Geometry.VectorTo(secondLine.StartSketchPoint.Geometry)
                v3 = v1.CrossProduct(v2).AsUnitVector().AsVector()
                v3.ScaleBy(gapFoldCM)
                endPoint.TranslateBy(v3)
            Else
                vk.ScaleBy(-1)
                endPoint.TranslateBy(vk)

            End If


            kl.Construction = True
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, endPoint, False)

            gapFold = sk3D.DimensionConstraints3D.AddLineLength(l)
            Try
                gapFold.Parameter._Value = gapFoldCM * 5
            Catch ex As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM)
                gapFold.Parameter._Value = gapFoldCM * 5
            End Try
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            'sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, firstLine)
            Try
                gc = TryPerpendicular(l, secondLine)
            Catch ex As Exception
                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, secondLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI / 2)
                dc.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            End Try
            Try
                gc = TryPerpendicular(l, firstLine)
                gc.Delete()
            Catch ex As Exception

            End Try

            CorrectSecondLine()
            CorrectThirdLine()

            lastLine = l
            constructionLines.Add(l)
            l.Construction = True
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function
    Function DrawFourthLine() As SketchLine3D
        Try
            Dim l, pl As SketchLine3D
            Dim dcl, acl, dcl4 As DimensionConstraint3D
            Dim b As Double = GetParameter("b")._Value
            CorrectThirdLine()
            CorrectSecondLine()
            CorrectThirdLine()

            pl = firstLine
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.StartPoint, lastLine.EndSketchPoint.Geometry, False)
            Try

                TryPerpendicular(l, lastLine)
            Catch ex As Exception
                acl = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, lastLine)
                adjuster.AdjustDimensionConstraint3DSmothly(acl, Math.PI / 2)
                acl.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
            End Try
            Try
                sk3D.GeometricConstraints3D.AddCoincident(lastLine.EndPoint, l)
            Catch ex As Exception
                dcl = sk3D.DimensionConstraints3D.AddTwoPointDistance(lastLine.EndPoint, l.EndPoint)
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, gapFoldCM)
                dcl.Delete()
                sk3D.GeometricConstraints3D.AddCoincident(lastLine.EndPoint, l)
            End Try
            dcl4 = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.EndPoint, curve.EndSketchPoint)
            adjuster.AdjustDimConstrain3DSmothly(dcl4, gapFoldCM * 4)
            Try
                AdjustThirdLine(2 * b)
            Catch ex As Exception

            End Try
            dcl4.Delete()
            Dim gc As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)


            gc.Delete()
            CorrectSecondLine()
            CorrectThirdLine()
            'CorrectSecondLine()
            lastLine = l
            bandLines.Add(l)
            AdjustFourLine()
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function

    Function TryPerpendicular(l As SketchLine3D, bl As SketchLine3D) As GeometricConstraint3D
        Dim ac As DimensionConstraint3D
        Dim gc As GeometricConstraint3D

        Try
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(bl, l)
            Try
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try

        Try
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, bl)
        Catch ex As Exception

        End Try



        Return gc
    End Function
    Function AdjustFourLine() As Boolean
        Dim b As Double = GetParameter("b")._Value
        Dim bl4 As SketchLine3D = bandLines(4)
        Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl4.EndPoint, curve.EndSketchPoint)
        For i = 1 To 8
            If bl4.Length < thirdLine.Length Then
                gapFold.Driven = True
                AdjustFourLine = adjuster.AdjustDimConstrain3DSmothly(dc, b)
                gapFold.Driven = False
            Else
                Exit For
            End If
        Next
        dc.Delete()
        Return AdjustFourLine
    End Function
    Function DrawFourthFloatingLine() As SketchLine3D
        Try
            Dim l, pl As SketchLine3D

            pl = firstLine
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.StartPoint, lastLine.EndSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
            sk3D.GeometricConstraints3D.AddCoincident(lastLine.EndPoint, l)
            'sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawFifthLine() As SketchLine3D
        Try
            Dim l, pl As SketchLine3D
            Dim gc As GeometricConstraint3D
            Dim dc As DimensionConstraint3D
            CorrectSecondLine()
            pl = thirdLine
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, pl.StartPoint, False)
            dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, thirdLine)
            Try

                adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI / 2)
            Catch ex As Exception
                dc.Driven = True
            End Try
            dc.Delete()

            Try
                gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, pl)
                sk3D.Solve()
                gc.Delete()
            Catch ex As Exception
                If gc.Deletable Then
                    gc.Delete()

                End If
            End Try


            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawSixthLine() As SketchLine3D
        Try
            Dim l, cl3, ml, bl4 As SketchLine3D
            Dim v As Vector
            Dim p As Plane
            Dim minDis, d As Double
            Dim optpoint As Point
            Dim dc, ac As DimensionConstraint3D
            Dim gc As GeometricConstraint3D
            Dim b As Double = GetParameter("b")._Value
            dcThirdLine = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
            If dcThirdLine.Parameter._Value > curve3D.DP.b * 2 / 10 Then
                If adjuster.AdjustDimensionConstraint3DSmothly(dcThirdLine, dcThirdLine.Parameter._Value * 4 / 5) Then
                    dcThirdLine.Driven = True
                Else
                    dcThirdLine.Delete()
                End If
            End If

            cl3 = constructionLines.Item(3)
            v = cl3.Geometry.Direction.AsVector
            p = tg.CreatePlane(cl3.EndSketchPoint.Geometry, v)

            Dim puntos As ObjectsEnumerator
            puntos = p.IntersectWithCurve(curve.Geometry)
            minDis = 9999999999
            Dim o2 As Point = puntos.Item(puntos.Count)
            optpoint = o2
            ml = bandLines.Item(4)
            Dim vc, vmjl, vbl4 As Vector
            vmjl = thirdLine.Geometry.Direction.AsVector
            vmjl.ScaleBy(-1)
            bl4 = bandLines.Item(4)
            Try
                gapFold.Driven = False
                dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl4.EndPoint, curve.EndSketchPoint)
                adjuster.AdjustDimConstrain3DSmothly(dc, dc.Parameter._Value / 2)
                dc.Delete()
                gc = sk3D.GeometricConstraints3D.AddCoincident(bl4.EndPoint, curve)
                gc.Delete()
                dc = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                adjuster.AdjustDimConstrain3DSmothly(dc, dc.Parameter._Value * 3 / 4)
                dc.Delete()
            Catch ex As Exception

            End Try
            vbl4 = bl4.Geometry.Direction.AsVector
            For Each o As Point In puntos
                'l = sk3D.SketchLines3D.AddByTwoPoints(cl3.EndSketchPoint.Geometry, o, False)
                vc = cl3.EndSketchPoint.Geometry.VectorTo(o)
                d = vc.CrossProduct(vmjl).Length * vc.DotProduct(vbl4)
                If d > 0 Then
                    If o.DistanceTo(cl3.EndSketchPoint.Geometry) < minDis Then
                        minDis = o.DistanceTo(cl3.EndSketchPoint.Geometry)

                        optpoint = o
                    End If
                End If


            Next

            vp = optpoint

            bl4 = bandLines.Item(4)
            l = sk3D.SketchLines3D.AddByTwoPoints(bl4.EndSketchPoint.Geometry, vp, False)
            gc = sk3D.GeometricConstraints3D.AddCoincident(bl4.EndPoint, l)
            gc = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            Try
                ' ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)

                Try
                    If l.Length > 4 * b / 2 Or l.Length < b Then
                        dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                        adjuster.AdjustDimConstrain3DSmothly(dc, 3 * b / 2)
                        dc.Delete()
                    End If
                Catch ex As Exception

                End Try
                ' ac.Delete()
            Catch ex As Exception

            End Try

            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawSeventhLine() As SketchLine3D
        Try
            Dim l, cl As SketchLine3D
            cl = constructionLines.Item(3)
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.StartPoint, secondLine.EndPoint, False)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(lastLine)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 4 / 3) Then
                dc.Delete()
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
                    dc.Delete()
                    sk3D.GeometricConstraints3D.AddEqual(l, cl)

                End If
            End If


            lastLine = l
            bandLines.Add(l)
            done = 1
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function

    Function AdjustlastAngle() As DimensionConstraint3D
        Try
            Dim fourLine, sixthLine, cl2, cl3, bl5 As SketchLine3D

            Dim ac, dc, acl4l6 As DimensionConstraint3D
            Dim gc As GeometricConstraint3D
            Dim limit As Double = Math.PI / 24
            Dim angleLimit As Double = 1.8
            Dim d As Double
            Dim counterLimit As Integer = 0
            Dim lastAngle As Double
            Dim hecho As Boolean = False


            Dim b As Boolean = False

            fourLine = bandLines.Item(4)
            bl5 = bandLines.Item(5)
            sixthLine = bandLines.Item(6)
            cl2 = constructionLines.Item(2)
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(thirdLine, bl5)
            Try
                dcThirdLine.Delete()
            Catch ex As Exception

            End Try
            dc = sk3D.DimensionConstraints3D.AddLineLength(sixthLine)
            Try
                acl4l6 = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)

            Catch ex As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 17 / 16)
                sk3D.Solve()

                acl4l6 = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)

            End Try
            ac.Driven = True
            acl4l6.Driven = True
            Try
                While ((acl4l6.Parameter._Value > angleLimit And acl4l6.Parameter._Value < Math.PI - limit) And counterLimit < 8)

                    adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2 * (1 - Math.Exp(-16 * ac.Parameter._Value / (Math.PI * 1))))
                    counterLimit = counterLimit + 1
                End While
            Catch ex As Exception

            End Try

            acl4l6.Delete()
            dc.Delete()
            dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)
            dc.Driven = True
            d = CalculateRoof()
            If d > 0 Then
                dc.Driven = True
                Try
                    If adjuster.IsLastAngleOk(dc, limit) Then
                        For index = 1 To 16
                            adjuster.AdjustDimConstrain3DSmothly(gapFold, gapFold.Parameter._Value * 17 / 16)
                            adjuster.AdjustDimConstrain3DSmothly(dc, dc.Parameter._Value * 15 / 16)
                            If (Not adjuster.IsLastAngleOk(dc, limit)) Or (gapFold.Parameter._Value > 6 * gapFoldCM) Then
                                Exit For
                            End If
                        Next















































                    End If
                    dc.Driven = True
                    If Not monitor.IsSketch3DHealthy(sk3D) Then
                        adjuster.RecoverUnhealthy3DSketch(sk3D)
                    End If
                    dc = AdjustPositiveAngle(dc, limit)
                    Try
                        counterLimit = 0
                        While (dc.Parameter._Value < angleLimit And dc.Parameter._Value > angleLimit / 2) And counterLimit < 4
                            lastAngle = ac.Parameter._Value
                            adjuster.AdjustGapSmothly(gapFold, gapFoldCM * 5 / 4, dc)
                            If ac.Parameter._Value < lastAngle Then
                                counterLimit = 4
                            End If
                            counterLimit = counterLimit + 1
                        End While
                        counterLimit = 0
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        ac.Driven = True

                    End Try
                    b = True
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                    dc.Driven = True
                End Try
                dc.Driven = True
            Else
                Try
                    If dc.Parameter._Value < Math.PI / 2 Then
                        Try
                            gc.Delete()
                            dc.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI * 3 / 4)
                            dc.Driven = True
                            Try
                                gc = sk3D.GeometricConstraints3D.AddPerpendicular(thirdLine, bl5)
                            Catch ex As Exception

                            End Try

                        Catch ex As Exception
                            Try
                                dc.Driven = True
                                gc = sk3D.GeometricConstraints3D.AddPerpendicular(thirdLine, bl5)
                            Catch ex9 As Exception

                            End Try
                        End Try
                    End If
                    dc.Driven = True
                    While ((d < 0 Or Not IsLastAngleOk(dc, 3 * limit / 4)) And counterLimit < 32)
                        Try
                            If adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value * 17 / 16) Then
                            Else
                                ac.Driven = True
                            End If
                            If d > 0 Then
                                If dc.Parameter._Value > Math.PI / 2 Then
                                    adjuster.AdjustDimConstrain3DSmothly(dc, Math.PI - 3 * limit / 4)
                                Else
                                    adjuster.AdjustDimConstrain3DSmothly(dc, 3 * limit / 4)
                                End If

                            Else
                                If gapFold.Parameter._Value > 15 / 10 And Not hecho Then
                                    If dc.Parameter._Value > Math.PI / 2 Then
                                        If adjuster.AdjustDimConstrain3DSmothly(dc, Math.PI) Then
                                            hecho = True
                                        End If

                                    Else
                                        If adjuster.AdjustDimConstrain3DSmothly(dc, 1 / 128) Then
                                            hecho = True
                                        End If

                                    End If


                                End If
                            End If
                            dc.Driven = True

                        Catch ex2 As Exception
                            counterLimit = counterLimit + 1
                            gapFold.Driven = True
                        End Try
                        d = CalculateRoof()
                        counterLimit = counterLimit + 1
                    End While
                    gapFold.Driven = False
                    b = True
                Catch ex As Exception
                    Try
                        gapFold.Delete()
                    Catch ex3 As Exception
                    End Try
                    gapFold = sk3D.DimensionConstraints3D.AddLineLength(cl3)
                    b = True
                End Try
            End If
            ac.Delete()
            Try
                gc.Delete()
            Catch ex As Exception

            End Try



            Return dc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function AdjustPositiveAngle(ac As DimensionConstraint3D, limit As Double) As DimensionConstraint3D
        Dim d As Double = CalculateRoof()
        For i = 1 To 32
            gapFold.Driven = True
            If d > 0 Then

                If ac.Parameter._Value > Math.PI / 2 Then
                    adjuster.AdjustDimConstrain3DSmothly(ac, Math.PI - limit)
                Else
                    adjuster.AdjustDimConstrain3DSmothly(ac, limit)
                End If
            Else

                If ac.Parameter._Value > Math.PI / 2 Then
                    adjuster.AdjustDimConstrain3DSmothly(ac, Math.PI)
                Else
                    adjuster.AdjustDimConstrain3DSmothly(ac, 1 / 128)
                End If
                ac.Driven = True
                adjuster.AdjustDimConstrain3DSmothly(gapFold, gapFold.Parameter._Value * 9 / 8)
                gapFold.Driven = True
                i = 1
            End If
            If adjuster.IsLastAngleOk(ac, limit) Then
                Exit For
            Else
                d = CalculateRoof()
            End If
        Next

        ac.Driven = True
        Return ac
    End Function
    Function IsLastAngleOk(dc As DimensionConstraint3D, d As Double) As Boolean
        Dim e As Double
        If dc.Parameter._Value > Math.PI / 2 Then
            e = (dc.Parameter._Value - (Math.PI - d)) / Math.PI
        Else
            e = (dc.Parameter._Value - d) / Math.PI
        End If
        Return Math.Abs(e) < 1 / 64

    End Function
    Function CorrectLastAngle(aci As DimensionConstraint3D) As DimensionConstraint3D
        Dim bl6, bl4, l, bl2 As SketchLine3D
        Dim dc, ac As DimensionConstraint3D
        Dim b As Double = GetParameter("b")._Value
        Dim currentQ As Integer = GetParameter("currentQ")._Value
        Dim oddQ As Integer
        Dim d As Double
        Try
            bl4 = bandLines(4)
            bl2 = bandLines(2)
            bl6 = bandLines(6)
            aci.Driven = True
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl4.EndPoint, bl6.EndPoint)
            dc.Driven = True
            l = sk3D.SketchLines3D.AddByTwoPoints(bl4.StartPoint, bl6.EndPoint)
            l.Construction = True
            acCl7Bl4 = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
            ' dc = RecoverGapFold(dc)
            '  adjuster.AdjustDimConstrain3DSmothly(acCl7Bl4, acCl7Bl4.Parameter._Value / 4)
            Math.DivRem(currentQ, 2, oddQ)
            If oddQ = 0 Then
                adjuster.AdjustDimConstrain3DSmothly(acCl7Bl4, acCl7Bl4.Parameter._Value * 3 / 2)
            Else
                adjuster.AdjustDimConstrain3DSmothly(acCl7Bl4, acCl7Bl4.Parameter._Value / 2)
            End If
            dimConstrainBandLine2.Driven = False
            If secondLine.Length < b Then
                adjuster.AdjustDimensionConstraint3DSmothly(dimConstrainBandLine2, 3 * b / 2)
            Else
                adjuster.AdjustDimensionConstraint3DSmothly(dimConstrainBandLine2, dimConstrainBandLine2.Parameter._Value * 4 / 3)
            End If
            dimConstrainBandLine2.Driven = True
            If gapFold.Parameter._Value > gapFoldCM * 2 Then
                'aci.Driven = False
                For i = 1 To 64
                    ' gapFold.Driven = True
                    d = (gapFold.Parameter._Value * (31 + (1 / (1 + Math.Exp(-(-dc.Parameter._Value + 2 * gapFoldCM) / gapFoldCM)))) / 32)
                    If adjuster.AdjustDimensionConstraint3DSmothly(gapFold, d) Then
                        '  aci.Driven = True
                        sk3D.Solve()
                    Else
                        ' aci.Driven = True
                        For j = 1 To 8
                            comando.UndoCommand()
                            comando.UndoCommand()
                            gapFold.Driven = False
                            sk3D.Solve()
                            If monitor.IsSketch3DHealthy(sk3D) Then

                                gapFold.Parameter._Value *= 32 / 31
                                Try
                                    sk3D.Solve()
                                Catch ex As Exception
                                    comando.UndoCommand()
                                    comando.UndoCommand()
                                    comando.UndoCommand()
                                    sk3D.Solve()
                                    gapFold.Driven = True
                                    adjuster.RecoveryUnhealthySketch(sk3D)

                                End Try

                                If monitor.IsSketch3DHealthy(sk3D) Then
                                    dc = RecoverGapFold(dc)
                                    dc.Driven = True
                                    Exit For
                                Else

                                    comando.UndoCommand()
                                    comando.UndoCommand()
                                    doku.Update()
                                    adjuster.RecoveryUnhealthySketch(sk3D)
                                End If

                            Else

                                comando.UndoCommand()
                                comando.UndoCommand()
                                sk3D.Solve()
                                adjuster.RecoveryUnhealthySketch(sk3D)
                            End If
                        Next


                    End If

                    If (gapFold.Parameter._Value < 2 * gapFoldCM) Or (dc.Parameter._Value < 3 * gapFoldCM / 4) Then
                        Exit For
                    End If

                Next
            End If
            gapFold.Driven = False
            sk3D.Solve()
            Try
                dc.Driven = False
            Catch ex As Exception
                dc.Driven = True
            End Try

            dimConstrainBandLine2.Driven = True
            gapFold.Driven = True
            sk3D.Solve()

            TryVerticalEntrance(dc, l)


            aci.Driven = False
            Return gapFold
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function TryVerticalEntrance(aci As DimensionConstraint3D, skli As SketchLine3D) As DimensionConstraint3D
        Dim dc As DimensionConstraint3D
        Dim bl6 As SketchLine3D = bandLines(6)
        Dim b As Double = GetParameter("b")._Value
        Dim currentQ As Integer = GetParameter("currentQ")._Value
        Dim oddQ As Integer
        Try
            aci.Driven = True
            acCl7Bl4.Driven = True
            adjuster.AdjustDimConstrain3DSmothly(gapFold, 2 * gapFoldCM)
            Math.DivRem(currentQ, 2, oddQ)
            If oddQ = 0 Then

                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(firstLine, skli)
                For index = 1 To 32
                    If aci.Parameter._Value > gapFoldCM / 2 Then
                        adjuster.AdjustDimConstrain3DSmothly(dc, dc.Parameter._Value * 3 / 4)
                    Else
                        Exit For
                    End If
                Next
                dc.Delete()

            Else
                acCl7Bl4.Driven = False
                For index = 1 To 8
                    adjuster.AdjustDimConstrain3DSmothly(dimConstrainBandLine2, 7 * dimConstrainBandLine2.Parameter._Value / 8)
                Next

                adjuster.AdjustDimConstrain3DSmothly(dimConstrainBandLine2, 3 * b / 2)
                For index = 1 To 16
                    If aci.Parameter._Value > gapFoldCM / 2 Then
                        adjuster.AdjustDimConstrain3DSmothly(acCl7Bl4, acCl7Bl4.Parameter._Value * 7 / 8)
                    Else
                        Exit For
                    End If
                Next
            End If
            acCl7Bl4.Driven = True
            Return acCl7Bl4

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function RecoverGapFold(dci As DimensionConstraint3D) As DimensionConstraint3D
        Try
            gapFold.Driven = True
            For i = 1 To 32
                If adjuster.AdjustDimensionConstraint3DSmothly(dci, dci.Parameter._Value * 15 / 16) Then
                Else
                    Exit For
                End If
                If (gapFold.Parameter._Value < 2 * gapFoldCM) Or (dci.Parameter._Value < 5 * gapFoldCM / 2) Then
                    Exit For
                End If
            Next

            Return dci
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CalculateRoof() As Double
        Dim fourLine, sixthLine, cl3, cl2 As SketchLine3D

        Dim dir As Integer = GetParameter("direction")._Value
        Dim v1, v2, v3, v4, vfl, vnwf, vflnwf, vcp As Vector
        Dim d, e As Double
        Dim currentQ As Integer = GetParameter("currentQ")._Value
        Dim oddQ As Integer

        Try
            fourLine = bandLines.Item(4)
            sixthLine = bandLines.Item(6)
            cl3 = constructionLines.Item(3)
            cl2 = constructionLines.Item(2)
            v3 = firstLine.StartSketchPoint.Geometry.VectorTo(sixthLine.EndSketchPoint.Geometry)
            v2 = fourLine.Geometry.Direction.AsVector

            v4 = cl3.Geometry.Direction.AsVector


            vfl = firstLine.Geometry.Direction.AsVector
            vnwf = cl2.Geometry.Direction.AsVector
            vcp = compDef.WorkPoints.Item(1).Point.VectorTo(firstLine.EndSketchPoint.Geometry)
            vflnwf = vfl.CrossProduct(vnwf)
            e = vcp.DotProduct(vflnwf)
            v1 = v2.CrossProduct(v3)
            d = v1.DotProduct(v4)
            If dir < 0 Then
                d = -1 * d
            End If
            Math.DivRem(currentQ, 2, oddQ)
            If oddQ = 0 Then
                d = -1 * d
            End If


            Return d
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function

    Function IsThereCoincidentConstrains(line As SketchLine3D) As Boolean
        For Each o As Object In refLine.Constraints3D
            If o.Type = ObjectTypeEnum.kCoincidentConstraint3DObject Then
                Return True
            End If
        Next
        Return False
    End Function
    Function GetFirstLine() As SketchLine3D
        Dim minLength As Double = 99999999999999
        Dim cc1, cc2 As CoincidentConstraint3D
        Dim v1, v2 As Vector

        Dim done As Boolean = False
        For Each cline As SketchLine3D In refSk.SketchLines3D
            If cline.Construction Then
                cc1 = FindCommonCoincidentConstrainCurve(cline)
                v1 = CreateVectorLineCoincident(cc1, cline)
                For Each line As SketchLine3D In refSk.SketchLines3D
                    If Not line.Equals(cline) Then
                        cc2 = FindCommonCoincidentConstrainCurve(line)

                        For Each cc As CoincidentConstraint3D In line.Constraints3D.OfType(Of CoincidentConstraint3D)
                            If AreCoincidentConstraintsClose(cc2, cc) Then
                                v2 = CreateVectorLineCoincident(cc2, line)
                                If v1.DotProduct(v2) < minLength Then
                                    minLength = v1.DotProduct(v2)
                                    firstLine = line
                                End If
                            End If



                        Next
                    End If
                Next
            End If
        Next
        Return firstLine
    End Function

    Function AreCoincidentConstraintsClose(cc1 As CoincidentConstraint3D, cc2 As CoincidentConstraint3D) As Boolean
        If cc1.SketchPoint.Geometry.IsEqualTo(cc2.SketchPoint.Geometry, 1 / 1000) Then
            Return True
        End If
        Return False
    End Function
    Function FindCommonCoincidentConstrainCurve(line As SketchLine3D) As CoincidentConstraint3D
        Try
            Dim o As Object
            o = line.Constraints3D.Count
            For Each gc As GeometricConstraint3D In line.Constraints3D
                If gc.Type = ObjectTypeEnum.kCoincidentConstraint3DObject Then
                    For Each gctc As CoincidentConstraint3D In refCurve.Constraints3D
                        If gctc.Type = ObjectTypeEnum.kCoincidentConstraint3DObject Then
                            If gc.Equals(gctc) Then
                                Return gc
                            End If
                        End If
                    Next

                End If


            Next
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try



        Return Nothing
    End Function

    Function CreateVectorLineCoincident(cc As CoincidentConstraint3D, line As SketchLine3D) As Vector
        Dim v1 As Vector
        Dim p1, p2 As Point
        If cc.Vertex.Point.IsEqualTo(line.StartPoint) Then
            p1 = line.StartPoint
            p2 = line.EndPoint
            v1 = p1.VectorTo(p2)
        Else
            p1 = line.EndPoint
            p2 = line.StartPoint
            v1 = p1.VectorTo(p2)
        End If
        Return v1
    End Function
    Function IsLengthSimilar(line As SketchLine3D, value As Double) As Boolean
        Dim tol As Double = 1
        Return IsLengthSimilar(line, value, tol)
        ' Return (line.Length * 10 > value * (1 - tol / 100) And line.Length * 10 < value * (1 + tol / 100))
    End Function
    Function IsLengthSimilar(line As SketchLine3D, value As Double, tol As Double) As Boolean

        Return Math.Abs((line.Length - value) / line.Length) > tol / 100
    End Function
    Public Function DrawKeyLine(line As SketchLine3D) As Sketch3D
        Dim sketch As Sketch3D
        sketch = doku.ComponentDefinition.Sketches3D.Add()
        sketch.Name = "MainSkecht"
        sketch.SketchLines3D.AddByTwoPoints(line.StartPoint, line.EndPoint)
        sketch.SketchLines3D.Item(sketch.SketchLines3D.Count).Construction = True

        Return sketch
    End Function
    Function IsSketchHealthy() As Boolean
        Dim b As Boolean
        If sk3D.HealthStatus = HealthStatusEnum.kOutOfDateHealth Or
         sk3D.HealthStatus = HealthStatusEnum.kUpToDateHealth Then
            If monitor.AreDimensionsHealthy(sk3D) Then
                If monitor.AreConstrainsHealthy(sk3D) Then
                    b = True
                Else
                    MsgBox("not healthy constrain")
                    b = True
                End If
            Else
                b = False

            End If
        Else
            b = False
        End If
        Return b
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
    Public Function IsDocHealthy() As Boolean
        Try
            If monitor.PartHasProblems(doku) Then
                healthy = False

            Else
                If doku._SickNodesCount > 0 Then
                    healthy = False


                Else
                    healthy = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
            healthy = False
        End Try


        Return healthy
    End Function

End Class
