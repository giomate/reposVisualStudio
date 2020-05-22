Imports Inventor
Imports System
Imports GetInitialConditions
Imports DrawInitialSketch

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
    Dim dimConstrainBandLine2, gapFold, dcThirdLine As DimensionConstraint3D
    Dim startPoint As GeometricConstraint3D
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim compDef As SheetMetalComponentDefinition
    Public ID1, ID2, qvalue As Integer
    Dim k1(), k2() As Byte
    Dim backwards As Boolean




    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)

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

    Public Function StartDrawing(ref As SketchLine3D) As Sketch3D
        refLine = ref
        wp1 = ref.StartSketchPoint.Geometry
        wp3 = ref.EndSketchPoint.Geometry
        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = "s1"
        curve = curve3D.DrawTrobinaCurve(sk3D)

        DrawInitialLine(refLine)
        If sk3D.SketchLines3D.Count > 0 Then

            'skRef = Form1.keyline.OpenMainSketch()
            DrawLines()

        End If

        Return sk3D
    End Function
    Public Function StartDrawingTranslated(refDoc As FindReferenceLine, q As Integer) As Sketch3D

        refSk = refDoc.OpenMainSketch(refDoc.oDoc)
        ' refCurve = refDoc.oDoc.ComponentDefinition.Sketches3D.Item("refCurve").SketchEquationCurves3D.Item(1)
        refLine = refDoc.GetKeyLine()
        qvalue = q
        doku.Activate()

        DrawTrobinaCurve(q)
        DrawInitialLine(refLine)
        comando.TopRightView()
        compDef = doku.ComponentDefinition
        sheetMetalFeatures = compDef.Features
        If sk3D.SketchLines3D.Count > 0 Then

            DrawLines()
            If AdjustlastAngle().Parameter._Value > gapFoldCM * 4 Then
                point1 = firstLine.StartSketchPoint.Geometry
                point2 = firstLine.EndSketchPoint.Geometry
                point3 = secondLine.EndSketchPoint.Geometry
                done = 1
            End If
        End If
        Return sk3D
    End Function
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
    Public Function DrawMainSketch(refDoc As FindReferenceLine, q As Integer) As Sketch3D

        Dim s As String
        Dim ac As DimensionConstraint3D

        refLine = refDoc.GetKeyLine()
        kanteLine = refDoc.GetKanteLine()
        doku.Activate()
        compDef = doku.ComponentDefinition
        sheetMetalFeatures = compDef.Features
        curve = DrawTrobinaCurveFitted(q)
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
                If gapFold.Parameter._Value > gapFoldCM * 4 Then
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
    Function DrawLines() As Sketch3D
        Dim i As Integer = 0
        While Not done
            DrawSingleOldLines(i)
            i = i + 1
        End While
        done = 0
        Return sk3D
    End Function
    Sub DrawSingleOldLines(i As Integer)

        DrawFirstLine()

        DrawSecondLine()

        DrawFirstConstructionLine()

        DrawThirdLine()

        DrawSecondConstructionLine()

        DrawFourthLine()

        DrawFifthLine()

        DrawThirdConstructionLine()
        DrawFourthConstructionLine()

        DrawSixthLine()

        DrawSeventhLine()





    End Sub
    Function DrawSingleFloatingLines() As Boolean
        Try
            If DrawFirstFloatingLine().Length > 0 Then
                If DrawSecondLine().Length > 0 Then
                    If DrawFirstConstructionLine().Construction Then
                        If DrawThirdLine().Length > 0 Then
                            If DrawSecondConstructionLine().Construction Then
                                If DrawFourthFloatingLine().Length > 0 Then
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
    Function DrawTrobinaCurveFitted(q As Integer, s As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(q, "s1", s)
    End Function
    Function DrawTrobinaCurveFitted(q As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(q, "s1", 0)
    End Function
    Function DrawTrobinaCurve(q As Integer, s As String) As SketchEquationCurve3D

        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = curve3D.DrawTrobinaCurve(sk3D, q)
        sk3D.GeometricConstraints3D.AddGround(curve)
        Return curve
    End Function
    Function DrawTrobinaCurve(q As Integer, s As String, f As Integer) As SketchEquationCurve3D

        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = curve3D.DrawTrobinaCurve(sk3D, q, f)
        sk3D.GeometricConstraints3D.AddGround(curve)
        Return curve
    End Function
    Function DrawFirstLine() As SketchLine3D
        Dim r1, r2 As Double
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
            'rl = sk3D.Include(refLine)
            l = sk3D.SketchLines3D.AddByTwoPoints(pt1, pt2)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            gc = sk3D.GeometricConstraints3D.AddGround(l.StartPoint)

            Try
                '  gpc = sk3D.GeometricConstraints3D.AddParallel(l, refLine)
            Catch ex As Exception

            End Try






            point1 = l.StartSketchPoint.Geometry
            point2 = l.EndSketchPoint.Geometry
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
            doku.Update2(True)
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, twistLine)
            doku.Update2(True)
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
            Dim l As SketchLine3D = Nothing
            Dim optpoint As Point = curve.EndSketchPoint.Geometry
            Dim dc, ac, dcfl As DimensionConstraint3D
            Dim d As Double
            v1 = lastLine.StartSketchPoint.Geometry.VectorTo(lastLine.EndSketchPoint.Geometry)
            v3 = doku.ComponentDefinition.WorkPoints.Item(1).Point.VectorTo(firstLine.StartSketchPoint.Geometry)
            p = tg.CreatePlane(lastLine.EndSketchPoint.Geometry, v1)
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
            If Math.Abs(firstLine.Length - GetParameter("b")._Value) < 1 / GetParameter("b")._Value Then
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
                adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value * 4 / 3)
                dc.Delete()
                ac.Delete()
                dcfl.Delete()

            End If
            dimConstrainBandLine2 = sk3D.DimensionConstraints3D.AddLineLength(l)
            dimConstrainBandLine2.Driven = True


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
            Dim otherLine, cl As SketchLine3D

            otherLine = bandLines.Item(4)
            cl = constructionLines.Item(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndPoint, otherLine.Geometry.MidPoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, otherLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, otherLine)
            sk3D.GeometricConstraints3D.AddEqual(l, cl)
            l.Construction = True

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
            Dim bl4, cl As SketchLine3D
            Dim dc As TwoPointDistanceDimConstraint3D
            Dim dc2, dccl4, ac As DimensionConstraint3D
            Dim gcel As GeometricConstraint3D
            bl4 = bandLines.Item(4)
            cl = constructionLines.Item(3)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndSketchPoint.Geometry, bl4.StartSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, bl4)
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.EndPoint, cl.EndPoint)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value / 2)
            AdjustTwoPointsSmothly(dc, 1 / 10)
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
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Driven = True
            sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
            dc2.Driven = False
            adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
            dc2.Driven = True
            Try
                sk3D.GeometricConstraints3D.AddEqual(l, cl)
            Catch ex2 As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 4)
                gapFold.Driven = True
                dc2.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                dc2.Driven = True
                gapFold.Driven = False
                Try
                    gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl)

                Catch ex3 As Exception
                    adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 4)
                    gapFold.Driven = True
                    dc2.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                    dc2.Driven = True
                    gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl)
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
    Function DrawThirdLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing
            Dim pl As SketchLine3D = Nothing
            pl = sk3D.SketchLines3D.Item(sk3D.SketchLines3D.Count - 1)
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.EndPoint, firstLine.StartPoint, False)
            thirdLine = l
            bandLines.Add(l)

            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawSecondConstructionLine() As SketchLine3D
        Try
            Dim v1, v2, v3, vk As Vector
            Dim l, kl As SketchLine3D
            Dim endPoint As Point
            Dim dc As DimensionConstraint3D
            v1 = firstLine.EndSketchPoint.Geometry.VectorTo(firstLine.StartSketchPoint.Geometry)
            v2 = secondLine.EndSketchPoint.Geometry.VectorTo(secondLine.StartSketchPoint.Geometry)
            v3 = v1.CrossProduct(v2).AsUnitVector().AsVector()
            v3.ScaleBy(gapFoldCM)
            kl = sk3D.SketchLines3D.AddByTwoPoints(kanteLine.StartSketchPoint.Geometry, kanteLine.EndSketchPoint.Geometry, False)

            vk = kl.Geometry.Direction.AsVector
            vk.ScaleBy(-1)
            endPoint = secondLine.StartSketchPoint.Geometry
            endPoint.TranslateBy(vk)

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
                sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            Catch ex As Exception
                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, secondLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI / 2)
                dc.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            End Try

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
            Dim dcl, acl As DimensionConstraint3D

            pl = firstLine
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.StartPoint, lastLine.EndSketchPoint.Geometry, False)
            Try
                sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
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
            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
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
                doku.Update2(True)
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
            Dim dc As DimensionConstraint3D
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
            sk3D.GeometricConstraints3D.AddCoincident(bl4.EndPoint, l)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
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
            Dim limit As Double = Math.PI / 32
            Dim angleLimit As Double = 1.8
            Dim d As Double
            Dim counterLimit As Integer = 0
            Dim lastAngle As Double


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
                doku.Update2(True)

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
                    Try
                        '  adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value * 2)
                    Catch ex As Exception
                    End Try
                    '  adjuster.AdjustGapSmothly(gapFold, gapFoldCM, dc)
                    Try
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
                    While ((d < 0 Or dc.Parameter._Value > Math.PI - limit / 1) And counterLimit < 16)
                        Try
                            gapFold.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value * 17 / 16)
                            gapFold.Driven = True
                            doku.Update2()

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



            Return dc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function CorrectLastAngle(aci As DimensionConstraint3D) As DimensionConstraint3D
        Dim bl6, bl4 As SketchLine3D
        Dim dc As DimensionConstraint3D
        Try
            bl4 = bandLines(4)
            bl6 = bandLines(6)
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl4.EndPoint, bl6.EndPoint)
            dc.Driven = True
            aci.Driven = False
            dimConstrainBandLine2.Driven = False
            adjuster.AdjustDimensionConstraint3DSmothly(dimConstrainBandLine2, dimConstrainBandLine2.Parameter._Value * 4 / 3)
            dimConstrainBandLine2.Delete()

            If gapFold.Parameter._Value > gapFoldCM * 2 Then
                For i = 1 To 16
                    gapFold.Driven = True
                    aci.Driven = False
                    If adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value * 15 / 16) Then
                        aci.Driven = True
                        doku.Update2(True)
                    Else
                        aci.Driven = True
                        For j = 1 To 8
                            comando.UndoCommand()
                            comando.UndoCommand()
                            comando.UndoCommand()

                            doku.Update2(True)
                            If monitor.IsSketch3DHealthy(sk3D) Then
                                gapFold.Parameter._Value *= 16 / 15
                                doku.Update()
                                If monitor.IsSketch3DHealthy(sk3D) Then
                                    dc = RecoverGapFold(dc)
                                Else
                                    comando.UndoCommand()
                                    comando.UndoCommand()
                                    comando.UndoCommand()
                                    doku.Update()
                                    adjuster.RecoveryUnhealthySketch(sk3D)
                                End If

                            Else
                                comando.UndoCommand()
                                comando.UndoCommand()
                                comando.UndoCommand()
                                doku.Update()
                                adjuster.RecoveryUnhealthySketch(sk3D)
                            End If
                        Next


                    End If

                    If (gapFold.Parameter._Value < gapFoldCM) Or (dc.Parameter._Value < 3 * gapFoldCM / 4) Then
                        Exit For
                    End If

                Next
            End If

            gapFold.Driven = True
            doku.Update2(True)
            aci.Driven = False
            Return gapFold
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function RecoverGapFold(dci As DimensionConstraint3D) As DimensionConstraint3D
        Try
            gapFold.Driven = True
            For i = 1 To 16
                If adjuster.AdjustDimensionConstraint3DSmothly(dci, dci.Parameter._Value * 15 / 16) Then
                Else
                    Exit For
                End If
                If (gapFold.Parameter._Value < 2 * gapFoldCM) Or (dci.Parameter._Value < 1 * gapFoldCM / 1) Then
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


        Dim v1, v2, v3, v4, vfl, vnwf, vflnwf, vcp As Vector
        Dim d, e As Double

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

            Return d * e
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function AdjustTwoPointsSmothly(dc As TwoPointDistanceDimConstraint3D, v As Double) As Boolean

        Dim dName As String
        Dim b As Boolean = False

        dName = dc.Parameter.Name

        If doku.Update2() Then
            If adjuster.UpdateDocu(doku) Then
                If adjuster.AdjustDimensionSmothly(dName, v) Then
                    b = True
                End If
            End If
        End If



        Return b
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
