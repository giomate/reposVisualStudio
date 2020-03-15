Imports Inventor
Imports System
Imports GetInitialConditions
Imports DrawInitialSketch

Public Class OriginSketch
    Dim doku As PartDocument
    Dim app As Application
    Public sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Public refLine, firstLine, secondLine, thirdLine, lastLine, twistLine, inputLine, doblezline, centroLine, gapFoldLine As SketchLine3D
    Public curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim curve3D As Curves3D
    Dim monitor As DesignMonitoring
    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3, centroPoint As Point
    Dim tg As TransientGeometry
    Dim gapFoldCM As Double
    Dim adjuster As SketchAdjust
    Public bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim nombrador As Nombres
    Dim metro, gapFold As DimensionConstraint3D
    Public ID1, ID2 As Integer




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
        refCurve = refDoc.oDoc.ComponentDefinition.Sketches3D.Item("refCurve").SketchEquationCurves3D.Item(1)
        refLine = refDoc.GetKeyLine()

        doku.Activate()

        DrawTrobinaCurve(q)
        DrawInitialLine(refLine)
        comando.TopRightView()
        If sk3D.SketchLines3D.Count > 0 Then

            DrawLines()
            If AdjustlastAngle() Then
                point1 = firstLine.StartSketchPoint.Geometry
                point2 = firstLine.EndSketchPoint.Geometry
                point3 = secondLine.EndSketchPoint.Geometry
                done = 1
            End If
        End If
        Return sk3D
    End Function
    Public Function DrawNextMainSketch(rl As SketchLine3D, tl As SketchLine3D, fl As SketchLine3D, gfl As SketchLine3D) As SketchLine3D

        twistLine = tl
        doblezline = fl
        sk3D = doku.ComponentDefinition.Sketches3D.Item(doku.ComponentDefinition.Sketches3D.Count)
        refLine = rl
        gapFoldLine = gfl
        refLine.Construction = True
        curve = sk3D.SketchEquationCurves3D.Item(sk3D.SketchEquationCurves3D.Count)
        If refLine.Length > 0 Then
            If DrawSingleFloatingLines() Then
                point1 = firstLine.StartSketchPoint.Geometry
                point2 = firstLine.EndSketchPoint.Geometry
                point3 = secondLine.EndSketchPoint.Geometry
                For Each sl As SketchLine3D In sk3D.SketchLines3D
                    If sl.Equals(firstLine) Then
                        ID1 = sl.AssociativeID

                    Else

                        If sl.Equals(secondLine) Then
                            ID1 = sl.AssociativeID

                        End If
                    End If
                Next
                inputLine = firstLine
                done = 1
            End If
        End If

        Return firstLine
    End Function
    Function DrawReferenceCurve(sk As Sketch3D) As SketchEquationCurve3D

        Return Nothing
    End Function
    Function DrawLines() As Sketch3D
        Dim i As Integer = 0
        While Not done
            DrawSingleLines(i)
            i = i + 1
        End While
        done = 0
        Return sk3D
    End Function
    Sub DrawSingleLines(i As Integer)

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
            If DrawCentroLine().Length > 0 Then
                If DrawFirstFloatingLine().Length > 0 Then
                    If DrawSecondLine().Length > 0 Then
                        If DrawFirstConstructionLine().Construction Then
                            If DrawThirdLine().Length > 0 Then
                                If DrawSecondConstructionLine().Construction Then
                                    If DrawFourthFloatingLine().Length > 0 Then
                                        If DrawFifthLine().Length > 0 Then
                                            If DrawThirdConstructionLine().Construction Then
                                                If DrawFourthConstructionLine().Construction Then
                                                    Return True
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
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return False
    End Function

    Function DrawCentroLine() As SketchLine3D

        Dim l As SketchLine3D
        Dim sp As SketchPoint3D
        centroPoint = tg.CreatePoint(0, 0, 0)
        sp = sk3D.SketchPoints3D.Add(centroPoint)
        sk3D.GeometricConstraints3D.AddGround(sp)
        l = sk3D.SketchLines3D.AddByTwoPoints(refLine.StartPoint, sp, False)
        l.Construction = True

        ' endPoint.TranslateBy(v)s Vector
        'v = firstLine.Geometry.Direction.AsVector
        'v.ScaleBy(GetParameter("b")._Value / 1)
        'endPoint = firstLine.StartSketchPoint.Geometry
        'endPoint.TranslateBy(v)
        centroLine = l
        Return l
    End Function
    Function DrawInitialLine(line As SketchLine3D) As SketchLine3D

        sk3D.SketchLines3D.AddByTwoPoints(line.StartSketchPoint.Geometry, line.EndSketchPoint.Geometry)
        sk3D.SketchLines3D.Item(sk3D.SketchLines3D.Count).Construction = True
        Return line
    End Function
    Function DrawTrobinaCurve(q As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(q, "s1")
    End Function
    Function DrawTrobinaCurve(q As Integer, s As String) As SketchEquationCurve3D

        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = curve3D.DrawTrobinaCurve(sk3D, q)
        sk3D.GeometricConstraints3D.AddGround(curve)
        Return curve
    End Function
    Function DrawFirstLine() As SketchLine3D
        Try

            Dim l As SketchLine3D = Nothing

            If IsThereCoincidentConstrains(refLine) Then
                Dim p1 As Point = FindCommonCoincidentConstrainCurve(refLine).SketchPoint.Geometry
            Else
                l = sk3D.SketchLines3D.AddByTwoPoints(refLine.StartSketchPoint.Geometry, refLine.EndSketchPoint.Geometry)
                sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            End If
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
    Function DrawFirstFloatingLine() As SketchLine3D
        Try

            Dim l As SketchLine3D
            ' Dim dc As DimensionConstraint3D
            Dim gc As GeometricConstraint3D
            Dim v1, v2, v3 As Vector
            v1 = doblezline.Geometry.Direction.AsVector
            v2 = twistLine.Geometry.Direction.AsVector
            v3 = v1.CrossProduct(v2)
            v3.AsUnitVector.AsVector.ScaleBy(GetParameter("b")._Value)
            Dim pt As Point = refLine.StartSketchPoint.Geometry
            pt.TranslateBy(v3)
            l = sk3D.SketchLines3D.AddByTwoPoints(refLine.StartPoint, pt, False)
            'dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            ' dc.Parameter._Value = curve3D.DP.b / 10
            doku.Update2(True)
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, twistLine)
            doku.Update2(True)
            gc.Delete()
            doku.Update2(True)
            gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, centroLine)
            doku.Update2(True)
            gc.Delete()
            'metro = dc
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
            v1 = lastLine.StartSketchPoint.Geometry.VectorTo(lastLine.EndSketchPoint.Geometry)
            v3 = twistLine.StartSketchPoint.Geometry.VectorTo(twistLine.EndSketchPoint.Geometry)
            Dim l As SketchLine3D = Nothing
            Dim d As Double
            Dim dc As DimensionConstraint3D
            Dim optpoint As Point = Nothing
            Dim p As Plane
            p = tg.CreatePlane(lastLine.EndSketchPoint.Geometry, v1)
            Dim minDis As Double = 9999999999
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
                'l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, o, False)
                v2 = lastLine.EndSketchPoint.Geometry.VectorTo(o)
                d = v2.DotProduct(v3)
                If d > 0 Then
                    If d < minDis Then
                        minDis = d
                        optpoint = o
                    End If
                End If
            Next
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, optpoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            Try
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                If dc.Parameter._Value > GetParameter("b")._Value * 2 Then
                    Try
                        adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value)
                    Catch ex As Exception
                        dc.Driven = True
                    End Try
                End If
                dc.Driven = True
            Catch ex As Exception
                dc.Driven = True
            End Try
            'metro.Driven = True
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
            Dim dc As DimensionConstraint3D

            'Dim v As Vector
            'v = firstLine.Geometry.Direction.AsVector
            'v.ScaleBy(GetParameter("b")._Value / 1)
            'endPoint = firstLine.StartSketchPoint.Geometry
            'endPoint.TranslateBy(v)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.StartPoint, firstLine.EndSketchPoint.Geometry, False)
            ' endPoint.TranslateBy(v)

            sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, secondLine)
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value) Then
                dc.Parameter._Value = GetParameter("b")._Value

            Else
                dc.Driven = True
            End If
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
            Dim bl4, cl1 As SketchLine3D
            Dim dc, ac As DimensionConstraint3D
            Dim gc As GeometricConstraint3D

            bl4 = bandLines.Item(4)
            cl1 = constructionLines.Item(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndPoint, bl4.Geometry.MidPoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, bl4)
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
            ac.Driven = True
            Try
                sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
            Catch ex As Exception
                ac.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
            End Try

            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            dc.Driven = True
            Try
                gc = sk3D.GeometricConstraints3D.AddEqual(l, cl1)
            Catch ex As Exception
                dc.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value)
                dc.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
            End Try

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
            Dim c As Integer
            Dim l, fl As SketchLine3D
            Dim bl4, cl3 As SketchLine3D
            Dim dc As TwoPointDistanceDimConstraint3D
            Dim dcl, dc2, dcbl1, dcbl4, ac As DimensionConstraint3D
            Dim gc As GeometricConstraint3D
            bl4 = bandLines.Item(4)
            cl3 = constructionLines.Item(3)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndSketchPoint.Geometry, bl4.StartSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, bl4)
            fl = bandLines.Item(4)
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.EndPoint, cl3.EndPoint)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value / 2)
            adjuster.AdjustTwoPointsSmothly(dc, 1 / 10)
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
            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
            Catch ex As Exception
                Try
                    dc2.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                    dc2.Driven = True
                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
                    adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                    ac.Driven = True
                    sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
                Catch ex3 As Exception

                End Try

            End Try

            Try
                dc2.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                dc2.Driven = True
                sk3D.GeometricConstraints3D.AddEqual(l, cl3)
            Catch ex As Exception
                Try
                    dc2.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                    dc2.Driven = True
                    sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                Catch ex1 As Exception
                    Try

                        gc = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                    Catch ex3 As Exception
                        dc2.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                        dc2.Driven = True
                        Try
                            gc = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                        Catch ex4 As Exception
                            dc2.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                            dc2.Driven = True
                            Try
                                gc = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                            Catch ex5 As Exception
                                If adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value) Then
                                    dc2.Parameter._Value = GetParameter("b")._Value
                                End If
                                Try
                                    gc = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                                Catch ex8 As Exception
                                    Try
                                        dcbl1 = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                                        dcbl4 = sk3D.DimensionConstraints3D.AddLineLength(fl)
                                        Try
                                            While Math.Abs(fl.Length - 2 * firstLine.Length) > gapFoldCM And c < 128
                                                adjuster.AdjustDimensionConstraint3DSmothly(dcbl1, fl.Length / 2)
                                                c = c + 1
                                            End While
                                        Catch ex9 As Exception
                                            dcbl1.Driven = True
                                            dcbl4.Driven = True
                                        End Try

                                        dcbl4.Delete()

                                        dcbl1.Delete()
                                        gc = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                                    Catch ex2 As Exception
                                        dcbl1.Driven = True
                                        dcbl4.Driven = True
                                    End Try
                                End Try
                            End Try

                        End Try

                    End Try
                End Try


            End Try
            Try
                gc = sk3D.GeometricConstraints3D.AddCoincident(fl.EndPoint, curve)
            Catch ex As Exception
                dcl = sk3D.DimensionConstraints3D.AddLineLength(fl)
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 3)
                dcl.Delete()
                Try
                    gc = sk3D.GeometricConstraints3D.AddCoincident(fl.EndPoint, curve)
                Catch ex4 As Exception
                    Dim v, v2 As Vector
                    Dim bl5 As SketchLine3D = bandLines.Item(5)
                    v = bl5.Geometry.Direction.AsVector
                    Dim pl As Plane
                    pl = tg.CreatePlane(bl5.StartSketchPoint.Geometry, v)
                    Dim opt As Point
                    Dim dMin As Double = 99999999
                    For Each o As Point In pl.IntersectWithCurve(curve.Geometry)
                        If o.DistanceTo(bl5.StartSketchPoint.Geometry) < dMin Then
                            dMin = o.DistanceTo(bl5.StartSketchPoint.Geometry)
                            opt = o
                        End If
                    Next
                    Dim skpt As SketchPoint3D
                    skpt = sk3D.SketchPoints3D.Add(opt)
                    Dim dcwo1 As DimensionConstraint3D
                    dcwo1 = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl5.StartPoint, skpt)
                    adjuster.AdjustDimensionConstraint3DSmothly(dcwo1, dcwo1.Parameter._Value / 4)
                    dcwo1.Driven = True
                    gc = sk3D.GeometricConstraints3D.AddCoincident(fl.EndPoint, curve)

                End Try

            End Try


            l.Construction = True
            constructionLines.Add(l)
            lastLine = l
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
            'pl = sk3D.SketchLines3D.Item(sk3D.SketchLines3D.Count - 1)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndPoint, firstLine.StartPoint, False)
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
            Dim v1, v2, v3 As Vector
            Dim l, pl As SketchLine3D
            Dim endPoint As Point
            Dim dc As DimensionConstraint3D
            v1 = firstLine.EndSketchPoint.Geometry.VectorTo(firstLine.StartSketchPoint.Geometry)
            v2 = secondLine.EndSketchPoint.Geometry.VectorTo(secondLine.StartSketchPoint.Geometry)
            v3 = gapFoldLine.Geometry.Direction.AsVector
            v3.ScaleBy(-1 * gapFoldCM)
            If sk3D.Name = "s9" Then
                ' v3.ScaleBy(-1)
            End If
            endPoint = secondLine.StartSketchPoint.Geometry
            endPoint.TranslateBy(v3)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, endPoint, False)
            endPoint.TranslateBy(v3)
            gapFold = sk3D.DimensionConstraints3D.AddLineLength(l)

            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, firstLine)
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

            pl = firstLine
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.StartPoint, lastLine.EndSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
            sk3D.GeometricConstraints3D.AddCoincident(lastLine.EndPoint, l)
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
            Dim ac As DimensionConstraint3D
            pl = thirdLine
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, pl.StartPoint, False)
            Try
                gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, pl)
            Catch ex As Exception
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, pl)
                Try
                    adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                Catch ex2 As Exception
                    ac.Driven = True
                End Try
                ac.Driven = True
                gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, pl)
            End Try
            doku.Update2(True)
            gc.Delete()
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
            Dim l, pl, ml As SketchLine3D
            Dim v As Vector
            Dim p As Plane
            Dim minDis, d As Double
            Dim optpoint As Point
            Dim dc As DimensionConstraint3D
            dc = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
            If dc.Parameter._Value > curve3D.DP.b * 2 / 10 Then
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 4 / 5) Then
                    dc.Driven = True
                Else
                    dc.Delete()
                End If
            End If
            pl = constructionLines.Item(3)
            v = pl.Geometry.Direction.AsVector
            p = tg.CreatePlane(pl.EndSketchPoint.Geometry, v)

            Dim puntos As ObjectsEnumerator
            puntos = p.IntersectWithCurve(curve.Geometry)
            minDis = 9999999999
            Dim o2 As Point = puntos.Item(puntos.Count)
            optpoint = o2
            ml = bandLines.Item(4)
            Dim vc, vmjl As Vector
            vmjl = thirdLine.Geometry.Direction.AsVector
            vmjl.ScaleBy(-1)

            For Each o As Point In puntos
                ' l = sk3D.SketchLines3D.AddByTwoPoints(pl.EndSketchPoint.Geometry, o, False)
                vc = pl.EndSketchPoint.Geometry.VectorTo(o)
                d = vc.CrossProduct(vmjl).Length * vc.DotProduct(vmjl)
                If d > 0 Then
                    If o.DistanceTo(pl.EndSketchPoint.Geometry) < minDis Then
                        minDis = o.DistanceTo(pl.EndSketchPoint.Geometry)

                        optpoint = o
                    End If
                End If


            Next

            vp = optpoint

            pl = bandLines.Item(4)
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.EndSketchPoint.Geometry, vp, False)
            sk3D.GeometricConstraints3D.AddCoincident(pl.EndPoint, l)
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
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 12 / 11) Then
                dc.Delete()
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 10) Then
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

    Function AdjustlastAngle() As Boolean
        Try
            Dim fourLine, sixthLine, cl2 As SketchLine3D
            Dim dc As TwoLineAngleDimConstraint3D
            Dim limit As Double = 0.1

            Dim b As Boolean = False

            fourLine = bandLines.Item(4)
            sixthLine = bandLines.Item(6)
            cl2 = constructionLines.Item(2)
            If adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value * 3) Then
                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)

                If adjuster.AdjustGapSmothly(gapFold, gapFoldCM, dc) Then
                    b = True
                Else
                    dc.Driven = True
                    gapFold.Delete()
                    gapFold = sk3D.DimensionConstraints3D.AddLineLength(cl2)
                    b = True
                End If
            Else
                gapFold.Driven = True
                b = True
            End If

            Return b
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
