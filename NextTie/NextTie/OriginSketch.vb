Imports Inventor
Imports System
Imports Subina_Design_Helpers


Public Class OriginSketch
    Dim doku As PartDocument
    Dim app As Application
    Public sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Dim tensorFirstLine As SketchLine3D
    Dim compDef As SheetMetalComponentDefinition
    Public refLine, firstLine, secondLine, thirdLine, lastLine, twistLine, inputLine, doblezline, centroLine, gapFoldLine, tangentLine, zAxisLine As SketchLine3D
    Public curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim curve3D As Curves3D
    Dim monitor As DesignMonitoring
    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3, centroPoint As Point
    Dim intersectionPoint As SketchPoint3D
    Dim tg As TransientGeometry
    Dim gapFoldCM, radioCylinder, diff2, diff1 As Double
    Dim adjuster As SketchAdjust
    Public bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim nombrador As Nombres
    Public curve_constraint As GeometricConstraint3D
    Public metro, gapFold, dcThirdLine, angleGap, angleTangent, inletGap, outletGap As DimensionConstraint3D
    Public ID1, ID2 As Integer






    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        compDef = doku.ComponentDefinition
        curve3D = New Curves3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        gapFoldCM = 3 / 10
        radioCylinder = 20 / 10
        done = False
        diff1 = Math.PI / 2 - 1.48
        diff2 = Math.PI / 2 - 1.58
    End Sub
    Public Sub New(docu As Inventor.Document, tl As SketchLine3D)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        compDef = doku.ComponentDefinition
        curve3D = New Curves3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        gapFoldCM = 3 / 10
        twistLine = tl
        done = False
        diff1 = Math.PI / 2 - 1.64
        diff2 = Math.PI / 2 - 1.43
    End Sub



    Public Function DrawNextStartSketch(rl As SketchLine3D, tl As SketchLine3D, fl As SketchLine3D, zline As SketchLine3D, spt As SketchPoint3D, oGap As DimensionConstraint3D) As SketchLine3D

        Dim currentQ As Integer = GetParameter("currentQ")._Value

        Dim oddQ As Integer
        Math.DivRem(currentQ, 2, oddQ)
        intersectionPoint = spt
        twistLine = tl
        tangentLine = tl
        doblezline = fl
        zAxisLine = zline
        sk3D = doku.ComponentDefinition.Sketches3D.Item(doku.ComponentDefinition.Sketches3D.Count)
        refLine = rl
        outletGap = oGap

        refLine.Construction = True
        curve = sk3D.SketchEquationCurves3D.Item(sk3D.SketchEquationCurves3D.Count)
        If refLine.Length > 0 Then
            If DrawSingleFloatingLines() Then
                point1 = firstLine.StartSketchPoint.Geometry
                point2 = firstLine.EndSketchPoint.Geometry
                point3 = secondLine.EndSketchPoint.Geometry
                inputLine = firstLine
                done = 1
            End If
        End If
#If Entry_Mode = "STIFT" Then
        If oddQ = 0 Then
        Else
#If Entry_Type = "Unilateral" Then
#Else


            curve_constraint.Delete()
#End If
        End If

#End If
            Return firstLine
    End Function
    Function DrawReferenceCurve(sk As Sketch3D) As SketchEquationCurve3D

        Return Nothing
    End Function
    Function IsFirstLineInsideCylinder() As Boolean
        Dim l As Line = compDef.WorkAxes.Item(3).Line
        If l.DistanceTo(firstLine.StartSketchPoint.Geometry) < radioCylinder Then
            Return True
        Else
            Return False

        End If
        Return False
    End Function
    Function ForceFirstLineOutside() As Integer
        Dim hecho As Boolean

        PullUpLine(tensorFirstLine)
        Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(tensorFirstLine)

        hecho = adjuster.AdjustDimensionConstraint3DSmothly(dc, radioCylinder + outletGap.Parameter._Value / 2)
        dc.Delete()
        Return hecho
    End Function
    Function PullUpLine(li As SketchLine3D) As Integer
        Dim hecho As Boolean
        Dim dc As DimensionConstraint3D
        Dim l As SketchLine3D
        Dim limit As Integer = 0
        If li.StartSketchPoint.Geometry.Z < 0 Then

            gapFold.Driven = True
            l = sk3D.SketchLines3D.AddByTwoPoints(intersectionPoint, li.StartPoint, False)
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            Do
                hecho = adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 15 / 16)

                limit = limit + 1
            Loop Until (l.EndSketchPoint.Geometry.Z > 1 Or limit > 16)
            dc.Delete()
            l.Delete()
            Return hecho
        Else
            Return True
        End If
        Return hecho
    End Function


    Function DrawSingleFloatingLines() As Boolean
        Try
            'If DrawCentroLine().Length > 0 Then
            If DrawFirstLine().Length > 0 Then
                If DrawSecondLine().Length > 0 Then
                    If DrawFirstConstructionLine().Construction Then
                        If DrawThirdLine().Length > 0 Then
                            If DrawSecondConstructionLine().Construction Then
                                If DrawFourLine().Length > 0 Then
                                    If DrawFifthLine().Length > 0 Then
                                        If DrawThirdConstructionLine().Construction Then
                                            If DrawFourthConstructionLine().Construction Then
                                                If DrawFifthConstructionLine().Length > 0 Then
                                                    Return True
                                                End If

                                                If AnchorFourLine().Length > 0 Then
                                                    If DrawFifthConstructionLine().Length > 0 Then
                                                        Return True
                                                    End If

                                                End If
                                                If DrawSixthLine().Length > 0 Then
                                                    If DrawSeventhLine().Length > 0 Then
                                                        ' Return True
                                                        If DrawFifthConstructionLine().Length > 0 Then
                                                            If DrawTangentParallel().Length > 0 Then
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

            End If
            '  End If

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return False
    End Function
    Function TryPerpendicular(l As SketchLine3D, bl As SketchLine3D) As GeometricConstraint3D
        Dim ac As DimensionConstraint3D
        Dim gc As GeometricConstraint3D
        ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(bl, l)
        adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
        ac.Delete()
        gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, bl)
        Return gc
    End Function
    Function DrawFifthConstructionLine() As SketchLine3D
        Dim l, cl2, cl1, bl2, l2 As SketchLine3D
        Dim dc, dccl4, ac As DimensionConstraint3D
        Dim gc, gcpp As GeometricConstraint3D
        Try

            cl1 = constructionLines.Item(1)
            Dim pt As Point = cl1.Geometry.MidPoint
            Dim vcl1, vbl2, v As Vector
            Dim currentQ As Integer = GetParameter("currentQ")._Value
            Dim dir As Integer = GetParameter("direction")._Value
            Dim oddQ As Integer
            FitEntryGap()

            CorrectSecondLine()
            vcl1 = cl1.Geometry.Direction.AsVector
            vbl2 = secondLine.Geometry.Direction.AsVector
            v = vcl1.CrossProduct(vbl2)
            Math.DivRem(currentQ + 1, 2, oddQ)
#If Entry_Type = "Unilateral" Then
            oddQ = 0
#Else
              If oddQ = 1 Then
                v.ScaleBy(-1)
            End If
#End If

            If dir > 0 Then
                v.ScaleBy(-1)
            End If
            pt.TranslateBy(v)
            bl2 = bandLines.Item(2)
            l = sk3D.SketchLines3D.AddByTwoPoints(cl1.Geometry.MidPoint, pt, False)

            Try
                GC = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, cl1)
                gcpp = TryPerpendicular(l, cl1)

                l2 = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartPoint, l.StartPoint, False)
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l2, l)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()
                l2.Construction = True
                gcpp = sk3D.GeometricConstraints3D.AddPerpendicular(l, l2)
            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                Return Nothing
            End Try
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, 1)
            l.Construction = True
            constructionLines.Add(l)





            refLine = l
            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawTangentParallel() As SketchLine3D
        Dim l, cl5 As SketchLine3D
        Dim gc As GeometricConstraint3D
        Try

            cl5 = constructionLines.Item(5)
            l = sk3D.SketchLines3D.AddByTwoPoints(cl5.StartPoint, tangentLine.Geometry.MidPoint, False)
            gc = sk3D.GeometricConstraints3D.AddParallel(l, tangentLine)
            angleTangent = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, cl5)
            angleTangent.Driven = True
            'gc = TryPerpendicular(l, cl5)
            l.Construction = True

            sk3D.Solve()
            If IsFirstLineInsideCylinder() Then
                ForceFirstLineOutside()
            End If
            CorrectSecondLine()
            lastLine = l
            constructionLines.Add(l)

            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Function DrawCentroLine() As SketchLine3D

        Dim l As SketchLine3D
        Dim sp As SketchPoint3D
        Dim dc As DimensionConstraint3D
        centroPoint = tg.CreatePoint(0, 0, 0)
        sp = sk3D.SketchPoints3D.Add(centroPoint)
        sk3D.GeometricConstraints3D.AddGround(sp)
        l = sk3D.SketchLines3D.AddByTwoPoints(sp, tangentLine.EndPoint, False)
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
        Dim dir As Integer = GetParameter("direction")._Value
        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = curve3D.DrawTrobinaCurve(sk3D, q, dir)
        sk3D.GeometricConstraints3D.AddGround(curve)
        Return curve
    End Function

    Function DrawFirstLine() As SketchLine3D
        Dim l As SketchLine3D
        Dim dc, ac As DimensionConstraint3D
        Dim gc As GeometricConstraint3D
        Dim v1, v2, v3 As Vector
        Dim limit As Integer = 0
        Dim delta As Double

        Dim pt As Point = refLine.StartSketchPoint.Geometry
        Dim b As Double = GetParameter("b")._Value
        Try


            l = sk3D.SketchLines3D.AddByTwoPoints(tangentLine.EndPoint, curve.EndSketchPoint.Geometry, False)
            Try
                gc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            Catch ex As Exception

            End Try

            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If dc.Parameter._Value > 2 * b Then
                adjuster.AdjustDimConstrain3DSmothly(dc, b)
            End If
            firstLine = l
            tensorFirstLine = sk3D.SketchLines3D.AddByTwoPoints(zAxisLine.Geometry.MidPoint, firstLine.StartPoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(tensorFirstLine.StartPoint, zAxisLine)
            TryPerpendicular(zAxisLine, tensorFirstLine)
            tensorFirstLine.Construction = True
            If dc.Parameter._Value > b Then
                Do
                    delta = (b - dc.Parameter._Value) / b
                    adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * Math.Exp(delta))
                    CorrectFirstLine()
                    limit = limit + 1
                Loop Until (limit > 8 Or Math.Abs(delta) < 1 / 1000)


            End If
            dc.Delete()
            bandLines.Add(l)
            firstLine = l
            lastLine = l

            If sk3D.Name = "s7" Then
                If IsFirstLineInside() Then
                Else
                    ForceFirstLineInside()
                    dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                    adjuster.AdjustDimensionConstraint3DSmothly(dc, b)
                    dc.Delete()
                End If
                If IsFirstLineInsideCylinder() Then
                    ForceFirstLineOutside()
                End If
            End If
            inletGap = sk3D.DimensionConstraints3D.AddTwoPointDistance(firstLine.StartPoint, intersectionPoint)
            If inletGap.Parameter._Value > 2 * b Then
                adjuster.AdjustDimConstrain3DSmothly(inletGap, inletGap.Parameter._Value / 2)
            End If
            inletGap.Driven = True
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function

    Public Function CorrectFirstLine() As Boolean
        If IsFirstLineInsideCylinder() Then
            Return ForceFirstLineOutside()
        Else
            Return True
        End If

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
            While (Not IsFirstLineInside() And limit < 8)
                Try
                    dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(firstLine.EndPoint, curve.EndSketchPoint)
                    adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 3 / 4)
                    dc.Delete()
                    Try
                        dc.Delete()
                    Catch ex2 As Exception

                    End Try
                    limit = limit + 1
                Catch ex As Exception
                    limit = limit + 1
                End Try

            End While
        Catch ex As Exception
            Try
                dc.Delete()
            Catch ex2 As Exception

            End Try
        End Try
        Return IsFirstLineInside()
    End Function

    Function DrawSecondLine() As SketchLine3D
        Dim b As Double = GetParameter("b")._Value
        Dim gc As GeometricConstraint3D
        Try
            Dim v1, v2, v3, v4 As Vector
            v4 = intersectionPoint.Geometry.VectorTo(curve.EndSketchPoint.Geometry)
            v4.Normalize()
            v1 = firstLine.Geometry.Direction.AsVector
            v1.Normalize()
            v3 = tangentLine.Geometry.Direction.AsVector
            v3.Normalize()
            v3.AddVector(v4)
            Dim l As SketchLine3D
            Dim d As Double
            Dim dc, ac As DimensionConstraint3D
            Dim optpoint As Point = curve.EndSketchPoint.Geometry
            Dim p As Plane
            p = tg.CreatePlane(lastLine.EndSketchPoint.Geometry, v1)
            Dim minDis As Double = 9999999999
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
                'l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, o, False)
                v2 = lastLine.EndSketchPoint.Geometry.VectorTo(o)
                d = v2.DotProduct(v3) * v1.CrossProduct(v2).Length
                If d > 0 Then
                    If d < minDis Then
                        minDis = d
                        optpoint = o
                    End If
                End If
            Next
            Try
                l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, optpoint, False)
            Catch ex As Exception
                l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, curve.EndSketchPoint.Geometry, False)
                gc = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value / 2)
                dc.Delete()
                gc.Delete()

                Try
                    dc.Delete()
                Catch ex3 As Exception

                End Try
            End Try
            Try
                sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            Catch ex As Exception
                dc = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, b)
                dc.Delete()
                CorrectFirstLine()
                sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            End Try

            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, firstLine)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()
            Catch ex2 As Exception
                ac.Delete()
            End Try

            Try
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                If dc.Parameter._Value > GetParameter("b")._Value * 2 Then
                    Try
                        adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value)
                    Catch ex As Exception
                        dc.Driven = True
                    End Try
                End If
                dc.Delete()
            Catch ex As Exception
                dc.Driven = True
            End Try
            'metro.Driven = True
            lastLine = l
            bandLines.Add(l)
            secondLine = l
            If sk3D.Name = "s7" Then
                If IsFirstLineInside() Then
                Else
                    ForceFirstLineInside()
                End If
                If IsFirstLineInsideCylinder() Then
                    ForceFirstLineOutside()
                End If
            End If
            FitEntryGap()
            CorrectAngleFirstLine()
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(secondLine.StartPoint, doblezline.EndPoint)
            adjuster.AdjustDimConstrain3DSmothly(dc, tangentLine.Length * 3 / 4)
            dc.Delete()
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
            Dim dc, ac As DimensionConstraint3D

            'Dim v As Vector
            'v = firstLine.Geometry.Direction.AsVector
            'v.ScaleBy(GetParameter("b")._Value / 1)
            'endPoint = firstLine.StartSketchPoint.Geometry
            'endPoint.TranslateBy(v)
            l = sk3D.SketchLines3D.AddByTwoPoints(firstLine.StartPoint, firstLine.EndSketchPoint.Geometry, False)
            ' endPoint.TranslateBy(v)

            Try
                sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, secondLine)
            Catch ex As Exception
                dc = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 6 / 5)
                dc.Delete()
                sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, secondLine)
            End Try
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, secondLine)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Delete()
            sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)

            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value) Then
                dc.Parameter._Value = GetParameter("b")._Value

            Else
                dc.Driven = True
            End If
            l.Construction = True
            constructionLines.Add(l)



            lastLine = l
            If sk3D.Name = "s7" Then
                If IsFirstLineInside() Then
                Else
                    ForceFirstLineInside()

                End If
                If IsFirstLineInsideCylinder() Then
                    ForceFirstLineOutside()
                End If
            End If
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(secondLine.StartPoint, doblezline.EndPoint)
            adjuster.AdjustDimConstrain3DSmothly(dc, tangentLine.Length * 3 / 4)
            dc.Delete()
            CorrectEntryGap()
            CorrectAngleFirstLine()
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
            Dim b As Double = GetParameter("b")._Value
            bl4 = bandLines.Item(4)
            cl1 = constructionLines.Item(1)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndPoint, bl4.Geometry.MidPoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, bl4)
            AdjustThirdLine(2 * b)
            ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
            adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            ac.Delete()
            Try

                sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
            Catch ex As Exception
                gapFold.Driven = True
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()

                Try
                    sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
                Catch ex3 As Exception
                    gapFold.Driven = True
                    dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                    adjuster.AdjustDimensionConstraint3DSmothly(dc, b * 2 / 3)
                    dc.Delete()
                    sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
                    gapFold.Driven = False
                End Try


            End Try
            AdjustThirdLine(2 * b)
            CorrectSecondLine()
            dc = sk3D.DimensionConstraints3D.AddLineLength(l)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, b)
            dc.Delete()

            Try
                gc = sk3D.GeometricConstraints3D.AddEqual(l, cl1)
            Catch ex As Exception
                gapFold.Driven = True
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, b * 2 / 3)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, b)
                dc.Delete()
                gc = sk3D.GeometricConstraints3D.AddEqual(l, cl1)
                gapFold.Driven = False
            End Try
            AdjustThirdLine(2 * b)
            FitEntryGap()
            CorrectSecondLine()
            CorrectGapFold()

            If gapFold.Parameter._Value > 2 * gapFoldCM Then
                adjuster.AdjustDimConstrain3DSmothly(gapFold, gapFoldCM * 2)
            End If
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl4.EndPoint, curve.EndSketchPoint)
            If dc.Parameter._Value > 2 * b Then
                adjuster.AdjustDimConstrain3DSmothly(dc, b)
            End If
            dc.Delete()
            ' gapFold.Driven = True
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
        Dim dcwo1 As DimensionConstraint3D
        Dim gc, gcel As GeometricConstraint3D
        Dim dcl4 As Double
        Try
            Dim c As Integer
            Dim b As Double = GetParameter("b")._Value
            Dim bl4, cl3, l As SketchLine3D
            Dim dc As TwoPointDistanceDimConstraint3D
            Dim dcl, dc2, dcbl1, dcbl4, ac As DimensionConstraint3D


            bl4 = bandLines.Item(4)
            cl3 = constructionLines.Item(3)

            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndSketchPoint.Geometry, bl4.StartSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, bl4)
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.EndPoint, cl3.EndPoint)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value / 2)
            adjuster.AdjustDimensionConstraint3DSmothly(dc, 1 / 10)
            dc2 = sk3D.DimensionConstraints3D.AddLineLength(l)
            If dc2.Parameter._Value > curve3D.DP.b * 3 / 20 Then
                If adjuster.AdjustDimensionConstraint3DSmothly(dc2, curve3D.DP.b * 5 / 40) Then
                    dc2.Driven = True
                Else
                    ' dc2.Delete()
                End If
            Else
                dc2.Driven = True
            End If
            AdjustThirdLine(2 * b)
            gapFold.Driven = False
            If gapFold.Parameter._Value > gapFoldCM * 3 Then
                adjuster.AdjustDimConstrain3DSmothly(gapFold, gapFoldCM * 2)
                gapFold.Driven = True
            End If
            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
                ac.Delete()
            Catch ex As Exception
                Try
                    dc2.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                    dc2.Delete()
                    ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, bl4)
                    adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                    ac.Delete()
                    sk3D.GeometricConstraints3D.AddPerpendicular(l, bl4)
                Catch ex3 As Exception

                End Try

            End Try
            ' CorrectFirstLine()
            FitEntryGap()
            adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 3)
            'gapFold.Driven = True
#If Entry_Type = "Unilateral" Then

            adjuster.AdjustDimConstrain3DSmothly(dc, 3 / 100)
            dcl4 = l.Length
            For i = 1 To 16
                adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                If Math.Abs(dcl4 - l.Length) < 1 / 1000 Then
                    Exit For
                Else
                    dcl4 = l.Length
                End If
            Next
            dc2.Delete()
#Else
            Try
                dc2.Driven = False
                adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                dc2.Driven = True
                gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                dc2.Delete()
            Catch ex As Exception
                Try
                    dc2.Driven = False
                    adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                    dc2.Driven = True
                    gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                Catch ex1 As Exception
                    Try

                        gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                    Catch ex3 As Exception
                        dc2.Driven = False
                        adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                        dc2.Driven = True
                        Try
                            gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                        Catch ex4 As Exception
                            dc2.Driven = False
                            adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                            dc2.Driven = True
                            Try
                                gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                            Catch ex5 As Exception
                                If adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value) Then
                                    dc2.Parameter._Value = GetParameter("b")._Value
                                End If
                                Try
                                    gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                                Catch ex8 As Exception
                                    Try
                                        dcbl1 = sk3D.DimensionConstraints3D.AddLineLength(firstLine)
                                        dcbl4 = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                                        Try
                                            While Math.Abs(bl4.Length - 2 * firstLine.Length) > gapFoldCM And c < 128
                                                adjuster.AdjustDimensionConstraint3DSmothly(dcbl1, bl4.Length / 2)
                                                c = c + 1
                                            End While
                                        Catch ex9 As Exception
                                            dcbl1.Driven = True
                                            dcbl4.Driven = True
                                        End Try
                                        dcbl4.Delete()
                                        dcbl1.Delete()
                                        Try
                                            dc2.Driven = True
                                            gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                                        Catch ex6 As Exception
                                            dc2.Driven = False
                                            adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value * 1.01)
                                            adjuster.AdjustDimensionConstraint3DSmothly(dc2, GetParameter("b")._Value)
                                            dc2.Driven = True
                                            gcel = sk3D.GeometricConstraints3D.AddEqual(l, cl3)
                                        End Try

                                    Catch ex2 As Exception
                                        Try
                                            dcbl1.Driven = True
                                            dcbl4.Driven = True
                                        Catch ex7 As Exception

                                        End Try

                                    End Try
                                End Try
                            End Try

                        End Try

                    End Try
                End Try


            End Try

#End If


            Try
                dcl = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                If dcl.Parameter._Value < GetParameter("b")._Value * 2 Then
                    adjuster.AdjustDimensionConstraint3DSmothly(dcl, GetParameter("b")._Value * 5)
                End If

                dcl.Delete()

            Catch ex As Exception



            End Try
            CorrectEntryGap()
            CorrectGapFold()
            FitEntryGap()
            adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 3)
            gapFold.Driven = False
            l.Construction = True
            constructionLines.Add(l)

            lastLine = l
            Return l

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function FitEntryGap() As Boolean


        Dim d As Double

        Dim limit As Integer = 0
        Dim dc, dcbl2 As DimensionConstraint3D
        Dim hecho As Boolean
        Try
            d = Math.Abs(-inletGap.Parameter._Value + outletGap.Parameter._Value) / inletGap.Parameter._Value * 100
            If d > 1 Then


                ' l = sk3D.SketchLines3D.AddByTwoPoints(cl2.StartPoint, tangentLine.StartPoint, False)
                ' dc = sk3D.DimensionConstraints3D.AddLineLength(tangentLine)
                '  dc.Driven = True

                Do
                    inletGap.Driven = False
                    hecho = adjuster.AdjustDimensionConstraint3DSmothly(inletGap, outletGap.Parameter._Value)
                    inletGap.Driven = True
                    CorrectFirstLine()
                    d = Math.Abs(-inletGap.Parameter._Value + outletGap.Parameter._Value) / inletGap.Parameter._Value * 100
                    limit = limit + 1
                Loop Until (d < 1 Or limit > 16)

                ' adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 2)
            Else
                Return True

            End If
            ' gapFold.Driven = True
            Return hecho
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try




    End Function
    Function CorrectGapFold() As Boolean
        Dim d As Double
        Dim currentQ As Integer = GetParameter("currentQ")._Value
        Dim oddQ As Integer
        Dim limit As Integer = 0
        Dim dc, dcbl2 As DimensionConstraint3D
        Dim hecho As Boolean
        Try

            d = CalculateGapError()
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(secondLine.StartPoint, doblezline.EndPoint)
            If d < 0 Then


                ' l = sk3D.SketchLines3D.AddByTwoPoints(cl2.StartPoint, tangentLine.StartPoint, False)
                ' dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(secondLine.StartPoint, doblezline.EndPoint)
                dc.Driven = True
                Math.DivRem(currentQ + 1, 2, oddQ)
#If Entry_Type = "Unilateral" Then
                oddQ = 1
#End If
                Do
                    If oddQ = 0 Then
                        adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value / 4)
                        gapFold.Driven = True
                        hecho = adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 17 / 16)
                    Else
                        adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value / 2)
                        gapFold.Driven = True
                        hecho = adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 7 / 8)
                    End If

                    dc.Driven = True

                    CorrectFirstLine()
                    d = CalculateGapError()
                    limit = limit + 1
                Loop Until (d > 0 Or limit > 16)
                dc.Driven = True
                adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 1)
            Else
                hecho = True

            End If
            dc.Delete()

            Return hecho
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function CalculateGapError() As Double
        Dim dir As Integer = GetParameter("direction")._Value
        Dim currentQ As Integer = GetParameter("currentQ")._Value
        Dim oddQ As Integer
        Dim vbl1, vbl4, vbl2 As Vector
        Dim bl4 As SketchLine3D
        Dim d As Double
        bl4 = bandLines.Item(4)
        vbl4 = bl4.Geometry.Direction.AsVector
        vbl1 = firstLine.Geometry.Direction.AsVector
        vbl2 = secondLine.Geometry.Direction.AsVector
        d = vbl1.CrossProduct(vbl4).DotProduct(vbl2)
        If dir < 0 Then
            d *= -1
        End If

#If Entry_Type = "Unilateral" Then
#Else
Math.DivRem(currentQ + 1, 2, oddQ)
   If oddQ = 0 Then
            d = -1 * d
        End If
#End If

        Return d
    End Function
    Function DrawThirdLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing

            Dim pl As SketchLine3D = Nothing
            Dim dc As DimensionConstraint3D
            Dim b As Double = GetParameter("b")._Value
            'pl = sk3D.SketchLines3D.Item(sk3D.SketchLines3D.Count - 1)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndPoint, firstLine.StartPoint, False)
            thirdLine = l
            bandLines.Add(l)
            If Not IsFirstLineInside() Then
                ForceFirstLineInside()
            End If
            AdjustThirdLine(b)

            If secondLine.Length > b Then
                dc = sk3D.DimensionConstraints3D.AddLineLength(secondLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, b)
                dc.Delete()
            End If
            FitEntryGap()
            If secondLine.Length > b Then
                dc = sk3D.DimensionConstraints3D.AddLineLength(secondLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, b)
                dc.Delete()
            End If
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(secondLine.StartPoint, doblezline.EndPoint)
            adjuster.AdjustDimConstrain3DSmothly(dc, tangentLine.Length * 3 / 4)
            dc.Delete()
            AdjustThirdLine(2 * b)
            CorrectEntryGap()
            CorrectAngleFirstLine()
            CorrectSecondLine()

            lastLine = l
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawSecondConstructionLine() As SketchLine3D
        Dim v1, v2, v3 As Vector
        Dim l As SketchLine3D
        Dim endPoint As Point
        Dim dc, ac As DimensionConstraint3D
        Dim gc As GeometricConstraint3D
        Try

            Dim b As Double = GetParameter("b")._Value
            Dim dir As Integer = GetParameter("direction")._Value
            Dim currentQ As Integer = GetParameter("currentQ")._Value
            Dim oddQ As Integer
            AdjustThirdLine(2 * b)
            endPoint = secondLine.StartSketchPoint.Geometry
#If Entry_Type = "Unilateral" Then
            v1 = firstLine.Geometry.Direction.AsVector()
            v2 = secondLine.Geometry.Direction.AsVector()
            v3 = v1.CrossProduct(v2).AsUnitVector().AsVector()
            v3.ScaleBy(gapFoldCM)
            endPoint.TranslateBy(v3)
#Else
            v2 = tangentLine.Geometry.Direction.AsVector
            v3 = firstLine.Geometry.Direction.AsVector
            v1 = v2.CrossProduct(v3)
            If dir < 0 Then
                v1.ScaleBy(-1)
            End If

            Math.DivRem(currentQ + 1, 2, oddQ)
            If oddQ = 0 Then
                v1.ScaleBy(-1)
            End If
            endPoint.TranslateBy(v1)

#End If



            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, endPoint, False)

            gapFold = sk3D.DimensionConstraints3D.AddLineLength(l)

            ' gc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, firstLine)
            gc = sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            Try
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, secondLine)
                adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
                ac.Delete()
                sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)
            Catch ex As Exception
                adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFold.Parameter._Value * 9 / 8)
                ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, secondLine)
                Try
                    ac.Parameter._Value = Math.PI / 2
                    sk3D.Solve()

                Catch ex3 As Exception

                End Try

                ac.Delete()
                sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)


            End Try
            Try
                dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, firstLine)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, Math.PI / 2)
                dc.Delete()

            Catch ex As Exception
                Try
                    dc.Delete()
                Catch ex3 As Exception

                End Try
            End Try
            If IsFirstLineInsideCylinder() Then
                ForceFirstLineOutside()
            End If
            adjuster.AdjustDimensionConstraint3DSmothly(gapFold, gapFoldCM * 4)
            gapFold.Driven = True
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

    Function DrawFourLine() As SketchLine3D
        Dim b As Double = GetParameter("b")._Value
        Try
            Dim l, pl As SketchLine3D
            Dim dcl, acl, dcl4 As DimensionConstraint3D
            pl = firstLine
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.StartPoint, lastLine.EndSketchPoint.Geometry, False)
            Try
                acl = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, lastLine)
                adjuster.AdjustDimensionConstraint3DSmothly(acl, Math.PI / 2)
                acl.Delete()
                sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
            Catch ex As Exception
                acl = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, lastLine)
                adjuster.AdjustDimensionConstraint3DSmothly(acl, Math.PI / 2)
                acl.Driven = True
                sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
            End Try
            AdjustThirdLine(2 * b)
            CorrectSecondLine()
            Try
                sk3D.GeometricConstraints3D.AddCoincident(lastLine.EndPoint, l)
            Catch ex As Exception
                dcl = sk3D.DimensionConstraints3D.AddTwoPointDistance(lastLine.EndPoint, l.EndPoint)
                adjuster.AdjustDimensionConstraint3DSmothly(dcl, gapFoldCM)
                dcl.Delete()
                Try
                    sk3D.GeometricConstraints3D.AddCoincident(lastLine.EndPoint, l)
                Catch ex2 As Exception
                    dcl = sk3D.DimensionConstraints3D.AddTwoPointDistance(lastLine.EndPoint, l.EndPoint)
                    adjuster.AdjustDimensionConstraint3DSmothly(dcl, gapFoldCM / 8)
                    dcl.Delete()
                    sk3D.GeometricConstraints3D.AddCoincident(lastLine.EndPoint, l)
                End Try

            End Try

            If IsFirstLineInsideCylinder() Then
                ForceFirstLineOutside()
            End If

            'sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            lastLine = l
            bandLines.Add(l)
            CorrectEntryGap()
            dcl4 = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.EndPoint, curve.EndSketchPoint)
            adjuster.AdjustDimConstrain3DSmothly(dcl4, gapFoldCM * 8)
            dcl4.Delete()
            curve_constraint = AnchorFourLine(l)
            ' Dim gc As GeometricConstraint3D = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)

            AdjustThirdLine(2 * b)
            CorrectSecondLine()
            AdjustFourLine()
            CorrectGapFold()

            CorrectEntryGap()
            adjuster.AdjustDimConstrain3DSmothly(gapFold, gapFoldCM * 2)
            'AdjustFourLine()
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function AnchorFourLine(li As SketchLine3D) As GeometricConstraint3D
        Dim l As SketchLine3D
        Dim dc As DimensionConstraint3D
        Dim gc As GeometricConstraint3D
        l = sk3D.SketchLines3D.AddByTwoPoints(li.EndSketchPoint, curve.EndSketchPoint.Geometry, False)
        gc = sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
        dc = sk3D.DimensionConstraints3D.AddLineLength(l)
        For index = 1 To 32
            adjuster.AdjustDimConstrain3DSmothly(dc, dc.Parameter._Value * 15 / 16)
            If l.Length < 1 / 10 Then
                Exit For
            End If
        Next
        dc.Delete()
        gc.Delete()
        l.Delete()
        curve_constraint = sk3D.GeometricConstraints3D.AddCoincident(li.EndPoint, curve)
        Return curve_constraint
    End Function
    Function CorrectAngleFirstLine() As Double
        Dim b As Double = GetParameter("b")._Value
        Dim l, fl As SketchLine3D
        Dim acl As DimensionConstraint3D


        CorrectAngleFirstLine = 0
        Try
            If firstLine.Equals(Nothing) Then
            Else
                fl = firstLine
                l = sk3D.SketchLines3D.AddByTwoPoints(fl.StartPoint, curve.EndSketchPoint, False)
                Try
                    acl = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, firstLine)
                    For i = 1 To 8
                        adjuster.AdjustDimensionConstraint3DSmothly(acl, acl.Parameter._Value * 7 / 8)
                        If acl.Parameter._Value < Math.PI / 128 Then
                            Exit For
                        End If
                        CorrectAngleFirstLine = acl.Parameter._Value
                    Next
                    acl.Delete()

                Catch ex As Exception
                    Try
                        acl.Delete()
                    Catch ex2 As Exception

                    End Try
                End Try


                l.Delete()
            End If


            Return CorrectAngleFirstLine
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
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
    Function AnchorFourLine() As SketchLine3D
        Dim bl4 As SketchLine3D = bandLines(4)
        Try

            Dim dc As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl4.EndPoint, curve.EndSketchPoint)
            adjuster.AdjustDimConstrain3DSmothly(dc, dc.Parameter._Value / 5)
            dc.Delete()
            sk3D.GeometricConstraints3D.AddCoincident(bl4.EndPoint, curve)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return bl4
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Function DrawFifthLine() As SketchLine3D
        Try
            Dim l, pl, bl4 As SketchLine3D
            Dim gc As GeometricConstraint3D
            Dim dc As DimensionConstraint3D
            Dim b As Double = GetParameter("b")._Value
            pl = thirdLine
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, pl.StartPoint, False)
            gapFold.Driven = False
            '   Try
            '   gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, pl)
            'Catch ex As Exception
            AdjustThirdLine(b * 2)
            '   ac = sk3D.DimensionConstraints3D.AddTwoLineAngle(l, pl)
            '  Try
            '  adjuster.AdjustDimensionConstraint3DSmothly(ac, Math.PI / 2)
            'Catch ex2 As Exception
            '   ac.Driven = True
            'End Try
            '    AdjustThirdLine(b * 2)
            '   gc = sk3D.GeometricConstraints3D.AddPerpendicular(l, pl)
            '  End Try
            'gapFold.Driven = True
            '  ac.Delete()
            sk3D.Solve()
            bl4 = bandLines(4)
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(bl4.EndPoint, curve.EndSketchPoint)
            If dc.Parameter._Value > 2 * b Then
                ' adjuster.AdjustDimConstrain3DSmothly(dc, b)
            End If
            dc.Delete()
            ' gc.Delete()
            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawSixthLine() As SketchLine3D
        Dim l, cl3, ml, bl4 As SketchLine3D
        Dim v As Vector
        Dim p As Plane
        Dim minDis, d As Double
        Dim optpoint As Point
        Dim dc As DimensionConstraint3D
        Dim puntos As ObjectsEnumerator
        Dim vc, vmjl, vbl4 As Vector
        Dim b As Double = GetParameter("b")._Value
        Try

            CorrectEntryGap()
            bl4 = bandLines.Item(4)
            If bl4.Length > 2 * b Then
                dc = sk3D.DimensionConstraints3D.AddLineLength(bl4)
                adjuster.AdjustDimensionConstraint3DSmothly(dc, 2 * b)
                dc.Delete()
            End If
            cl3 = constructionLines.Item(3)
            v = cl3.Geometry.Direction.AsVector
            p = tg.CreatePlane(cl3.EndSketchPoint.Geometry, v)


            puntos = p.IntersectWithCurve(curve.Geometry)

            Dim o2 As Point = puntos.Item(puntos.Count)
            optpoint = o2


            vmjl = thirdLine.Geometry.Direction.AsVector
            vmjl.ScaleBy(-1)

            vbl4 = bl4.Geometry.Direction.AsVector
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
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
    Public Function CorrectEntryGap() As Boolean


        Dim b As Double = GetParameter("b")._Value
        Dim factor As Double = Math.Exp((17.062 / 10 - outletGap.Parameter._Value) / (outletGap.Parameter._Value * 1))
        Dim d As Double = Math.Abs(inletGap.Parameter._Value - outletGap.Parameter._Value * factor)

        If d > 1 / 16 Then

            Dim dct As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(tangentLine)
            dct.Driven = True
            Dim limit As Integer = 0
            Dim hecho As Boolean
            CorrectFirstLine()
            Do
                hecho = adjuster.AdjustDimensionConstraint3DSmothly(inletGap, outletGap.Parameter._Value * factor)
                inletGap.Driven = True
                If tangentLine.Length > 5 * b Then
                    adjuster.AdjustDimConstrain3DSmothly(dct, dct.Parameter._Value * 31 / 32)
                    dct.Driven = True
                End If

                'FitEntryGap()
                d = Math.Abs(inletGap.Parameter._Value - outletGap.Parameter._Value * factor)
                limit += 1
            Loop Until (d < 1 / 16 Or limit > 16)
            inletGap.Driven = True
            dct.Delete()

            Return hecho
        Else
            Return True

        End If
        Return 0
    End Function
    Function CorrectThirdLine() As Boolean

        Dim b As Double = GetParameter("b")._Value
        Dim d As Double = Math.Abs(inletGap.Parameter._Value - outletGap.Parameter._Value)

        If d > 1 / 16 Then

            Dim dct As DimensionConstraint3D = sk3D.DimensionConstraints3D.AddLineLength(tangentLine)
            Dim limit As Integer = 0
            Dim hecho As Boolean
            CorrectFirstLine()
            Do
                inletGap.Driven = False
                AdjustThirdLine(2 * b)
                If inletGap.Parameter._Value > b Then
                    hecho = adjuster.AdjustDimensionConstraint3DSmothly(inletGap, inletGap.Parameter._Value * 7 / 8)
                Else
                    hecho = adjuster.AdjustDimensionConstraint3DSmothly(inletGap, outletGap.Parameter._Value)

                End If

                inletGap.Driven = True
                If tangentLine.Length > 4 * b Then
                    adjuster.AdjustDimConstrain3DSmothly(dct, 7 * b / 2)
                    dct.Driven = True
                End If

                d = Math.Abs(inletGap.Parameter._Value - outletGap.Parameter._Value)
                limit += 1
            Loop Until (d < 1 / 16 Or limit > 16)

            dct.Delete()

            Return hecho
        Else
            Return True

        End If
        Return 0
    End Function
    Public Function CorrectSecondLine() As Boolean
        Dim dc As DimensionConstraint3D
        Dim b As Double = GetParameter("b")._Value

        Dim currentQ As Integer = GetParameter("currentQ")._Value
        Dim oddQ As Integer

        Dim factor As Double
        Math.DivRem(currentQ + 1, 2, oddQ)
#If Entry_Type = "Unilateral" Then
        oddQ = 0
#End If
        If oddQ = 0 Then
            factor = 4 / 3
        Else
            factor = 5 / 6
        End If

        If secondLine.Length < b Or secondLine.Length > 2 * b Then
            dc = sk3D.DimensionConstraints3D.AddLineLength(secondLine)
            CorrectSecondLine = adjuster.AdjustDimConstrain3DSmothly(dc, b)
            dc.Delete()
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(secondLine.StartPoint, doblezline.EndPoint)
            '  dc2 = sk3D.DimensionConstraints3D.AddTwoPointDistance(intersectionPoint, doblezline.EndPoint)
            If dc.Parameter._Value > (intersectionPoint.Geometry.DistanceTo(doblezline.EndSketchPoint.Geometry)) * factor Then
                Try
                    gapFold.Driven = True
                    adjuster.AdjustDimConstrain3DSmothly(dc, tangentLine.Length * 7 / 8)
                    dc.Driven = True
                    adjuster.AdjustDimConstrain3DSmothly(gapFold, gapFoldCM * 2)
                Catch ex As Exception

                End Try

            End If
            dc.Delete()
            Return CorrectSecondLine

        End If
        Return True
    End Function
    Function CalculateEntryRodFactor() As Double
        Dim vz As Vector = tg.CreateVector(0, 0, -1)
        vz.Normalize()
        Dim vnsbl3 As Vector = thirdLine.Geometry.Direction.AsVector
        vnsbl3.Normalize()
        Dim d As Double = vnsbl3.DotProduct(vz)
        Return d
    End Function
    Function DrawSeventhLine() As SketchLine3D
        Dim l, cl As SketchLine3D
        Dim dc As DimensionConstraint3D
        Try

            cl = constructionLines.Item(3)
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.StartPoint, secondLine.EndPoint, False)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)

            dc = sk3D.DimensionConstraints3D.AddLineLength(lastLine)
            If adjuster.AdjustDimensionConstraint3DSmothly(dc, dc.Parameter._Value * 4 / 3) Then
                dc.Delete()
                dc = sk3D.DimensionConstraints3D.AddLineLength(l)
                If adjuster.AdjustDimensionConstraint3DSmothly(dc, GetParameter("b")._Value / 1) Then
                    dc.Delete()
                    sk3D.GeometricConstraints3D.AddEqual(l, cl)

                End If
            End If

            CorrectSecondLine()
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
