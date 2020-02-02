Imports Inventor
Imports System
Imports GetInitialConditions
Imports DrawInitialSketch

Public Class Sketcher3D
    Dim doku As PartDocument
    Dim app As Application
    Public sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Public refLine, firstLine, secondLine, thirdLine, lastLine As SketchLine3D
    Public curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim curve3D As Curve3D
    Dim monitor As DesignMonitoring
    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3 As Point
    Dim tg As TransientGeometry
    Dim gap1 As Double
    Dim adjuster As SketchAdjust
    Public bandLines, constructionLines As ObjectCollection
    Dim comando As Commands




    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)

        curve3D = New Curve3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        gap1 = 3 / 10
        done = False
    End Sub

    Public Function StartDrawing(ref As SketchLine3D) As Sketch3D
        refLine = ref
        wp1 = ref.StartSketchPoint.Geometry
        wp3 = ref.EndSketchPoint.Geometry
        sk3D = doku.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = "s0"
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
    Function DrawInitialLine(line As SketchLine3D) As SketchLine3D

        sk3D.SketchLines3D.AddByTwoPoints(line.StartSketchPoint.Geometry, line.EndSketchPoint.Geometry)
        sk3D.SketchLines3D.Item(sk3D.SketchLines3D.Count).Construction = True
        Return line
    End Function
    Function DrawTrobinaCurve(q As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(q, "s0")
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
            Debug.Print(ex.ToString())
            Return Nothing
        End Try
        Return Nothing
    End Function
    Function DrawSecondLine() As SketchLine3D
        Try
            Dim v As Vector = lastLine.StartSketchPoint.Geometry.VectorTo(lastLine.EndSketchPoint.Geometry)
            Dim p As Plane
            Dim optpoint As Point = Nothing
            p = tg.CreatePlane(lastLine.EndSketchPoint.Geometry, v)
            Dim minDis As Double = 9999999999
            For Each o As Point In p.IntersectWithCurve(curve.Geometry)
                If o.DistanceTo(lastLine.EndSketchPoint.Geometry) < minDis Then
                    minDis = o.DistanceTo(lastLine.EndSketchPoint.Geometry)
                    optpoint = o
                End If

            Next
            Dim l As SketchLine3D = Nothing
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, optpoint, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            point3 = l.EndSketchPoint.Geometry
            lastLine = l
            bandLines.Add(l)
            secondLine = lastLine


            Return l
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawFirstConstructionLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing
            Dim endPoint As Point
            Dim v As Vector
            v = firstLine.Geometry.Direction.AsVector
            v.ScaleBy(GetParameter("b")._Value / 10)
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
            Debug.Print(ex.ToString())
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
            Debug.Print(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawFourthConstructionLine() As SketchLine3D
        Try
            Dim l As SketchLine3D = Nothing
            Dim ol, cl As SketchLine3D
            Dim dc As TwoPointDistanceDimConstraint3D
            ol = bandLines.Item(4)
            cl = constructionLines.Item(3)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.EndSketchPoint.Geometry, ol.StartSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, curve)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, ol)
            dc = sk3D.DimensionConstraints3D.AddTwoPointDistance(l.EndPoint, cl.EndPoint)
            If AdjustTwoPointsSmothly(dc, 1) Then
                sk3D.GeometricConstraints3D.AddPerpendicular(l, ol)
                sk3D.GeometricConstraints3D.AddEqual(l, cl)
                l.Construction = True

                constructionLines.Add(l)

                lastLine = l
                Return l
            End If
            Return Nothing
        Catch ex As Exception
            Debug.Print(ex.ToString())
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
            Debug.Print(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DrawSecondConstructionLine() As SketchLine3D
        Try
            Dim v1, v2, v3 As Vector
            Dim l, pl As SketchLine3D
            Dim endPoint As Point
            v1 = firstLine.EndSketchPoint.Geometry.VectorTo(firstLine.StartSketchPoint.Geometry)
            v2 = secondLine.EndSketchPoint.Geometry.VectorTo(secondLine.StartSketchPoint.Geometry)
            v3 = v1.CrossProduct(v2).AsUnitVector().AsVector()
            v3.ScaleBy(gap1)
            endPoint = secondLine.StartSketchPoint.Geometry
            endPoint.TranslateBy(v3)
            l = sk3D.SketchLines3D.AddByTwoPoints(secondLine.StartSketchPoint.Geometry, endPoint, False)
            endPoint.TranslateBy(v3)
            sk3D.DimensionConstraints3D.AddLineLength(l, endPoint, False)

            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, secondLine)
            sk3D.GeometricConstraints3D.AddCoincident(l.StartPoint, firstLine)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, secondLine)

            lastLine = l
            constructionLines.Add(l)
            l.Construction = True
            Return l
        Catch ex As Exception
            Debug.Print(ex.ToString())
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
            Debug.Print(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawFifthLine() As SketchLine3D
        Try
            Dim l, pl As SketchLine3D
            pl = thirdLine
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.EndPoint, pl.StartPoint, False)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, pl)

            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawSixthLine() As SketchLine3D
        Try
            Dim l, pl As SketchLine3D
            Dim v As Vector
            Dim p As Plane
            Dim minDis As Double
            Dim optpoint As Point
            pl = constructionLines.Item(3)
            v = pl.Geometry.Direction.AsVector
            p = tg.CreatePlane(pl.EndSketchPoint.Geometry, v)

            Dim puntos As ObjectsEnumerator
            puntos = p.IntersectWithCurve(curve.Geometry)
            minDis = 9999999999
            Dim o2 As Point = puntos.Item(puntos.Count)
            optpoint = o2
            For Each o As Point In puntos
                If o.DistanceTo(pl.EndSketchPoint.Geometry) < minDis Then
                    minDis = o.DistanceTo(pl.EndSketchPoint.Geometry)
                    o2 = optpoint
                    optpoint = o
                End If

            Next

            vp = o2

            pl = bandLines.Item(4)
            l = sk3D.SketchLines3D.AddByTwoPoints(pl.EndSketchPoint.Geometry, vp, False)
            sk3D.GeometricConstraints3D.AddCoincident(pl.EndPoint, l)
            sk3D.GeometricConstraints3D.AddCoincident(l.EndPoint, curve)
            lastLine = l
            bandLines.Add(l)
            Return l
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function DrawSeventhLine() As SketchLine3D
        Try
            Dim l, cl As SketchLine3D
            cl = constructionLines.Item(3)
            l = sk3D.SketchLines3D.AddByTwoPoints(lastLine.StartPoint, secondLine.EndPoint, False)
            sk3D.GeometricConstraints3D.AddPerpendicular(l, lastLine)
            sk3D.GeometricConstraints3D.AddEqual(l, cl)
            lastLine = l
            bandLines.Add(l)
            done = 1
            Return l
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try
    End Function

    Function AdjustlastAngle() As Boolean
        Dim fourLine, sixthLine As SketchLine3D
        Dim dc As TwoLineAngleDimConstraint3D
        Dim dName As String
        Dim b As Boolean = False

        fourLine = bandLines.Item(4)
        sixthLine = bandLines.Item(6)
        dc = sk3D.DimensionConstraints3D.AddTwoLineAngle(fourLine, sixthLine)
        dName = String.Concat(sk3D.Name & "a1p1z100c200f50")
        dc.Parameter.Name = dName

        If doku.Update2() Then
            If adjuster.UpdateDocu(doku) Then
                If GetParameter(dName)._Value > Math.PI / 2 Then
                    If adjuster.AdjustDimensionSmothly(dName, Math.PI - 0.1) Then
                        b = True
                    End If
                Else
                    If adjuster.AdjustDimensionSmothly(dName, 0.1) Then
                        b = True
                    End If
                End If
            End If
        End If



        Return b
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
            Debug.Print(ex.ToString())
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
                    Debug.Print("not healthy constrain")
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
                    Debug.Print(ex2.ToString())
                    Debug.Print("Parameter not found: " & name)
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
            Debug.Print(ex.ToString())
            healthy = False
        End Try


        Return healthy
    End Function

End Class
