Imports Inventor
Public Class FoldingEvaluator
    Dim doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D
    Dim comando As Commands
    Dim refLine, firstLine, secondLine, thirdLine, lastLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim ring As Curves3D

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Dim partNumber As Integer
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines As ObjectCollection

    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim bendLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge As Edge
    Dim minorLine, majorLine As SketchLine3D
    Dim workFace, adjacentFace, bendFace As Face
    Dim bendAngle As DimensionConstraint
    Dim gapFold, crucialAngle As DimensionConstraint3D
    Dim folded As FoldFeature
    Dim sheetFeatures As SheetMetalFeatures
    Dim foldFeatures As FoldFeatures

    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        ring = New Curves3D(doku)

        adjuster = New SketchAdjust(doku)
        compDef = doku.ComponentDefinition
        sheetFeatures = compDef.Features
        foldFeatures = sheetFeatures.FoldFeatures
        comando = New Commands(app)


        done = False
    End Sub
    Public Sub Update(docu As Inventor.Document)
        doku = docu
        app = doku.Parent

        compDef = doku.ComponentDefinition
        sheetFeatures = compDef.Features
        foldFeatures = sheetFeatures.FoldFeatures


        done = False
    End Sub
    Public Function IsReadyForLastFold() As Boolean
        Try

            If GetCrucialAngle().Parameter._Value > (Math.PI / 2) * 148 / 157 Then
                Return True
            Else
                Return False
            End If
            comando.MakeInvisibleSketches(doku)
            Return False
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return False
        End Try

        Return False
    End Function
    Function GetLastBend() As Face
        Dim maxArea1, maxArea2, maxArea3 As Double

        Dim maxface1, maxface2, maxface3 As Face


        maxface2 = compDef.Bends.Item(compDef.Bends.Count).BackFaces.Item(1)
        maxface1 = maxface2
        maxArea2 = 0
        maxArea1 = maxArea2
        For Each f As Face In compDef.Bends.Item(compDef.Bends.Count).BackFaces
            'lamp.HighLighFace(f)
            If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                'lamp.HighLighFace(f)
                If f.Evaluator.Area > maxArea1 Then
                    maxArea2 = maxArea1
                    maxface2 = maxface1
                    maxArea1 = f.Evaluator.Area
                    maxface1 = f

                Else
                    maxface2 = f
                End If
            Else
                maxface3 = f
            End If
        Next
        bendFace = maxface2
        Return bendFace
    End Function
    Function GetLastBendLine() As SketchLine
        Try
            Dim i, j, k As Integer
            k = foldFeatures.Count
            i = Math.DivRem(k, 2, j)
            bendLine = foldFeatures.Item(2 * (i + j) - 1).Definition.BendLine
            Return bendLine
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function DrawEvaluationSketch() As Sketch3D
        Try
            sk3D = doku.ComponentDefinition.Sketches3D.Add()
            firstLine = sk3D.Include(GetLastBendLine())
            curve = ring.DrawLowerRing(sk3D)
            sk3D.GeometricConstraints3D.AddGround(curve)
            secondLine = sk3D.SketchLines3D.AddByTwoPoints(firstLine.EndPoint, curve.StartSketchPoint.Geometry, False)
            sk3D.GeometricConstraints3D.AddCoincident(secondLine.EndPoint, curve)
            thirdLine = sk3D.SketchLines3D.AddByTwoPoints(firstLine.StartPoint, secondLine.EndPoint, False)
            Return sk3D
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try



        Return Nothing
    End Function
    Function GetCrucialAngle() As DimensionConstraint3D
        Try
            Dim dc1, dc2, a As DimensionConstraint3D
            Try
                If compDef.Sketches3D.Item("crucialAngle").DimensionConstraints3D.Count > 0 Then
                    sk3D = compDef.Sketches3D.Item("crucialAngle")
                    crucialAngle = GetDimensionConstraint("crucialAngle")
                End If
                Return crucialAngle
            Catch ex As Exception
                sk3D = DrawEvaluationSketch()
                sk3D.Name = "crucialAngle"
                dc2 = sk3D.DimensionConstraints3D.AddLineLength(secondLine)
                adjuster.UpdateDocu(doku)
                Try
                    adjuster.GetMinimalDimension(dc2)
                Catch ex2 As Exception

                End Try
                dc2.Driven = True
                dc1 = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
                Try
                    adjuster.GetMinimalDimension(dc1)
                Catch ex3 As Exception
                    dc1.Driven = True
                End Try
                dc1.Delete()
                dc2.Driven = False
                If adjuster.GetMinimalDimension(dc2) Then

                    dc2.Delete()
                    dc1 = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
                    If adjuster.GetMinimalDimension(dc1) Then
                        dc1.Driven = True
                        a = sk3D.DimensionConstraints3D.AddTwoLineAngle(secondLine, thirdLine)
                        crucialAngle = a
                    Else
                        dc1.Delete()
                        a = sk3D.DimensionConstraints3D.AddTwoLineAngle(secondLine, thirdLine)
                        crucialAngle = a

                    End If

                Else
                    dc1 = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
                    If adjuster.GetMinimalDimension(dc1) Then
                        dc1.Delete()
                        dc2 = sk3D.DimensionConstraints3D.AddLineLength(thirdLine)
                        If adjuster.GetMinimalDimension(dc2) Then
                            dc2.Driven = True
                            a = sk3D.DimensionConstraints3D.AddTwoLineAngle(secondLine, thirdLine)
                            crucialAngle = a
                        Else
                            dc1.Delete()
                            a = sk3D.DimensionConstraints3D.AddTwoLineAngle(secondLine, thirdLine)
                            crucialAngle = a

                        End If
                    Else
                        dc1.Delete()
                        a = sk3D.DimensionConstraints3D.AddTwoLineAngle(secondLine, thirdLine)
                        crucialAngle = a
                    End If
                    Return Nothing
                End If
                crucialAngle.Parameter.Name = "crucialAngle"
                Return a

            End Try

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Public Function GetLastFold(d As Document) As FoldFeature
        Update(d)
        For index = 1 To foldFeatures.Count
            If foldFeatures.Item(foldFeatures.Count - index + 1).Type = ObjectTypeEnum.kFoldFeatureObject Then
                Return foldFeatures.Item(foldFeatures.Count - index + 1)

            End If
        Next

        Return Nothing
    End Function
    Function GetDimensionConstraint(name As String) As DimensionConstraint3D
        For Each dimension In sk3D.DimensionConstraints3D
            If dimension.Parameter.Name = name Then
                Return dimension
            End If
        Next
        Return Nothing
    End Function
End Class
