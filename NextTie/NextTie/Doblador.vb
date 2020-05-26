Imports Inventor

Public Class Doblador
    Dim doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D
    Dim lines3D As SketchLines3D
    Dim refLine, firstLine, secondLine, thirdLine, lastLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean
    Dim curve3D As Curves3D
    Dim monitor As DesignMonitoring

    Public wp1, wp2, wp3 As WorkPoint
    Public vp, point1, point2, point3 As Point
    Dim tg As TransientGeometry
    Dim gap1 As Double
    Dim adjuster As SketchAdjust
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim mainSketch As OriginSketch
    Dim pro As Profile
    Dim feature As FaceFeature
    Public bendLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim edgeBand As Edge
    Dim bendAngle As DimensionConstraint
    Public folded As FoldFeature
    Dim foldDefinition As FoldDefinition
    Dim features As SheetMetalFeatures
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)

        curve3D = New Curves3D(doku)
        monitor = New DesignMonitoring(doku)
        adjuster = New SketchAdjust(doku)
        compDef = doku.ComponentDefinition
        features = compDef.Features

        mainSketch = New OriginSketch(doku)
        tg = app.TransientGeometry

        done = False
    End Sub


    Public Function MakeFirstFold(refDoc As FindReferenceLine) As Boolean
        Dim b As Boolean
        Try


            sk3D = mainSketch.StartDrawingTranslated(refDoc, 1)
            If mainSketch.done Then
                doku.ComponentDefinition.Sketches3D.Item("s1").Visible = False
                If DrawBandStripe().Count > 0 Then
                    If MakeStartingFace(pro).GetHashCode > 0 Then


                    End If
                End If

            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
    Function DrawBandStripe() As Profile
        Dim ps As PlanarSketch

        Dim p1, p2, p3 As SketchPoint
        Dim v As Vector2d
        Dim r As SketchEntitiesEnumerator
        Dim ln, l2d As SketchLine
        Try
            mainWorkPlane = doku.ComponentDefinition.WorkPlanes.AddByThreePoints(mainSketch.firstLine.StartPoint, mainSketch.firstLine.EndPoint, mainSketch.secondLine.EndPoint)
            ps = doku.ComponentDefinition.Sketches.Add(mainWorkPlane)
            p1 = ps.AddByProjectingEntity(mainSketch.firstLine.StartPoint)
            p2 = ps.AddByProjectingEntity(mainSketch.firstLine.EndPoint)
            p3 = ps.AddByProjectingEntity(mainSketch.secondLine.EndPoint)
            ln = ps.AddByProjectingEntity(mainSketch.bandLines.Item(2))
            v = ln.Geometry.Direction.AsVector
            v.ScaleBy(-3 / 10)
            p1.MoveBy(v)
            l2d = ps.SketchLines.AddByTwoPoints(p1, p2.Geometry)
            ps.GeometricConstraints.AddPerpendicular(l2d, ln)
            Dim dc As DimensionConstraint
            dc = ps.DimensionConstraints.AddTwoPointDistance(p1, l2d.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, p2.Geometry)
            dc.Parameter._Value = GetParameter("b")._Value / 10
            v.Normalize()
            v.ScaleBy(-50)
            p3.MoveBy(v)
            r = ps.SketchLines.AddAsThreePointRectangle(p1.Geometry, l2d.EndSketchPoint.Geometry, p3.Geometry)

            pro = ps.Profiles.AddForSolid

            Return pro

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeStartingFace(pro As Profile) As FaceFeature
        Dim oSheetMetalCompDef As SheetMetalComponentDefinition
        oSheetMetalCompDef = doku.ComponentDefinition
        oSheetMetalCompDef.UseSheetMetalStyleThickness = False
        Dim oThicknessParam As Parameter
        oThicknessParam = oSheetMetalCompDef.Thickness
        oThicknessParam.Units = UnitsTypeEnum.kMillimeterLengthUnits
        oThicknessParam._Value = 0.03

        Dim oSheetMetalFeatures As SheetMetalFeatures
        oSheetMetalFeatures = doku.ComponentDefinition.Features

        compDef = oSheetMetalCompDef
        Dim oFaceFeatureDefinition As FaceFeatureDefinition
        oFaceFeatureDefinition = oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(pro)
        oFaceFeatureDefinition.Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection
        Dim oFaceFeature As FaceFeature
        oFaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition)
        feature = oFaceFeature
        Return oFaceFeature
    End Function
    Public Function FoldBand(i As Integer) As FoldFeature
        Try



            Dim oFoldFeature As FoldFeature
            oFoldFeature = features.FoldFeatures.Add(AdjustFoldDefinition(i))
            folded = oFoldFeature

            Return folded

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function CorrectFold(ff As FoldFeature) As FoldFeature
        Try
            ff.Delete()
            foldDefinition.IsPositiveBendSide = Not foldDefinition.IsPositiveBendSide

            folded = features.FoldFeatures.Add(foldDefinition)
            Return folded

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function AdjustFoldDefinition(i As Integer) As FoldDefinition
        Try
            features = doku.ComponentDefinition.Features
            Dim oFoldDefinition As FoldDefinition
            If i > 3 Then

                If bendAngle.Parameter._Value > Math.PI / 2 Then
                    oFoldDefinition = features.FoldFeatures.CreateFoldDefinition(bendLine, bendAngle.Parameter._Value)
                Else
                    oFoldDefinition = features.FoldFeatures.CreateFoldDefinition(bendLine, Math.PI - bendAngle.Parameter._Value)
                End If

            Else
                If bendAngle.Parameter._Value > Math.PI / 2 Then
                    oFoldDefinition = features.FoldFeatures.CreateFoldDefinition(bendLine, Math.PI - bendAngle.Parameter._Value)
                Else
                    oFoldDefinition = features.FoldFeatures.CreateFoldDefinition(bendLine, bendAngle.Parameter._Value)
                End If

            End If

            foldDefinition = oFoldDefinition
            Return oFoldDefinition
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Public Function GetBendLine(workFace As Face, line As SketchLine3D) As SketchLine
        Try

            Dim ps As PlanarSketch

            ps = compDef.Sketches.Add(workFace)

            Dim sl, bl As SketchLine
            sl = ps.AddByProjectingEntity(line)

            sl.Construction = True
            bl = ps.SketchLines.AddByTwoPoints(sl.EndSketchPoint.Geometry, sl.StartSketchPoint.Geometry)



            bendLine = bl
            Return bl
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

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
    Public Function GetFoldingAngle(mne As Edge, sl3D As SketchLine3D) As DimensionConstraint

        Try
            Dim ps As PlanarSketch

            Dim wp As WorkPlane
            wp = compDef.WorkPlanes.AddByNormalToCurve(bendLine, bendLine.StartSketchPoint)

            ps = doku.ComponentDefinition.Sketches.Add(wp)
            Dim sl, el As SketchLine
            sl = ps.AddByProjectingEntity(sl3D)
            el = ps.AddByProjectingEntity(mne)
            Dim dc As DimensionConstraint
            dc = ps.DimensionConstraints.AddTwoLineAngle(el, sl, sl.EndSketchPoint.Geometry)
            dc.Driven = True
            bendAngle = dc


            Return dc

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
End Class
