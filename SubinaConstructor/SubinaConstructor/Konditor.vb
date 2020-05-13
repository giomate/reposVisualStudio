Imports Inventor
Imports System
Imports System.IO

Imports System.Text
Imports System.IO.Directory
Public Class Konditor
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application

    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile
    Dim adjuster As SketchAdjust


    Public wp1, wp2, wp3, wpConverge As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, convergePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double

    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Dim surfaceBodies As ObjectCollection




    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Public compDef As ComponentDefinition


    Dim features As PartFeatures
    Dim lamp As Highlithing
    Dim di As System.IO.DirectoryInfo
    Dim fi As System.IO.File
    Dim nf As System.IO.Path
    Dim bandFaces As WorkSurface



    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane

    Dim fullFileNames As String()

    Public Structure DesignParam
        Public p As Integer
        Public q As Integer
        Public b As Integer
        Public Dmax As Double
        Public Dmin As Double

    End Structure

    Dim DP As DesignParam
    Dim Tr As Double
    Dim Cr As Double
    Public Sub New(docu As Inventor.Document)
        doku = docu
        App = doku.Parent
        comando = New Commands(App)
        monitor = New DesignMonitoring(doku)
        invFile = New InventorFile(App)
        projectManager = App.DesignProjectManager

        compDef = doku.ComponentDefinition

        SurfaceBodies = app.TransientObjects.CreateObjectCollection

        tg = app.TransientGeometry
        bandLines = App.TransientObjects.CreateObjectCollection
        constructionLines = App.TransientObjects.CreateObjectCollection


        lamp = New Highlithing(doku)


        nombrador = New Nombres(doku)


        DP.Dmax = 200 / 10
        DP.Dmin = 1 / 10
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 13
        DP.q = 31
        DP.b = 25
        done = False
    End Sub
    Function DocUpdate(docu As PartDocument) As PartDocument
        doku = docu

        compDef = docu.ComponentDefinition
        features = docu.ComponentDefinition.Features
        lamp = New Highlithing(doku)
        monitor = New DesignMonitoring(doku)
        'adjuster = New SketchAdjust(doku)

        Return doku
    End Function
    Public Function MakeWedgesCake(i As Integer) As PartDocument
        Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim q As Integer
        Dim nd As String
        nd = String.Concat(p, "\Iteration", i.ToString)

        Dim t As PartDocument
        Try
            t = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
            Conversions.SetUnitsToMetric(t)
            t.Update()
            doku = DocUpdate(t)
            MakeRing()

            If (Directory.Exists(nd)) Then
                fullFileNames = Directory.GetFiles(nd, "*.ipt")
                For Each s As String In fullFileNames
                    If nombrador.ContainWedge(s) Then
                        MergeSingleWedge(s)
                    End If
                Next
                If monitor.IsFeatureHealthy(CombineBodies()) Then
                    If monitor.IsFeatureHealthy(RemoveSpikes()) Then

                        If doku.ComponentDefinition.SurfaceBodies.Count < 2 Then
                            doku.SaveAs(String.Concat(p, "\Iteration", i.ToString, "\Subina23.ipt"), False)
                            doku.Save()
                            done = True
                        End If
                    End If

                End If



            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return doku
    End Function
    Public Function MakeSkeletonsCake(i As Integer) As PartDocument
        Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim q As Integer
        Dim nd As String
        nd = String.Concat(p, "\Iteration", i.ToString)

        Dim t As PartDocument
        Try
            t = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
            Conversions.SetUnitsToMetric(t)
            t.Update()
            doku = DocUpdate(t)


            If (Directory.Exists(nd)) Then
                fullFileNames = Directory.GetFiles(nd, "*.ipt")
                For Each s As String In fullFileNames
                    If nombrador.ContainSkeleton(s) Then
                        DeriveSingleSkeleton(s)
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return doku
    End Function
    Function CombineBodies() As CombineFeature
        Dim cf As CombineFeature
        If doku.ComponentDefinition.SurfaceBodies.Count > 1 Then
            SurfaceBodies.Clear()

            For index = 2 To doku.ComponentDefinition.SurfaceBodies.Count
                SurfaceBodies.Add(doku.ComponentDefinition.SurfaceBodies.Item(index))
            Next
            Try
                cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), SurfaceBodies, PartFeatureOperationEnum.kJoinOperation)

            Catch ex As Exception


            End Try
        End If

        ' cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)

        Return cf
    End Function
    Function RemoveSpikes() As RevolveFeature
        Dim rf As RevolveFeature
        Dim pro As Profile = DrawRemoveFace()

        Try

            rf = doku.ComponentDefinition.Features.RevolveFeatures.AddFull(pro, doku.ComponentDefinition.WorkAxes.Item(3), PartFeatureOperationEnum.kCutOperation)
            Return rf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function DrawRemoveFace() As Profile
        Dim pro As Profile
        Dim skpl As PlanarSketch
        Dim l, cl As SketchLine
        Dim spt1, spt2, spt3 As SketchPoint
        Dim pt1, pt2, pt3 As Point2d
        Dim a, b, c, d, e As SketchArc
        Dim dc As DimensionConstraint
        Dim gc As GeometricConstraint
        Try
            skpl = doku.ComponentDefinition.Sketches.Add(doku.ComponentDefinition.WorkPlanes.Item(2))

            spt1 = skpl.SketchPoints.Add(tg.CreatePoint2d(0, Cr))
            spt2 = skpl.SketchPoints.Add(tg.CreatePoint2d(0, -Cr))
            l = skpl.SketchLines.AddByTwoPoints(spt1, spt2)
            cl = skpl.AddByProjectingEntity(doku.ComponentDefinition.WorkAxes.Item(3))
            cl.Construction = True
            gc = skpl.GeometricConstraints.AddCollinear(l, cl)

            dc = skpl.DimensionConstraints.AddTwoPointDistance(l.StartSketchPoint, l.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, l.EndSketchPoint.Geometry)
            spt1 = skpl.SketchPoints.Add(tg.CreatePoint2d(0, Cr / 2))
            spt2 = skpl.SketchPoints.Add(tg.CreatePoint2d(Cr / 2, Cr / 2))
            a = skpl.SketchArcs.AddByCenterStartEndPoint(spt1, l.StartSketchPoint, spt2, False)
            dc = skpl.DimensionConstraints.AddRadius(a, spt2.Geometry)
            ' dc.Parameter._Value = 37 / 10
            gc = skpl.GeometricConstraints.AddCoincident(a.CenterSketchPoint, l)
            spt1 = skpl.SketchPoints.Add(tg.CreatePoint2d(0, -Cr / 2))
            spt2 = skpl.SketchPoints.Add(tg.CreatePoint2d(Cr / 2, -Cr / 2))
            b = skpl.SketchArcs.AddByCenterStartEndPoint(spt1, l.EndSketchPoint, spt2)
            dc = skpl.DimensionConstraints.AddRadius(b, spt2.Geometry)
            gc = skpl.GeometricConstraints.AddCoincident(b.CenterSketchPoint, l)
            spt1 = skpl.SketchPoints.Add(tg.CreatePoint2d(Tr, 0))
            spt2 = skpl.SketchPoints.Add(tg.CreatePoint2d(Tr, Cr / 2 - 5 / 10))
            spt3 = skpl.SketchPoints.Add(tg.CreatePoint2d(Tr, -Cr / 2 + 5 / 10))
            c = skpl.SketchArcs.AddByCenterStartEndPoint(spt1, spt2, spt3, False)
            dc = skpl.DimensionConstraints.AddRadius(c, spt2.Geometry)
            gc = skpl.GeometricConstraints.AddGround(c.CenterSketchPoint)
            spt2 = skpl.SketchPoints.Add(tg.CreatePoint2d(Tr, Cr))
            l = skpl.SketchLines.AddByTwoPoints(c.CenterSketchPoint, spt2)
            l.Construction = True
            cl = skpl.AddByProjectingEntity(doku.ComponentDefinition.WorkAxes.Item(1))
            cl.Construction = True
            gc = skpl.GeometricConstraints.AddPerpendicular(cl, l)

            d = skpl.SketchArcs.AddByCenterStartEndPoint(l.EndSketchPoint, a.StartSketchPoint, c.EndSketchPoint)
            gc = skpl.GeometricConstraints.AddTangent(a, d)
            gc = skpl.GeometricConstraints.AddCoincident(d.EndSketchPoint, c)
            gc = skpl.GeometricConstraints.AddTangent(d, c)
            gc = skpl.GeometricConstraints.AddCoincident(c.EndSketchPoint, d)
            gc = skpl.GeometricConstraints.AddCoincident(d.CenterSketchPoint, l)
            spt2 = skpl.SketchPoints.Add(tg.CreatePoint2d(Tr, -Cr))
            l = skpl.SketchLines.AddByTwoPoints(l.StartSketchPoint, spt2)
            l.Construction = True
            gc = skpl.GeometricConstraints.AddPerpendicular(cl, l)

            e = skpl.SketchArcs.AddByCenterStartEndPoint(l.EndSketchPoint, b.EndSketchPoint, c.StartSketchPoint, False)
            gc = skpl.GeometricConstraints.AddTangent(b, e)
            gc = skpl.GeometricConstraints.AddCoincident(e.StartSketchPoint, c)
            gc = skpl.GeometricConstraints.AddTangent(e, c)
            gc = skpl.GeometricConstraints.AddCoincident(c.StartSketchPoint, e)
            gc = skpl.GeometricConstraints.AddCoincident(e.CenterSketchPoint, l)
            gc = skpl.GeometricConstraints.AddSymmetry(d, e, cl)
            pro = skpl.Profiles.AddForSolid
            Return pro
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function MergeSingleWedge(s As String) As DerivedPartComponent
        Dim derivedDefinition As DerivedPartDefinition
        Dim dpc As DerivedPartComponent
        Try



            derivedDefinition = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
            derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
            dpc = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)




            Return dpc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function DeriveSingleSkeleton(s As String) As DerivedPartComponent
        Dim derivedDefinition As DerivedPartDefinition
        Dim dpc As DerivedPartComponent
        Try



            derivedDefinition = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
            derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
            derivedDefinition.IncludeAllSurfaces = DerivedComponentOptionEnum.kDerivedIncludeAll
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
            dpc = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)




            Return dpc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

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
    Function DrawRing() As Profile
        Dim pro As Profile
        Dim ps As PlanarSketch
        Dim spt As SketchPoint
        Try
            ps = doku.ComponentDefinition.Sketches.Add(doku.ComponentDefinition.WorkPlanes.Item(2))
            spt = ps.SketchPoints.Add(tg.CreatePoint2d(Tr, 0))
            ps.SketchCircles.AddByCenterRadius(spt, Cr / 2 - 1 / 10)
            ps.SketchCircles.AddByCenterRadius(spt, Cr / 2 - 6 / 10)
            pro = ps.Profiles.AddForSolid
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return pro
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
End Class
