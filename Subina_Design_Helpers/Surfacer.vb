
Imports Inventor
Imports Subina_Design_Helpers

Public Class Surfacer
    Public doku, reference As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk As Sketch3D
    Dim comando As Commands
    Public nombrador As Nombres
    Dim monitor As DesignMonitoring

    Public compDef As PartComponentDefinition
    Dim tg As TransientGeometry
    Dim tangentials, sticks, cylinders As ObjectCollection
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)

        projectManager = app.DesignProjectManager

        compDef = doku.ComponentDefinition
        cylinders = app.TransientObjects.CreateObjectCollection

        tangentials = app.TransientObjects.CreateObjectCollection


    End Sub
    Function MakeCylinders() As WorkSurface
        Dim ws As WorkSurface
        Dim sb1 As SurfaceBody
        Dim sb As SurfaceBody = compDef.SurfaceBodies.Item(1)
        Dim np As NonParametricBaseFeature
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition
        Try
            cylinders.Clear()
            For Each f As Face In sb.Faces
                If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                    If f.Evaluator.Area > 0.5 And f.Evaluator.Area < 1 Then
                        cylinders.Add(f)
                    End If
                End If
            Next

            npDef.BRepEntities = cylinders
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
            np.Name = "rodillos"
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sb1 = ws.SurfaceBodies.Item(1)
            sb1.Name = "cylinders"
            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function

    Function MakeTangentials() As WorkSurface
        Dim ws As WorkSurface
        Dim sbtan As SurfaceBody
        Dim sb As SurfaceBody = compDef.SurfaceBodies.Item(1)
        Dim np, npAUx As NonParametricBaseFeature
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition


        Try
            tangentials.Clear()
            For Each f As Face In sb.Faces
                If f.Evaluator.Area > 0.09 And f.Evaluator.Area < 9 Then
                    If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Or f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        Try
                            If f.TangentiallyConnectedFaces.Count > 13 Then
                                tangentials.Add(f)
                            End If
                        Catch ex As Exception

                        End Try
                    End If



                End If

            Next

            npDef.BRepEntities = tangentials
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)

            ' npAUx = CheckSurfaceValidity(np)
            '  If monitor.IsFeatureHealthy(npAUx) Then
            '   np = npAUx
            '   End If

            np.Name = "tangentials"
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sb = ws.SurfaceBodies.Item(1)
            sb.Name = "tangents"

            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Public Function GetTangentials(fi As Face) As WorkSurface
        Dim ws As WorkSurface
        Dim sb As SurfaceBody
        Dim np, npAUx As NonParametricBaseFeature
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition


        Try
            tangentials.Clear()
            tangentials.Add(fi)
            For Each f As Face In fi.TangentiallyConnectedFaces
                tangentials.Add(f)
            Next

            npDef.BRepEntities = tangentials
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)

            ' npAUx = CheckSurfaceValidity(np)
            '  If monitor.IsFeatureHealthy(npAUx) Then
            '   np = npAUx
            '   End If

            np.Name = "tangentials"
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sb = ws.SurfaceBodies.Item(1)
            sb.Name = "tangents"

            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Public Function GetTangentialsNumber(fi As Face, n As String) As WorkSurface
        Dim ws As WorkSurface
        Dim sb As SurfaceBody
        Dim np, npAUx As NonParametricBaseFeature
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition


        Try
            tangentials.Clear()
            tangentials.Add(fi)
            For Each f As Face In fi.TangentiallyConnectedFaces
                tangentials.Add(f)
            Next

            npDef.BRepEntities = tangentials
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)

            ' npAUx = CheckSurfaceValidity(np)
            '  If monitor.IsFeatureHealthy(npAUx) Then
            '   np = npAUx
            '   End If

            np.Name = String.Concat("np", n)
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sb = ws.SurfaceBodies.Item(1)
            sb.Name = String.Concat("ws", n)

            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Public Function CreateSmallVolume(fi As Face, n As String) As WorkSurface
        Dim ws As WorkSurface
        Dim sb As SurfaceBody
        Dim np, npAUx As NonParametricBaseFeature
        Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition


        Try
            tangentials.Clear()
            tangentials.Add(fi)
            For Each f As Face In fi.TangentiallyConnectedFaces
                tangentials.Add(f)
            Next

            npDef.BRepEntities = tangentials
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)

            ' npAUx = CheckSurfaceValidity(np)
            '  If monitor.IsFeatureHealthy(npAUx) Then
            '   np = npAUx
            '   End If

            np.Name = String.Concat("np", n)
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sb = ws.SurfaceBodies.Item(1)
            sb.Name = String.Concat("ws", n)

            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function ReferencedTangentials(ref As PartDocument) As WorkSurface
        Dim ws, wsi As WorkSurface
        Dim np As NonParametricBaseFeature
        Dim sbo As SurfaceBody
        Dim npDef As NonParametricBaseFeatureDefinition

        Try
            npDef = compDef.Features.NonParametricBaseFeatures.CreateDefinition
            wsi = ref.ComponentDefinition.WorkSurfaces.Item("tangents")
            tangentials.Clear()
            For Each sb As SurfaceBody In wsi.SurfaceBodies
                For Each f As Face In sb.Faces
                    tangentials.Add(f)
                Next
            Next


            npDef.BRepEntities = tangentials
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)

            ' npAUx = CheckSurfaceValidity(np)
            '  If monitor.IsFeatureHealthy(npAUx) Then
            '   np = npAUx
            '   End If

            np.Name = "tangentials"
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sbo = ws.SurfaceBodies.Item(1)
            sbo.Name = "tangents"

            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function ReferencedCylinders(ref As PartDocument) As WorkSurface
        Dim ws, wsi As WorkSurface
        Dim np As NonParametricBaseFeature
        Dim sbo As SurfaceBody
        Dim npDef As NonParametricBaseFeatureDefinition

        Try
            npDef = compDef.Features.NonParametricBaseFeatures.CreateDefinition
            wsi = ref.ComponentDefinition.WorkSurfaces.Item("cylinders")
            cylinders.Clear()
            For Each sb As SurfaceBody In wsi.SurfaceBodies
                For Each f As Face In sb.Faces
                    cylinders.Add(f)
                Next
            Next


            npDef.BRepEntities = cylinders
            npDef.OutputType = BaseFeatureOutputTypeEnum.kCompositeOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)

            ' npAUx = CheckSurfaceValidity(np)
            '  If monitor.IsFeatureHealthy(npAUx) Then
            '   np = npAUx
            '   End If

            np.Name = "rodillos"
            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
            sbo = ws.SurfaceBodies.Item(1)
            sbo.Name = "cylinders"

            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function

End Class
