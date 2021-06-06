Imports Inventor
Imports System
Imports System.IO

Imports System.Text
Imports System.IO.Directory
Imports Subina_Design_Helpers
Public Class SubinaStruct
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application

    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile
    Dim adjuster As SketchAdjust

    Dim scaleFactor As Double


    Public wp1, wp2, wp3, wpConverge As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint, convergePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double

    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres

    Dim caraTrabajo As Surfacer
    Dim cylinderFace As Face



    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Public compDef As PartComponentDefinition


    Dim features As PartFeatures
    Dim lamp As Highlithing
    Dim directoryName As System.IO.DirectoryInfo
    Dim archivo As System.IO.File
    Dim pahtFile As System.IO.Path
    Dim bandFaces As WorkSurface
    Dim qValue1, qValue2, qValue, qVecinos(), sequence() As Integer
    Dim vecino1, vecino2, vecinos() As String

    Dim sections, caras, surfaceBodies, facesToDelete, surfacesSculpt As ObjectCollection

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
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        invFile = New InventorFile(app)
        projectManager = app.DesignProjectManager

        compDef = doku.ComponentDefinition

        surfaceBodies = app.TransientObjects.CreateObjectCollection

        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        sections = app.TransientObjects.CreateObjectCollection

        caras = app.TransientObjects.CreateObjectCollection
        surfaceBodies = app.TransientObjects.CreateObjectCollection
        facesToDelete = app.TransientObjects.CreateFaceCollection
        surfacesSculpt = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)


        nombrador = New Nombres(doku)

        scaleFactor = 1.01
        sequence = {21, 19, 17, 15, 13, 11, 9, 7, 5, 3, 1, 22, 20, 18, 16, 14, 12, 10, 8, 6, 4, 2, 23}

        DP.Dmax = 200 / 10
        DP.Dmin = 1 / 10
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 11
        DP.q = 23
        DP.b = 25
        done = False
    End Sub
    Public Function MakeNestStruct(i As Integer) As PartDocument
        Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim q As Integer
        Dim dn, rn, bn, sn As String
        dn = String.Concat(p, "\Iteration", i.ToString)
        Dim dpc As DerivedPartComponent
        Dim t As PartDocument
        Dim ws As WorkSurface
        Dim partCounter As Integer = 0
        Dim cf As CombineFeature

        Try
            If (Directory.Exists(dn)) Then
                fullFileNames = Directory.GetFiles(dn, "*.ipt")
                For i = 1 To DP.q
                    rn = nombrador.MakeRibFileName(dn, i)
                    If System.IO.File.Exists(rn) Then
                        partCounter += 1
                    Else
                        sn = nombrador.ConstructFileName(dn, "Skeleton", i)
                        If System.IO.File.Exists(sn) Then
                            If MakeCollidedRib(sn) = 1 Then
                                partCounter += 1

                            End If
                        End If

                    End If
                Next
                If partCounter = DP.q Then
                    t = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
                    Conversions.SetUnitsToMetric(t)
                    partCounter = 0
                    t.Update2(True)
                    doku = DocUpdate(t)
                    fullFileNames = Directory.GetFiles(dn, "*.ipt")
                    For i = 1 To DP.q
                        rn = nombrador.MakeRibFileName(dn, i)
                        If System.IO.File.Exists(rn) Then
                            dpc = DeriveRibBand(rn)
                            lamp.FitView(doku)
                            If monitor.IsDerivedPartHealthy(dpc) Then
                                partCounter += 1
                            Else
                                Exit For
                            End If
                        End If

                    Next


                End If
                If partCounter = DP.q Then
                    cf = CombineBodies()
                    If monitor.IsFeatureHealthy(cf) Then
                        SaveAsSilent(String.Concat(dn, "\Subina.ipt"))
                        done = True
                    End If
                Else
                    done = False
                End If

            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return doku
    End Function
    Public Function MakeSimpleNestStruct(i As Integer) As PartDocument
        Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim q As Integer
        Dim dn, rn, bn, sn As String
        dn = String.Concat(p, "\Iteration", i.ToString)
        Dim dpc As DerivedPartComponent
        Dim t As PartDocument
        Dim ws As WorkSurface
        Dim partCounter As Integer = 0
        Dim cf As CombineFeature

        Try
            If (Directory.Exists(dn)) Then
                fullFileNames = Directory.GetFiles(dn, "*.ipt")
                rn = nombrador.MakeSkeletonFileName(dn, 1)
                If System.IO.File.Exists(rn) Then
                    t = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
                    Conversions.SetUnitsToMetric(t)
                    partCounter = 0
                    t.Update2(True)
                    doku = DocUpdate(t)
                    fullFileNames = Directory.GetFiles(dn, "*.ipt")
                    For i = 1 To DP.q
                        rn = nombrador.MakeSkeletonFileName(dn, i)
                        If System.IO.File.Exists(rn) Then
                            dpc = DeriveRibBand(rn)
                            lamp.FitView(doku)
                            If monitor.IsDerivedPartHealthy(dpc) Then
                                partCounter += 1
                            Else
                                Exit For
                            End If
                        End If

                    Next


                End If
                If partCounter = DP.q Then
                    cf = CombineBodies()
                    If monitor.IsFeatureHealthy(cf) Then
                        SaveAsSilent(String.Concat(dn, "\Subina.ipt"))
                        done = True
                    End If
                Else
                    done = False
                End If

            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return doku
    End Function
    Public Function MakeCondesatorStruct(i As Integer) As PartDocument
        Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim q As Integer
        Dim dn, rn, bn, sn As String
        dn = String.Concat(p, "\Iteration", i.ToString)
        Dim dpc As DerivedPartComponent
        Dim t As PartDocument
        Dim ws As WorkSurface
        Dim partCounter As Integer = 0
        Dim cf As CombineFeature
        Dim ffn As String = doku.FullFileName

        Try
            If (Directory.Exists(dn)) Then
                t = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
                Conversions.SetUnitsToMetric(t)
                partCounter = 0
                t.Update2(True)
                doku = DocUpdate(t)
                fullFileNames = Directory.GetFiles(dn, "*.ipt")
                For i = 1 To DP.q
                    bn = nombrador.MakeBandFileName(ffn, i)
                    If System.IO.File.Exists(bn) Then
                        dpc = DeriveBand(bn)
                        lamp.FitView(doku)
                        If monitor.IsDerivedPartHealthy(dpc) Then
                            partCounter += 1
                            done = True
                        Else
                            done = False
                            Exit For
                        End If
                    End If

                Next
                If partCounter = DP.q Then
                    SaveAsSilent(String.Concat(dn, "\Condesator1.ipt"))
                    done = True
                Else
                    done = False
                End If

            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return doku
    End Function
    Public Function AssemblyNest(i As Integer) As PartDocument
        Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim q As Integer
        Dim dn, rn, bn, sn As String
        dn = String.Concat(p, "\Iteration", i.ToString)
        Dim dpc As DerivedPartComponent
        Dim t As PartDocument
        Dim ws As WorkSurface
        Dim partCounter As Integer = 0

        Try
            If (Directory.Exists(dn)) Then
                fullFileNames = Directory.GetFiles(dn, "*.ipt")

                t = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
                Conversions.SetUnitsToMetric(t)
                partCounter = 0
                t.Update2(True)
                doku = DocUpdate(t)
                fullFileNames = Directory.GetFiles(dn, "*.ipt")
                For i = 1 To DP.q
                    rn = nombrador.MakeRibFileName(dn, i)
                    If System.IO.File.Exists(rn) Then
                        dpc = DeriveRibBand(rn)
                        lamp.FitView(doku)
                        If monitor.IsDerivedPartHealthy(dpc) Then
                            partCounter += 1
                        End If
                    Else
                        Exit For
                    End If

                Next


                If partCounter = DP.q Then
                    SaveAsSilent(String.Concat(dn, "\Subina.ipt"))
                    done = True
                Else
                    done = False
                End If

            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return doku
    End Function
    Public Function RemoveFaceShells(i As Integer) As PartDocument
        Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim q As Integer
        Dim dn, rn, bn, sn As String
        dn = String.Concat(p, "\Iteration", i.ToString)
        Dim dpc As DerivedPartComponent
        Dim t As PartDocument
        Dim ws As WorkSurface
        Dim partCounter As Integer = 0
        Dim sb As SurfaceBody

        Try
            If (Directory.Exists(dn)) Then
                fullFileNames = Directory.GetFiles(dn, "*.ipt")
                For i = 1 To DP.q
                    sn = nombrador.ConstructFileName(dn, "Skeleton", i)
                    If System.IO.File.Exists(sn) Then
                        t = app.Documents.Open(sn)
                        doku = DocUpdate(t)
                        sb = CutSmallBodies()
                        If compDef.SurfaceBodies.Count < 2 Then
                            SaveSilent()
                            DokuClose()
                        End If

                    End If
                Next

                done = True
            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return doku
    End Function
    Function DeriveRib(rn As String) As DerivedPartComponent
        Dim dpc As DerivedPartComponent
        Dim ws As WorkSurface
        Try
            dpc = DeriveSinglePart(rn)
            qValue = nombrador.GetQNumberString(rn, "Rib")
            dpc.SurfaceBodies(1).Name = String.Concat("Rib", qValue)
            If monitor.IsDerivedPartHealthy(dpc) Then
                ws = MakeBandSurface(nombrador.MakeBandFileName(rn))
                If ws.SurfaceBodies.Count > 1 Then
                    Return dpc
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Public Function MakeRib(sn As String) As Integer
        Dim pf As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim p As PartDocument
        Dim q As Integer
        Dim dpc As DerivedPartComponent
        Dim newComponent As DerivedPartComponent

        Dim lf As LoftFeature
        Dim cf As CombineFeature
        Dim bodyCounter As Integer
        Dim d, e As Double
        Dim skt As Sketch3D
        Dim cc As Integer = 1
        Dim sn1, sn2, rn As String
        Dim solitario As Boolean = True



        Try
            p = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
            Conversions.SetUnitsToMetric(p)
            p.Update2(True)
            doku = DocUpdate(p)
            qValue = nombrador.GetQNumberString(sn, "Skeleton")
            sn1 = GetNeigboorValues(sn)

            dpc = DeriveSinglePart(sn)
            If monitor.IsDerivedPartHealthy(dpc) Then
                lamp.FitView(doku)
                For i = 0 To vecinos.Length - 1
                    If System.IO.File.Exists(vecinos(i)) Then
                        Try
                            If monitor.IsFeatureHealthy(DeriveSingleNeigboor(sn, i)) Then
                                solitario = False
                            End If
                        Catch ex As Exception

                        End Try



                    End If

                Next
                If Not solitario Then

                    Return IntersectCollisions(sn)


                Else
                    Return DeriveNormalSkeleton(sn)
                End If
            End If







        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function MakeCollidedRib(sn As String) As Integer
        Dim pf As String = projectManager.ActiveDesignProject.WorkspacePath
        Dim p As PartDocument
        Dim q As Integer
        Dim dpc As DerivedPartComponent
        Dim newComponent As DerivedPartComponent

        Dim lf As LoftFeature
        Dim cf As CombineFeature
        Dim bodyCounter As Integer
        Dim d, e As Double
        Dim skt As Sketch3D
        Dim nFS As Integer
        Dim sn1, sn2, rn As String
        Dim solitario As Boolean = True
        Try
            p = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
            Conversions.SetUnitsToMetric(p)
            p.Update2(True)
            doku = DocUpdate(p)
            qValue = nombrador.GetQNumberString(sn, "Skeleton")
            sn1 = GetNeigboorValues(sn)

            dpc = DeriveSinglePart(sn)
            If monitor.IsDerivedPartHealthy(dpc) Then
                lamp.FitView(doku)

                nFS = IntersectCollisions(sn)
                    If nFS < 2 Then
                        Return nFS
                    End If

            End If





            Return 0

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function IntersectCollisions(sn As String) As Integer
        Dim dpc As DerivedPartComponent
        Dim cf As CombineFeature
        Dim bodyCounter As Integer
        Try


    
                If compDef.SurfaceBodies(1).FaceShells.Count > 1 Then
                    CutSmallBodies()
                End If

                DocUpdate(doku)
                IntersectCollisions = compDef.SurfaceBodies(1).FaceShells.Count
                If IntersectCollisions < 2 Then
                    SaveAsSilent(nombrador.MakeRibFileName(sn))
                    DokuClose()
                    Return IntersectCollisions
                End If


        Catch ex As Exception

        End Try

        Return 0
    End Function
    Function CutFloatingBodies(sn As String) As Integer

        Try


            If compDef.SurfaceBodies(1).FaceShells.Count > 1 Then
                    CutSmallBodies()
                End If

            DocUpdate(doku)
            CutFloatingBodies = compDef.SurfaceBodies(1).FaceShells.Count
            Return CutFloatingBodies

        Catch ex As Exception

        End Try

        Return 0
    End Function
    Function DeriveSecondNeigboor(sn As String) As Integer
        Dim dpc As DerivedPartComponent

        Dim cf As CombineFeature
        Dim bodyCounter As Integer
        Dim d, e As Double
        Try
            doku.Update2(True)


            bodyCounter = compDef.SurfaceBodies.Count
            If monitor.IsDerivedPartHealthy(DeriveSingleSkeletonScaled(vecino2)) Then
                lamp.FitView(doku)
                cf = CutTwoBodies(compDef.SurfaceBodies(bodyCounter + 1), compDef.SurfaceBodies(1), False)
                If monitor.IsFeatureHealthy(cf) Then
                    dpc = DeriveSinglePart(sn)
                    If monitor.IsDerivedPartHealthy(dpc) Then
                        cf = RemoveCollisionBodies(compDef.SurfaceBodies(compDef.SurfaceBodies.Count))
                        lamp.FitView(doku)
                        If monitor.IsFeatureHealthy(cf) Then
                            CutSmallBodies()
                            DocUpdate(doku)
                            DeriveSecondNeigboor = compDef.SurfaceBodies(1).FaceShells.Count
                            If DeriveSecondNeigboor < 2 Then
                                SaveAsSilent(nombrador.MakeRibFileName(sn))
                                DokuClose()
                                Return DeriveSecondNeigboor
                            End If

                        End If
                    End If

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


        Return 0
    End Function
    Function DeriveSingleNeigboor(sn As String, i As Integer) As CombineFeature
        Dim dpc As DerivedPartComponent

        Dim cf As CombineFeature
        Dim bodyCounter As Integer
        Dim d, e As Double
        Try

            If System.IO.File.Exists(vecinos(i)) Then
                doku.Update2(True)
                DocUpdate(doku)
                bodyCounter = compDef.SurfaceBodies.Count
                dpc = DeriveSingleSkeletonScaled(vecinos(i))
                If monitor.IsDerivedPartHealthy(dpc) Then
                    lamp.FitView(doku)
                    Try
                        cf = CutTwoBodies(compDef.SurfaceBodies(1), compDef.SurfaceBodies(bodyCounter + 1), False)
                        If monitor.IsFeatureHealthy(cf) Then
                            Return cf
                        Else
                            cf.Delete()
                            dpc.Delete()
                            Return Nothing
                        End If
                    Catch ex As Exception

                    End Try

                End If
            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try



    End Function
    Function DeriveNormalSkeleton(sn) As Integer
        Dim dpc As DerivedPartComponent
        Try

            DocUpdate(doku)
            DeriveNormalSkeleton = compDef.SurfaceBodies(1).FaceShells.Count
            lamp.FitView(doku)
            If DeriveNormalSkeleton < 2 Then
                SaveAsSilent(nombrador.MakeRibFileName(sn))
                DokuClose()
                Return DeriveNormalSkeleton
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DokuClose() As Integer
        SaveSilent()
        doku.Close()
        app.Documents.CloseAll()
        Return 0
    End Function
    Function SaveSilent() As PartDocument
        app.SilentOperation = True
        doku.Save()

        app.SilentOperation = False
        Return doku
    End Function
    Function SaveAsSilent(rn As String) As PartDocument
        app.SilentOperation = True
        doku.SaveAs(rn, False)
        doku.Save()
        app.SilentOperation = False
        Return doku
    End Function
    Function FindRib() As PartDocument
        Try

        Catch ex As Exception

        End Try

    End Function
    Function Removeintersections(s As String) As CombineFeature
        Try

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function SingleItersection(sv As String) As Integer
        Dim dpc As DerivedPartComponent
        Dim cf As CombineFeature
        Try
            dpc = DeriveSinglePart(sv)
            If monitor.IsDerivedPartHealthy(dpc) Then
                cf = CutTwoBodies(dpc.SurfaceBodies(1), compDef.SurfaceBodies(qValue), True)
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CutTwoBodies(sbb As SurfaceBody, sbt As SurfaceBody, keep As Boolean) As CombineFeature

        Dim cf As CombineFeature
        Dim factor As Double
        Try
            surfaceBodies.Clear()

            surfaceBodies.Add(sbt)
            If doku.ComponentDefinition.SurfaceBodies.Count > 1 Then
                Try
                    cf = compDef.Features.CombineFeatures.Add(sbb, surfaceBodies, PartFeatureOperationEnum.kCutOperation, keep)
                Catch ex As Exception

                    MsgBox(ex.ToString())
                    Return Nothing

                End Try

            End If
            ' cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)
            Return cf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function RemoveCollisionBodies(sbb As SurfaceBody) As CombineFeature

        Dim cf As CombineFeature

        Try
            surfaceBodies.Clear()
            For i = 2 To compDef.SurfaceBodies.Count
                surfaceBodies.Add(compDef.SurfaceBodies(i))
            Next

            If doku.ComponentDefinition.SurfaceBodies.Count > 1 Then
                Try
                    cf = compDef.Features.CombineFeatures.Add(sbb, surfaceBodies, PartFeatureOperationEnum.kCutOperation, True)
                Catch ex As Exception

                    MsgBox(ex.ToString())
                    Return Nothing

                End Try

            End If
            ' cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)
            Return cf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function DeriveRibBand(rn As String) As DerivedPartComponent
        Dim dpc As DerivedPartComponent = DeriveSinglePart(rn)
        Dim ws As WorkSurface
        If monitor.IsDerivedPartHealthy(dpc) Then
            ws = MakeBandSurface(rn)
            Return dpc
        End If
        Return Nothing
    End Function
    Function DeriveBand(bn As String) As DerivedPartComponent
        Dim dpc As DerivedPartComponent = DeriveSinglePart(bn)
        Dim ws As WorkSurface
        Dim q As Integer
        If monitor.IsDerivedPartHealthy(dpc) Then
            doku.Update2(True)
            caraTrabajo = New Surfacer(doku)
            q = nombrador.GetQNumberString(bn)
            ws = caraTrabajo.GetTangentialsNumber(GetInitialFace(compDef.SurfaceBodies.Item(compDef.SurfaceBodies.Count)), q.ToString)
            ws.Visible = False
            Return dpc
        End If

        Return Nothing
    End Function
    Function CutSmallBodies() As SurfaceBody
        Try
            Dim vMax As Double = 0
            Dim sbMax As SurfaceBody = compDef.SurfaceBodies(1)
            Dim fsMax As FaceShell = sbMax.FaceShells(1)
            Dim sbTemp As SurfaceBody
            Dim ws As WorkSurface
            Dim sf As SculptFeature
            Dim nFaceMax As Integer = 0
            Dim nFaceShells, k As Integer
            Dim fs As FaceShell
            k = 1
            Dim v As Double
            doku.Update2(True)
            If sbMax.FaceShells.Count > 1 Then

                nFaceShells = doku.ComponentDefinition.SurfaceBodies(1).FaceShells.Count
                For i = 1 To nFaceShells * 2

                    fs = doku.ComponentDefinition.SurfaceBodies(1).FaceShells(k)
                    fsMax = GetLargerFaceShell()
                    If Not fs.Equals(fsMax) Then
                        facesToDelete.Clear()
                        For Each f As Face In fs.Faces
                            lamp.HighLighFace(f)
                            facesToDelete.Add(f)
                        Next
                        If monitor.IsFeatureHealthy(CreateNonParametricBody()) Then
                            ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)

                            sf = SculptRemove(ws)
                            If monitor.IsFeatureHealthy(sf) Then
                                doku.Update2(True)
                                nFaceShells = doku.ComponentDefinition.SurfaceBodies(1).FaceShells.Count
                                If nFaceShells < 2 Then
                                    Exit For
                                Else
                                    k = 1
                                    DocUpdate(doku)
                                End If
                            Else
                                MsgBox(sf.ToString)
                                Return Nothing
                            End If

                        End If
                    Else
                        k = k + 1

                    End If
                    If doku.ComponentDefinition.SurfaceBodies(1).FaceShells.Count < 2 Then
                        Exit For
                    End If
                Next



            End If



            Return doku.ComponentDefinition.SurfaceBodies(1)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetLargerFaceShell() As FaceShell
        Dim fsMax As FaceShell = doku.ComponentDefinition.SurfaceBodies(1).FaceShells(1)

        Dim nFaceMax As Integer = 0

        For Each fsi As FaceShell In doku.ComponentDefinition.SurfaceBodies(1).FaceShells
            If fsi.Faces.Count > nFaceMax Then
                nFaceMax = fsi.Faces.Count
                fsMax = fsi
            End If
        Next
        Return fsMax
    End Function
    Function LeaveBiggerPart(sn As String, cf As CombineFeature) As SurfaceBody
        Try
            Dim vMax As Double = 0
            Dim sbMax As SurfaceBody = compDef.SurfaceBodies(1)
            Dim fsMax As FaceShell = sbMax.FaceShells(1)
            Dim sbTemp As SurfaceBody
            Dim ws As WorkSurface
            Dim sf As SculptFeature

            Dim v As Double
            If sbMax.FaceShells.Count > 1 Then
                For Each fs As FaceShell In sbMax.FaceShells
                    ' If Not fs.Equals(fsMax) Then
                    facesToDelete.Clear()
                    For Each f As Face In fs.Faces
                        lamp.HighLighFace(f)
                        facesToDelete.Add(f)
                    Next
                    If monitor.IsFeatureHealthy(CreateNonParametricBody()) Then
                        ws = compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)
                        ' sf = SculptRemove(ws)
                        ' If monitor.IsFeatureHealthy(sf) Then

                        'Else
                        '  MsgBox(sf.ToString)
                        '  Return Nothing
                        '  End If
                    End If



                    'End If
                Next

            End If



            Return sbMax
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CreateNonParametricBody() As NonParametricBaseFeature
        Dim np As NonParametricBaseFeature
        Try
            Dim npDef As NonParametricBaseFeatureDefinition = compDef.Features.NonParametricBaseFeatures.CreateDefinition

            npDef.BRepEntities = facesToDelete
            npDef.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType
            npDef.IsAssociative = False
            np = compDef.Features.NonParametricBaseFeatures.AddByDefinition(npDef)
            Return np
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function SculptRemove(wsi As WorkSurface) As SculptFeature
        Dim sf As SculptFeature
        Dim ss As SculptSurface

        Try
            surfacesSculpt.Clear()
            ss = compDef.Features.SculptFeatures.CreateSculptSurface(wsi, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            surfacesSculpt.Add(ss)
            Try
                sf = compDef.Features.SculptFeatures.Add(surfacesSculpt, PartFeatureOperationEnum.kCutOperation)
                If Not monitor.IsFeatureHealthy(sf) Then
                    sf = CorrectSculpt(sf, wsi)

                End If
            Catch ex As Exception
                sf = compDef.Features.SculptFeatures.Add(surfacesSculpt, PartFeatureOperationEnum.kNewBodyOperation)
                If monitor.IsFeatureHealthy(CombineBodies()) Then
                    Return sf
                Else
                    MsgBox(ex.ToString())
                    Return Nothing
                End If
            End Try


            Return sf
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function CorrectSculpt(sf As SculptFeature, wsi As WorkSurface) As SculptFeature
        Dim ss As SculptSurface
        Try
            sf.Delete(True, True, True)
            CorrectSculpt = compDef.Features.SculptFeatures.Add(surfacesSculpt, PartFeatureOperationEnum.kNewBodyOperation)
            If monitor.IsFeatureHealthy(CorrectSculpt) Then
                If monitor.IsFeatureHealthy(CombineBodies()) Then
                    Return CorrectSculpt
                Else
                    MsgBox(sf.ToString)

                    Return Nothing
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return Nothing
    End Function
    Function CombineBodies() As CombineFeature
        Dim cf As CombineFeature
        If compDef.SurfaceBodies.Count > 1 Then
            surfaceBodies.Clear()

            For index = 2 To doku.ComponentDefinition.SurfaceBodies.Count
                surfaceBodies.Add(compDef.SurfaceBodies.Item(index))
            Next
            Try
                cf = doku.ComponentDefinition.Features.CombineFeatures.Add(compDef.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)

            Catch ex As Exception


            End Try
        End If

        ' cf = doku.ComponentDefinition.Features.CombineFeatures.Add(doku.ComponentDefinition.SurfaceBodies.Item(1), surfaceBodies, PartFeatureOperationEnum.kJoinOperation)

        Return cf
    End Function
    Function GetNeighboorNames(si As String) As String
        Dim qi As Integer = nombrador.GetQNumberString(si, "Skeleton")
        Dim qm, j As Integer
        Dim pathName As String = projectManager.ActiveDesignProject.WorkspacePath
        ReDim qVecinos(3)
        ReDim vecinos(3)
        Dim s() As String = Strings.Split(si, "Skeleton")

        Math.DivRem(qValue1 + 11, 23, qm)
        Dim ri, rq, rMin1, rMin2, r, rMins(3) As Double
        For i = 0 To rMins.Length - 1
            rMins(i) = 9999
            qVecinos(i) = i
        Next
        rMin1 = 99999
        rMin2 = rMin1
        rq = Math.IEEERemainder((qi * DP.p / DP.q) * 2 * Math.PI, (2 * Math.PI))
        'rq = Math.Min(Math.Abs(rq), Math.Abs(2 * Math.PI - rq))
        For i = 1 To DP.q
            Math.DivRem(CInt(qi + i - 1), DP.q, j)
            j = j + 1
            If Not j = qi Then
                ri = Math.IEEERemainder((j * DP.p / DP.q) * 2 * Math.PI, (2 * Math.PI))
                ' ri = Math.Min(Math.Abs(ri), Math.Abs(2 * Math.PI - ri))
                r = Math.Min(Math.Abs(rq - ri), Math.Abs(ri - rq))
                For n = 0 To rMins.Length - 1
                    If r < rMins(n) + 1 / 128 And r > Math.PI / DP.q Then
                        If n < 3 Then
                            If r < rMins(n + 1) + 1 / 128 And r > Math.PI / DP.q Then
                            Else
                                If n > 0 Then
                                    For k = 0 To n - 1
                                        rMins(k) = rMins(k + 1)
                                        qVecinos(k) = qVecinos(k + 1)
                                    Next
                                End If

                                rMins(n) = r
                                qVecinos(n) = j
                            End If
                        Else
                            If n > 0 Then
                                For k = 0 To n - 1
                                    rMins(k) = rMins(k + 1)
                                    qVecinos(k) = qVecinos(k + 1)
                                Next
                            End If
                            rMins(n) = r
                            qVecinos(n) = j

                        End If
                    End If
                Next

            End If

        Next
        For m = 0 To rMins.Length - 1
            vecinos(m) = String.Concat(s(0), "Rib", qVecinos(m).ToString, ".ipt")
        Next

        Return vecinos(0)
    End Function
    Function GetNeigboorValues(si As String) As String
        Dim qi As Integer = nombrador.GetQNumberString(si, "Skeleton")
        Dim qm, j, k, sign As Integer
        Dim pathName As String = projectManager.ActiveDesignProject.WorkspacePath
        ReDim qVecinos(3)
        ReDim vecinos(3)
        Dim s() As String = Strings.Split(si, "Skeleton")


        For i = 1 To qVecinos.Length
            sign = CInt(Math.Pow(-1, i + 1))
            If sign > 0 Then
                Math.DivRem(qi + i + 23, 23, j)
                j = j + 1
            Else
                Math.DivRem(qi - i + 22, 23, j)
                j = j + 1
            End If
            qVecinos(i - 1) = j

        Next

        For m = 0 To qVecinos.Length - 1
            vecinos(m) = String.Concat(s(0), "Rib", qVecinos(m).ToString, ".ipt")
        Next

        Return vecinos(0)
    End Function
    Function DocUpdate(docu As PartDocument) As PartDocument
        doku = docu

        compDef = docu.ComponentDefinition
        features = docu.ComponentDefinition.Features
        lamp = New Highlithing(doku)
        monitor = New DesignMonitoring(doku)
        'adjuster = New SketchAdjust(doku)

        Return doku
    End Function
    Public Function DeriveSinglePart(s As String) As DerivedPartComponent
        Dim dpd As DerivedPartDefinition
        Dim dpc As DerivedPartComponent
        Try
            dpd = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
            dpd.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
            ' derivedDefinition.IncludeAllSurfaces = DerivedComponentOptionEnum.kDerivedIncludeAll
            dpd.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
            dpc = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(dpd)
            Return dpc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Public Function DeriveSingleSkeletonScaled(s As String) As DerivedPartComponent
        Dim derivedDefinition As DerivedPartUniformScaleDef
        Dim dpc As DerivedPartComponent
        Try
            derivedDefinition = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateUniformScaleDef(s)
            derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
            derivedDefinition.ScaleFactor = scaleFactor
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
            dpc = doku.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)
            Return dpc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function MakeBandSurface(s As String) As WorkSurface
        Dim bn As String = nombrador.MakeBandFileName(s)
        Dim dpd As DerivedPartDefinition
        Dim ws As WorkSurface
        Dim q As Integer
        Try
            dpd = compDef.ReferenceComponents.DerivedPartComponents.CreateDefinition(bn)
            dpd.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsWorkSurface
            Dim dpc As DerivedPartComponent = compDef.ReferenceComponents.DerivedPartComponents.Add(dpd)
            doku.Update2(True)
            caraTrabajo = New Surfacer(doku)
            compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count).Visible = False
            q = nombrador.GetQNumberString(bn)
            ws = caraTrabajo.GetTangentialsNumber(GetInitialFace(compDef.WorkSurfaces.Item(compDef.WorkSurfaces.Count)), q.ToString)
            ws.Visible = False
            Return ws
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function GetInitialFace(ws As WorkSurface) As Face

        Dim ptc As Point

        Dim min2, min1 As Double

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
            lamp.HighLighFace(fmin1)
            cylinderFace = fmin1


            Return cylinderFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
    Function GetInitialFace(sbi As SurfaceBody) As Face

        Dim ptc As Point

        Dim min2, min1 As Double

        Dim fmin1, fmin2 As Face
        Try
            ptc = doku.ComponentDefinition.WorkPoints.Item(1).Point
            caras.Clear()
            min1 = sbi.Faces.Item(1).Evaluator.Area
            min2 = min1
            fmin1 = sbi.Faces.Item(1)
            fmin2 = fmin1

            For Each fc As Face In sbi.Faces
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

            lamp.HighLighFace(fmin1)
            cylinderFace = fmin1


            Return cylinderFace
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
    End Function
End Class
