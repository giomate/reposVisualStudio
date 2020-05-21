Imports Inventor
Imports System
Imports System.IO
Public Class STLfileGenerator
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D As Sketch3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invFile As InventorFile
    Dim adjuster As SketchAdjust
    Dim barrido As Sweeper
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
    Dim vecino1, vecino2, vecinos(), strCLSID, sabinaName As String

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

        SurfaceBodies = app.TransientObjects.CreateObjectCollection

        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        sections = app.TransientObjects.CreateObjectCollection

        caras = app.TransientObjects.CreateObjectCollection
        SurfaceBodies = app.TransientObjects.CreateObjectCollection
        facesToDelete = app.TransientObjects.CreateFaceCollection
        surfacesSculpt = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)


        nombrador = New Nombres(doku)
        barrido = New Sweeper(doku)
        scaleFactor = 1.01
        sequence = {21, 19, 17, 15, 13, 11, 9, 7, 5, 3, 1, 22, 20, 18, 16, 14, 12, 10, 8, 6, 4, 2, 23}

        DP.Dmax = 200 / 10
        DP.Dmin = 1 / 10
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 11
        DP.q = 23
        DP.b = 25
        strCLSID = "{533E9A98-FC3B-11D4-8E7E-0010B541CD80}"
        done = False
    End Sub

    Public Function MakeAllSTLFiles(doc As PartDocument) As DataMedium
        Dim s1, s2 As String
        Dim dm As DataMedium
        Dim dpc As DerivedPartComponent
        Dim n As Integer
        Try

            doku = DocUpdate(doc)
            sabinaName = doku.FullFileName
            s1 = projectManager.ActiveDesignProject.WorkspacePath
            doku.Close(True)
            If CheckSTLAddIn() Then
                For i = 1 To DP.q
                    s2 = String.Concat(s1, "\Iteration8\stl", i.ToString, ".stl")
                    If System.IO.File.Exists(s2) Then
                        done = True
                    Else
                        dpc = ImportSingleSkeleton(i)
                        If monitor.IsDerivedPartHealthy(dpc) Then
                            If compDef.SurfaceBodies(1).FaceShells.Count > 1 Then
                                doku = barrido.DocUpdate(doku)
                                n = barrido.CutSmallBodies()
                            Else
                                n = 1
                            End If
                            If n = 1 Then
                                If AlignBody() > Math.PI / 2 - 1 / 128 Then
                                    dm = MakeSingleSTL(dpc)
                                    If dm.MediumType = MediumTypeEnum.kDataObjectMedium Then
                                        done = True
                                        doku.Close(True)
                                    Else
                                        done = False
                                        Exit For
                                    End If
                                End If
                            End If


                        End If

                    End If

                Next
            Else
                Return Nothing
            End If
            Return dm
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function AlignBody() As Double
        Dim vBox, vr, v1, v2, v3 As Vector
        Dim oMoveFeature As MoveFeature
        Dim oRotateAboutAxis As RotateAboutLineMoveOperation
        Dim oMoveDef As MoveDefinition
        Dim wa As WorkAxis
        Dim angle, d As Double
        Dim spt1, spt2 As SketchPoint3D
        Dim skl As SketchLine3D
        Try
            Dim sb As SurfaceBody = compDef.SurfaceBodies(1)
            Dim ptMin, ptMax As Point
            ptMin = MoveBodyAtCenter(sb)
            sk3D = compDef.Sketches3D.Add
            Dim vz As Vector = tg.CreateVector(0, 0, 1)

            lamp.FitView(doku)

            For i = 1 To 32
                sb = compDef.SurfaceBodies(1)
                v1 = sb.OrientedMinimumRangeBox.DirectionOne
                v2 = sb.OrientedMinimumRangeBox.DirectionTwo
                v3 = sb.OrientedMinimumRangeBox.DirectionThree
                surfaceBodies.Clear()
                surfaceBodies.Add(sb)
                oMoveDef = compDef.Features.MoveFeatures.CreateMoveDefinition(surfaceBodies)
                If v1.Length > v2.Length Then
                    If v1.Length > v3.Length Then
                        vBox = v1
                    Else
                        vBox = v3

                    End If
                Else
                    If v2.Length > v3.Length Then
                        vBox = v2
                    Else
                        vBox = v3
                    End If

                End If
                ' vBox = v1
                '  vBox.AddVector(v2)
                ' vBox.AddVector(v3)
                ptMin = sb.OrientedMinimumRangeBox.CornerPoint
                spt1 = sk3D.SketchPoints3D.Add(ptMin)
                ptMax = ptMin
                ptMax.TranslateBy(vBox)
                spt2 = sk3D.SketchPoints3D.Add(ptMax)
                skl = sk3D.SketchLines3D.AddByTwoPoints(spt1, spt2, False)

                vr = vz.CrossProduct(vBox)
                wa = compDef.WorkAxes.AddFixed(compDef.WorkPoints(1).Point, vr.AsUnitVector)
                wa.Visible = False
                angle = vz.AngleTo(vBox)
                d = Math.Abs(vz.DotProduct(vBox))
                If d > 1 / 32 Then
                    oRotateAboutAxis = oMoveDef.AddRotateAboutAxis(wa, True, (Math.PI / 2 - angle) / 2)

                    oMoveFeature = compDef.Features.MoveFeatures.Add(oMoveDef)
                    doku.Update2(True)
                    DocUpdate(doku)
                Else
                    Exit For
                End If
                skl.Construction = True
            Next
            sk3D.Visible = False
            doku.Update2(True)
            Return angle
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try


    End Function
    Function MoveBodyAtCenter(sb As SurfaceBody) As Point
        Dim oMoveDef As MoveDefinition
        Try
            Dim ptMin, ptMax As Point
            ptMin = sb.RangeBox.MinPoint
            ptMax = sb.RangeBox.MaxPoint
            Dim ls As LineSegment = tg.CreateLineSegment(ptMin, ptMax)
            surfaceBodies.Clear()
            surfaceBodies.Add(sb)
            oMoveDef = compDef.Features.MoveFeatures.CreateMoveDefinition(surfaceBodies)
            Dim oFreeDrag As FreeDragMoveOperation
            oFreeDrag = oMoveDef.AddFreeDrag(-ls.MidPoint.X, ls.MidPoint.Y, ls.MidPoint.Z)
            Dim oMoveFeature As MoveFeature
            oMoveFeature = compDef.Features.MoveFeatures.Add(oMoveDef)
            doku.Update2(True)
            DocUpdate(doku)
            sb = compDef.SurfaceBodies(1)
            Return sb.RangeBox.MinPoint
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeSingleSTL(dpc As DerivedPartComponent) As DataMedium
        Dim stlTranslator As TranslatorAddIn
        Dim oFileName As String
        Dim oData As DataMedium
        Dim oOptions As NameValueMap
        Dim oContext As TranslationContext
        Try
            stlTranslator = app.ApplicationAddIns.ItemById(strCLSID)

            oContext = app.TransientObjects.CreateTranslationContext

            oOptions = app.TransientObjects.CreateNameValueMap
            'Configure options and write out
            If stlTranslator.HasSaveCopyAsOptions(app.ActiveDocument, oContext, oOptions) Then
                oOptions.Value("Units") = 4   'Inch	'Based off of index in combo box in STL Export Dialog
                ' oOptions.Value("ExportUnits") = 2   'Inch	'Based off of index in combo box in STL Export Dialog
                oOptions.Value("Resolution") = 1    'High	'Based off of index of radio buttons in STL Export Dialog
                oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
                oData = app.TransientObjects.CreateDataMedium

                oFileName = projectManager.ActiveDesignProject.WorkspacePath
                oData.FileName = String.Concat(oFileName, "\Iteration8\stl", qValue.ToString, ".stl")

                stlTranslator.SaveCopyAs(doku, oContext, oOptions, oData)
            Else
                MsgBox("!!!ERROR: STL File was not created!!!")
            End If
            Return oData
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

        Return Nothing
    End Function
    Function ImportSingleSkeleton(si As Integer) As DerivedPartComponent
        Dim p As PartDocument

        Dim dpd As DerivedPartUniformScaleDef
        Dim newComponent As DerivedPartComponent
        Dim dpe As DerivedPartEntity



        Dim skt As Sketch3D
        Dim cc As Integer = 1


        Try
            p = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
            Conversions.SetUnitsToMetric(p)
            dpd = p.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateUniformScaleDef(sabinaName)
            dpd.ExcludeAll()

            dpd.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
            dpd.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIndividualDefined
            dpd.IncludeAllParameters = False
            dpd.IncludeAllSurfaces = False
            dpd.IncludeAllWorkFeatures = DerivedComponentOptionEnum.kDerivedExcludeAll
            dpe = dpd.Solids(si)
            dpe.IncludeEntity = True
            newComponent = p.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(dpd)
            p.Update2(True)
            doku = DocUpdate(p)
            qValue = si


            Return newComponent
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try
        Return p
    End Function
    Function DocUpdate(docu As PartDocument) As PartDocument
        doku = docu
        app = doku.Parent
        compDef = docu.ComponentDefinition
        features = docu.ComponentDefinition.Features
        lamp = New Highlithing(doku)
        monitor = New DesignMonitoring(doku)
        comando = New Commands(app)

        Return doku
    End Function
    Function CheckSTLAddIn() As Boolean
        Dim stlTraslator As TranslatorAddIn
        stlTraslator = app.ApplicationAddIns.ItemById(strCLSID)
        If stlTraslator Is Nothing Then
            MsgBox("Could not access STEP translator.")
            Return False
        Else
            Return True
        End If
        Return False
    End Function
End Class
