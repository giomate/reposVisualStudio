Imports Inventor
Imports GetInitialConditions
Imports DrawingMainSketch
Imports InitialFolding
Imports Twist7


Public Class TieMaker1
    Public doku As PartDocument
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, connectLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invDoc As InventorFile
    Dim mainSketch As DrawingMainSketch.Sketcher3D
    Dim trobinaCurve As Curves3D

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Dim partNumber, qNext As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Dim nombrador As Nombres
    Dim nextSketch As OriginSketch
    Dim cutProfile As Profile

    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Dim compDef As SheetMetalComponentDefinition
    Dim mainWorkPlane As WorkPlane
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, cutEdge1, cutEdge2, CutEsge3 As Edge
    Dim minorLine, majorLine, cutLine3D, kante3D, tante3D As SketchLine3D
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, twistFace As Face
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex As DimensionConstraint3D
    Dim folded As FoldFeature
    Dim features As SheetMetalFeatures
    Dim lamp As Highlithing

    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim refDoc As FindReferenceLine
    Dim doblez1 As InitialFolding.Doblador
    Dim doblez2 As MicroFold
    Dim doblez3 As MacroFold
    Dim doblez4 As MicroFold4
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        invDoc = New InventorFile(app)


        compDef = doku.ComponentDefinition
        features = compDef.Features
        tg = app.TransientGeometry
        bandLines = app.TransientObjects.CreateObjectCollection
        constructionLines = app.TransientObjects.CreateObjectCollection
        sections = app.TransientObjects.CreateObjectCollection
        esquinas = app.TransientObjects.CreateObjectCollection
        rails = app.TransientObjects.CreateObjectCollection
        lamp = New Highlithing(doku)
        thicknessCM = compDef.Thickness._Value
        gap1CM = 3 / 10
        refDoc = New FindReferenceLine(doku)
        nombrador = New Nombres(doku)


        done = False
    End Sub
    Public Function MakeNextTie() As PartDocument
        Try
            If GetInitialConditions() > 0 Then
                If doblez1.MakeFirstFold(refDoc, qNext) Then
                    doblez2 = New MicroFold(doblez1.doku)
                    If doblez2.MakeSecondFold() Then
                        doblez3 = New MacroFold(doblez2.doku)
                        If doblez3.MakeThirdFold() Then
                            doblez4 = New MicroFold4(doblez3.doku)
                            If doblez4.MakeForthFold() Then

                            End If

                        End If

                    End If
                End If
            End If

            Return Nothing
        Catch ex As Exception

            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetInitialConditions() As Integer
        Try
            Dim q As Integer = 0
            doblez1 = New InitialFolding.Doblador(invDoc.CreateSheetMetalFile(nombrador.MakeNextFileName(refDoc.oDoc)))
            doku = doblez1.doku
            trobinaCurve = New Curves3D(doblez1.doku)
            trobinaCurve.DefineTrobinaParameters(doku)
            q = nombrador.GetQNumber(doku)
            qNext = q
            Return q
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
End Class
