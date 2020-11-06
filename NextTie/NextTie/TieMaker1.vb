Imports Inventor

Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory

Public Class TieMaker1
    Public doku As PartDocument
    Public projectManager As DesignProjectManager
    Dim app As Application
    Dim sk3D, refSk As Sketch3D

    Dim refLine, firstLine, secondLine, thirdLine, lastLine, connectLine As SketchLine3D
    Dim curve, refCurve As SketchEquationCurve3D
    Public done, healthy As Boolean

    Dim monitor As DesignMonitoring
    Dim invDoc As InventorFile

    Public trobinaCurve As Curves3D

    Public wp1, wp2, wp3 As WorkPoint
    Public farPoint, point1, point2, point3, curvePoint As Point
    Dim tg As TransientGeometry
    Dim gap1CM, thicknessCM As Double
    Public partNumber, qNext, qLastTie, direction As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Dim nextSketch As OriginSketch
    Dim cutProfile As Profile

    Dim pro As Profile

    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Public compDef As SheetMetalComponentDefinition
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim mainWorkPlane As WorkPlane
    Dim minorEdge, majorEdge, bendEdge, adjacentEdge, cutEdge1, cutEdge2, CutEsge3 As Edge
    Dim minorLine, majorLine, cutLine3D, kante3D, tante3D As SketchLine3D
    Dim workFace, adjacentFace, bendFace, frontBendFace, cutFace, twistFace As Face
    Dim bendAngle As DimensionConstraint
    Dim gapFold, gapVertex As DimensionConstraint3D
    Dim folded As FoldFeature
    Public foldFeatures As FoldFeatures
    Dim features As SheetMetalFeatures
    Dim lamp As Highlithing
    Dim di As System.IO.DirectoryInfo
    Dim fi As System.IO.File
    Dim nf As System.IO.Path

    Dim foldFeature As FoldFeature
    Dim sections, esquinas, rails As ObjectCollection

    Dim edgeColl As EdgeCollection
    Dim twistPlane As WorkPlane
    Dim refDoc As FindReferenceLine
    Dim doblez1 As InitFold
    Dim doblez2 As MicroFold6
    Dim doblez3 As MacroFold5
    Dim doblez4 As MicroFold6
    Dim doblez5 As MacroFold5
    Dim doblez6 As MicroFold6
    Dim doblez7 As MacroFold5
    Dim doblez8 As MicroFold6
    Dim manager As FoldingEvaluator
    Dim giro As TwistFold7
    Dim arrayFunctions As Collection
    Dim fullFileNames As String()
    Public fullFileName As String
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        invDoc = New InventorFile(app)
        projectManager = app.DesignProjectManager

        compDef = doku.ComponentDefinition
        sheetMetalFeatures = compDef.Features
        foldFeatures = sheetMetalFeatures.FoldFeatures
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
        manager = New FoldingEvaluator(doku)
        trobinaCurve = New Curves3D(doku)
        direction = -1

        done = False
    End Sub
    Public Function MakeNextTie() As PartDocument
        Try
            Dim i As Integer
            'Dim fp As FlatPattern
            i = GetStartingFeature()
            doku.Update2(True)

            If i < 1 Then
                If GetInitialConditions() > 0 Then

                    If doblez1.MakeFirstFold(refDoc, qNext, direction) Then
                        doku = doblez1.doku

                        doku.Activate()
                        MakeRestTie(1)

                    End If
                End If
            Else
                If sheetMetalFeatures.CutFeatures.Count > 0 Then
                    cutfeature = sheetMetalFeatures.CutFeatures.Item(1)
                    If cutfeature.Name = "initCut" Then
                        MakeRestTie(i)
                    Else
                        If monitor.IsFeatureHealthy(doblez1.MakeInitCut()) Then
                            MakeRestTie(1)
                        End If
                    End If

                Else
                    If monitor.IsFeatureHealthy(doblez1.MakeInitCut()) Then
                        MakeRestTie(1)
                    End If

                End If

            End If
            doku.Update()
            doku.Update()
            compDef = doku.ComponentDefinition
            If compDef.Sketches3D.Item("introLine").SketchLines3D.Count > 0 Then
                ' fp = compDef.FlatPattern
                If compDef.HasFlatPattern Then
                    done = True
                    fullFileName = doku.FullFileName
                    Return doku
                Else
                    Return Nothing

                End If

            End If


        Catch ex As Exception

            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeRestTie(i As Integer) As PartDocument
        Try
            Select Case i
                Case 1
                    doblez2 = New MicroFold6(doblez1.doku)
                    doblez2.direction = direction
                    If doblez2.MakeSecondFold() Then
                        doku = doblez2.doku
                        MakeRestTie(i + 1)

                    End If
                Case 2
                    doblez3 = New MacroFold5(doblez2.doku)
                    doblez3.direction = direction
                    If doblez3.MakeThirdFold() Then
                        doku = doblez3.doku
                        MakeRestTie(i + 1)

                    End If
                Case 3
                    doblez4 = New MicroFold6(doblez3.doku)
                    doblez4.direction = direction
                    If doblez4.MakeFourthFold() Then
                        doku = doblez4.doku
                        MakeRestTie(i + 1)
                    End If
                Case 4
                    doblez5 = New MacroFold5(doblez4.doku)
                    doblez5.direction = direction
                    If doblez5.MakeFifthFold() Then
                        doku = doblez5.doku
                        MakeRestTie(i + 1)
                    End If
                Case 5
                    doblez6 = New MicroFold6(doblez5.doku)
                    doblez6.direction = direction
                    If doblez6.MakeSixthFold() Then
                        doku = doblez6.doku
                        MakeRestTie(i + 1)
                    End If
                Case 6
                    manager.Update(doblez6.doku)
                    If (manager.IsReadyForLastFold() Or True) Then
                        comando.MakeInvisibleSketches(doku)
                        MakeRestTie(10)
                    Else
                        comando.MakeInvisibleSketches(doku)
                        doblez7 = New MacroFold5(doblez6.doku)
                        If doblez7.MakeSeventhFold() Then
                            doku = doblez7.doku
                            MakeRestTie(i + 1)
                        End If
                    End If
                Case 7
                    giro = New TwistFold7(doblez7.doku)
                    giro.direction = direction
                    If giro.MakeFinalTwist() Then
                        Return doku
                    End If
                Case 8
                    giro = New TwistFold7(doblez8.doku)
                    giro.direction = direction
                    If giro.MakeFinalTwist() Then
                        Return doku
                    End If
                Case 9

                    If giro.MakeFinalTwist() Then
                        Return doku
                    End If

                Case 10
                    Try
                        giro = New TwistFold7(doblez8.doku)
                        giro.direction = direction
                    Catch ex As Exception
                        giro = New TwistFold7(doblez6.doku)
                        giro.direction = direction
                    End Try

                    If giro.MakeFinalTwist() Then
                        Return doku
                    End If
                Case Else
                    Return doku
            End Select
            Return doku
        Catch ex As Exception

            Return Nothing
        End Try

    End Function
    Function CreateFoldObjects(i As Integer) As PartDocument
        Select Case i
            Case 1
                doblez1 = New InitFold(doku)
                doblez1.direction = direction
                Return doku
            Case 2
                CreateFoldObjects(i - 1)
                doblez2 = New MicroFold6(doblez1.doku)
                doblez2.direction = direction
            Case 3
                CreateFoldObjects(i - 1)
                doblez3 = New MacroFold5(doblez2.doku)
                doblez3.direction = direction
            Case 4
                CreateFoldObjects(i - 1)
                doblez4 = New MicroFold6(doblez3.doku)
                doblez4.direction = direction
            Case 5
                CreateFoldObjects(i - 1)
                doblez5 = New MacroFold5(doblez4.doku)
                doblez5.direction = direction
            Case 6
                CreateFoldObjects(i - 1)
                doblez6 = New MicroFold6(doblez5.doku)
                doblez6.direction = direction
            Case 7
                CreateFoldObjects(i - 1)
                doblez7 = New MacroFold5(doblez6.doku)
                doblez7.direction = direction
            Case 8
                CreateFoldObjects(i - 1)
                doblez8 = New MicroFold6(doblez7.doku)
                doblez8.direction = direction
            Case 9
                CreateFoldObjects(i - 1)
                giro = New TwistFold7(doblez8.doku)
                giro.direction = direction
            Case Else
                Return doku
        End Select
        Return Nothing
    End Function
    Function GetInitialConditions() As Integer
        Try
            Dim q As Integer = 0
            doblez1 = New InitFold(invDoc.CreateSheetMetalFile(nombrador.MakeNextFileName(refDoc.oDoc)))
            doku = doblez1.doku
            trobinaCurve = New Curves3D(doblez1.doku)
            trobinaCurve.direction = direction
            trobinaCurve.DefineTrobinaParameters(doku)
            q = nombrador.GetQNumber(doku)
            doku.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(q - 1, UnitsTypeEnum.kUnitlessUnits, "currentQ")
            qNext = q
            Return q
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function GetStartingFeature() As Integer
        Try
            Dim i As Integer = 0
            If invDoc.IsAlreadyCreated(nombrador.MakeNextFileName(refDoc.oDoc)) Then
                doku = invDoc.documento
                compDef = doku.ComponentDefinition
                sheetMetalFeatures = compDef.Features
                If doku.ComponentDefinition.Features.Count > 0 Then
                    i = nombrador.GetFeatureNumber(manager.GetLastFold(doku))
                    CreateFoldObjects(i)
                Else
                    i = 0
                End If

            End If
            Return i
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function FindLastTie() As PartDocument
        Dim max, max2 As Integer
        max = 0
        max2 = 0
        Dim ltn, slt As String
        ltn = doku.FullFileName
        slt = ltn
        Dim lastTie As PartDocument
        fullFileNames = Directory.GetFiles(projectManager.ActiveDesignProject.WorkspacePath, "*.ipt")
        For Each s As String In fullFileNames
            If nombrador.GetFileNumber(s) > max2 Then
                If nombrador.GetFileNumber(s) > max Then
                    max2 = max
                    max = nombrador.GetFileNumber(s)
                    slt = ltn
                    ltn = s
                Else
                    max2 = nombrador.GetFileNumber(s)
                    slt = s

                End If
            Else

            End If

        Next
        lastTie = invDoc.OpenFullFileName(ltn)
        If IsLastTieFinish(lastTie) Then

        Else
            lastTie = invDoc.OpenFullFileName(slt)
        End If

        qLastTie = nombrador.GetQNumber(lastTie) - 1
        Return lastTie
    End Function
    Public Function MoveAllTies(i As Integer) As Integer
        Try

            Dim newFile, newFullFileName As String

            Dim p As String = projectManager.ActiveDesignProject.WorkspacePath
            Dim nd As String
            nd = String.Concat(p, "\Iteration", i.ToString)
            If (Directory.Exists(nd)) Then
                Debug.Print("That path exists already.")
                Directory.Delete(nd)
            End If
            di = Directory.CreateDirectory(nd)
            fullFileNames = Directory.GetFiles(p, "*.ipt")
            For Each s As String In fullFileNames
                Try
                    newFile = System.IO.Path.GetFileName(s)
                    newFullFileName = System.IO.Path.Combine(di.FullName, newFile)
                    If System.IO.File.Exists(newFullFileName) Then
                        System.IO.File.Delete(newFullFileName)
                    End If
                    System.IO.File.Move(s, newFullFileName)
                Catch ex As Exception
                    MsgBox(ex.ToString())
                    Return Nothing
                End Try
            Next
            Return di.GetFiles.Count
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function

    Public Function IsLastTieFinish(d As PartDocument) As Boolean
        Dim b As Boolean = False

        If d.ComponentDefinition.Features.Count > 5 Then
            Try
                If d.ComponentDefinition.Sketches3D.Item("introLine").SketchLines3D.Count > 0 Then
                    b = True
                End If
            Catch ex As Exception
                b = False
            End Try


        End If
        Return b
    End Function

End Class
