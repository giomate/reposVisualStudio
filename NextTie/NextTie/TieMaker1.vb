﻿Imports Inventor

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
    Public partNumber, qNext, qLastTie As Integer
    Dim bandLines, constructionLines As ObjectCollection
    Dim comando As Commands
    Public nombrador As Nombres
    Dim nextSketch As OriginSketch
    Dim cutProfile As Profile

    Dim pro As Profile
    Dim direction As Vector
    Dim feature As FaceFeature
    Dim cutfeature As CutFeature
    Dim bendLine, cutLine As SketchLine
    Public compDef As SheetMetalComponentDefinition
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
    Dim doblez1 As InitFold
    Dim doblez2 As MicroFold
    Dim doblez3 As MacroFold
    Dim doblez4 As MicroFold4
    Dim doblez5 As MacroFold5
    Dim doblez6 As MicroFold6
    Dim doblez7 As MacroFold5
    Dim manager As FoldingEvaluator
    Dim giro As TwistFold7
    Dim arrayFunctions As Collection
    Dim fullFileNames As String()
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        comando = New Commands(app)
        monitor = New DesignMonitoring(doku)
        invDoc = New InventorFile(app)
        projectManager = app.DesignProjectManager

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
        manager = New FoldingEvaluator(doku)


        done = False
    End Sub
    Public Function MakeNextTie() As PartDocument
        Try
            Dim i As Integer

            i = GetStartingFeature()
            If i < 1 Then
                If GetInitialConditions() > 0 Then
                    If doblez1.MakeFirstFold(refDoc, qNext) Then
                        doku = doblez1.doku
                        MakeRestTie(1)

                    End If
                End If
            Else
                MakeRestTie(i)
            End If
            If compDef.Sketches3D.Item("last").SketchLines3D.Count > 0 Then
                done = True
            End If

            Return doku
        Catch ex As Exception

            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MakeRestTie(i As Integer) As PartDocument
        Try
            Select Case i
                Case 1
                    doblez2 = New MicroFold(doblez1.doku)
                    If doblez2.MakeSecondFold() Then
                        doku = doblez2.doku
                        MakeRestTie(i + 1)

                    End If
                Case 2
                    doblez3 = New MacroFold(doblez2.doku)
                    If doblez3.MakeThirdFold() Then
                        doku = doblez3.doku
                        MakeRestTie(i + 1)

                    End If
                Case 3
                    doblez4 = New MicroFold4(doblez3.doku)
                    If doblez4.MakeForthFold() Then
                        doku = doblez4.doku
                        MakeRestTie(i + 1)
                    End If
                Case 4
                    doblez5 = New MacroFold5(doblez4.doku)
                    If doblez5.MakeFifthFold() Then
                        doku = doblez5.doku
                        MakeRestTie(i + 1)
                    End If
                Case 5
                    doblez6 = New MicroFold6(doblez5.doku)
                    If doblez6.MakeSixthFold() Then
                        manager.Update(doblez6.doku)
                        If manager.IsReadyForLastFold() Then
                            comando.MakeInvisibleSketches(doku)
                            MakeRestTie(10)
                        Else
                            comando.MakeInvisibleSketches(doku)
                            MakeRestTie(i + 1)
                        End If

                    End If
                Case 6
                    manager.Update(doblez6.doku)
                    If manager.IsReadyForLastFold() Then
                        comando.MakeInvisibleSketches(doku)
                        MakeRestTie(10)
                    Else
                        comando.MakeInvisibleSketches(doku)
                        doblez7 = New MacroFold5(doblez6.doku)
                        If doblez7.MakeFifthFold() Then
                            doku = doblez7.doku
                            MakeRestTie(10)
                        End If
                    End If
                Case 7
                    If giro.MakeFinalTwist() Then
                        Return doku
                    End If

                Case 10
                    giro = New TwistFold7(doblez6.doku)
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
                Return doku
            Case 2
                CreateFoldObjects(i - 1)
                doblez2 = New MicroFold(doblez1.doku)

            Case 3
                CreateFoldObjects(i - 1)
                doblez3 = New MacroFold(doblez2.doku)

            Case 4
                CreateFoldObjects(i - 1)
                doblez4 = New MicroFold4(doblez3.doku)
            Case 5
                CreateFoldObjects(i - 1)
                doblez5 = New MacroFold5(doblez4.doku)
            Case 6
                CreateFoldObjects(i - 1)
                doblez6 = New MicroFold6(doblez5.doku)
            Case 7
                CreateFoldObjects(i - 1)
                giro = New TwistFold7(doblez6.doku)
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
            trobinaCurve.DefineTrobinaParameters(doku)
            q = nombrador.GetQNumber(doku)
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

                i = nombrador.GetFeatureNumber(manager.GetLastFold(doku))
                CreateFoldObjects(i)
            End If
            Return i
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function FindLastTie() As PartDocument
        Dim max As Integer = 0
        Dim ltn As String
        Dim lastTie As PartDocument
        fullFileNames = Directory.GetFiles(projectManager.ActiveDesignProject.WorkspacePath, "*.ipt")
        For Each s As String In fullFileNames
            If nombrador.GetFileNumber(s) > max Then
                max = nombrador.GetFileNumber(s)
                ltn = s

            End If
        Next
        lastTie = invDoc.OpenFullFileName(ltn)
        qLastTie = nombrador.GetQNumber(lastTie)
        Return lastTie
    End Function
End Class