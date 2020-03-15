Imports Inventor
Imports System
Imports System.IO
Public Class BandAssembly
    Dim oApp As Inventor.Application
    Dim oDesignProjectMgr As DesignProjectManager
    Dim oAsmDoc As AssemblyDocument
    Dim doku As PartDocument
    Public done As Boolean
    Dim di As DirectoryInfo
    Dim fullFileNames As String()
    Dim invApp As InventorFile
    Dim oTransBRep As TransientBRep
    Dim oSurfaces As ObjectCollection
    Dim tg As TransientGeometry

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


    Public Sub New(App As Inventor.Application)
        oApp = App
        oDesignProjectMgr = oApp.DesignProjectManager
        'oAsmDoc = oApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, oApp.FileManager.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject))
        'Conversions.SetUnitsToMetric(oAsmDoc)
        invApp = New InventorFile(oApp)
        oSurfaces = oApp.TransientObjects.CreateObjectCollection()
        oTransBRep = oApp.TransientBRep
        tg = App.TransientGeometry
        DP.Dmax = 200
        DP.Dmin = 50
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 17
        DP.q = 37
        done = False
    End Sub
    Function createFileName(fileName As String) As String
        Dim strFilePath As String
        strFilePath = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
        ' Dim strFileName As String
        'strFileName = "Embossed" & CStr(I) & ".ipt"
        Dim strFullFileName As String
        strFullFileName = strFilePath & "\" & fileName
        Return strFullFileName
    End Function

    Function rotationTotal(qValue As Integer) As Matrix
        Dim oTG As TransientGeometry
        oTG = oApp.TransientGeometry
        Dim rMatrix As Matrix
        rMatrix = oTG.CreateMatrix
        rMatrix = rotationRadial(qValue, rMatrix)
        'oMatrix = rotationPhi(qValue, oMatrix)
        Dim pMatrix As Matrix
        pMatrix = oTG.CreateMatrix
        pMatrix = rotationPhi(qValue, pMatrix)

        Dim tMatrix As Matrix
        tMatrix = oTG.CreateMatrix
        Call tMatrix.SetToIdentity()
        Call tMatrix.PreMultiplyBy(pMatrix)
        Call tMatrix.PostMultiplyBy(rMatrix)


        Return tMatrix
    End Function
    Function rotationPhi(qValue As Integer, matrix As Matrix) As Matrix
        Dim oTG As TransientGeometry
        oTG = oApp.TransientGeometry
        Dim oMatrix As Matrix
        oMatrix = matrix
        Call oMatrix.SetToRotation(2 * Math.PI * DP.p * CDbl(qValue) / DP.q,
                            oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0))
        Return oMatrix
    End Function


    Function rotationRadial(qValue As Integer, matrix As Matrix) As Matrix
        Dim oTG As TransientGeometry
        oTG = oApp.TransientGeometry
        Dim oMatrix As Matrix
        oMatrix = matrix
        Dim x As Double
        Dim y As Double
        Dim Tx As Double
        Dim Ty As Double

        x = Tr * Math.Cos(2 * Math.PI * CDbl(qValue) / DP.q)
        y = Tr * Math.Sin(2 * Math.PI * CDbl(qValue) / DP.q)
        Tx = Tr * Math.Cos(2 * Math.PI * CDbl(qValue) / DP.q + Math.PI / 2)
        Ty = Tr * Math.Sin(2 * Math.PI * CDbl(qValue) / DP.q + Math.PI / 2)
        Dim Remainder As Integer
        Call Math.DivRem(CInt(DP.p * qValue), CInt(DP.q), Remainder)
        Call oMatrix.SetToRotation(0.0045 * (2 * CDbl(Remainder) / DP.q - 1) / 1,
                            oTG.CreateVector(0 * Tx / 10, Tr / 10, 0), oTG.CreatePoint(Tr / 10, 0 * y / 10, 0))
        Return oMatrix
    End Function

    Public Sub putPart(qvalue As Integer, fileName As String)
        Dim oOcc As ComponentOccurrence
        oOcc = oAsmDoc.ComponentDefinition.Occurrences.Add(fileName, rotationTotal(qvalue))

    End Sub
    Function PutSingleTie(fileName As String) As ComponentDefinition
        Dim oTG As TransientGeometry
        oTG = oApp.TransientGeometry
        Dim tMatrix As Matrix
        tMatrix = oTG.CreateMatrix
        Call tMatrix.SetToIdentity()
        Dim oOcc As ComponentOccurrence
        oOcc = oAsmDoc.ComponentDefinition.Occurrences.Add(fileName, tMatrix)
        Return oOcc.Definition
    End Function
    Function PutDonnut(fileName As String) As ComponentOccurrence
        Dim oTG As TransientGeometry
        oTG = oApp.TransientGeometry
        Dim tMatrix As Matrix
        tMatrix = oTG.CreateMatrix
        Call tMatrix.SetToIdentity()
        Dim oOcc As ComponentOccurrence
        oOcc = oAsmDoc.ComponentDefinition.Occurrences.Add(fileName, tMatrix)
        Return oOcc
    End Function

    Public Function createBandAsm() As AssemblyDocument

        Dim oDoc As AssemblyDocument
        oDoc = oApp.ActiveDocument
        Dim counter As Integer
        Dim fileName As String = "CondesatorBand4.ipt"
        Geometrics.AddSketchCircle(oApp)

        For counter = 1 To CInt(DP.q)
            Call putPart(counter, createFileName(fileName))

        Next
        Return oDoc
    End Function
    Public Function PutTies() As AssemblyDocument
        Try
            Dim oDoc As AssemblyDocument
            Dim counter As Integer
            Dim fileName As String = "Band0.ipt"
            Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
            Dim nd As String
            nd = String.Concat(p, "\Iteration1")
            If (Directory.Exists(nd)) Then
                Debug.Print("That path exists already.")
                Geometrics.AddSketchCircle(oApp)
                fullFileNames = Directory.GetFiles(nd, "*.ipt")
                For Each s As String In fullFileNames
                    Try
                        PutSingleTie(s)
                    Catch ex As Exception
                        MsgBox(ex.ToString())
                        Return Nothing
                    End Try
                Next

            End If
            oDoc = oApp.ActiveDocument
            Return oDoc
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function TrimDonnut() As PartDocument
        Try
            Dim workDonnut, oWorkSurface As WorkSurface
            Dim donnutBody, tieBody As SurfaceBody
            Dim ocu As ComponentOccurrence
            Dim guia, donnut As PartDocument
            Dim derivedDefinition As DerivedPartDefinition
            Dim newComponent As DerivedPartComponent
            Dim bolsa As SculptSurface
            Dim j, k As Integer

            Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath

            Dim nd As String
            nd = String.Concat(p, "\donnut2.ipt")
            donnut = invApp.OpenFullFileName(nd)
            nd = String.Concat(p, "\Iteration1")
            guia = donnut
            If (Directory.Exists(nd)) Then
                Debug.Print("That path exists already.")
                'Geometrics.AddSketchCircle(oApp)
                fullFileNames = Directory.GetFiles(nd, "*.ipt")
                For Each s As String In fullFileNames
                    Try
                        j = guia.ComponentDefinition.WorkSurfaces.Count
                        derivedDefinition = guia.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
                        derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsWorkSurface
                        newComponent = guia.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)
                        guia.Update2(True)
                        k = guia.ComponentDefinition.WorkSurfaces.Count
                        For index = j + 1 To k
                            oSurfaces.Add(guia.ComponentDefinition.Features.SculptFeatures.CreateSculptSurface(guia.ComponentDefinition.WorkSurfaces(index), PartFeatureExtentDirectionEnum.kNegativeExtentDirection))
                        Next
                        Try
                            guia.ComponentDefinition.Features.SculptFeatures.Add(oSurfaces, PartFeatureOperationEnum.kCutOperation)
                            oSurfaces.Clear()

                        Catch ex As Exception
                            oSurfaces.Clear()
                            oSurfaces.Add(guia.ComponentDefinition.Features.SculptFeatures.CreateSculptSurface(guia.ComponentDefinition.WorkSurfaces(j + 1), PartFeatureExtentDirectionEnum.kNegativeExtentDirection))
                            guia.ComponentDefinition.Features.SculptFeatures.Add(oSurfaces, PartFeatureOperationEnum.kCutOperation)
                            oSurfaces.Clear()
                        End Try



                    Catch ex As Exception
                        MsgBox(ex.ToString())
                        Return Nothing
                    End Try
                Next

                done = True
            End If

            Return donnut
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Function MergeAllWedges() As PartDocument
        Dim p As PartDocument
        Dim derivedDefinition As DerivedPartDefinition
        Dim newComponent As DerivedPartComponent
        Try
            p = oApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject,, True)
            Conversions.SetUnitsToMetric(p)
            p.Update2(True)
            MakeRing()
            derivedDefinition = p.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
            derivedDefinition.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsWorkSurface
            newComponent = p.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)

        Catch ex As Exception

        End Try
        Return doku
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
            ps.SketchCircles.AddByCenterRadius(spt, Cr / 2 + 3 / 10)
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
    Public Function DeriveAllTies() As PartDocument
        Try
            Dim guia As PartDocument
            Dim derivedDefinition As DerivedPartDefinition

            Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath

            Dim nd As String
            nd = String.Concat(p, "\Iteration1")
            guia = oApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, , True)
            If (Directory.Exists(nd)) Then
                Debug.Print("That path exists already.")
                Geometrics.AddSketchCircle(oApp)
                fullFileNames = Directory.GetFiles(nd, "*.ipt")
                For Each s As String In fullFileNames
                    Try


                        derivedDefinition = guia.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
                        derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
                        guia.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)



                    Catch ex As Exception
                        MsgBox(ex.ToString())
                        Return Nothing
                    End Try
                Next

                done = True
            End If

            Return guia
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function DeriveAllWedges() As PartDocument
        Try
            Dim guia As PartDocument
            Dim derivedDefinition As DerivedPartDefinition

            Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath

            Dim nd As String
            nd = String.Concat(p, "\Iteration1")
            If (Directory.Exists(nd)) Then
                fullFileNames = Directory.GetFiles(nd, "*.ipt")
                For Each s As String In fullFileNames
                    If  Then

                    End If
                    Try
                        derivedDefinition = guia.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
                        derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
                        guia.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)

                    Catch ex As Exception
                        MsgBox(ex.ToString())
                        Return Nothing
                    End Try
                Next

                done = True
            End If

            Return guia
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function MakeRails() As PartDocument
        Try
            Dim guia As PartDocument
            Dim derivedDefinition As DerivedPartDefinition

            Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath

            Dim nd As String
            nd = String.Concat(p, "\Iteration1")
            guia = oApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, , True)
            If (Directory.Exists(nd)) Then
                Debug.Print("That path exists already.")
                Geometrics.AddSketchCircle(oApp)
                fullFileNames = Directory.GetFiles(nd, "*.ipt")
                For Each s As String In fullFileNames
                    Try


                        derivedDefinition = guia.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(s)
                        derivedDefinition.IncludeAllSolids = DerivedComponentOptionEnum.kDerivedIncludeAll
                        guia.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(derivedDefinition)




                    Catch ex As Exception
                        MsgBox(ex.ToString())
                        Return Nothing
                    End Try
                Next

                done = True
            End If

            Return guia
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return Nothing
        End Try

    End Function
    Public Function MoveAllTies(i As Integer) As Integer
        Try

            Dim newFile, newFullFileName As String

            Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
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
End Class
