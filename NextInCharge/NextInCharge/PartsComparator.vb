Imports Inventor
Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory

Public Class PartsComparator
    Dim oApp As Inventor.Application
    Dim oDesignProjectMgr As DesignProjectManager
    Dim oAsmDoc As AssemblyDocument
    Dim Candidate As PartDocument
    Dim newName As Nombres
    Dim fileName, pattern As String
    Public emsamble As InventorFile
    Dim rotatedCheck, collidedCheck, qCheckSet As ObjectCollection
    Public done As Boolean
    Dim FileNames As String()
    Dim counter As Integer
    Dim maxVolume, minVolume As Double


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
        oAsmDoc = oApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, oApp.FileManager.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject), True)
        Conversions.SetUnitsToMetric(oAsmDoc)
        collidedCheck = oApp.TransientObjects.CreateObjectCollection
        rotatedCheck = oApp.TransientObjects.CreateObjectCollection
        qCheckSet = oApp.TransientObjects.CreateObjectCollection
        DP.Dmax = 200
        DP.Dmin = 50
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 17
        DP.q = 37
        emsamble = New InventorFile(oApp)
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
    Function MatrixProblematicCollision(qValue As Integer) As Matrix
        Dim oTG As TransientGeometry
        oTG = oApp.TransientGeometry
        Dim rMatrix As Matrix
        rMatrix = oTG.CreateMatrix
        rMatrix = rotationRadial(qValue, rMatrix)

        Dim pMatrix As Matrix
        pMatrix = oTG.CreateMatrix
        pMatrix = rotationPhi(1, pMatrix)

        Dim tMatrix As Matrix
        tMatrix = oTG.CreateMatrix
        Call tMatrix.SetToIdentity()
        Call tMatrix.PreMultiplyBy(pMatrix)
        Call tMatrix.PostMultiplyBy(rMatrix)


        Return tMatrix
    End Function
    Function MatrixPhiRotation(qValue As Integer) As Matrix
        Dim oTG As TransientGeometry
        oTG = oApp.TransientGeometry

        Dim pMatrix As Matrix
        pMatrix = oTG.CreateMatrix
        pMatrix = rotationPhi(qValue, pMatrix)

        Return pMatrix
    End Function
    Function MatrixRadialRotation(qValue As Integer) As Matrix
        Dim oTG As TransientGeometry
        oTG = oApp.TransientGeometry
        Dim rMatrix As Matrix
        rMatrix = oTG.CreateMatrix
        Return rotationRadial(qValue, rMatrix)
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
        'Call oMatrix.SetToRotation(0.0045 * (2 * CDbl(Remainder) / DP.q - 1) / 1,
        'oTG.CreateVector(0 * Tx / 10, Tr / 10, 0), oTG.CreatePoint(Tr / 10, 0 * y / 10, 0))
        Call oMatrix.SetToRotation(2 * Math.PI * CDbl(qValue) / DP.q, oTG.CreateVector(0 * Tx / 10, Tr / 10, 0), oTG.CreatePoint(Tr / 10, 0 * y / 10, 0))
        Return oMatrix
    End Function

    Public Sub putPart(qvalue As Integer, fileName As String)
        Dim oOcc As ComponentOccurrence
        oOcc = oAsmDoc.ComponentDefinition.Occurrences.Add(fileName, rotationTotal(qvalue))

    End Sub
    Public Sub PutRotatedPart(qvalue As Integer, fileName As String)
        Dim oOcc As ComponentOccurrence
        oOcc = oAsmDoc.ComponentDefinition.Occurrences.Add(fileName, MatrixPhiRotation(qvalue))

    End Sub
    Public Sub PutRadialRotatedPart(qvalue As Integer, fileName As String)
        Dim oOcc As ComponentOccurrence
        oOcc = oAsmDoc.ComponentDefinition.Occurrences.Add(fileName, MatrixRadialRotation(qvalue))

    End Sub
    Public Sub PutProblematicRotatedPart(qvalue As Integer, fileName As String)
        Dim oOcc As ComponentOccurrence
        oOcc = oAsmDoc.ComponentDefinition.Occurrences.Add(fileName, MatrixProblematicCollision(qvalue))

    End Sub

    Public Function createBandAsm() As AssemblyDocument

        Dim oDoc As AssemblyDocument
        oDoc = oApp.ActiveDocument
        Dim counter As Integer
        Dim fileName As String = "BandFoldingIteration.ipt"
        pattern = "Iteration"
        newName = New Nombres
        Geometrics.AddSketchCircle(oApp)

        For counter = 1 To CInt(DP.q)
            Call putPart(counter, createFileName(fileName))

        Next
        Return oDoc
    End Function

    Function CompareTwoRotatedParts(fullNamea As String, fullNameb As String) As InterferenceResults

        DeleteAllOcurrences(oAsmDoc)
        PutRotatedPart(1, fullNamea)
        PutRotatedPart(2, fullNameb)
        Return oAsmDoc.ComponentDefinition.AnalyzeInterference(CreateObjectCollection(oAsmDoc))
    End Function
    Function CreateObjectCollection(oDoc As AssemblyDocument) As ObjectCollection
        Dim collection As ObjectCollection = oApp.TransientObjects.CreateObjectCollection
        For Each occu As ComponentOccurrence In oDoc.ComponentDefinition.Occurrences
            collection.Add(occu)
        Next
        Return collection
    End Function
    Function CreateSingleObjectCollection(oDoc As AssemblyDocument) As ObjectCollection
        Dim collection As ObjectCollection = oApp.TransientObjects.CreateObjectCollection
        Dim singles As ObjectCollection = oApp.TransientObjects.CreateObjectCollection
        For Each occu As ComponentOccurrence In oDoc.ComponentDefinition.Occurrences
            singles.Add(occu)
            collection.Add(singles.Item(singles.Count))
            singles.RemoveByObject(occu)
        Next
        Return collection
    End Function
    Function CompareTwoCollidedParts(fullNamea As String, fullNameb As String) As InterferenceResults
        Try
            oAsmDoc.Activate()
            DeleteAllOcurrences(oAsmDoc)
            PutRotatedPart(1, fullNamea)
            PutRotatedPart(1, fullNameb)

            Return oAsmDoc.ComponentDefinition.AnalyzeInterference(CreateObjectCollection(oAsmDoc))
        Catch ex As Exception
            Debug.Print(ex.ToString())

            counter = 0
            Return CompareTwoProblematicParts(fullNamea, fullNameb)
        End Try

    End Function
    Function CompareTwoProblematicParts(fullNamea As String, fullNameb As String) As InterferenceResults
        Try
            DeleteAllOcurrences(oAsmDoc)
            counter = counter + 1
            PutRotatedPart(1, fullNamea)
            PutProblematicRotatedPart(counter, fullNameb)

            Return oAsmDoc.ComponentDefinition.AnalyzeInterference(CreateObjectCollection(oAsmDoc))
        Catch ex As Exception
            Debug.Print(ex.ToString())

            Return CompareTwoProblematicParts(fullNamea, fullNameb)
        End Try

    End Function
    Public Function CompareAllRotatedParts(fullname As String) As ObjectCollection

        oAsmDoc.Activate()


        For Each file As String In FileNames
            If Not file = fullname Then
                rotatedCheck.Add(CompareTwoRotatedParts(fullname, file))
                DeleteAllOcurrences(oAsmDoc)

            End If


        Next

        Return rotatedCheck
    End Function
    Function DeleteAllOcurrences(ass As AssemblyDocument) As Integer
        oAsmDoc.Activate()
        For Each occ As ComponentOccurrence In ass.ComponentDefinition.Occurrences
            occ.Delete2()
        Next


        Return 0
    End Function
    Public Function CompareAllCollidedParts(fullname As String) As ObjectCollection
        oAsmDoc.Activate()
        For Each file As String In FileNames
            If Not file = fullname Then
                collidedCheck.Add(CompareTwoCollidedParts(fullname, file))
                DeleteAllOcurrences(oAsmDoc)
            End If
        Next

        Return collidedCheck
    End Function

    Function FindSmallerInteference(results As ObjectCollection) As InterferenceResults

        Dim collision As Double = 999999999999999999
        Dim volume As Double = 0
        Dim minResults As InterferenceResults = Nothing
        For Each oResults As InterferenceResults In results
            volume = GetVolumeTotalInterferences(oResults)

            If volume < collision Then
                collision = volume
                minResults = oResults
            End If

        Next
        minVolume = collision

        Return minResults
    End Function
    Function FindLargerCollision(results As ObjectCollection) As InterferenceResults

        Dim collision As Double = 0
        Dim volume As Double = 0
        Dim maxResults As InterferenceResults = Nothing
        Dim maxinterference As InterferenceResult
        For Each oResults As InterferenceResults In results
            volume = GetVolumeTotalInterferences(oResults)

            If volume > 0 Then
                If volume > collision Then
                    collision = volume
                    maxResults = oResults
                End If
            End If

        Next
        maxVolume = collision
        Return maxResults
    End Function
    Public Function FindBestMatch(name As String) As String
        Dim fullname As String = emsamble.createFileName(name)
        FileNames = Directory.GetFiles(emsamble.manager.ActiveDesignProject.WorkspacePath, "*.ipt")
        CompareAllCollidedParts(fullname)
        CompareAllRotatedParts(fullname)
        Return Searchingloop(fullname)
    End Function
    Public Function FindCandidateFor(current As String) As PartDocument

        Return emsamble.openFileFullName(FindBestMatch(current))
    End Function
    Function Searchingloop(fullname As String) As String
        Try

            Dim minInterference As String = Nothing

            Dim i As Integer = 0
            For Each file As String In FileNames
                If Not file = fullname Then

                    If IsVolumeSimilar(CompareTwoRotatedParts(fullname, file), FindSmallerInteference(rotatedCheck)) Then
                        i = i + 1
                        minInterference = file
                        If IsVolumeSimilar(CompareTwoCollidedParts(fullname, file), FindLargerCollision(collidedCheck)) Then
                            done = True
                            Return file
                        End If
                    Else
                        RemoveCollidedInterferenceObject(i)
                    End If
                End If
            Next
            Return minInterference
        Catch ex As Exception
            Debug.Print(ex.ToString())
        End Try
    End Function
    Function GetVolumeTotalInterferences(interferences As InterferenceResults) As Double
        Dim volume As Double = 0
        For Each interference As InterferenceResult In interferences
            volume = interference.Volume + volume
        Next
        Return volume
    End Function
    Function IsVolumeSimilar(a As InterferenceResults, b As InterferenceResults) As Boolean
        Dim tolerance, delta As Double

        delta = GetVolumeTotalInterferences(a) - GetVolumeTotalInterferences(b)
        If maxVolume = 0 Then
            FindLargerCollision(collidedCheck)
        End If
        tolerance = Math.Exp(delta / (maxVolume - minVolume))
        If (tolerance < 1.1 And tolerance > 0.9) Then
            Return True
        Else
            Return False


        End If

        Return False
    End Function
    Function RemoveRotatedInterferenceObject(i As Integer) As Integer

        rotatedCheck.Remove(i + 1)

        Return i
    End Function
    Function RemoveCollidedInterferenceObject(i As Integer) As Integer

        collidedCheck.RemoveByObject(collidedCheck.Item(i + 1))

        Return i
    End Function
    Function RemoveInterferenceObject(i As Integer) As Integer
        RemoveCollidedInterferenceObject(i)
        RemoveRotatedInterferenceObject(i)
        Return i
    End Function
End Class
