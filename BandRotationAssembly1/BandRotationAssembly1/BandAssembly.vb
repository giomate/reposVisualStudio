Imports Inventor

Public Class BandAssembly
    Dim oApp As Inventor.Application
    Dim oDesignProjectMgr As DesignProjectManager
    Dim oAsmDoc As AssemblyDocument

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
        oAsmDoc = oApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, oApp.FileManager.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject))
        Conversions.SetUnitsToMetric(oAsmDoc)
        DP.Dmax = 200
        DP.Dmin = 50
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 17
        DP.q = 37
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
End Class
