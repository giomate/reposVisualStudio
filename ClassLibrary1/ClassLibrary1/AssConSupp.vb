Imports Inventor
Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices

Public Class AssConSupp
    Dim InvApp As Inventor.Application
    Dim started As Boolean = True
    Dim partDoc As PartDocument
    Dim partDef As PartComponentDefinition

    Public Sub AssCon()
        Dim oApp As Inventor.Application
        oApp = ThisApplication
        Dim oDesignProjectMgr As DesignProjectManager
        oDesignProjectMgr = oApp.DesignProjectManager

        Dim oAsmDoc As AssemblyDocument
        oAsmDoc = ThisApplication.Documents.Add(kAssemblyDocumentObject,
                ThisApplication.FileManager.GetTemplateFile(kAssemblyDocumentObject))



        Dim p As Double
        Dim q As Double
        p = 13
        q = 31


        Dim strFilePath As String
        strFilePath = oDesignProjectMgr.ActiveDesignProject.WorkspacePath

        For I = 1 To CInt(q)
            'p = CDbl(I)
            Dim strFileName As String
            strFileName = "Embossed" & CStr(I) & ".ipt"
            Dim strFullFileName As String
            strFullFileName = strFilePath & "\" & strFileName

            Call EmbossFaces(CInt(I), strFullFileName)
            Dim oTG As TransientGeometry
            oTG = ThisApplication.TransientGeometry

            Dim oMatrix As Matrix
            oMatrix = oTG.CreateMatrix
            trans = ThisApplication.TransientGeometry.CreateMatrix
            Call oMatrix.SetToRotation(8 * Atn(1) * p * CDbl(I) / q,
                            oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0))

            Dim oOcc As ComponentOccurrence
            oOcc = oAsmDoc.ComponentDefinition.Occurrences.Add(strFullFileName, oMatrix)


        Next
        Dim strAsmName As String
        strAsmName = "MoldSample2.iam"

        Call oAsmDoc.SaveAs(strFilePath & "\" & strAsmName, False)
        'oAsmDoc.Close

        Dim oPartDoc1 As PartDocument
        oPartDoc1 = ThisApplication.Documents.Add(kPartDocumentObject,
                 ThisApplication.FileManager.GetTemplateFile(kPartDocumentObject))

        Dim oDerivedAsmDef As DerivedAssemblyDefinition
        oDerivedAsmDef = oPartDoc1.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(oAsmDoc.FullDocumentName)

        oDerivedAsmDef.DeriveStyle = kDeriveAsSingleBodyNoSeams
        Dim strMoldName As String
        strMoldName = "CondesatorSupport2.ipt"

        Call oPartDoc1.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.Add(oDerivedAsmDef)
        Call oPartDoc1.SaveAs(strFilePath & "\" & strMoldName, False)

    End Sub
End Class
