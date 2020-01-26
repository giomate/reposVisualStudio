Imports Inventor

Public Sub EmbossFaces(value As Integer, strFullFileName As String)

    Dim oApp As Inventor.Application
    oApp = ThisApplication
    Dim oDesignProjectMgr As DesignProjectManager
    oDesignProjectMgr = oApp.DesignProjectManager



    Dim oDoc As PartDocument

    'If oDesignProjectMgr.IsFileInActiveProject(strFileName, kWorkspaceLocation, strFilePath) = True Then

    'On Error Resume Next
    oDoc = ThisApplication.Documents.Add(kPartDocumentObject,
                 ThisApplication.FileManager.GetTemplateFile(kPartDocumentObject), True)
    'End If

    'Set oDoc = ThisApplication.Documents.Open(strFullFileName, True)

    Call DeleteAll()

    Dim strFilePath As String
    strFilePath = oDesignProjectMgr.ActiveDesignProject.WorkspacePath

    Dim oDerivedPartDef As DerivedPartUniformScaleDef
    oDerivedPartDef = ThisApplication.ActiveDocument.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateUniformScaleDef(strFilePath & "\Ribeta2.ipt")
    oDerivedPartDef.ScaleFactor = 1
    Dim oDerivedPart As DerivedPartComponent
    oDerivedPart = oDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add(oDerivedPartDef)

    Dim oPartCompDef As PartComponentDefinition
    oPartCompDef = ThisApplication.ActiveDocument.ComponentDefinition

    Dim wp(2) As WorkPlane
    Dim face As Integer

    For I = 1 To 2

        face = I
        wp(I) = oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oDerivedPart.WorkFeatures.Item(I), 0)
        wp(I).FlipNormal

        Call ExtrudeNumber(value, oPartCompDef, wp(I), face)
    Next


    ThisApplication.SilentOperation = True
    Call oDoc.SaveAs(strFullFileName, False)
    oDoc.Close
End Sub






