Imports Inventor
Imports 3D

Module BandSheetMetal
    Public Sub CreateFile(oApp As Inventor.Application)
        ' Create a new sheet metal document, using the default sheet metal template.
        Dim Doc As Inventor.Document = oApp.ActiveDocument
        Dim oSheetMetalDoc As PartDocument
        oSheetMetalDoc = oApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                 oApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject, , , "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"))

        Dim oPartDoc As PartDocument
        oPartDoc = oApp.ActiveDocument

        ' Set a reference to the component definition.
        Dim oCompDef As SheetMetalComponentDefinition
        oCompDef = oSheetMetalDoc.ComponentDefinition

        ' Set a reference to the sheet metal features collection.
        Dim oSheetMetalFeatures As SheetMetalFeatures
        oSheetMetalFeatures = oCompDef.Features

        Dim oSheetMetalCompDef As SheetMetalComponentDefinition
        oSheetMetalCompDef = oPartDoc.ComponentDefinition

        ' Override the thickness for the document
        oSheetMetalCompDef.UseSheetMetalStyleThickness = False

        ' Get a reference to the parameter controlling the thickness.
        Dim oThicknessParam As Parameter
        oThicknessParam = oSheetMetalCompDef.Thickness

        ' Change the value of the parameter.
        oThicknessParam.Value = oThicknessParam.Value / 10

        ' Update the part.
        oApp.ActiveDocument.Update()
    End Sub
    Function CreateFaceFeature(oDoc As PartDocument) As FaceFeature
        Dim o3DSketch As New CondesatorBandRolling.Class3DSketch
        o3DSketch.DrawFirstFace(oDoc)
        o3DSketch.
        Dim oProfile As Profile
        oProfile = oDoc.ComponentDefinition.Sketches.Item(1).Profiles.AddForSolid
        Dim oFaceFeatureDefinition As FaceFeatureDefinition
        oFaceFeatureDefinition = oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(oProfile)
        ' Create a face feature.
        Dim oFaceFeature As FaceFeature
        oFaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition
        Return oFaceFeature
    End Function


End Module
