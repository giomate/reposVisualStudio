Imports Inventor
Module DeleteElements
    Public Sub DeleteAll()
        Call DeleteBodies()
        Call DeleteReferences()
        Call DeleteFeatures()
        Call DeleteSketches()
        'Call DeleteWorkplanes
    End Sub
    Public Sub DeleteBodies()
        Dim oObject As ObjectCollection

        For Each oObject In ThisApplication.ActiveDocument.ComponentDefinition.Occurrences

            Call ThisApplication.ActiveDocument.ComponentDefinition.DeleteObjects(oObject, False, False, False)


        Next
    End Sub
    Public Sub DeleteReferences()
        Dim oRef As DerivedPartComponent

        For Each oRef In ThisApplication.ActiveDocument.ComponentDefinition.ReferenceComponents.DerivedPartComponents

            'Call ThisApplication.ActiveDocument.ComponentDefinition.DeleteObjects(oObject, False, False, False)
            oRef.Delete


        Next
    End Sub

    Public Sub DeleteFeatures()

        Dim partDoc As PartDocument
        partDoc = ThisApplication.ActiveDocument

        Dim feature As PartFeature
        For Each feature In partDoc.ComponentDefinition.Features

            feature.Delete

        Next
    End Sub

    Public Sub DeleteSketches()

        Dim partDoc As PartDocument
        partDoc = ThisApplication.ActiveDocument

        Dim feature As Sketch
        For Each feature In partDoc.ComponentDefinition.Sketches

            feature.Delete

        Next
    End Sub


    Public Sub DeleteWorkplanes()
        Dim partDoc As PartDocument
        partDoc = ThisApplication.ActiveDocument

        Dim feature As WorkPlane
        For Each feature In partDoc.ComponentDefinition.WorkPlanes
            feature.Delete(True)

        Next
    End Sub

End Module
