Imports Inventor
Module Geometrics
    Public Sub AddSketchCircle(oApp As Inventor.Application)
        Dim oDoc As Document
        oDoc = oApp.ActiveDocument
        ' Get the active part document.
        Dim partDef As ComponentDefinition
        partDef = oDoc.ComponentDefinition

        ' Create a new sketch.
        Dim sketch As PlanarSketch
        sketch = partDef.Sketches.Add(partDef.WorkPlanes.Item(3))

        Dim tg As TransientGeometry
        tg = oApp.TransientGeometry



        ' Create an circle
        Dim circle As SketchCircle
        circle = sketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(0, 0), 100 / 10)


    End Sub
End Module