Imports Inventor
Imports System
Public Class FindReferenceLine
    Public oDoc As PartDocument
    Dim sk3D As Sketch3D
    Dim lines3D As SketchLines3D
    Public line As SketchLine3D
    Public foldFeatures As FoldFeatures
    Dim sheetMetalFeatures As SheetMetalFeatures
    Dim compDef As SheetMetalComponentDefinition
    Public Sub New(docu As Inventor.Document)
        oDoc = docu
        compDef = docu.ComponentDefinition
        sheetMetalFeatures = compDef.Features
        foldFeatures = sheetMetalFeatures.FoldFeatures
    End Sub

    Public Function GetKeyLine(docu As Inventor.Document) As SketchLines3D
        oDoc = docu
        Return GetKeyLine()
    End Function
    Public Function GetKeyLine() As SketchLine3D

        lines3D = OpenIntroLine(oDoc).SketchLines3D
        Debug.Print("Number of lines:  " & lines3D.Count.ToString)
        line = lines3D.Item(lines3D.Count)
        Return line
    End Function
    Public Function GetKanteLine() As SketchLine3D

        lines3D = OpenKanteSketch(oDoc).SketchLines3D
        Debug.Print("Number of lines:  " & lines3D.Count.ToString)
        line = lines3D.Item(lines3D.Count)
        Return line
    End Function
    Function OpenIntroLine(oDoc As PartDocument) As Sketch3D
        sk3D = oDoc.ComponentDefinition.Sketches3D.Item("introLine")
        Return sk3D
    End Function
    Function OpenKanteSketch(oDoc As PartDocument) As Sketch3D
        sk3D = oDoc.ComponentDefinition.Sketches3D.Item("kante")
        Return sk3D
    End Function
    Public Function OpenMainSketch(docu As PartDocument) As Sketch3D
        oDoc = docu
        sk3D = oDoc.ComponentDefinition.Sketches3D.Item("s1")
        Return sk3D
    End Function
    Public Function OpenMainSketch() As Sketch3D

        Return OpenMainSketch(oDoc)
    End Function
    Public Function DrawKeyLine(line As SketchLine3D) As Sketch3D
        Dim sketch As Sketch3D
        sketch = oDoc.ComponentDefinition.Sketches3D.Add()
        sketch.Name = "MainSkecht"
        sketch.SketchLines3D.AddByTwoPoints(line.StartPoint, line.EndPoint)
        sketch.SketchLines3D.Item(sketch.SketchLines3D.Count).Construction = True

        Return sketch
    End Function
End Class
