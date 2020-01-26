Imports Inventor
Imports System
Public Class FindReferenceLine
    Public oDoc As PartDocument
    Dim sk3D As Sketch3D
    Dim line3D As SketchLines3D
    Public line As SketchLine3D
    Public Sub New(docu As Inventor.Document)
        oDoc = docu
    End Sub

    Public Function GetKeyLine(docu As Inventor.Document) As SketchLines3D
        oDoc = docu
        Return GetKeyLine()
    End Function
    Public Function GetKeyLine() As SketchLine3D

        line3D = OpenLastSketch(oDoc).SketchLines3D
        Debug.Print("Number of lines:  " & line3D.Count.ToString)
        line = line3D.Item(line3D.Count)
        Return line
    End Function
    Function OpenLastSketch(oDoc As PartDocument) As Sketch3D
        sk3D = oDoc.ComponentDefinition.Sketches3D.Item("last")
        Return sk3D
    End Function
    Public Function OpenMainSketch(docu As PartDocument) As Sketch3D
        oDoc = docu
        sk3D = oDoc.ComponentDefinition.Sketches3D.Item("s0")
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
