Imports Inventor
Public Class Highlithing
    Dim doku As PartDocument
    Dim oCommandMgr As CommandManager
    Dim oControlDef As ControlDefinition
    Dim app As Application
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
    End Sub
    Sub HighLighFace(f As Face)


        Dim oEndHLSet As HighlightSet

        oEndHLSet = doku.CreateHighlightSet

        ' Change the highlight color for the set to green.
        Dim oGreen As Color
        oGreen = app.TransientObjects.CreateColor(0, 255, 0)
        oEndHLSet.Color = oGreen

        ' Add all end faces to the highlightset.

        oEndHLSet.AddItem(f)
        oEndHLSet.Delete()

    End Sub
    Sub HighLighObject(o As Object)
        Dim oEndHLSet As HighlightSet
        oEndHLSet = doku.CreateHighlightSet
        ' Change the highlight color for the set to green.
        Dim oGreen As Color
        oGreen = app.TransientObjects.CreateColor(0, 255, 0)

        oEndHLSet.Color = oGreen

        ' Add all end faces to the highlightset.

        oEndHLSet.AddItem(o)
        oEndHLSet.Delete()

    End Sub
End Class
