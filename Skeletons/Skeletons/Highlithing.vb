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
    Public Sub LookAtFace(f As Face)
        Dim oSSet As SelectSet
        Try
            If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                oSSet = doku.SelectSet


                oSSet.Select(f)

                'change active view camera orientation
                Dim oControlDef As ControlDefinition
                oControlDef = app.CommandManager.ControlDefinitions.Item("AppLookAtCmd")
                oControlDef.Execute()
                Beep()
                oSSet.Clear()
                Try
                    ' FitView(doku)
                Catch ex As Exception

                End Try
            End If

        Catch ex As Exception

        End Try



    End Sub
    Public Sub LookAtPlane(wpl As WorkPlane)
        Dim oSSet As SelectSet

        oSSet = doku.SelectSet


        oSSet.Select(wpl)

        'change active view camera orientation
        Dim oControlDef As ControlDefinition
        oControlDef = app.CommandManager.ControlDefinitions.Item("AppLookAtCmd")
        oControlDef.Execute()
        Beep()
        oSSet.Clear()
        Try
            ' FitView(doku)
        Catch ex As Exception

        End Try




    End Sub
    Public Sub FitView(d As PartDocument)
        Dim dMax As Double = 0
        Dim fMax As Face
        Dim oCamera As Camera
        Try

            oCamera = app.ActiveView.Camera

            oCamera.Fit()
            oCamera.Apply()
        Catch ex As Exception

        End Try


    End Sub
End Class
