﻿Imports Inventor
Public Class Commands
    Dim oCommandMgr As CommandManager
    Dim oControlDef As ControlDefinition
    Dim mainApp As Application
    Dim doku As Document
    Dim app As Application

    Public Sub New(App As Inventor.Application)
        oCommandMgr = App.CommandManager
        doku = App.ActiveDocument
        mainApp = App
    End Sub

    Public Sub UndoCommand()


        ' Get control definition for the line command. 

        oControlDef = oCommandMgr.ControlDefinitions.Item("AppUndoCmd")

        ' Execute the command. 
        Call oControlDef.Execute()
    End Sub
    Function IsUndoable() As Boolean
        Dim ud As Boolean
        oControlDef = oCommandMgr.ControlDefinitions.Item("AppUndoCmd")
        ud = oControlDef.Enabled
        Return ud
    End Function
    Function CopyCommand() As Boolean
        Dim b As Boolean
        oControlDef = oCommandMgr.ControlDefinitions.Item("AppCopyCmd")
        b = oControlDef.Enabled
        Return b
    End Function
    Public Sub TopRightView()


        Dim oCamera As Camera
        oCamera = mainApp.ActiveView.Camera
        oCamera.ViewOrientationType = ViewOrientationTypeEnum.kTopViewOrientation
        oCamera.Fit()
        oCamera.Apply()

    End Sub
    Public Sub WireFrameView(docu As PartDocument)

        doku = docu
        mainApp = doku.Parent
        mainApp.ActiveView.DisplayMode = DisplayModeEnum.kWireframeWithHiddenEdgesRendering
        doku.Update2(True)


    End Sub
    Public Sub RealisticView(docu As PartDocument)

        doku = docu
        mainApp = doku.Parent
        mainApp.ActiveView.DisplayMode = DisplayModeEnum.kRealisticRendering
        doku.Update2(True)


    End Sub
    Public Sub MakeInvisibleSketches(docu As Inventor.Document)
        For Each sk As Sketch In docu.ComponentDefinition.Sketches
            sk.Visible = False

        Next
        For Each sk3 As Sketch3D In docu.ComponentDefinition.Sketches3D
            sk3.Visible = False

        Next


    End Sub
    Public Sub HideSketches(docu As Inventor.Document)
        Dim part As PartDocument = docu
        part.ObjectVisibility.Sketches3D = False
        part.ObjectVisibility.Sketches = False



    End Sub
    Public Sub HideSurfaces(docu As Inventor.Document)
        Dim part As PartDocument = docu
        part.ObjectVisibility.ConstructionSurfaces = False
    End Sub
    Public Sub MakeInvisibleWorkPlanes(docu As Inventor.Document)
        For Each wp As WorkPlane In docu.ComponentDefinition.WorkPlanes
            wp.Visible = False

        Next



    End Sub

    Public Sub HideAllSurfaces(docu As Inventor.Document)
        For Each ws As WorkSurface In docu.ComponentDefinition.WorkSurfaces
            ws.Visible = False

        Next



    End Sub
End Class
