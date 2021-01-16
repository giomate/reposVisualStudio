Imports Inventor
Public Class Commands
    Dim oCommandMgr As CommandManager
    Dim oControlDef As ControlDefinition
    Dim app As Application
    Dim doku As Document


    Public Sub New(App As Inventor.Application)
        oCommandMgr = App.CommandManager
        doku = App.ActiveDocument
        Me.app = App
    End Sub

    Public Sub UndoCommand()


        ' Get control definition for the line command. 

        oControlDef = oCommandMgr.ControlDefinitions.Item("AppUndoCmd")

        ' Execute the command. 
        Call oControlDef.Execute()
    End Sub
    Public Sub UndoCommand(doku As PartDocument)


        ' Get control definition for the line command. 

        oControlDef = oCommandMgr.ControlDefinitions.Item("AppUndoCmd")

        ' Execute the command. 
        Call oControlDef.Execute()
        doku.Update()
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
        oCamera = app.ActiveView.Camera
        oCamera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation
        oCamera.Fit()
        oCamera.Apply()

    End Sub
    Public Sub UnfoldBand(doku As PartDocument)
        Dim oDef As ControlDefinition
        Dim oCompDef As SheetMetalComponentDefinition
        Try

            app = doku.Parent
            oDef = app.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
            oDef.Execute()

            oCompDef = doku.ComponentDefinition
            oCompDef.Unfold()
        Catch ex As Exception
            MsgBox(ex.ToString())

        End Try


    End Sub
    Public Sub RefoldBand(doku As PartDocument)

        Dim oDef As ControlDefinition
        app = doku.Parent
        oDef = app.CommandManager.ControlDefinitions.Item("PartSwitchRepresentationCmd")
        'oDef.Execute()
        Dim oCompDef As SheetMetalComponentDefinition
        oCompDef = doku.ComponentDefinition
        oCompDef.FlatPattern.ExitEdit()

    End Sub

    Public Sub MakeInvisibleSketches(docu As Inventor.Document)
        For Each sk As Sketch In docu.ComponentDefinition.Sketches
            sk.Visible = False

        Next
        For Each sk3 As Sketch3D In docu.ComponentDefinition.Sketches3D
            sk3.Visible = False

        Next


    End Sub
    Public Sub MakeInvisibleWorkPlanes(docu As Inventor.Document)
        For Each wp As WorkPlane In docu.ComponentDefinition.WorkPlanes
            wp.Visible = False

        Next



    End Sub


End Class
