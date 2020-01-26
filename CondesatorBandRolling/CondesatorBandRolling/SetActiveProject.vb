Imports Inventor

Module ProjectDefinitions
    Public Sub SetActiveProject(oApp As Inventor.Application)

        ' Check to make sure a document isn't open.
        If oApp.Documents.Count > 0 Then
            MsgBox("All documents must be closed before changing the project.")
            Exit Sub
        End If

        ' Set a reference to the DesignProjectManager object.
        Dim oDesignProjectMgr As DesignProjectManager
        oDesignProjectMgr = oApp.DesignProjectManager

        ' Show the current project.
        Debug.Print("Active project: " & oDesignProjectMgr.ActiveDesignProject.FullFileName)

        ' Get the project to activate
        ' This assumes that "C:\Temp\MyProject.ipj" exists.
        Dim oProject As DesignProject
        oProject = oDesignProjectMgr.ActiveDesignProject

        ' Activate the project
        oProject.Activate()

        ' Show the current project after making the project change.
        Debug.Print("New active project: " & oDesignProjectMgr.ActiveDesignProject.FullFileName)
    End Sub
End Module