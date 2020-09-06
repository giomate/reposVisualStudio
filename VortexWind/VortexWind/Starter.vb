Imports Inventor
Imports System
Imports System.IO

Imports System.Text
Imports System.IO.Directory
Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Public Class Starter
    Dim oApp As Inventor.Application
    Dim oDoc As Inventor.Document
    Dim started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim invDoc As InventorFile
    Dim iteration As Integer
    Dim vortice As VortexRod
    Dim monitor As DesignMonitoring

    Public Sub New(app As Inventor.Application)
        oApp = app
        iteration = 15
        oDesignProjectMgr = oApp.DesignProjectManager

        If oApp.Visible Then
            started = True
        End If
    End Sub
    Public Function AddInStamper(i As Integer) As Boolean
        iteration = i
        done = StampGuides()
        Return done
    End Function



    Public Function StampGuides() As Boolean
        Dim b As Boolean
        Try
            If (started) Then
                oDesignProjectMgr = oApp.DesignProjectManager
                Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
                Dim ffn As String
                ffn = String.Concat(p, "\Iteration", iteration.ToString, "\Condesator1.ipt")
                invDoc = New InventorFile(oApp)
                oDoc = invDoc.OpenFullFileName(ffn)
                vortice = New VortexRod(oDoc)
                vortice.iteration = iteration
                monitor = New DesignMonitoring(vortice.doku)
                If monitor.IsFeatureHealthy(vortice.StampAllWireGuides(vortice.doku, True)) Then
                    b = vortice.done
                End If

            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function

End Class
