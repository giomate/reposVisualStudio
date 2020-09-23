Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory
Imports Inventor
Public Class Starter
    Dim oApp As Inventor.Application
    Dim oDoc As Inventor.Document
    Public started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim invDoc As InventorFile
    Dim iteration As Integer
    Dim piedra As Wedges
    Dim esqueleto As Skeletons
    Public Sub New(app As Inventor.Application)
        oApp = app
        iteration = 20
        oDesignProjectMgr = oApp.DesignProjectManager
        invDoc = New InventorFile(oApp)
        If oApp.Visible Then
            started = True
        End If
    End Sub
    Public Function AddInSkeletons(i As Integer) As Boolean
        iteration = i
        done = MakeAllSkeletons()
        Return done
    End Function
    Public Function MakeAllSkeletons() As Boolean
        Dim b As Boolean
        Try
            If (started) Then
                Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
                Dim ffn As String
                ffn = String.Concat(p, "\Iteration", iteration.ToString, "\Band9.ipt")
                oDoc = invDoc.OpenFullFileName(ffn)
                esqueleto = New Skeletons(oDoc)
                esqueleto.MakeSkeletonIteration(iteration)

                b = esqueleto.done
            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
End Class
