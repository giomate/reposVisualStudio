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
    Dim started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim invDoc As InventorFile
    Dim iteration As Integer
    Dim piedra As Wedges
    Dim esqueleto As Skeletons
    Public Sub New(app As Inventor.Application)
        oApp = app
        iteration = 13
        If oApp.Visible Then
            started = True
        End If
    End Sub
    Public Function AddInSkeletons() As Boolean
        done = MakeAllSkeletons()
        Return done
    End Function
    Public Function MakeAllSkeletons() As Boolean
        Dim b As Boolean
        Try
            If (started) Then
                Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
                Dim ffn As String
                ffn = String.Concat(p, "\Iteration13\Band9.ipt")
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
