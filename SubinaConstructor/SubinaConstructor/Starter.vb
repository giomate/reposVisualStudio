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
    Public started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim invDoc As InventorFile
    Public iteration As Integer
    Dim nido As SubinaStruct

    Public Sub New(app As Inventor.Application)
        oApp = app
        iteration = 14
        oDesignProjectMgr = oApp.DesignProjectManager

        If oApp.Visible Then
            started = True
        End If
    End Sub
    Public Function AddInSubina(i As Integer) As Boolean
        iteration = i
        done = MakeCondesator()
        Return done
    End Function


    Public Function MakeTort(i As Integer) As Boolean
        Dim b As Boolean
        iteration = i
        Try
            If (started) Then
                oDesignProjectMgr = oApp.DesignProjectManager
                Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
                Dim ffn As String
                ffn = String.Concat(p, "\Iteration", iteration.ToString, "\Skeleton1.ipt")
                invDoc = New InventorFile(oApp)
                oDoc = invDoc.OpenFullFileName(ffn)
                nido = New SubinaStruct(oDoc)
                nido.MakeSimpleNestStruct(iteration)
                b = nido.done
            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
    Public Function MakeCondesator() As Boolean
        Dim b As Boolean
        Try
            If (started) Then
                oDesignProjectMgr = oApp.DesignProjectManager
                Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
                Dim ffn As String
                ffn = String.Concat(p, "\Iteration", iteration.ToString, "\Band1.ipt")
                invDoc = New InventorFile(oApp)
                oDoc = invDoc.OpenFullFileName(ffn)
                nido = New SubinaStruct(oDoc)
                nido.MakeCondesatorStruct(iteration)
                b = nido.done
            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function

End Class
