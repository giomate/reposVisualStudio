Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory
Imports Inventor
Public Class Form1
    Dim oApp As Inventor.Application
    Dim oDoc As Inventor.Document
    Dim started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim invDoc As InventorFile
    Dim iteration As Integer
    Dim piedra As Wedges
    Dim esqueleto As Skeletons

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        iteration = 20
        ' Add any initialization after the InitializeComponent() call.
        Try
            oApp = Marshal.GetActiveObject("Inventor.Application")
            started = True

        Catch ex As Exception
            Try
                Dim oInvAppType As Type = GetTypeFromProgID("Inventor.Application")

                oApp = CreateInstance(oInvAppType)
                oApp.Visible = True

                'Note: if you shut down the Inventor session that was started
                'this(way) there is still an Inventor.exe running. We will use
                'this Boolean to test whether or not the Inventor App  will
                'need to be shut down.
                started = True

            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                MsgBox("Unable to get or start Inventor")
            End Try
        End Try
    End Sub
    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If started Then
            'oApp.ActiveDocument.Close()
            MsgBox("Closing Session")
        End If
        oApp = Nothing
    End Sub


    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Try
            oDesignProjectMgr = oApp.DesignProjectManager
            invDoc = New InventorFile(oApp)
            ' done = MakeSkeletonTest()
            done = MakeAllSkeletons()

            If done Then
                Me.Close()
            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Unable to find Document")
        End Try
    End Sub
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
    Public Function MakeSkeletonTest() As Boolean
        Dim b As Boolean
        Try
            If (started) Then
                Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
                Dim ffn As String
                ffn = String.Concat(p, "\Iteration", iteration.ToString, "\Band9.ipt")
                oDoc = invDoc.OpenFullFileName(ffn)
                esqueleto = New Skeletons(oDoc)
                esqueleto.RecoverSkeletonIteration(iteration)

                b = esqueleto.done
            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
End Class
