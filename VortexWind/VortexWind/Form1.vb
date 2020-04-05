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
    Dim vortice As VortexRod
    Dim monitor As DesignMonitoring
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        iteration = 1
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

            done = MakeTort()
            If done Then
                Me.Close()
            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Unable to find Document")
        End Try
    End Sub

    Public Function MakeTort() As Boolean
        Dim b As Boolean
        Try
            If (started) Then
                oDesignProjectMgr = oApp.DesignProjectManager
                Dim p As String = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
                Dim ffn As String
                ffn = String.Concat(p, "\Iteration6\Sabina.ipt")
                invDoc = New InventorFile(oApp)
                oDoc = invDoc.OpenFullFileName(ffn)
                vortice = New VortexRod(oDoc)
                monitor = New DesignMonitoring(vortice.doku)
                If monitor.IsFeatureHealthy(vortice.MakeAllWiresGuides(vortice.doku)) Then
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
